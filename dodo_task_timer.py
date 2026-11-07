# -*- coding: utf-8 -*-
"""Dodo Task Timer - a compact Windows app for tracking daily task productivity."""
from __future__ import annotations

import math
import os
import sqlite3
import sys
import threading
import time
import queue
from collections import defaultdict
from datetime import datetime, date, time as dt_time, timedelta
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from pynput import keyboard
    PYNPUT_AVAILABLE = True
except Exception:
    keyboard = None
    PYNPUT_AVAILABLE = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    OPENPYXL_AVAILABLE = True
except Exception:
    Workbook = None
    OPENPYXL_AVAILABLE = False

APP_NAME = "Task Timer"
BG_BLUE = "#1687E8"
BG_GREEN = "#16A34A"
BG_YELLOW = "#FACC15"
BG_RED = "#DC2626"
DARK_TEXT = "#172033"
WHITE = "#FFFFFF"
BLUE_BUTTON = "#075EAA"
RED_BUTTON = "#991B1B"
NEUTRAL_BUTTON = "#E8EDF3"
NEUTRAL_TEXT = "#263445"
DAY_TARGET_SECONDS = 4 * 60 * 60
TASK_TARGET_SECONDS = 4 * 60


def utc_now() -> datetime:
    """Return a timezone-naive UTC timestamp for storage in SQLite."""
    return datetime.utcnow().replace(microsecond=0)


def local_now(offset_minutes: int = 180) -> datetime:
    """Return wall-clock time at a fixed UTC offset, independent of OS timezone."""
    return utc_now() + timedelta(minutes=int(offset_minutes))


def iso_now() -> str:
    """Persist all event timestamps in UTC (naive ISO format)."""
    return utc_now().isoformat(timespec="seconds")


def parse_dt(value) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value)
    except (TypeError, ValueError):
        return None


def format_stopwatch(seconds: int | float) -> str:
    """Compact stopwatch display: MM:SS.hh (hundredths; no separate hour field)."""
    total_centiseconds = max(0, int(round(float(seconds) * 100)))
    total_seconds, centiseconds = divmod(total_centiseconds, 100)
    total_minutes, whole_seconds = divmod(total_seconds, 60)
    return f"{total_minutes:02d}:{whole_seconds:02d}.{centiseconds:02d}"


def format_duration(seconds: int | float, include_seconds: bool = True) -> str:
    """Format durations left-to-right to avoid bidirectional text reordering."""
    total_seconds = max(0, int(round(seconds)))
    hours, rem = divmod(total_seconds, 3600)
    minutes, secs = divmod(rem, 60)
    pieces = []
    if hours:
        pieces.append(f"{hours} hr")
    if minutes or (hours and include_seconds):
        pieces.append(f"{minutes} min")
    if include_seconds and (secs or not pieces):
        pieces.append(f"{secs} sec")
    elif not include_seconds and total_seconds < 60:
        return f"{total_seconds} sec"
    return " ".join(pieces) if pieces else "0 sec"


def format_average(seconds: float | int | None) -> str:
    if seconds is None or seconds < 0:
        return "N/A"
    whole = int(round(seconds))
    minutes, secs = divmod(whole, 60)
    if minutes:
        return f"{minutes} min {secs} sec"
    return f"{secs} sec"


def normalize_shortcut(value: str) -> str:
    """Normalize keyboard shortcut names for comparisons and display."""
    parts = [part.strip().upper() for part in str(value).replace("CONTROL+", "CTRL+").split("+") if part.strip()]
    return "+".join(parts)


def tkinter_shortcut_sequence(value: str) -> str:
    """Convert e.g. F8 or Ctrl+F10 to a Tk binding sequence."""
    normalized = normalize_shortcut(value)
    if normalized.startswith("CTRL+"):
        return f"<Control-{normalized.split('+', 1)[1]}>"
    return f"<{normalized}>"


def shortcut_parts(value: str) -> tuple[bool, str]:
    """Return (requires_ctrl, function_key) for configured function-key shortcuts."""
    normalized = normalize_shortcut(value)
    if normalized.startswith("CTRL+"):
        return True, normalized.split("+", 1)[1]
    return False, normalized


def format_date(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def format_ampm_clock(value: str | None) -> str:
    """Format stored 24-hour HH:MM time for people using AM/PM."""
    if not value:
        return "—"
    try:
        return datetime.strptime(value, "%H:%M").strftime("%I:%M %p").lstrip("0")
    except (TypeError, ValueError):
        return str(value)


def format_datetime(value: str | None, utc_offset_minutes: int = 180) -> str:
    """Format a stored UTC timestamp in the record's configured fixed-offset timezone."""
    parsed = parse_dt(value)
    if not parsed:
        return "—"
    shown = parsed + timedelta(minutes=int(utc_offset_minutes or 0))
    return shown.strftime("%d/%m/%Y %H:%M:%S")


def legacy_local_timestamp_to_utc(value: str | None, legacy_offset_minutes: int = 180) -> str | None:
    """Convert timestamps from previous versions, which stored Egypt local wall time, to UTC."""
    parsed = parse_dt(value)
    if not parsed:
        return value
    return (parsed - timedelta(minutes=legacy_offset_minutes)).isoformat(timespec="seconds")


def color_for_elapsed(seconds: int | float) -> tuple[str, str]:
    # Blue: under 3 minutes; green: 3 to under 4 minutes; red: 4 minutes and above.
    if seconds < 3 * 60:
        return BG_BLUE, WHITE
    if seconds < 4 * 60:
        return BG_GREEN, WHITE
    return BG_RED, WHITE


def split_duration_by_date(start: datetime | None, end: datetime | None, duration_seconds: int | float) -> dict[date, float]:
    """Allocate measured task time across calendar dates, conserving the exact total."""
    duration_seconds = max(0.0, float(duration_seconds))
    if duration_seconds == 0:
        return {}
    if start is None:
        return {date.today(): duration_seconds}
    if end is None or end <= start:
        return {start.date(): duration_seconds}

    wall_total = (end - start).total_seconds()
    if wall_total <= 0:
        return {start.date(): duration_seconds}

    intervals: list[tuple[date, float]] = []
    cursor = start
    while cursor.date() <= end.date():
        next_midnight = datetime.combine(cursor.date() + timedelta(days=1), dt_time.min)
        segment_end = min(end, next_midnight)
        overlap = max(0.0, (segment_end - cursor).total_seconds())
        if overlap > 0:
            intervals.append((cursor.date(), overlap))
        if segment_end >= end:
            break
        cursor = segment_end

    if not intervals:
        return {start.date(): duration_seconds}

    result: dict[date, float] = {}
    allocated = 0.0
    for idx, (day, overlap) in enumerate(intervals):
        if idx == len(intervals) - 1:
            piece = duration_seconds - allocated
        else:
            piece = duration_seconds * overlap / wall_total
            piece = max(0.0, min(piece, duration_seconds - allocated))
        result[day] = result.get(day, 0) + piece
        allocated += piece
    return result


def workday_for_timestamp(value: datetime | None, utc_offset_minutes: int, start_minutes: int) -> date:
    """Map a stored UTC-naive timestamp to a workday using a saved UTC offset and day boundary."""
    if value is None:
        value = utc_now()
    local_value = value + timedelta(minutes=int(utc_offset_minutes or 0))
    boundary = max(0, min(1439, int(start_minutes or 0)))
    return (local_value - timedelta(minutes=boundary)).date()


def split_duration_by_workday(start: datetime | None, end: datetime | None, duration_seconds: int | float,
                              utc_offset_minutes: int, start_minutes: int) -> dict[date, float]:
    """Allocate duration across workdays, using the record's saved offset and boundary."""
    duration_seconds = max(0.0, float(duration_seconds))
    if duration_seconds <= 0:
        return {}
    if start is None:
        return {workday_for_timestamp(end, utc_offset_minutes, start_minutes): duration_seconds}
    if end is None or end <= start:
        return {workday_for_timestamp(start, utc_offset_minutes, start_minutes): duration_seconds}
    shifted_start = start + timedelta(minutes=int(utc_offset_minutes or 0))
    shifted_end = end + timedelta(minutes=int(utc_offset_minutes or 0))
    boundary = max(0, min(1439, int(start_minutes or 0)))
    wall_total = (shifted_end - shifted_start).total_seconds()
    if wall_total <= 0:
        return {workday_for_timestamp(start, utc_offset_minutes, start_minutes): duration_seconds}
    intervals = []
    cursor = shifted_start
    while cursor < shifted_end:
        workday = (cursor - timedelta(minutes=boundary)).date()
        next_boundary = datetime.combine(workday + timedelta(days=1), dt_time.min) + timedelta(minutes=boundary)
        segment_end = min(shifted_end, next_boundary)
        overlap = max(0.0, (segment_end - cursor).total_seconds())
        if overlap > 0:
            intervals.append((workday, overlap))
        if segment_end >= shifted_end:
            break
        cursor = segment_end
    result = {}
    allocated = 0.0
    for idx, (day, overlap) in enumerate(intervals):
        piece = duration_seconds - allocated if idx == len(intervals) - 1 else duration_seconds * overlap / wall_total
        piece = max(0.0, min(piece, duration_seconds - allocated))
        result[day] = result.get(day, 0.0) + piece
        allocated += piece
    return result


class DodoTaskTimer:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_NAME)
        # Compact footprint; trim the unused space below the action buttons.
        self.root.geometry("250x228")
        self.root.minsize(246, 224)
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.close_app)

        local_appdata = os.environ.get("LOCALAPPDATA")
        data_root = Path(local_appdata) if local_appdata else (Path.home() / ".local" / "share")
        self.data_dir = data_root / "DodoTaskTimer"
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.db_path = self.data_dir / "dodo_data.sqlite3"
        self.conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._init_db()
        self.settings = self._load_settings()
        self._migrate_legacy_local_timestamps()
        self.settings = self._load_settings()
        # Freeze defaults on any malformed/null legacy metadata after the one-time conversion.
        with self.conn:
            self.conn.execute("UPDATE tasks SET utc_offset_minutes=? WHERE utc_offset_minutes IS NULL", (self._current_offset_minutes(),))
            self.conn.execute("UPDATE tasks SET workday_start_minutes=? WHERE workday_start_minutes IS NULL", (self._current_workday_start_minutes(),))
            self.conn.execute("UPDATE manual_adjustments SET utc_offset_minutes=? WHERE utc_offset_minutes IS NULL", (self._current_offset_minutes(),))
            self.conn.execute("UPDATE manual_adjustments SET workday_start_minutes=? WHERE workday_start_minutes IS NULL", (self._current_workday_start_minutes(),))
        self._recover_orphaned_tasks()

        self.current_task_id: int | None = None
        self.current_base_elapsed = 0
        self.current_mono_start = 0.0
        self.last_display_elapsed = 0
        self.last_undo: dict | None = None
        self.last_checkpoint_mono = 0.0
        self.last_stats_update = 0.0
        self.current_bg = None
        self.report_window: tk.Toplevel | None = None
        self.report_daily_tree = None
        self.report_tasks_tree = None
        self.report_manual_tree = None
        self.report_week_selector = None
        self.report_week_tree = None
        self.report_week_summary = None
        self.report_notebook = None
        self.report_summary_frame = None
        self.keyboard_listener = None
        self.hotkey_queue: queue.Queue[str] = queue.Queue()
        self._hotkey_last_fired = {name: 0.0 for name in ("dodo", "undo", "stop", "skip")}
        self._hotkey_lock = threading.Lock()
        self._pynput_pressed_keys: set[str] = set()

        self._build_ui()
        self._apply_topmost()
        self._bind_local_hotkeys()
        self._start_global_listeners()
        self.root.after(100, self._tick)

    # ---------- Database and settings ----------
    def _init_db(self):
        with self.conn:
            self.conn.execute(
                """CREATE TABLE IF NOT EXISTS tasks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    started_at TEXT NOT NULL,
                    ended_at TEXT,
                    duration_seconds INTEGER NOT NULL DEFAULT 0,
                    duration_ms INTEGER NOT NULL DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'active',
                    completion_source TEXT,
                    last_checkpoint_at TEXT,
                    utc_offset_minutes INTEGER,
                    workday_start_minutes INTEGER
                )"""
            )
            self.conn.execute(
                """CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )"""
            )
            self.conn.execute(
                """CREATE TABLE IF NOT EXISTS manual_adjustments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    work_date TEXT NOT NULL,
                    task_count INTEGER NOT NULL CHECK(task_count >= 0),
                    total_ms INTEGER NOT NULL CHECK(total_ms >= 0),
                    created_at TEXT NOT NULL,
                    start_time TEXT,
                    end_time TEXT,
                    utc_offset_minutes INTEGER,
                    workday_start_minutes INTEGER
                )"""
            )
            manual_columns = {row["name"] for row in self.conn.execute("PRAGMA table_info(manual_adjustments)").fetchall()}
            if "start_time" not in manual_columns:
                self.conn.execute("ALTER TABLE manual_adjustments ADD COLUMN start_time TEXT")
            if "end_time" not in manual_columns:
                self.conn.execute("ALTER TABLE manual_adjustments ADD COLUMN end_time TEXT")
            if "utc_offset_minutes" not in manual_columns:
                self.conn.execute("ALTER TABLE manual_adjustments ADD COLUMN utc_offset_minutes INTEGER")
            if "workday_start_minutes" not in manual_columns:
                self.conn.execute("ALTER TABLE manual_adjustments ADD COLUMN workday_start_minutes INTEGER")
            # Migrate databases created by earlier versions without losing history.
            task_columns = {row["name"] for row in self.conn.execute("PRAGMA table_info(tasks)").fetchall()}
            if "duration_ms" not in task_columns:
                self.conn.execute("ALTER TABLE tasks ADD COLUMN duration_ms INTEGER NOT NULL DEFAULT 0")
            if "utc_offset_minutes" not in task_columns:
                self.conn.execute("ALTER TABLE tasks ADD COLUMN utc_offset_minutes INTEGER")
            if "workday_start_minutes" not in task_columns:
                self.conn.execute("ALTER TABLE tasks ADD COLUMN workday_start_minutes INTEGER")
            self.conn.execute(
                "UPDATE tasks SET duration_ms = duration_seconds * 1000 "
                "WHERE duration_ms = 0 AND duration_seconds > 0"
            )
            self.conn.execute(
                "DELETE FROM tasks WHERE status='incomplete' AND completion_source='close'"
            )
            # Drop malformed legacy rows marked complete without an end timestamp.
            # Such rows could add work-time seconds without counting as a task.
            self.conn.execute(
                "DELETE FROM tasks WHERE status='completed' AND (ended_at IS NULL OR ended_at = '')"
            )

    def _load_settings(self) -> dict:
        defaults = {
            "topmost": "1",
            "dodo_key": "F8",
            "undo_key": "F9",
            # Modified function keys avoid colliding with common browser actions.
            "stop_key": "Ctrl+F10",
            "skip_key": "Ctrl+F11",
            # Timestamps are stored in UTC and displayed at the configured fixed offset.
            "utc_offset_minutes": "180",
            "workday_start_minutes": "180",
        }
        with self.conn:
            for key, value in defaults.items():
                self.conn.execute("INSERT OR IGNORE INTO settings(key,value) VALUES(?,?)", (key, value))
        result = defaults.copy()
        # Remove the legacy mouse-button binding; this version uses only F8/F9.
        with self.conn:
            self.conn.execute("DELETE FROM settings WHERE key = 'mouse_button'")
        rows = self.conn.execute("SELECT key,value FROM settings").fetchall()
        for row in rows:
            if row["key"] in defaults:
                result[row["key"]] = row["value"]
        configured = [normalize_shortcut(result[key]) for key in ("dodo_key", "undo_key", "stop_key", "skip_key")]
        if len(configured) != len(set(configured)):
            result.update({key: defaults[key] for key in ("dodo_key", "undo_key", "stop_key", "skip_key")})
        return result

    def _current_offset_minutes(self) -> int:
        try:
            return int(self.settings.get("utc_offset_minutes", "180"))
        except (TypeError, ValueError):
            return 180

    def _current_workday_start_minutes(self) -> int:
        try:
            return max(0, min(1439, int(self.settings.get("workday_start_minutes", "180"))))
        except (TypeError, ValueError):
            return 180

    def _migrate_legacy_local_timestamps(self):
        """One-time migration from older local-naive timestamps to UTC-naive storage.

        Earlier V23 builds used datetime.now() for stored event timestamps. On the
        user's Egypt setup those values represented UTC+3 wall time, despite being
        interpreted as UTC by the workday calculations. Shift timestamp fields back
        three hours once, then standardize all existing history to UTC+3 / 03:00 AM.
        Task durations and manual work-date/time entries are intentionally untouched.
        """
        marker = self.conn.execute("SELECT value FROM settings WHERE key='timestamp_storage_utc_v1'").fetchone()
        if marker and marker["value"] == "1":
            return
        with self.conn:
            for row in self.conn.execute("SELECT id, started_at, ended_at, last_checkpoint_at FROM tasks").fetchall():
                self.conn.execute(
                    "UPDATE tasks SET started_at=?, ended_at=?, last_checkpoint_at=?, utc_offset_minutes=180, workday_start_minutes=180 WHERE id=?",
                    (legacy_local_timestamp_to_utc(row["started_at"], 180),
                     legacy_local_timestamp_to_utc(row["ended_at"], 180),
                     legacy_local_timestamp_to_utc(row["last_checkpoint_at"], 180), row["id"]),
                )
            for row in self.conn.execute("SELECT id, created_at FROM manual_adjustments").fetchall():
                self.conn.execute(
                    "UPDATE manual_adjustments SET created_at=?, utc_offset_minutes=180, workday_start_minutes=180 WHERE id=?",
                    (legacy_local_timestamp_to_utc(row["created_at"], 180), row["id"]),
                )
            self.conn.execute("INSERT INTO settings(key,value) VALUES('utc_offset_minutes','180') ON CONFLICT(key) DO UPDATE SET value='180'")
            self.conn.execute("INSERT INTO settings(key,value) VALUES('workday_start_minutes','180') ON CONFLICT(key) DO UPDATE SET value='180'")
            self.conn.execute("INSERT INTO settings(key,value) VALUES('timestamp_storage_utc_v1','1') ON CONFLICT(key) DO UPDATE SET value='1'")

    def _workday_today(self) -> date:
        return workday_for_timestamp(utc_now(), self._current_offset_minutes(), self._current_workday_start_minutes())

    def open_workday_settings(self):
        """Change workday boundary and UTC offset without changing original timestamps."""
        parent = self.report_window if self.report_window is not None else self.root
        win = tk.Toplevel(parent)
        win.title("Workday Settings")
        win.geometry("510x345")
        win.resizable(False, False)
        win.transient(parent)
        win.grab_set()
        frame = tk.Frame(win, bg="#F6F8FB", padx=16, pady=14)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text="Workday Settings", font=("Segoe UI", 12, "bold"), bg="#F6F8FB", fg="#152238").pack(anchor="w", pady=(0, 10))

        fields = tk.Frame(frame, bg="#F6F8FB")
        fields.pack(fill="x")
        tk.Label(fields, text="Workday starts at", bg="#F6F8FB", fg="#263445", font=("Segoe UI", 9)).grid(row=0, column=0, sticky="w", pady=5)
        boundary = self._current_workday_start_minutes()
        boundary_dt = datetime.combine(date.today(), dt_time.min) + timedelta(minutes=boundary)
        start_var = tk.StringVar(value=boundary_dt.strftime("%I:%M %p").lstrip("0"))
        ttk.Entry(fields, textvariable=start_var, width=15).grid(row=0, column=1, sticky="w", padx=(12, 0), pady=5)

        tk.Label(fields, text="UTC offset", bg="#F6F8FB", fg="#263445", font=("Segoe UI", 9)).grid(row=1, column=0, sticky="w", pady=5)
        offset_var = tk.StringVar(value=f"UTC{self._current_offset_minutes()/60:+g}")
        offsets = [f"UTC{minutes/60:+g}" for minutes in range(-720, 841, 30)]
        ttk.Combobox(fields, textvariable=offset_var, values=offsets, state="readonly", width=12).grid(row=1, column=1, sticky="w", padx=(12, 0), pady=5)

        tk.Label(frame, text="Apply these settings to", bg="#F6F8FB", fg="#263445", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(12, 3))
        apply_var = tk.StringVar(value="all")
        tk.Radiobutton(frame, text="Selected week only, plus future records", variable=apply_var, value="week", bg="#F6F8FB", anchor="w").pack(fill="x")
        tk.Radiobutton(frame, text="All weeks (all history and future records)", variable=apply_var, value="all", bg="#F6F8FB", anchor="w").pack(fill="x")
        tk.Radiobutton(frame, text="New records only (keep all old records unchanged)", variable=apply_var, value="future", bg="#F6F8FB", anchor="w").pack(fill="x")

        def save():
            try:
                parsed_boundary = datetime.strptime(start_var.get().strip().upper(), "%I:%M %p")
                start_minutes = parsed_boundary.hour * 60 + parsed_boundary.minute
                offset_text = offset_var.get().replace("UTC", "")
                offset_hours = float(offset_text)
                offset_minutes = int(round(offset_hours * 60))
                if offset_minutes < -720 or offset_minutes > 840:
                    raise ValueError("UTC offset is outside the supported range.")
            except (ValueError, TypeError):
                messagebox.showerror("Invalid settings", "Enter the workday start as e.g. 3:00 AM and select a valid UTC offset.", parent=win)
                return
            scope = apply_var.get()
            if scope in {"all", "week"}:
                scope_text = "all existing weeks" if scope == "all" else "the selected week only"
                if not messagebox.askyesno("Recalculate history", f"Apply the new workday settings to {scope_text}, and use them for future records? Original event timestamps and task durations will be preserved.", parent=win):
                    return
                if scope == "all":
                    with self.conn:
                        self.conn.execute("UPDATE tasks SET utc_offset_minutes=?, workday_start_minutes=?", (offset_minutes, start_minutes))
                        self.conn.execute("UPDATE manual_adjustments SET utc_offset_minutes=?, workday_start_minutes=?", (offset_minutes, start_minutes))
                else:
                    week_start = None
                    if self.report_week_selector is not None and self.report_week_selector.winfo_exists():
                        week_start = getattr(self, "_week_archive_starts", {}).get(self.report_week_selector.get())
                    if week_start is None:
                        today = self._workday_today()
                        week_start = today - timedelta(days=(today.weekday() + 2) % 7)
                    week_end = week_start + timedelta(days=7)
                    task_rows = self.conn.execute("SELECT id, started_at, ended_at, utc_offset_minutes, workday_start_minutes FROM tasks").fetchall()
                    manual_rows = self.conn.execute("SELECT id, work_date, start_time, utc_offset_minutes, workday_start_minutes FROM manual_adjustments").fetchall()
                    task_ids = []
                    manual_ids = []
                    for record in task_rows:
                        stamp = parse_dt(record["ended_at"]) or parse_dt(record["started_at"])
                        old_offset = int(record["utc_offset_minutes"] if record["utc_offset_minutes"] is not None else self._current_offset_minutes())
                        old_boundary = int(record["workday_start_minutes"] if record["workday_start_minutes"] is not None else self._current_workday_start_minutes())
                        record_day = workday_for_timestamp(stamp, old_offset, old_boundary)
                        if week_start <= record_day < week_end:
                            task_ids.append((offset_minutes, start_minutes, record["id"]))
                    for record in manual_rows:
                        try:
                            record_day = date.fromisoformat(record["work_date"])
                            if record["start_time"]:
                                clock = datetime.strptime(record["start_time"], "%H:%M").time()
                                record_day = workday_for_timestamp(datetime.combine(record_day, clock), 0, int(record["workday_start_minutes"] if record["workday_start_minutes"] is not None else self._current_workday_start_minutes()))
                        except (TypeError, ValueError):
                            continue
                        if week_start <= record_day < week_end:
                            manual_ids.append((offset_minutes, start_minutes, record["id"]))
                    with self.conn:
                        self.conn.executemany("UPDATE tasks SET utc_offset_minutes=?, workday_start_minutes=? WHERE id=?", task_ids)
                        self.conn.executemany("UPDATE manual_adjustments SET utc_offset_minutes=?, workday_start_minutes=? WHERE id=?", manual_ids)
            self._save_setting("utc_offset_minutes", str(offset_minutes))
            self._save_setting("workday_start_minutes", str(start_minutes))
            win.destroy()
            self._update_main_stats()
            self.refresh_report()

        buttons = tk.Frame(frame, bg="#F6F8FB")
        buttons.pack(fill="x", pady=(12, 0))
        tk.Button(buttons, text="Save Settings", command=save, bg=BLUE_BUTTON, fg=WHITE, relief="flat", padx=10, pady=4).pack(side="left", padx=(0, 6))
        tk.Button(buttons, text="Cancel", command=win.destroy, bg="#E2E8F0", fg="#172033", relief="flat", padx=10, pady=4).pack(side="left")

    def _save_setting(self, key: str, value: str):
        self.settings[key] = value
        with self.conn:
            self.conn.execute(
                "INSERT INTO settings(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
                (key, value),
            )

    def _recover_orphaned_tasks(self):
        """If Windows or the app crashed, preserve the last checkpoint but never count downtime."""
        rows = self.conn.execute("SELECT * FROM tasks WHERE status='active'").fetchall()
        if not rows:
            return
        with self.conn:
            for row in rows:
                end_at = row["last_checkpoint_at"] or row["started_at"]
                duration_ms = max(0, int(row["duration_ms"] or 0))
                duration = max(0, int(round(duration_ms / 1000)))
                self.conn.execute(
                    "UPDATE tasks SET status='incomplete', ended_at=?, duration_seconds=?, duration_ms=?, completion_source='app_recovery' WHERE id=?",
                    (end_at, duration, duration_ms, row["id"]),
                )

    def _insert_active_task(self) -> int:
        now = iso_now()
        with self.conn:
            cursor = self.conn.execute(
                "INSERT INTO tasks(started_at, ended_at, duration_seconds, duration_ms, status, completion_source, last_checkpoint_at, utc_offset_minutes, workday_start_minutes) VALUES(?,NULL,0,0,'active',NULL,?,?,?)",
                (now, now, self._current_offset_minutes(), self._current_workday_start_minutes()),
            )
        self.current_task_id = int(cursor.lastrowid)
        self.current_base_elapsed = 0
        self.current_mono_start = time.monotonic()
        self.last_checkpoint_mono = self.current_mono_start
        self.last_display_elapsed = 0
        return self.current_task_id

    def _current_elapsed(self) -> float:
        if self.current_task_id is None:
            return max(0.0, float(self.last_display_elapsed))
        return max(0.0, float(self.current_base_elapsed) + (time.monotonic() - self.current_mono_start))

    def _checkpoint_active_task(self, force: bool = False):
        if self.current_task_id is None:
            return
        now_mono = time.monotonic()
        if not force and now_mono - self.last_checkpoint_mono < 1.0:
            return
        elapsed = self._current_elapsed()
        elapsed_ms = max(0, int(round(elapsed * 1000)))
        elapsed_seconds = max(0, int(round(elapsed_ms / 1000)))
        checkpoint = iso_now()
        with self.conn:
            self.conn.execute(
                "UPDATE tasks SET duration_seconds=?, duration_ms=?, last_checkpoint_at=? WHERE id=? AND status='active'",
                (elapsed_seconds, elapsed_ms, checkpoint, self.current_task_id),
            )
        self.last_checkpoint_mono = now_mono

    def _finish_current_task(self, source: str) -> tuple[int | None, float]:
        if self.current_task_id is None:
            return None, 0.0
        task_id = self.current_task_id
        duration = self._current_elapsed()
        duration_ms = max(0, int(round(duration * 1000)))
        duration_seconds = max(0, int(round(duration_ms / 1000)))
        end_at = iso_now()
        with self.conn:
            self.conn.execute(
                "UPDATE tasks SET ended_at=?, duration_seconds=?, duration_ms=?, status='completed', completion_source=?, last_checkpoint_at=? WHERE id=?",
                (end_at, duration_seconds, duration_ms, source, end_at, task_id),
            )
        self.current_task_id = None
        self.current_base_elapsed = 0
        self.current_mono_start = 0.0
        self.last_display_elapsed = duration
        return task_id, duration

    # ---------- Statistics ----------
    def _all_task_records(self, include_live: bool = True) -> list[dict]:
        rows = self.conn.execute("SELECT * FROM tasks ORDER BY id DESC").fetchall()
        now = utc_now()
        records = []
        for row in rows:
            item = dict(row)
            if include_live and self.current_task_id is not None and item["id"] == self.current_task_id:
                live_elapsed = self._current_elapsed()
                item["duration_seconds"] = max(0, int(round(live_elapsed)))
                item["duration_ms"] = max(0, int(round(live_elapsed * 1000)))
                item["ended_at"] = now.isoformat(timespec="seconds")
                item["_live"] = True
            else:
                item["_live"] = False
            records.append(item)
        return records

    def _statistics(self) -> dict:
        records = self._all_task_records(include_live=True)
        daily_work = defaultdict(int)
        daily_tasks = defaultdict(int)
        daily_task_duration = defaultdict(int)
        total_completed_count = 0
        total_completed_duration = 0
        total_task_targets = 0

        for row in records:
            start = parse_dt(row.get("started_at"))
            end = parse_dt(row.get("ended_at"))
            duration_ms_value = row.get("duration_ms")
            duration_seconds_value = row.get("duration_seconds") or 0
            if duration_ms_value is None or (int(duration_ms_value or 0) == 0 and float(duration_seconds_value) > 0):
                duration = max(0.0, float(duration_seconds_value))
            else:
                duration = max(0.0, float(duration_ms_value or 0) / 1000.0)
            status = row.get("status")
            is_live_task = status == "active" and bool(row.get("_live"))
            if end is None and is_live_task:
                end = utc_now()

            # Work-time totals include only completed tasks. A running attempt is
            # shown by the main stopwatch, but is not added to today's work time
            # until Dodo or Stop completes it. This also guarantees that closing
            # the app mid-task cannot leave stray seconds in the daily statistics.
            if status == "completed" and row.get("ended_at"):
                offset_minutes = row.get("utc_offset_minutes")
                boundary_minutes = row.get("workday_start_minutes")
                if offset_minutes is None:
                    offset_minutes = self._current_offset_minutes()
                if boundary_minutes is None:
                    boundary_minutes = self._current_workday_start_minutes()
                for day, seconds in split_duration_by_workday(start, end, duration, int(offset_minutes), int(boundary_minutes)).items():
                    daily_work[day] += seconds

            if status == "completed" and row.get("ended_at"):
                completed_at = parse_dt(row["ended_at"])
                if completed_at:
                    offset_minutes = row.get("utc_offset_minutes")
                    boundary_minutes = row.get("workday_start_minutes")
                    if offset_minutes is None:
                        offset_minutes = self._current_offset_minutes()
                    if boundary_minutes is None:
                        boundary_minutes = self._current_workday_start_minutes()
                    day = workday_for_timestamp(completed_at, int(offset_minutes), int(boundary_minutes))
                    daily_tasks[day] += 1
                    daily_task_duration[day] += duration
                    total_completed_count += 1
                    total_completed_duration += duration

        manual_rows = self.conn.execute(
            "SELECT * FROM manual_adjustments ORDER BY id DESC"
        ).fetchall()
        for adjustment in manual_rows:
            try:
                work_day = date.fromisoformat(adjustment["work_date"])
            except (TypeError, ValueError):
                continue
            added_tasks = max(0, int(adjustment["task_count"] or 0))
            added_seconds = max(0, int(adjustment["total_ms"] or 0)) / 1000.0
            offset_minutes = adjustment["utc_offset_minutes"] if "utc_offset_minutes" in adjustment.keys() else None
            boundary_minutes = adjustment["workday_start_minutes"] if "workday_start_minutes" in adjustment.keys() else None
            if offset_minutes is None:
                offset_minutes = self._current_offset_minutes()
            if boundary_minutes is None:
                boundary_minutes = self._current_workday_start_minutes()
            if adjustment["start_time"]:
                try:
                    start_clock = datetime.strptime(adjustment["start_time"], "%H:%M").time()
                    work_day = workday_for_timestamp(datetime.combine(work_day, start_clock), 0, int(boundary_minutes))
                except (TypeError, ValueError):
                    pass
            daily_work[work_day] += added_seconds
            daily_tasks[work_day] += added_tasks
            daily_task_duration[work_day] += added_seconds
            total_completed_count += added_tasks
            total_completed_duration += added_seconds

        all_days = sorted(set(daily_work) | set(daily_tasks), reverse=True)
        daily_rows = []
        for day in all_days:
            work_precise = max(0.0, float(daily_work.get(day, 0)))
            work = int(round(work_precise))
            tasks = int(daily_tasks.get(day, 0))
            target = int(math.ceil(work_precise / TASK_TARGET_SECONDS)) if work_precise else 0
            avg = (daily_task_duration[day] / tasks) if tasks else None
            if work > 0 or tasks > 0:
                daily_rows.append({
                    "date": day,
                    "work_seconds": work,
                    "required_seconds": DAY_TARGET_SECONDS,
                    "actual_tasks": tasks,
                    "target_tasks": target,
                    "difference": tasks - target,
                    "avg_seconds": avg,
                    "completed_duration": int(daily_task_duration.get(day, 0)),
                })
        days_worked = len(daily_rows)
        total_work = sum(r["work_seconds"] for r in daily_rows)
        total_required = days_worked * DAY_TARGET_SECONDS
        total_target_tasks = sum(r["target_tasks"] for r in daily_rows)
        total_difference = total_completed_count - total_target_tasks
        total_avg = total_completed_duration / total_completed_count if total_completed_count else None
        today = self._workday_today()
        today_work_precise = max(0.0, float(daily_work.get(today, 0)))
        today_work = int(round(today_work_precise))
        today_tasks = int(daily_tasks.get(today, 0))
        today_avg = (daily_task_duration[today] / today_tasks) if today_tasks else None
        today_target = int(math.ceil(today_work_precise / TASK_TARGET_SECONDS)) if today_work_precise else 0
        return {
            "records": records,
            "daily_rows": daily_rows,
            "days_worked": days_worked,
            "total_work_seconds": total_work,
            "total_required_seconds": total_required,
            "total_tasks": total_completed_count,
            "total_target_tasks": total_target_tasks,
            "total_difference": total_difference,
            "total_avg_seconds": total_avg,
            "today_work_seconds": today_work,
            "today_work_precise": today_work_precise,
            "today_tasks": today_tasks,
            "today_avg_seconds": today_avg,
            "today_target_tasks": today_target,
        }

    # ---------- Main UI ----------
    def _build_ui(self):
        self.root.configure(bg=BG_BLUE)
        self.main = tk.Frame(self.root, bg=BG_BLUE, padx=6, pady=1)
        self.main.pack(fill="both", expand=True)

        header = tk.Frame(self.main, bg=BG_BLUE)
        header.pack(fill="x")
        self.title_label = tk.Label(
            header, text="Task Timer", bg=BG_BLUE, fg=WHITE,
            font=("Segoe UI", 11, "bold"), anchor="w"
        )
        self.title_label.pack(side="left", expand=True, fill="x")
        self.report_button = tk.Button(
            header, text="Log", command=self.open_report,
            bg="#FFFFFF", fg=NEUTRAL_TEXT, relief="flat",
            activebackground="#EAF2F8", font=("Segoe UI", 8, "bold"),
            padx=5, pady=3, cursor="hand2"
        )
        self.report_button.pack(side="right", padx=(0, 0))
        self.settings_button = tk.Button(
            header, text="⚙", command=self.open_settings,
            bg="#FFFFFF", fg=NEUTRAL_TEXT, relief="flat",
            activebackground="#EAF2F8", font=("Segoe UI", 10, "bold"),
            width=2, pady=1, cursor="hand2"
        )
        # Leave a clear gap between Settings (shortcut configuration) and Log.
        self.settings_button.pack(side="right", padx=(0, 8))

        self.timer_label = tk.Label(
            self.main, text="00:00.00", bg=BG_BLUE, fg=WHITE,
            font=("Segoe UI", 20, "bold"), anchor="center"
        )
        self.timer_label.pack(fill="x", pady=(2, 1))

        self.pacing_frame = tk.Frame(self.main, bg="#0F172A", padx=4, pady=1)
        self.pacing_frame.pack(fill="x", pady=(0, 2))
        self.pacing_label = tk.Label(
            self.pacing_frame, text="Pacing: On target", bg="#0F172A", fg=WHITE,
            font=("Segoe UI", 8, "bold"), anchor="center"
        )
        self.pacing_label.pack(fill="x")

        stats = tk.Frame(self.main, bg=BG_BLUE)
        stats.pack(fill="x", pady=(0, 2))
        # Work Time and Average are on the left; task counts are on the right.
        self.today_work_label = self._make_stat_card(stats, 0, 0, "Work Time Today", "0 min")
        self.today_tasks_label = self._make_stat_card(stats, 0, 1, "Tasks Today", "0")
        self.today_avg_label = self._make_stat_card(stats, 1, 0, "Avg. Task Time", "N/A")
        self.today_target_label = self._make_stat_card(stats, 1, 1, "Target Tasks", "0")

        buttons = tk.Frame(self.main, bg=BG_BLUE)
        buttons.pack(fill="x", pady=(1, 0))
        self.dodo_button = tk.Button(
            buttons, text="▶ Start +1", command=self.handle_dodo,
            bg=BLUE_BUTTON, fg=WHITE, activebackground="#064B87",
            activeforeground=WHITE, relief="flat", font=("Segoe UI", 9, "bold"),
            padx=5, pady=3, cursor="hand2"
        )
        self.dodo_button.pack(side="left", expand=True, fill="x", padx=(0, 3))
        self.undo_button = tk.Button(
            buttons, text="↶", command=self.handle_undo,
            bg=NEUTRAL_BUTTON, fg=NEUTRAL_TEXT, activebackground="#D6DEE8",
            relief="flat", font=("Segoe UI Symbol", 11, "bold"),
            width=2, pady=2, cursor="hand2"
        )
        self.undo_button.pack(side="left", padx=2)
        self.skip_button = tk.Button(
            buttons, text="Skip", command=self.handle_skip,
            bg="#64748B", fg=WHITE, activebackground="#475569",
            activeforeground=WHITE, relief="flat", font=("Segoe UI", 8, "bold"),
            padx=4, pady=3, cursor="hand2"
        )
        self.skip_button.pack(side="right", padx=(0, 3))
        self.stop_button = tk.Button(
            buttons, text="Stop", command=self.handle_stop,
            bg=RED_BUTTON, fg=WHITE, activebackground="#7F1D1D",
            activeforeground=WHITE, relief="flat", font=("Segoe UI", 8, "bold"),
            padx=5, pady=3, cursor="hand2"
        )
        self.stop_button.pack(side="right")

        self._update_main_stats()
        self._set_colors(BG_BLUE, WHITE)

    def _make_stat_card(self, parent, row, col, title, value):
        card = tk.Frame(parent, bg="#FFFFFF", padx=3, pady=1, highlightthickness=0)
        card.grid(row=row, column=col, sticky="nsew", padx=2, pady=1)
        parent.grid_columnconfigure(col, weight=1, uniform="stats")
        label = tk.Label(card, text=value, bg="#FFFFFF", fg="#172033", font=("Segoe UI", 11, "bold"), anchor="center")
        label.pack(fill="x")
        caption = tk.Label(card, text=title, bg="#FFFFFF", fg="#536273", font=("Segoe UI", 8), anchor="center")
        caption.pack(fill="x", pady=(1, 0))
        return label

    def _set_colors(self, bg: str, fg: str):
        self.current_bg = bg
        try:
            self.root.configure(bg=bg)
            self.main.configure(bg=bg)
            for child in self.main.winfo_children():
                # Header, stats and buttons have their own controlled button/card colors.
                if isinstance(child, tk.Frame):
                    child.configure(bg=bg)
                    for grandchild in child.winfo_children():
                        if isinstance(grandchild, tk.Frame) and grandchild not in (getattr(self, "main", None),):
                            # Only the cards themselves retain white backgrounds.
                            if grandchild.winfo_children():
                                pass
                        elif isinstance(grandchild, tk.Label):
                            grandchild.configure(bg=bg, fg=fg)
                elif isinstance(child, tk.Label):
                    child.configure(bg=bg, fg=fg)
            self.title_label.configure(bg=bg, fg=fg)
            self.timer_label.configure(bg=bg, fg=fg)
            if hasattr(self, "pacing_frame") and hasattr(self, "pacing_label"):
                self.pacing_frame.configure(bg="#0F172A")
                self.pacing_label.configure(bg="#0F172A", fg=WHITE)
        except tk.TclError:
            pass

    def _hotkey_description(self):
        return f"{self.settings['dodo_key']} : Start +1   •   {self.settings['undo_key']} : Undo"

    def _update_status(self, text: str):
        # Status/instruction text was intentionally removed from the compact UI.
        # Keep this no-op so internal actions remain readable without displaying prompts.
        return

    def _update_main_stats(self):
        stats = self._statistics()
        self.today_tasks_label.configure(text=str(stats["today_tasks"]))
        self.today_work_label.configure(text=format_duration(stats["today_work_seconds"], include_seconds=False))
        self.today_target_label.configure(text=str(stats["today_target_tasks"]))
        self.today_avg_label.configure(text=format_average(stats["today_avg_seconds"]))
        self._update_pacing_label(stats)

    def _update_pacing_label(self, stats: dict | None = None):
        """Show the current time bank against the four-minute-per-task benchmark."""
        if stats is None:
            stats = self._statistics()
        work_seconds = float(stats.get("today_work_precise", stats.get("today_work_seconds", 0)))
        task_count = int(stats.get("today_tasks", 0))
        live_elapsed_today = 0.0

        # Count the live attempt only for this temporary pacing readout; it is not
        # added to official Work Time until Dodo or Stop completes the task.
        if self.current_task_id is not None:
            elapsed = self._current_elapsed()
            active_row = self.conn.execute(
                "SELECT started_at FROM tasks WHERE id=? AND status='active'",
                (self.current_task_id,),
            ).fetchone()
            started = parse_dt(active_row["started_at"]) if active_row else None
            today = self._workday_today()
            active_row = self.conn.execute("SELECT utc_offset_minutes, workday_start_minutes FROM tasks WHERE id=?", (self.current_task_id,)).fetchone()
            active_offset = int(active_row["utc_offset_minutes"] if active_row and active_row["utc_offset_minutes"] is not None else self._current_offset_minutes())
            active_boundary = int(active_row["workday_start_minutes"] if active_row and active_row["workday_start_minutes"] is not None else self._current_workday_start_minutes())
            if started is None or workday_for_timestamp(started, active_offset, active_boundary) == today:
                live_elapsed_today = elapsed
            elif started and workday_for_timestamp(started, active_offset, active_boundary) < today:
                today_start_utc = datetime.combine(today, dt_time.min) + timedelta(minutes=active_boundary-active_offset)
                live_elapsed_today = min(elapsed, max(0.0, (utc_now() - today_start_utc).total_seconds()))

        balance = int(round(task_count * TASK_TARGET_SECONDS - work_seconds - live_elapsed_today))
        if balance > 0:
            minutes, seconds = divmod(balance, 60)
            text = f"Pacing: +{minutes}:{seconds:02d} min ahead"
        elif balance < 0:
            minutes, seconds = divmod(abs(balance), 60)
            text = f"Pacing: {minutes}:{seconds:02d} min behind"
        else:
            text = "Pacing: On target"
        self.pacing_frame.configure(bg="#0F172A")
        self.pacing_label.configure(text=text, bg="#0F172A", fg=WHITE)

    def _tick(self):
        if not self.root.winfo_exists():
            return
        # pynput callbacks run on a background thread; execute Tk operations only
        # from Tk's own event loop to avoid silent cross-thread failures.
        while True:
            try:
                action = self.hotkey_queue.get_nowait()
            except queue.Empty:
                break
            handlers = {
                "dodo": self.handle_dodo,
                "undo": self.handle_undo,
                "stop": self.handle_stop,
                "skip": self.handle_skip,
            }
            handler = handlers.get(action)
            if handler:
                handler()
        elapsed = self._current_elapsed() if self.current_task_id is not None else self.last_display_elapsed
        bg, fg = color_for_elapsed(elapsed)
        if bg != self.current_bg:
            self._set_colors(bg, fg)
        self.timer_label.configure(text=format_stopwatch(elapsed))
        if self.current_task_id is not None:
            self._checkpoint_active_task()
        now = time.monotonic()
        if now - self.last_stats_update >= 1.25:
            self._update_main_stats()
            self.last_stats_update = now
        # Refresh fast enough for the two decimal places shown by Windows Stopwatch.
        self.root.after(10, self._tick)

    # ---------- Task actions ----------
    def handle_dodo(self):
        if self.current_task_id is None:
            new_id = self._insert_active_task()
            # Undo can reverse the start itself, which is useful after an accidental keypress.
            self.last_undo = {"kind": "start", "new_id": new_id}
            self._update_main_stats()
            self.refresh_report()
            return

        completed_id, duration = self._finish_current_task("dodo")
        new_id = self._insert_active_task()
        self.last_undo = {
            "kind": "dodo",
            "completed_id": completed_id,
            "new_id": new_id,
            "duration": duration,
        }
        self._update_main_stats()
        self.refresh_report()

    def handle_stop(self):
        if self.current_task_id is None:
            return
        task_id, duration = self._finish_current_task("stop")
        self.last_undo = {"kind": "stop", "completed_id": task_id, "duration": duration}
        self._update_main_stats()
        self.refresh_report()

    def handle_skip(self):
        """Discard the active attempt without adding it to task or work-time statistics."""
        if self.current_task_id is None:
            return
        task_id = self.current_task_id
        row = self.conn.execute("SELECT * FROM tasks WHERE id=?", (task_id,)).fetchone()
        if row is None:
            # Defensive recovery; do not invent a logged task when its row is missing.
            duration = self._current_elapsed()
            snapshot = None
        else:
            snapshot = dict(row)
            duration = self._current_elapsed()
            duration_ms = max(0, int(round(duration * 1000)))
            snapshot["duration_ms"] = duration_ms
            snapshot["duration_seconds"] = max(0, int(round(duration_ms / 1000)))

        if snapshot is not None:
            with self.conn:
                self.conn.execute("DELETE FROM tasks WHERE id=? AND status='active'", (task_id,))
        self.last_undo = {"kind": "skip", "task_snapshot": snapshot, "duration": duration}
        self.current_task_id = None
        self.current_base_elapsed = 0.0
        self.current_mono_start = 0.0
        self.last_checkpoint_mono = 0.0
        self.last_display_elapsed = 0.0
        self._update_main_stats()
        self.refresh_report()

    def handle_undo(self):
        """Undo the latest action: start, Dodo, Stop, Skip, or manual adjustment."""
        action = self.last_undo
        if not action:
            return

        now = iso_now()
        now_mono = time.monotonic()
        kind = action.get("kind")

        if kind == "start":
            if self.current_task_id != action.get("new_id"):
                return
            with self.conn:
                self.conn.execute("DELETE FROM tasks WHERE id=? AND status='active'", (self.current_task_id,))
            self.current_task_id = None
            self.current_base_elapsed = 0.0
            self.current_mono_start = 0.0
            self.last_checkpoint_mono = 0.0
            self.last_display_elapsed = 0.0

        elif kind == "dodo":
            if self.current_task_id != action.get("new_id"):
                return
            previous_id = action.get("completed_id")
            if previous_id is None:
                return
            # Preserve the existing V7 behavior: elapsed time since Dodo is applied
            # to the restored task, so an accidental Dodo press does not lose time.
            combined_duration = float(action.get("duration", 0.0)) + self._current_elapsed()
            duration_ms = max(0, int(round(combined_duration * 1000)))
            duration_seconds = max(0, int(round(duration_ms / 1000)))
            with self.conn:
                self.conn.execute("DELETE FROM tasks WHERE id=? AND status='active'", (self.current_task_id,))
                self.conn.execute(
                    "UPDATE tasks SET status='active', ended_at=NULL, duration_seconds=?, duration_ms=?, completion_source=NULL, last_checkpoint_at=? WHERE id=?",
                    (duration_seconds, duration_ms, now, previous_id),
                )
            self.current_task_id = previous_id
            self.current_base_elapsed = combined_duration
            self.current_mono_start = now_mono
            self.last_checkpoint_mono = now_mono
            self.last_display_elapsed = combined_duration

        elif kind == "stop":
            if self.current_task_id is not None:
                return
            task_id = action.get("completed_id")
            if task_id is None:
                return
            duration = float(action.get("duration", 0.0))
            duration_ms = max(0, int(round(duration * 1000)))
            duration_seconds = max(0, int(round(duration_ms / 1000)))
            with self.conn:
                self.conn.execute(
                    "UPDATE tasks SET status='active', ended_at=NULL, duration_seconds=?, duration_ms=?, completion_source=NULL, last_checkpoint_at=? WHERE id=?",
                    (duration_seconds, duration_ms, now, task_id),
                )
            self.current_task_id = task_id
            self.current_base_elapsed = duration
            self.current_mono_start = now_mono
            self.last_checkpoint_mono = now_mono
            self.last_display_elapsed = duration

        elif kind == "manual_adjustment":
            adjustment_id = action.get("adjustment_id")
            if adjustment_id is None:
                return
            with self.conn:
                self.conn.execute("DELETE FROM manual_adjustments WHERE id=?", (adjustment_id,))

        elif kind == "skip":
            if self.current_task_id is not None:
                return
            snapshot = action.get("task_snapshot")
            if snapshot is not None:
                task_id = int(snapshot["id"])
                duration_ms = max(0, int(snapshot.get("duration_ms") or 0))
                duration_seconds = max(0, int(round(duration_ms / 1000)))
                started_at = snapshot.get("started_at") or now
                with self.conn:
                    self.conn.execute(
                        "INSERT INTO tasks(id, started_at, ended_at, duration_seconds, duration_ms, status, completion_source, last_checkpoint_at, utc_offset_minutes, workday_start_minutes) "
                        "VALUES(?, ?, NULL, ?, ?, 'active', NULL, ?, ?, ?) "
                        "ON CONFLICT(id) DO UPDATE SET started_at=excluded.started_at, ended_at=NULL, duration_seconds=excluded.duration_seconds, duration_ms=excluded.duration_ms, status='active', completion_source=NULL, last_checkpoint_at=excluded.last_checkpoint_at, utc_offset_minutes=excluded.utc_offset_minutes, workday_start_minutes=excluded.workday_start_minutes",
                        (task_id, started_at, duration_seconds, duration_ms, now, snapshot.get("utc_offset_minutes", self._current_offset_minutes()), snapshot.get("workday_start_minutes", self._current_workday_start_minutes())),
                    )
                self.current_task_id = task_id
                self.current_base_elapsed = duration_ms / 1000.0
                self.current_mono_start = now_mono
                self.last_checkpoint_mono = now_mono
                self.last_display_elapsed = self.current_base_elapsed
            else:
                # No database row existed before Skip; safely return to idle.
                self.last_display_elapsed = 0.0

        else:
            return

        self.last_undo = None
        self._update_main_stats()
        self.refresh_report()

    # ---------- Global keyboard shortcuts ----------
    def _bind_local_hotkeys(self):
        """Bind configured shortcuts when the app window itself is focused."""
        for sequence in getattr(self, "_local_hotkey_sequences", []):
            try:
                self.root.unbind_all(sequence)
            except tk.TclError:
                pass
        self._local_hotkey_sequences = []
        for setting_key, action in (
            ("dodo_key", "dodo"), ("undo_key", "undo"),
            ("stop_key", "stop"), ("skip_key", "skip"),
        ):
            sequence = tkinter_shortcut_sequence(self.settings.get(setting_key, "F8"))
            try:
                self.root.bind_all(sequence, lambda event, act=action: self._queue_hotkey(act))
                self._local_hotkey_sequences.append(sequence)
            except tk.TclError:
                pass

    def _start_global_listeners(self):
        """Prefer a low-level Windows keyboard hook; use pynput if unavailable."""
        native_ok = False
        hook_ok = False
        self.hotkey_status_message = "Could not enable global hotkeys."
        if os.name == "nt":
            self._native_hotkey_ready = threading.Event()
            self._native_hotkey_error = None
            self._native_hotkey_thread_id = None
            self.native_hotkey_thread = threading.Thread(
                target=self._windows_hotkey_loop, daemon=True, name="DodoGlobalHotkeys"
            )
            self.native_hotkey_thread.start()
            if self._native_hotkey_ready.wait(timeout=2.0) and self._native_hotkey_error is None:
                native_ok = True

        # Fallback works globally with pynput, but Windows hook is preferred because it
        # can consume configured key combinations before a browser handles them.
        if not native_ok and PYNPUT_AVAILABLE:
            try:
                self._pynput_pressed_keys = set()
                self.keyboard_listener = keyboard.Listener(
                    on_press=self._on_global_key,
                    on_release=self._on_global_key_release,
                )
                self.keyboard_listener.daemon = True
                self.keyboard_listener.start()
                hook_ok = True
            except Exception as exc:
                self.hotkey_status_message = f"Could not enable hotkeys: {exc}"

        if native_ok:
            self.hotkey_status_message = "Global keyboard shortcuts enabled."
        elif hook_ok:
            self.hotkey_status_message = "Global keyboard shortcuts enabled using the fallback method."
        elif os.name == "nt" and self._native_hotkey_error:
            self.hotkey_status_message = f"Could not enable hotkeys: {self._native_hotkey_error}"
        self._update_status(self.hotkey_status_message)

    def _windows_hotkey_loop(self):
        """Install a low-level Windows hook for configured function-key shortcuts."""
        import ctypes
        from ctypes import wintypes

        user32 = ctypes.WinDLL("user32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        WH_KEYBOARD_LL = 13
        HC_ACTION = 0
        WM_KEYDOWN = 0x0100
        WM_SYSKEYDOWN = 0x0104
        WM_KEYUP = 0x0101
        WM_SYSKEYUP = 0x0105
        virtual_keys = {f"F{i}": 0x70 + (i - 1) for i in range(1, 13)}
        shortcut_settings = (
            ("dodo_key", "dodo"), ("undo_key", "undo"),
            ("stop_key", "stop"), ("skip_key", "skip"),
        )
        selected: dict[int, list[tuple[str, bool]]] = {}
        for setting_key, action in shortcut_settings:
            needs_ctrl, key_name = shortcut_parts(self.settings.get(setting_key, ""))
            vk = virtual_keys.get(key_name)
            if vk is not None:
                selected.setdefault(vk, []).append((action, needs_ctrl))

        class KBDLLHOOKSTRUCT(ctypes.Structure):
            _fields_ = [
                ("vkCode", wintypes.DWORD),
                ("scanCode", wintypes.DWORD),
                ("flags", wintypes.DWORD),
                ("time", wintypes.DWORD),
                ("dwExtraInfo", ctypes.c_size_t),
            ]

        LRESULT = ctypes.c_ssize_t
        HOOKPROC = ctypes.WINFUNCTYPE(LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
        user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, ctypes.c_void_p, wintypes.DWORD]
        user32.SetWindowsHookExW.restype = ctypes.c_void_p
        user32.CallNextHookEx.argtypes = [ctypes.c_void_p, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
        user32.CallNextHookEx.restype = LRESULT
        user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
        user32.UnhookWindowsHookEx.restype = wintypes.BOOL
        user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), ctypes.c_void_p, wintypes.UINT, wintypes.UINT]
        user32.GetMessageW.restype = ctypes.c_int
        user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
        user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
        user32.DispatchMessageW.restype = LRESULT
        kernel32.GetCurrentThreadId.restype = wintypes.DWORD
        kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
        kernel32.GetModuleHandleW.restype = ctypes.c_void_p

        hook_handle = None
        pressed_keys: set[int] = set()
        consumed_keys: set[int] = set()
        ctrl_codes = {0x11, 0xA2, 0xA3}
        msg = wintypes.MSG()

        def hook_callback(n_code, w_param, l_param):
            try:
                if n_code == HC_ACTION:
                    info = ctypes.cast(l_param, ctypes.POINTER(KBDLLHOOKSTRUCT)).contents
                    vk = int(info.vkCode)
                    if w_param in (WM_KEYDOWN, WM_SYSKEYDOWN):
                        already_down = vk in pressed_keys
                        pressed_keys.add(vk)
                        if vk in consumed_keys:
                            return 1  # swallow repeats for an assigned shortcut
                        if not already_down and vk in selected:
                            ctrl_down = any(code in pressed_keys for code in ctrl_codes)
                            for action, needs_ctrl in selected[vk]:
                                if ctrl_down == needs_ctrl:
                                    consumed_keys.add(vk)
                                    self._queue_hotkey(action)
                                    return 1  # prevent F10/F11/browser side effects
                    elif w_param in (WM_KEYUP, WM_SYSKEYUP):
                        was_consumed = vk in consumed_keys
                        consumed_keys.discard(vk)
                        pressed_keys.discard(vk)
                        if was_consumed:
                            return 1
            except Exception:
                pass
            return user32.CallNextHookEx(hook_handle, n_code, w_param, l_param)

        callback_ref = HOOKPROC(hook_callback)
        try:
            hook_handle = user32.SetWindowsHookExW(
                WH_KEYBOARD_LL, callback_ref, kernel32.GetModuleHandleW(None), 0
            )
            if not hook_handle:
                self._native_hotkey_error = f"Windows error {ctypes.get_last_error()}"
                self._native_hotkey_ready.set()
                return

            self._native_hotkey_thread_id = kernel32.GetCurrentThreadId()
            self._native_hotkey_error = None
            self._native_hotkey_ready.set()
            while True:
                result = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if result <= 0:
                    break
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        except Exception as exc:
            self._native_hotkey_error = str(exc)
            self._native_hotkey_ready.set()
        finally:
            if hook_handle:
                try:
                    user32.UnhookWindowsHookEx(hook_handle)
                except Exception:
                    pass
            self._native_hotkey_thread_id = None

    def _queue_hotkey(self, action: str):
        now = time.monotonic()
        with self._hotkey_lock:
            if now - self._hotkey_last_fired.get(action, 0.0) < 0.35:
                return
            self._hotkey_last_fired[action] = now
        self.hotkey_queue.put(action)

    @staticmethod
    def _pynput_key_name(key) -> str:
        name = getattr(key, "name", None)
        if not name:
            name = str(key).replace("Key.", "").replace("<", "").replace(">", "")
        name = str(name).upper()
        if name in {"CTRL", "CTRL_L", "CTRL_R", "CONTROL", "CONTROL_L", "CONTROL_R"}:
            return "CTRL"
        return name

    def _on_global_key(self, key):
        try:
            name = self._pynput_key_name(key)
            already_down = name in self._pynput_pressed_keys
            self._pynput_pressed_keys.add(name)
            if already_down:
                return
            ctrl_down = "CTRL" in self._pynput_pressed_keys
            for setting_key, action in (
                ("dodo_key", "dodo"), ("undo_key", "undo"),
                ("stop_key", "stop"), ("skip_key", "skip"),
            ):
                needs_ctrl, expected_name = shortcut_parts(self.settings.get(setting_key, ""))
                if name == expected_name and ctrl_down == needs_ctrl:
                    self._queue_hotkey(action)
                    break
        except Exception:
            pass

    def _on_global_key_release(self, key):
        try:
            self._pynput_pressed_keys.discard(self._pynput_key_name(key))
        except Exception:
            pass

    def _update_hotkey_label(self):
        # Keyboard shortcuts no longer take up space in the main window.
        return

    def _stop_global_listeners(self):
        native_thread = getattr(self, "native_hotkey_thread", None)
        native_thread_id = getattr(self, "_native_hotkey_thread_id", None)
        if native_thread is not None and native_thread.is_alive() and native_thread_id:
            try:
                import ctypes
                from ctypes import wintypes
                user32 = ctypes.WinDLL("user32", use_last_error=True)
                user32.PostThreadMessageW.argtypes = [wintypes.DWORD, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
                user32.PostThreadMessageW.restype = wintypes.BOOL
                user32.PostThreadMessageW(native_thread_id, 0x0012, 0, 0)
                native_thread.join(timeout=1.0)
            except Exception:
                pass
        self.native_hotkey_thread = None
        if self.keyboard_listener is not None:
            try:
                self.keyboard_listener.stop()
                self.keyboard_listener.join(timeout=0.5)
            except Exception:
                pass
            self.keyboard_listener = None

    def _restart_global_listeners(self):
        self._stop_global_listeners()
        self._bind_local_hotkeys()
        self._start_global_listeners()

    # ---------- Settings ----------
    def open_settings(self):
        win = tk.Toplevel(self.root)
        win.title("Task Timer Settings")
        win.geometry("365x370")
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()
        frame = tk.Frame(win, padx=14, pady=12, bg="#F6F8FB")
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Settings", font=("Segoe UI", 12, "bold"), bg="#F6F8FB", fg="#152238").pack(anchor="w", pady=(0, 8))
        topmost_var = tk.BooleanVar(value=self.settings.get("topmost", "1") == "1")
        tk.Checkbutton(
            frame, text="Keep window on top", variable=topmost_var,
            bg="#F6F8FB", fg="#152238", activebackground="#F6F8FB",
            font=("Segoe UI", 9), anchor="w", justify="left"
        ).pack(fill="x", anchor="w", pady=(0, 8))

        shortcut_values = [f"F{i}" for i in range(1, 13)] + [f"Ctrl+F{i}" for i in range(1, 13)]
        shortcut_vars = {}
        shortcut_specs = (
            ("dodo_key", "Start +1 shortcut", "F8"),
            ("undo_key", "Undo shortcut", "F9"),
            ("stop_key", "Stop shortcut", "Ctrl+F10"),
            ("skip_key", "Skip shortcut", "Ctrl+F11"),
        )
        for setting_key, label_text, default in shortcut_specs:
            row = tk.Frame(frame, bg="#F6F8FB")
            row.pack(fill="x", pady=3)
            tk.Label(row, text=label_text, bg="#F6F8FB", font=("Segoe UI", 9), width=17, anchor="w").pack(side="left")
            variable = tk.StringVar(value=self.settings.get(setting_key, default))
            ttk.Combobox(
                row, textvariable=variable, values=shortcut_values,
                state="readonly", width=12, justify="center"
            ).pack(side="left")
            shortcut_vars[setting_key] = variable

        status_color = "#166534" if "enabled" in getattr(self, "hotkey_status_message", "").lower() else "#B91C1C"
        tk.Label(
            frame,
            text=f"Global hotkey status: {getattr(self, 'hotkey_status_message', 'Unknown')}",
            bg="#F6F8FB", fg=status_color, font=("Segoe UI", 8), wraplength=330, justify="left"
        ).pack(anchor="w", pady=(8, 5))

        def save_settings():
            values = {key: variable.get() for key, variable in shortcut_vars.items()}
            normalized = [normalize_shortcut(values[key]) for key, _, _ in shortcut_specs]
            if len(normalized) != len(set(normalized)):
                messagebox.showerror("Duplicate shortcuts", "Choose a different shortcut for each action.", parent=win)
                return
            self._save_setting("topmost", "1" if topmost_var.get() else "0")
            for key, value in values.items():
                self._save_setting(key, value)
            self._apply_topmost()
            self._update_hotkey_label()
            self._restart_global_listeners()
            win.destroy()

        button_row = tk.Frame(frame, bg="#F6F8FB")
        button_row.pack(fill="x", pady=(6, 0))
        tk.Button(button_row, text="Save Settings", command=save_settings, bg=BLUE_BUTTON, fg=WHITE,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=10, pady=5, cursor="hand2").pack(side="left", padx=(0, 6))
        tk.Button(button_row, text="Close", command=win.destroy, bg="#E2E8F0", fg="#172033",
                  relief="flat", font=("Segoe UI", 9), padx=10, pady=5, cursor="hand2").pack(side="left")

    def _apply_topmost(self):
        try:
            self.root.attributes("-topmost", self.settings.get("topmost", "1") == "1")
        except tk.TclError:
            pass

    # ---------- Reports and export ----------
    def open_report(self):
        if self.report_window is not None and self.report_window.winfo_exists():
            self.report_window.deiconify()
            self.report_window.lift()
            self.refresh_report()
            return
        win = tk.Toplevel(self.root)
        self.report_window = win
        win.title("Work & Productivity Log — Task Timer")
        win.geometry("800x465")
        win.minsize(720, 405)
        win.configure(bg="#F4F7FB")
        win.transient(self.root)
        win.protocol("WM_DELETE_WINDOW", self._close_report)

        top = tk.Frame(win, bg="#F4F7FB", padx=10, pady=8)
        top.pack(fill="x")
        tk.Label(top, text="Work & Productivity Log", font=("Segoe UI", 13, "bold"), bg="#F4F7FB", fg="#152238").pack(anchor="w")
        self.report_summary_frame = tk.Frame(top, bg="#F4F7FB")
        self.report_summary_frame.pack(fill="x", pady=(6, 0), anchor="w")

        body = tk.Frame(win, bg="#F4F7FB", padx=8)
        body.pack(fill="both", expand=True)
        notebook = ttk.Notebook(body)
        self.report_notebook = notebook
        notebook.pack(fill="both", expand=True)
        daily_tab = tk.Frame(notebook, bg="#FFFFFF")
        tasks_tab = tk.Frame(notebook, bg="#FFFFFF")
        manual_tab = tk.Frame(notebook, bg="#FFFFFF")
        weekly_tab = tk.Frame(notebook, bg="#FFFFFF")
        notebook.add(daily_tab, text="Daily Statistics")
        notebook.add(tasks_tab, text="Task Details")
        notebook.add(manual_tab, text="Manual Adjustments")
        notebook.add(weekly_tab, text="Weekly Archive")

        week_top = tk.Frame(weekly_tab, bg="#FFFFFF", padx=10, pady=10)
        week_top.pack(fill="x")
        tk.Label(week_top, text="Saved weeks (Saturday 3:00 AM UTC+3 to next Saturday 3:00 AM UTC+3; shown Saturday–Friday):",
                 bg="#FFFFFF", fg="#172033", font=("Segoe UI", 9, "bold")).pack(side="left")
        self.report_week_selector = ttk.Combobox(week_top, state="readonly", width=34)
        self.report_week_selector.pack(side="left", padx=(8, 0))
        self.report_week_selector.bind("<<ComboboxSelected>>", lambda _event: self._refresh_week_archive())
        self.report_week_summary = tk.Label(weekly_tab, text="", bg="#FFFFFF", fg="#465568",
                                            font=("Segoe UI", 9), anchor="w", justify="left")
        self.report_week_summary.pack(fill="x", padx=10, pady=(0, 6))
        week_cols = ("average", "tasks", "work", "date")
        self.report_week_tree = ttk.Treeview(weekly_tab, columns=week_cols, show="headings", height=10)
        for col, label, width in (("average", "Average Task Time", 150), ("tasks", "Tasks", 100),
                                  ("work", "Work Time", 150), ("date", "Workday", 150)):
            self.report_week_tree.heading(col, text=label, anchor="center")
            self.report_week_tree.column(col, width=width, minwidth=80, anchor="center", stretch=True)
        week_scroll = ttk.Scrollbar(weekly_tab, orient="vertical", command=self.report_week_tree.yview)
        self.report_week_tree.configure(yscrollcommand=week_scroll.set)
        self.report_week_tree.pack(side="left", fill="both", expand=True, padx=(10, 0), pady=(0, 8))
        week_scroll.pack(side="right", fill="y", padx=(0, 10), pady=(0, 8))

        style = ttk.Style(win)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=24, background="#FFFFFF", fieldbackground="#FFFFFF")
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#E5EDF6", foreground="#172033")
        style.map("Treeview", background=[("selected", "#D8EAFE")], foreground=[("selected", "#172033")])

        # Show report columns in reverse order so Date is at the right edge.
        daily_cols = ("average", "target_tasks", "actual_tasks", "actual_time", "date")
        self.report_daily_tree = ttk.Treeview(daily_tab, columns=daily_cols, show="headings", height=11)
        headings = {
            "date": "Date", "actual_time": "Actual Work Time",
            "actual_tasks": "Actual Tasks", "target_tasks": "Target Tasks",
            "average": "Average Task Time",
        }
        widths = {"date": 115, "actual_time": 150, "actual_tasks": 110, "target_tasks": 125, "average": 130}
        for col in daily_cols:
            self.report_daily_tree.heading(col, text=headings[col], anchor="center")
            self.report_daily_tree.column(col, width=widths[col], minwidth=80, anchor="center", stretch=True)
        daily_scroll = ttk.Scrollbar(daily_tab, orient="vertical", command=self.report_daily_tree.yview)
        self.report_daily_tree.configure(yscrollcommand=daily_scroll.set)
        self.report_daily_tree.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=8)
        daily_scroll.pack(side="right", fill="y", pady=8)
        self.report_daily_tree.tag_configure("total", background="#E1ECF8", font=("Segoe UI", 9, "bold"))
        self.report_daily_tree.tag_configure("negative", foreground="#B91C1C")
        self.report_daily_tree.tag_configure("positive", foreground="#15803D")

        # Reverse task-detail columns so Task # is at the right edge.
        task_cols = ("source", "status", "duration", "end", "start", "id")
        self.report_tasks_tree = ttk.Treeview(tasks_tab, columns=task_cols, show="headings", height=11)
        task_heads = {"id": "Task #", "start": "Start Time", "end": "End Time", "duration": "Duration", "status": "Status", "source": "Completed By"}
        task_widths = {"id": 65, "start": 145, "end": 145, "duration": 95, "status": 90, "source": 115}
        for col in task_cols:
            self.report_tasks_tree.heading(col, text=task_heads[col], anchor="center")
            self.report_tasks_tree.column(col, width=task_widths[col], minwidth=75, anchor="center", stretch=True)
        task_scroll = ttk.Scrollbar(tasks_tab, orient="vertical", command=self.report_tasks_tree.yview)
        self.report_tasks_tree.configure(yscrollcommand=task_scroll.set)
        self.report_tasks_tree.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=8)
        task_scroll.pack(side="right", fill="y", pady=8)

        manual_cols = ("created", "total_time", "tasks", "start_time", "end_time", "date")
        self.report_manual_tree = ttk.Treeview(manual_tab, columns=manual_cols, show="headings", height=11)
        manual_heads = {"created": "Added At", "total_time": "Total Time", "tasks": "Tasks Added",
                        "start_time": "From", "end_time": "To", "date": "Date"}
        manual_widths = {"created": 155, "total_time": 125, "tasks": 100, "end_time": 85, "start_time": 85, "date": 115}
        for col in manual_cols:
            self.report_manual_tree.heading(col, text=manual_heads[col], anchor="center")
            self.report_manual_tree.column(col, width=manual_widths[col], minwidth=85, anchor="center", stretch=True)
        manual_scroll = ttk.Scrollbar(manual_tab, orient="vertical", command=self.report_manual_tree.yview)
        self.report_manual_tree.configure(yscrollcommand=manual_scroll.set)
        self.report_manual_tree.pack(side="left", fill="both", expand=True, padx=(0, 8), pady=8)
        manual_scroll.pack(side="right", fill="y", pady=8)

        actions = tk.Frame(win, bg="#F4F7FB", padx=10, pady=7)
        actions.pack(fill="x")
        tk.Button(actions, text="Export to Excel", command=self.export_excel, bg=BLUE_BUTTON, fg=WHITE,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=9, pady=5, cursor="hand2").pack(side="right", padx=(0, 5))
        tk.Button(actions, text="Refresh", command=self.refresh_report, bg="#E2E8F0", fg="#172033",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=9, pady=5, cursor="hand2").pack(side="right", padx=5)
        tk.Button(actions, text="Workday Settings", command=self.open_workday_settings, bg="#E8EDF3", fg="#263445",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=9, pady=5, cursor="hand2").pack(side="right", padx=(0, 6))
        tk.Button(actions, text="Add Manual", command=self.open_add_manual, bg="#DCEFE2", fg="#166534",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=9, pady=5, cursor="hand2").pack(side="left", padx=(0, 6))
        tk.Button(actions, text="Edit Record", command=self.edit_selected_record, bg="#DBEAFE", fg="#1D4ED8",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=9, pady=5, cursor="hand2").pack(side="left", padx=(0, 6))
        tk.Button(actions, text="Delete Record", command=self.delete_selected_record, bg="#FEE2E2", fg="#991B1B",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=9, pady=5, cursor="hand2").pack(side="left", padx=(0, 6))
        tk.Button(actions, text="Clear History", command=self.clear_history, bg="#FEE2E2", fg="#991B1B",
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=9, pady=5, cursor="hand2").pack(side="left")
        self.refresh_report()

    def _selected_log_record(self):
        """Return (kind, record_id) for the selected row in an editable log tab."""
        if self.report_notebook is None:
            return None, None
        selected_tab = self.report_notebook.select()
        if selected_tab == str(self.report_tasks_tree.master):
            tree, kind = self.report_tasks_tree, "task"
        elif selected_tab == str(self.report_manual_tree.master):
            tree, kind = self.report_manual_tree, "manual"
        else:
            # Notebook tab IDs are frame paths; compare against the actual tab widgets.
            tree = None
            kind = None
            for tab_widget, candidate_tree, candidate_kind in (
                (self.report_tasks_tree.master, self.report_tasks_tree, "task"),
                (self.report_manual_tree.master, self.report_manual_tree, "manual"),
            ):
                if selected_tab == str(tab_widget):
                    tree, kind = candidate_tree, candidate_kind
                    break
        if tree is None:
            messagebox.showinfo("Select a record", "Open Task Details or Manual Adjustments, then select a record.", parent=self.report_window or self.root)
            return None, None
        selection = tree.selection()
        if not selection:
            messagebox.showinfo("Select a record", "Select one record first.", parent=self.report_window or self.root)
            return None, None
        iid = selection[0]
        try:
            record_id = int(iid.split("_", 1)[1])
        except (IndexError, ValueError):
            messagebox.showerror("Record unavailable", "This row cannot be edited.", parent=self.report_window or self.root)
            return None, None
        return kind, record_id

    def delete_selected_record(self):
        kind, record_id = self._selected_log_record()
        if not kind:
            return
        label = "task record" if kind == "task" else "manual adjustment"
        if not messagebox.askyesno("Delete Record", f"Delete this {label}? This action cannot be undone.", icon="warning", parent=self.report_window or self.root):
            return
        if kind == "task":
            existing = self.conn.execute("SELECT status FROM tasks WHERE id=?", (record_id,)).fetchone()
            if existing is None:
                messagebox.showerror("Record not found", "This task record no longer exists.", parent=self.report_window or self.root)
                return
            if existing["status"] == "active":
                messagebox.showinfo("Task is running", "A running task cannot be deleted from the log. Skip or stop it first.", parent=self.report_window or self.root)
                return
        with self.conn:
            if kind == "task":
                self.conn.execute("DELETE FROM tasks WHERE id=?", (record_id,))
            else:
                self.conn.execute("DELETE FROM manual_adjustments WHERE id=?", (record_id,))
        self.last_undo = None
        self._update_main_stats()
        self.refresh_report()

    def edit_selected_record(self):
        kind, record_id = self._selected_log_record()
        if not kind:
            return
        if kind == "task":
            self._edit_task_record(record_id)
        else:
            self._edit_manual_record(record_id)

    def _edit_task_record(self, record_id):
        row = self.conn.execute("SELECT * FROM tasks WHERE id=?", (record_id,)).fetchone()
        if row is None:
            messagebox.showerror("Record not found", "This task record no longer exists.", parent=self.report_window or self.root)
            return
        if row["status"] == "active":
            messagebox.showinfo("Task is running", "A running task cannot be edited. Stop, complete, or skip it first.", parent=self.report_window or self.root)
            return
        win = tk.Toplevel(self.report_window or self.root)
        win.title(f"Edit Task Record #{record_id}")
        win.geometry("440x270")
        win.resizable(False, False)
        win.transient(self.report_window or self.root)
        win.grab_set()
        frame = tk.Frame(win, bg="#F6F8FB", padx=14, pady=12)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text=f"Edit Task Record #{record_id}", font=("Segoe UI", 11, "bold"), bg="#F6F8FB", fg="#152238").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 8))
        record_offset = int(row["utc_offset_minutes"] if row["utc_offset_minutes"] is not None else self._current_offset_minutes())
        stored_start = parse_dt(row["started_at"])
        stored_end = parse_dt(row["ended_at"])
        shown_start = stored_start + timedelta(minutes=record_offset) if stored_start else None
        shown_end = stored_end + timedelta(minutes=record_offset) if stored_end else None
        start_var = tk.StringVar(value=(shown_start.strftime("%d/%m/%Y %H:%M:%S") if shown_start else ""))
        end_var = tk.StringVar(value=(shown_end.strftime("%d/%m/%Y %H:%M:%S") if shown_end else ""))
        duration_var = tk.StringVar(value=f"{int(row['duration_ms'] or int(row['duration_seconds'] or 0)*1000)/1000:.2f}")
        source_var = tk.StringVar(value=row["completion_source"] or "stop")
        for r, label, var, width in (
            (1, "Start (DD/MM/YYYY HH:MM:SS)", start_var, 25),
            (2, "End (DD/MM/YYYY HH:MM:SS)", end_var, 25),
            (3, "Duration (seconds)", duration_var, 12),
        ):
            tk.Label(frame, text=label, bg="#F6F8FB", fg="#263445", font=("Segoe UI", 9)).grid(row=r, column=0, sticky="w", pady=4)
            ttk.Entry(frame, textvariable=var, width=width).grid(row=r, column=1, sticky="w", pady=4)
        tk.Label(frame, text="Completed By", bg="#F6F8FB", fg="#263445", font=("Segoe UI", 9)).grid(row=4, column=0, sticky="w", pady=4)
        ttk.Combobox(frame, textvariable=source_var, values=("dodo", "stop", "manual"), state="readonly", width=12).grid(row=4, column=1, sticky="w", pady=4)

        def save():
            try:
                start_dt = datetime.strptime(start_var.get().strip(), "%d/%m/%Y %H:%M:%S")
                end_dt = datetime.strptime(end_var.get().strip(), "%d/%m/%Y %H:%M:%S")
                duration = float(duration_var.get().strip())
                if duration < 0 or duration > 7 * 24 * 3600:
                    raise ValueError("Duration must be between 0 and 604800 seconds.")
                if end_dt < start_dt:
                    raise ValueError("End time cannot be before start time.")
            except ValueError as exc:
                messagebox.showerror("Invalid values", str(exc) or "Check the date, time, and duration format.", parent=win)
                return
            duration_ms = int(round(duration * 1000))
            with self.conn:
                self.conn.execute(
                    "UPDATE tasks SET started_at=?, ended_at=?, duration_seconds=?, duration_ms=?, status='completed', completion_source=? WHERE id=?",
                    ((start_dt - timedelta(minutes=record_offset)).isoformat(timespec="seconds"), (end_dt - timedelta(minutes=record_offset)).isoformat(timespec="seconds"), int(round(duration)), duration_ms, source_var.get(), record_id),
                )
            self.last_undo = None
            win.destroy()
            self._update_main_stats()
            self.refresh_report()
        buttons = tk.Frame(frame, bg="#F6F8FB")
        buttons.grid(row=5, column=0, columnspan=2, sticky="w", pady=(12, 0))
        tk.Button(buttons, text="Save Changes", command=save, bg=BLUE_BUTTON, fg=WHITE, relief="flat", padx=10, pady=4).pack(side="left", padx=(0, 6))
        tk.Button(buttons, text="Cancel", command=win.destroy, bg="#E2E8F0", fg="#172033", relief="flat", padx=10, pady=4).pack(side="left")

    def _edit_manual_record(self, record_id):
        row = self.conn.execute("SELECT * FROM manual_adjustments WHERE id=?", (record_id,)).fetchone()
        if row is None:
            messagebox.showerror("Record not found", "This manual adjustment no longer exists.", parent=self.report_window or self.root)
            return
        win = tk.Toplevel(self.report_window or self.root)
        win.title(f"Edit Manual Adjustment #{record_id}")
        win.geometry("450x300")
        win.resizable(False, False)
        win.transient(self.report_window or self.root)
        win.grab_set()
        frame = tk.Frame(win, bg="#F6F8FB", padx=14, pady=12)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text=f"Edit Manual Adjustment #{record_id}", font=("Segoe UI", 11, "bold"), bg="#F6F8FB", fg="#152238").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        date_var = tk.StringVar(value=format_date(date.fromisoformat(row["work_date"])))
        tasks_var = tk.StringVar(value=str(row["task_count"]))
        start_var = tk.StringVar(value=format_ampm_clock(row["start_time"]).replace("—", ""))
        end_var = tk.StringVar(value=format_ampm_clock(row["end_time"]).replace("—", ""))
        for r, label, var, width in ((1, "Date (DD/MM/YYYY)", date_var, 15), (2, "Tasks completed", tasks_var, 10), (3, "From (e.g. 9:30 AM)", start_var, 16), (4, "To (e.g. 1:15 PM)", end_var, 16)):
            tk.Label(frame, text=label, bg="#F6F8FB", fg="#263445", font=("Segoe UI", 9)).grid(row=r, column=0, sticky="w", pady=4)
            ttk.Entry(frame, textvariable=var, width=width).grid(row=r, column=1, columnspan=2, sticky="w", pady=4)

        def save():
            try:
                date_text = date_var.get().strip()
                try:
                    work_day = datetime.strptime(date_text, "%d/%m/%Y").date()
                except ValueError:
                    work_day = datetime.strptime(date_text, "%Y-%m-%d").date()
                task_count = int(tasks_var.get().strip())
                if task_count < 1:
                    raise ValueError("Tasks completed must be at least 1.")
                start_clock = datetime.strptime(start_var.get().strip().upper(), "%I:%M %p").time()
                end_clock = datetime.strptime(end_var.get().strip().upper(), "%I:%M %p").time()
                start_dt = datetime.combine(work_day, start_clock)
                end_dt = datetime.combine(work_day, end_clock)
                if end_dt == start_dt:
                    raise ValueError("From and To cannot be the same time.")
                if end_dt < start_dt:
                    end_dt += timedelta(days=1)
                total_ms = int(round((end_dt - start_dt).total_seconds() * 1000))
                if total_ms <= 0 or total_ms > 24 * 3600 * 1000:
                    raise ValueError("The selected time range must be less than 24 hours.")
            except (TypeError, ValueError) as exc:
                messagebox.showerror("Invalid values", str(exc) or "Check the date, task count, and AM/PM times.", parent=win)
                return
            with self.conn:
                self.conn.execute(
                    "UPDATE manual_adjustments SET work_date=?, task_count=?, total_ms=?, start_time=?, end_time=? WHERE id=?",
                    (work_day.isoformat(), task_count, total_ms, start_clock.strftime("%H:%M"), end_clock.strftime("%H:%M"), record_id),
                )
            self.last_undo = None
            win.destroy()
            self._update_main_stats()
            self.refresh_report()
        buttons = tk.Frame(frame, bg="#F6F8FB")
        buttons.grid(row=5, column=0, columnspan=3, sticky="w", pady=(12, 0))
        tk.Button(buttons, text="Save Changes", command=save, bg=BLUE_BUTTON, fg=WHITE, relief="flat", padx=10, pady=4).pack(side="left", padx=(0, 6))
        tk.Button(buttons, text="Cancel", command=win.destroy, bg="#E2E8F0", fg="#172033", relief="flat", padx=10, pady=4).pack(side="left")

    def open_add_manual(self):
        """Add tasks for a chosen date and time interval without altering task-level records."""
        parent = self.report_window if self.report_window is not None else self.root
        win = tk.Toplevel(parent)
        win.title("Add Manual")
        win.geometry("450x330")
        win.resizable(False, False)
        win.transient(parent)
        win.grab_set()
        frame = tk.Frame(win, padx=14, pady=12, bg="#F6F8FB")
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text="Add Manual Adjustment", font=("Segoe UI", 11, "bold"),
                 bg="#F6F8FB", fg="#152238").pack(anchor="w", pady=(0, 4))
        tk.Label(frame, text="Choose the work date and time range.",
                 font=("Segoe UI", 8), bg="#F6F8FB", fg="#536273").pack(anchor="w", pady=(0, 8))

        fields = tk.Frame(frame, bg="#F6F8FB")
        fields.pack(fill="x")
        date_var = tk.StringVar(value=format_date(self._workday_today()))
        tasks_var = tk.StringVar()
        from_var = tk.StringVar()
        to_var = tk.StringVar()
        from_period = tk.StringVar(value="AM")
        to_period = tk.StringVar(value="PM")
        field_specs = [
            ("Date (DD/MM/YYYY)", date_var),
            ("Tasks completed", tasks_var),
        ]
        entry_widgets = {}
        for row_index, (label_text, variable) in enumerate(field_specs):
            tk.Label(fields, text=label_text, bg="#F6F8FB", fg="#263445",
                     font=("Segoe UI", 9), width=19, anchor="w").grid(row=row_index, column=0, sticky="w", pady=4)
            entry = ttk.Entry(fields, textvariable=variable, width=17, justify="center")
            entry.grid(row=row_index, column=1, sticky="w", pady=4)
            entry_widgets[label_text] = entry

        for row_index, (label_text, time_var, period_var) in enumerate([
            ("From time", from_var, from_period),
            ("To time", to_var, to_period),
        ], start=2):
            tk.Label(fields, text=label_text, bg="#F6F8FB", fg="#263445",
                     font=("Segoe UI", 9), width=19, anchor="w").grid(row=row_index, column=0, sticky="w", pady=4)
            time_entry = ttk.Entry(fields, textvariable=time_var, width=12, justify="center")
            time_entry.grid(row=row_index, column=1, sticky="w", pady=4)
            period_box = ttk.Combobox(fields, textvariable=period_var, values=("AM", "PM"),
                                      state="readonly", width=5, justify="center")
            period_box.grid(row=row_index, column=2, sticky="w", padx=(7, 0), pady=4)
            entry_widgets[label_text] = time_entry

        duration_var = tk.StringVar(value="Calculated duration: —")
        duration_label = tk.Label(frame, textvariable=duration_var, bg="#F6F8FB", fg="#075EAA",
                                  font=("Segoe UI", 9, "bold"), anchor="w")
        duration_label.pack(anchor="w", pady=(6, 0))

        def parse_form_values(show_errors=False):
            try:
                date_text = date_var.get().strip()
                try:
                    work_day = datetime.strptime(date_text, "%d/%m/%Y").date()
                except ValueError:
                    work_day = datetime.strptime(date_text, "%Y-%m-%d").date()
                task_count = int(tasks_var.get().strip())
                if task_count < 1:
                    raise ValueError("Tasks completed must be at least 1.")
                from_text = from_var.get().strip()
                to_text = to_var.get().strip()
                start_clock = datetime.strptime(f"{from_text} {from_period.get()}", "%I:%M %p").time()
                end_clock = datetime.strptime(f"{to_text} {to_period.get()}", "%I:%M %p").time()
                start_dt = datetime.combine(work_day, start_clock)
                end_dt = datetime.combine(work_day, end_clock)
                if end_dt == start_dt:
                    raise ValueError("From and To cannot be the same time.")
                if end_dt < start_dt:
                    end_dt += timedelta(days=1)  # supports a range that crosses midnight
                total_seconds = (end_dt - start_dt).total_seconds()
                if total_seconds <= 0 or total_seconds > 24 * 3600:
                    raise ValueError("The selected time range must be less than 24 hours.")
                return work_day, task_count, start_clock.strftime("%H:%M"), end_clock.strftime("%H:%M"), int(round(total_seconds * 1000))
            except (TypeError, ValueError) as exc:
                if show_errors:
                    messagebox.showerror(
                        "Invalid values",
                        str(exc) if str(exc) and not str(exc).startswith("time data") else
                        "Enter a valid date (DD/MM/YYYY), whole task count, and times in 12-hour HH:MM format with AM or PM selected.",
                        parent=win,
                    )
                return None

        def update_duration(*_args):
            parsed = parse_form_values(False)
            if parsed is None:
                duration_var.set("Calculated duration: —")
            else:
                duration_var.set(f"Calculated duration: {format_duration(parsed[4] / 1000.0, True)}")

        for variable in (date_var, tasks_var, from_var, to_var, from_period, to_period):
            variable.trace_add("write", update_duration)

        def save_adjustment():
            parsed = parse_form_values(True)
            if parsed is None:
                return
            work_day, task_count, start_text, end_text, total_ms = parsed
            with self.conn:
                cursor = self.conn.execute(
                    "INSERT INTO manual_adjustments(work_date, task_count, total_ms, created_at, start_time, end_time, utc_offset_minutes, workday_start_minutes) "
                    "VALUES(?, ?, ?, ?, ?, ?, ?, ?)",
                    (work_day.isoformat(), task_count, total_ms, iso_now(), start_text, end_text, self._current_offset_minutes(), self._current_workday_start_minutes()),
                )
            self.last_undo = {"kind": "manual_adjustment", "adjustment_id": int(cursor.lastrowid)}
            self._update_main_stats()
            self.refresh_report()
            win.destroy()

        buttons = tk.Frame(frame, bg="#F6F8FB")
        buttons.pack(fill="x", pady=(10, 0))
        tk.Button(buttons, text="Add", command=save_adjustment, bg=BLUE_BUTTON, fg=WHITE,
                  relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=4, cursor="hand2").pack(side="left", padx=(0, 6))
        tk.Button(buttons, text="Cancel", command=win.destroy, bg="#E2E8F0", fg="#172033",
                  relief="flat", font=("Segoe UI", 9), padx=12, pady=4, cursor="hand2").pack(side="left")
        entry_widgets["Tasks completed"].focus_set()

    def _close_report(self):
        if self.report_window is not None:
            try:
                self.report_window.destroy()
            except tk.TclError:
                pass
        self.report_window = None
        self.report_daily_tree = None
        self.report_tasks_tree = None
        self.report_manual_tree = None
        self.report_week_selector = None
        self.report_week_tree = None
        self.report_week_summary = None
        self.report_notebook = None

    def _refresh_week_archive(self):
        if self.report_week_selector is None or not self.report_week_selector.winfo_exists():
            return
        stats = self._statistics()
        rows = stats["daily_rows"]
        # Weeks are keyed by the workday date; Saturday 10 AM starts a new week.
        current_day = self._workday_today()
        current_week_start = current_day - timedelta(days=(current_day.weekday() + 2) % 7)
        starts = {current_week_start}
        for row in rows:
            day = row["date"]
            starts.add(day - timedelta(days=(day.weekday() + 2) % 7))
        starts = sorted(starts, reverse=True)
        labels = []
        self._week_archive_starts = {}
        for start in starts:
            end = start + timedelta(days=7)
            visible_end = end - timedelta(days=1)
            label = f"{start.strftime('%d %b %Y')} – {visible_end.strftime('%d %b %Y')}"
            labels.append(label)
            self._week_archive_starts[label] = start
        self.report_week_selector["values"] = labels
        selected = self.report_week_selector.get()
        if selected not in self._week_archive_starts:
            selected = labels[0] if labels else ""
            self.report_week_selector.set(selected)
        self._show_selected_week(rows, selected)

    def _show_selected_week(self, rows, label):
        if self.report_week_tree is None or not self.report_week_tree.winfo_exists():
            return
        for item in self.report_week_tree.get_children(""):
            self.report_week_tree.delete(item)
        start = getattr(self, "_week_archive_starts", {}).get(label)
        if start is None:
            self.report_week_summary.config(text="No saved weekly data yet.")
            return
        end = start + timedelta(days=7)
        selected = [row for row in rows if start <= row["date"] < end]
        total_work = sum(row["work_seconds"] for row in selected)
        total_tasks = sum(row["actual_tasks"] for row in selected)
        avg = sum(row["completed_duration"] for row in selected) / total_tasks if total_tasks else None
        visible_end = end - timedelta(days=1)
        boundary = self._current_workday_start_minutes()
        boundary_clock = (datetime.combine(date.today(), dt_time.min) + timedelta(minutes=boundary)).strftime("%I:%M %p").lstrip("0")
        self.report_week_summary.config(text=(
            f"Week: {start.strftime('%A, %d %b %Y')} – {visible_end.strftime('%A, %d %b %Y')}  "
            f"(daily boundary {boundary_clock} UTC{self._current_offset_minutes()/60:+g})   |   "
            f"Work: {format_duration(total_work, False)}   |   Tasks: {total_tasks}   |   Average: {format_average(avg)}"
        ))
        for row in sorted(selected, key=lambda item: item["date"], reverse=True):
            self.report_week_tree.insert("", "end", values=(format_average(row["avg_seconds"]), row["actual_tasks"],
                format_duration(row["work_seconds"], False), format_date(row["date"])))

    def refresh_report(self):
        if self.report_window is None or not self.report_window.winfo_exists():
            return
        stats = self._statistics()
        try:
            # Build each summary field separately so Tk packs it explicitly left-to-right.
            for child in self.report_summary_frame.winfo_children():
                child.destroy()
            summary_items = [
                ("Average Task", format_average(stats["total_avg_seconds"])),
                ("Target Tasks", str(stats["total_target_tasks"])),
                ("Actual Tasks", str(stats["total_tasks"])),
                ("Total Work Time", format_duration(stats["total_work_seconds"], False)),
                ("Days Worked", str(stats["days_worked"])),
            ]
            for index, (label_text, value_text) in enumerate(summary_items):
                if index:
                    tk.Label(self.report_summary_frame, text="|", font=("Segoe UI", 8),
                             bg="#F4F7FB", fg="#8290A2").pack(side="left", padx=6)
                tk.Label(self.report_summary_frame, text=f"{label_text}: {value_text}",
                         font=("Segoe UI", 8), bg="#F4F7FB", fg="#465568", anchor="w").pack(side="left")
            for tree in (self.report_daily_tree, self.report_tasks_tree, self.report_manual_tree):
                for item in tree.get_children(""):
                    tree.delete(item)

            for row in stats["daily_rows"]:
                self.report_daily_tree.insert("", "end", values=(
                    format_average(row["avg_seconds"]), row["target_tasks"],
                    row["actual_tasks"], format_duration(row["work_seconds"], False),
                    format_date(row["date"]),
                ))
            if stats["daily_rows"]:
                self.report_daily_tree.insert("", "end", values=(
                    format_average(stats["total_avg_seconds"]), stats["total_target_tasks"],
                    stats["total_tasks"], format_duration(stats["total_work_seconds"], False),
                    f"Total ({stats['days_worked']} days)",
                ), tags=("total",))
            records = stats["records"]
            for row in records:
                status_map = {"completed": "Completed", "incomplete": "Incomplete", "active": "Active"}
                source_map = {"dodo": "Start +1", "stop": "Stop", "close": "App Closed", "app_recovery": "Recovered After Crash", None: "—", "": "—"}
                duration = row.get("duration_seconds") or 0
                row_offset = int(row.get("utc_offset_minutes") if row.get("utc_offset_minutes") is not None else self._current_offset_minutes())
                start_text = format_datetime(row.get("started_at"), row_offset)
                end_text = format_datetime(row.get("ended_at"), row_offset)
                if row.get("_live"):
                    end_text = "Running"
                self.report_tasks_tree.insert("", "end", iid=f"task_{row['id']}", values=(
                    source_map.get(row.get("completion_source"), "—"),
                    status_map.get(row.get("status"), row.get("status", "—")),
                    format_stopwatch(float(row.get("duration_ms") or int(duration) * 1000) / 1000.0),
                    end_text, start_text, row["id"],
                ))
            manual_rows = self.conn.execute(
                "SELECT * FROM manual_adjustments ORDER BY id DESC"
            ).fetchall()
            for row in manual_rows:
                try:
                    adjusted_date = date.fromisoformat(row["work_date"])
                except (TypeError, ValueError):
                    continue
                total_seconds = max(0, int(row["total_ms"] or 0)) / 1000.0
                self.report_manual_tree.insert("", "end", iid=f"manual_{row['id']}", values=(
                    format_datetime(row["created_at"], int(row["utc_offset_minutes"] if row["utc_offset_minutes"] is not None else self._current_offset_minutes())), format_duration(total_seconds, True),
                    int(row["task_count"]), format_ampm_clock(row["start_time"]), format_ampm_clock(row["end_time"]),
                    format_date(adjusted_date),
                ))
            self._refresh_week_archive()
        except tk.TclError:
            pass

    def export_excel(self):
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Missing Library", "Excel export requires openpyxl. Run the included setup script, then try again.", parent=self.report_window or self.root)
            return
        path = filedialog.asksaveasfilename(
            parent=self.report_window or self.root,
            title="Save Work Log",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile=f"Task_Timer_Work_Log_{date.today().strftime('%Y-%m-%d')}.xlsx",
        )
        if not path:
            return
        stats = self._statistics()
        wb = Workbook()
        ws = wb.active
        ws.title = "Daily Summary"
        ws.sheet_view.rightToLeft = True
        headers = ["Date", "Actual Work Time", "Actual Tasks", "Target Tasks", "Average Task Time"]
        ws.append(headers)
        for row in stats["daily_rows"]:
            ws.append([
                format_date(row["date"]), format_duration(row["work_seconds"], False),
                row["actual_tasks"], row["target_tasks"], format_average(row["avg_seconds"]),
            ])
        if stats["daily_rows"]:
            ws.append([
                f"Total ({stats['days_worked']} days)",
                format_duration(stats["total_work_seconds"], False),
                stats["total_tasks"], stats["total_target_tasks"], format_average(stats["total_avg_seconds"]),
            ])
        tasks_ws = wb.create_sheet("Task Details")
        tasks_ws.sheet_view.rightToLeft = True
        tasks_ws.append(["Task #", "Start Time", "End Time", "Duration (ms)", "Duration (MM:SS.hh)", "Status", "Completed By"])
        status_map = {"completed": "Completed", "incomplete": "Incomplete", "active": "Active"}
        source_map = {"dodo": "Start +1", "stop": "Stop", "close": "App Closed", "app_recovery": "Recovered After Crash", None: "—", "": "—"}
        for row in stats["records"]:
            tasks_ws.append([
                row["id"], format_datetime(row.get("started_at"), int(row.get("utc_offset_minutes") if row.get("utc_offset_minutes") is not None else self._current_offset_minutes())),
                "Running" if row.get("_live") else format_datetime(row.get("ended_at"), int(row.get("utc_offset_minutes") if row.get("utc_offset_minutes") is not None else self._current_offset_minutes())),
                int(row.get("duration_ms") or int(row.get("duration_seconds") or 0) * 1000),
                format_stopwatch(float(row.get("duration_ms") or int(row.get("duration_seconds") or 0) * 1000) / 1000.0),
                status_map.get(row.get("status"), row.get("status", "—")),
                source_map.get(row.get("completion_source"), "—"),
            ])

        manual_ws = wb.create_sheet("Manual Adjustments")
        manual_ws.sheet_view.rightToLeft = True
        manual_ws.append(["Date", "From", "To", "Tasks Added", "Total Time", "Added At"])
        for row in self.conn.execute("SELECT * FROM manual_adjustments ORDER BY id DESC").fetchall():
            try:
                adjusted_date = date.fromisoformat(row["work_date"])
            except (TypeError, ValueError):
                continue
            manual_ws.append([
                format_date(adjusted_date), format_ampm_clock(row["start_time"]), format_ampm_clock(row["end_time"]),
                int(row["task_count"]), format_duration(max(0, int(row["total_ms"] or 0)) / 1000.0, True),
                format_datetime(row["created_at"], int(row["utc_offset_minutes"] if row["utc_offset_minutes"] is not None else self._current_offset_minutes())),
            ])

        header_fill = PatternFill("solid", fgColor="075EAA")
        for sheet in wb.worksheets:
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions if sheet.max_row > 1 else "A1"
            for cell in sheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for row_cells in sheet.iter_rows(min_row=2):
                for cell in row_cells:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            for column_cells in sheet.columns:
                col_letter = column_cells[0].column_letter
                max_len = max((len(str(c.value or "")) for c in column_cells), default=10)
                sheet.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 34)
        try:
            wb.save(path)
            messagebox.showinfo("Export Complete", f"Excel file saved successfully:\n{path}", parent=self.report_window or self.root)
        except Exception as exc:
            messagebox.showerror("Export Failed", f"Could not save the file.\n{exc}", parent=self.report_window or self.root)

    def clear_history(self):
        active_desc = "\nThe current task, if any, will be stopped." if self.current_task_id is not None else ""
        confirm = messagebox.askyesno(
            "Confirm Clear History",
            "Delete all task records and historical statistics?\nThis cannot be undone. Your settings and shortcuts will be kept." + active_desc,
            icon="warning", parent=self.report_window or self.root,
        )
        if not confirm:
            return
        if self.current_task_id is not None:
            self.current_task_id = None
            self.current_base_elapsed = 0
            self.current_mono_start = 0.0
        with self.conn:
            self.conn.execute("DELETE FROM tasks")
            self.conn.execute("DELETE FROM manual_adjustments")
        self.last_undo = None
        self.last_display_elapsed = 0
        self._update_status("History cleared. Press Start +1 to begin.")
        self._update_main_stats()
        self.refresh_report()

    # ---------- Shutdown ----------
    def close_app(self):
        # An unfinished task is discarded on normal exit: no task count and no
        # work time from the current attempt is carried into future statistics.
        if self.current_task_id is not None:
            task_id = self.current_task_id
            with self.conn:
                self.conn.execute("DELETE FROM tasks WHERE id=?", (task_id,))
            self.current_task_id = None
            self.current_base_elapsed = 0.0
            self.current_mono_start = 0.0
            self.last_display_elapsed = 0.0
        self._stop_global_listeners()
        try:
            self.conn.commit()
            self.conn.close()
        except Exception:
            pass
        self.root.destroy()


def main():
    if sys.platform != "win32":
        # Native global hotkeys are used on Windows; pynput is the fallback elsewhere.
        pass
    root = tk.Tk()
    app = DodoTaskTimer(root)
    root.mainloop()


if __name__ == "__main__":
    main()
