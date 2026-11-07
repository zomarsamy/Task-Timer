Task Timer — V23 Weekly Archive (UTC+3 / 3:00 AM)

This version corrects the timestamp model and weekly grouping.

- Stored task/event timestamps use UTC internally; task details and exports display the configured UTC offset.
- Default timezone is UTC+3 and the workday starts at 3:00 AM UTC+3 (00:00 UTC).
- Weekly intervals run from Saturday 3:00 AM UTC+3 to the following Saturday 3:00 AM UTC+3.
- Archive labels show the seven included workday dates, Saturday through Friday (for example, 10 Oct–16 Oct), while the internal next-week boundary is Saturday 3:00 AM.
- On first launch, the app migrates legacy timestamps that were previously stored as Egypt UTC+3 wall time to UTC. It preserves task durations and manual work-date/time entries, then recalculates old records under UTC+3 / 3:00 AM. This migration runs once.
- Workday Settings can apply changes to the selected week, all weeks, or new records only.

Run_DodoTaskTimer.bat starts the app. Build_EXE.bat builds a Windows executable on Windows with Python and required packages installed.
