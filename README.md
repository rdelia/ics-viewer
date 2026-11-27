# ICS Viewer

A simple Windows viewer for iCalendar (`.ics`) files, with two views:

- **Calendar**: month grid; click an event for details or a day to see all its events.
- **List**: compact, searchable list that scrolls to the next upcoming event.

The interface is in **Italian** or **English**. By default it follows the Windows display language;
pick a language from the **Language** menu in the toolbar to override it (the choice is remembered in
`%APPDATA%\ICSViewer\settings.json`).

**Export .xlsx** (Ctrl+E) saves the events to an Excel workbook, with real date/time cells, a frozen
header row and filters. In the list view it exports what the list shows (search and "hide past events"
filters applied); in the calendar view it exports every event.

Supports time zones (IANA and Windows/Outlook names), all-day and multi-day events,
recurring events (RRULE, RDATE, EXDATE) and moved occurrences (RECURRENCE-ID).

## Download

Run [`dist/ICSViewer.exe`](dist/ICSViewer.exe): a single file, no installation needed.
Open a file with **Open…** (Ctrl+O), drag an `.ics` onto the `.exe`, or set it as the "Open with" program for `.ics` files.

## Build from source

Requires Python 3 with `python-dateutil`, `tzdata` and `pyinstaller`.

```powershell
python ics_viewer.py sample.ics
python -m PyInstaller --noconfirm --onefile --windowed --name ICSViewer --collect-data tzdata --hidden-import tzdata --version-file version_info.txt ics_viewer.py
```

## Author

Raffaele D'Elia — <https://raffaeledelia.com>

Copyright © 2026 Raffaele D'Elia
