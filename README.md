# ICS Viewer

A simple Windows viewer for iCalendar (`.ics`) files, with two views:

- **Calendar**: month grid; click an event for details or a day to see all its events.
- **List**: compact, searchable list that scrolls to the next upcoming event.

The interface is in **Italian** or **English**, picked automatically from the Windows display language
(set `ICSVIEWER_LANG=it` or `en` to force one).

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
