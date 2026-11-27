"""ICS Viewer - a simple viewer for iCalendar (.ics) files.

Two views: a month calendar and a compact list.
UI language (Italian / English) is detected automatically from the system;
set the environment variable ICSVIEWER_LANG=it|en to force one.
"""
import sys
import os
import re
import locale
import zlib
import calendar
import json
import zipfile
from xml.sax.saxutils import escape as xml_escape
import webbrowser
from datetime import datetime, date, time, timedelta, timezone
from zoneinfo import ZoneInfo

import tkinter as tk
import tkinter.font as tkfont
from tkinter import ttk, filedialog, messagebox

from dateutil.rrule import rrulestr, rruleset

APP_NAME = "ICS Viewer"
APP_VERSION = "1.1.0"
AUTHOR = "Raffaele D'Elia"
AUTHOR_URL = "https://raffaeledelia.com"
COPYRIGHT = "Copyright © 2026 Raffaele D'Elia"

# --------------------------------------------------------------------------
# Localisation
# --------------------------------------------------------------------------

STRINGS = {
    "en": {
        "open": "Open…", "calendar": "Calendar", "list": "List", "today": "Today",
        "search": "Search:", "hide_past": "Hide past events", "all_day": "All day",
        "no_file": "Open an .ics file to get started  (Ctrl+O)",
        "events": "{n} events", "more": "+{n} more",
        "when": "When", "location": "Location", "description": "Description",
        "organizer": "Organizer", "attendees": "Attendees", "url": "Link",
        "status": "Status", "recurring": "Recurring event", "close": "Close", "about": "About",
        "language": "Language:", "lang_auto": "Automatic", "export": "Export .xlsx",
        "export_done": "Exported {n} events to:\n{path}", "nothing_to_export": "There are no events to export.",
        "error_write": "Could not save the file:\n{e}", "sheet": "Events", "yes": "Yes", "no": "No",
        "xlsx_cols": ["Start date", "Start time", "End date", "End time", "All day", "Title", "Location",
                      "Description", "Organizer", "Attendees", "Status", "Recurring", "Link"],
        "about_text": "A simple viewer for iCalendar (.ics) files.", "version": "Version", "author": "Author",
        "error": "Error", "error_read": "Could not read the file:\n{e}",
        "cal_files": "Calendar files", "all_files": "All files",
        "col_date": "Date", "col_time": "Time", "col_title": "Event", "col_loc": "Location",
        "no_title": "(no title)", "no_events_day": "No events",
        "months": ["January", "February", "March", "April", "May", "June", "July",
                   "August", "September", "October", "November", "December"],
        "months_short": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug",
                         "Sep", "Oct", "Nov", "Dec"],
        "days": ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"],
        "days_short": ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"],
        "long_date": "{wd}, {d} {month} {y}",
    },
    "it": {
        "open": "Apri…", "calendar": "Calendario", "list": "Elenco", "today": "Oggi",
        "search": "Cerca:", "hide_past": "Nascondi eventi passati", "all_day": "Tutto il giorno",
        "no_file": "Apri un file .ics per iniziare  (Ctrl+O)",
        "events": "{n} eventi", "more": "+{n} altri",
        "when": "Quando", "location": "Luogo", "description": "Descrizione",
        "organizer": "Organizzatore", "attendees": "Partecipanti", "url": "Link",
        "status": "Stato", "recurring": "Evento ricorrente", "close": "Chiudi", "about": "Informazioni",
        "language": "Lingua:", "lang_auto": "Automatica", "export": "Esporta .xlsx",
        "export_done": "Esportati {n} eventi in:\n{path}", "nothing_to_export": "Non ci sono eventi da esportare.",
        "error_write": "Impossibile salvare il file:\n{e}", "sheet": "Eventi", "yes": "Sì", "no": "No",
        "xlsx_cols": ["Data inizio", "Ora inizio", "Data fine", "Ora fine", "Tutto il giorno", "Titolo", "Luogo",
                      "Descrizione", "Organizzatore", "Partecipanti", "Stato", "Ricorrente", "Link"],
        "about_text": "Un semplice visualizzatore di file iCalendar (.ics).", "version": "Versione", "author": "Autore",
        "error": "Errore", "error_read": "Impossibile leggere il file:\n{e}",
        "cal_files": "File calendario", "all_files": "Tutti i file",
        "col_date": "Data", "col_time": "Ora", "col_title": "Evento", "col_loc": "Luogo",
        "no_title": "(senza titolo)", "no_events_day": "Nessun evento",
        "months": ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio",
                   "agosto", "settembre", "ottobre", "novembre", "dicembre"],
        "months_short": ["gen", "feb", "mar", "apr", "mag", "giu", "lug", "ago",
                         "set", "ott", "nov", "dic"],
        "days": ["lunedì", "martedì", "mercoledì", "giovedì", "venerdì", "sabato", "domenica"],
        "days_short": ["lun", "mar", "mer", "gio", "ven", "sab", "dom"],
        "long_date": "{wd} {d} {month} {y}",
    },
}


LANGUAGE_NAMES = {"en": "English", "it": "Italiano"}
SETTINGS_PATH = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"),
                             "ICSViewer", "settings.json")


def load_settings():
    try:
        with open(SETTINGS_PATH, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_settings(settings):
    try:
        os.makedirs(os.path.dirname(SETTINGS_PATH), exist_ok=True)
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except OSError:
        pass


def system_language():
    if sys.platform == "win32":
        try:
            import ctypes
            langid = ctypes.windll.kernel32.GetUserDefaultUILanguage()
            return "it" if (langid & 0x3FF) == 0x10 else "en"
        except Exception:
            pass
    for value in (os.environ.get("LANG"), os.environ.get("LC_ALL"), locale.getlocale()[0]):
        if value and (value.lower().startswith("it") or "italian" in value.lower()):
            return "it"
    return "en"


def detect_language():
    """ICSVIEWER_LANG env var > language chosen in the app > Windows display language."""
    forced = os.environ.get("ICSVIEWER_LANG", "").lower()[:2]
    if forced in STRINGS:
        return forced
    chosen = load_settings().get("language", "auto")
    return chosen if chosen in STRINGS else system_language()


LANG = detect_language()
T = dict(STRINGS[LANG])  # updated in place by set_language()


def set_language(lang):
    global LANG
    LANG = lang if lang in STRINGS else system_language()
    T.clear()
    T.update(STRINGS[LANG])


def fmt_date_short(d):
    return f"{T['days_short'][d.weekday()]} {d.day:02d} {T['months_short'][d.month - 1]} {d.year}"


def fmt_date_long(d):
    s = T["long_date"].format(wd=T["days"][d.weekday()], d=d.day,
                              month=T["months"][d.month - 1], y=d.year)
    return s[0].upper() + s[1:]


def fmt_time(dt):
    return dt.strftime("%H:%M")


# --------------------------------------------------------------------------
# iCalendar parsing
# --------------------------------------------------------------------------

WINDOWS_TZ = {
    "W. Europe Standard Time": "Europe/Berlin",
    "Romance Standard Time": "Europe/Paris",
    "Central Europe Standard Time": "Europe/Budapest",
    "Central European Standard Time": "Europe/Warsaw",
    "GMT Standard Time": "Europe/London",
    "Greenwich Standard Time": "Atlantic/Reykjavik",
    "E. Europe Standard Time": "Europe/Chisinau",
    "FLE Standard Time": "Europe/Kiev",
    "GTB Standard Time": "Europe/Bucharest",
    "Russian Standard Time": "Europe/Moscow",
    "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago",
    "Mountain Standard Time": "America/Denver",
    "Pacific Standard Time": "America/Los_Angeles",
    "US Mountain Standard Time": "America/Phoenix",
    "Alaskan Standard Time": "America/Anchorage",
    "Hawaiian Standard Time": "Pacific/Honolulu",
    "Atlantic Standard Time": "America/Halifax",
    "E. South America Standard Time": "America/Sao_Paulo",
    "India Standard Time": "Asia/Kolkata",
    "China Standard Time": "Asia/Shanghai",
    "Tokyo Standard Time": "Asia/Tokyo",
    "Singapore Standard Time": "Asia/Singapore",
    "AUS Eastern Standard Time": "Australia/Sydney",
    "Arabian Standard Time": "Asia/Dubai",
    "Israel Standard Time": "Asia/Jerusalem",
    "South Africa Standard Time": "Africa/Johannesburg",
    "UTC": "UTC", "Coordinated Universal Time": "UTC",
}

_tz_cache = {}


def get_tz(tzid):
    if not tzid:
        return None
    if tzid in _tz_cache:
        return _tz_cache[tzid]
    candidates = [tzid, WINDOWS_TZ.get(tzid)]
    m = re.search(r"([A-Za-z]+/[A-Za-z_\-]+(?:/[A-Za-z_\-]+)?)$", tzid)
    if m:
        candidates.append(m.group(1))
    tz = None
    for c in candidates:
        if not c:
            continue
        try:
            tz = ZoneInfo(c)
            break
        except Exception:
            pass
    if tz is None and tzid.upper() in ("UTC", "GMT", "Z"):
        tz = timezone.utc
    _tz_cache[tzid] = tz
    return tz


def unfold(text):
    lines = []
    for raw in text.splitlines():
        if raw[:1] in (" ", "\t") and lines:
            lines[-1] += raw[1:]
        else:
            lines.append(raw)
    return lines


def split_line(line):
    """Split 'NAME;P1=a;P2="b:c":value' into (NAME, {P1: a, P2: b:c}, value)."""
    in_quotes = False
    idx = -1
    for i, c in enumerate(line):
        if c == '"':
            in_quotes = not in_quotes
        elif c == ":" and not in_quotes:
            idx = i
            break
    if idx < 0:
        return None
    head, value = line[:idx], line[idx + 1:]
    parts, cur, in_quotes = [], "", False
    for c in head:
        if c == '"':
            in_quotes = not in_quotes
            cur += c
        elif c == ";" and not in_quotes:
            parts.append(cur)
            cur = ""
        else:
            cur += c
    parts.append(cur)
    name = parts[0].upper().split(".")[-1]  # drop optional group prefix
    params = {}
    for p in parts[1:]:
        if "=" in p:
            k, v = p.split("=", 1)
            params[k.upper()] = v.strip('"')
    return name, params, value


def unescape(s):
    return re.sub(r"\\([nN,;\\])",
                  lambda m: "\n" if m.group(1) in "nN" else m.group(1), s)


def parse_ics(text):
    """Return (calendar_name, [event_dict]) where event_dict maps NAME -> [(params, value)]."""
    events, stack, cur, calname = [], [], None, ""
    for line in unfold(text):
        if not line.strip():
            continue
        r = split_line(line)
        if not r:
            continue
        name, params, value = r
        if name == "BEGIN":
            stack.append(value.strip().upper())
            if stack[-1] == "VEVENT" and cur is None:
                cur = {}
            continue
        if name == "END":
            v = value.strip().upper()
            if stack:
                stack.pop()
            if v == "VEVENT" and cur is not None:
                events.append(cur)
                cur = None
            continue
        if cur is not None and stack and stack[-1] == "VEVENT":
            cur.setdefault(name, []).append((params, value))
        elif stack and stack[-1] == "VCALENDAR" and name == "X-WR-CALNAME":
            calname = unescape(value)
    return calname, events


def parse_dt(value, params):
    value = value.strip()
    if params.get("VALUE", "").upper() == "DATE" or re.fullmatch(r"\d{8}", value):
        return datetime.strptime(value[:8], "%Y%m%d").date()
    m = re.fullmatch(r"(\d{8})T(\d{4,6})(Z?)", value)
    if not m:
        raise ValueError(f"Invalid date: {value}")
    dt = datetime.strptime(m.group(1) + m.group(2).ljust(6, "0"), "%Y%m%d%H%M%S")
    if m.group(3):
        return dt.replace(tzinfo=timezone.utc)
    tz = get_tz(params.get("TZID"))
    return dt.replace(tzinfo=tz) if tz else dt


def is_date(v):
    return isinstance(v, date) and not isinstance(v, datetime)


def to_local(v):
    """Dates stay dates; aware datetimes become naive local time; floating stays as is."""
    if is_date(v):
        return v
    if v.tzinfo is not None:
        return v.astimezone().replace(tzinfo=None)
    return v


def parse_duration(s):
    m = re.fullmatch(r"([+-])?P(?:(\d+)W)?(?:(\d+)D)?(?:T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?)?",
                     s.strip())
    if not m:
        return None
    sign = -1 if m.group(1) == "-" else 1
    w, d, h, mi, se = (int(x or 0) for x in m.groups()[1:])
    return sign * timedelta(weeks=w, days=d, hours=h, minutes=mi, seconds=se)


def first(ev, name, default=""):
    vals = ev.get(name)
    return vals[0][1] if vals else default


PALETTE = ["#1a73e8", "#d93025", "#188038", "#e37400", "#9334e6",
           "#c5221f", "#007b83", "#b06000", "#3949ab", "#ad1457"]


class Occurrence:
    __slots__ = ("start", "end", "all_day", "info")

    def __init__(self, start, end, all_day, info):
        self.start, self.end, self.all_day, self.info = start, end, all_day, info

    def start_dt(self):
        return datetime.combine(self.start, time()) if self.all_day else self.start

    def end_dt(self):
        return datetime.combine(self.end, time()) if self.all_day else self.end

    def days(self):
        """Dates covered by this occurrence (end is exclusive)."""
        if self.all_day:
            first_d, last_d = self.start, self.end - timedelta(days=1)
        else:
            first_d = self.start.date()
            last_d = (self.end - timedelta(microseconds=1)).date() if self.end > self.start else first_d
        if last_d < first_d:
            last_d = first_d
        n = min((last_d - first_d).days, 366)
        return [first_d + timedelta(days=i) for i in range(n + 1)]


def _fix_until(rule, tz):
    """Make RRULE UNTIL comparable with naive wall-clock DTSTART."""
    def rep(m):
        v = m.group(1)
        if re.fullmatch(r"\d{8}", v):
            return "UNTIL=" + v + "T235959"
        mm = re.fullmatch(r"(\d{8}T\d{6})Z", v, re.I)
        if mm:
            dt = datetime.strptime(mm.group(1), "%Y%m%dT%H%M%S").replace(tzinfo=timezone.utc)
            dt = dt.astimezone(tz) if tz else dt.astimezone()
            return "UNTIL=" + dt.strftime("%Y%m%dT%H%M%S")
        return m.group(0)
    return re.sub(r"UNTIL=([0-9TZtz]+)", rep, rule, flags=re.I)


def _to_wall(v, tz, all_day, base_time):
    if all_day:
        return datetime.combine(v if is_date(v) else v.date(), time())
    if is_date(v):
        return datetime.combine(v, base_time)
    if v.tzinfo is not None:
        return (v.astimezone(tz) if tz else v.astimezone()).replace(tzinfo=None)
    return v


MAX_OCCURRENCES = 2000


def build_occurrences(events):
    def okey(v):
        return to_local(v)

    overrides = set()
    for e in events:
        if "RECURRENCE-ID" in e:
            p, v = e["RECURRENCE-ID"][0]
            try:
                overrides.add((first(e, "UID"), okey(parse_dt(v, p))))
            except ValueError:
                pass

    result = []
    now = datetime.now()
    for e in events:
        if "DTSTART" not in e:
            continue
        try:
            p, v = e["DTSTART"][0]
            start = parse_dt(v, p)
        except ValueError:
            continue
        all_day = is_date(start)
        end = None
        try:
            if "DTEND" in e:
                end = parse_dt(e["DTEND"][0][1], e["DTEND"][0][0])
            elif "DURATION" in e:
                dur = parse_duration(first(e, "DURATION"))
                if dur is not None:
                    end = start + (timedelta(days=dur.days) if all_day else dur)
        except ValueError:
            end = None
        if end is None:
            end = start + timedelta(days=1) if all_day else start
        if all_day and not is_date(end):
            end = end.date()
        elif not all_day and is_date(end):
            end = datetime.combine(end, time(), tzinfo=start.tzinfo)
        try:
            duration = end - start
        except TypeError:  # naive vs aware mix
            duration = end.replace(tzinfo=None) - start.replace(tzinfo=None)
        if duration < timedelta(0):
            duration = timedelta(0)
        if all_day and duration < timedelta(days=1):
            duration = timedelta(days=1)

        summary = unescape(first(e, "SUMMARY")).strip()
        organizer = first(e, "ORGANIZER")
        org_params = e["ORGANIZER"][0][0] if "ORGANIZER" in e else {}
        org_name = org_params.get("CN") or re.sub(r"^mailto:", "", organizer, flags=re.I)
        attendees = []
        for ap, av in e.get("ATTENDEE", []):
            attendees.append(ap.get("CN") or re.sub(r"^mailto:", "", av, flags=re.I))
        uid = first(e, "UID")
        is_rec = "RRULE" in e or "RDATE" in e
        info = {
            "summary": summary or T["no_title"],
            "location": unescape(first(e, "LOCATION")).strip(),
            "description": unescape(first(e, "DESCRIPTION")).strip(),
            "organizer": org_name,
            "attendees": attendees,
            "url": first(e, "URL").strip(),
            "status": first(e, "STATUS").strip(),
            "uid": uid,
            "recurring": is_rec or "RECURRENCE-ID" in e,
            "color": PALETTE[zlib.crc32((summary or uid).encode("utf-8")) % len(PALETTE)],
        }

        if not is_rec or "RECURRENCE-ID" in e:
            result.append(Occurrence(to_local(start), to_local(start + duration) if not all_day
                                     else start + duration, all_day, info))
            continue

        tz = None if all_day else start.tzinfo
        base = datetime.combine(start, time()) if all_day else start.replace(tzinfo=None)
        rs = rruleset()
        try:
            for _, rule in e.get("RRULE", []):
                rs.rrule(rrulestr(_fix_until(rule.strip(), tz), dtstart=base))
            for name, adder in (("RDATE", rs.rdate), ("EXDATE", rs.exdate)):
                for pp, vv in e.get(name, []):
                    if pp.get("VALUE", "").upper() == "PERIOD":
                        continue
                    for item in vv.split(","):
                        if item.strip():
                            adder(_to_wall(parse_dt(item, pp), tz, all_day, base.time()))
            rs.rdate(base)  # DTSTART is always the first instance (duplicates are merged)
        except Exception:
            result.append(Occurrence(to_local(start), to_local(start + duration) if not all_day
                                     else start + duration, all_day, info))
            continue

        horizon = max(now, base) + timedelta(days=3 * 366)
        count = 0
        for occ in rs:
            if occ > horizon or count >= MAX_OCCURRENCES:
                break
            count += 1
            if all_day:
                s = occ.date()
                o = Occurrence(s, s + duration, True, info)
            else:
                aware = occ.replace(tzinfo=tz) if tz else occ
                o = Occurrence(to_local(aware), to_local(aware + duration), False, info)
            if (uid, o.start) in overrides:
                continue
            result.append(o)

    result.sort(key=lambda o: (o.start_dt().date(), 0 if o.all_day else 1, o.start_dt()))
    return result


def load_ics_file(path):
    with open(path, "rb") as f:
        data = f.read()
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("cp1252", errors="replace")
    calname, events = parse_ics(text)
    return calname, build_occurrences(events)


# --------------------------------------------------------------------------
# Minimal .xlsx writer (no external dependencies)
# --------------------------------------------------------------------------

_XML_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")
_EXCEL_EPOCH = date(1899, 12, 30)
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def _col_letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _xlsx_cell(ref, v, header=False):
    """Styles: 1 = bold header, 2 = date, 3 = time."""
    if v is None or v == "":
        return ""
    if header or isinstance(v, str):
        text = xml_escape(_XML_BAD.sub("", str(v))[:32000])
        st = ' s="1"' if header else ""
        return f'<c r="{ref}" t="inlineStr"{st}><is><t xml:space="preserve">{text}</t></is></c>'
    if isinstance(v, datetime):
        v = v.date()
    if isinstance(v, date):
        return f'<c r="{ref}" s="2"><v>{(v - _EXCEL_EPOCH).days}</v></c>'
    if isinstance(v, time):
        return f'<c r="{ref}" s="3"><v>{(v.hour * 3600 + v.minute * 60 + v.second) / 86400!r}</v></c>'
    return f'<c r="{ref}"><v>{v}</v></c>'


def write_xlsx(path, sheet_name, headers, rows, widths):
    """Write a single-sheet workbook. Row values: str, int/float, date, time or None."""
    xml_rows = []
    for r, values in enumerate([headers] + rows, start=1):
        cells = "".join(_xlsx_cell(f"{_col_letter(c)}{r}", v, r == 1) for c, v in enumerate(values))
        xml_rows.append(f'<row r="{r}">{cells}</row>')
    last_col = _col_letter(len(headers) - 1)
    last_row = len(rows) + 1
    cols = "".join(f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>'
                   for i, w in enumerate(widths))
    sheet = (f'{_XML_HEAD}<worksheet xmlns="{_NS_MAIN}">'
             '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
             'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
             f'<cols>{cols}</cols><sheetData>{"".join(xml_rows)}</sheetData>'
             f'<autoFilter ref="A1:{last_col}{last_row}"/></worksheet>')
    styles = (f'{_XML_HEAD}<styleSheet xmlns="{_NS_MAIN}">'
              '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
              '<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
              '<fills count="2"><fill><patternFill patternType="none"/></fill>'
              '<fill><patternFill patternType="gray125"/></fill></fills>'
              '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
              '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
              '<xf numFmtId="20" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
              '</cellXfs></styleSheet>')
    name = re.sub(r"[\\/?*\[\]:']", "", sheet_name)[:31] or "Sheet1"
    workbook = (f'{_XML_HEAD}<workbook xmlns="{_NS_MAIN}" xmlns:r="{_NS_REL}">'
                f'<sheets><sheet name="{xml_escape(name)}" sheetId="1" r:id="rId1"/></sheets>'
                '<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'
                f"'{xml_escape(name)}'!$A$1:${last_col}${last_row}</definedName></definedNames>"
                '</workbook>')
    ct = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    files = {
        "[Content_Types].xml":
            f'{_XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{ct}.sheet.main+xml"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{ct}.worksheet+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{ct}.styles+xml"/>'
            '</Types>',
        "_rels/.rels":
            f'{_XML_HEAD}<Relationships xmlns="{_NS_PKG}">'
            f'<Relationship Id="rId1" Type="{_NS_REL}/officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>',
        "xl/_rels/workbook.xml.rels":
            f'{_XML_HEAD}<Relationships xmlns="{_NS_PKG}">'
            f'<Relationship Id="rId1" Type="{_NS_REL}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{_NS_REL}/styles" Target="styles.xml"/>'
            '</Relationships>',
        "xl/workbook.xml": workbook,
        "xl/styles.xml": styles,
        "xl/worksheets/sheet1.xml": sheet,
    }
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in files.items():
            z.writestr(n, data.encode("utf-8"))
    os.replace(tmp, path)


# --------------------------------------------------------------------------
# GUI
# --------------------------------------------------------------------------

BG = "#ffffff"
GRID = "#dadce0"
MUTED = "#80868b"
OTHER_MONTH_BG = "#f8f9fa"
TODAY_BG = "#e8f0fe"
ACCENT = "#1a73e8"


class App(tk.Tk):
    def __init__(self, path=None):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1150x760")
        self.minsize(720, 480)
        self.configure(bg=BG)

        self.occs = []
        self.by_day = {}
        self.calname = ""
        self.path = None
        today = date.today()
        self.year, self.month = today.year, today.month
        self.view = tk.StringVar(value="cal")
        self.search_var = tk.StringVar()
        self.hide_past = tk.BooleanVar(value=False)
        self._cap = 4
        self._resize_job = None

        self.font = tkfont.Font(family="Segoe UI", size=9)
        self.font_bold = tkfont.Font(family="Segoe UI", size=9, weight="bold")
        self.font_title = tkfont.Font(family="Segoe UI", size=14, weight="bold")
        self.font_head = tkfont.Font(family="Segoe UI", size=9, weight="bold")

        style = ttk.Style(self)
        try:
            style.theme_use("vista" if sys.platform == "win32" else "clam")
        except tk.TclError:
            pass
        style.configure("Treeview", font=("Segoe UI", 9), rowheight=int(self.font.metrics("linespace") * 1.45))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        self.search_var.trace_add("write", lambda *a: self.render_list())
        self._build_ui()
        self.bind_all("<MouseWheel>", self._on_wheel)
        self.bind("<Control-o>", lambda e: self.open_dialog())
        self.bind("<Control-e>", lambda e: self.export_xlsx())
        self.bind("<F1>", lambda e: self.show_about())
        self.bind("<Prior>", lambda e: self.shift_month(-1))
        self.bind("<Next>", lambda e: self.shift_month(1))
        self.bind("<Home>", lambda e: self.go_today())

        self.show_view()
        if path:
            self.after(50, lambda: self.load(path))

    def _build_ui(self):
        self._build_toolbar()
        self.body = tk.Frame(self, bg=BG)
        self.body.pack(fill="both", expand=True)
        self.cal_frame = tk.Frame(self.body, bg=BG)
        self.list_frame = ttk.Frame(self.body)
        self._build_list()
        self.status = ttk.Label(self, text=T["no_file"], anchor="w", padding=(8, 3))
        self.status.pack(fill="x", side="bottom")
        self.cal_frame.bind("<Configure>", self._on_cal_resize)

    def change_language(self, _event=None):
        codes = ["auto"] + list(LANGUAGE_NAMES)
        idx = self.lang_combo.current()
        choice = codes[idx] if 0 <= idx < len(codes) else "auto"
        settings = load_settings()
        settings["language"] = choice
        save_settings(settings)
        set_language(choice)
        for w in self.winfo_children():
            w.destroy()
        self._build_ui()
        if self.path:
            self.load(self.path, keep_month=True)
        else:
            self.title(APP_NAME)
        self.show_view()

    # ---- toolbar -------------------------------------------------------
    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(fill="x")
        ttk.Button(bar, text=T["open"], command=self.open_dialog).pack(side="left")
        ttk.Button(bar, text=T["export"], command=self.export_xlsx).pack(side="left", padx=(4, 0))
        ttk.Button(bar, text=T["about"], command=self.show_about).pack(side="right")
        codes = ["auto"] + list(LANGUAGE_NAMES)
        chosen = load_settings().get("language", "auto")
        self.lang_combo = ttk.Combobox(bar, state="readonly", width=11,
                                       values=[T["lang_auto"]] + list(LANGUAGE_NAMES.values()))
        self.lang_combo.current(codes.index(chosen) if chosen in codes else 0)
        self.lang_combo.pack(side="right", padx=(0, 10))
        self.lang_combo.bind("<<ComboboxSelected>>", self.change_language)
        ttk.Label(bar, text=T["language"]).pack(side="right", padx=(0, 4))
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        for key, text in (("cal", T["calendar"]), ("list", T["list"])):
            ttk.Radiobutton(bar, text=text, value=key, variable=self.view, style="Toolbutton",
                            command=self.show_view).pack(side="left", padx=1)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)

        self.nav = ttk.Frame(bar)
        ttk.Button(self.nav, text="◀", width=3, command=lambda: self.shift_month(-1)).pack(side="left")
        ttk.Button(self.nav, text=T["today"], command=self.go_today).pack(side="left", padx=2)
        ttk.Button(self.nav, text="▶", width=3, command=lambda: self.shift_month(1)).pack(side="left")
        self.month_lbl = ttk.Label(self.nav, font=self.font_title, padding=(12, 0))
        self.month_lbl.pack(side="left")

        self.search_box = ttk.Frame(bar)
        ttk.Label(self.search_box, text=T["search"]).pack(side="left")
        entry = ttk.Entry(self.search_box, textvariable=self.search_var, width=30)
        entry.pack(side="left", padx=4)
        ttk.Checkbutton(self.search_box, text=T["hide_past"], variable=self.hide_past,
                        command=self.render_list).pack(side="left", padx=8)

    # ---- list view -----------------------------------------------------
    def _build_list(self):
        cols = ("date", "time", "title", "loc")
        self.tree = ttk.Treeview(self.list_frame, columns=cols, show="headings", selectmode="browse")
        for c, text, w, stretch in (("date", T["col_date"], 130, False),
                                    ("time", T["col_time"], 200, False),
                                    ("title", T["col_title"], 380, True),
                                    ("loc", T["col_loc"], 260, True)):
            self.tree.heading(c, text=text, anchor="w")
            self.tree.column(c, width=w, stretch=stretch, anchor="w")
        sb = ttk.Scrollbar(self.list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.tag_configure("past", foreground=MUTED)
        self.tree.tag_configure("today", background=TODAY_BG)
        self.tree.tag_configure("newday", font=self.font_bold)
        self.tree.bind("<Double-1>", self._on_tree_open)
        self.tree.bind("<Return>", self._on_tree_open)
        self.tree_items = {}

    def _on_tree_open(self, _event=None):
        sel = self.tree.selection()
        if sel and sel[0] in self.tree_items:
            self.show_details(self.tree_items[sel[0]])

    @staticmethod
    def _is_past(o, now):
        return o.end_dt() <= now if o.end_dt() > o.start_dt() else o.start_dt() < now

    def list_filtered(self):
        """Occurrences matching the list view's search box and 'hide past' option."""
        q = self.search_var.get().strip().casefold()
        now = datetime.now()
        out = []
        for o in self.occs:
            if self.hide_past.get() and self._is_past(o, now):
                continue
            if q:
                hay = " ".join((o.info["summary"], o.info["location"], o.info["description"])).casefold()
                if q not in hay:
                    continue
            out.append(o)
        return out

    def render_list(self):
        self.tree.delete(*self.tree.get_children())
        self.tree_items = {}
        now = datetime.now()
        today = now.date()
        first_upcoming = None
        prev_day = None
        shown = 0
        for o in self.list_filtered():
            past = self._is_past(o, now)
            d = o.start_dt().date()
            if o.all_day:
                last = o.end - timedelta(days=1)
                t = T["all_day"] if last <= o.start else f"{T['all_day']} → {fmt_date_short(last)}"
            elif o.end.date() != o.start.date() and o.end > o.start:
                t = f"{fmt_time(o.start)} → {o.end.day:02d} {T['months_short'][o.end.month - 1]} {fmt_time(o.end)}"
            elif o.end > o.start:
                t = f"{fmt_time(o.start)} – {fmt_time(o.end)}"
            else:
                t = fmt_time(o.start)
            tags = []
            if past:
                tags.append("past")
            if d == today:
                tags.append("today")
            date_txt = fmt_date_short(d) if d != prev_day else ""
            prev_day = d
            iid = self.tree.insert("", "end", values=(date_txt, t, o.info["summary"],
                                                      o.info["location"].replace("\n", ", ")),
                                   tags=tags)
            self.tree_items[iid] = o
            if first_upcoming is None and not past:
                first_upcoming = (iid, shown)
            shown += 1
        if first_upcoming and shown:
            self.tree.yview_moveto(max(0, first_upcoming[1] - 2) / shown)

    # ---- calendar view -------------------------------------------------
    def _on_cal_resize(self, event):
        nweeks = len(calendar.Calendar(0).monthdatescalendar(self.year, self.month))
        header = self.font_head.metrics("linespace") + 8
        row_h = (event.height - header) / max(nweeks, 1)
        chip_h = self.font.metrics("linespace") + 3
        day_h = self.font_bold.metrics("linespace") + 6
        cap = max(1, int((row_h - day_h) // chip_h))
        if cap != self._cap:
            self._cap = cap
            if self._resize_job:
                self.after_cancel(self._resize_job)
            self._resize_job = self.after(120, self.render_calendar)

    def _on_wheel(self, event):
        if self.view.get() != "cal":
            return
        if not str(event.widget).startswith(str(self.cal_frame)):
            return
        self.shift_month(-1 if event.delta > 0 else 1)

    def shift_month(self, delta):
        if self.view.get() != "cal":
            return
        m = self.month - 1 + delta
        self.year += m // 12
        self.month = m % 12 + 1
        self.render_calendar()

    def go_today(self):
        t = date.today()
        self.year, self.month = t.year, t.month
        if self.view.get() == "cal":
            self.render_calendar()

    def render_calendar(self):
        f = self.cal_frame
        for w in f.winfo_children():
            w.destroy()
        self.month_lbl.config(text=f"{T['months'][self.month - 1].capitalize()} {self.year}")
        weeks = calendar.Calendar(firstweekday=0).monthdatescalendar(self.year, self.month)
        for c in range(7):
            f.columnconfigure(c, weight=1, uniform="col")
            tk.Label(f, text=T["days_short"][c].upper(), bg=BG, fg=MUTED, font=self.font_head,
                     pady=4).grid(row=0, column=c, sticky="nsew")
        f.rowconfigure(0, weight=0)
        for r in range(1, 7):
            f.rowconfigure(r, weight=1 if r <= len(weeks) else 0, uniform="row" if r <= len(weeks) else "")
        today = date.today()
        cap = self._cap
        for r, week in enumerate(weeks, start=1):
            for c, d in enumerate(week):
                in_month = d.month == self.month
                bg = TODAY_BG if d == today else (BG if in_month else OTHER_MONTH_BG)
                cell = tk.Frame(f, bg=bg, width=1, height=1, highlightthickness=1,
                                highlightbackground=GRID, highlightcolor=GRID)
                cell.grid(row=r, column=c, sticky="nsew")
                cell.pack_propagate(False)
                num_fg = "#ffffff" if d == today else ("#202124" if in_month else MUTED)
                num_bg = ACCENT if d == today else bg
                head = tk.Frame(cell, bg=bg)
                head.pack(fill="x")
                num = tk.Label(head, text=str(d.day), bg=num_bg, fg=num_fg, font=self.font_bold, padx=4)
                num.pack(side="left", padx=3, pady=2)
                day_occs = self.by_day.get(d, [])
                show = day_occs if len(day_occs) <= cap else day_occs[:max(cap - 1, 0)]
                for o in show:
                    self._chip(cell, o, d, bg)
                if len(show) < len(day_occs):
                    more = tk.Label(cell, text=T["more"].format(n=len(day_occs) - len(show)),
                                    bg=bg, fg=MUTED, font=self.font_bold, anchor="w", padx=4, cursor="hand2")
                    more.pack(fill="x")
                    more.bind("<Button-1>", lambda e, dd=d: self.show_day(dd))
                for w in (cell, head, num):
                    w.bind("<Button-1>", lambda e, dd=d: self.show_day(dd))

    def _chip(self, parent, o, d, bg):
        color = o.info["color"]
        multi_day = len(o.days()) > 1
        if o.all_day or multi_day:
            label = o.info["summary"]
            if not o.all_day and o.start.date() == d:
                label = f"{fmt_time(o.start)} {label}"
            w = tk.Label(parent, text=label, bg=color, fg="#ffffff", font=self.font,
                         anchor="w", padx=4, width=1, cursor="hand2")
            w.pack(fill="x", padx=2, pady=1)
        else:
            w = tk.Label(parent, text=f"● {fmt_time(o.start)} {o.info['summary']}", bg=bg,
                         fg=color, font=self.font, anchor="w", padx=2, width=1, cursor="hand2")
            w.pack(fill="x", padx=2)
        w.bind("<Button-1>", lambda e, oo=o: self.show_details(oo))

    # ---- popups --------------------------------------------------------
    def _popup(self, title, width=520, height=420):
        top = tk.Toplevel(self)
        top.title(title)
        top.transient(self)
        top.configure(bg=BG)
        x = self.winfo_rootx() + (self.winfo_width() - width) // 2
        y = self.winfo_rooty() + (self.winfo_height() - height) // 3
        top.geometry(f"{width}x{height}+{max(x, 0)}+{max(y, 0)}")
        top.bind("<Escape>", lambda e: top.destroy())
        return top

    def show_day(self, d):
        occs = self.by_day.get(d, [])
        if not occs:
            return
        top = self._popup(fmt_date_long(d), 460, min(120 + 34 * len(occs), 600))
        tk.Label(top, text=fmt_date_long(d), font=self.font_title, bg=BG, anchor="w",
                 padx=12, pady=8).pack(fill="x")
        frame = tk.Frame(top, bg=BG)
        frame.pack(fill="both", expand=True, padx=12)
        for o in occs:
            if o.all_day or len(o.days()) > 1:
                when = T["all_day"] if o.all_day else f"{fmt_time(o.start)} →"
            else:
                when = f"{fmt_time(o.start)}–{fmt_time(o.end)}" if o.end > o.start else fmt_time(o.start)
            row = tk.Frame(frame, bg=BG, cursor="hand2")
            row.pack(fill="x", pady=2)
            bar = tk.Frame(row, bg=o.info["color"], width=4)
            bar.pack(side="left", fill="y")
            l1 = tk.Label(row, text=when, bg=BG, fg=MUTED, font=self.font, width=16, anchor="w")
            l1.pack(side="left", padx=6)
            l2 = tk.Label(row, text=o.info["summary"], bg=BG, font=self.font_bold, anchor="w")
            l2.pack(side="left", fill="x")
            for w in (row, bar, l1, l2):
                w.bind("<Button-1>", lambda e, oo=o: self.show_details(oo))
        ttk.Button(top, text=T["close"], command=top.destroy).pack(pady=8)

    def show_details(self, o):
        info = o.info
        top = self._popup(info["summary"], 560, 460)
        head = tk.Frame(top, bg=BG)
        head.pack(fill="x", padx=12, pady=(12, 4))
        tk.Frame(head, bg=info["color"], width=6).pack(side="left", fill="y")
        tk.Label(head, text=info["summary"], font=self.font_title, bg=BG, anchor="w",
                 justify="left", wraplength=500, padx=8).pack(side="left", fill="x")

        txt = tk.Text(top, wrap="word", relief="flat", bg=BG, font=("Segoe UI", 10),
                      padx=14, pady=6, highlightthickness=0)
        sb = ttk.Scrollbar(top, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        ttk.Button(top, text=T["close"], command=top.destroy).pack(side="bottom", pady=8)
        sb.pack(side="right", fill="y")
        txt.pack(fill="both", expand=True)
        txt.tag_configure("label", font=("Segoe UI", 9, "bold"), foreground=MUTED, spacing1=8)

        def add(label, value):
            if value:
                txt.insert("end", label.upper() + "\n", "label")
                txt.insert("end", value + "\n")

        add(T["when"], self._when_text(o) + (f"\n↻ {T['recurring']}" if info["recurring"] else ""))
        add(T["location"], info["location"])
        add(T["organizer"], info["organizer"])
        add(T["attendees"], "\n".join(info["attendees"]))
        add(T["status"], info["status"].capitalize())
        add(T["url"], info["url"])
        add(T["description"], info["description"])
        txt.configure(state="disabled")

    def _when_text(self, o):
        if o.all_day:
            last = o.end - timedelta(days=1)
            if last <= o.start:
                return f"{fmt_date_long(o.start)} · {T['all_day']}"
            return f"{fmt_date_long(o.start)} – {fmt_date_long(last)}"
        if o.end <= o.start:
            return f"{fmt_date_long(o.start.date())}, {fmt_time(o.start)}"
        if o.end.date() == o.start.date():
            return f"{fmt_date_long(o.start.date())}, {fmt_time(o.start)} – {fmt_time(o.end)}"
        return (f"{fmt_date_long(o.start.date())}, {fmt_time(o.start)} –\n"
                f"{fmt_date_long(o.end.date())}, {fmt_time(o.end)}")

    # ---- loading / switching ------------------------------------------
    def show_about(self):
        top = self._popup(T["about"], 400, 250)
        top.resizable(False, False)
        tk.Label(top, text=APP_NAME, font=self.font_title, bg=BG).pack(pady=(18, 0))
        tk.Label(top, text=f"{T['version']} {APP_VERSION}", font=self.font, bg=BG, fg=MUTED).pack()
        tk.Label(top, text=T["about_text"], font=self.font, bg=BG).pack(pady=(10, 6))
        tk.Label(top, text=f"{T['author']}: {AUTHOR}", font=self.font_bold, bg=BG).pack()
        link = tk.Label(top, text=AUTHOR_URL, font=tkfont.Font(family="Segoe UI", size=9, underline=True),
                        bg=BG, fg=ACCENT, cursor="hand2")
        link.pack()
        link.bind("<Button-1>", lambda e: webbrowser.open(AUTHOR_URL))
        tk.Label(top, text=COPYRIGHT, font=self.font, bg=BG, fg=MUTED).pack(pady=(8, 0))
        ttk.Button(top, text=T["close"], command=top.destroy).pack(pady=10)
        top.grab_set()

    def open_dialog(self):
        path = filedialog.askopenfilename(
            parent=self,
            filetypes=[(T["cal_files"], "*.ics *.ical *.ifb *.icalendar"), (T["all_files"], "*.*")])
        if path:
            self.load(path)

    def export_xlsx(self):
        # The list view exports what it shows (search / hide-past filters); the calendar exports everything.
        occs = self.list_filtered() if self.view.get() == "list" else self.occs
        if not occs:
            messagebox.showinfo(APP_NAME, T["nothing_to_export"], parent=self)
            return
        base = os.path.splitext(os.path.basename(self.path))[0] if self.path else "calendar"
        path = filedialog.asksaveasfilename(parent=self, defaultextension=".xlsx",
                                            initialfile=base + ".xlsx",
                                            filetypes=[("Excel", "*.xlsx"), (T["all_files"], "*.*")])
        if not path:
            return
        yes_no = {True: T["yes"], False: T["no"]}
        rows = []
        for o in occs:
            i = o.info
            if o.all_day:
                start_d, start_t = o.start, None
                end_d, end_t = max(o.end - timedelta(days=1), o.start), None
            else:
                start_d, start_t = o.start.date(), o.start.time()
                end_d, end_t = o.end.date(), o.end.time()
            rows.append([start_d, start_t, end_d, end_t, yes_no[o.all_day], i["summary"],
                         i["location"], i["description"], i["organizer"], "; ".join(i["attendees"]),
                         i["status"].capitalize(), yes_no[bool(i["recurring"])], i["url"]])
        try:
            write_xlsx(path, T["sheet"], T["xlsx_cols"], rows,
                       [12, 10, 12, 10, 10, 40, 30, 60, 25, 40, 12, 11, 40])
        except Exception as e:  # noqa: BLE001 - show any failure to the user
            messagebox.showerror(T["error"], T["error_write"].format(e=e), parent=self)
            return
        messagebox.showinfo(APP_NAME, T["export_done"].format(n=len(rows), path=path), parent=self)

    def load(self, path, keep_month=False):
        try:
            calname, occs = load_ics_file(path)
        except Exception as e:  # noqa: BLE001 - show any failure to the user
            messagebox.showerror(T["error"], T["error_read"].format(e=e), parent=self)
            return
        self.path, self.calname, self.occs = path, calname, occs
        self.by_day = {}
        for o in occs:
            for d in o.days():
                self.by_day.setdefault(d, []).append(o)
        self.title(f"{calname or os.path.basename(path)} — {APP_NAME}")
        status = T["events"].format(n=len(occs))
        self.status.config(text=f"{os.path.basename(path)}   ·   {calname + '   ·   ' if calname else ''}{status}")
        # jump to the month with the next upcoming event (or the last one)
        today = date.today()
        if keep_month:
            pass
        elif occs and not any(d.year == today.year and d.month == today.month for d in self.by_day):
            nxt = next((o for o in occs if o.start_dt().date() >= today), occs[-1])
            d = nxt.start_dt().date()
            self.year, self.month = d.year, d.month
        else:
            self.year, self.month = today.year, today.month
        self.render_calendar()
        self.render_list()

    def show_view(self):
        if self.view.get() == "cal":
            self.list_frame.pack_forget()
            self.search_box.pack_forget()
            self.nav.pack(side="left")
            self.cal_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
            self.render_calendar()
        else:
            self.cal_frame.pack_forget()
            self.nav.pack_forget()
            self.search_box.pack(side="left")
            self.list_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
            self.tree.focus_set()


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    path = sys.argv[1] if len(sys.argv) > 1 and os.path.isfile(sys.argv[1]) else None
    App(path).mainloop()


if __name__ == "__main__":
    main()
