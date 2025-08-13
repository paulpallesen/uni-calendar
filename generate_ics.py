#!/usr/bin/env python3
import argparse
from openpyxl import load_workbook
from datetime import datetime
import hashlib
import os, sys

AUS_TZ_VTIMEZONE = """BEGIN:VTIMEZONE
TZID:Australia/Sydney
BEGIN:STANDARD
DTSTART:19700405T030000
TZOFFSETFROM:+1100
TZOFFSETTO:+1000
TZNAME:AEST
RRULE:FREQ=YEARLY;BYMONTH=4;BYDAY=1SU
END:STANDARD
BEGIN:DAYLIGHT
DTSTART:19701004T020000
TZOFFSETFROM:+1000
TZOFFSETTO:+1100
TZNAME:AEDT
RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=1SU
END:DAYLIGHT
END:VTIMEZONE
"""

def fmt_local(dt: datetime) -> str:
    return dt.strftime("%Y%m%dT%H%M%S")

def parse_date(s):
    if not s:
        return None
    return datetime.strptime(s.strip(), "%Y-%m-%d")

def parse_time(s):
    if not s:
        return None
    return datetime.strptime(s.strip(), "%H:%M")

def parse_datetime(date_str, time_str):
    d = parse_date(date_str)
    if d is None:
        return None
    if not time_str:
        return datetime(d.year, d.month, d.day, 0, 0, 0)
    t = parse_time(time_str)
    return datetime(d.year, d.month, d.day, t.hour, t.minute, 0)

def make_uid(fields):
    h = hashlib.sha1("|".join(fields).encode("utf-8")).hexdigest()[:16]
    return f"{h}@youruni"

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--sheet", default="Events")
    ap.add_argument("--out", default="docs/calendar.ics")
    ap.add_argument("--tz", default="Australia/Sydney")
    args = ap.parse_args()

    wb = load_workbook(args.xlsx, data_only=True)
    if args.sheet not in wb.sheetnames:
        print(f"Sheet '{args.sheet}' not found", file=sys.stderr)
        sys.exit(1)
    ws = wb[args.sheet]

    headers = [ (c.value or "").strip() for c in next(ws.iter_rows(min_row=1, max_row=1, values_only=False)) ]
    def idx(name):
        try:
            return headers.index(name)
        except ValueError:
            return -1

    col_UID = idx("UID (optional)")
    col_Course = idx("CourseCode")
    col_Title = idx("Title")
    col_Cat = idx("Category (optional)")
    col_SDate = idx("StartDate (YYYY-MM-DD)")
    col_STime = idx("StartTime (HH:MM)")
    col_EDate = idx("EndDate (YYYY-MM-DD, optional)")
    col_ETime = idx("EndTime (HH:MM)")
    col_TZ = idx("Timezone (default Australia/Sydney)")
    col_Loc = idx("Location")
    col_Desc = idx("Description (optional)")
    col_RRULE = idx("RRULE (optional) e.g. FREQ=WEEKLY;COUNT=12;BYDAY=MO")
    col_EXDATE = idx("EXDATE (optional, comma-separated, ISO dates/times)")
    col_URL = idx("URL (optional)")
    col_SEQ = idx("Sequence (optional integer)")

    now_utc = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR",
        "PRODID:-//YourUni//Class Feeds 1.0//EN",
        "VERSION:2.0",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        AUS_TZ_VTIMEZONE.strip()
    ]

    for r in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in r):
            continue

        uid = (r[col_UID] or "").strip() if col_UID >= 0 and r[col_UID] else ""
        course = (r[col_Course] or "").strip() if col_Course >= 0 and r[col_Course] else ""
        title = (r[col_Title] or "").strip() if col_Title >= 0 and r[col_Title] else ""
        cat = (r[col_Cat] or "").strip() if col_Cat >= 0 and r[col_Cat] else ""
        sdate = (r[col_SDate] or "").strip() if col_SDate >= 0 and r[col_SDate] else ""
        stime = (r[col_STime] or "").strip() if col_STime >= 0 and r[col_STime] else ""
        edate = (r[col_EDate] or "").strip() if col_EDate >= 0 and r[col_EDate] else ""
        etime = (r[col_ETime] or "").strip() if col_ETime >= 0 and r[col_ETime] else ""
        tz = (r[col_TZ] or "").strip() if col_TZ >= 0 and r[col_TZ] else args.tz
        location = (r[col_Loc] or "").strip() if col_Loc >= 0 and r[col_Loc] else ""
        desc = (r[col_Desc] or "").strip() if col_Desc >= 0 and r[col_Desc] else ""
        rrule = (r[col_RRULE] or "").strip() if col_RRULE >= 0 and r[col_RRULE] else ""
        exdate_raw = (r[col_EXDATE] or "").strip() if col_EXDATE >= 0 and r[col_EXDATE] else ""
        url = (r[col_URL] or "").strip() if col_URL >= 0 and r[col_URL] else ""
        seq = str(r[col_SEQ]).strip() if col_SEQ >= 0 and r[col_SEQ] is not None else ""

        if not title or not sdate or not stime or not etime:
            continue

        dt_start = parse_datetime(sdate, stime)
        if edate:
            dt_end = parse_datetime(edate, etime)
        else:
            dt_end = parse_datetime(sdate, etime)

        if not uid:
            uid = make_uid([course, title, fmt_local(dt_start), fmt_local(dt_end), location])

        summary = f"{course} — {title}" if course else title

        lines.append("BEGIN:VEVENT")
        lines.append(f"UID:{uid}")
        lines.append(f"DTSTAMP:{now_utc}")
        if seq:
            lines.append(f"SEQUENCE:{seq}")
        lines.append(f"DTSTART;TZID={tz}:{fmt_local(dt_start)}")
        lines.append(f"DTEND;TZID={tz}:{fmt_local(dt_end)}")
        lines.append(f"SUMMARY:{summary}")
        if location:
            lines.append(f"LOCATION:{location}")
        if desc:
            lines.append(f"DESCRIPTION:{desc.replace('\\n', '\\n')}")
        if url:
            lines.append(f"URL:{url}")
        cats = []
        if course: cats.append(course)
        if cat: cats.append(cat)
        if location: cats.append(location)
        if cats:
            lines.append(f"CATEGORIES:{','.join(cats)}")
        if rrule:
            lines.append(f"RRULE:{rrule}")
        if exdate_raw:
            parts = [p.strip() for p in exdate_raw.split(",") if p.strip()]
            ex_vals = []
            for p in parts:
                try:
                    if "T" in p:
                        dt = datetime.fromisoformat(p)
                        ex_vals.append(fmt_local(dt))
                    else:
                        d = datetime.fromisoformat(p + "T00:00:00")
                        ex_vals.append(fmt_local(d))
                except Exception:
                    pass
            if ex_vals:
                lines.append(f"EXDATE;TZID={tz}:{','.join(ex_vals)}")
        lines.append("END:VEVENT")

    lines.append("END:VCALENDAR")

    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    with open(args.out, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(lines))

    print(f"Wrote {args.out}")

if __name__ == "__main__":
    main()
