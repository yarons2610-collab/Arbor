#!/usr/bin/env python3
"""Export an Arbor workspace to a presentable Excel workbook.

Runs outside the browser, so the twice-daily scheduled export does not need
Arbor to be open. Reads the workspace either from the synced Gist (the source
of truth) or from a JSON file, and writes a real .xlsx — not the SpreadsheetML
the in-app button used to produce, which made Excel warn about the file format
on every open.

  arbor_export.py --gist <id> --out <dir>     # what the scheduled job runs
  arbor_export.py --json <file> --out <dir>
  arbor_export.py --json <file> --xlsx <file>

Each run also keeps the raw JSON it exported from, so the folder doubles as a
backup that does not depend on a browser, on GitHub, or on this script.
"""

import argparse
import datetime as dt
import json
import os
import re
import subprocess
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ── look ──────────────────────────────────────────────────────────────────────
# Arbor's own palette, muted for paper: the app is dark, a spreadsheet is not.
INK = "1F2733"
HEAD_BG = "22303F"
RULE = "D9DDE5"
BAND = "F4F6F9"
GREEN = "1B7F4B"   # alive
GREY = "6B7685"    # no longer alive
AMBER = "9A6B00"   # practice
RED = "B3243F"     # excluded

HEAD_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEAD_FILL = PatternFill("solid", fgColor=HEAD_BG)
BODY_FONT = Font(name="Calibri", size=11, color=INK)
MONO_FONT = Font(name="Consolas", size=10, color=INK)
TITLE_FONT = Font(name="Calibri", size=18, bold=True, color=INK)
SUB_FONT = Font(name="Calibri", size=11, color="5B6678")
THIN = Side(style="thin", color=RULE)
CELL_BORDER = Border(bottom=THIN)

DATE_FMT = "dd/mm/yyyy"          # day/month/year everywhere, as in the app
DATETIME_FMT = "dd/mm/yyyy hh:mm"

ISO_DATE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})$")
ISO_DATETIME = re.compile(r"^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})")


# ── small helpers ─────────────────────────────────────────────────────────────
def as_date(v):
    """'2026-03-17' / '2026-03-17T09:00' -> a real date, so Excel can sort and
    filter it. Anything else is handed back untouched."""
    if not isinstance(v, str):
        return v
    m = ISO_DATETIME.match(v)
    if m:
        y, mo, d, h, mi = (int(x) for x in m.groups()[:5])
        return dt.datetime(y, mo, d, h, mi)
    m = ISO_DATE.match(v)
    if m:
        return dt.date(*(int(x) for x in m.groups()))
    return v


def ms_to_dt(ms):
    if not ms:
        return ""
    try:
        return dt.datetime.fromtimestamp(ms / 1000)
    except (OverflowError, OSError, ValueError):
        return ""


TAG_RE = re.compile(r"<[^>]+>")
WS_RE = re.compile(r"[ \t\r\f\v]+")
ENTITIES = {"&nbsp;": " ", "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&#39;": "'"}


def plain(html):
    """Rich-text field -> readable text. <div>/<br> become line breaks so a
    multi-line note stays multi-line in the cell."""
    if not html:
        return ""
    s = re.sub(r"<(br|/div|/p|/li)[^>]*>", "\n", str(html), flags=re.I)
    s = TAG_RE.sub("", s)
    for k, v in ENTITIES.items():
        s = s.replace(k, v)
    s = WS_RE.sub(" ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def natural_key(v):
    """'M2' before 'M10'. Case-insensitive, digits compared as numbers."""
    s = "" if v is None else str(v)
    if isinstance(v, (dt.date, dt.datetime)):
        return [(0, "", v.toordinal())]
    return [(1, "", int(p)) if p.isdigit() else (2, p.lower(), 0)
            for p in re.split(r"(\d+)", s) if p != ""] or [(3, "", 0)]


def sort_rows(rows, keys):
    for k in reversed(keys or []):
        rows.sort(key=lambda r: natural_key(r[k]) if k < len(r) else [])
    return rows


def yes(v):
    return "Yes" if v else ""


# the app's own words (REPEATS and DOW in index.html), so both exports agree
DOW = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]
PLAIN_REPEATS = {"day": "every day", "week": "every week", "2week": "every 2 weeks",
                 "month": "every month", "year": "every year"}


def repeat_label(r):
    """The repeat rule in words, matching what the app shows."""
    if not r or not r.get("every"):
        return ""
    if r["every"] != "custom":
        base = PLAIN_REPEATS.get(r["every"], str(r["every"]))
    else:
        freq = r.get("freq") or "week"
        iv = max(1, int(r.get("interval") or 1))
        if freq == "week":
            days = ", ".join(DOW[i] for i in sorted(r.get("days") or []) if 0 <= i < 7)
            base = (f"{days} every {iv} weeks" if iv > 1 else days) if days else "custom"
        else:
            unit = {"day": "day", "month": "month", "year": "year"}.get(freq, freq)
            base = f"every {iv} {unit}s" if iv > 1 else f"every {unit}"
    until = as_date(r.get("until"))
    if isinstance(until, dt.date):
        base += f", until {until.strftime('%d/%m/%Y')}"
    return base


# ── record kinds ──────────────────────────────────────────────────────────────
# Mirrors RECORD_KINDS in index.html: in Arbor 2.0 the kinds are defined in
# code, not in the workspace's config, so the exporter carries its own copy.
RECORD_KINDS = {
    "weight": ("Weighing", [("grams", "Weight (g)", "number"), ("ration", "Ration (g)", "number")]),
    "surgery": ("Surgery", [("duration", "Duration (min)", "number"), ("operator", "Operator", "text"),
                            ("complications", "Complications", "text")]),
    "postop": ("Post-op check", [("given", "Post-op given", "yesno"), ("condition", "Condition", "text")]),
    "training": ("Behaviour session", [("stage", "Stage", "text"), ("trials", "Trials", "number"),
                                       ("success", "Success %", "number"), ("length", "Session length (min)", "number"),
                                       ("dataAt", "Data location", "text")]),
    "microscopy": ("Imaging session", [("purpose", "Purpose", "text"), ("laser", "Laser power", "text"),
                                       ("channels", "Channels", "text"), ("dataAt", "Data location", "text"),
                                       ("status", "Analysis", "text")]),
    "perfusion": ("Perfusion", [("operator", "Done by", "text")]),
    "histology": ("Slicing", [("operator", "Sliced by", "text"), ("plane", "Plane", "text"),
                              ("thickness", "Thickness (µm)", "number"), ("storage", "Stored at", "text")]),
    "histo": ("Histology", [("operator", "Done by", "text"), ("dataAt", "Scans at", "text")]),
}
FOV_SETTINGS = [("depth", "Depth (µm)"), ("zoom", "Zoom"), ("planes", "Planes"),
                ("frames", "Frames / plane"), ("res", "Resolution")]
RESTRICTION_TARGET = 90
BEHAVIOUR_TASKS = {"Lever": ["Threshold", "Sound frequency (Hz)", "Table angle"]}


def ration_of(d):
    """The day's ration — 2.0 stores it as "ration", 1.0 in a custom field."""
    v = d.get("ration") if d.get("ration") is not None else d.get("f_a847sm")
    return "" if v in (None, "") else v


def note_of(d):
    """Same reading as diaryNoteOf in the app."""
    if d.get("note") is not None:
        return plain(d.get("note"))
    if d.get("label") is not None:
        return ": ".join(x for x in (d.get("label"), d.get("text")) if x)
    return plain(d.get("text") or "")


# ── workspace model ───────────────────────────────────────────────────────────
class Workspace:
    """Reads an Arbor workspace the way the app does: resolving the links that
    connect animals, experiments and projects, so the sheets name things
    instead of showing ids."""

    def __init__(self, db):
        self.db = db
        self.animals = db.get("animals") or []
        self.experiments = db.get("experiments") or []
        self.projects = db.get("projects") or []
        self.schedule = db.get("schedule") or []
        self.plans = db.get("plans") or []
        self.notes = db.get("notes") or []
        self.config = db.get("config") or {}
        self.proj_by_id = {p["id"]: p for p in self.projects if p.get("id")}
        self.exp_by_id = {e["id"]: e for e in self.experiments if e.get("id")}
        self.diary_types = {t.get("key"): t for t in (self.config.get("diaryTypes") or [])}

    # -- naming
    @staticmethod
    def proj_name(p):
        return (p or {}).get("name") or "(unnamed project)"

    @staticmethod
    def exp_title(e):
        return (e or {}).get("title") or ("(unnamed workstream)" if is_workstream(e) else "(unnamed experiment)")

    def project_of_exp(self, e):
        return self.proj_by_id.get((e or {}).get("projectId"))

    # -- membership, mirroring animalSteps()/animalGroups()
    def experiments_of_animal(self, aid):
        out = []
        for e in self.experiments:
            in_group = (e.get("groups") or {}).get(aid)
            linked = any(l.get("kind") == "animal" and l.get("id") == aid
                         for l in (e.get("links") or []))
            if in_group or linked:
                out.append(e)
        return out

    def group_name(self, e, aid):
        gid = (e.get("groups") or {}).get(aid)
        if not gid:
            return ""
        for g in e.get("groupDefs") or []:
            if g.get("id") == gid:
                return g.get("name") or ""
        return ""

    def where(self, aid):
        bits = []
        for e in self.experiments_of_animal(aid):
            g = self.group_name(e, aid)
            bits.append(self.exp_title(e) + (f" ({g})" if g else ""))
        return ", ".join(bits)

    def animals_of_exp(self, e):
        ids = set((e.get("groups") or {}).keys())
        ids |= {l["id"] for l in (e.get("links") or [])
                if l.get("kind") == "animal" and l.get("id")}
        names = [a.get("name") or "" for a in self.animals if a.get("id") in ids]
        names.sort(key=natural_key)
        return names

    def attached_to(self, kind, rid):
        """Which experiments/projects link to this note or library record."""
        names = []
        for e in self.experiments:
            if any(l.get("kind") == kind and l.get("id") == rid for l in (e.get("links") or [])):
                names.append(self.exp_title(e))
        for p in self.projects:
            if any(l.get("kind") == kind and l.get("id") == rid for l in (p.get("links") or [])):
                names.append(self.proj_name(p))
        return ", ".join(names)

    # -- animal facts
    @staticmethod
    def status(a):
        if a.get("alive"):
            return "Alive"
        return {"sacrificed": "Sacrificed",
                "early_death": "Early death"}.get(a.get("deathKind"), "No longer alive")

    @staticmethod
    def latest_weight(a):
        """The newest weighing that carries a value. This used to read
        a.weights - the map weighings were migrated OUT of, and which the app
        has kept frozen since only as a fallback for old devices - so every
        weighing logged after that migration was missing from this sheet.
        Mirrors latestWeightEntry() in the app."""
        w = Workspace.weight_entries(a)
        if w:
            return w[-1]["data"]["grams"]
        legacy = a.get("weights") or {}
        if isinstance(legacy, dict) and legacy:
            try:
                return legacy[max(legacy.keys())]
            except (ValueError, TypeError):
                return ""
        return ""

    # ---- the animal model, mirrored from the app -------------------------
    # This exporter runs on its own, twice a day, without the app open, so it
    # cannot ask the app what an animal is. These follow the app's own rules
    # one for one (entryHappened, baselineAt, pctOfBaseline, animalViruses...)
    # and must change when those do. Kept together so that is one place.

    @staticmethod
    def entry_happened(e):
        d = as_date(e.get("date"))
        if not isinstance(d, dt.date) or d > dt.date.today():
            return False
        if e.get("fromPlan") or e.get("fromTask"):
            return bool(e.get("done")) or Workspace.has_content(e)
        return True

    @staticmethod
    def has_content(e):
        if e.get("done"):
            return True
        for v in (e.get("data") or {}).values():
            if v is None or v == "":
                continue
            if isinstance(v, (list, dict)) and not v:
                continue
            return True
        return False

    @staticmethod
    def weight_entries(a):
        out = []
        for e in a.get("diary") or []:
            g = (e.get("data") or {}).get("grams")
            if e.get("type") == "weight" and g not in (None, "") and e.get("date"):
                out.append(e)
        out.sort(key=lambda e: (str(e.get("date")), e.get("created") or 0))
        return out

    @staticmethod
    def baseline_at(a, date):
        """The baseline in force on a date: the animal's own, set from that
        date or earlier, else its first weighing."""
        d = str(date)[:10]
        own = sorted((b for b in (a.get("baselines") or [])
                      if b.get("date") and b.get("grams") not in (None, "")
                      and str(b["date"])[:10] <= d),
                     key=lambda b: str(b["date"]))
        if own:
            return float(own[-1]["grams"]), str(own[-1]["date"])[:10], True
        w = Workspace.weight_entries(a)
        if w:
            return float(w[0]["data"]["grams"]), str(w[0]["date"])[:10], False
        return None

    @staticmethod
    def restrictions(a):
        return sorted((r for r in (a.get("restrictions") or []) if r and r.get("start")),
                      key=lambda r: str(r["start"]))

    @staticmethod
    def restriction_at(a, date):
        d = str(date)[:10]
        live = [r for r in Workspace.restrictions(a)
                if str(r["start"])[:10] <= d and (not r.get("end") or str(r["end"])[:10] > d)]  # end = first day off
        return live[-1] if live else None

    @staticmethod
    def weighing_context(a, date):
        """Why a weighing was taken — worked out as weighingContext does."""
        d = as_date(date)
        if Workspace.restriction_at(a, date):
            return "restriction weighing"
        if isinstance(d, dt.date):
            for e in a.get("diary") or []:
                s = as_date(e.get("date")) if e.get("type") == "surgery" else None
                if isinstance(s, dt.date) and abs((s - d).days) <= 3:
                    return "pre-op weighing" if d < s else "post-op weighing"
        return "routine weighing"

    @staticmethod
    def fovs(a):
        return [f for f in (a.get("fovs") or []) if f and f.get("id")]

    @staticmethod
    def pct_baseline(a):
        w = Workspace.weight_entries(a)
        if not w:
            return ""
        last = w[-1]
        b = Workspace.baseline_at(a, last["date"])
        if not b or not b[0]:
            return ""
        return round(float(last["data"]["grams"]) / b[0] * 100)

    @staticmethod
    def injections(a):
        return (a.get("surgeryDetail") or {}).get("injections") or []

    @staticmethod
    def joined(values):
        seen = []
        for v in values:
            if v and v not in seen:
                seen.append(v)
        return ", ".join(seen)

    @staticmethod
    def viruses(a):
        return Workspace.joined(i.get("virus") for i in Workspace.injections(a)) or (a.get("virus") or "")

    @staticmethod
    def sites(a):
        return Workspace.joined(i.get("siteRegion") for i in Workspace.injections(a)) or (a.get("targets") or "")

    @staticmethod
    def projections(a):
        return Workspace.joined(i.get("partnerRegion") for i in Workspace.injections(a))

    def next_stage(self, a):
        """The earliest stage a plan scheduled for this animal that has not
        happened yet, read from the entries the plan wrote."""
        pending = [e for e in (a.get("diary") or [])
                   if e.get("fromPlan") and not Workspace.entry_happened(e) and e.get("date")]
        if not pending:
            return ""
        e = min(pending, key=lambda x: str(x["date"]))
        ex = next((x for x in self.experiments if x.get("id") == e.get("fromPlan")), None)
        label = self.diary_label(e.get("type"))
        if ex:
            for pid in [ex.get("planId")] + [g.get("planId") for g in (ex.get("groupDefs") or [])]:
                pl = next((p for p in self.plans if p.get("id") == pid), None)
                st = next((s for s in ((pl or {}).get("stages") or []) if s.get("id") == e.get("planStage")), None)
                if st and st.get("label"):
                    label = st["label"]
                    break
        d = as_date(e["date"])
        late = isinstance(d, dt.date) and d < dt.date.today()
        return f"{label} {d.strftime('%d/%m/%Y') if isinstance(d, dt.date) else e['date']}" + (" (overdue)" if late else "")

    @staticmethod
    def age_weeks(a):
        dob = as_date(a.get("dob"))
        if not isinstance(dob, dt.date):
            return ""
        end = as_date(a.get("sacrificed"))
        end = end if isinstance(end, dt.date) else dt.date.today()
        return round((end - dob).days / 7)

    def surgery_dates(self, a):
        """Every surgery that has happened: the profile's own date plus any
        logged as diary entries. A surgery booked for next week is not one the
        animal has had - counting it made "days since last surgery" negative."""
        out = set()
        if a.get("surgery"):
            d = as_date(a["surgery"])
            if isinstance(d, dt.date) and d <= dt.date.today():
                out.add(d)
        for e in a.get("diary") or []:
            if e.get("type") == "surgery" and Workspace.entry_happened(e):
                d = as_date(e.get("date"))
                if isinstance(d, dt.date):
                    out.add(d)
        return sorted(out)

    def diary_label(self, key):
        if key in RECORD_KINDS:
            return RECORD_KINDS[key][0]
        t = self.diary_types.get(key)
        return (t or {}).get("label") or (key or "").title()

    def belongs_to(self, e):
        """What an entry belongs to, when it was not written by hand: the task
        whose round it is part of, or the experiment stage that scheduled it.
        The app says the same sentence from diaryBelongsTo; a sheet that only
        knew about rounds left every planned stage blank."""
        tid = e.get("fromTask")
        if tid:
            for t in self.schedule:
                if t.get("id") == tid:
                    return t.get("title") or ""
            return ""
        pid = e.get("fromPlan")
        if not pid:
            return ""
        ex = next((x for x in self.experiments if x.get("id") == pid), None)
        if not ex:
            return ""
        wanted = e.get("planStage")
        plan_ids = [ex.get("planId")] + [g.get("planId") for g in (ex.get("groupDefs") or [])]
        for plid in [p for p in plan_ids if p]:
            pl = next((p for p in self.plans if p.get("id") == plid), None)
            for st in ((pl or {}).get("stages") or []):
                if st.get("id") == wanted:
                    return f"{ex.get('title') or 'experiment'} › {stage_name(pl, st)}".strip()
        return ex.get("title") or ""

    @staticmethod
    def kind_summary(e, a):
        """The record's one-line summary — KIND_SUMMARY in the app, part for part."""
        d = record_data(e)
        k = e.get("type")
        has = lambda v: v not in (None, "")
        n = lambda v: f"{v:g}" if isinstance(v, float) else str(v)
        if k == "weight":
            r = ration_of(d)
            if not has(d.get("grams")) and r == "":
                # a 1.0 feeding round ticked without numbers
                parts = ["no weight written" if e.get("done") else ""]
            else:
                parts = [n(d["grams"]) + "g" if has(d.get("grams")) else "", f"fed {n(r)}g" if r != "" else ""]
        elif k == "training":
            st = d.get("settings") or {}
            parts = [d.get("task"), f"session {d['session']}" if d.get("session") else "", d.get("stage"), f"{n(d['trials'])} trials" if has(d.get("trials")) else "",
                     f"{n(d['success'])}%" if has(d.get("success")) else "",
                     f"{n(d['length'])} min" if d.get("length") else "",
                     f"angle {n(st['angle'])}" if st.get("angle") is not None else "",
                     ("imaged" + (" " + d["target"].lower() if d.get("target") else "")) if d.get("imaged") else ""]
        elif k == "microscopy":
            f = d.get("fovs") or []
            parts = [d.get("target"), d.get("purpose"), f"{len(f)} FOV{'' if len(f) == 1 else 's'}" if f else "",
                     d.get("status") if d.get("status") and d.get("status") != "raw" else ""]
        elif k == "surgery":
            parts = ["head-post" if d.get("headpost") else "", "window" if d.get("window") else "",
                     f"{n(d['duration'])} min" if d.get("duration") else "", d.get("operator"),
                     "⚠ " + d["complications"] if d.get("complications") else ""]
        elif k == "postop":
            parts = ["post-op given" if d.get("given") is True else "not given" if d.get("given") is False else "",
                     d.get("condition")]
        elif k == "perfusion":
            pbs = as_date(d.get("pbs"))
            parts = [d.get("operator"), "→ PBS " + pbs.strftime("%d/%m") if isinstance(pbs, dt.date) else "",
                     "not processed" if d.get("notProcessed") else ""]
        elif k == "histology":
            sl = d.get("slides") or []
            parts = [d.get("plane"), f"{n(d['thickness'])}µm" if d.get("thickness") else "",
                     f"{len(sl)} slide{'' if len(sl) == 1 else 's'}, {sum(len(x.get('aps') or []) for x in sl)} slices" if sl else "",
                     d.get("storage")]
        elif k == "histo":
            parts = [", ".join(" ".join(x for x in (f.get("region"), f.get("dir") if f.get("dir") != "labelled" else "") if x)
                               for f in (d.get("findings") or []))]
        else:
            return None
        parts = [str(x) for x in parts if x is not None and str(x).strip() != ""]
        if k == "weight" and has(d.get("grams")):
            b = Workspace.baseline_at(a, e.get("date"))
            w = Workspace.weight_entries(a)
            cur = float(d["grams"])
            is_self = b and ((not b[2] and w and w[0] is e) or
                             (b[2] and b[1] == str(e.get("date"))[:10] and b[0] == cur))
            if b and not is_self:
                delta = cur - b[0]
                pct = f" ({'+' if delta >= 0 else ''}{delta / b[0] * 100:.1f}%)" if b[0] else ""
                parts.append(f"{'+' if delta >= 0 else ''}{delta:.1f}g{pct} vs baseline")
        note = note_of(d)
        if note:
            parts.append(note)
        return " · ".join(parts)

    def diary_detail(self, e, a=None):
        if a is not None and e.get("type") in RECORD_KINDS:
            return Workspace.kind_summary(e, a)
        """The entry's own fields, as 'Label: value' pairs — the same fields the
        app shows under an entry."""
        data = e.get("data") or {}
        kind = RECORD_KINDS.get(e.get("type"))
        t = ({"fields": [{"key": k, "label": l, "type": ty} for k, l, ty in kind[1]]} if kind
             else self.diary_types.get(e.get("type")) or {})
        bits = []
        for f in t.get("fields") or []:
            v = data.get(f.get("key"))
            if v in (None, "", []):
                continue
            if f.get("type") == "table":
                # the rows themselves get their own sheet; here, only how many,
                # because a table dumped into a cell is unreadable either way
                n = len(v) if isinstance(v, list) else len(v or {})
                if n:
                    bits.append(f"{f.get('label') or f.get('key')}: {n} rows")
                continue
            if f.get("type") == "yesno":
                v = "Yes" if v else "No"
            bits.append(f"{f.get('label') or f.get('key')}: {v}")
        # a weighing can be marked as the baseline every later one is measured
        # from; the app says so in its own summary, so the sheet does too
        if e.get("baseline"):
            bits.append("baseline")
        if e.get("type") == "training":
            if data.get("task"):
                bits.insert(0, f"Task: {data['task']}")
            for k, v in (data.get("metrics") or {}).items():
                if v not in (None, ""):
                    bits.append(f"{k}: {v}")
        note = data.get("note")
        if note:
            bits.append(plain(note))
        return " · ".join(bits)


# ── sheet writing ─────────────────────────────────────────────────────────────
class Sheet:
    def __init__(self, name, headers, rows, sort=(), numeric=(), wide=()):
        self.name = name
        self.headers = headers
        self.rows = sort_rows([list(r) for r in rows], list(sort))
        self.numeric = set(numeric)
        self.wide = set(wide)


def write_sheet(wb, sheet, colour_status_col=None):
    ws = wb.create_sheet(sheet.name[:31])
    ws.append(sheet.headers)
    for c, _ in enumerate(sheet.headers, start=1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEAD_FONT
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 26

    for row in sheet.rows:
        ws.append([as_date(v) if isinstance(v, str) else v for v in row])

    ncol = len(sheet.headers)
    widths = [len(str(h)) + 2 for h in sheet.headers]

    for r in range(2, ws.max_row + 1):
        banded = (r % 2 == 0)
        for c in range(1, ncol + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = BODY_FONT
            cell.border = CELL_BORDER
            if banded:
                cell.fill = PatternFill("solid", fgColor=BAND)
            v = cell.value
            if isinstance(v, dt.datetime):
                cell.number_format = DATETIME_FMT
                cell.alignment = Alignment(vertical="top", horizontal="left")
                widths[c - 1] = max(widths[c - 1], 17)
            elif isinstance(v, dt.date):
                cell.number_format = DATE_FMT
                cell.alignment = Alignment(vertical="top", horizontal="left")
                widths[c - 1] = max(widths[c - 1], 12)
            elif isinstance(v, (int, float)):
                cell.alignment = Alignment(vertical="top", horizontal="right")
                widths[c - 1] = max(widths[c - 1], 10)
            else:
                text = "" if v is None else str(v)
                wrap = (c - 1) in sheet.wide or "\n" in text
                cell.alignment = Alignment(vertical="top", wrap_text=wrap)
                longest = max((len(x) for x in text.split("\n")), default=0)
                widths[c - 1] = max(widths[c - 1], min(longest + 2, 60))

        # the one column worth colouring: an animal's status reads at a glance
        if colour_status_col is not None:
            cell = ws.cell(row=r, column=colour_status_col)
            val = str(cell.value or "")
            colour = {"Alive": GREEN}.get(val, GREY)
            cell.font = Font(name="Calibri", size=11, bold=(val == "Alive"), color=colour)

    for c in range(1, ncol + 1):
        ws.column_dimensions[get_column_letter(c)].width = min(max(widths[c - 1], 9), 62)

    ws.freeze_panes = "A2"
    if ws.max_row > 1:
        ref = f"A1:{get_column_letter(ncol)}{ws.max_row}"
        # a real Excel table: banded rows, filter arrows and a name in the
        # Name Box, so the sheet behaves like something a person built
        safe = re.sub(r"\W", "", sheet.name) or "Table"
        t = Table(displayName=f"t{safe}"[:30], ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight1", showRowStripes=False)
        try:
            ws.add_table(t)
        except ValueError:
            ws.auto_filter.ref = ref
    ws.sheet_view.showGridLines = False
    return ws


def write_overview(wb, w, sheets, source_label):
    """The first thing you see: what this file is, when it was made, and what is
    in it — so a 14-sheet workbook does not open on a wall of data."""
    ws = wb.create_sheet("Overview", 0)
    ws.sheet_view.showGridLines = False
    ws["B2"] = "Arbor workspace"
    ws["B2"].font = TITLE_FONT
    now = dt.datetime.now()
    ws["B3"] = f"Exported {now.strftime('%d/%m/%Y at %H:%M')} · {source_label}"
    ws["B3"].font = SUB_FONT

    alive = [a for a in w.animals if a.get("alive")]
    in_use = [a for a in alive if not a.get("excluded") and not a.get("practice")]
    diary_n = sum(len(a.get("diary") or []) for a in w.animals)
    upcoming = 0
    today = dt.date.today()
    for t in w.schedule:
        d = as_date((t.get("when") or "")[:10])
        if isinstance(d, dt.date) and d >= today:
            upcoming += 1

    facts = [
        ("Animals", len(w.animals)),
        ("— alive", len(alive)),
        ("— alive and in use", len(in_use)),
        ("Animal records", diary_n),
        ("Projects", len(w.projects)),
        ("Experiments", sum(1 for e in w.experiments if not is_workstream(e))),
        ("Programs and courses", sum(1 for e in w.experiments if is_workstream(e))),
        ("Scheduled items", len(w.schedule)),
        ("— today or later", upcoming),
        ("Notes", len(w.notes)),
    ]
    r = 5
    ws.cell(row=r, column=2, value="At a glance").font = Font(
        name="Calibri", size=11, bold=True, color=INK)
    r += 1
    for label, value in facts:
        ws.cell(row=r, column=2, value=label).font = (
            SUB_FONT if label.startswith("—") else BODY_FONT)
        c = ws.cell(row=r, column=3, value=value)
        c.font = Font(name="Calibri", size=11, bold=not label.startswith("—"), color=INK)
        c.alignment = Alignment(horizontal="right")
        r += 1

    r += 1
    ws.cell(row=r, column=2, value="Sheets in this workbook").font = Font(
        name="Calibri", size=11, bold=True, color=INK)
    r += 1
    for s in sheets:
        link = ws.cell(row=r, column=2, value=s.name)
        link.font = Font(name="Calibri", size=11, color="0B62A8", underline="single")
        link.hyperlink = f"#'{s.name}'!A1"
        c = ws.cell(row=r, column=3, value=len(s.rows))
        c.font = BODY_FONT
        c.alignment = Alignment(horizontal="right")
        ws.cell(row=r, column=4, value="rows").font = SUB_FONT
        r += 1

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 10
    ws.column_dimensions["D"].width = 10
    return ws


# ── the sheets ────────────────────────────────────────────────────────────────
def cell_of(v, sf):
    """A sub-field value as a spreadsheet cell."""
    if v is None:
        return ""
    t = (sf or {}).get("type")
    if t == "yesno":
        return "Yes" if v else "No"
    if t == "number":
        return v
    return v


def table_sheets(w):
    """One sheet per table field, in long format — a line per row, per entry.

    A table cannot live in a cell, and flattening it to "3 rows" throws away
    the thing it was built to keep. Long format is what a pivot table or a plot
    wants, and it is what makes "how did this row change" answerable in the
    spreadsheet as well as in the app."""
    out, used = [], {}

    def name_for(label):
        n = (label or "Table")[:28]
        used[n] = used.get(n, 0) + 1
        return f"{n} {used[n]}" if used[n] > 1 else n

    for t in (w.config.get("diaryTypes") or []):
        for f in (t.get("fields") or []):
            if f.get("type") != "table":
                continue
            cols = [c for c in (f.get("fields") or []) if c.get("key")]
            pers = [c for c in (f.get("per") or []) if c.get("key")]
            rows = []
            if f.get("anchor") == "animal":
                for a in w.animals:
                    defined = (a.get("tables") or {}).get(f["key"]) or []
                    if not defined:
                        continue
                    entries = [e for e in (a.get("diary") or []) if e.get("type") == t.get("key")]
                    for r in defined:
                        base = [a.get("name") or ""] + [cell_of(r.get(c["key"]), c) for c in cols]
                        measured = [e for e in entries
                                    if ((e.get("data") or {}).get(f["key"]) or {}).get(r.get("id"))]
                        if not measured:
                            rows.append(base + ["", ""] + ["" for _ in pers])
                            continue
                        for e in measured:
                            v = ((e.get("data") or {}).get(f["key"]) or {}).get(r.get("id")) or {}
                            rows.append(base + [e.get("date") or "", e.get("time") or ""]
                                        + [cell_of(v.get(c["key"]), c) for c in pers])
                if rows:
                    out.append(Sheet(name_for(f.get("label")),
                        ["Animal"] + [c.get("label") or c["key"] for c in cols]
                        + ["Date", "Time"] + [c.get("label") or c["key"] for c in pers],
                        rows, sort=[0, 1, 2]))
            else:
                for a in w.animals:
                    for e in (a.get("diary") or []):
                        if e.get("type") != t.get("key"):
                            continue
                        v = (e.get("data") or {}).get(f["key"])
                        for i, r in enumerate(v if isinstance(v, list) else []):
                            rows.append([a.get("name") or "", e.get("date") or "",
                                         e.get("time") or "", i + 1]
                                        + [cell_of(r.get(c["key"]), c) for c in cols])
                if rows:
                    out.append(Sheet(name_for(f.get("label")),
                        ["Animal", "Date", "Time", "Row"]
                        + [c.get("label") or c["key"] for c in cols], rows, sort=[0, 1, 2]))

    for sec in (w.config.get("animalSections") or []):
        for f in (sec.get("fields") or []):
            if f.get("type") != "table":
                continue
            cols = [c for c in (f.get("fields") or []) if c.get("key")]
            rows = []
            for a in w.animals:
                for r in ((a.get("tables") or {}).get(f["key"]) or []):
                    rows.append([a.get("name") or ""] + [cell_of(r.get(c["key"]), c) for c in cols])
            if rows:
                out.append(Sheet(name_for(f.get("label")),
                    ["Animal"] + [c.get("label") or c["key"] for c in cols], rows, sort=[0]))
    return out


def stage_name(pl, st):
    """A stage's own name, or what it is — stageName in the app."""
    if st.get("label"):
        return st["label"]
    t = st.get("type")
    if t == "training" and st.get("task"):
        n = 0
        for x in (pl or {}).get("stages") or []:
            if x.get("type") == "training" and x.get("task") == st["task"]:
                n += 1
            if x is st:
                break
        tg = st.get("target")
        return f"{st['task']} session {n}" + (f" (imaged{' · ' + tg.lower() if tg else ''})" if st.get("imaged") else "")
    if t == "microscopy" and st.get("target"):
        return f"{st['target']} imaging"
    if t == "procedure":
        return "Procedure"
    return (RECORD_KINDS.get(t) or ("stage",))[0]


def is_workstream(e):
    return (e or {}).get("kind") == "workstream"


def entity_status(x):
    s = (x or {}).get("status")
    return s if s in ("active", "paused", "done") else "active"


def dmy(iso):
    s = str(iso or "")[:10]
    return f"{s[8:10]}/{s[5:7]}/{s[:4]}" if len(s) == 10 else s


# the workstream kinds and their own fields — WS_KINDS in the app
WS_KINDS = {
    "tool": ("Program", [("codeAt", "Lives at", "text")]),
    "course": ("Course you take", [("semester", "Semester", "text"), ("grade", "Grade", "text")]),
    "teaching": ("Course you teach", [("room", "Room", "text")]),
}


def record_data(e):
    """A record as it reads: what was written in it over what its plan stage
    said it would be (e.preset) — recordData in the app."""
    d = dict(e.get("preset") or {})
    d.update(e.get("data") or {})
    return d


def kind_sheets(w):
    """One sheet per record kind — the same sheets the app's export makes."""
    def recs(kind):
        return [(a, e, record_data(e)) for a in w.animals
                for e in (a.get("diary") or []) if e.get("type") == kind]

    def planned(e):
        return "" if Workspace.entry_happened(e) else "planned"

    def num(v):
        return "" if v in (None, "") else v

    def fov_name(a, fid):
        return next((f.get("name") or "" for f in Workspace.fovs(a) if f.get("id") == fid), "")

    out = []
    rows = []
    for a, e, d in recs("weight"):
        g = d.get("grams")
        if g in (None, "") and ration_of(d) == "":
            continue
        b = Workspace.baseline_at(a, e.get("date"))
        pct = round(1000 * float(g) / b[0]) / 10 if (b and b[0] and g not in (None, "")) else ""
        rows.append([a.get("name") or "", e.get("date") or "", e.get("time") or "", num(g),
                     ration_of(d), pct, Workspace.weighing_context(a, e.get("date")), note_of(d)])
    out.append(Sheet("Weighings", ["Animal", "Date", "Time", "Weight (g)", "Ration (g)", "% of baseline",
                                   "Why", "Note"], rows, sort=[0, 1, 2]))

    rows = []
    for a in w.animals:
        for r in Workspace.restrictions(a):
            b = Workspace.baseline_at(a, r["start"])
            rows.append([a.get("name") or "", r.get("start") or "", r.get("end") or "",
                         r.get("target") or RESTRICTION_TARGET, b[0] if b else "", b[1] if b else "",
                         r.get("endReason") or ""])
    out.append(Sheet("Food restrictions", ["Animal", "Start", "End", "Target %", "Baseline (g)",
                                           "Baseline date", "Ended because"], rows, sort=[0, 1]))

    rows = []
    for a, e, d in recs("surgery"):
        # as surgeryOfInjection: the surgery it names, else the first one on or
        # after the day it was planned (the first at all, if never dated)
        surgs = sorted((x for x in (a.get("diary") or []) if x.get("type") == "surgery" and x.get("date")),
                       key=lambda x: str(x.get("date")))

        def surgery_of(i):
            hit = next((x for x in surgs if i.get("surgeryId") and x.get("id") == i.get("surgeryId")), None)
            return hit or next((x for x in surgs if str(x.get("date"))[:10] >= (i.get("plannedAt") or "")), None)
        injs = [i for i in Workspace.injections(a) if surgery_of(i) is e]
        rows.append([a.get("name") or "", e.get("date") or "", e.get("time") or "", planned(e),
                     num(d.get("duration")), d.get("operator") or "", d.get("complications") or "",
                     yes(d.get("headpost")), yes(d.get("window")),
                     "; ".join((i.get("virus") or "?") + (" → " + i["siteRegion"] if i.get("siteRegion") else "")
                               for i in injs), note_of(d)])
    out.append(Sheet("Surgeries", ["Animal", "Date", "Time", "Status", "Duration (min)", "Operator",
                                   "Complications", "Head-post", "Window", "Injections", "Note"],
                     rows, sort=[0, 1], wide=[9]))

    rows = [[a.get("name") or "", e.get("date") or "",
             "Yes" if d.get("given") is True else "No" if d.get("given") is False else "",
             d.get("condition") or "", note_of(d)] for a, e, d in recs("postop")]
    out.append(Sheet("Post-op checks", ["Animal", "Date", "Post-op given", "Condition", "Note"],
                     rows, sort=[0, 1]))

    rows = [[a.get("name") or "", e.get("date") or "", e.get("time") or "", planned(e), d.get("task") or "",
             num(d.get("session")), d.get("stage") or "", num(d.get("trials")), num(d.get("success")), num(d.get("length")),
             ", ".join(f"{k}={v}" for k, v in (d.get("settings") or {}).items()), yes(d.get("imaged")),
             d.get("dataAt") or "", note_of(d)] for a, e, d in recs("training")]
    out.append(Sheet("Behaviour sessions", ["Animal", "Date", "Time", "Status", "Task", "Session #", "Stage", "Trials",
                                            "Success %", "Length (min)", "Settings", "Imaged", "Data at", "Note"],
                     rows, sort=[0, 1, 2], wide=[13]))

    rows = []
    for a, e, d in recs("microscopy") + [x for x in recs("training") if x[2].get("imaged")]:
        beh = e.get("type") == "training"
        rows.append([a.get("name") or "", e.get("date") or "", e.get("time") or "", planned(e),
                     d.get("target") or "", "functional" if beh else d.get("purpose") or "", yes(beh),
                     ", ".join(n for n in (fov_name(a, f.get("id")) for f in (d.get("fovs") or [])) if n),
                     d.get("laser") or "", d.get("channels") or "", d.get("status") or "",
                     d.get("dataAt") or "", note_of(d)])
    out.append(Sheet("Imaging sessions", ["Animal", "Date", "Time", "Status", "Imaging", "Purpose", "During behaviour",
                                          "FOVs", "Laser", "Channels", "Analysis", "Data at", "Note"],
                     rows, sort=[0, 1, 2], wide=[12]))

    rows = []
    for a in w.animals:
        for f in Workspace.fovs(a):
            den = f.get("dendrites") or []
            n = sum(1 for e in (a.get("diary") or []) if Workspace.entry_happened(e)
                    and any(x.get("id") == f["id"] for x in ((e.get("data") or {}).get("fovs") or [])))
            rows.append([a.get("name") or "", f.get("name") or ""] + [num(f.get(k)) for k, _ in FOV_SETTINGS]
                        + [", ".join(x.get("name") or "" for x in den),
                           ", ".join(x.get("name") or "" for x in den if x.get("reconstructed")), n])
    out.append(Sheet("FOVs", ["Animal", "FOV"] + [l for _, l in FOV_SETTINGS]
                     + ["Dendrites", "Reconstructed", "Times imaged"], rows, sort=[0, 1]))

    rows = [[a.get("name") or "", e.get("date") or "", planned(e), d.get("operator") or "",
             d.get("pbs") or "", yes(d.get("notProcessed")), note_of(d)] for a, e, d in recs("perfusion")]
    out.append(Sheet("Perfusions", ["Animal", "Date", "Status", "Done by", "Moved to PBS", "Not processed",
                                    "Note"], rows, sort=[0, 1]))

    rows = []
    for a, e, d in recs("histology"):
        head = [a.get("name") or "", e.get("date") or "", d.get("operator") or "", d.get("plane") or "",
                num(d.get("thickness")), d.get("storage") or ""]
        slides = d.get("slides") or []
        for x in slides:
            aps = x.get("aps") or []
            rows.append(head + [x.get("n"), len(aps), ", ".join(str(v) for v in aps), note_of(d)])
        if not slides:
            rows.append(head + ["", "", "", note_of(d)])
    out.append(Sheet("Slicing", ["Animal", "Date", "Sliced by", "Plane", "Thickness (µm)", "Stored at",
                                 "Slide", "Slices", "APs", "Note"], rows, sort=[0, 1, 6]))

    rows = []
    for a, e, d in recs("histo"):
        head = [a.get("name") or "", e.get("date") or "", d.get("operator") or "", d.get("dataAt") or ""]
        finds = d.get("findings") or []
        for x in finds:
            rows.append(head + [x.get("region") or "", x.get("dir") or "", x.get("slides") or "",
                                x.get("note") or "", note_of(d)])
        if not finds:
            rows.append(head + ["", "", "", "", note_of(d)])
    out.append(Sheet("Histology", ["Animal", "Date", "Done by", "Scans at", "Region", "Finding",
                                   "On slides", "Finding note", "Note"], rows, sort=[0, 1]))
    return out


def animal_fields_sheet(w):
    """Whatever the lab added to the animal window, beside the built-in columns."""
    cols = [(sec, f) for sec in (w.config.get("animalSections") or [])
            for f in (sec.get("fields") or []) if f.get("type") != "table"]
    if not cols:
        return None
    rows = []
    for a in w.animals:
        r = [a.get("name") or ""] + [cell_of((a.get("custom") or {}).get(f["key"]), f) for _, f in cols]
        if any(v != "" for v in r[1:]):
            rows.append(r)
    if not rows:
        return None
    return Sheet("Animal fields",
                 ["Animal"] + [f"{(sec.get('label') or 'Section')} · {f.get('label') or f['key']}"
                               for sec, f in cols], rows, sort=[0])


def build_sheets(w):
    sheets = []

    animals = []
    for a in w.animals:
        surg = w.surgery_dates(a)
        last = surg[-1] if surg else None
        animals.append([
            a.get("name") or "", a.get("cage") or "", a.get("sex") or "",
            a.get("strain") or "", Workspace.status(a),
            yes(a.get("practice")),
            (a.get("excludedReason") or "Yes") if a.get("excluded") else "",
            a.get("dob") or "", Workspace.age_weeks(a), a.get("sacrificed") or "",
            ", ".join(d.strftime("%d/%m/%Y") for d in surg),
            (dt.date.today() - last).days if last else "",
            Workspace.latest_weight(a),
            (lambda b: b[0] if b else "")(Workspace.baseline_at(a, dt.date.today())),
            Workspace.pct_baseline(a),
            Workspace.viruses(a), Workspace.sites(a), Workspace.projections(a),
            yes((a.get("surgeryDetail") or {}).get("headpost")),
            yes((a.get("surgeryDetail") or {}).get("window")),
            w.where(a.get("id")),
            w.next_stage(a),
            plain(a.get("comments")),
        ])
    sheets.append(Sheet(
        "Animals",
        ["Name", "Cage", "Sex", "Strain", "Status", "Practice", "Excluded", "Born",
         "Age (weeks)", "Sacrificed", "Surgeries", "Days since last surgery",
         "Latest weight (g)", "Baseline (g)", "% of baseline", "Virus",
         "Injection sites", "Projection targets", "Head-post", "Cranial window",
         "Experiments (group)", "Next plan stage", "Comments"],
        animals, sort=[0], wide=[10, 20, 22]))

    diary = []
    for a in w.animals:
        for e in a.get("diary") or []:
            diary.append([
                a.get("name") or "", e.get("date") or "", e.get("time") or "",
                e.get("endTime") or "", w.diary_label(e.get("type")),
                w.diary_detail(e, a), yes(Workspace.entry_happened(e)),
                w.belongs_to(e),
                ms_to_dt(e.get("created")),
            ])
    sheets.append(Sheet(
        "All records",
        ["Animal", "Date", "Time", "End time", "Kind", "Details", "Done",
         "Part of", "Logged"],
        diary, sort=[0, 1, 2], wide=[5]))
    sheets.extend(kind_sheets(w))

    inj = []
    for a in w.animals:
        for i, x in enumerate((a.get("surgeryDetail") or {}).get("injections") or [], 1):
            depths = [{"dv": x.get("dv"), "volume": x.get("volume")}] + list(x.get("depths") or [])
            for di, d in enumerate(depths, 1):
                inj.append([
                    a.get("name") or "", i, di, x.get("virus") or "", x.get("type") or "",
                    x.get("siteRegion") or "", x.get("ap"), x.get("ml"),
                    d.get("dv"), d.get("volume"),
                    x.get("partnerRegion") or "", x.get("pAp"), x.get("pMl"), x.get("pDv"),
                    x.get("verified") or "",
                ])
    sheets.append(Sheet(
        "Injections",
        ["Animal", "Injection #", "Depth #", "Virus", "Virus type", "Site region",
         "AP", "ML", "DV", "Volume (nl)", "Target region", "Target AP", "Target ML",
         "Target DV", "Histology"],
        inj, sort=[0, 1, 2]))

    timeline = []
    for t in w.schedule:
        e = w.exp_by_id.get(t.get("experimentId"))
        p = w.proj_by_id.get(t.get("projectId")) or (w.project_of_exp(e) if e else None)
        timeline.append([
            t.get("when") or "", t.get("endDate") or "", t.get("duration") or "",
            "Event" if t.get("isEvent") else "Task", t.get("title") or "",
            repeat_label(t.get("repeat")),
            Workspace.proj_name(p) if p else "", Workspace.exp_title(e) if e else "",
            len(t.get("doneOccurrences") or []) or "",
            plain(t.get("notes")),
        ])
    sheets.append(Sheet(
        "Timeline",
        ["When", "Until", "Duration (min)", "Kind", "Title", "Repeats", "Project",
         "Experiment", "Occurrences done", "Notes"],
        timeline, sort=[0, 4], wide=[9]))

    exps_only = [e for e in w.experiments if not is_workstream(e)]
    wss = [e for e in w.experiments if is_workstream(e)]
    titles = lambda xs: ", ".join(sorted((Workspace.exp_title(x) if x.get("title") else "(unnamed)" for x in xs), key=natural_key))

    exps = []
    for e in exps_only:
        p = w.project_of_exp(e)
        exps.append([
            Workspace.proj_name(p) if p else "", e.get("title") or "", entity_status(e), e.get("summary") or "",
            e.get("question") or "", e.get("deadline") or "", e.get("endDate") or "",
            ", ".join(w.animals_of_exp(e)),
            plain(e.get("description")), plain(e.get("outcome")),
        ])
    sheets.append(Sheet(
        "Experiments",
        ["Project", "Title", "Status", "Where it stands", "Question", "Deadline", "End", "Animals",
         "Description", "Outcome"],
        exps, sort=[0, 1], wide=[4, 7, 8, 9]))

    rows = []
    for e in wss:
        p = w.project_of_exp(e)
        k = WS_KINDS.get(e.get("wsKind")) or WS_KINDS["tool"]
        kd = e.get("kindData") or {}
        rows.append([Workspace.proj_name(p) if p else "", e.get("title") or "", k[0], entity_status(e),
                     e.get("summary") or "", "; ".join(f"{lab}: {kd[key]}" for key, lab, typ in k[1] if kd.get(key))])
    sheets.append(Sheet("Programs and courses", ["Project", "Title", "Kind", "Status", "Where it stands", "Details"],
                        rows, sort=[0, 1], wide=[4, 5]))

    projs = []
    for p in [p for p in w.projects if p.get("kind") != "area"]:
        mine = [e for e in w.experiments if e.get("projectId") == p.get("id")]
        projs.append([p.get("name") or "", entity_status(p), p.get("summary") or "", p.get("question") or "",
                      titles([e for e in mine if not is_workstream(e)]),
                      titles([e for e in mine if is_workstream(e) and e.get("wsKind") not in ("course", "teaching")]),
                      plain(p.get("notes"))])
    sheets.append(Sheet("Projects", ["Name", "Status", "Where it stands", "Question", "Experiments",
                                     "Programs", "Notes"],
                        projs, sort=[0], wide=[3, 4, 5, 6]))

    # what is inside a program, a course and a teaching course — "Work items" in
    # the app; a date is its task's
    tasks_by_id = {t.get("id"): t for t in w.schedule}

    def tday(i):
        t = tasks_by_id.get(i) if i else None
        return str(t.get("when"))[:10] if t and t.get("when") else ""

    def tdone(i, own):
        t = tasks_by_id.get(i) if i else None
        return (t.get("when") in (t.get("doneOccurrences") or [])) if t and t.get("when") else (False if t else bool(own))
    work = []
    for e in wss:
        of, k = Workspace.exp_title(e), (WS_KINDS.get(e.get("wsKind")) or WS_KINDS["tool"])[0]
        for x in e.get("ideas") or []:
            work.append([of, k, "idea", x.get("text") or "", "", "", "", ""])
        for x in e.get("deliverables") or []:
            work.append([of, k, "deliverable", x.get("text") or "", x.get("dkind") or "", "", tday(x.get("taskId")),
                         yes(tdone(x.get("taskId"), x.get("done")))])
        for x in e.get("uploads") or []:
            work.append([of, k, "upload", ("נושא " + x["topic"]) if x.get("topic") else "", x.get("what") or "",
                         tday(x.get("prepTaskId")), tday(x.get("dueTaskId")), yes(tdone(x.get("dueTaskId"), x.get("uploaded")))])
        for key, v in (e.get("missed") or {}).items():
            tid, _, when = key.partition("|")
            t = tasks_by_id.get(tid)
            work.append([of, k, "missed lecture", (t.get("title") or "") if t else "",
                         "caught up" if v == "caught" else "to catch up", "", when[:10], yes(v == "caught")])
    sheets.append(Sheet("Work items", ["Of", "Kind", "Item kind", "Item", "Detail", "Prepare from", "Due", "Done"],
                        work, sort=[0, 2, 6]))
    sheets.append(Sheet("Weekend feeding", ["Weekend (Friday)", "Who"],
                        [[f, who] for f, who in ((w.config.get("weekendFeeding") or {}).items())], sort=[0]))

    # animal orders, mirroring orderList()/orderState(): the request week and the
    # arrival are the linked tasks' dates, falling back to the order week
    def add_days(iso, n):
        return (dt.date.fromisoformat(iso[:10]) + dt.timedelta(days=n)).isoformat()
    orders = []
    for o in (w.db.get("animalOrders") or []):
        if not o or not o.get("id"):
            continue
        st, sd, ad = tasks_by_id.get(o.get("sendTask")), o.get("week") or "", add_days(o["week"], 7) if o.get("week") else ""
        send = str(st["when"])[:10] if st and st.get("when") else sd
        at = tasks_by_id.get(o.get("arriveTask"))
        arrive = str(at["when"])[:10] if at and at.get("when") else ad
        sent = bool(o.get("received")) or bool(st and st.get("when") in (st.get("doneOccurrences") or []))
        f = w.exp_by_id.get(o.get("forId")) or w.proj_by_id.get(o.get("forId"))
        label = (Workspace.exp_title(f) if "title" in f else Workspace.proj_name(f)) if f else ""
        g = next((x for x in (f or {}).get("groupDefs") or [] if x.get("id") == o.get("groupId")), None) if f else None
        age = o.get("age")
        orders.append([send, arrive, label, (g.get("name") or "") if g else "", int(o.get("count") or 0),
                       {"M": "male", "F": "female"}.get(o.get("sex"), "either"), o.get("strain") or "",
                       "" if age in ("", None) else age, o.get("names") or "",
                       "arrived" if o.get("received") else "on the way" if sent else "to send",
                       o.get("received") or "", o.get("note") or ""])
    sheets.append(Sheet("Animal orders", ["Request goes in", "Arrives", "For", "Group", "How many", "Sex", "Strain",
                                          "Age (weeks)", "Names", "Status", "Received", "Note"], orders, sort=[0]))

    # every checklist; a dated item is a task, so its title and date are the task's
    tasks = {t.get("id"): t for t in w.schedule}
    checks = []
    for kind, x in [("Project", p) for p in w.projects] + \
            [((WS_KINDS.get(e.get("wsKind")) or WS_KINDS["tool"])[0] if is_workstream(e) else "Experiment", e)
             for e in w.experiments]:
        for it in x.get("checklist") or []:
            t = tasks.get(it.get("taskId")) if it.get("taskId") else None
            done = (t.get("when") in (t.get("doneOccurrences") or [])) if t and t.get("when") else bool(it.get("done")) if not t else False
            checks.append([kind, Workspace.proj_name(x) if kind == "Project" else Workspace.exp_title(x),
                           (t.get("title") or "") if t else (it.get("text") or ""), yes(done),
                           str(t.get("when"))[:10] if t and t.get("when") else ""])
    sheets.append(Sheet("Checklists", ["Kind", "Of", "Item", "Done", "Date"], checks, sort=[0, 1]))

    notes = []
    for n in w.notes:
        notes.append([n.get("title") or "", w.attached_to("note", n.get("id")),
                      ms_to_dt(n.get("created")), ms_to_dt(n.get("updated")),
                      plain(n.get("body"))])
    sheets.append(Sheet("Notes", ["Title", "Attached to", "Created", "Updated", "Text"],
                        notes, sort=[0], wide=[4]))

    lib = []
    for kind, key in (("Paper", "readings"), ("Protocol", "protocols"), ("Other", "other")):
        for r in w.db.get(key) or []:
            lib.append([
                kind, r.get("title") or "", r.get("authors") or r.get("writtenBy") or "",
                r.get("year") or r.get("date") or "", r.get("journal") or "",
                w.attached_to({"readings": "reading", "protocols": "protocol",
                               "other": "other"}[key], r.get("id")),
                r.get("link") or "", r.get("fileName") or "", plain(r.get("notes")),
            ])
    sheets.append(Sheet(
        "Library",
        ["Kind", "Title", "Authors / written by", "Year / date", "Journal",
         "Attached to", "Link", "File", "Notes"],
        lib, sort=[0, 1], wide=[1, 8]))

    viruses = [[v.get("name") or "", v.get("type") or "", v.get("promoter") or "",
                v.get("titer") or "", v.get("source") or "", plain(v.get("note")),
                v.get("backupTubes", ""), v.get("aliquotTubes", ""),
                v.get("boxNumber") or "", v.get("boxLocation") or "",
                yes(v.get("needOrder"))]
               for v in (w.config.get("virusLib") or [])]
    sheets.append(Sheet(
        "Virus library",
        ["Name", "Type", "Promoter", "Titer", "Source", "Note", "Backup tubes",
         "Aliquoted tubes", "Box #", "Box location", "Need to order"],
        viruses, sort=[0], wide=[0, 5]))

    sites = [[s.get("name") or "", s.get("ap"), s.get("ml"), s.get("dv")]
             for s in (w.config.get("siteLib") or [])]
    sheets.append(Sheet("Injection sites", ["Name", "AP", "ML", "DV"], sites, sort=[0]))

    # Lever's settings are designed into Arbor (BEHAVIOUR_TASKS in the app);
    # any other task carries its own short list in the task library
    lib = {t.get("name"): t for t in (w.config.get("taskLib") or []) if t.get("name")}
    used = {e.get("data", {}).get("task") for a in w.animals for e in (a.get("diary") or [])
            if e.get("type") == "training" and (e.get("data") or {}).get("task")}
    names = sorted(set(BEHAVIOUR_TASKS) | set(lib) | used, key=natural_key)
    tasks = []
    for n in names:
        sets = list(BEHAVIOUR_TASKS.get(n) or []) + [p.get("label") or p.get("key") or ""
                                                     for p in (lib.get(n, {}).get("params") or [])
                                                     if p.get("key") and (p.get("label") or "").strip() and not re.match(r"(?i)^(time|trials|success)", p.get("label") or "")]
        sessions = sum(1 for a in w.animals for e in (a.get("diary") or [])
                       if e.get("type") == "training" and (e.get("data") or {}).get("task") == n)
        tasks.append([n, ", ".join(sets), sessions])
    sheets.append(Sheet("Behavioural tasks", ["Task", "Settings", "Sessions"], tasks,
                        sort=[0], wide=[1]))

    inv = [[k, (v or {}).get("status") or "", plain((v or {}).get("note"))]
           for k, v in (w.db.get("inventory") or {}).items()]
    sheets.append(Sheet("Stock", ["Material", "Status", "Note"], inv, sort=[0], wide=[2]))

    # the legacy outline, last: superseded by Timeline, but exporting it keeps
    # the workbook a complete record of what the workspace still holds
    doms = {d["id"]: d for d in (w.db.get("domains") or []) if d.get("id")}
    outline = [[t.get("title") or "", (doms.get(t.get("domainId")) or {}).get("name") or "",
                t.get("status") or "", t.get("deadline") or "", t.get("endDate") or "",
                plain(t.get("notes"))]
               for t in (w.db.get("tasks") or [])]
    sheets.append(Sheet("Outline tasks",
                        ["Name", "Soma", "Status", "Start", "End", "Notes"],
                        outline, sort=[0], wide=[5]))

    # whatever this lab added for itself, after the sheets that always exist
    extra = animal_fields_sheet(w)
    if extra:
        sheets.append(extra)
    sheets.extend(table_sheets(w))

    return [s for s in sheets if s.rows]


def build_workbook(db, source_label):
    w = Workspace(db)
    sheets = build_sheets(w)
    wb = Workbook()
    wb.remove(wb.active)
    for s in sheets:
        write_sheet(wb, s, colour_status_col=(5 if s.name == "Animals" else None))
    write_overview(wb, w, sheets, source_label)
    wb.active = 0
    return wb, w, sheets


# ── input ─────────────────────────────────────────────────────────────────────
def load_from_gist(gist_id, gh="gh"):
    out = subprocess.run(
        [gh, "api", f"/gists/{gist_id}", "--jq", '.files["arbor_workspace.json"].content'],
        capture_output=True, text=True, timeout=120)
    if out.returncode != 0:
        raise SystemExit(f"Could not read the Gist: {out.stderr.strip()[:300]}")
    text = out.stdout
    if not text.strip():
        raise SystemExit("The Gist returned no workspace file.")
    return json.loads(text), text


def main():
    ap = argparse.ArgumentParser(description="Export an Arbor workspace to Excel.")
    src = ap.add_mutually_exclusive_group(required=True)
    src.add_argument("--gist", help="Gist id holding arbor_workspace.json")
    src.add_argument("--json", help="A workspace JSON file")
    ap.add_argument("--out", help="Folder to write a dated workbook into")
    ap.add_argument("--xlsx", help="Write to exactly this path instead")
    ap.add_argument("--keep", type=int, default=30,
                    help="How many dated exports to keep in --out (default 30)")
    ap.add_argument("--gh", default=os.path.expanduser("~/.local/bin/gh"),
                    help="Path to the gh CLI")
    args = ap.parse_args()

    if args.gist:
        db, raw = load_from_gist(args.gist, args.gh)
        source = "from the synced workspace"
    else:
        raw = open(args.json, encoding="utf-8").read()
        db = json.loads(raw)
        source = f"from {os.path.basename(args.json)}"

    counts = {k: len(db.get(k) or []) for k in ("animals", "schedule", "experiments")}
    if not db.get("animals") and not db.get("schedule"):
        raise SystemExit(f"Refusing to export an empty-looking workspace: {counts}")

    wb, w, sheets = build_workbook(db, source)

    if args.xlsx:
        path = args.xlsx
        os.makedirs(os.path.dirname(os.path.abspath(path)) or ".", exist_ok=True)
    else:
        if not args.out:
            raise SystemExit("Give either --out or --xlsx.")
        os.makedirs(args.out, exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M")
        path = os.path.join(args.out, f"Arbor {stamp}.xlsx")
        # the JSON beside it: this folder is also the backup that depends on
        # nothing — not a browser, not GitHub, not this script
        with open(os.path.join(args.out, f"Arbor {stamp}.json"), "w", encoding="utf-8") as f:
            f.write(raw)

    wb.save(path)

    if args.out and args.keep > 0:
        keep_recent(args.out, args.keep)

    rows = sum(len(s.rows) for s in sheets)
    print(f"{path}  ·  {len(sheets)} sheets, {rows} rows  ·  "
          f"{len(w.animals)} animals, {sum(len(a.get('diary') or []) for a in w.animals)} animal records")


def keep_recent(folder, keep):
    """Dated exports pile up twice a day. Keep the most recent `keep` of each
    kind and delete the rest, so the folder stays a usable list."""
    for ext in (".xlsx", ".json"):
        files = sorted((f for f in os.listdir(folder)
                        if f.startswith("Arbor ") and f.endswith(ext)), reverse=True)
        for old in files[keep:]:
            try:
                os.remove(os.path.join(folder, old))
            except OSError:
                pass


if __name__ == "__main__":
    sys.exit(main())
