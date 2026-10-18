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


DOW = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]
PLAIN_REPEATS = {"day": "daily", "week": "weekly", "2week": "every 2 weeks",
                 "month": "monthly", "year": "yearly"}


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
        return (e or {}).get("title") or "(unnamed experiment)"

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
        w = a.get("weights") or {}
        if isinstance(w, dict) and w:
            try:
                return w[max(w.keys())]
            except (ValueError, TypeError):
                return ""
        return ""

    @staticmethod
    def age_weeks(a):
        dob = as_date(a.get("dob"))
        if not isinstance(dob, dt.date):
            return ""
        end = as_date(a.get("sacrificed"))
        end = end if isinstance(end, dt.date) else dt.date.today()
        return round((end - dob).days / 7)

    def surgery_dates(self, a):
        """Every surgery on the animal: the profile's own date plus any logged
        as diary entries, deduplicated."""
        out = set()
        if a.get("surgery"):
            d = as_date(a["surgery"])
            if isinstance(d, dt.date):
                out.add(d)
        for e in a.get("diary") or []:
            if e.get("type") == "surgery":
                d = as_date(e.get("date"))
                if isinstance(d, dt.date):
                    out.add(d)
        return sorted(out)

    def diary_label(self, key):
        t = self.diary_types.get(key)
        return (t or {}).get("label") or (key or "").title()

    def diary_detail(self, e):
        """The entry's own fields, as 'Label: value' pairs — the same fields the
        app shows under an entry."""
        data = e.get("data") or {}
        t = self.diary_types.get(e.get("type")) or {}
        bits = []
        for f in t.get("fields") or []:
            v = data.get(f.get("key"))
            if v in (None, "", []):
                continue
            if f.get("type") == "yesno":
                v = "Yes" if v else "No"
            bits.append(f"{f.get('label') or f.get('key')}: {v}")
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
        ("Diary entries", diary_n),
        ("Projects", len(w.projects)),
        ("Experiments", len(w.experiments)),
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
            yes((a.get("surgeryDetail") or {}).get("headpost")),
            yes((a.get("surgeryDetail") or {}).get("window")),
            w.where(a.get("id")),
            plain(a.get("comments")),
        ])
    sheets.append(Sheet(
        "Animals",
        ["Name", "Cage", "Sex", "Strain", "Status", "Practice", "Excluded", "Born",
         "Age (weeks)", "Sacrificed", "Surgeries", "Days since last surgery",
         "Latest weight (g)", "Head-post", "Cranial window", "Experiments (group)",
         "Comments"],
        animals, sort=[0], wide=[10, 15, 16]))

    diary = []
    for a in w.animals:
        for e in a.get("diary") or []:
            diary.append([
                a.get("name") or "", e.get("date") or "", e.get("time") or "",
                e.get("endTime") or "", w.diary_label(e.get("type")),
                w.diary_detail(e), yes(e.get("done")),
                ms_to_dt(e.get("created")),
            ])
    sheets.append(Sheet(
        "Diary entries",
        ["Animal", "Date", "Time", "End time", "Type", "Details", "Done", "Logged"],
        diary, sort=[0, 1, 2], wide=[5]))

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
                ])
    sheets.append(Sheet(
        "Injections",
        ["Animal", "Injection #", "Depth #", "Virus", "Virus type", "Site region",
         "AP", "ML", "DV", "Volume (nl)", "Target region", "Target AP", "Target ML",
         "Target DV"],
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

    exps = []
    for e in w.experiments:
        p = w.project_of_exp(e)
        exps.append([
            Workspace.proj_name(p) if p else "", Workspace.exp_title(e),
            e.get("question") or "", e.get("deadline") or "", e.get("endDate") or "",
            ", ".join(w.animals_of_exp(e)),
            plain(e.get("description")), plain(e.get("outcome")),
        ])
    sheets.append(Sheet(
        "Experiments",
        ["Project", "Title", "Question", "Deadline", "End", "Animals", "Description",
         "Outcome"],
        exps, sort=[0, 1], wide=[2, 5, 6, 7]))

    projs = []
    for p in w.projects:
        kids = [Workspace.exp_title(e) for e in w.experiments
                if e.get("projectId") == p.get("id")]
        kids.sort(key=natural_key)
        projs.append([p.get("name") or "", p.get("question") or "",
                      ", ".join(kids), plain(p.get("notes"))])
    sheets.append(Sheet("Projects", ["Name", "Question", "Experiments", "Notes"],
                        projs, sort=[0], wide=[1, 2, 3]))

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

    tasks = [[t.get("name") or "",
              ", ".join(p.get("label") or p.get("key") or "" for p in (t.get("params") or []))]
             for t in (w.config.get("taskLib") or [])]
    sheets.append(Sheet("Behavioural tasks", ["Task", "Parameters"], tasks,
                        sort=[0], wide=[1]))

    inv = [[k, (v or {}).get("status") or "", plain((v or {}).get("note"))]
           for k, v in (w.db.get("inventory") or {}).items()]
    sheets.append(Sheet("Stock", ["Material", "Status", "Note"], inv, sort=[0], wide=[2]))

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
          f"{len(w.animals)} animals, {sum(len(a.get('diary') or []) for a in w.animals)} diary entries")


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
