# Arbor 2.0 — screen-by-screen audit

The record kinds are rebuilt (see `arbor-2.0-design.md`). The screens around
them are still 1.0: they were written for "any entry type with any fields",
and many still speak that language. This lists every tab and window, what in
it is left over from 1.0, and what it should become. Items marked **Decide**
need the user; the rest follow from decisions already made.

Status key: ✅ fine as is · 🔧 rework · 🗑 retire · ✨ new

## Decisions (02/11/2026)

- **Restriction board** — a view inside the Animals tab: *Show as: Restriction*.
- **🔖 Fields (own sections/fields/tables on the animal)** — retired. Values
  already recorded stay visible read-only until each is designed into the
  animal model.
- **Filters** — a fixed set of questions replaces the condition builder; the
  list below is shown to the user before it is built.
- **Plans library in Materials** — removed; plans live in their experiment.

## Progress — full pass done (02/11/2026)

Built, then checked on screen and on a fresh migration of the real data:
Animals tab (restriction board, fixed filter questions, 2.0 sorts, cards,
own fields retired, the Log with notes and kinds) · animal window (Diary
month list, form-only records, done only for scheduled records, starting
values, Save only on Details) · dates dd/mm/yyyy and 24h everywhere ·
"record" as the one word · Find searches record content · Experiments
wording · Materials: plans library gone, Behaviour tasks sub-tab · a day's
feeding can move · phone layout (feeding cards, filter chips, stray sheet)
· both exports identical row for row. Left switched off rather than cut:
the old Arbor graph/outline views (unreachable; data kept and exported).

---

## The Animals tab

| Piece | Now | 2.0 |
|---|---|---|
| Filter row | Cage, experiment, project, group, status, excluded, practice, and **"Entry: any entry type"** | 🔧 Keep cage / experiment / project / group / status / excluded / practice. Replace "Entry" with questions the lab actually asks — see *Filters* below |
| "⊕ Add condition" | A generic query builder: *any entry type → any field → over / under / contains*, plus task metrics that no longer exist | 🗑 Retire the builder. Replace with a fixed set of questions (below) |
| Show as | List / Cages | 🔧 Add **Restriction** — see the restriction board below |
| Sort by | Name, born, added, sacrificed | 🔧 Add % of baseline, restriction day, stage, last session, days post-op |
| Summary line chooser | Facts, including "Your own fields" | ✅ Keep; it already reads the 2.0 facts. Drops "own fields" if those retire (**Decide**, below) |
| 🔖 Fields | Opens "Fields on the animal": make your own sections and fields on the animal | **Decide** — this is the same open-ended flexibility 2.0 removed from records. Keep, limit, or retire? |
| 📋 Event log | Every record, date by date; still mentions repeating entries | 🔧 Reword; group by kind; read each kind's own summary |
| Animal card | Name, cage, dob chip, "sacrificed?" chip, surgery chip | 🔧 Show what 2.0 knows: restriction %, stage, warnings; "sacrificed?" only where it makes sense |

### Filters — proposed fixed questions

Instead of a builder, a short list of switches, each one a real question:

- **Food restriction:** on restriction · reaching / maintaining · below 85% · not fed today
- **Procedure:** had surgery · surgery booked · post-op period · virus is … · site is …
- **Behaviour:** task is … · stage is … · no session in N days
- **Imaging:** has FOVs · imaged in last N days · reconstruction pending
- **Brain:** perfused · in PBS · sliced · histology done · not processed
- **Warnings:** has any warning

Every one of these reads a fact the animal's window already shows, so the
filter and the window can never disagree.

**Decide:** is this the right list? Anything you filter by today that isn't here?

### ✨ Restriction board

One place for every animal on restriction, as a table: animal · day · baseline
· last weight · % · phase (reaching / maintaining) · ration · trend (small
weight curve) · fed today ✓ · warnings. Weight and ration typed straight into
the row, the same as the 16:00 feeding window. Ending a restriction or starting
a new one from here.

**Decide:** a view inside the Animals tab ("Show as: Restriction"), or its own tab?

---

## An animal's window

| Piece | Now | 2.0 |
|---|---|---|
| Overview | 2.0 sections (body, restriction, procedure, behaviour, imaging, brain, notes) | ✅ Mostly done. 🔧 "Recorded on the animal" (own fields) depends on the Fields decision |
| Diary tab | Month calendar + list + form; filter chips by type | 🔧 Filter chips become kinds; the calendar marks kinds by icon, not by user colour |
| Details tab | Profile, procedure editor, own sections | 🔧 Own sections depend on the Fields decision; review every profile field against the animal model |

---

## The Timeline

| Piece | Now | 2.0 |
|---|---|---|
| Tasks / Events / Animal work filters | ✅ | Keep |
| Feeding, post-op, brain steps | ✅ Drawn | Keep. **Decide:** let one day's feeding be moved (design said any day is movable; not built) |
| Clicking an empty slot | Task / Event / Record | ✅ Record now hands over to the animal's form |
| Colours | Entry types had user colours | 🔧 Kinds have fixed colours; check the scope ladder (project → experiment → record) still reads |
| Month view | Agenda chips | ✅ |
| .ics export | Counts "diary entries" | 🔧 Reword, export kinds by name |

---

## Experiments tab and windows

| Piece | Now | 2.0 |
|---|---|---|
| Experiment window | Arms, plans per arm, animals, outcome | 🔧 Check every stage type is a kind; the per-animal summary should read kind summaries (stage, success, % baseline) |
| Plan editor | Stage types come from the kinds | ✅ Check wording |
| "Subsumes diary entries" (records under an experiment) | Shows each record as "animal — type" | 🔧 Show the kind's own summary and icon |
| Project window | Experiments, links, sharing | ✅ Review wording only |

---

## Materials tab

| Sub-tab | Now | 2.0 |
|---|---|---|
| Injection sites | ✅ | Keep |
| Viruses | ✅ | Keep |
| Procedures | Surgery templates | ✅ Keep; check they feed the surgery kind |
| Plans | Library of experiment plans | 🔧 Plans are now made inside an experiment; **Decide:** keep a library here at all? |
| Inventory | Materials and stock | ✅ Keep |
| Behavioural tasks (in task library) | Each task has generic "params" | 🔧 Becomes *task settings* (what Lever's threshold/frequency/angle are) — the one flexibility kept |

---

## Library, Notes

Not record-shaped; ✅ keep. Check that notes attached to an animal (the
migrated comments) show in the animal's window and in the event log.

---

## Menu, search, bell, sync

| Piece | 2.0 |
|---|---|
| Bell | ✅ Reads the warnings layer |
| Find (search) | 🔧 Search records by kind summary, not by generic field text |
| Export Excel | ✅ One sheet per kind |
| Reset to seed | 🔧 The seed (demo workspace) is 1.0 — rebuild it in 2.0 shape or remove the option |
| Sharing, history, backup | ✅ |

---

## Retire outright

- 🗑 The **Arbor graph** (somata and dendrites, the "outline" of `DB.tasks` /
  `DB.domains`): no tab reaches it any more; it still renders and exports an
  "Outline tasks" sheet. Remove, keeping the data in backups.
- 🗑 The generic condition builder and its operators.
- 🗑 The generic table field editor (`renderTableField`, per-row values),
  unless the Fields decision keeps tables on the animal.
- 🗑 Leftover colour entries for entry types in the colour picker's
  "used by" list.

---

## Order of work (proposed)

1. Decisions above, with the user.
2. Animals tab: filters, restriction board, sort, cards.
3. Animal window: diary tab, details, own fields per the decision.
4. Timeline wording/colours; movable feeding day if wanted.
5. Experiments: per-animal summaries, records shown by kind.
6. Materials: behavioural task settings; plans library per the decision.
7. Search, event log, seed, menu.
8. Retirements; full regression and two-device test again.
