# Arbor 2.0 — design record

Decisions agreed with Yaron, type by type. Nothing here is built yet; the types
are designed first and then built together.

## Principles

1. **Arbor stores what you did and what is true.** Everything you look at — the
   timeline, overviews, tables, warnings — is computed from those, never a
   second copy.
2. **One source per fact.** A weight lives in a weighing; everything else reads it.
3. **Every important thing has a hub** (overview, history, links): animal,
   experiment, project.
4. **Links run both ways.** A note about an animal shows in the animal's hub.
5. **Records are specific.** No generic diary entry, no generic task. A new
   recording need gets designed, not configured.
6. **Warnings come from facts, in one place.**
7. **The timeline is derived**, except plain events (meetings, deadlines).
8. **Built incrementally inside the current app.** Each 2.0 piece replaces a 1.0
   piece, migrates its data, then retires it. The sync/sharing layer carries over.

## Pieces

| piece | what it is |
|---|---|
| Project | research question; sharing boundary |
| Experiment | arms, each arm's protocol, cohort, hypothesis, outcome |
| Animal | profile, current state, full history — the centre |
| Records | purpose-built, one kind per real activity |
| Protocols | the records an arm calls for, with timing (1.0 "plans") |
| Timeline | derived: what's due + events |
| Events | the one generic calendar item |
| Notes & Library | documents linked to anything, shown in its hub |
| Materials | viruses, sites, procedure templates |

## Types

### Food restriction — designed (02/11/2026)

A **period** on the animal, not an entry.

- **Start:** date; baseline = one weighing taken at start (becomes the animal's
  baseline); target **90%** of baseline.
- **Phases:** *reaching target* → (at or below 90%) **ready for behaviour** →
  *maintaining* (restriction continues through behaviour trials; show drift
  from 90%).
- **Safety floor 85%:** the loudest warning in the app — animal, timeline,
  experiment table, notification.
- **Daily:** one line per animal = weight + ration given. The weight is a real
  weighing (single source). Ration defaults to yesterday's (typically
  2.65–2.8 g).
- **Timeline:** the restriction itself produces one item per day, default
  **16:00** (any day movable): "🍽 Feeding · N on restriction · k/N done", with
  warnings. Opens the daily form for all; replaces the האכלות recurring task.
  Roster = whoever is on restriction (no manual roster). Done = every animal
  weighed and fed.
- **End:** end date + reason; past periods kept as history.
- **Shows in:** animal hub (section), animal card, experiment cohort table,
  filters (on restriction / ready / % of baseline / days on restriction),
  notifications, export (periods sheet + daily log sheet).
- **Migration:** existing food-restriction entries and the האכלות task become
  restriction periods, checked against real data before retiring them.

### Weighing — designed (02/11/2026)

- Its own small record: grams + date/time.
- **Context is derived, never tagged:** part of a restriction day; *pre-op* /
  *post-op* when within a few days of a surgery; otherwise *routine*.
- **Two references:** restriction → restriction baseline; recovery → the pre-op
  weight. Both warn below **85%** (assumed same floor; confirm).
- No dose calculator — surgery doses are roughly constant; the surgery record
  notes what was given.

### Behaviour — designed (02/11/2026)

- **Session summaries only** (no per-trial data in Arbor for now).
- **Shared skeleton (fixed):** date/time, task, stage (beginner / middle /
  expert), trials, success rate (**read off the rig**, entered), session length,
  notes, **data location**.
- **Lever settings, initially:** threshold, sound frequency, **table angle**.
  Table angle can differ between animals but is kept the same within an
  animal — warn if a session's angle differs from that animal's previous ones.
- **Stage is set by hand.** Criteria vary (sometimes a session count,
  sometimes a success rate that may be coincidental), so Arbor doesn't decide —
  it shows the evidence beside the stage control: sessions at this stage and
  the last few success rates.
- **No imaging settings on the session** for now; revisit under Imaging.
- **Per task:** its own settings and measures. Lever: threshold, sound
  frequency, … Unknown tasks start as a short list of named settings/measures
  and can later get a purpose-built design.
- **Settings carry session to session** per animal until changed (the one place
  "stays until changed" belongs).
- **Views:** learning curve per animal (success rate per session, stage changes
  marked); learning curves by arm in the experiment; date each stage reached.
- **Restriction:** session form shows current % of baseline and readiness.
  Rewards do **not** count toward the daily ration.
- **Imaged sessions:** one session carrying imaging; two files per trial
  (imaging + video); Arbor stores *where the data is*, not the files.

### Surgery & Post-op — designed (02/11/2026)

- **Surgery record:** date/time, **duration**, **operator** (a lab member),
  **complications**, notes, and the procedure parts *it* performed. Drugs are
  not recorded — standard protocol.
- **Procedure parts belong to the surgery that did them.** Injection in
  surgery 1, window in surgery 2: each surgery holds its own parts; the animal's
  procedure is the union. (1.0 kept one `surgeryDetail` per animal with
  `inj.surgeryId` as a patch; 2.0 makes it the model.)
- **Booking:** a surgery can be booked ahead (planned or not — a second surgery
  is often unplanned) from the animal's hub; booked ones are on the timeline.
- **Post-op, 72 h:** once a surgery has happened, Arbor puts **one check per
  day for 3 days** on the timeline (grouped across animals, like feeding). Each
  check = *post-op given* ✓ + *condition* (+ note). Done = both recorded.
  Weighings in this window are post-op weighings and warn below 85% of the
  pre-op weight.
- **Waiting periods → "ready from" dates**, shown in the hub, the cohort table
  and when starting the next step:
  - virus expression ≈ **1 month** after an injection surgery;
  - restriction / behaviour from **3–4 weeks** after surgery (starting a
    restriction earlier asks first).
  Lab-wide defaults, adjustable.
- **Open (minor):** condition scale — good / watch / poor + note assumed.

### Imaging (Microscopy, Spines, Reconstruction) — proposed (02/11/2026)

Yaron is least sure here, so the design tracks rather than models; expand later.

- **One record: imaging session**, with a purpose — **functional** (during
  behaviour → attached to that behaviour session, one event) or **structural**
  (stands alone, e.g. "Baseline Spines"). Not every animal gets both.
- **Fields of view belong to the animal** (chronic imaging returns to them):
  name, depth, zoom, number of planes. A session lists which FOVs it imaged;
  per-FOV settings **carry from the last session** that imaged that FOV.
- **Session-wide:** laser power, channels, data location.
- **Reconstruction = analysis of a session's data**, not a session: each FOV
  holds its dendrites; each dendrite records reconstructed ✓, when, by whom,
  file location.
- **Analysis status (optional):** raw → processed → analysed per session; the
  hub lists what's waiting. Drop if unused.
- **Not modelled yet:** spine counts / analysis results (track that analysis
  happened and where results are).

### Tissue (Perfusion, Slicing, Histology) — designed (02/11/2026)

- **Perfusion:** date, who, notes (for anything out of the ordinary). Ends the
  animal's life (already handled by the end-of-life role).
- **The brain as a specimen**, with its own derived steps on the timeline:
  **PFA → PBS** 2 days after perfusion (a check, like post-op), then **ready
  to slice** a day or two later. Hub shows the brain's state (in PFA → in PBS →
  sliced). Not every brain is processed: can be marked *not processed*, which
  stops the prompts.
- **Slicing:** date, who, plane (coronal default; sagittal possible),
  thickness, storage location; **slides and slices modelled** — each slice
  unique, each slide holds slices of descending AP coordinates (similar areas
  per slide); a slide shows its AP range.
- **Histology:** (a) **verification** — each injection marked confirmed /
  off-target / not expressed, on the animal; (b) **exploratory** — findings
  (regions projecting to/from, with notes) pointing to the slides/slices where
  seen; data location of scans.
- **Slide layout is not fixed** (varies per person and per slide): a slide
  lists its slices, uppermost = most anterior, typically 5–8 per slide, each
  slice with its own AP. Arbor suggests the next AP from the previous slice and
  the thickness; it never imposes a pattern.

### Comment — designed

- Becomes a **dated note on the animal** (Notes system, shown in the animal's
  history) rather than an entry type. Not throwaway — just not an "activity".

## What the real workspace contains (read-only copy, 02/11/2026)

36 animals (10 alive), 4 projects, 4 experiments, 1 plan ("Two Tasks", 36
stages), 411 scheduled items, 37 notes.

| 1.0 type (key → label) | rows | with content | notes |
|---|---|---|---|
| weight → "Food restriction" (+ field *Amount (g)*) | 266 | 79 | 202 in repeating series; 119 empty future placeholders (T2I-1 runs to 04/2027) |
| training → "Behavior" | 37 | 4 | 30 from the plan, 32 empty future; tasks *Lever* (Time, Trials, Success %), *Droplet*; note "זוית 5" = table angle |
| surgery | 24 | 12 | 11 empty past rows |
| postop | 14 | 12 | |
| perfusion | 10 | 3 | |
| histology → "Slicing" | 5 | 2 | |
| microscopy | 4 | 2 | note "512*512 / 1.5 zoom / 430um / 100 per plane, 42 planes" |
| Spines (custom) | 5 | 0 | plan stages, future |
| Reconstruction (custom) | 1 | 0 | plan stage |
| comment | 7 | 7 | "Switch to pbs", "PBS", "Make feeding request" |

- Animal section *Two Photon* holds a *FOVs* table.
- **8 האכלות tasks** cover weight → all become restriction periods.
- **Anomaly to flag, not fix:** T2NX-1 has a ration of 26.3 g (a weight in the
  ration box).
- Imaging design gains **resolution** and **frames per plane** (from that note).

## Storage decision

The generic problem is in how types are **defined and shown**, not how rows
are **stored and synced**. 2.0 records remain rows in the animal's history and
use the existing, tested sync/merge/sharing engine unchanged; each record kind's
shape, form, summary and meaning is **defined in code** (no user-defined
fields). Restriction periods and the specimen (brain) state are stored on the
animal and merged by id, like baselines. Generated rows (post-op checks,
feeding lines) use deterministic ids, like rounds.

Built on branch `arbor-2`; nothing deploys until all types are done and
verified together.

## Build plan

All types are built together, on shared foundations first.

**Phase 1 — foundations**
1. A record store for typed records (one shape per kind, each owning its own
   form, summary, hub section and export columns), syncing and sharing through
   the existing engine.
2. Derived timeline items: a source (restriction, surgery, specimen, protocol)
   produces dated items with done/warning state; the timeline draws them; only
   events are stored as calendar items.
3. Hub sections: each record kind contributes a section to the animal hub, and
   rows to experiment/project cohort tables.
4. One warnings layer (below 85%, overdue, not done today) feeding the hub,
   cards, timeline and notifications.

**Phase 2 — the types**, in dependency order: Weighing → Food restriction →
Surgery & Post-op → Behaviour (+ Lever) → Imaging → Tissue → Comment as notes.

**Phase 3 — migration & retirement**
- Each 1.0 entry type and the relevant tasks (e.g. האכלות) are migrated into
  the new records, checked against Yaron's real workspace (read-only copy)
  before anything is retired.
- Then the generic diary entry and generic task are retired; 1.0 data is kept
  in the backup until confirmed.

**Phase 4 — verification**: unit checks per type, real-data migration dry
run, two-device sharing test, both exporters' parity.
