# Sharing tests

Two or three Arbors, a fake GitHub, and the sharing paths driven end to end.
Everything asserts against the repository or the *other* device — never against
the function just called, which is how four bugs survived being reasoned about
and died on the first run.

## Run it

```bash
rm -rf tools/sharing-tests/.devices
bash tools/sharing-tests/build-device.sh sA WORKTREE
bash tools/sharing-tests/build-device.sh sB WORKTREE
python3 tools/sharing-tests/fake-github.py tools/sharing-tests/.devices 8777
```

Then open two tabs — **different origins, so they are different devices with
separate storage**:

| device | url |
|---|---|
| A | `http://localhost:8777/sA/index.html` |
| B | `http://127.0.0.1:8777/sB/index.html` |
| C | `http://[::1]:8777/sC/index.html` (build `sC` too) |

Before loading each one, seed its identity from that origin's `setup.html`:

```js
localStorage.clear();
localStorage.setItem('arbor_gist_token','fA');      // any string
localStorage.setItem('arbor_gist_id','gA1');        // DISTINCT per device
localStorage.setItem('arbor_space_token','fA');
localStorage.setItem('arbor_space_repo','lab/arbor-shared');
```

Then in each device's console:

```js
const s = await (await fetch('/lib.js')).text(); (0,eval)(s);
DB.me.name = 'Yaron'; persist(); await spaceJoin('Yaron');
```

`T` is now the suite. `T.mkProject/mkAnimal/mkDiary/mkTask/mkExperiment` build
fixtures, `T.push()` / `T.pull()` move data, `T.file(pid)` and `T.index()` read
the repository, `T.until(fn)` waits for a condition instead of guessing a delay,
and `T.ok/eq` record results for `T.report()`.

## Rules that earned themselves

- **Assert against the repo or the other device.** A function returning
  happily proves nothing: `publishProject` worked perfectly while the file it
  wrote stayed empty.
- **Re-fetch references after a push.** `attachProjectFile` replaces record
  objects, so a reference held across a sync is detached and `spOf()` is
  undefined on it. This bit the suite before it bit anyone else.
- **Check the rendered page, not just `DB`.** A project arrived correctly and
  was still undrawable (v245).
- **Test as the receiver.** Every bug in v242–v250 was on the receiving side;
  the account that creates everything never meets them.
- **After weakening a safety check, re-prove it still fails.** Narrowing
  `leakCheck`'s id scan needed four negative tests before it could be trusted.

## Also covered (second pass)

repeating tasks shared between people, and two or three of them ticking
different occurrences — including an un-tick not being resurrected by the
other side's tick · animal weights, surgery details, and the virus and site a
surgery names (and a virus it does not name staying behind) · an animal taken
out of a shared project leaving the collaborator while staying, private and
whole, with the person who took it back · an experiment and its tasks moved to
a private project · going offline with work queued on both sides and
reconnecting · three people editing different fields of one record at once.

## Also covered (third pass)

notes, papers, protocols and library records attached through a project's links
— travelling when linked, leaving when unlinked, never travelling when not ·
the same animal linked to two shared projects (it lives in exactly one file;
links pointing outside a file are dropped so nothing dangles) · deleting a
shared project outright, which must stop sharing it as well · "See as
collaborator": it hides private work, and writing, syncing and saving are all
refused while it is on · an audience narrowed after records exist · a diary
entry overriding its project's default in both directions · a project renamed
and recoloured · a record moved from one shared project to another.

## What is covered

definitions and libraries travelling and staying their author's · a shared
entry keeping its own duration · the resolver (private project, explicitly
private record, the ambiguous-animal rule) · no private id or text reaching the
file · concurrent edits to different fields of a record and of a diary entry ·
same-field collisions · delete → shared trash → restore, onto the right animal ·
unshare keeping your copy · withdrawing a project · per-project audience ·
discovery without a reload · and that `leakCheck` still catches a smuggled id,
smuggled text, text hidden in the activity log, and a private record placed in
the file.
