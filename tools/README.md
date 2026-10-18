# Arbor tools

## Twice-daily export and backup

`arbor_export.py` reads the workspace — from the synced Gist, or from a JSON
file — and writes a styled `.xlsx`. It runs outside the browser, so the
scheduled export does not need Arbor to be open, and it produces a real xlsx
rather than the SpreadsheetML the in-app button used to emit, which made Excel
warn about the file format on every open.

```
arbor_export.py --gist <id> --out <dir>      # what the scheduled job runs
arbor_export.py --json <file> --xlsx <file>  # one-off, from a saved snapshot
```

Each scheduled run also saves the raw JSON it exported from, so
`~/Arbor exports` is a backup that depends on nothing else — not a browser, not
GitHub, not Arbor. The last 30 of each are kept.

### The schedule

`install_export_job.sh` installs a launchd agent (`com.arbor.export`) that runs
at 08:00 and 20:00. If the Mac is asleep at those times, launchd runs the job
when it next wakes.

**Re-run `install_export_job.sh` after changing `arbor_export.py`.** launchd
runs an installed copy in `~/.local/share/arbor`, because macOS refuses
background jobs any access to `~/Desktop`, where this repo lives. That is also
why exports go to `~/Arbor exports`, with a symlink from
`Lab/Arbor exports` so they show up where you would look for them.

```
launchctl kickstart -p gui/$UID/com.arbor.export   # run it now
launchctl bootout   gui/$UID/com.arbor.export      # stop it
tail -f "$HOME/Arbor exports/export.log"           # watch it
```

The job refuses to write an export from an empty-looking workspace, and a
failed run leaves previous exports untouched.
