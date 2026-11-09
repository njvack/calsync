# calsync

Smart-but-simple bidirectional Outlook <-> Google calendar syncing.

Two-way sync between your Google Calendar and Outlook calendar. Copies are attendee-less "busy" events with no reminders, tagged with a hidden property so the script only ever edits or deletes its own copies.

Each pass (every 90 s while the Mac is awake, and once on wake):

1. Fetch yesterday through +60 days from both calendars. Recurring events come back as individual occurrences, and occurrences are what get copied.
2. Ignore our own copies as sources, plus declined, cancelled, "free", and working-location events.
3. Two events are the same meeting if start and end are identical and one normalized title (lowercased, whitespace collapsed, leading `Copy:` stripped) is a prefix of the other, with the shorter at least 4 characters.
4. If the other calendar has a matching native event, don't copy, and delete our copy if one exists.
5. Otherwise create the copy (once the source hasn't changed for 60 s), or update it if the source moved or was renamed.
6. Delete copies whose source was deleted, declined, or marked free.

Private events are copied with the title "Busy". If any fetch fails, the run makes no changes.

## Setup

1. **Microsoft app:** under Authentication, add a "Mobile and desktop applications" platform with redirect URI `http://localhost`. Copy the client ID and tenant ID from Overview.
2. **Google app:** under Google Auth Platform › Clients, create an OAuth client of type **Desktop app** and download its JSON.
3. Install uv if needed: `brew install uv`.
4. Set up the config:

   ```sh
   mkdir -p ~/.config/calsync ~/bin
   cp config.example.toml ~/.config/calsync/config.toml        # then fill in the IDs
   mv ~/Downloads/client_secret_*.json ~/.config/calsync/google_client_secret.json
   cp calsync.py ~/bin/ && chmod +x ~/bin/calsync.py
   ```

5. Sign in to both accounts (opens a browser twice): `~/bin/calsync.py auth`. Add `--device` for a device-code sign-in to Microsoft instead.
6. Preview what it would do: `~/bin/calsync.py run --dry-run`. Run that a few times and read the output before going live.
7. Turn off CalendarBridge.
8. Install the launchd job:

   ```sh
   sed -e "s|UV_PATH|$(which uv)|" -e "s|SCRIPT_PATH|$HOME/bin/calsync.py|" \
       -e "s|LOG_PATH|$HOME/Library/Logs/calsync.log|" \
       local.calsync.plist > ~/Library/LaunchAgents/local.calsync.plist
   launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/local.calsync.plist
   tail -f ~/Library/Logs/calsync.log
   ```

To stop: `launchctl bootout gui/$(id -u)/local.calsync`.

To undo everything the script created: `~/bin/calsync.py ours` lists its copies, and `~/bin/calsync.py ours --delete` removes them.

## Notes

- Tokens are stored in `~/.config/calsync/` with mode 600. If Microsoft sign-in expires (your university's policies may force this occasionally), the log says so; rerun `auth`.
- Old CalendarBridge "Copy" events look like native events to this script, so they suppress copies rather than causing new ones. Remove them when convenient.
- To catch title variations the prefix rule misses, run `--dry-run` and look for `create` lines that duplicate something already on the other calendar.

## Tests

`python3 tests/test_plan.py` checks the matching and reconcile logic against fake calendars. It needs no dependencies or network access.
