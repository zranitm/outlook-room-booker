# Outlook Room Booker

Auto-books Microsoft Outlook meeting rooms at midnight, the instant a
rolling 14-day window opens new slots. Built for one company's resource-mailbox
rooms (`room-*@example.com`) but reusable for any Exchange tenant that
exposes rooms as auto-accepting resource mailboxes.

Internal rooms get claimed seconds after midnight. This program prepares
every draft ahead of time and fires `AppointmentItem.Send()` on the
wall-clock second the room becomes reservable.

## Prerequisites

- Windows + desktop Outlook signed into the booking account
- Go 1.25+
- Booking account (or a delegate) must be the default Outlook profile
- Before first run, edit `main.go`:
  - `ownerTag` — your short identifier (appears in every subject)
  - `rooms` — resource mailboxes to claim
  - `localTZName` — your timezone if not `Asia/Bangkok`
  - `responseFolderName` — Inbox subfolder for room responses (see
    "Optional: Room booking folder" below)

### Optional: Room booking folder

`responseFolderName = "Room booking"` is a hint for the second cleanup
pass — it tells the program where you've routed room accept/decline
emails. **Not strictly required**: if the folder doesn't exist, the
pass falls back to scanning Inbox directly (logs `no non-empty "Room
booking" folder found, will rely on Inbox scan`).

Recommended setup (Outlook web or desktop):

1. Inbox → right-click → **New folder** → name it `Room booking`.
2. Add a rule: *from* contains `room-` (or whichever pattern your
   rooms use), *move to* `Room booking`.

Pass-2 cleanup runs faster against a small dedicated folder than a
4000-item Inbox. Set `responseFolderName = ""` to skip the folder
lookup entirely if you don't want to set it up.

> **Sender identity vs `ownerTag`.** The program does not specify an email
> address anywhere — it talks to whatever Outlook profile is currently
> running, and meetings are sent **as that account**. `ownerTag` is only
> a label baked into the subject line so people can tell whose booking is
> whose. Switching the booking account = sign in to Outlook with a
> different profile (no code change needed). Optionally update `ownerTag`
> to match.

## Commands

```sh
# Normal run — startup cleanup + scheduler waits for midnight
go run .

# One-shot: delete declined/pending bookings on a specific day, then exit
go run . --clear-day=2026-07-15

# Manual clock step (elevated PowerShell, value from the [debug] log line)
Set-Date -Adjust ([TimeSpan]::FromMilliseconds(-840))
```

### Debug logs

Every run is tee'd to a date-stamped file in the working directory —
`booking-YYYY-MM-DD.log` (date = local Asia/Bangkok at process start,
append-mode so multiple same-day runs accumulate). Console output is
unchanged; the file gets an exact copy.

If something looks wrong the morning after a run — wrong number of
wins, unexpected declines, cleanup deleting too much — open last
night's log and grep for `[trigger]`, `[poll]`, `[cleanup]`, or
`[recover]`. The log captures every Send attempt, every room response,
every cleanup decision and the reason for it. Paste the relevant
chunk into Claude and ask what went wrong; the log is structured
enough that the failure mode is usually obvious from a few lines.

## What `go run .` does

1. Measures NTP skew vs `time.cloudflare.com`. If > 50 ms, prints the
   exact `Set-Date -Adjust` command for you to paste into an elevated
   PowerShell.
2. Kills orphan `OUTLOOK.EXE` and spawns COM workers attached to a fresh
   Outlook instance.
3. Sweeps the calendar (today through today + 14 days) for previous
   nights' leftover declined/pending bookings and deletes them.
   Accepted meetings are kept.
4. Idles until 90 s before the next midnight, then prepares all drafts
   in parallel.
5. At 00:00 local, fires `Send()` per (room, slot) with a 15 ms stagger.
   Polls each room's response; on decline, retries (up to 10 attempts at
   30 s intervals — 5-minute total window).

## Configuration

All knobs live at the top of `main.go`:

| Constant | Default | What it controls |
|---|---|---|
| `ownerTag` | `"YOURTAG"` | Identifier embedded in every meeting subject. Cleanup matcher derives its prefixes from this too. |
| `rooms` | 9 example rooms | Resource mailboxes to book. Each must resolve in Exchange. |
| `slots` | 09:00–12:00, 14:00–17:00 | Time slots booked on each target day. |
| `bookingDaysAhead` | 14 | Rolling window — at midnight day N, the program books day N+14. |
| `triggerHr`, `triggerMin` | 00:00 | Local trigger time. Set to a few minutes ahead for smoke tests. |
| `localTZName` | Asia/Bangkok | Wall-clock zone the trigger evaluates in. |
| `prepLead` | 90 s | How far ahead of the trigger drafts get prepared. |
| `sendAttempts` | 10 | Max retries per (room, slot) on decline. |
| `pollInterval` | 30 s | Wait for the room's reply after each `Send` before retrying. |
| `sendStagger` | 15 ms | Delay between concurrent first-Sends (avoids WordEditor contention). |
| `responseFolderName` | `"Room booking"` | Inbox subfolder for room responses. |

## Architecture (key decisions)

**COM over Microsoft Graph.** Graph needs tenant-admin consent for
`Calendars.ReadWrite` and `Mail.Send`, which was refused. Desktop
Outlook is already authenticated, so we drive it via OLE Automation
(`github.com/go-ole/go-ole`) — no extra credentials, no admin signoff.

Tradeoff: COM is a Windows-only OS feature (lives in
`ole32.dll`/`oleaut32.dll`), and `go-ole` won't compile or run on
macOS or Linux. There's no equivalent on those platforms — macOS
Outlook exposes a different automation API (AppleScript/ScriptingBridge),
and Linux has no Outlook client at all. The only path off Windows
would be to swap COM for Graph API, which puts us back at the
admin-consent wall.

**Per-(room, slot) COM worker pool.** Outlook's COM server is STA;
cross-apartment calls serialize through the marshaller. We spawn one
goroutine per (room, slot), each pinned to its own OS thread
(`runtime.LockOSThread`) with its own STA apartment and its own
`Outlook.Application` proxy. With 9 rooms × 2 slots = 18 workers, 18
Sends fire from 18 independent apartments — the Go-side queueing wait
disappears.

**Drafts prepared 90 s ahead.** The trigger's hot path is just
`appt.Send()`. Everything else (CreateItem, set props, ResolveAll)
happens during `prepLead`. By midnight every draft is ready.

**15 ms send stagger.** Without it, 18 first-Sends collide on Outlook's
shared `WordEditor` and most fail with "text formatting command is not
available." 15 ms between launches clears that contention.

**Decline-only retry, not bulk fire-and-forget.** First Send fires at
trigger; we poll the room's response for 30 s; only on **decline** do we
retry. Catches competitors who book first and release seconds later
without spamming Exchange.

**Startup-only cleanup.** Earlier versions ran a calendar sweep 15 min
after the trigger and deleted accepts-in-flight (room moderators can
take hours to approve). Now cleanup runs once at startup — the previous
night's leftovers go, the current night's batch stays untouched.

Cleanup is two passes:
1. Calendar scan by date + subject prefix, delete if Declined or Pending.
2. Walk the response folder (and Inbox, for empty-subject declines) for
   "Declined:" emails, call `GetAssociatedAppointment(false)` to link
   back, delete the appointment.

**Outlook-died recovery.** If `OUTLOOK.EXE` dies mid-run, `w.app`
returns `RPC_S_SERVER_UNAVAILABLE` (0x800706BA). `prepareDraft` detects
this, reattaches via `CreateObject("Outlook.Application")` on the same
pinned thread, and retries.

**Clock skew matters because the race is millisecond-tight.** The
trigger fires when the **laptop's** wall-clock reads `00:00:00.000`
Asia/Bangkok. Exchange evaluates "is this slot reservable yet?"
against **its own** clock (NTP-disciplined, treated as ground truth).
If the laptop is 300 ms fast, our Send hits Exchange at *its*
`23:59:59.700` — before the rolling 14-day window opens — and the
request gets bounced as "slot not yet available." Conversely, a slow
laptop just hands the race to whoever's clock is more accurate. We
need to be within a few tens of ms of true network time for the trigger
to fire at the right Exchange moment.

**Why manual `Set-Date`, not auto-sync.** Three reasons:
1. `w32tm /resync` and the `SetSystemTime` API both require admin. The
   program runs unelevated by design (no UAC prompt every night), so
   it can't sync the OS clock itself.
2. The standard workaround — a Scheduled Task running as SYSTEM that
   the program triggers via `schtasks /run` — needs one-time admin
   setup that often doesn't stick and adds a moving part that can
   silently break.
3. Even if `w32tm /resync` runs, Windows W32Time **slews** phase
   offsets under 1 second (gradually adjusting clock rate over
   minutes) instead of **stepping** them. So a 500 ms-fast clock
   stays 500 ms-fast for the booking window.

`Set-Date -Adjust` steps immediately and is a one-line PowerShell
command. The program measures the residual offset post-NTP and prints
the exact `Set-Date` value to apply — you copy, paste into an elevated
PS, and the clock is correct.

## Limitations

- **Windows only.** COM is a Windows OS feature; the program won't
  compile or run on macOS or Linux (no containerization either —
  Outlook isn't supported in Windows Server Core / Nano images, and
  the profile setup is interactive).
- Single-user, single-laptop. Outlook is desktop-bound.
- Doesn't handle the Object Model Guard prompt ("A program is trying to
  send…"). Disable it in Trust Center the first time.
- Clock skew is reported, not corrected.
