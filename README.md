# Outlook Room Booker

A Go program that auto-books Microsoft Outlook meeting rooms at midnight, the
instant a rolling 14-day window opens new slots for reservation. Built for
one company's resource-mailbox rooms (`room-*@example.com`) but reusable for any
Exchange tenant that exposes rooms as auto-accepting resource mailboxes.

## Why

Internal company rooms get claimed seconds after midnight. A human can't
compete; even one room takes minutes to book manually, and by then someone
else has the slot. This program prepares everything ahead of time and fires
`AppointmentItem.Send()` on the wall-clock second when the room becomes
reservable.

## Usage

Prerequisites:

- Windows + desktop Outlook signed into the booking account
- Go 1.25+
- The account that owns the booking (set via `ownerTag` in `main.go`) — or
  any account with delegate access to the resource mailboxes you want to
  book — must be the default Outlook profile
- Clock synced — the trigger fires at the local wall-clock second, so a
  300 ms-fast laptop fires before Exchange thinks the slot is open

Steps:

```
cd D:\Study\assessment
go run .
```

That's it. The program:

1. Measures NTP skew against `time.cloudflare.com`. If drift > 50 ms, the
   program prints the exact `Set-Date -Adjust` command you can paste into
   an elevated PowerShell to step the clock. The program does NOT sync
   automatically — `w32tm` requires admin every run, so the cleaner UX is
   to surface the precise correction and let you apply it once.
2. Kills any orphan `OUTLOOK.EXE` and starts COM workers attached to a
   fresh instance.
3. Sweeps the calendar (today through today + 14 days) for previous nights'
   leftover declined/pending bookings and deletes them.
4. Waits idle until 90 s before the next midnight (Asia/Bangkok).
5. Prepares all drafts in parallel — each (room, slot, attempt) tuple is a
   fully populated `AppointmentItem` with the room resolved against the
   directory.
6. At 00:00 local, fires `Send()` on every first-attempt draft with a 15 ms
   stagger between groups. Polls each room's response; on decline, fires
   the next attempt for that group (up to 10 attempts at 30 s intervals).

A `booking-YYYY-MM-DD.log` file is written for every run (date-stamped to
the day the program started in Asia/Bangkok).

## Configuration

All knobs live at the top of `main.go`:

| Constant | Default | What it controls |
|---|---|---|
| `ownerTag` | `"YOURTAG"` | Short identifier embedded in the meeting subject. Change this when a different user runs the program — the cleanup matcher derives its prefixes from this too. |
| `rooms` | 9 example rooms | Resource mailboxes to book. Each must resolve to a real Exchange recipient. |
| `slots` | 09:00–12:00, 14:00–17:00 | Time slots booked on each target day. |
| `bookingDaysAhead` | 14 | Rolling window — at midnight day N, the program books day N+14. |
| `triggerHr`, `triggerMin` | 00:00 | Local trigger time. Set to a few minutes ahead for smoke tests. |
| `localTZName` | Asia/Bangkok | Wall-clock zone the trigger evaluates in. |
| `prepLead` | 90 s | How far ahead of the trigger drafts get prepared. |
| `sendAttempts` | 10 | Max retries per (room, slot) on decline. Total retry window = `sendAttempts × pollInterval`. |
| `pollInterval` | 30 s | Time the program waits for the room's response after each `Send` before retrying. |
| `pollEvery` | 1 s | Cadence of polling the room's reply during the `pollInterval` window. |
| `sendStagger` | 15 ms | Delay between concurrent first-Sends across groups (avoids Outlook WordEditor contention). |
| `meetingSubject` | `"Room booking for " + ownerTag` | Subject of every booked meeting. Derived from `ownerTag`. |
| `cleanupSubjects` | derived | Subject prefixes considered "ours" by cleanup. Derived from `ownerTag`; supports historical variants. |
| `responseFolderName` | `"Room booking"` | Inbox subfolder where the user's rule routes meeting responses. Searched for orphaned declines during pass-2 cleanup. |

## Architecture

### Why Outlook COM, not Microsoft Graph

Graph requires tenant-admin consent for `Calendars.ReadWrite` and
`Mail.Send` on the user's behalf, which was refused. Desktop Outlook is
already signed in and authenticated, so we drive it directly via OLE
Automation (`github.com/go-ole/go-ole`) — same Object Model the VBA macros
use. No extra credentials, no admin signoff.

### Per-(room, slot) COM worker pool

Outlook's COM server is a Single-Threaded Apartment (STA). Cross-apartment
calls serialize through the marshaller, so a single goroutine ends up the
bottleneck even though Outlook itself can handle multiple concurrent Sends
internally.

Solution: one goroutine per (room, slot) tuple, each pinned to a unique OS
thread (`runtime.LockOSThread`) and owning its own STA apartment
(`CoInitializeEx(COINIT_APARTMENTTHREADED)`) and its own
`Outlook.Application` proxy. With 9 rooms × 2 slots = 18 workers (+ 1 for
cleanup), 18 Sends can fire at the same wall-clock millisecond from 18
independent COM apartments. Outlook still has internal serialization
points, but the Go-side queueing wait is eliminated.

Tradeoff: pure firehose 18× simultaneous Sends still collide on Outlook's
shared `WordEditor`, surfacing as `"text formatting command is not
available"`. We solve that with a 15 ms send stagger (see below).

### Drafts prepared 90s before the trigger

The trigger's hot path is just `appt.Send()` — that's the only step we
can't perform early. Everything else (CreateItem, set properties, add
recipient, ResolveAll) happens during the `prepLead` window. 18 workers
preparing 10 attempts each = 180 drafts, parallel across groups,
finishing in ~1–2 s. By midnight every draft is fully populated and the
trigger window is dominated by one Send round-trip per group.

### Send stagger

Without staggering, 18 first-Sends at the same microsecond cause Outlook's
shared `WordEditor` to refuse most of them. 15 ms between launches gives
each Send enough room to clear that internal before the next starts —
total fan-out for 18 groups is ~270 ms, well under any meaningful
trigger window.

### Transient COM retry

`prepareDraft` and `sendDraft` both retry on:

- `RPC_E_CALL_REJECTED` (0x80010001)
- `RPC_E_SERVERCALL_RETRYLATER` (0x8001010A)
- `"text formatting command is not available"` (WordEditor contention)

Backoffs are 30/80/160/320 ms, max 4 attempts. These all clear within tens
of ms in practice.

### Outlook-died recovery (RPC_S_SERVER_UNAVAILABLE)

If `OUTLOOK.EXE` dies mid-run (auto-update, crash, killed), `w.app` becomes
a dead RPC handle and every call returns `0x800706BA`. `prepareDraft`
detects this and reattaches in place: `CreateObject("Outlook.Application")`
on the same pinned thread, replacing `w.app` with a fresh proxy. Drafts
prepared via the dead handle stay dead — their Send will fail and the
retry sequence moves on — but new prep calls succeed.

### Smart retry loop, not bulk fire-and-forget

A naive approach would queue all 10 attempts × 18 groups upfront. That
creates 180 racing requests against Exchange, with no way to abort once
one room accepts. We do the opposite: for each (room, slot), the first
Send fires at the trigger, we poll the response for 30 s, and only on a
**decline** do we fire the next attempt. On accept (or no decisive
response within `sendAttempts × pollInterval`), we stop.

Why retry at all? Other people are also racing. A competitor who books
first sometimes releases their slot seconds later when they notice they
got the wrong room. A 5-minute retry window catches those late releases
without spamming Exchange.

### Startup-only cleanup

Earlier versions ran a calendar sweep 15 min after the trigger to delete
declined/pending meetings. This was wrong: room moderators can take
**hours** to approve a booking, and the +15 min sweep was deleting
accepts-in-flight. Symptom: accept emails arrived at noon for meetings
that no longer existed on the calendar.

Fixed by moving cleanup to startup. Every run sweeps the previous nights'
leftovers (declined + still-pending), but **never** touches the freshly
sent batch. Pending items from the current run stay on the calendar all
day; whatever the room approves stays, whatever doesn't gets cleared at
the start of the next run.

Cleanup is two passes:

1. **Calendar scan** — `Restrict` calendar items by date range, filter by
   subject prefix, delete if room status is Declined or Pending.
2. **Response folder scan** — walk the user's `"Room booking"` Inbox
   subfolder (and Inbox itself, since empty-subject declines bypass
   rules). For each `"Declined:"` email, call
   `GetAssociatedAppointment(false)` to link back to the original meeting,
   confirm it's ours (subject prefix) and within the cleanup window, and
   delete the appointment directly. Catches stuck-pending items whose
   decline emails Outlook's auto-process didn't apply.

### Clock skew diagnostic

The trigger fires at the laptop's wall-clock 00:00:00.000 in
Asia/Bangkok. Exchange uses its own atomic-clock-disciplined time. If the
laptop is 300 ms ahead, the program sends 300 ms before Exchange opens
the slot — the request gets rejected as "not yet available."

At startup the program runs 5 NTP queries against
`time.cloudflare.com` (anycast, low and symmetric RTT from VN), reports
the median offset and the spread. A spread tight to a few ms means the
displayed offset is real drift; a wide spread means path noise and the
clock is probably fine.

If drift exceeds 50 ms, the program prints the exact `Set-Date` command
to apply — copy it into an elevated PowerShell:

```
[debug]   Set-Date -Adjust ([TimeSpan]::FromMilliseconds(-843))
```

No auto-sync: `w32tm /resync` requires admin, the scheduled-task
workaround needs one-time SYSTEM setup, and Windows W32Time slews phase
offsets under 1 s rather than stepping. `Set-Date -Adjust` steps
immediately and the program tells you the precise value to use.

### Date-stamped logs

Every run is tee'd to `booking-YYYY-MM-DD.log` in the working directory
(date = start time in Asia/Bangkok, append mode for same-day re-runs).
Console output is unchanged; the file gets an exact copy via
`io.MultiWriter` over an `os.Pipe()` replacing `os.Stdout`.

## Limitations

- Single-user, single-laptop: Outlook is desktop-bound and the COM
  apartment dies if the user logs out.
- Doesn't handle the case where Outlook *prompts* for the
  Object Model Guard ("A program is trying to send..."). Disable that
  prompt in Trust Center the first time it appears.
- If Outlook dies and the reattach also fails (no `Outlook.Application`
  registered), the program logs the error and continues but every
  subsequent COM call will fail. Restart the program to recover.
- Clock skew is reported, not corrected. Sync the OS clock yourself.

## License

Personal use. No warranty.
