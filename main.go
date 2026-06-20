package main

// Set-Date -Adjust ([TimeSpan]::FromMilliseconds(-298))
import (
	"fmt"
	"io"
	"os"
	"os/exec"
	"runtime"
	"strings"
	"sync"
	"time"

	"github.com/beevik/ntp"
	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

// ---------------------------------------------------------------------------
// Configuration — fill these in before running.
// ---------------------------------------------------------------------------

// Rooms to attempt to book in parallel. Each entry is the room's mailbox.
// Outlook treats rooms as resource mailboxes.
var rooms = []Room{
	// {Email: "room-a@example.com", Name: "Room A - Floor 1 - 6 Seats"},
	// {Email: "room-b@example.com", Name: "Room B - Floor 1 - 6 Seats"},
	// {Email: "room-c@example.com", Name: "Room C - Floor 1 - 6 Seats"},
	// {Email: "room-d@example.com", Name: "Room D - Floor 1 - 6 Seats"},
	{Email: "room-e@example.com", Name: "Room E - Floor 1 - 6 Seats"},
	{Email: "room-f@example.com", Name: "Room F - Floor 1 - Lobby - 12 Seats"},
	{Email: "room-g@example.com", Name: "Room G - Floor 1 - Lobby - 8 Seats"},
	{Email: "room-h@example.com", Name: "Room H - Floor 1 - 6 Seats"},
	{Email: "room-i@example.com", Name: "Room I - Floor 1 - 8 Seats"},
}

// Booking owner — short identifier embedded in every meeting subject so
// the booker is recognizable in calendar listings. Change this when a
// different user runs the program; the cleanup matcher derives its
// subject prefixes from this value too, so a single edit re-tags the
// whole pipeline.
const ownerTag = "YOURTAG"

// Meeting template applied to every booking attempt.
const (
	meetingSubject = "Room booking for " + ownerTag
	meetingBody    = "Reserved automatically."

	// Number of days between trigger date and meeting date. Matches the
	// rolling booking window — at midnight on day N, the slot for day
	// N+bookingDaysAhead just became reservable.
	bookingDaysAhead = 14
)

// Time slots booked for each room. Add / remove entries to change
// which parts of the day get reserved at the trigger. Every entry
// here will be booked for every room.
var slots = []meetingSlot{
	{StartHr: 9, StartMin: 0, Duration: 180 * time.Minute},  // 09:00 – 12:00
	{StartHr: 14, StartMin: 0, Duration: 180 * time.Minute}, // 14:00 – 17:00
}

type meetingSlot struct {
	StartHr  int
	StartMin int
	Duration time.Duration
}

// IANA zone name used for the trigger evaluation. Must match the timezone
// the rooms are released in (GMT+7 in the original deployment).
const localTZName = "Asia/Bangkok"

// Booking trigger: local hh:mm at which the bookings fire. Defaults to
// 00:00 so a new day's rooms are claimed the instant they open. For
// smoke-testing, set this to a few minutes in the future.
const (
	triggerHr  = 00
	triggerMin = 00
)

// How far ahead of the trigger we prepare the meeting drafts. Each draft
// is a fully-populated Outlook AppointmentItem with a resolved room
// recipient — at the trigger we only call Send() on it, which is the
// only step we can't run early. With sendAttempts=10 we prep 10 × 4
// rooms × 2 slots = 80 drafts; the per-(room, slot) workers run in
// parallel so wallclock is bounded by ~10 sequential preps, but the
// extra headroom also gives Outlook ~30s of idle time AFTER the last
// prep before the trigger fires — without that buffer, Outlook is
// still settling from prep activity when sends start, making the
// first-Send Outlook time balloon to 700+ms.
const prepLead = 90 * time.Second

// Maximum invites per (room, slot). The first invite fires at the
// trigger; if the room is declined within pollInterval, the next
// attempt fires. The retry only happens on a decline — never on an
// accept — so we never create our own conflicts. 10 attempts × 30s
// pollInterval gives a 5-minute window to catch late cancellations
// from competitors who released their slot a few minutes after the
// initial midnight race.
const sendAttempts = 10

// pollInterval — maximum total time to wait for the room's reply
// after each Send. If the room hasn't given a definitive accept or
// decline by then, we move to the next attempt. 5s was too short
// in practice (Exchange routing + Outlook sync took longer than
// that), so we use a more generous window.
const pollInterval = 30 * time.Second

// pollEvery — cadence at which we re-check the room's reply during
// the pollInterval window. As soon as we see Accepted (we won) or
// Declined (lost the race, send retry), we exit early instead of
// waiting the full pollInterval.
const pollEvery = 1000 * time.Millisecond

// sendStagger — delay between consecutive goroutine first-Sends at
// the trigger. Without staggering, 8 worker apartments hitting
// Outlook simultaneously contend for Outlook's shared WordEditor and
// most Sends fail with "text formatting command is not available."
// 15ms is enough room for each Send to clear that internal before
// the next starts; total fan-out cost for 8 groups is ~105ms (the
// 8th group's first send), which is still well under the previous
// single-worker fan-out time.
const sendStagger = 15 * time.Millisecond

// Subjects considered "ours" by cleanup. We match by PREFIX, so
// historical variants that appended room names (e.g. "Room booking
// for <owner> - Room A - ...") are still recognized.
var cleanupSubjects = []string{
	"Room booking for " + ownerTag,
	ownerTag + " Room booking",
}

// Outlook Inbox subfolder where the user's rule routes meeting
// accept/decline responses. We also walk this folder during cleanup
// to catch decline responses that Outlook's auto-process didn't apply
// to the corresponding appointments (which leaves them stuck "Pending"
// on the calendar). Empty string disables the second-pass folder scan.
const responseFolderName = "Room booking"

// isOurSubject reports whether s starts with one of the subjects we
// consider ours. Whitespace is trimmed, comparison is case-sensitive
// (Outlook preserves case as set).
func isOurSubject(s string) bool {
	s = strings.TrimSpace(s)
	for _, want := range cleanupSubjects {
		if strings.HasPrefix(s, want) {
			return true
		}
	}
	return false
}

// NTP server we query at startup to estimate how far the laptop clock
// drifts from the network. Cloudflare's anycast NTP routes to a nearby
// PoP from VN, giving a symmetric round-trip — important because the
// skew formula assumes symmetric latency and asymmetric paths leak
// noise (~hundreds of ms against time.windows.com). Diagnostic only —
// the program does NOT compensate, just reports. Requires UDP/123
// outbound; if blocked, the check logs an error and startup continues.
const clockSkewSource = "time.cloudflare.com"

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type Room struct {
	Email string
	Name  string
}

// preparedDraft pairs a prepared (but unsent) Outlook AppointmentItem
// IDispatch handle with the room and slot it targets and which attempt
// number it represents (1..sendAttempts), so sendAll can label log
// lines unambiguously when there are multiple slots per room.
type preparedDraft struct {
	room    Room
	appt    *ole.IDispatch
	attempt int
	start   time.Time
}

// bookMetrics breaks down the latency of a single sendDraft call.
//
//	Wait     time queued behind the COM worker
//	Outlook  time spent inside Send() on the worker thread
//	Total    from the trigger tick to function return
type bookMetrics struct {
	Wait    time.Duration
	Outlook time.Duration
	Total   time.Duration
}

func (m bookMetrics) String() string {
	return fmt.Sprintf("queued %dms, outlook %dms, total %dms",
		m.Wait.Milliseconds(),
		m.Outlook.Milliseconds(),
		m.Total.Milliseconds())
}

// ---------------------------------------------------------------------------
// Clock sync + skew check
// ---------------------------------------------------------------------------

// clockSyncTaskName is a pre-registered Windows Scheduled Task that
// runs as SYSTEM (highest privileges) so it can start the W32Time
// service and call w32tm /resync without prompting for elevation.
// Created once, then triggered every run from this non-elevated
// process via `schtasks /run`. See README → Clock sync diagnostic.
const clockSyncTaskName = "RoomBooker-TimeSync"

// syncSystemClock triggers the pre-registered scheduled task to sync
// the OS clock against the W32Time peer list. Falls back to a helpful
// log message if the task isn't registered — measureClockSkew still
// runs immediately after, so the user sees the remaining drift either
// way.
func syncSystemClock() {
	out, err := exec.Command("schtasks", "/run", "/tn", clockSyncTaskName).CombinedOutput()
	msg := strings.TrimSpace(string(out))
	if err != nil {
		fmt.Printf("[init] schtasks /run %s failed: %v (%s)\n", clockSyncTaskName, err, msg)
		fmt.Println("[init] one-time setup — paste into an elevated PowerShell:")
		fmt.Println(`[init]   $a = New-ScheduledTaskAction -Execute 'cmd.exe' -Argument '/c net start w32time & w32tm /resync & net stop w32time'`)
		fmt.Println(`[init]   $t = New-ScheduledTaskTrigger -AtStartup`)
		fmt.Println(`[init]   $p = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -RunLevel Highest`)
		fmt.Printf("[init]   Register-ScheduledTask -TaskName '%s' -Action $a -Trigger $t -Principal $p -Force\n", clockSyncTaskName)
		return
	}
	fmt.Printf("[init] time sync triggered: %s\n", msg)
	// Task runs async (net start + w32tm /resync + net stop chain).
	// Sleep long enough for the resync to land before the skew check.
	time.Sleep(4 * time.Second)
}

// clockSkewSample holds one NTP query result. We collect several to
// distinguish real drift (offsets cluster tightly) from path-asymmetry
// noise (offsets correlate with RTT).
type clockSkewSample struct {
	Offset time.Duration // positive = local ahead of server
	RTT    time.Duration
}

// measureClockSkew issues a handful of NTP queries to clockSkewSource
// and returns the median-offset sample. Single-shot NTP can be off by
// hundreds of ms on asymmetric paths; the median across samples damps
// that down to something representative of real drift. We log min/max
// RTT and a min/max offset spread so the user can tell measurement
// noise from real skew.
func measureClockSkew() (clockSkewSample, error) {
	const n = 5
	samples := make([]clockSkewSample, 0, n)
	var firstErr error
	for i := 0; i < n; i++ {
		resp, err := ntp.Query(clockSkewSource)
		if err != nil {
			if firstErr == nil {
				firstErr = err
			}
			continue
		}
		// resp.ClockOffset is "server minus local"; flip so positive
		// means the laptop is ahead of the server.
		samples = append(samples, clockSkewSample{
			Offset: -resp.ClockOffset,
			RTT:    resp.RTT,
		})
		time.Sleep(50 * time.Millisecond)
	}
	if len(samples) == 0 {
		return clockSkewSample{}, firstErr
	}
	// Median by offset — robust to a single bad packet.
	sortSamplesByOffset(samples)
	med := samples[len(samples)/2]
	// Report spread to give the user signal about reliability.
	minOff, maxOff := samples[0].Offset, samples[len(samples)-1].Offset
	var minRTT, maxRTT time.Duration = samples[0].RTT, samples[0].RTT
	for _, s := range samples {
		if s.RTT < minRTT {
			minRTT = s.RTT
		}
		if s.RTT > maxRTT {
			maxRTT = s.RTT
		}
	}
	fmt.Printf("[debug] ntp samples=%d offset range=[%s..%s] rtt range=[%s..%s]\n",
		len(samples),
		minOff.Truncate(time.Millisecond), maxOff.Truncate(time.Millisecond),
		minRTT.Truncate(time.Millisecond), maxRTT.Truncate(time.Millisecond))
	return med, nil
}

func sortSamplesByOffset(s []clockSkewSample) {
	// Tiny n — insertion sort avoids importing sort for one line.
	for i := 1; i < len(s); i++ {
		for j := i; j > 0 && s[j-1].Offset > s[j].Offset; j-- {
			s[j-1], s[j] = s[j], s[j-1]
		}
	}
}

// ---------------------------------------------------------------------------
// Outlook process hygiene — kill any orphan before COM startup
// ---------------------------------------------------------------------------

// killExistingOutlook taskkills any running OUTLOOK.EXE so the COM
// worker's subsequent CreateObject launches a fresh instance in the
// current user's session. Without this, an orphan Outlook left over
// from a different security context (e.g. a prior session or another
// user) gets COM-attached to instead, and every CreateItem call fails
// with "Exception occurred. (The operation failed.)" because the
// IDispatch points at a process that can't reach this user's mailbox.
//
// taskkill /F succeeds silently if no process is running. If the kill
// is denied (e.g. orphan owned by another user and we're not elevated),
// we log it and proceed — the downstream COM error is diagnostic
// enough, and the user knows to retry from an elevated shell.
func killExistingOutlook() {
	cmd := exec.Command("taskkill", "/F", "/IM", "OUTLOOK.EXE")
	out, err := cmd.CombinedOutput()
	output := strings.TrimSpace(string(out))

	switch {
	case err == nil:
		fmt.Printf("[init] %s\n", output)
		// Brief pause so the OS finishes releasing the process before
		// we ask COM to launch a fresh one.
		time.Sleep(500 * time.Millisecond)
	case strings.Contains(output, "not found"):
		// No Outlook running — nothing to do.
	default:
		fmt.Printf("[init] taskkill OUTLOOK.EXE failed: %v (%s)\n", err, output)
		fmt.Println("[init] if booking fails, retry from an elevated PowerShell so the orphan can be killed")
	}
}

// ---------------------------------------------------------------------------
// COM worker pool — one OS thread + STA + Outlook proxy per (room, slot)
// ---------------------------------------------------------------------------

// Outlook Object Model constants.
const (
	olAppointmentItem = 1
	olMeeting         = 1
	olOrganizer       = 0 // Recipient.Type
	olRequired        = 1
	olOptional        = 2
	olResource        = 3

	olResponseNone         = 0 // Recipient.MeetingResponseStatus
	olResponseOrganized    = 1
	olResponseTentative    = 2
	olResponseAccepted     = 3
	olResponseDeclined     = 4
	olResponseNotResponded = 5

	olFolderCalendar = 9 // NameSpace.GetDefaultFolder
	olFolderInbox    = 6
)

const rpcEChangedMode = 0x80010106

// comWorker is a goroutine pinned to one OS thread with its own COM
// apartment and Outlook proxy. All AppointmentItem operations for a
// given (room, slot) run on the same comWorker so the IDispatch never
// crosses apartments. Multiple comWorkers running concurrently let
// goroutines submit Send() calls to Outlook from independent apartments;
// Outlook itself may still serialize internally, but the Go-side
// queueing wait is eliminated, so all first-attempt sends can fire
// at the trigger without waiting for prior sends to complete.
type comWorker struct {
	queue chan func()
	label string
	app   *ole.IDispatch // cached Outlook.Application IDispatch for this apartment
}

func newComWorker(label string) (*comWorker, error) {
	w := &comWorker{queue: make(chan func(), 4), label: label}
	ready := make(chan error, 1)
	go func() {
		runtime.LockOSThread() // never released — apartment must outlive bookings
		if err := ole.CoInitializeEx(0, ole.COINIT_APARTMENTTHREADED); err != nil {
			if oleErr, ok := err.(*ole.OleError); !ok || oleErr.Code() != rpcEChangedMode {
				ready <- fmt.Errorf("CoInitializeEx (%s): %w", w.label, err)
				return
			}
		}
		unk, err := oleutil.CreateObject("Outlook.Application")
		if err != nil {
			ready <- fmt.Errorf("CreateObject Outlook.Application (%s): %w", w.label, err)
			return
		}
		app, err := unk.QueryInterface(ole.IID_IDispatch)
		if err != nil {
			ready <- fmt.Errorf("QueryInterface (%s): %w", w.label, err)
			return
		}
		// Both unk and app hold refs to Outlook — intentionally never
		// Released, so Outlook stays alive for the program's lifetime.
		w.app = app
		// Warmup: pay the cold-marshalling cost now rather than at the
		// trigger. One trivial property read is enough to prime the path.
		_, _ = oleutil.GetProperty(app, "Name")
		ready <- nil
		for fn := range w.queue {
			fn()
		}
	}()
	return w, <-ready
}

// do submits fn to the worker and blocks until it has run.
func (w *comWorker) do(fn func()) {
	done := make(chan struct{})
	w.queue <- func() {
		defer close(done)
		fn()
	}
	<-done
}

// reattachOutlookLocked replaces w.app with a fresh Outlook.Application
// IDispatch. Used after the cached handle starts returning
// RPC_S_SERVER_UNAVAILABLE — Outlook.exe died (crash, auto-update,
// killed) and CreateObject will launch / attach to a fresh instance.
// Must be called from inside a w.do() callback because the COM apartment
// is pinned to the worker's OS thread; calling it from another goroutine
// is undefined behavior. The previous w.app is leaked intentionally —
// Release would just call into the dead process and hang.
func (w *comWorker) reattachOutlookLocked() error {
	unk, err := oleutil.CreateObject("Outlook.Application")
	if err != nil {
		return fmt.Errorf("CreateObject Outlook.Application (%s): %w", w.label, err)
	}
	app, err := unk.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		return fmt.Errorf("QueryInterface (%s): %w", w.label, err)
	}
	w.app = app
	return nil
}

// Workers are keyed by (room email, slot start hh:mm). Every COM
// operation for a given (room, slot) — prepareDraft, sendDraft,
// getRoomResponse, releaseDraft — always hits the same worker, so the
// IDispatch never escapes its apartment.
type slotKey struct {
	email    string
	startHr  int
	startMin int
}

var (
	slotWorkers   map[slotKey]*comWorker
	cleanupWorker *comWorker
)

// startWorkers spawns one comWorker per (room, slot) plus one
// dedicated to the calendar-sweep cleanup. Returns when every worker
// has its own Outlook proxy ready.
func startWorkers() error {
	slotWorkers = map[slotKey]*comWorker{}
	for _, r := range rooms {
		for _, sl := range slots {
			label := fmt.Sprintf("%s@%02d:%02d", r.Email, sl.StartHr, sl.StartMin)
			w, err := newComWorker(label)
			if err != nil {
				return err
			}
			slotWorkers[slotKey{r.Email, sl.StartHr, sl.StartMin}] = w
		}
	}
	var err error
	cleanupWorker, err = newComWorker("cleanup")
	return err
}

// workerForSlot returns the comWorker that owns the apartment for the
// given (room, slot start) pair.
func workerForSlot(email string, start time.Time) *comWorker {
	return slotWorkers[slotKey{email, start.Hour(), start.Minute()}]
}

// ---------------------------------------------------------------------------
// Booking via local Outlook COM
// ---------------------------------------------------------------------------
//
// NOTE: Outlook's Programmatic Access ("Object Model Guard") may show a
// "A program is trying to send an email message on your behalf" prompt
// the first time. Accept it, or disable in:
//
//	File → Options → Trust Center → Trust Center Settings →
//	Programmatic Access → Never warn me about suspicious activity

// prepareDraft creates an AppointmentItem and populates every field
// (subject/body/location/start/duration/recipient) plus runs ResolveAll.
// The returned IDispatch is ready for Send() — at the trigger this is
// the only call we make, so the trigger-window latency drops to roughly
// one Send round-trip per room. The caller owns the returned handle and
// must arrange for it to be Released (sendDraft does so on success).
func prepareDraft(w *comWorker, room Room, start, end time.Time) (*ole.IDispatch, error) {
	var appt *ole.IDispatch
	var rerr error
	w.do(func() {
		// CreateItem is the first COM call that hits Outlook from this
		// apartment. When 8 worker apartments all enter prep at the
		// same moment, most of these calls collide on Outlook's STA
		// and bounce with RPC_E_CALL_REJECTED. The error clears within
		// tens of ms, so retry on transient with backoff — same shape
		// as sendDraft's retry loop.
		backoffs := []time.Duration{
			30 * time.Millisecond,
			80 * time.Millisecond,
			160 * time.Millisecond,
			320 * time.Millisecond,
		}
		var apptVar *ole.VARIANT
		var err error
		reattached := false
		for attempt := 0; ; attempt++ {
			apptVar, err = oleutil.CallMethod(w.app, "CreateItem", olAppointmentItem)
			if err == nil {
				break
			}
			// Outlook process died — reattach once (still on this
			// worker's pinned thread / apartment) and retry. Drafts
			// previously created via the dead handle remain dead;
			// only future creates benefit. The send-time loop
			// already treats dead drafts as failed attempts and
			// moves on, so partial recovery is still useful.
			if isOutlookGoneError(err) && !reattached {
				reattached = true
				fmt.Printf("[recover %s] Outlook unreachable, reattaching...\n", w.label)
				if rErr := w.reattachOutlookLocked(); rErr != nil {
					rerr = fmt.Errorf("CreateItem failed (%v) and reattach failed: %w", err, rErr)
					return
				}
				fmt.Printf("[recover %s] reattached, retrying CreateItem\n", w.label)
				continue
			}
			if !isTransientCOMError(err) || attempt >= len(backoffs) {
				rerr = fmt.Errorf("CreateItem: %w", err)
				return
			}
			time.Sleep(backoffs[attempt])
		}
		a := apptVar.ToIDispatch()
		// no defer Release — caller takes ownership

		for _, set := range []struct {
			name  string
			value any
		}{
			{"MeetingStatus", olMeeting},
			{"Subject", meetingSubject},
			{"Body", meetingBody},
			{"Location", room.Name},
			{"Start", oleDate(start)},
			{"Duration", int(end.Sub(start).Minutes())},
		} {
			if _, e := oleutil.PutProperty(a, set.name, set.value); e != nil {
				a.Release()
				rerr = fmt.Errorf("set %s: %w", set.name, e)
				return
			}
		}

		recipsVar, err := oleutil.GetProperty(a, "Recipients")
		if err != nil {
			a.Release()
			rerr = fmt.Errorf("get Recipients: %w", err)
			return
		}
		recips := recipsVar.ToIDispatch()
		defer recips.Release()

		recipVar, err := oleutil.CallMethod(recips, "Add", room.Email)
		if err != nil {
			a.Release()
			rerr = fmt.Errorf("Recipients.Add: %w", err)
			return
		}
		recip := recipVar.ToIDispatch()
		defer recip.Release()

		if _, e := oleutil.PutProperty(recip, "Type", olResource); e != nil {
			a.Release()
			rerr = fmt.Errorf("set Recipient.Type: %w", e)
			return
		}

		resolved, err := oleutil.CallMethod(recips, "ResolveAll")
		if err != nil {
			a.Release()
			rerr = fmt.Errorf("ResolveAll: %w", err)
			return
		}
		if resolved.Val == 0 {
			a.Release()
			rerr = fmt.Errorf("could not resolve %s against the directory", room.Email)
			return
		}

		appt = a
	})
	return appt, rerr
}

// sendDraft calls Send() on a previously prepared draft. Does NOT
// release the IDispatch — caller keeps it alive so it can poll the
// room's response via getRoomResponse, and releases it later via
// releaseDraft. tickAt is the trigger time so Total reflects
// end-to-end delay.
//
// Outlook sometimes returns transient COM errors when many apartments
// hit Send() simultaneously ("call was rejected by callee", "text
// formatting command is not available"). These un-contend in tens of
// milliseconds, so we retry the same Send a few times with a short
// backoff before giving up. Microsoft's COM guidance recommends this
// shape for RPC_E_CALL_REJECTED; the WordEditor "text formatting"
// error behaves identically in practice.
func sendDraft(w *comWorker, appt *ole.IDispatch, tickAt time.Time) (m bookMetrics, err error) {
	submitAt := time.Now()
	w.do(func() {
		m.Wait = time.Since(submitAt)
		workStart := time.Now()
		defer func() { m.Outlook = time.Since(workStart) }()

		backoffs := []time.Duration{
			30 * time.Millisecond,
			80 * time.Millisecond,
			160 * time.Millisecond,
			320 * time.Millisecond,
		}
		var lastErr error
		for attempt := 0; ; attempt++ {
			_, e := oleutil.CallMethod(appt, "Send")
			if e == nil {
				return
			}
			lastErr = e
			if !isTransientCOMError(e) || attempt >= len(backoffs) {
				break
			}
			time.Sleep(backoffs[attempt])
		}
		err = fmt.Errorf("Send: %w", lastErr)
	})
	m.Total = time.Since(tickAt)
	return
}

// isTransientCOMError matches the COM errors that come from Outlook
// momentarily losing a shared internal (WordEditor) or refusing a
// re-entrant call. They consistently clear within tens of ms, so a
// brief retry succeeds where a single attempt fails.
func isTransientCOMError(err error) bool {
	if err == nil {
		return false
	}
	msg := strings.ToLower(err.Error())
	signals := []string{
		"text formatting command",     // WordEditor contention
		"call was rejected by callee", // RPC_E_CALL_REJECTED
		"server is too busy",          // RPC_E_SERVERCALL_RETRYLATER
	}
	for _, s := range signals {
		if strings.Contains(msg, s) {
			return true
		}
	}
	if oleErr, ok := err.(*ole.OleError); ok {
		switch oleErr.Code() {
		case 0x80010001, // RPC_E_CALL_REJECTED
			0x8001010A: // RPC_E_SERVERCALL_RETRYLATER
			return true
		}
	}
	return false
}

// isOutlookGoneError matches the COM errors that mean Outlook.exe is no
// longer reachable from our process — the cached Application proxy
// references a dead RPC channel. Reattaching via CreateObject launches
// (or attaches to) a fresh Outlook and produces a usable handle.
func isOutlookGoneError(err error) bool {
	if err == nil {
		return false
	}
	if strings.Contains(strings.ToLower(err.Error()), "rpc server is unavailable") {
		return true
	}
	if oleErr, ok := err.(*ole.OleError); ok {
		switch oleErr.Code() {
		case 0x800706BA, // RPC_S_SERVER_UNAVAILABLE
			0x800706BE, // RPC_S_CALL_FAILED
			0x800706BF, // RPC_S_CALL_FAILED_DNE
			0x80080005: // CO_E_SERVER_EXEC_FAILURE
			return true
		}
	}
	return false
}

// getRoomResponse reads the current resource-recipient response state
// off a sent appointment. The IDispatch reflects whatever updates
// Outlook has received so far, so calling this in a loop after Send
// reveals the room's reply as soon as it lands.
func getRoomResponse(w *comWorker, appt *ole.IDispatch) roomState {
	state := stateNoRoom
	w.do(func() {
		recipsVar, err := oleutil.GetProperty(appt, "Recipients")
		if err != nil {
			return
		}
		recips := recipsVar.ToIDispatch()
		defer recips.Release()

		countVar, err := oleutil.GetProperty(recips, "Count")
		if err != nil {
			return
		}
		n := int(countVar.Val)
		for i := 1; i <= n; i++ {
			rVar, err := oleutil.CallMethod(recips, "Item", i)
			if err != nil {
				continue
			}
			r := rVar.ToIDispatch()
			typeVar, _ := oleutil.GetProperty(r, "Type")
			if int(typeVar.Val) == olResource {
				statusVar, _ := oleutil.GetProperty(r, "MeetingResponseStatus")
				switch int(statusVar.Val) {
				case olResponseAccepted:
					state = stateAccepted
				case olResponseDeclined:
					state = stateDeclined
				default:
					state = statePending
				}
				r.Release()
				// Only one room per draft — no need to inspect further.
				return
			}
			r.Release()
		}
	})
	return state
}

// waitForResponse polls the room's response every pollEvery until it
// sees a definitive Accepted or Declined, or `max` elapses. Logs the
// first observed state and any subsequent transitions, prefixed with
// `label`, so it's visible whether the response arrives in time and
// what it was.
func waitForResponse(w *comWorker, appt *ole.IDispatch, max time.Duration, label string) roomState {
	deadline := time.Now().Add(max)
	state := statePending
	last := roomState(-1) // sentinel — anything seen first will log
	for time.Now().Before(deadline) {
		time.Sleep(pollEvery)
		state = getRoomResponse(w, appt)
		if state != last {
			fmt.Printf("[poll] %s -> %s\n", label, state)
			last = state
		}
		if state == stateAccepted || state == stateDeclined {
			return state
		}
	}
	return state
}

// releaseDraft drops our COM reference. For sent items the underlying
// calendar entry survives (Outlook holds its own ref); for unsent
// drafts (attempts we skipped because we already won) the in-memory
// item is freed.
func releaseDraft(w *comWorker, appt *ole.IDispatch) {
	w.do(func() { appt.Release() })
}

// oleDate converts a Go time.Time to the OLE Automation date format that
// Outlook's AppointmentItem.Start expects: float64 days since 1899-12-30
// at midnight in the system's local timezone.
func oleDate(t time.Time) float64 {
	epoch := time.Date(1899, 12, 30, 0, 0, 0, 0, time.Local)
	return float64(t.Sub(epoch)) / float64(24*time.Hour)
}

// preparedDrafts is populated by prepareAll at the prep window and
// drained by sendAll at the trigger. Guarded because the two run on
// separate goroutines kicked off from the scheduler tick loop.
var (
	preparedDraftsMu sync.Mutex
	preparedDrafts   []preparedDraft
)

func prepareAll(loc *time.Location, fireAt time.Time) {
	if len(rooms) == 0 {
		fmt.Println("[prep] no rooms configured")
		return
	}

	// time.Date normalizes day overflow, so Day()+14 across a month
	// boundary works without us doing calendar math.
	targetDay := time.Date(fireAt.Year(), fireAt.Month(), fireAt.Day()+bookingDaysAhead,
		0, 0, 0, 0, loc)

	total := len(rooms) * len(slots) * sendAttempts
	fmt.Printf("[prep] preparing %d draft(s) (%d rooms × %d slots × %d attempts) for %s (%s)\n",
		total, len(rooms), len(slots), sendAttempts,
		targetDay.Format("2006-01-02"), loc)

	// One goroutine per (room, slot) — each dispatches its sendAttempts
	// preps to its own worker in parallel with other (room, slot)
	// groups. Within a group attempts are still serial because the
	// worker is single-threaded, but across groups the workers run
	// concurrently.
	var (
		draftsMu sync.Mutex
		drafts   []preparedDraft
		wg       sync.WaitGroup
	)
	for _, sl := range slots {
		for _, r := range rooms {
			wg.Add(1)
			go func(r Room, sl meetingSlot) {
				defer wg.Done()
				start := time.Date(targetDay.Year(), targetDay.Month(), targetDay.Day(),
					sl.StartHr, sl.StartMin, 0, 0, loc)
				end := start.Add(sl.Duration)
				w := workerForSlot(r.Email, start)
				slotLabel := start.Format("15:04")
				for attempt := 1; attempt <= sendAttempts; attempt++ {
					t0 := time.Now()
					appt, err := prepareDraft(w, r, start, end)
					elapsed := time.Since(t0).Milliseconds()
					if err != nil {
						fmt.Printf("[prep fail %d/%d] %s @ %s: %v (%dms)\n",
							attempt, sendAttempts, r.Email, slotLabel, err, elapsed)
						continue
					}
					fmt.Printf("[prep ok %d/%d]   %s @ %s (%dms)\n",
						attempt, sendAttempts, r.Email, slotLabel, elapsed)
					draftsMu.Lock()
					drafts = append(drafts, preparedDraft{r, appt, attempt, start})
					draftsMu.Unlock()
				}
			}(r, sl)
		}
	}
	wg.Wait()

	preparedDraftsMu.Lock()
	preparedDrafts = drafts
	preparedDraftsMu.Unlock()

	fmt.Printf("[prep] %d/%d draft(s) ready\n", len(drafts), total)
}

// sendGroup holds all the prepared drafts for a single (room, slot)
// across attempts. runRetrySequence walks them in order, stopping
// early on acceptance.
type sendGroup struct {
	room     Room
	start    time.Time
	attempts []*ole.IDispatch // ordered by attempt number
}

type sendGroupKey struct {
	email string
	start time.Time
}

func sendAll(loc *time.Location, fireAt time.Time) {
	preparedDraftsMu.Lock()
	drafts := preparedDrafts
	preparedDrafts = nil
	preparedDraftsMu.Unlock()

	if len(drafts) == 0 {
		fmt.Println("[trigger] no prepared drafts (prep window missed?) — nothing to send")
		return
	}

	// Group drafts by (room, slot start) so each group's attempts can
	// be sent serially with a poll-and-retry loop on a single goroutine.
	groups := map[sendGroupKey]*sendGroup{}
	for _, d := range drafts {
		k := sendGroupKey{d.room.Email, d.start}
		if _, ok := groups[k]; !ok {
			groups[k] = &sendGroup{room: d.room, start: d.start}
		}
		groups[k].attempts = append(groups[k].attempts, d.appt)
	}

	// Build a canonical group order (slots × rooms, matching the
	// configured slices) so the first room/slot fires first regardless
	// of Go map iteration order. The stagger below applies in this
	// order, so the priority room is the one that gets the 0ms start.
	targetDay := time.Date(fireAt.Year(), fireAt.Month(), fireAt.Day()+bookingDaysAhead,
		0, 0, 0, 0, loc)
	var orderedGroups []*sendGroup
	for _, sl := range slots {
		start := time.Date(targetDay.Year(), targetDay.Month(), targetDay.Day(),
			sl.StartHr, sl.StartMin, 0, 0, loc)
		for _, r := range rooms {
			if g, ok := groups[sendGroupKey{r.Email, start}]; ok {
				orderedGroups = append(orderedGroups, g)
			}
		}
	}

	fmt.Printf("[trigger] sending across %d (room, slot) group(s), up to %d attempts each (stagger %s)\n",
		len(orderedGroups), sendAttempts, sendStagger)

	// One goroutine per (room, slot). Each waits index*sendStagger
	// before firing its first Send so only one apartment touches
	// Outlook's WordEditor at a time — that's what was failing with
	// "text formatting command is not available" when all 8 fired
	// at the same microsecond. Inside the goroutine, attempts are
	// sequential (send → poll → maybe retry) on the group's own
	// dedicated comWorker.
	var wg sync.WaitGroup
	for i, g := range orderedGroups {
		wg.Add(1)
		delay := time.Duration(i) * sendStagger
		go func(g *sendGroup, delay time.Duration) {
			defer wg.Done()
			if delay > 0 {
				time.Sleep(delay)
			}
			w := workerForSlot(g.room.Email, g.start)
			defer func() {
				for _, a := range g.attempts {
					releaseDraft(w, a)
				}
			}()
			runRetrySequence(w, g, loc, fireAt)
		}(g, delay)
	}
	wg.Wait()
}

// runRetrySequence sends each attempt for one (room, slot) in order,
// polling the room's response between attempts. Stops as soon as the
// room accepts; otherwise keeps trying until attempts are exhausted.
func runRetrySequence(w *comWorker, g *sendGroup, loc *time.Location, fireAt time.Time) {
	slotLabel := g.start.Format("15:04")
	for i, appt := range g.attempts {
		m, err := sendDraft(w, appt, fireAt)
		done := time.Now().In(loc).Format("15:04:05.000")
		tag := fmt.Sprintf("%d/%d", i+1, len(g.attempts))
		if err != nil {
			fmt.Printf("%s [fail %s] %s @ %s: %v (%s)\n",
				done, tag, g.room.Email, slotLabel, err, m)
			// Send never reached Exchange — polling would just sit at
			// "pending" for the full pollInterval. Skip the wait and
			// fire the next attempt immediately so we don't burn 30s
			// of race-time on an item that has no room response.
			if i == len(g.attempts)-1 {
				return
			}
			continue
		}
		fmt.Printf("%s [ok %s]   %s @ %s sent (%s)\n",
			done, tag, g.room.Email, slotLabel, m)

		// No point waiting/polling after the last attempt.
		if i == len(g.attempts)-1 {
			return
		}

		state := waitForResponse(w, appt, pollInterval,
			fmt.Sprintf("%s @ %s attempt %s", g.room.Email, slotLabel, tag))
		now := time.Now().In(loc).Format("15:04:05.000")
		switch state {
		case stateAccepted:
			fmt.Printf("%s [accepted] %s @ %s after attempt %s — skipping retries\n",
				now, g.room.Email, slotLabel, tag)
			return
		case stateDeclined:
			fmt.Printf("%s [declined] %s @ %s attempt %s — sending next attempt\n",
				now, g.room.Email, slotLabel, tag)
		default: // statePending or stateNoRoom
			fmt.Printf("%s [timeout] %s @ %s attempt %s no response within %s — sending next attempt\n",
				now, g.room.Email, slotLabel, tag, pollInterval)
		}
	}
}

// cleanupLeftovers is the startup-time entry point — sweeps any of our
// auto-booked meetings still stuck Declined or Pending from previous
// nights' runs. Scans from today through today + bookingDaysAhead + 1
// day (the widest window any prior run could have targeted that's
// still in the future). Past meetings are skipped because we can't
// unbook history. Accepts stay — we never delete them.
func cleanupLeftovers(loc *time.Location) {
	now := time.Now().In(loc)
	start := time.Date(now.Year(), now.Month(), now.Day(), 0, 0, 0, 0, loc)
	end := start.AddDate(0, 0, bookingDaysAhead+1)
	cleanupOurMeetings(loc, start, end)
}

// cleanupOurMeetings scans the calendar in [start, end) for meetings
// matching our auto-booking subject(s) and deletes those whose room
// recipient declined or hasn't replied yet. Runs at program start
// only — post-trigger cleanup is intentionally absent because room
// moderators can take many hours to approve, and an aggressive sweep
// at +15min nukes accepts-in-flight. Pending items from previous
// nights either get accepted (kept) or stay pending until the next
// startup, where we treat them as lost and delete them.
func cleanupOurMeetings(loc *time.Location, start, end time.Time) {
	fmt.Printf("[cleanup] scanning calendar %s to %s, removing declined or pending bookings\n",
		start.Format("2006-01-02"), end.Format("2006-01-02"))

	cleanupWorker.do(func() {
		// Reuse the cleanup worker's cached Outlook.Application IDispatch
		// rather than CreateObject + QueryInterface every cleanup pass.
		sessVar, err := oleutil.GetProperty(cleanupWorker.app, "Session")
		if err != nil {
			fmt.Printf("[cleanup] Session: %v\n", err)
			return
		}
		session := sessVar.ToIDispatch()
		defer session.Release()

		// Force a sync with Exchange before we scan. Without this, any
		// room-decline emails that Outlook hasn't yet pulled down leave
		// their appointments stuck in "Pending" — which our defensive
		// rule then keeps. SendAndReceive(false) runs the default
		// send/receive group with no UI; the short sleep gives the
		// background sync a moment to apply the responses to local
		// items before we read MeetingResponseStatus.
		fmt.Println("[cleanup] forcing Outlook send/receive to pull latest room responses")
		if _, e := oleutil.CallMethod(session, "SendAndReceive", false); e != nil {
			fmt.Printf("[cleanup] SendAndReceive: %v (proceeding anyway)\n", e)
		}
		time.Sleep(3 * time.Second)

		calVar, err := oleutil.CallMethod(session, "GetDefaultFolder", olFolderCalendar)
		if err != nil {
			fmt.Printf("[cleanup] GetDefaultFolder: %v\n", err)
			return
		}
		cal := calVar.ToIDispatch()
		defer cal.Release()

		itemsVar, err := oleutil.GetProperty(cal, "Items")
		if err != nil {
			fmt.Printf("[cleanup] Items: %v\n", err)
			return
		}
		items := itemsVar.ToIDispatch()
		defer items.Release()

		// Sort + IncludeRecurrences are recommended before Restrict on date ranges.
		oleutil.PutProperty(items, "Sort", "[Start]")
		oleutil.PutProperty(items, "IncludeRecurrences", true)

		// Restrict expects "M/d/yyyy h:mm AM/PM" in current-locale form.
		// We can no longer match Subject exactly because each booking
		// includes the room's display name suffix (e.g. "<owner> Room
		// booking - Room A - ..."). Filter by date here and
		// match the subject prefix manually below.
		filter := fmt.Sprintf("[Start] >= '%s' AND [Start] <= '%s'",
			start.Format("1/2/2006 3:04 PM"),
			end.Format("1/2/2006 3:04 PM"))
		filteredVar, err := oleutil.CallMethod(items, "Restrict", filter)
		if err != nil {
			fmt.Printf("[cleanup] Restrict: %v\n", err)
			return
		}
		filtered := filteredVar.ToIDispatch()
		defer filtered.Release()

		countVar, err := oleutil.GetProperty(filtered, "Count")
		if err != nil {
			fmt.Printf("[cleanup] Count: %v\n", err)
			return
		}
		count := int(countVar.Val)
		fmt.Printf("[cleanup] %d candidate meeting(s) in range\n", count)

		var kept, deleted int
		// Iterate backwards so deletions don't shift indices we haven't visited.
		for i := count; i >= 1; i-- {
			itemVar, err := oleutil.CallMethod(filtered, "Item", i)
			if err != nil {
				continue
			}
			item := itemVar.ToIDispatch()

			// Subject gate: skip anything that isn't one of our known
			// auto-booking subjects. Log what we skip (with the subject
			// quoted) so unexpected items can be diagnosed.
			subjectVar, err := oleutil.GetProperty(item, "Subject")
			if err != nil {
				fmt.Printf("[cleanup] skip   <unreadable subject>: %v\n", err)
				item.Release()
				continue
			}
			subject := subjectVar.ToString()
			if !isOurSubject(subject) {
				fmt.Printf("[cleanup] skip   %q (subject not in cleanup list)\n", subject)
				item.Release()
				continue
			}

			ex := examineMeeting(item)
			suffix := ""
			if ex.humanDeclined {
				suffix = " (human attendee also declined)"
			}

			// Delete on Declined OR Pending. At startup, "Pending"
			// means a prior run created the request and the room
			// hasn't approved (or has but the response is lost).
			// Either way the meeting isn't going to happen — clear it.
			if ex.roomState == stateDeclined || ex.roomState == statePending {
				fmt.Printf("[cleanup] delete %s — %s%s\n", ex.roomEmail, ex.roomState, suffix)
				if _, e := oleutil.CallMethod(item, "Delete"); e != nil {
					fmt.Printf("[cleanup]        delete failed: %v\n", e)
				} else {
					deleted++
				}
			} else {
				fmt.Printf("[cleanup] keep   %s — %s%s\n", ex.roomEmail, ex.roomState, suffix)
				kept++
			}
			item.Release()
		}

		fmt.Printf("[cleanup] pass-1 (calendar scan) — kept %d, deleted %d\n", kept, deleted)

		// Pass 2 — walk the "Room booking" folder for decline responses
		// whose associated appointments are still stuck Pending in the
		// local calendar (Outlook's auto-process can fail to apply the
		// response when a rule routes the email before processing).
		// Each Declined response's GetAssociatedAppointment gives us
		// the appointment to delete directly.
		if responseFolderName == "" {
			return
		}
		extra := sweepResponseFolder(session, start, end)
		fmt.Printf("[cleanup] pass-2 (response folder) — deleted %d additional stuck-pending appointment(s)\n", extra)
	})
}

// sweepResponseFolder walks the user's mailbox to find decline
// responses whose associated appointment is stuck "Pending" in the
// local calendar (Outlook's auto-process can fail to apply them, and
// the user's mailbox rule may not match responses whose subject is
// just "Declined: " with nothing after — those stay in Inbox).
//
// We scan both the configured response folder (if it has items) AND
// the Inbox itself, since responses with empty subjects don't match
// typical "Room booking"-keyword rules and stay in Inbox.
//
// Runs inside the cleanup worker's apartment.
func sweepResponseFolder(session *ole.IDispatch, start, end time.Time) int {
	inboxVar, err := oleutil.CallMethod(session, "GetDefaultFolder", olFolderInbox)
	if err != nil {
		fmt.Printf("[cleanup] GetDefaultFolder(Inbox): %v\n", err)
		return 0
	}
	inbox := inboxVar.ToIDispatch()
	defer inbox.Release()

	rootVar, err := oleutil.GetProperty(inbox, "Parent")
	if err != nil {
		fmt.Printf("[cleanup] Inbox.Parent: %v\n", err)
		return 0
	}
	root := rootVar.ToIDispatch()
	defer root.Release()

	deleted := 0

	// Scan configured response folder (if it exists and has items).
	if folder := findResponseFolder(root, responseFolderName); folder != nil {
		deleted += scanFolderForDeclines(folder, start, end, responseFolderName)
		folder.Release()
	} else {
		fmt.Printf("[cleanup] no non-empty %q folder found, will rely on Inbox scan\n", responseFolderName)
	}

	// Always also scan Inbox — many room declines have empty subjects
	// ("Declined: " with nothing after) and never match a keyword rule,
	// so they sit in Inbox unprocessed.
	deleted += scanFolderForDeclines(inbox, start, end, "Inbox")

	return deleted
}

// scanFolderForDeclines iterates a folder's items looking for any
// "Declined:" response email, calls GetAssociatedAppointment on each,
// and deletes the appointment if it's on the target day AND its
// subject matches one of ours. This is the cross-validation step:
// the email subject alone is unreliable (many rooms send "Declined: "
// with nothing after), so we identify ours via the associated
// appointment instead. label is just for log readability.
func scanFolderForDeclines(folder *ole.IDispatch, start, end time.Time, label string) int {
	itemsVar, err := oleutil.GetProperty(folder, "Items")
	if err != nil {
		return 0
	}
	items := itemsVar.ToIDispatch()
	defer items.Release()

	countVar, err := oleutil.GetProperty(items, "Count")
	if err != nil {
		return 0
	}
	n := int(countVar.Val)
	fmt.Printf("[cleanup] scanning %d item(s) in %s\n", n, label)

	deleted := 0

	for i := n; i >= 1; i-- {
		itemVar, err := oleutil.CallMethod(items, "Item", i)
		if err != nil {
			continue
		}
		item := itemVar.ToIDispatch()
		// Folders can contain item types whose VARIANT isn't a usable
		// IDispatch (post items, report items, etc.). Skip those
		// defensively — calling COM on a nil dispatch panics.
		if item == nil {
			continue
		}

		subjVar, err := oleutil.GetProperty(item, "Subject")
		if err != nil {
			item.Release()
			continue
		}
		// Match any "Declined:" subject — many rooms send these with
		// nothing after the colon, so we can't rely on a more specific
		// prefix. Instead, identify ours via the associated appointment.
		if !strings.HasPrefix(strings.TrimSpace(subjVar.ToString()), "Declined:") {
			item.Release()
			continue
		}

		apptVar, err := oleutil.CallMethod(item, "GetAssociatedAppointment", false)
		if err != nil {
			// Appointment already gone (likely deleted in pass-1) — fine.
			item.Release()
			continue
		}
		appt := apptVar.ToIDispatch()
		if appt == nil {
			item.Release()
			continue
		}

		// Confirm this decline is for OUR booking (subject matches).
		apptSubjVar, err := oleutil.GetProperty(appt, "Subject")
		if err != nil || !isOurSubject(apptSubjVar.ToString()) {
			appt.Release()
			item.Release()
			continue
		}

		// Date guard — only delete if the appointment falls inside
		// the configured [start, end) window. Prevents touching
		// older / unrelated declines.
		startVar, err := oleutil.GetProperty(appt, "Start")
		if err == nil {
			if apptStart, ok := startVar.Value().(time.Time); ok {
				apptStart = apptStart.In(start.Location())
				if apptStart.Before(start) || !apptStart.Before(end) {
					appt.Release()
					item.Release()
					continue
				}
			}
		}

		if _, e := oleutil.CallMethod(appt, "Delete"); e != nil {
			fmt.Printf("[cleanup] pass-2 delete failed: %v\n", e)
		} else {
			deleted++
		}
		appt.Release()
		item.Release()
	}
	return deleted
}

// findResponseFolder walks the folder tree under `parent` searching
// for a folder whose name matches `target` (case-insensitive) and
// which has at least one item in it. Skips empty matches — multiple
// folders with the same name can exist (e.g., an old empty subfolder
// under Inbox plus the real one elsewhere); we want the populated
// one. Returns nil if no non-empty match is found. Caller must
// Release the returned folder.
func findResponseFolder(parent *ole.IDispatch, target string) *ole.IDispatch {
	foldersVar, err := oleutil.GetProperty(parent, "Folders")
	if err != nil {
		return nil
	}
	folders := foldersVar.ToIDispatch()
	defer folders.Release()

	countVar, err := oleutil.GetProperty(folders, "Count")
	if err != nil {
		return nil
	}
	n := int(countVar.Val)

	for i := 1; i <= n; i++ {
		fVar, err := oleutil.CallMethod(folders, "Item", i)
		if err != nil {
			continue
		}
		f := fVar.ToIDispatch()

		nameVar, _ := oleutil.GetProperty(f, "Name")
		if strings.EqualFold(strings.TrimSpace(nameVar.ToString()), target) {
			// Found a candidate — confirm it has items before using it.
			itemsVar, err := oleutil.GetProperty(f, "Items")
			if err == nil {
				items := itemsVar.ToIDispatch()
				icountVar, err := oleutil.GetProperty(items, "Count")
				items.Release()
				if err == nil && int(icountVar.Val) > 0 {
					pathVar, _ := oleutil.GetProperty(f, "FolderPath")
					fmt.Printf("[cleanup] found %q at %s (%d items)\n",
						target, pathVar.ToString(), int(icountVar.Val))
					return f
				}
			}
			// Empty match — keep searching for a non-empty one.
		}

		// Recurse into subfolders.
		if found := findResponseFolder(f, target); found != nil {
			f.Release()
			return found
		}
		f.Release()
	}
	return nil
}

// roomState classifies the resource recipient's response on a meeting.
// "Pending" lumps None / Tentative / NotResponded — anything that isn't
// a definitive accept or decline.
type roomState int

const (
	stateNoRoom   roomState = iota // no recipient with Type=olResource
	stateAccepted                  // room responded accept
	stateDeclined                  // room responded decline
	statePending                   // room hasn't given a definitive response yet
)

func (s roomState) String() string {
	switch s {
	case stateNoRoom:
		return "no room recipient"
	case stateAccepted:
		return "room accepted"
	case stateDeclined:
		return "room declined"
	case statePending:
		return "room pending"
	}
	return "unknown"
}

// meetingExamination summarizes what the cleanup needs to know about a
// candidate meeting:
//
//	roomEmail      — the resource recipient's address (empty if none)
//	roomState      — what the room replied (or stateNoRoom)
//	humanDeclined  — informational; ANY human required/optional declined
type meetingExamination struct {
	roomEmail     string
	roomState     roomState
	humanDeclined bool
}

// examineMeeting walks appt's Recipients collection and reports the
// state of the resource recipient plus whether any human declined.
func examineMeeting(appt *ole.IDispatch) meetingExamination {
	out := meetingExamination{roomState: stateNoRoom}

	recipsVar, err := oleutil.GetProperty(appt, "Recipients")
	if err != nil {
		return out
	}
	recips := recipsVar.ToIDispatch()
	defer recips.Release()

	countVar, err := oleutil.GetProperty(recips, "Count")
	if err != nil {
		return out
	}
	n := int(countVar.Val)

	for i := 1; i <= n; i++ {
		rVar, err := oleutil.CallMethod(recips, "Item", i)
		if err != nil {
			continue
		}
		r := rVar.ToIDispatch()

		typeVar, _ := oleutil.GetProperty(r, "Type")
		statusVar, _ := oleutil.GetProperty(r, "MeetingResponseStatus")
		recipType := int(typeVar.Val)
		status := int(statusVar.Val)

		switch recipType {
		case olResource:
			addressVar, _ := oleutil.GetProperty(r, "Address")
			out.roomEmail = addressVar.ToString()
			switch status {
			case olResponseAccepted:
				out.roomState = stateAccepted
			case olResponseDeclined:
				out.roomState = stateDeclined
			default:
				out.roomState = statePending
			}
		case olRequired, olOptional:
			if status == olResponseDeclined {
				out.humanDeclined = true
			}
		}
		r.Release()
	}
	return out
}

// ---------------------------------------------------------------------------
// Scheduler
// ---------------------------------------------------------------------------

// teeStdoutToFile redirects everything written to os.Stdout to BOTH the
// original console and a date-stamped log file in the working directory
// (booking-YYYY-MM-DD.log, append-mode so re-runs on the same day keep
// accumulating). Returns a cleanup func the caller must defer — it
// flushes the pipe, restores os.Stdout, and closes the file. On error
// we fall back to console-only output and the cleanup is a no-op.
func teeStdoutToFile() func() {
	tz, err := time.LoadLocation(localTZName)
	if err != nil {
		tz = time.Local
	}
	name := fmt.Sprintf("booking-%s.log", time.Now().In(tz).Format("2006-01-02"))
	f, err := os.OpenFile(name, os.O_APPEND|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		fmt.Fprintf(os.Stderr, "[init] log file %s open failed: %v — console only\n", name, err)
		return func() {}
	}
	r, w, err := os.Pipe()
	if err != nil {
		fmt.Fprintf(os.Stderr, "[init] log pipe failed: %v — console only\n", err)
		f.Close()
		return func() {}
	}
	realStdout := os.Stdout
	os.Stdout = w

	done := make(chan struct{})
	go func() {
		defer close(done)
		// MultiWriter copies every byte to both the original console
		// and the log file. io.Copy returns when the write end of the
		// pipe is closed (i.e. on shutdown below).
		_, _ = io.Copy(io.MultiWriter(realStdout, f), r)
	}()

	fmt.Printf("[init] logging to %s\n", name)

	return func() {
		os.Stdout = realStdout
		w.Close()
		<-done
		f.Close()
	}
}

func main() {
	defer teeStdoutToFile()()

	logSkew := func(label string) *clockSkewSample {
		sample, err := measureClockSkew()
		if err != nil {
			fmt.Printf("[debug] clock skew (%s) check failed: %v\n", label, err)
			return nil
		}
		direction := "ahead of"
		mag := sample.Offset
		if sample.Offset < 0 {
			direction = "behind"
			mag = -sample.Offset
		}
		fmt.Printf("[debug] clock skew (%s): laptop is %s %s network (rtt %s)\n",
			label, mag.Truncate(time.Millisecond), direction,
			sample.RTT.Truncate(time.Millisecond))
		return &sample
	}

	logSkew("before sync")
	syncSystemClock()
	if after := logSkew("after sync"); after != nil {
		mag := after.Offset
		if mag < 0 {
			mag = -mag
		}
		// 50ms is roughly the per-Send variance — drift above this
		// starts shifting the trigger fire away from when Exchange
		// thinks the slot opens, so it's worth surfacing.
		if mag > 50*time.Millisecond {
			// Negate the offset: positive Offset = laptop ahead → step
			// backward; negative = laptop behind → step forward.
			adjustMs := -after.Offset.Milliseconds()
			fmt.Printf("[debug] residual drift still %s — auto-sync did not fully correct\n",
				mag.Truncate(time.Millisecond))
			fmt.Println("[debug] paste in an elevated PowerShell to step the clock:")
			fmt.Printf("[debug]   Set-Date -Adjust ([TimeSpan]::FromMilliseconds(%d))\n", adjustMs)
		}
	}

	killExistingOutlook()

	startup := time.Now()
	workerCount := len(rooms)*len(slots) + 1
	fmt.Printf("starting %d COM workers (one per room/slot + cleanup)...\n", workerCount)
	if err := startWorkers(); err != nil {
		fmt.Printf("worker start failed: %v\n", err)
		return
	}
	fmt.Printf("Outlook ready (%d workers) in %dms\n", workerCount, time.Since(startup).Milliseconds())

	loc, err := time.LoadLocation(localTZName)
	if err != nil {
		fmt.Printf("LoadLocation(%q) failed: %v — falling back to fixed UTC+7\n", localTZName, err)
		loc = time.FixedZone("GMT+7", 7*60*60)
	}

	// Startup-time cleanup of leftovers from previous nights' runs.
	// Post-trigger cleanup is intentionally absent — room moderators
	// can take many hours to approve, so a sweep right after sending
	// would delete accepts-in-flight. Anything still pending tomorrow
	// will be cleared by tomorrow's startup pass.
	cleanupLeftovers(loc)

	// Align the first tick close to a wall-clock second boundary so the
	// detection lag at the trigger is small (single-digit ms instead of
	// up to ~1s of jitter from process start time).
	now := time.Now()
	nextSec := now.Truncate(time.Second).Add(time.Second)
	time.Sleep(time.Until(nextSec))

	ticker := time.NewTicker(1 * time.Second)
	defer ticker.Stop()

	fmt.Printf("scheduler started — local zone %s, polling every 1s\n", loc)
	fmt.Printf("booking trigger at %02d:%02d local (drafts prepared %s before)\n",
		triggerHr, triggerMin, prepLead)

	// One-fire-per-trigger guards. lastPreparedKey is keyed by the
	// upcoming trigger date (which can be tomorrow if prep is on the
	// previous calendar day, e.g. trigger=00:00 + prepLead=30s).
	// lastFiredDate is keyed by the actual day the trigger fires on.
	var (
		lastPreparedKey string
		lastFiredDate   string
	)

	for tick := range ticker.C {
		local := tick.In(loc)
		date := local.Format("2006-01-02")

		// Compute the next trigger fire (today's, or tomorrow's if
		// today's has already passed).
		nextTrigger := time.Date(local.Year(), local.Month(), local.Day(),
			triggerHr, triggerMin, 0, 0, loc)
		if !nextTrigger.After(local) {
			nextTrigger = nextTrigger.AddDate(0, 0, 1)
		}
		nextPrep := nextTrigger.Add(-prepLead)
		nextTriggerKey := nextTrigger.Format("2006-01-02")

		// Prep window: between [nextPrep, nextTrigger), once per upcoming trigger.
		if !local.Before(nextPrep) && local.Before(nextTrigger) && nextTriggerKey != lastPreparedKey {
			lastPreparedKey = nextTriggerKey
			fmt.Printf("[%s] prep window opened — preparing for trigger at %s\n",
				local.Format("15:04:05"), nextTrigger.Format("15:04:05"))
			go prepareAll(loc, nextTrigger)
		}

		// Trigger window: hh:mm equals trigger, once per calendar day.
		if local.Hour() == triggerHr && local.Minute() == triggerMin && date != lastFiredDate {
			lastFiredDate = date
			fmt.Printf("trigger reached at %s — sending\n",
				local.Format("2006-01-02 15:04:05.000 MST"))
			fireAt := local
			go sendAll(loc, fireAt)
		}
	}
}
