/*
AutoHotkey

Copyright 2003-2009 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#include "stdafx.h" // pre-compiled headers
#include "application.h"
#include "globaldata.h" // for access to g_clip, the "g" global struct, etc.
#include "window.h" // for several MsgBox and window functions
#include "util.h" // for strlcpy()
#include "resources/resource.h"  // For ID_TRAY_OPEN.


bool MsgSleep(int aSleepDuration, MessageMode aMode)
// Returns true if it launched at least one thread, and false otherwise.
// aSleepDuration can be be zero to do a true Sleep(0), or less than 0 to avoid sleeping or
// waiting at all (i.e. messages are checked and if there are none, the function will return
// immediately).  aMode is either RETURN_AFTER_MESSAGES (default) or WAIT_FOR_MESSAGES.
// If the caller doesn't specify aSleepDuration, this function will return after a
// time less than or equal to SLEEP_INTERVAL (i.e. the exact amount of the sleep
// isn't important to the caller).  This mode is provided for performance reasons
// (it avoids calls to GetTickCount and the TickCount math).  However, if the
// caller's script subroutine is suspended due to action by us, an unknowable
// amount of time may pass prior to finally returning to the caller.
{
	bool we_turned_on_defer = false; // Set default.
	if (aMode == RETURN_AFTER_MESSAGES_SPECIAL_FILTER)
	{
		aMode = RETURN_AFTER_MESSAGES; // To simplify things further below, eliminate the mode RETURN_AFTER_MESSAGES_SPECIAL_FILTER from further consideration.
		// g_DeferMessagesForUnderlyingPump is a global because the instance of MsgSleep on the call stack
		// that set it to true could launch new thread(s) that call MsgSleep again (i.e. a new layer), and a global
		// is the easiest way to inform all such MsgSleeps that there's a non-standard msg pump beneath them on the
		// call stack.
		if (!g_DeferMessagesForUnderlyingPump)
		{
			g_DeferMessagesForUnderlyingPump = true;
			we_turned_on_defer = true;
		}
		// So now either we turned it on or some layer beneath us did.  Therefore, we know there's at least one
		// non-standard msg pump beneath us on the call stack.
	}

	// While in mode RETURN_AFTER_MESSAGES, there are different things that can happen:
	// 1) We launch a new hotkey subroutine, interrupting/suspending the old one.  But
	//    subroutine calls this function again, so now it's recursed.  And thus the
	//    new subroutine can be interrupted yet again.
	// 2) We launch a new hotkey subroutine, but it returns before any recursed call
	//    to this function discovers yet another hotkey waiting in the queue.  In this
	//    case, this instance/recursion layer of the function should process the
	//    hotkey messages linearly rather than recursively?  No, this doesn't seem
	//    necessary, because we can just return from our instance/layer and let the
	//    caller handle any messages waiting in the queue.  Eventually, the queue
	//    should be emptied, especially since most hotkey subroutines will run
	//    much faster than the user could press another hotkey, with the possible
	//    exception of the key-repeat feature triggered by holding a key down.
	//    Even in that case, the worst that would happen is that messages would
	//    get dropped off the queue because they're too old (I think that's what
	//    happens).
	// Based on the above, when mode is RETURN_AFTER_MESSAGES, we process
	// all messages until a hotkey message is encountered, at which time we
	// launch that subroutine only and then return when it returns to us, letting
	// the caller handle any additional messages waiting on the queue.  This avoids
	// the need to have a "run the hotkeys linearly" mode in a single iteration/layer
	// of this function.  Note: The WM_QUIT message does not receive any higher
	// precedence in the queue than other messages.  Thus, if there's ever concern
	// that that message would be lost, as a future change perhaps can use PeekMessage()
	// with a filter to explicitly check to see if our queue has a WM_QUIT in it
	// somewhere, prior to processing any messages that might take result in
	// a long delay before the remainder of the queue items are processed (there probably
	// aren't any such conditions now, so nothing to worry about?)

	// Above is somewhat out-of-date.  The objective now is to spend as much time
	// inside GetMessage() as possible, since it's the keystroke/mouse engine
	// whenever the hooks are installed.  Any time we're not in GetMessage() for
	// any length of time (say, more than 20ms), keystrokes and mouse events
	// will be lagged.  PeekMessage() is probably almost as good, but it probably
	// only clears out any waiting keys prior to returning.  CONFIRMED: PeekMessage()
	// definitely routes to the hook, perhaps only if called regularly (i.e. a single
	// isolated call might not help much).

	// Decided to support a true Sleep(0) for aSleepDuration == 0, as well
	// as no delay at all if aSleepDuration < 0.  This is needed to implement
	// "SetKeyDelay, 0" and possibly other things.  I believe a Sleep(0)
	// is always <= Sleep(1) because both of these will wind up waiting
	// a full timeslice if the CPU is busy.

	// Reminder for anyone maintaining or revising this code:
	// Giving each subroutine its own thread rather than suspending old ones is
	// probably not a good idea due to the exclusive nature of the GUI
	// (i.e. it's probably better to suspend existing subroutines rather than
	// letting them continue to run because they might activate windows and do
	// other stuff that would interfere with the window automation activities of
	// other threads)

	// If caller didn't specify, the exact amount of the Sleep() isn't
	// critical to it, only that we handles messages and do Sleep()
	// a little.
	// Most of this initialization section isn't needed if aMode == WAIT_FOR_MESSAGES,
	// but it's done anyway for consistency:
	bool allow_early_return;
	if (aSleepDuration == INTERVAL_UNSPECIFIED)
	{
		aSleepDuration = SLEEP_INTERVAL;  // Set interval to be the default length.
		allow_early_return = true;
	}
	else
		// The timer resolution makes waiting for half or less of an
		// interval too chancy.  The correct thing to do on average
		// is some kind of rounding, which this helps with:
		allow_early_return = (aSleepDuration <= SLEEP_INTERVAL_HALF);

	// Record the start time when the caller first called us so we can keep
	// track of how much time remains to sleep (in case the caller's subroutine
	// is suspended until a new subroutine is finished).  But for small sleep
	// intervals, don't worry about it.
	// Note: QueryPerformanceCounter() has very high overhead compared to GetTickCount():
	DWORD start_time = allow_early_return ? 0 : GetTickCount();

	// This check is also done even if the main timer will be set (below) so that
	// an initial check is done rather than waiting 10ms more for the first timer
	// message to come in.  Some of our many callers would want this, and although some
	// would not need it, there are so many callers that it seems best to just do it
	// unconditionally, especially since it's not a high overhead call (e.g. it returns
	// immediately if the tickcount is still the same as when it was last run).
	// Another reason for doing this check immediately is that our msg queue might
	// contains a time-consuming msg prior to our WM_TIMER msg, e.g. a hotkey msg.
	// In that case, the hotkey would be processed and launched without us first having
	// emptied the queue to discover the WM_TIMER msg.  In other words, WM_TIMER msgs
	// might get buried in the queue behind others, so doing this check here should help
	// ensure that timed subroutines are checked often enough to keep them running at
	// their specified frequencies.
	// Note that ExecUntil() no longer needs to call us solely for prevention of lag
	// caused by the keyboard & mouse hooks, so checking the timers early, rather than
	// immediately going into the GetMessage() state, should not be a problem:
	POLL_JOYSTICK_IF_NEEDED  // Do this first since it's much faster.
	bool return_value = false; //  Set default.  Also, this is used by the macro below.
	CHECK_SCRIPT_TIMERS_IF_NEEDED

	// Because this function is called recursively: for now, no attempt is
	// made to improve performance by setting the timer interval to be
	// aSleepDuration rather than a standard short interval.  That would cause
	// a problem if this instance of the function invoked a new subroutine,
	// suspending the one that called this instance.  The new subroutine
	// might need a timer of a longer interval, which would mess up
	// this layer.  One solution worth investigating is to give every
	// layer/instance its own timer (the ID of the timer can be determined
	// from info in the WM_TIMER message).  But that can be a real mess
	// because what if a deeper recursion level receives our first
	// WM_TIMER message because we were suspended too long?  Perhaps in
	// that case we wouldn't our WM_TIMER pulse because upon returning
	// from those deeper layers, we would check to see if the current
	// time is beyond our finish time.  In addition, having more timers
	// might be worse for overall system performance than having a single
	// timer that pulses very frequently (because the system must keep
	// them all up-to-date).  UPDATE: Timer is now also needed whenever an
	// aSleepDuration greater than 0 is about to be done and there are some
	// script timers that need to be watched (this happens when aMode == WAIT_FOR_MESSAGES).
	// UPDATE: Make this a macro so that it is dynamically resolved every time, in case
	// the value of g_script.mTimerEnabledCount changes on-the-fly.
	// UPDATE #2: The below has been changed in light of the fact that the main timer is
	// now kept always-on whenever there is at least one enabled timed subroutine.
	// This policy simplifies ExecUntil() and long-running commands such as FileSetAttrib.
	// UPDATE #3: Use aMode == RETURN_AFTER_MESSAGES, not g_nThreads > 0, because the
	// "Edit This Script" menu item (and possibly other places) might result in an indirect
	// call to us and we will need the timer to avoid getting stuck in the GetMessageState()
	// with hotkeys being disallowed due to filtering:
	bool this_layer_needs_timer = (aSleepDuration > 0 && aMode == RETURN_AFTER_MESSAGES);
	if (this_layer_needs_timer)
	{
		++g_nLayersNeedingTimer;  // IsCycleComplete() is responsible for decrementing this for us.
		SET_MAIN_TIMER
		// Reasons why the timer might already have been on:
		// 1) g_script.mTimerEnabledCount is greater than zero or there are joystick hotkeys.
		// 2) another instance of MsgSleep() (beneath us in the stack) needs it (see the comments
		//    in IsCycleComplete() near KILL_MAIN_TIMER for details).
	}

	// Only used when aMode == RETURN_AFTER_MESSAGES:
	// True if the current subroutine was interrupted by another:
	//bool was_interrupted = false;
	bool sleep0_was_done = false;
	bool empty_the_queue_via_peek = false;
	int messages_received = 0; // This is used to ensure we Sleep() at least a minimal amount if no messages are received.

	bool msg_was_handled;
	HWND fore_window, focused_control, criterion_found_hwnd;
	TCHAR wnd_class_name[32];
	UserMenuItem *menu_item;
	HotkeyIDType hk_id;
	Hotkey *hk;
	USHORT variant_id;
	HotkeyVariant *variant;
	int priority;
	Hotstring *hs;
	GuiType *pgui; // This is just a temp variable and should not be referred to once the below has been determined.
	GuiControlType *pcontrol, *ptab_control;
	GuiIndexType gui_control_index;
	GuiEventType gui_action;
	UCHAR gui_event_kind;
	UINT gui_event_code;
	DWORD_PTR gui_event_info;
	BYTE gui_event_byte;
	bool event_is_control_generated;
	MsgMonitorList *check_if_running;
	ExprTokenType gui_event_args[6]; // Current maximum number of arguments for Gui event handlers.
	int gui_event_arg_count;
	POINT gui_point;
	HDROP hdrop_to_free;
	input_type *input_hook;
	LRESULT msg_reply;
	BOOL peek_result;
	MSG msg;

	for (;;) // Main event loop.
	{
		if (aSleepDuration > 0 && !empty_the_queue_via_peek && !g_DeferMessagesForUnderlyingPump) // g_Defer: Requires a series of Peeks to handle non-contiguous ranges, which is why GetMessage() can't be used.
		{
			// The following comment is mostly obsolete as of v1.0.39 (which introduces a thread
			// dedicated to the hooks).  However, using GetMessage() is still superior to
			// PeekMessage() for performance reason.  Add to that the risk of breaking things
			// and it seems clear that it's best to retain GetMessage().
			// Older comment:
			// Use GetMessage() whenever possible -- rather than PeekMessage() or a technique such
			// MsgWaitForMultipleObjects() -- because it's the "engine" that passes all keyboard
			// and mouse events immediately to the low-level keyboard and mouse hooks
			// (if they're installed).  Otherwise, there's greater risk of keyboard/mouse lag.
			// PeekMessage(), depending on how, and how often it's called, will also do this, but
			// I'm not as confident in it.
			if (GetMessage(&msg, NULL, 0, MSG_FILTER_MAX) == -1) // -1 is an error, 0 means WM_QUIT
				continue; // Error probably happens only when bad parameters were passed to GetMessage().
			//else let any WM_QUIT be handled below.
		}
		else // aSleepDuration < 1 || empty_the_queue_via_peek || g_DeferMessagesForUnderlyingPump
		{
			bool do_special_msg_filter, peek_was_done = false; // Set default.
			// Check the active window in each iteration in case a significant amount of time has passed since
			// the previous iteration (due to launching threads, etc.)
			if (g_DeferMessagesForUnderlyingPump && (fore_window = GetForegroundWindow()) != NULL  // There is a foreground window.
				&& GetWindowThreadProcessId(fore_window, NULL) == g_MainThreadID) // And it belongs to our main thread (the main thread is the only one that owns any windows).
			{
				do_special_msg_filter = false; // Set default.
                if (g_nFileDialogs) // v1.0.44.12: Also do the special Peek/msg filter below for FileSelect because testing shows that frequently-running timers disrupt the ability to double-click.
				{
					GetClassName(fore_window, wnd_class_name, _countof(wnd_class_name));
					do_special_msg_filter = !_tcscmp(wnd_class_name, _T("#32770"));  // Due to checking g_nFileDialogs above, this means that this dialog is probably FileSelect rather than MsgBox/InputBox/DirSelect (even if this guess is wrong, it seems fairly inconsequential to filter the messages since other pump beneath us on the call-stack will handle them ok).
				}
				if (!do_special_msg_filter && (focused_control = GetFocus()))
				{
					GetClassName(focused_control, wnd_class_name, _countof(wnd_class_name));
					do_special_msg_filter = !_tcsicmp(wnd_class_name, _T("SysTreeView32")) // A TreeView owned by our thread has focus (includes DirSelect's TreeView).
						|| !_tcsicmp(wnd_class_name, _T("SysListView32"));
				}
				if (do_special_msg_filter)
				{
					// v1.0.48.03: Below now applies to SysListView32 because otherwise a timer that runs
					// while the user is dragging a rectangle around a selection (Marquee) can cause the
					// mouse button to appear to be stuck down down after the user releases it.
					// v1.0.44.12: Below now applies to FileSelect dialogs too (see reason above).
					// v1.0.44.11: Since one of our thread's TreeViews has focus (even in DirSelect), this
					// section is a work-around for the fact that the TreeView's message pump (somewhere beneath
					// us on the call stack) is apparently designed to process some mouse messages directly rather
					// than receiving them indirectly (in its WindowProc) via our call to DispatchMessage() here
					// in this pump.  The symptoms of this issue are an inability of a user to reliably select
					// items in a TreeView (the selection sometimes snaps back to the previously selected item),
					// which can be reproduced by showing a TreeView while a 10ms script timer is running doing
					// a trivial single line such as x=1.
					// NOTE: This happens more often in DirSelect dialogs, I believe because it's msg
					// pump is ALWAYS running but that of a GUI TreeView is running only during mouse capture
					// (i.e. when left/right button is down).
					// This special handling for TreeView can someday be broadened so that focused control's
					// class isn't checked: instead, could check whether left and/or right mouse button is
					// logically down (which hasn't yet been tested).  Or it could be broadened to include
					// other system dialogs and/or common controls that have unusual processing in their
					// message pumps -- processing that requires them to directly receive certain messages
					// rather than having them dispatched directly to their WindowProc.
					peek_was_done = true;
					// Peek() must be used instead of Get(), and Peek() must be called more than once to handle
					// the two ranges on either side of the mouse messages.  But since it would be improper
					// to process messages out of order (and might lead to side-effects), force the retrieval
					// to be in chronological order by checking the timestamps of each Peek first message, and
					// then fetching the one that's oldest (since it should be the one that's been waiting the
					// longest and thus generally should be ahead of the other Peek's message in the queue):
					UINT filter_max = (IsInterruptible() ? UINT_MAX : WM_HOTKEY - 1); // Fixed in v1.1.16 to not use MSG_FILTER_MAX, which would produce 0 when IsInterruptible(). Although WM_MOUSELAST+1..0 seems to produce the right results, MSDN does not indicate that it is valid.
#define PEEK1(mode) PeekMessage(&msg, NULL, 0, WM_MOUSEFIRST-1, mode) // Relies on the fact that WM_MOUSEFIRST < MSG_FILTER_MAX
#define PEEK2(mode) PeekMessage(&msg, NULL, WM_MOUSELAST+1, filter_max, mode)
					if (!PEEK1(PM_NOREMOVE))  // Since no message in Peek1, safe to always use Peek2's (even if it has no message either).
						peek_result = PEEK2(PM_REMOVE);
					else // Peek1 has a message.  So if Peek2 does too, compare their timestamps.
					{
						DWORD peek1_time = msg.time; // Save it due to overwrite in next line.
						if (!PEEK2(PM_NOREMOVE)) // Since no message in Peek2, use Peek1's.
							peek_result = PEEK1(PM_REMOVE);
						else // Both Peek2 and Peek1 have a message waiting, so to break the tie, retrieve the oldest one.
						{
							// In case tickcount has wrapped, compare it the better way (must cast to int to avoid
							// loss of negative values):
							peek_result = ((int)(msg.time - peek1_time) > 0) // Peek2 is newer than Peek1, so treat peak1 as oldest and thus first in queue.
								? PEEK1(PM_REMOVE) : PEEK2(PM_REMOVE);
						}
					}
				}
			}
			if (!peek_was_done) // Since above didn't Peek(), fall back to doing the Peek with the standard filter.
				peek_result = PeekMessage(&msg, NULL, 0, MSG_FILTER_MAX, PM_REMOVE);
			if (!peek_result) // No more messages
			{
				// UPDATE: The section marked "OLD" below is apparently not quite true: although Peek() has been
				// caught yielding our timeslice, it's now difficult to reproduce.  Perhaps it doesn't consistently
				// yield (maybe it depends on the relative priority of competing processes) and even when/if it
				// does yield, it might somehow not be as long or as good as Sleep(0).  This is evidenced by the fact
				// that some of my script's WinWaitClose's finish too quickly when the Sleep(0) is omitted after a
				// Peek() that returned FALSE.
				// OLD (mostly obsolete in light of above): It is not necessary to actually do the Sleep(0) when
				// aSleepDuration == 0 because the most recent PeekMessage() has just yielded our prior timeslice.
				// This is because when Peek() doesn't find any messages, it automatically behaves as though it
				// did a Sleep(0).
				if (aSleepDuration == 0 && !sleep0_was_done)
				{
					Sleep(0);
					sleep0_was_done = true;
					// Now start a new iteration of the loop that will see if we
					// received any messages during the up-to-20ms delay (perhaps even more)
					// that just occurred.  It's done this way to minimize keyboard/mouse
					// lag (if the hooks are installed) that will occur if any key or
					// mouse events are generated during that 20ms.  Note: It seems that
					// the OS knows not to yield our timeslice twice in a row: once for
					// the Sleep(0) above and once for the upcoming PeekMessage() (if that
					// PeekMessage() finds no messages), so it does not seem necessary
					// to check HIWORD(GetQueueStatus(QS_ALLEVENTS)).  This has been confirmed
					// via the following test, which shows that while BurnK6 (CPU maxing program)
					// is foreground, a Sleep(0) really does a Sleep(60).  But when it's not
					// foreground, it only does a Sleep(20).  This behavior is UNAFFECTED by
					// the added presence of a HIWORD(GetQueueStatus(QS_ALLEVENTS)) check here:
					//SplashTextOn,,, xxx
					//WinWait, xxx  ; set last found window
					//Loop
					//{
					//	start = %a_tickcount%
					//	Sleep, 0
					//	elapsed = %a_tickcount%
					//	elapsed -= %start%
					//	WinSetTitle, %elapsed%
					//}
					continue;
				}
				// Otherwise: aSleepDuration is non-zero or we already did the Sleep(0)
				if (messages_received == 0 && allow_early_return)
				{
					// Fix for v1.1.05.04: Since Peek() didn't find a message, avoid maxing the CPU.
					// This specific section is needed for Wait() when an underlying thread
					// is displaying a dialog, and perhaps in other cases.
					// Fix for v1.1.07.00: Avoid Sleep() if caller specified a duration of zero;
					// otherwise SendEvent with a key delay of 0 will be slower than expected.
					// This affects auto-replace hotstrings in SendEvent mode (which is the default
					// when SendInput is unavailable).  Note that if aSleepDuration == 0, Sleep(0)
					// was already called above or by a prior iteration.
					if (aSleepDuration > 0)
						Sleep(5); // This is a somewhat arbitrary value: the intent of a value below 10 is to avoid yielding more than one timeslice on all systems even if they have unusual timeslice sizes / system timers.
					++messages_received; // Don't repeat this section.
					continue;
				}
				// Notes for the macro further below:
				// Must decrement prior to every RETURN to balance it.
				// Do this prior to checking whether timer should be killed, below.
				// Kill the timer only if we're about to return OK to the caller since the caller
				// would still need the timer if FAIL was returned above.  But don't kill it if
				// there are any enabled timed subroutines, because the current policy it to keep
				// the main timer always-on in those cases.  UPDATE: Also avoid killing the timer
				// if there are any script threads running.  To do so might cause a problem such
				// as in this example scenario: MsgSleep() is called for any reason with a delay
				// large enough to require the timer.  The timer is set.  But a msg arrives that
				// MsgSleep() dispatches to MainWindowProc().  If it's a hotkey or custom menu,
				// MsgSleep() is called recursively with a delay of -1.  But when it finishes via
				// IsCycleComplete(), the timer would be wrongly killed because the underlying
				// instance of MsgSleep still needs it.  Above is even more wide-spread because if
				// MsgSleep() is called recursively for any reason, even with a duration >10, it will
				// wrongly kill the timer upon returning, in some cases.  For example, if the first call to
				// MsgSleep(-1) finds a hotkey or menu item msg, and executes the corresponding subroutine,
				// that subroutine could easily call MsgSleep(10+) for any number of reasons, which
				// would then kill the timer.
				// Also require that aSleepDuration > 0 so that MainWindowProc()'s receipt of a
				// WM_HOTKEY msg, to which it responds by turning on the main timer if the script
				// is uninterruptible, is not defeated here.  In other words, leave the timer on so
				// that when the script becomes interruptible once again, the hotkey will take effect
				// almost immediately rather than having to wait for the displayed dialog to be
				// dismissed (if there is one).
				//
				// "we_turned_on_defer" is necessary to prevent us from turning it off if some other
				// instance of MsgSleep beneath us on the calls stack turned it on.  Only it should
				// turn it off because it might still need the "true" value for further processing.
				#define RETURN_FROM_MSGSLEEP \
				{\
					if (we_turned_on_defer)\
						g_DeferMessagesForUnderlyingPump = false;\
					if (this_layer_needs_timer)\
					{\
						--g_nLayersNeedingTimer;\
						if (aSleepDuration > 0 && !g_nLayersNeedingTimer && !g_script.mTimerEnabledCount && !Hotkey::sJoyHotkeyCount)\
							KILL_MAIN_TIMER \
					}\
					return return_value;\
				}
				// IsCycleComplete should always return OK in this case.  Also, was_interrupted
				// will always be false because if this "aSleepDuration < 1" call really
				// was interrupted, it would already have returned in the hotkey cases
				// of the switch().  UPDATE: was_interrupted can now the hotkey case in
				// the switch() doesn't return, relying on us to do it after making sure
				// the queue is empty.
				// The below is checked here rather than in IsCycleComplete() because
				// that function is sometimes called more than once prior to returning
				// (e.g. empty_the_queue_via_peek) and we only want this to be decremented once:
				if (IsCycleComplete(aSleepDuration, start_time, allow_early_return)) // v1.0.44.11: IsCycleComplete() must be called for all modes, but now its return value is checked due to the new g_DeferMessagesForUnderlyingPump mode.
					RETURN_FROM_MSGSLEEP
				// Otherwise (since above didn't return) combined logic has ensured that all of the following are true:
				// 1) aSleepDuration > 0
				// 2) !empty_the_queue_via_peek
				// 3) The above two combined with logic above means that g_DeferMessagesForUnderlyingPump==true.
				Sleep(5); // Since Peek() didn't find a message, avoid maxing the CPU.  This is a somewhat arbitrary value: the intent of a value below 10 is to avoid yielding more than one timeslice on all systems even if they have unusual timeslice sizes / system timers.
				continue;
			}
			// else Peek() found a message, so process it below.
		} // PeekMessage() vs. GetMessage()

		// Since above didn't return or "continue", a message has been received that is eligible
		// for further processing.
		++messages_received;

		// For max. flexibility, it seems best to allow the message filter to have the first
		// crack at looking at the message, before even TRANSLATE_AHK_MSG:
		if (g_MsgMonitor.Count() && MsgMonitor(msg.hwnd, msg.message, msg.wParam, msg.lParam, &msg, msg_reply))  // Count is checked here to avoid function-call overhead.
		{
			continue; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.
			// NOTE: Above does "continue" and ignores msg_reply.  This is because testing shows that
			// messages received via Get/PeekMessage() were always sent via PostMessage.  If an
			// another thread sends ours a message, MSDN implies that Get/PeekMessage() internally
			// calls the message's WindowProc directly and sends the reply back to the other thread.
			// That makes sense because it seems unlikely that DispatchMessage contains any means
			// of replying to a message because it has no way of knowing whether the MSG struct
			// arrived via Post vs. SendMessage.
		}

		// If this message might be for one of our GUI windows, check that before doing anything
		// else with the message.  This must be done first because some of the standard controls
		// also use WM_USER messages, so we must not assume they're generic thread messages just
		// because they're >= WM_USER.  The exception is AHK_GUI_ACTION should always be handled
		// here rather than by IsDialogMessage().
		// UPDATE: MSDN isn't clear about which messages IsDialogMessage() dispatches, so it's
		// now only used for keyboard handling.  This reduces code complexity a little and
		// eliminates some uncertainty about message routing.  All of the cases for WM_USER
		// range messages below already checked msg.hwnd to ensure it is one of our messages.
		if (g_firstGui // Checked first to help performance, since all messages must come through this bottleneck.
			&& msg.message >= WM_KEYFIRST && msg.message <= WM_KEYLAST
			&& (pgui = GuiType::FindGuiParent(msg.hwnd))) // Ordered for short-circuit performance.
		{
			focused_control = msg.hwnd; // Alias for maintainability.  Seems more appropriate (and efficient) to use this vs. GetFocus().

			if (pgui->mAccel) // v1.1.04: Keyboard accelerators.
				if (TranslateAccelerator(pgui->mHwnd, pgui->mAccel, &msg))
					continue; // Above call handled it.

			// Relies heavily on short-circuit boolean order:
			if (  msg.message == WM_KEYDOWN && pgui->mTabControlCount
				&& (msg.wParam == VK_NEXT || msg.wParam == VK_PRIOR || msg.wParam == VK_TAB
					|| msg.wParam == VK_LEFT || msg.wParam == VK_RIGHT)
				&& (pcontrol = pgui->FindControl(focused_control)) && pcontrol->type != GUI_CONTROL_HOTKEY   )
			{
				ptab_control = NULL; // Set default.
				if (pcontrol->type == GUI_CONTROL_TAB) // The focused control is a tab control itself.
				{
					ptab_control = pcontrol;
					// For the below, note that Alt-left and Alt-right are automatically excluded,
					// as desired, since any key modified only by alt would be WM_SYSKEYDOWN vs. WM_KEYDOWN.
					if (msg.wParam == VK_LEFT || msg.wParam == VK_RIGHT)
					{
						pgui->SelectAdjacentTab(*ptab_control, msg.wParam == VK_RIGHT, false, false);
						// Pass false for both the above since that's the whole point of having arrow
						// keys handled separately from the below: Focus should stay on the tabs
						// rather than jumping to the first control of the tab, it focus should not
						// wrap around to the beginning or end (to conform to standard behavior for
						// arrow keys).
						continue; // Suppress this key even if the above failed (probably impossible in this case).
					}
					//else fall through to the next part.
				}
				// If focus is in a multiline edit control, don't act upon Control-Tab (and
				// shift-control-tab -> for simplicity & consistency) since Control-Tab is a special
				// keystroke that inserts a literal tab in the edit control:
				if (   msg.wParam != VK_LEFT && msg.wParam != VK_RIGHT
					&& (GetKeyState(VK_CONTROL) & 0x8000) // Even if other modifiers are down, it still qualifies. Use GetKeyState() vs. GetAsyncKeyState() because the former's definition is more suitable.
					&& (msg.wParam != VK_TAB || pcontrol->type != GUI_CONTROL_EDIT
						|| !(GetWindowLong(pcontrol->hwnd, GWL_STYLE) & ES_MULTILINE))   )
				{
					// If ptab_control wasn't determined above, check if focused control is owned by a tab control:
					if (!ptab_control && !(ptab_control = pgui->FindTabControl(pcontrol->tab_control_index))   )
						// Fall back to the first tab control (for consistency & simplicity, seems best
						// to always use the first rather than something fancier such as "nearest in z-order".
						ptab_control = pgui->FindTabControl(0);
					if (ptab_control && IsWindowEnabled(ptab_control->hwnd))
					{
						pgui->SelectAdjacentTab(*ptab_control
							, msg.wParam == VK_NEXT || (msg.wParam == VK_TAB && !(GetKeyState(VK_SHIFT) & 0x8000)) // Use GetKeyState() vs. GetAsyncKeyState() because the former's definition is more suitable.
							, true, true);
						// Update to the below: Must suppress the tab key at least, to prevent it
						// from navigating *and* changing the tab.  And since this one is suppressed,
						// might as well suppress the others for consistency.
						// Older: Since WM_KEYUP is not handled/suppressed here, it seems best not to
						// suppress this WM_KEYDOWN either (it should do nothing in this case
						// anyway, but for balance this seems best): Fall through to the next section.
						continue;
					}
					//else fall through to the below.
				}
				//else fall through to the below.
			} // Interception of keystrokes for navigation in tab control.

			// v1.0.34: Fix for the fact that a multiline edit control will send WM_CLOSE to its parent
			// when user presses ESC while it has focus.  The following check is similar to the block's above.
			// The alternative to this approach would have been to override the edit control's WindowProc,
			// but the following seemed to be less code. Although this fix is only necessary for multiline
			// edits, its done for all edits since it doesn't do any harm.  In addition, there is no need to
			// check what modifiers are down because we never receive the keystroke for Ctrl-Esc and Alt-Esc
			// (the OS handles those beforehand) and both Win-Esc and Shift-Esc are identical to a naked Esc
			// inside an edit.  The following check relies heavily on short-circuit eval. order.
			if (   msg.message == WM_KEYDOWN
				&& (msg.wParam == VK_ESCAPE || msg.wParam == VK_TAB // v1.0.38.03: Added VK_TAB handling for "WantTab".
					|| (msg.wParam == 'A' && (GetKeyState(VK_CONTROL) & 0x8000) // v1.0.44: Added support for "WantCtrlA".
						&& !(GetKeyState(VK_RMENU) & 0x8000))) // v1.1.17: Exclude AltGr+A (Ctrl+Alt+A).
				&& (pcontrol = pgui->FindControl(focused_control))
				&& pcontrol->type == GUI_CONTROL_EDIT)
			{
				switch(msg.wParam)
				{
				case 'A': // v1.0.44: Support for Ctrl-A to select all text.
					if (!(pcontrol->attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT)) // i.e. presence of AltSubmit bit DISABLES Ctrl-A handling.
					{
						SendMessage(pcontrol->hwnd, EM_SETSEL, 0, -1); // Select all text.
						continue; // Omit this keystroke from any further processing.
					}
					break;
				case VK_ESCAPE:
					pgui->Escape();
					continue; // Omit this keystroke from any further processing.
				default: // VK_TAB
					if ((pcontrol->attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR) // It has the "WantTab" property.
						&& !(GetWindowLong(pcontrol->hwnd, GWL_STYLE) & ES_READONLY)) // It doesn't have "ReadOnly".
					{
						// For flexibility, do this even for single-line edit controls, though in that
						// case the tab keystroke will produce an "empty box" character.
						// Strangely, if a message pump other than this one (MsgSleep) is running,
						// such as that of a MsgBox, "WantTab" is already in effect unconditionally,
						// perhaps because MsgBox and others respond to WM_GETDLGCODE with DLGC_WANTTAB.
						SendMessage(pcontrol->hwnd, EM_REPLACESEL, TRUE, (LPARAM)"\t");
						continue; // Omit this keystroke from any further processing.
					}
				} // switch()
			}

			if (GuiType::sTreeWithEditInProgress && msg.message == WM_KEYDOWN)
			{
				if (msg.wParam == VK_RETURN)
				{
					TreeView_EndEditLabelNow(GuiType::sTreeWithEditInProgress, FALSE); // Save changes to label/text.
					continue;
				}
				else if (msg.wParam == VK_ESCAPE)
				{
					TreeView_EndEditLabelNow(GuiType::sTreeWithEditInProgress, TRUE); // Cancel without saving.
					continue;
				}
			}

			// IsDialogMessage() takes care of standard keyboard handling within the dialog,
			// such as tab to change focus and Enter to activate the default button.
			// Avoid calling IsDialogMessage() for WM_SYSCHAR if the GUI has no controls.
			// It seems that if a GUI has no controls, IsDialogMessage() will return true for
			// Alt+n combinations without invoking the default processing, such as focusing
			// a menu bar item.  IsDialogMessage() still needs to be called for some messages;
			// at the very least, WM_KEYDOWN (VK_ESC) must be intercepted for GuiEscape to work.
			if (pgui->mControlCount || msg.message != WM_SYSCHAR)
			{
				g->CalledByIsDialogMessageOrDispatch = true;
				g->CalledByIsDialogMessageOrDispatchMsg = msg.message; // Added in v1.0.44.11 because it's known that IsDialogMessage can change the message number (e.g. WM_KEYDOWN->WM_NOTIFY for UpDowns)
				msg_was_handled = IsDialogMessage(pgui->mHwnd, &msg); // Pass the dialog HWND, not msg.hwnd, which is often a control.
				g->CalledByIsDialogMessageOrDispatch = false;
				if (msg_was_handled) // This message was handled by IsDialogMessage() above.
					continue; // Continue with the main message loop.
			}
		} // if (keyboard message posted to GUI)

		// v1.0.44: There's no reason to call TRANSLATE_AHK_MSG here because all WM_COMMNOTIFY messages
		// are sent to g_hWnd. Thus, our call to DispatchMessage() later below will route such messages to
		// MainWindowProc(), which will then call TRANSLATE_AHK_MSG().
		//TRANSLATE_AHK_MSG(msg.message, msg.wParam)

		switch(msg.message)
		{
		// MSG_FILTER_MAX should prevent us from receiving this first group of messages whenever g_AllowInterruption or
		// g->AllowThreadToBeInterrupted is false.
		case AHK_HOOK_HOTKEY:  // Sent from this app's keyboard or mouse hook.
		case AHK_HOTSTRING:    // Sent from keybd hook to activate a non-auto-replace hotstring.
		case AHK_CLIPBOARD_CHANGE:
			// This extra handling is present because common controls and perhaps other OS features tend
			// to use WM_USER+NN messages, a practice that will probably be even more common in the future.
			// To cut down on message conflicts, dispatch such messages whenever their HWND isn't a what
			// it should be for a message our own code generated. That way, the target control (if any)
			// will still get the msg.
			if (msg.hwnd && msg.hwnd != g_hWnd) // v1.0.44: It's wasn't sent by our code; perhaps by a common control's code.
				break; // Dispatch it vs. discarding it, in case it's for a control.
			//ELSE FALL THROUGH:
		case AHK_GUI_ACTION:   // The user pressed a button on a GUI window, or some other actionable event. Listed before the below for performance.
		case WM_HOTKEY:        // As a result of this app having previously called RegisterHotkey(), or from TriggerJoyHotkeys().
		case AHK_USER_MENU:    // The user selected a custom menu item.
		case AHK_INPUT_END:    // Input ended (sent by the hook thread).
		case AHK_INPUT_KEYDOWN:
		case AHK_INPUT_CHAR:
		case AHK_INPUT_KEYUP:
		{
			hdrop_to_free = NULL;  // Set default for this message's processing (simplifies code).
			switch(msg.message)
			{
			case AHK_GUI_ACTION: // Listed first for performance.
				// Assume that it is possible that this message's GUI window has been destroyed
				// (and maybe even recreated) since the time the msg was posted.  If this can happen,
				// that's another reason for finding which GUI this control is associate with (it also
				// needs to be found so that we can call the correct GUI window object to perform
				// the action):
				if (   !(pgui = GuiType::FindGui(msg.hwnd))   ) // No associated GUI object, so ignore this event.
					// v1.0.44: Dispatch vs. continue/discard since it's probably for a common control
					// whose msg number happens to be AHK_GUI_ACTION.  Do this *only* when HWND isn't recognized,
					// not when msg content is invalid, because dispatching a msg whose HWND is one of our own
					// GUI windows might cause GuiWindowProc to fwd it back to us, creating an infinite loop.
					goto break_out_of_main_switch; // Goto seems preferably in this case for code size & performance.
				
				gui_event_info =    (DWORD_PTR)msg.lParam;
				gui_action =        LOBYTE(msg.wParam); // Byte 0.
				gui_event_byte =    HIBYTE(msg.wParam); // Byte 1.  Sometimes used as an additional parameter.
				gui_control_index = HIWORD(msg.wParam); // Bytes 2-3.  Caller has set it to NO_CONTROL_INDEX if it isn't applicable.
				gui_event_arg_count = 0;
				gui_event_code = gui_action; // Set default.  Using both this and gui_action simplifies some things.

				pcontrol = gui_control_index < pgui->mControlCount ? pgui->mControl[gui_control_index] : NULL; // Set for use in other places below.

				check_if_running = &pgui->mEvents; // Set default.
				event_is_control_generated = false; // Set default.
				gui_event_kind = GUI_EVENTKIND_EVENT; // Set default.

				switch(gui_action)
				{
				case GUI_EVENT_CLOSE:
				case GUI_EVENT_ESCAPE:
				case GUI_EVENT_RESIZE:
					break;
				case GUI_EVENT_CONTEXTMENU:
					// Must allow multiple threads because otherwise the user cannot right-click twice consecutively
					// (the second click is blocked because the menu is still displayed at the instant of the click).
					check_if_running = NULL;
					break;
				case GUI_EVENT_DROPFILES:
					hdrop_to_free = pgui->mHdrop; // This variable simplifies the code further below.
					if (   !hdrop_to_free // Checked just in case, so that the below can query it.
						|| !(gui_event_info = DragQueryFile(hdrop_to_free, 0xFFFFFFFF, NULL, 0))   ) // Probably impossible, but if it ever can happen, seems best to ignore it.
					{
						if (hdrop_to_free) // Checked again in case short-circuit boolean above never checked it.
						{
							DragFinish(hdrop_to_free); // Since the drop-thread will not be launched, free the memory.
							pgui->mHdrop = NULL; // Indicate that this GUI window is ready for another drop.
						}
						continue;
					}
					// mHdrop != NULL prevents another DROPFILES event from being queued, up until the
					// message is received here and we remove WS_EX_ACCEPTFILES.  After that the user
					// can no longer drop files onto the window, but it may still be necessary to check
					// for running instances in case of:
					//  1) Another WM_DROPFILES already in the queue, but not yet processed.
					//  2) Someone posting WM_DROPFILES (i.e. fake drag and drop).
					break;
				case GUI_EVENT_WM_COMMAND:
					gui_event_kind = GUI_EVENTKIND_COMMAND;
					gui_event_code = (GuiEventType)gui_event_info;
					// Fall through:
				default: // This is an action from a particular control in the GUI window.
					if (!pcontrol) // gui_control_index was beyond the quantity of controls, possibly due to parent window having been destroyed since the msg was sent (or bogus msg).
						continue;  // Discarding an invalid message here is relied upon both other sections below.
					// It seems best by default not to allow multiple threads for the same event.
					// Such events are discarded because it seems likely that most script designers
					// would want to see the effects of faulty design (e.g. long running timers or
					// hotkeys that interrupt gui threads) rather than having events for later,
					// when they might suddenly take effect unexpectedly.
					check_if_running = &pcontrol->events;
					event_is_control_generated = true; // As opposed to a drag-and-drop or context-menu event that targets a specific control.
				} // switch(gui_action)
				
				// IsMonitoring() isn't checked at this stage for performance and code size, and because
				// it was checked prior to posting AHK_GUI_ACTION.  If the handler was removed since then,
				// the section below will do some unnecessary but harmless work.
				if (check_if_running && check_if_running->IsRunning(gui_event_code, gui_event_kind))
				{
					if (hdrop_to_free) // Unlikely but possible, in theory.  Doesn't seem to affect code size.
					{
						DragFinish(hdrop_to_free); // Since the drop-thread will not be launched, free the memory.
						pgui->mHdrop = NULL; // Indicate that this GUI window is ready for another drop.
					}
					continue;
				}
				priority = 0;  // Always use default for now.
				break; // case AHK_GUI_ACTION

			case AHK_USER_MENU: // user-defined menu item
				if (msg.wParam) // Poster specified that this menu item was from a gui's menu bar (since wParam is unsigned, any incoming -1 is seen as greater than max).
				{
					// The menu type is passed with the message so that its value will be in sync with
					// the timestamp of the message (in case this message has been stuck in the queue
					// for a long time).
					// msg.wParam is the HWND rather than a pointer to avoid any chance of problems with
					// a gui object having been destroyed while the msg was waiting in the queue.
					if (!(pgui = GuiType::FindGui((HWND)msg.wParam)) // Not a GUI's menu bar item...
						&& msg.hwnd && msg.hwnd != g_hWnd) // ...and not a script menu item.
						goto break_out_of_main_switch; // See "goto break_out_of_main_switch" higher above for complete explanation.
				}
				else
					pgui = NULL; // Set for use in later sections.
				if (   !(menu_item = g_script.FindMenuItemByID((UINT)msg.lParam))   ) // Item not found.
					continue; // ignore the msg
				// And just in case a menu item that lacks a label (such as a separator) is ever
				// somehow selected (perhaps via someone sending direct messages to us, bypassing
				// the menu):
				if (!menu_item->mCallback)
					continue;
				// Ignore/discard a hotkey or custom menu item event if the current thread's priority
				// is higher than it's:
				priority = menu_item->mPriority;
				break;

			case AHK_HOTSTRING:
				if (msg.wParam >= Hotstring::sHotstringCount) // Invalid hotstring ID (perhaps spoofed by external app)
					continue; // Do nothing.
				hs = Hotstring::shs[msg.wParam];  // For performance and convenience.
				if (hs->mHotCriterion)
				{
					// For details, see comments in the hotkey section of this switch().
					if (   !(criterion_found_hwnd = HotCriterionAllowsFiring(hs->mHotCriterion, hs->mName))   )
						// Hotstring is no longer eligible to fire even though it was when the hook sent us
						// the message.  Abort the firing even though the hook may have already started
						// executing the hotstring by suppressing the final end-character or other actions.
						// It seems preferable to abort midway through the execution than to continue sending
						// keystrokes to the wrong window, or when the hotstring has become suspended.
						continue;
					// For details, see comments in the hotkey section of this switch().
					if (!(hs->mHotCriterion->Type == HOT_IF_ACTIVE || hs->mHotCriterion->Type == HOT_IF_EXIST))
						criterion_found_hwnd = NULL; // For "NONE" and "NOT", there is no last found window.
				}
				else // No criterion, so it's a global hotstring.  It can always fire, but it has no "last found window".
					criterion_found_hwnd = NULL;
				// Do a simple replacement for the hotstring if that's all that's called for.
				// Don't create a new quasi-thread or any of that other complexity done further
				// below.  But also do the backspacing (if specified) for a non-autoreplace hotstring,
				// even if it can't launch due to MaxThreads, MaxThreadsPerHotkey, or some other reason:
				hs->DoReplace(msg.lParam);  // Does only the backspacing if it's not an auto-replace hotstring.
				if (hs->mReplacement) // Fully handled by the above; i.e. it's an auto-replace hotstring.
					continue;
				// Otherwise, continue on and let a new thread be created to handle this hotstring.
				// Since this isn't an auto-replace hotstring, set this value to support
				// the built-in variable A_EndChar:
				g_script.mEndChar = hs->mEndCharRequired ? (TCHAR)LOWORD(msg.lParam) : 0; // v1.0.48.04: Explicitly set 0 when hs->mEndCharRequired==false because LOWORD is used for something else in that case.
				priority = hs->mPriority;
				break;

			case AHK_CLIPBOARD_CHANGE: // Due to the registration of an OnClipboardChange function in the script.
				if (g_script.mOnClipboardChangeIsRunning)
					continue;
				priority = 0;  // Always use default for now.
				break;

			case AHK_INPUT_END:
				input_hook = InputRelease((input_type *)msg.wParam); // The function will verify that it is a valid input_type pointer.
				if (!input_hook)
					continue; // No OnEnd callback/invalid message.
				priority = 0;
				break;

			case AHK_INPUT_KEYDOWN:
			case AHK_INPUT_CHAR:
			case AHK_INPUT_KEYUP:
				for (input_hook = g_input; input_hook && input_hook != (input_type *)msg.wParam; input_hook = input_hook->Prev);
				if (!input_hook)
					continue; // Invalid message or Input already ended (and therefore may have been deleted).
				if (!(msg.message == AHK_INPUT_KEYDOWN ? input_hook->ScriptObject->onKeyDown
					: msg.message == AHK_INPUT_KEYUP ? input_hook->ScriptObject->onKeyUp
					: input_hook->ScriptObject->onChar))
					continue;
				priority = 0;
				break;

			default: // hotkey
				hk_id = msg.wParam & HOTKEY_ID_MASK;
				if (hk_id >= Hotkey::sHotkeyCount) // Invalid hotkey ID.
					continue;
				hk = Hotkey::shk[hk_id];
				// Check if criterion allows firing.
				// For maintainability, this is done here rather than a little further down
				// past the g_MaxThreadsTotal and thread-priority checks.  Those checks hardly
				// ever abort a hotkey launch anyway.
				//
				// If message is WM_HOTKEY, it's either:
				// 1) A joystick hotkey from TriggerJoyHotkeys(), in which case the lParam is ignored.
				// 2) A hotkey message sent by the OS, in which case lParam contains currently-unused info set by the OS.
				//
				// An incoming WM_HOTKEY can be subject to #HotIf Win. at this stage under the following conditions:
				// 1) Joystick hotkey, because it relies on us to do the check so that the check is done only
				//    once rather than twice.
				// 2) #HotIf Win. keybd hotkeys that were made non-hook because they have a non-suspended, global variant.
				//
				// If message is AHK_HOOK_HOTKEY:
				// Rather than having the hook pass the qualified variant to us, it seems preferable
				// to search through all the criteria again and rediscover it.  This is because conditions
				// may have changed since the message was posted, and although the hotkey might still be
				// eligible for firing, a different variant might now be called for (e.g. due to a change
				// in the active window).  Since most criteria hotkeys have at most only a few criteria,
				// and since most such criteria are #HotIf WinActive rather than Exist, the performance will
				// typically not be reduced much at all.  Furthermore, trading performance for greater
				// reliability seems worth it in this case.
				// 
				// The inefficiency of calling HotCriterionAllowsFiring() twice for each hotkey --
				// once in the hook and again here -- seems justified for the following reasons:
				// - It only happens twice if the hotkey a hook hotkey (multi-variant keyboard hotkeys
				//   that have a global variant are usually non-hook, even on NT/2k/XP).
				// - The hook avoids doing its first check of WinActive/Exist if it sees that the hotkey
				//   has a non-suspended, global variant.  That way, hotkeys that are hook-hotkeys for
				//   reasons other than #HotIf Win. (such as mouse, overriding OS hotkeys, or hotkeys
				//   that are too fancy for RegisterHotkey) will not have to do the check twice.
				// - It provides the ability to set the last-found-window for #HotIf WinActive/Exist
				//   (though it's not needed for the "Not" counterparts).  This HWND could be passed
				//   via the message, but that would require malloc-there and free-here, and might
				//   result in memory leaks if its ever possible for messages to get discarded by the OS.
				// - It allows hotkeys that were eligible for firing at the time the message was
				//   posted but that have since become ineligible to be aborted.  This seems like a
				//   good precaution for most users/situations because such hotkey subroutines will
				//   often assume (for scripting simplicity) that the specified window is active or
				//   exists when the subroutine executes its first line.
				// - Most criterion hotkeys use #HotIf WinActive(), which is a very fast call.  Also, although
				//   WinText and/or "SetTitleMatchMode 'Slow'" slow down window searches, those are rarely
				//   used too.
				//
				variant = NULL; // Set default.
				// For #HotIf hotkey variants, we don't want to evaluate the expression a second time. If the hook
				// thread determined that a specific variant should fire, it is passed via the high word of wParam:
				if (variant_id = HIWORD(msg.wParam))
				{
					// The following relies on the fact that variants can't be removed or re-ordered;
					// variant_id should always be the variant's one-based index in the linked list:
					--variant_id; // i.e. index 1 should be mFirstVariant, not mFirstVariant->mNextVariant.
					for (variant = hk->mFirstVariant; variant_id; variant = variant->mNextVariant, --variant_id);
				}
				if (   !(variant || (variant = hk->CriterionAllowsFiring(&criterion_found_hwnd
					, msg.message == AHK_HOOK_HOTKEY ? KEY_IGNORE_LEVEL(HIWORD(msg.lParam)) : 0)))   )
					continue; // No criterion is eligible, so ignore this hotkey event (see other comments).
					// If this is AHK_HOOK_HOTKEY, criterion was eligible at time message was posted,
					// but not now.  Seems best to abort (see other comments).
				
				// Due to the key-repeat feature and the fact that most scripts use a value of 1
				// for their #MaxThreadsPerHotkey, this check will often help average performance
				// by avoiding a lot of unnecessary overhead that would otherwise occur:
				if (!hk->PerformIsAllowed(*variant))
				{
					// The key is buffered in this case to boost the responsiveness of hotkeys
					// that are being held down by the user to activate the keyboard's key-repeat
					// feature.  This way, there will always be one extra event waiting in the queue,
					// which will be fired almost the instant the previous iteration of the subroutine
					// finishes (this above description applies only when MaxThreadsPerHotkey is 1,
					// which it usually is).
					hk->RunAgainAfterFinished(*variant); // Wheel notch count (g->EventInfo below) should be okay because subsequent launches reuse the same thread attributes to do the repeats.
					continue;
				}

				// Now that above has ensured variant is non-NULL:
				HotkeyCriterion *hc = variant->mHotCriterion;
				if (!hc || hc->Type == HOT_IF_NOT_ACTIVE || hc->Type == HOT_IF_NOT_EXIST)
					criterion_found_hwnd = NULL; // For "NONE" and "NOT", there is no last found window.
				else if (HOT_IF_REQUIRES_EVAL(hc->Type))
					criterion_found_hwnd = g_HotExprLFW; // For #HotIf WinExist(WinTitle) and similar.

				priority = variant->mPriority;
			} // switch(msg.message)

			if (g_nThreads >= g_MaxThreadsTotal)
			{
				// Allow only a limited number of recursion levels to avoid any chance of
				// stack overflow.  So ignore this message.  Later, can devise some way
				// to support "queuing up" these launch-thread events for use later when
				// there is "room" to run them, but that might cause complications because
				// in some cases, the user didn't intend to hit the key twice (e.g. due to
				// "fat fingers") and would have preferred to have it ignored.  Doing such
				// might also make "infinite key loops" harder to catch because the rate
				// of incoming hotkeys would be slowed down to prevent the subroutines from
				// running concurrently.
				if (hdrop_to_free) // This is only non-NULL when pgui is non-NULL and gui_action==GUI_EVENT_DROPFILES
				{
					DragFinish(hdrop_to_free); // Since the drop-thread will not be launched, free the memory.
					pgui->mHdrop = NULL; // Indicate that this GUI window is ready for another drop.
				}
				if (msg.message == AHK_INPUT_END)
					input_hook->ScriptObject->Release();
				continue;
				// If the above "continued", it seems best not to re-queue/buffer the key since
				// it might be a while before the number of threads drops back below the limit.
			}

			// Discard the event if it's priority is lower than that of the current thread:
			if (priority < g->Priority)
			{
				if (hdrop_to_free) // This is only non-NULL when pgui is non-NULL and gui_action==GUI_EVENT_DROPFILES
				{
					DragFinish(hdrop_to_free); // Since the drop-thread will not be launched, free the memory.
					pgui->mHdrop = NULL; // Indicate that this GUI window is ready for another drop.
				}
				if (msg.message == AHK_INPUT_END)
					input_hook->ScriptObject->Release();
				continue;
			}

			// Now it is certain that the new thread will be launched, so set everything up.
			// Perform the new thread's subroutine:
			return_value = true; // We will return this value to indicate that we launched at least one new thread.

			// UPDATE v1.0.48: The main timer is no longer killed because testing shows that
			// SetTimer() and/or KillTimer() are relatively slow calls.  Thus it is likely that
			// on average, it's better to receive some unnecessary WM_TIMER messages (especially
			// since WM_TIMER processing is fast when there's nothing to do) than it is to
			// KILL and then RE-SET the main timer for every new thread (especially rapid-fire
			// threads like some GUI threads can be).  This also makes the thread types that are
			// handled here more like other threads such as timers, callbacks, and OnMessage.
			// Also, not doing the following KILL_MAIN_TIMER avoids have to check whether
			// SET_MAIN_TIMER is needed in two places further below (e.g. RETURN_FROM_MSGSLEEP).
			// OLDER COMMENTS:
			// Always kill the main timer, for performance reasons and for simplicity of design,
			// prior to embarking on new subroutine whose duration may be long (e.g. the called
			// subroutine may not return to us for seconds, minutes, or more; during which time
			// we don't want the timer running because it will only fill up the queue with
			// WM_TIMER messages and thus hurt performance).
			// UPDATE: But don't kill it if it should be always-on to support the existence of
			// at least one enabled timed subroutine or joystick hotkey:
			//if (!g_script.mTimerEnabledCount && !Hotkey::sJoyHotkeyCount)
			//	KILL_MAIN_TIMER;

			switch(msg.message)
			{
			case AHK_GUI_ACTION: // Listed first for performance.
			case AHK_CLIPBOARD_CHANGE:
			case AHK_INPUT_END:
			case AHK_INPUT_KEYDOWN:
			case AHK_INPUT_CHAR:
			case AHK_INPUT_KEYUP:
			case AHK_USER_MENU: // user-defined menu item
				break; // Do nothing at this stage.
			default: // hotkey or hotstring
				// Just prior to launching the hotkey, update these values to support built-in
				// variables such as A_TimeSincePriorHotkey:
				g_script.mPriorHotkeyName = g_script.mThisHotkeyName;
				g_script.mPriorHotkeyStartTime = g_script.mThisHotkeyStartTime;
				// Unlike hotkeys -- which can have a name independent of their label by being created or updated
				// with the HOTKEY command -- a hot string's unique name is always its label since that includes
				// the options that distinguish between (for example) :c:ahk:: and ::ahk::
				g_script.mThisHotkeyName = (msg.message == AHK_HOTSTRING) ? hs->mName : hk->mName;
				g_script.mThisHotkeyStartTime = GetTickCount(); // Fixed for v1.0.35.10 to not happen for GUI threads.
			}

			// Make every newly launched subroutine start off with the global default values that
			// the user set up in the auto-execute part of the script (e.g. KeyDelay, WinDelay, etc.).
			InitNewThread(priority, false, true);
			global_struct &g = *::g; // ONLY AFTER above is it safe to "lock in". Reduces code size a bit (31 bytes currently) and may improve performance.  Eclipsing ::g with local g makes compiler remind/enforce the use of the right one.

			// Do this nearly last, right before launching the thread:
			g_script.mLastPeekTime = GetTickCount();
			// v1.0.38.04: The above now resets mLastPeekTime too to reduce situations in which a thread
			// doesn't even run one line before being interrupted by another thread.  Here's how that would
			// happen: ExecUntil() would see that a Peek() is due and call PeekMessage().  The Peek() will
			// yield if we have no messages and the CPU is under heavy load, and thus the script might not
			// get another timeslice for 20ms (or even longer if there is more than one other needy process).
			// Even if the Peek() doesn't yield (i.e. we have messages), those messages might take a long time
			// to process (such as WM_PAINT) even though the script is uninterruptible.  Either way, when the
			// Peek-check completes, a long time might have passed, and the thread might now be interruptible
			// due to the interruptible-timer having expired (which is probably possible only in the no-yield
			// scenario above, since in the case of yield, ExecUntil wouldn't check messages again after the
			// yield).  Thus, the Peek-check's MsgSleep() might launch an interrupting thread before the prior
			// thread had a chance to execute even one line.  Resetting mLastPeekTime above should alleviate that,
			// perhaps even completely resolve it due to the way tickcounts tend not to change early on in
			// a timeslice (perhaps because timeslices fall exactly upon tick-count boundaries).  If it doesn't
			// completely resolve it, mLastPeekTime could instead be set to zero as a special value that
			// ExecUntil recognizes to do the following processing, but this processing reduces performance
			// by 2.5% in a simple addition-loop benchmark:
			//if (g_script.mLastPeekTime)
			//	LONG_OPERATION_UPDATE
			//else
			//	g_script.mLastPeekTime = GetTickCount();

			switch (msg.message)
			{
			case AHK_GUI_ACTION: // Listed first for performance.
			{
#define EVT_ARG_ADD(_value) gui_event_args[gui_event_arg_count++].SetValue(_value)

				// The following check isn't used because a Gui with an event sink can also have events handled
				// by unrelated function objects.  Instead, MsgMonitorList::Call adjusts parameters as needed:
				//if (event_is_control_generated || pgui->mEventSink != pgui)
				EVT_ARG_ADD(event_is_control_generated ? (IObject*)pcontrol : (IObject*)pgui);

				switch(gui_action)
				{
				case GUI_EVENT_CONTEXTMENU:
					bool from_keyboard; // True if this context-menu was generated via the keyboard (such as AppsKey or Shift-F10):
					from_keyboard = gui_event_byte;
					gui_point = msg.pt; // Set default. v1.0.38: More accurate/customary to use msg.pt than GetCursorPos().
					if (pcontrol) // i.e. this context menu is for a control rather than a click somewhere in the parent window itself.
					{
						// By definition, pcontrol should be the focused control.  However, testing shows that
						// it can also be the only control in a window that lacks any focus-capable controls.
						// If the window has no controls at all, testing shows that pcontrol will be NULL,
						// in which case GuiPoint default set earlier is retained (for AppsKey too).
						if (from_keyboard)
							pgui->ControlGetPosOfFocusedItem(*pcontrol, gui_point); // Since pcontrol!=NULL, find out which item is focused in this control.
						//else this is a context-menu event that was invoked via normal mouse click.  Leave
						// g.GuiPoint at its default set earlier.
						switch(pcontrol->type)
						{
						case GUI_CONTROL_LISTBOX:
							// Retrieve the focused item, which is not necessarily the one which was clicked.
							// LB_ITEMFROMPOINT could be used to get the item at the mouse position, but since
							// the control doesn't provide any visual feedback that the item was clicked, it
							// doesn't feel right.  We could provide feedback by automatically selecting the
							// item, but that would be non-trivial due to single- vs. multi-select, and it
							// might not be what the user wants anyway.
							gui_event_info = 1 + (int)SendMessage(pcontrol->hwnd, LB_GETCARETINDEX, 0, 0); // Cast to int to preserve any -1 value.
							break;
						case GUI_CONTROL_LISTVIEW:
							if (from_keyboard) // AppsKey or Shift+F10.
								gui_event_info = 1 + ListView_GetNextItem(pcontrol->hwnd, -1, LVNI_FOCUSED);
							// Otherwise, this event should have been generated by NM_RCLICK, in which case
							// gui_event_info already contains the row number.
							break;
						case GUI_CONTROL_TREEVIEW:
							// Retrieve the HTREEITEM that is the true target of this event.
							if (from_keyboard) // AppsKey or Shift+F10.
								gui_event_info = (DWORD_PTR)SendMessage(pcontrol->hwnd, TVM_GETNEXTITEM, TVGN_CARET, NULL); // Get focused item.
							else // Context menu invoked via right-click.  Find out which item (if any) was clicked on.
							{
								// Use HitTest because the focused item isn't necessarily the one that was
								// clicked on.  This is because unlike a ListView, right-clicking a TreeView
								// item doesn't change the focus.  Instead, the target item is momentarily
								// selected for visual confirmation and then unselected.
								TVHITTESTINFO ht;
								ht.pt = msg.pt;
								ScreenToClient(pcontrol->hwnd, &ht.pt);
								gui_event_info = (DWORD_PTR)TreeView_HitTest(pcontrol->hwnd, &ht);
							}
							break;
						}
					}
					//else pcontrol==NULL: Since there is no focused control, it seems best to report the
					// cursor's position rather than some arbitrary center-point, or top-left point in the
					// parent window.  This is because it might be more convenient for the user to move the
					// mouse to select a menu item (since menu will be close to mouse cursor).
					
					// Convert to client coordinates, since that's what we use for positioning controls,
					// and it's usually the default Menu CoordMode.  Since the script might not use the
					// coordinates to display a menu, it seems best to use this even if the script has
					// changed the default CoordMode.
					ScreenToClient(pgui->mHwnd, &gui_point);
					
					// Build event arguments.
					if (pcontrol)
						EVT_ARG_ADD(pcontrol);
					else
						EVT_ARG_ADD(_T(""));
					EVT_ARG_ADD((__int64)gui_event_info);
					EVT_ARG_ADD(!from_keyboard); // True if mouse-activated.
					EVT_ARG_ADD(gui_point.x);
					EVT_ARG_ADD(gui_point.y);
					break; // case GUI_CONTEXT_MENU.

				case GUI_EVENT_DROPFILES:
					gui_point = msg.pt; // v1.0.38: More accurate/customary to use msg.pt than GetCursorPos().
					ScreenToClient(pgui->mHwnd, &gui_point); // Seems more useful/appropriate than window coordinates, especially now that Client CoordMode is the starting default.
					// Visually indicate that drops aren't allowed while and existing drop is still being
					// processed. Fix for v1.0.31.02: The window's current ExStyle is fetched every time
					// in case a non-GUI command altered it (such as making it transparent):
					SetWindowLong(pgui->mHwnd, GWL_EXSTYLE, GetWindowLong(pgui->mHwnd, GWL_EXSTYLE) & ~WS_EX_ACCEPTFILES);

					// Build event arguments.
					if (pcontrol)
						EVT_ARG_ADD(pcontrol);
					else
						EVT_ARG_ADD(_T(""));
					EVT_ARG_ADD(GuiType::CreateDropArray(hdrop_to_free)); // Currently must be the third-last arg; see "Free the drop array."
					EVT_ARG_ADD(pgui->Unscale(gui_point.x));
					EVT_ARG_ADD(pgui->Unscale(gui_point.y));

					// Free the drop object.
					DragFinish(hdrop_to_free);
					pgui->mHdrop = NULL;
					break;

				case GUI_EVENT_CLOSE:
				case GUI_EVENT_ESCAPE:
				case GUI_EVENT_RESIZE:
					if (gui_action == GUI_EVENT_RESIZE)
					{
						// Build event arguments.
						EVT_ARG_ADD((char)gui_event_byte); // MinMax state
						EVT_ARG_ADD(pgui->Unscale(LOWORD(gui_event_info))); // Width
						EVT_ARG_ADD(pgui->Unscale(HIWORD(gui_event_info))); // Height
					}
					break;

				case GUI_EVENT_WM_COMMAND: // Control-generated, but no additional parameters.
					break;

				default: // Other control-generated event (i.e. event_is_control_generated==true).
					if (pcontrol->type == GUI_CONTROL_LINK && gui_action == GUI_EVENT_CLICK)
					{
						LITEM item;
						item.mask = LIF_URL|LIF_ITEMID|LIF_ITEMINDEX;
						item.iLink = (int)gui_event_info - 1;
						if (!SendMessage(pcontrol->hwnd, LM_GETITEM, NULL, (LPARAM)&item))
						{
							// Make sure they are initialized to sensible values.
							*item.szID = '\0';
							*item.szUrl = '\0';
						}
						if (*item.szID)
							EVT_ARG_ADD(item.szID); // ID attribute.
						else
							EVT_ARG_ADD((int)gui_event_info); // One-based index.
						EVT_ARG_ADD(item.szUrl); // Empty unless it had a href which failed to execute.
						break;
					}
					EVT_ARG_ADD((__int64)gui_event_info);
					switch (gui_action)
					{
					case GUI_EVENT_ITEMSELECT:
						if (pcontrol->type != GUI_CONTROL_LISTVIEW)
							break; // Should be a TreeView, which does not raise an event for de-selection.
						// Otherwise, this is a ListView, so add the parameter indicating whether this
						// is selection or de-selection:
					case GUI_EVENT_ITEMCHECK: // Parameter indicates checked (1) vs unchecked (0).
					case GUI_EVENT_ITEMEXPAND: // Parameter indicates expanded (1) vs collapsed (0).
						EVT_ARG_ADD((int)gui_event_byte - 1);
						break;
					}
				} // switch (msg.message)

				pgui->AddRef(); // Keep the pointer valid at least until the thread finishes.

				ResultType result;
				INT_PTR retval;

				if (gui_action == GUI_EVENT_CONTEXTMENU && pcontrol)
				{
					// Call the control's context menu handler, if any, omitting the "Gui" parameter
					// for consistency with other Ctrl events (and perhaps convenience).  But don't
					// omit the first parameter if the Gui parameter was already omitted.
					int arg_to_omit = pgui->mEventSink != pgui ? 1 : 0;
					result = pcontrol->events.Call(gui_event_args + arg_to_omit, gui_event_arg_count - arg_to_omit, gui_event_code, gui_event_kind, pgui);
					if (result == EARLY_RETURN // Suppress the GUI's handler for this event, if any.
						|| !pgui->mHwnd) // Gui was destroyed.
					{
						pgui->Release();
						break;
					}
				}

				MsgMonitorList &events = event_is_control_generated ? pcontrol->events : pgui->mEvents;
				result = events.Call(gui_event_args, gui_event_arg_count, gui_event_code, gui_event_kind, pgui, &retval);
				
				if (pgui->mHwnd) // i.e. GUI was not destroyed.
				{
					switch (gui_action)
					{
					case GUI_EVENT_CLOSE:
						// If the return value is false/unspecified, hide/destroy the Gui.
						if (!retval)
							pgui->Cancel();
						break;
					case GUI_EVENT_DROPFILES:
						if (pgui->IsMonitoring(gui_action)) // Reapply the style only if we're still monitoring this event.
						{
							// The window's current ExStyle is fetched every time in case a non-GUI
							// command altered it (such as making it transparent):
							SetWindowLong(pgui->mHwnd, GWL_EXSTYLE, GetWindowLong(pgui->mHwnd, GWL_EXSTYLE) | WS_EX_ACCEPTFILES);
						}
						break;
					}
				}
				if (gui_action == GUI_EVENT_DROPFILES) // Must be done regardless of pgui->mHwnd.
					gui_event_args[gui_event_arg_count - 3].object->Release(); // Free the drop array.
				// Counteract the earlier AddRef(). If the Gui was destroyed (and none of this
				// Gui's other labels are still running), this will free the Gui structure.
				pgui->Release();
				break;
			} // case AHK_GUI_ACTION

			case AHK_USER_MENU: // user-defined menu item
			{
				// Below: the menu type is passed with the message so that its value will be in sync
				// with the timestamp of the message (in case this message has been stuck in the
				// queue for a long time):
				if (pgui) // Set by an earlier stage. It means poster specified that this menu item was from a gui's menu bar.
				{
					// As documented, set the last found window if possible/applicable.  It's not necessary to
					// check IsWindow/IsWindowVisible/DetectHiddenWindows since GetValidLastUsedWindow()
					// takes care of that whenever the script actually tries to use the last found window.
					g.hWndLastUsed = pgui->mHwnd; // OK if NULL.
					// This flags GUI menu items as being GUI so that the script has a way of detecting
					// whether a given submenu's item was selected from inside a menu bar vs. a popup:
					g.EventInfo = (EventInfoType)pgui->mHwnd;
					pgui->AddRef();
				}
				UserMenu *menu = menu_item->mMenu; // In case the item is deleted.
				menu->AddRef();
				ExprTokenType param[] =
				{
					menu_item->mName, // This should be immediately copied into a script variable, so it's okay if menu_item is deleted.
					(__int64)(menu_item->Pos() + 1), // +1 to convert zero-based to one-based.
					menu
				};
				menu_item->mCallback->ExecuteInNewThread(_T("Menu"), param, _countof(param));
				menu->Release();
				if (pgui)
					pgui->Release();
				break;
			}

			case AHK_HOTSTRING:
				g.hWndLastUsed = criterion_found_hwnd; // v1.0.42. Even if the window is invalid for some reason, IsWindow() and such are called whenever the script accesses it (GetValidLastUsedWindow()).
				g.SendLevel = hs->mInputLevel;
				g.HotCriterion = hs->mHotCriterion; // v2: Let the Hotkey command use the criterion of this hotstring by default.
				hs->PerformInNewThreadMadeByCaller();
				break;

			case AHK_CLIPBOARD_CHANGE:
			{
				int type = CountClipboardFormats() ? (IsClipboardFormatAvailable(CF_NATIVETEXT) || IsClipboardFormatAvailable(CF_HDROP) ? 1 : 2) : 0;
				ExprTokenType param ((__int64)type);
				g_script.mOnClipboardChangeIsRunning = true;
				DEBUGGER_STACK_PUSH(_T("OnClipboardChange"))
				g_script.mOnClipboardChange.Call(&param, 1, 1);
				DEBUGGER_STACK_POP()
				g_script.mOnClipboardChangeIsRunning = false;
				break;
			}

			case AHK_INPUT_END:
			{
				ExprTokenType param = input_hook->ScriptObject;
				IObjectPtr(input_hook->ScriptObject->onEnd)->ExecuteInNewThread(_T("InputHook"), &param, 1);
				input_hook->ScriptObject->Release();
				break;
			}
			
			case AHK_INPUT_KEYDOWN:
			case AHK_INPUT_KEYUP:
			{
				ExprTokenType params[] =
				{
					input_hook->ScriptObject,
					__int64(vk_type(msg.lParam)),
					__int64(sc_type(msg.lParam >> 16)),
				};
				IObjectPtr onKey = msg.message == AHK_INPUT_KEYDOWN ? input_hook->ScriptObject->onKeyDown : input_hook->ScriptObject->onKeyUp;
				onKey->ExecuteInNewThread(_T("InputHook"), params, _countof(params));
				break;
			}

			case AHK_INPUT_CHAR:
			{
				TCHAR chars[] = { TCHAR(msg.lParam), TCHAR(msg.lParam >> 16), '\0' };
				ExprTokenType params[] =
				{
					input_hook->ScriptObject,
					chars
				};
				IObjectPtr(input_hook->ScriptObject->onChar)->ExecuteInNewThread(_T("InputHook"), params, _countof(params));
				break;
			}

			default: // hotkey
				if (IS_WHEEL_VK(hk->mVK)) // If this is true then also: msg.message==AHK_HOOK_HOTKEY
					g.EventInfo = LOWORD(msg.lParam); // v1.0.43.03: Override the thread default of 0 with the number of notches by which the wheel was turned.
					// Above also works for RunAgainAfterFinished since that feature reuses the same thread attributes set above.
				g.hWndLastUsed = criterion_found_hwnd; // v1.0.42. Even if the window is invalid for some reason, IsWindow() and such are called whenever the script accesses it (GetValidLastUsedWindow()).
				g.SendLevel = variant->mInputLevel;
				g.HotCriterion = variant->mHotCriterion; // v2: Let the Hotkey command use the criterion of this hotkey variant by default.
				hk->PerformInNewThreadMadeByCaller(*variant);
				
			}

			// v1.0.37.06: Call ResumeUnderlyingThread() even if aMode==WAIT_FOR_MESSAGES; this is for
			// maintainability and also in case the pause command has been used to unpause the idle thread.
			ResumeUnderlyingThread();
			// DUE TO THE --g DONE BY THE LINE ABOVE, ANYTHING BETWEEN THIS POINT AND THE NEXT '}' MUST USE :;g INSTEAD OF g.

			if (aMode == WAIT_FOR_MESSAGES) // This is the idle state, so we were called directly from WinMain().
				continue; // The above condition is possible only when the AutoExec section had ended prior to the thread we just launched.

			// Otherwise a script thread other than the idle thread has just been resumed.
			if (IsCycleComplete(aSleepDuration, start_time, allow_early_return))
			{
				// Check for messages once more in case the subroutine that just completed
				// above didn't check them that recently.  This is done to minimize the time
				// our thread spends *not* pumping messages, which in turn minimizes keyboard
				// and mouse lag if the hooks are installed (even though this is no longer
				// true due to v1.0.39's dedicated hook thread, it seems best to continue this
				// practice to maximize responsiveness of hotkeys, the app itself [e.g. tray
				// menu], and also to retain backward compatibility).  Set the state of this
				// function/layer/instance so that it will use peek-mode.  UPDATE: Don't change
				// the value of aSleepDuration to -1 because IsCycleComplete() needs to know the
				// original sleep time specified by the caller to determine whether
				// to decrement g_nLayersNeedingTimer:
				empty_the_queue_via_peek = true;
				allow_early_return = true;
				// And now let it fall through to the "continue" statement below.
			}
			// and now if the cycle isn't complete, stay in the blessed GetMessage() state until the time
			// has expired.
			continue;
		} // End of cases that launch new threads, such as hotkeys and GUI events.

		case WM_TIMER:
			if (msg.lParam // This WM_TIMER is intended for a TimerProc...
				|| msg.hwnd != g_hWnd) // ...or it's intended for a window other than the main window, which implies that it doesn't belong to program internals (i.e. the script is probably using it). This fix was added in v1.0.47.02 and it also fixes the ES_NUMBER balloon bug.
				break; // Fall through to let a later section do DispatchMessage() on it.
			// It seems best to poll the joystick for every WM_TIMER message (i.e. every 10ms or so on
			// NT/2k/XP).  This is because if the system is under load, it might be 20ms, 40ms, or even
			// longer before we get a timeslice again and that is a long time to be away from the poll
			// (a fast button press-and-release might occur in less than 50ms, which could be missed if
			// the polling frequency is too low):
			POLL_JOYSTICK_IF_NEEDED // Do this first since it's much faster.
			CHECK_SCRIPT_TIMERS_IF_NEEDED
			if (aMode == WAIT_FOR_MESSAGES)
				// Timer should have already been killed if we're in this state.
				// But there might have been some WM_TIMER msgs already in the queue
				// (they aren't removed when the timer is killed).  Or perhaps
				// a caller is calling us with this aMode even though there
				// are suspended subroutines (currently never happens).
				// In any case, these are always ignored in this mode because
				// the caller doesn't want us to ever return.  UPDATE: This can now
				// happen if there are any enabled timed subroutines we need to keep an
				// eye on, which is why the mTimerEnabledCount value is checked above
				// prior to starting a new iteration.
				continue;
			if (aSleepDuration < 1) // In this case, WM_TIMER messages have already fulfilled their function, above.
				continue;
			// Otherwise aMode == RETURN_AFTER_MESSAGES:
			// Realistically, there shouldn't be any more messages in our queue
			// right now because the queue was stripped of WM_TIMER messages
			// prior to the start of the loop, which means this WM_TIMER
			// came in after the loop started.  So in the vast majority of
			// cases, the loop would have had enough time to empty the queue
			// prior to this message being received.  Therefore, just return rather
			// than trying to do one final iteration in peek-mode (which might
			// complicate things, i.e. the below function makes certain changes
			// in preparation for ending this instance/layer, only to be possibly,
			// but extremely rarely, interrupted/recursed yet again if that final
			// peek were to detect a recursable message):
			if (IsCycleComplete(aSleepDuration, start_time, allow_early_return))
				RETURN_FROM_MSGSLEEP
			// Otherwise, stay in the blessed GetMessage() state until the time has expired:
			continue;

		case WM_CANCELJOURNAL:
			// IMPORTANT: It's tempting to believe that WM_CANCELJOURNAL might be lost/dropped if the script
			// is displaying a MsgBox or other dialog that has its own msg pump (since such a pump would
			// discard any msg with a NULL HWND).  However, that is not true in this case because such a dialog's
			// msg pump would be beneath this one on the call stack.  This is because our caller is calling us in
			// a loop that does not permit the script to display any *additional* dialogs.  Thus, our msg pump
			// here should always receive the OS-generated WM_CANCELJOURNAL msg reliably.
			// v1.0.44: This message is now received only when the user presses Ctrl-Alt-Del or Ctrl-Esc during
			// playback. For performance and simplicity, the playback hook itself no longer sends this message,
			// instead directly sets g_PlaybackHook = NULL to notify the installer of the hook that it's gone.
			g_PlaybackHook = NULL; // A signal for caller.
			empty_the_queue_via_peek = true;
			// Above is set to so that we return faster, since our caller should be SendKeys() whenever
			// WM_CANCELJOURNAL is received, and SendKeys() benefits from a faster return.
			continue;

		case WM_KEYDOWN:
			if (msg.hwnd == g_hWndEdit && msg.wParam == VK_ESCAPE)
			{
				// This won't work if a MessageBox() window is displayed because its own internal
				// message pump will dispatch the message to our edit control, which will just
				// ignore it.  And avoiding setting the focus to the edit control won't work either
				// because the user can simply click in the window to set the focus.  But for now,
				// this is better than nothing:
				ShowWindow(g_hWnd, SW_HIDE);  // And it's okay if this msg gets dispatched also.
				continue;
			}
			// Otherwise, break so that the messages will get dispatched.  We need the other
			// WM_KEYDOWN msgs to be dispatched so that the cursor is keyboard-controllable in
			// the edit window:
			break;

		case WM_QUIT:
			// The app normally terminates before WM_QUIT is ever seen here because of the way
			// WM_CLOSE is handled by MainWindowProc().  However, this is kept here in case anything
			// external ever explicitly posts a WM_QUIT to our thread's queue:
			g_script.ExitApp(EXIT_CLOSE);
			continue; // Since ExitApp() won't necessarily exit.
		} // switch()
break_out_of_main_switch:

		// If a "continue" statement wasn't encountered somewhere in the switch(), we want to
		// process this message in a more generic way.
		// This little part is from the Miranda source code.  But it doesn't seem
		// to provide any additional functionality: You still can't use keyboard
		// keys to navigate in the dialog unless it's the topmost dialog.
		// UPDATE: The reason it doesn't work for non-topmost MessageBoxes is that
		// this message pump isn't even the one running.  It's the pump of the
		// top-most MessageBox itself, which apparently doesn't properly dispatch
		// all types of messages to other MessagesBoxes.  However, keeping this
		// here is probably a good idea because testing reveals that it does
		// sometimes receive messages intended for MessageBox windows (which makes
		// sense because our message pump here retrieves all thread messages).
		// It might cause problems to dispatch such messages directly, since
		// IsDialogMessage() is supposed to be used in lieu of DispatchMessage()
		// for these types of messages.
		// NOTE: THE BELOW IS CONFIRMED to be needed, at least for a FileSelect()
		// dialog whose quasi-thread has been suspended, and probably for some of the other
		// types of dialogs as well:
		if ((fore_window = GetForegroundWindow()) != NULL  // There is a foreground window.
			&& GetWindowThreadProcessId(fore_window, NULL) == g_MainThreadID) // And it belongs to our main thread (the main thread is the only one that owns any windows).
		{
			GetClassName(fore_window, wnd_class_name, _countof(wnd_class_name));
			if (!_tcscmp(wnd_class_name, _T("#32770")))  // MsgBox, InputBox, FileSelect, DirSelect dialog.
			{
				g->CalledByIsDialogMessageOrDispatch = true; // In case there is any way IsDialogMessage() can call one of our own window proc's rather than that of a MsgBox, etc.
				g->CalledByIsDialogMessageOrDispatchMsg = msg.message; // Added in v1.0.44.11 because it's known that IsDialogMessage can change the message number (e.g. WM_KEYDOWN->WM_NOTIFY for UpDowns)
				if (IsDialogMessage(fore_window, &msg))  // This message is for it, so let it process it.
				{
					g->CalledByIsDialogMessageOrDispatch = false;
					continue;  // This message is done, so start a new iteration to get another msg.
				}
				g->CalledByIsDialogMessageOrDispatch = false;
			}
		}
		// Translate keyboard input for any of our thread's windows that need it:
		if (!g_hAccelTable || !TranslateAccelerator(g_hWnd, g_hAccelTable, &msg))
		{
			g->CalledByIsDialogMessageOrDispatch = true; // Relies on the fact that the types of messages we dispatch can't result in a recursive call back to this function.
			g->CalledByIsDialogMessageOrDispatchMsg = msg.message; // Added in v1.0.44.11. Do it prior to Translate & Dispatch in case either one of them changes the message number (it is already known that IsDialogMessage can change message numbers).
			TranslateMessage(&msg);
			DispatchMessage(&msg); // This is needed to send keyboard input and other messages to various windows and for some WM_TIMERs.
			g->CalledByIsDialogMessageOrDispatch = false;
		}
	} // infinite-loop
}



ResultType IsCycleComplete(int aSleepDuration, DWORD aStartTime, bool aAllowEarlyReturn)
// This function is used just to make MsgSleep() more readable/understandable.
{
	// Note: Even if TickCount has wrapped due to system being up more than about 49 days,
	// DWORD subtraction still gives the right answer as long as aStartTime itself isn't more
	// than about 49 days ago. Note: must cast to int or any negative result will be lost
	// due to DWORD type:
	DWORD tick_now = GetTickCount();
	if (!aAllowEarlyReturn && (int)(aSleepDuration - (tick_now - aStartTime)) > SLEEP_INTERVAL_HALF)
		// Early return isn't allowed and the time remaining is large enough that we need to
		// wait some more (small amounts of remaining time can't be effectively waited for
		// due to the 10ms granularity limit of SetTimer):
		return FAIL; // Tell the caller to wait some more.

	// v1.0.38.04: Reset mLastPeekTime because caller has just done a GetMessage() or PeekMessage(),
	// both of which should have routed events to the keyboard/mouse hooks like LONG_OPERATION_UPDATE's
	// PeekMessage() and thus satisfied the reason that mLastPeekTime is tracked in the first place.
	// UPDATE: Although the hooks now have a dedicated thread, there's a good chance mLastPeekTime is
	// beneficial in terms of increasing GUI & script responsiveness, so it is kept.
	// The following might also improve performance slightly by avoiding extra Peek() calls, while also
	// reducing premature thread interruptions.
	g_script.mLastPeekTime = tick_now;
	return OK;
}



bool CheckScriptTimers()
// Returns true if it launched at least one thread, and false otherwise.
// It's best to call this function only directly from MsgSleep() or when there is an instance of
// MsgSleep() closer on the call stack than the nearest dialog's message pump (e.g. MsgBox).
// This is because threads some events might get queued up for our thread during the execution
// of the timer subroutines here.  When those subroutines finish, if we return directly to a dialog's
// message pump, and such pending messages might be discarded or mishandled.
// Caller should already have checked the value of g_script.mTimerEnabledCount to ensure it's
// greater than zero, since we don't check that here (for performance).
// This function will go through the list of timed subroutines only once and then return to its caller.
// It does it only once so that it won't keep a thread beneath it permanently suspended if the sum
// total of all timer durations is too large to be run at their specified frequencies.
// This function is allowed to be called recursively, which handles certain situations better:
// 1) A hotkey subroutine interrupted and "buried" one of the timer subroutines in the stack.
//    In this case, we don't want all the timers blocked just because that one is, so recursive
//    calls from ExecUntil() are allowed, and they might discover other timers to run.
// 2) If the script is idle but one of the timers winds up taking a long time to execute (perhaps
//    it gets stuck in a long WinWait), we want a recursive call (from MsgSleep() in this example)
//    to launch any other enabled timers concurrently with the first, so that they're not neglected
//    just because one of the timers happens to be long-running.
// Of course, it's up to the user to design timers so that they don't cause problems when they
// interrupted hotkey subroutines, or when they themselves are interrupted by hotkey subroutines
// or other timer subroutines.
{
	// When the following is true, such as during a SendKeys() operation, it seems best not to launch any
	// new timed subroutines.  The reasons for this are similar to the reasons for not allowing hotkeys
	// to fire during such times.  Those reasons are discussed in other comments.  In addition,
	// it seems best as a policy not to allow timed subroutines to run while the script's current
	// quasi-thread is paused.  Doing so would make the tray icon flicker (were it even updated below,
	// which it currently isn't) and in any case is probably not what the user would want.  Most of the
	// time, the user would want all timed subroutines stopped while the current thread is paused.
	// And even if this weren't true, the confusion caused by the subroutines still running even when
	// the current thread is paused isn't worth defaulting to the opposite approach.  In the future,
	// and if there's demand, perhaps a config option can added that allows a different default behavior.
	// UPDATE: It seems slightly better (more consistent) to disallow all timed subroutines whenever
	// there is even one paused thread anywhere in the "stack".
	// v1.0.48: Since g_IdleIsPaused was removed (to simplify a lot of things), g_nPausedThreads now
	// counts the idle thread if it's paused.  Also, to avoid array overflow, g_MaxThreadsTotal must not
	// be exceeded except where otherwise documented.
	// v2.0: g->AllowTimers is ignored when !g_nThreads so that its value can be retained for use as
	// the default for new threads, after the auto-execute thread finishes.
	if (g_nPausedThreads > 0 || (!g->AllowTimers && g_nThreads) || g_nThreads >= g_MaxThreadsTotal || !IsInterruptible()) // See above.
		return false;

	ScriptTimer *ptimer, *next_timer;
	BOOL at_least_one_timer_launched;
	DWORD tick_start;

	// Note: It seems inconsequential if a subroutine that the below loop executes causes a
	// new timer to be added to the linked list while the loop is still enumerating the timers.

	for (at_least_one_timer_launched = FALSE, ptimer = g_script.mFirstTimer
		; ptimer != NULL
		; ptimer = next_timer)
	{
		ScriptTimer &timer = *ptimer; // For performance and convenience.
		if (!timer.mEnabled || timer.mExistingThreads > 0 || timer.mPriority < g->Priority) // thread priorities
		{
			next_timer = timer.mNextTimer;
			continue;
		}

		tick_start = GetTickCount(); // Call GetTickCount() every time in case a previous iteration of the loop took a long time to execute.
		// As of v1.0.36.03, the following subtracts two DWORDs to support intervals of 49.7 vs. 24.8 days.
		// This should work as long as the distance between the values being compared isn't greater than
		// 49.7 days. This is because 1 minus 2 in unsigned math yields 0xFFFFFFFF milliseconds (49.7 days).
		// If the distance between time values is greater than 49.7 days (perhaps because the computer was
		// suspended/hibernated for 50+ days), the next launch of the affected timer(s) will be delayed
		// by up to 100% of their periods.  See IsInterruptible() for more discussion.
		if (tick_start - timer.mTimeLastRun < (DWORD)timer.mPeriod) // Timer is not yet due to run.
		{
			next_timer = timer.mNextTimer;
			continue;
		}
		// Otherwise, this timer is due to run.
		if (!at_least_one_timer_launched) // This will be the first timer launched here.
		{
			at_least_one_timer_launched = TRUE;
			// Since this is the first subroutine that will be launched during this call to
			// this function, we know it will wind up running at least one subroutine, so
			// certain changes are made:
			// Increment the count of quasi-threads only once because this instance of this
			// function will never create more than 1 thread (i.e. if there is more than one
			// enabled timer subroutine, the will always be run sequentially by this instance).
			// If g_nThreads is zero, incrementing it will also effectively mark the script as
			// non-idle, the main consequence being that an otherwise-idle script can be paused
			// if the user happens to do it at the moment a timed subroutine is running, which
			// seems best since some timed subroutines might take a long time to run:
			++g_nThreads; // These are the counterparts the decrements that will be done further
			++g;          // below by ResumeUnderlyingThread().
			// But never kill the main timer, since the mere fact that we're here means that
			// there's at least one enabled timed subroutine.  Though later, performance can
			// be optimized by killing it if there's exactly one enabled subroutine, or if
			// all the subroutines are already in a running state (due to being buried beneath
			// the current quasi-thread).  However, that might introduce unwanted complexity
			// in other places that would need to start up the timer again because we stopped it, etc.
		} // if (!at_least_one_timer_launched)

		// Fix for v1.0.31: mTimeLastRun is now given its new value *before* the thread is launched
		// rather than after.  This allows a timer to be reset by its own thread -- by means of
		// "SetTimer, TimerName", which is otherwise impossible because the reset was being
		// overridden by us here when the thread finished.
		// Seems better to store the start time rather than the finish time, though it's clearly
		// debatable.  The reason is that it's sometimes more important to ensure that a given
		// timed subroutine is *begun* at the specified interval, rather than assuming that
		// the specified interval is the time between when the prior run finished and the new
		// one began.  This should make timers behave more consistently (i.e. how long a timed
		// subroutine takes to run SHOULD NOT affect its *apparent* frequency, which is number
		// of times per second or per minute that we actually attempt to run it):
		timer.mTimeLastRun = tick_start;
		if (timer.mRunOnlyOnce)
			timer.Disable();  // This is done prior to launching the thread for reasons similar to above.

		// v1.0.38.04: The following line is done prior to the timer launch to reduce situations
		// in which a timer thread is interrupted before it can execute even a single line.
		// Search for mLastPeekTime in MsgSleep() for detailed explanation.
		g_script.mLastPeekTime = tick_start; // It's valid to reset this because by definition, "msg" just came in to our caller via Get() or Peek(), both of which qualify as a Peek() for this purpose.

		// This next line is necessary in case a prior iteration of our loop invoked a different
		// timed subroutine that changed any of the global struct's values.  In other words, make
		// every newly launched subroutine start off with the global default values that
		// the user set up in the auto-execute part of the script (e.g. KeyDelay, WinDelay, etc.).
		// Pass false as 3rd param below because ++g_nThreads should be done only once rather than
		// for each Init(), and also it's not necessary to call update the tray icon since timers
		// won't run if there is any paused thread, thus the icon can't currently be showing "paused".
		InitNewThread(timer.mPriority, false, false);

		// This is used to determine which timer SetTimer,,xxx acts on:
		g->CurrentTimer = &timer;

		++timer.mExistingThreads;
		timer.mCallback->ExecuteInNewThread(_T("Timer"));
		--timer.mExistingThreads;

		for (auto *this_timer = &timer; this_timer; this_timer = next_timer)
		{
			// Resolve the next timer only now, in case other timers were created or deleted while
			// this timer was executing.  Must be done before the timer is potentially deleted below.
			// Currently timers are disabled only when they can't be deleted (because they're
			// running).  So now that this one has finished, check if it needs to be deleted.
			if (this_timer->mEnabled || this_timer->mExistingThreads || this_timer->mDeleteLocked)
			{
				if (this_timer == &timer) // If this_timer itself has just executed.
					next_timer = this_timer->mNextTimer;
				//else leave next_timer == this_timer, in case it is ready to execute.
				break;
			}
			next_timer = this_timer->mNextTimer;
			if (next_timer)
				next_timer->mDeleteLocked++; // Prevent next_timer from being deleted.
			// The following call can trigger __delete, which can cause further changes to timers,
			// either directly via SetTimer or via thread interruption.
			g_script.DeleteTimer(this_timer->mCallback->ToObject());
			if (next_timer)
				next_timer->mDeleteLocked--; // Might still be non-zero due to thread interruption.
			// Now also check next_timer, in case it was disabled while __delete was executing.
		} // for() series of timers being deleted.
		
	} // for() each timer.

	if (at_least_one_timer_launched) // Since at least one subroutine was run above, restore various values for our caller.
	{
		ResumeUnderlyingThread();
		return true;
	}
	return false;
}



void PollJoysticks()
// It's best to call this function only directly from MsgSleep() or when there is an instance of
// MsgSleep() closer on the call stack than the nearest dialog's message pump (e.g. MsgBox).
// This is because events posted to the thread indirectly by us here would be discarded or mishandled
// by a non-standard (dialog) message pump.
//
// Polling the joysticks this way rather than using joySetCapture() is preferable for several reasons:
// 1) I believe joySetCapture() internally polls the joystick anyway, via a system timer, so it probably
//    doesn't perform much better (if at all) than polling "manually".
// 2) joySetCapture() only supports 4 buttons;
// 3) joySetCapture() will fail if another app is already capturing the joystick;
// 4) Even if the joySetCapture() succeeds, other programs (e.g. older games), would be prevented from
//    capturing the joystick while the script in question is running.
{
	// Even if joystick hotkeys aren't currently allowed to fire, poll it anyway so that hotkey
	// messages can be buffered for later.
	static DWORD sButtonsPrev[MAX_JOYSTICKS] = {0}; // Set initial state to "all buttons up for all joysticks".
	JOYINFOEX jie;
	DWORD buttons_newly_down;

	for (int i = 0; i < MAX_JOYSTICKS; ++i)
	{
		if (!Hotkey::sJoystickHasHotkeys[i])
			continue;
		// Reset these every time in case joyGetPosEx() ever changes them. Also, most systems have only one joystick,
		// so this code will hardly ever be executed more than once (and sometimes zero times):
		jie.dwSize = sizeof(JOYINFOEX);
		jie.dwFlags = JOY_RETURNBUTTONS; // vs. JOY_RETURNALL
		if (joyGetPosEx(i, &jie) != JOYERR_NOERROR) // Skip this joystick and try the others.
			continue;
		// The exclusive-or operator determines which buttons have changed state.  After that,
		// the bitwise-and operator determines which of those have gone from up to down (the
		// down-to-up events are currently not significant).
		buttons_newly_down = (jie.dwButtons ^ sButtonsPrev[i]) & jie.dwButtons;
		sButtonsPrev[i] = jie.dwButtons;
		if (!buttons_newly_down)
			continue;
		// See if any of the joystick hotkeys match this joystick ID and one of the buttons that
		// has just been pressed on it.  If so, queue up (buffer) the hotkey events so that they will
		// be processed when messages are next checked:
		Hotkey::TriggerJoyHotkeys(i, buttons_newly_down);
	}
}



inline bool MsgMonitor(MsgMonitorInstance &aInstance, HWND aWnd, UINT aMsg, WPARAM awParam, LPARAM alParam, MSG *apMsg, LRESULT &aMsgReply);
bool MsgMonitor(HWND aWnd, UINT aMsg, WPARAM awParam, LPARAM alParam, MSG *apMsg, LRESULT &aMsgReply)
// Returns false if the message is not being monitored, or it is but the called function indicated
// that the message should be given its normal processing.  Returns true when the caller should
// not process this message but should instead immediately reply with aMsgReply (if a reply is possible).
// When false is returned, caller should ignore the value of aMsgReply.
{
	// This function directly launches new threads rather than posting them as something like
	// AHK_GUI_ACTION (which would allow them to be queued by means of MSG_FILTER_MAX) because a message
	// monitor function in the script can return "true" to exempt the message from further processing.
	// Consequently, the MSG_FILTER_MAX queuing effect will only occur for monitored messages that are
	// numerically greater than WM_HOTKEY. Other messages will not be subject to the filter and thus
	// will arrive here even when the script is currently uninterruptible, in which case it seems best
	// for flexibility to allow the interruption (the same is done for CreateCallback).
	//if (!INTERRUPTIBLE_IN_EMERGENCY)
	//	return false;

	bool result = false; // Set default: Tell the caller to give this message any additional/default processing.
	MsgMonitorInstance inst (g_MsgMonitor); // Register this instance so that index can be adjusted by BIF_OnMessage if an item is deleted.

	// Linear search vs. binary search should perform better on average because the vast majority
	// of message monitoring scripts are expected to monitor only a few message numbers.
	for (inst.index = 0; inst.index < inst.count; ++inst.index)
		if (g_MsgMonitor[inst.index].msg == aMsg)
		{
			if (MsgMonitor(inst, aWnd, aMsg, awParam, alParam, apMsg, aMsgReply))
			{
				result = true;
				break;
			}
		}

	return result;
}

bool MsgMonitor(MsgMonitorInstance &aInstance, HWND aWnd, UINT aMsg, WPARAM awParam, LPARAM alParam, MSG *apMsg, LRESULT &aMsgReply)
{
	MsgMonitorStruct *monitor = &g_MsgMonitor[aInstance.index];
	IObject *func = monitor->func; // In case monitor item gets deleted while the function is running (e.g. by the function itself).

	// Many of the things done below are similar to the thread-launch procedure used in MsgSleep(),
	// so maintain them together and see MsgSleep() for more detailed comments.
	if (g_nThreads >= g_MaxThreadsTotal)
		return false;
	if (monitor->instance_count >= monitor->max_instances || g->Priority > 0) // Monitor is already running more than the max number of instances, or existing thread's priority is too high to be interrupted.
		return false;
	// Since above didn't return, the launch of the new thread is now considered unavoidable.

	// See MsgSleep() for comments about the following section.
	InitNewThread(0, false, true);
	DEBUGGER_STACK_PUSH(_T("OnMessage")) // Push a "thread" onto the debugger's stack.  For simplicity and performance, use the function name vs something like "message 0x123".

	// Set last found window (as documented).  Can be NULL.
	// Nested controls like ComboBoxes require more than a simple call to GetParent().
	g->hWndLastUsed = GetNonChildParent(aWnd); // Assign parent window as the last found window (it's ok if it's hidden).
	if (apMsg)
	{
		g->EventInfo = apMsg->time;
	}
	//else leave them at their init-thread defaults.
	
	// v1.0.38.04: Below was added to maximize responsiveness to incoming messages.  The reasoning
	// is similar to why the same thing is done in MsgSleep() prior to its launch of a thread, so see
	// MsgSleep for more comments:
	g_script.mLastPeekTime = GetTickCount();
	++monitor->instance_count;

	// Set up the array of parameters for func->Invoke().
	ExprTokenType param[] =
	{
		(__int64)awParam,
		(__int64)(DWORD_PTR)alParam, // Additional type-cast prevents sign-extension on 32-bit, since LPARAM is signed.
		(__int64)aMsg,
		(__int64)(size_t)aWnd
	};

	ResultType result;
	__int64 retval;

	result = CallMethod(func, func, nullptr, param, _countof(param), &retval);

	bool block_further_processing = (result == EARLY_RETURN);
	if (block_further_processing)
		aMsgReply = (LRESULT)retval;
	//else leave aMsgReply uninitialized because we'll be returning false later below, which tells our caller
	// to ignore aMsgReply.

	DEBUGGER_STACK_POP()

	ResumeUnderlyingThread();

	// Check that the msg monitor still exists (it may have been deleted during the thread that just finished,
	// either by the thread itself or some other thread that interrupted it).  The following cases are possible:
	// 1) This msg monitor is deleted, so g_MsgMonitor[aInstance.index] is either an obsolete array item
	//    or some other msg monitor.  Ensure this item matches before decrementing instance_count.
	// 2) An older msg monitor is deleted, so aInstance.index has been adjusted by BIF_OnMessage
	//    and still points at the correct monitor.
	// 3) This msg monitor is deleted and recreated.  aInstance.index might point to the new monitor,
	//    in which case instance_count is zero and must not be decremented.  If other monitors were also
	//    deleted, aInstance.index might point at a different monitor or an obsolete array item.
	// 4) A newer msg monitor is deleted; nothing needs to be done since this item wasn't affected.
	// 5) Some other msg monitor is created; nothing needs to be done since it's added at the end of the list.
	//
	// UPDATE: We now use a simpler method which flags this specific instance as having been deleted,
	// so if the monitor is deleted and then recreated (with instance_count == 0), there's no chance
	// it will be picked up as the same instance.
	//
	// If "monitor" is defunct due to deletion, decrementing its instance_count is harmless.  However,
	// "monitor" might have been reused by BIF_OnMessage() to create a new msg monitor, so it must be
	// checked to avoid wrongly decrementing some other msg monitor's instance_count.
	if (!aInstance.deleted)
	{
		monitor = &g_MsgMonitor[aInstance.index]; // Retrieve it again in case it was moved.
		if (monitor->instance_count) // Checked for maintainability.  Zero should be impossible due to the "deleted" check.
			--monitor->instance_count;
	}
	else
	{
		// "monitor" is now some other msg-monitor, so don't change it (see comments above).
		aInstance.deleted = false; // Reset for subsequent iterations.
	}
	return block_further_processing; // If false, the caller will ignore aMsgReply and process this message normally. If true, aMsgReply contains the reply the caller should immediately send for this message.
}



void InitNewThread(int aPriority, bool aSkipUninterruptible, bool aIncrementThreadCountAndUpdateTrayIcon
	, bool aIsCritical)
// The value of aIsCritical is ignored when aSkipUninterruptible==true.
// To reduce the expectation that a newly launched hotkey or timed subroutine will
// be immediately interrupted by a timed subroutine or hotkey, interruptions are
// forbidden for a short time (user-configurable).  If the subroutine is a quick one --
// finishing prior to when ExecUntil() or the Timer would have set g_AllowInterruption to be
// true -- we will set it to be true afterward so that it gets done as quickly as possible.
// The following rules of precedence apply:
// If either UninterruptibleTime or UninterruptedLineCountMax is zero, newly launched subroutines
// are always interruptible.  Otherwise: If both are negative, newly launched subroutines are
// never interruptible.  If only one is negative, newly launched subroutines cannot be interrupted
// due to that component, only the other one (which is now known to be positive otherwise the
// first rule of precedence would have applied).
{
	if (aIncrementThreadCountAndUpdateTrayIcon)
	{
		++g_nThreads; // It is the caller's responsibility to avoid calling us if the thread count is too high.
		// Once g_array[0] is used by AutoExec section, it's never used by any other thread because:
		//  1) the auto-execute thread might never finish, in which case it needs to keep consulting the values in g_array[0].
		//  2) it's used to retain the default settings for each newly launched thread.
		++g; 
	}
	// Copy only settings, not state, from the auto-execute thread.
	memcpy(static_cast<ScriptThreadSettings*>(g), static_cast<ScriptThreadSettings*>(&g_default), sizeof(ScriptThreadSettings));

	global_struct &g = *::g; // Must be done AFTER the ++g above. Reduces code size and may improve performance.
	global_clear_state(g);
	g.Priority = aPriority;

	// If the current quasi-thread is paused, the thread we're about to launch will not be, so the tray icon
	// needs to be checked unless the caller said it wasn't needed.  In any case, if the tray icon is already
	// in the right state (which it usually, since paused threads are rare), UpdateTrayIcon() is a very fast call.
	if (aIncrementThreadCountAndUpdateTrayIcon)
		g_script.UpdateTrayIcon(); // Must be done ONLY AFTER updating "g" (e.g, ++g) and/or g->IsPaused.

	if (aSkipUninterruptible)
		return;

	if (!g.ThreadIsCritical)
		g.ThreadIsCritical = aIsCritical;

	if (g_script.mUninterruptibleTime && g_script.mUninterruptedLineCountMax // Both components must be non-zero to start off uninterruptible.
		|| g.ThreadIsCritical) // v1.0.38.04.
	{
		g.AllowThreadToBeInterrupted = false;
		if (!g.ThreadIsCritical)
		{
			if (g_script.mUninterruptibleTime < 0) // A setting of -1 (or any negative) means the thread's uninterruptibility never times out.
				g.UninterruptibleDuration = -1; // "Lock in" the above because for backward compatibility, above is not supposed to affect threads after they're created. Override the default value contained in g_default.
				//g.ThreadStartTime doesn't need to be set when g.UninterruptibleDuration < 0.
			else // It's now known to be >0 (due to various checks above).
			{
				// For backward compatibility, "lock in" the time this thread will become interruptible
				// because that's how previous versions behaved (i.e. 'Thread "Interrupt", NewTimeout'
				// doesn't affect the current thread, only the thread creation behavior in the future).
				// This also makes it more predictable, since AllowThreadToBeInterrupted is only changed
				// when IsInterruptible() is called, which might not happen in between changes to the setting.
				// For explanation of why two fields instead of one are used, see comments in IsInterruptible().
				g.ThreadStartTime = GetTickCount();
				g.UninterruptibleDuration = g_script.mUninterruptibleTime;
			}
		}
		//else g.ThreadIsCritical==true, in which case the values set above won't matter; so they're not set.
	}
	//else g.AllowThreadToBeInterrupted is left at its default of true, in which case the values set
	// above won't matter; so they're not set.
}



void ResumeUnderlyingThread()
{
	if (g->ThrownToken)
		g_script.FreeExceptionToken(g->ThrownToken);


	// The following section handles the switch-over to the former/underlying "g" item:
	--g_nThreads; // Other sections below might rely on this having been done early.
	--g;
	// The below relies on the above having restored "g" to be the global_struct of the underlying thread.

	// If the thread to be resumed was paused and has not been unpaused above, it will automatically be
	// resumed in a paused state because when we return from this function, we should be returning to
	// an instance of ExecUntil() (our caller), which should be in a pause loop still.  Conversely,
	// if the thread to be resumed wasn't paused but was just paused above, the icon will be changed now
	// but the thread won't actually pause until ExecUntil() goes into its pause loop (which should be
	// immediately after the current command finishes, if execution is right in the middle of a command
	// due to the command having done a MsgSleep to allow a thread to interrupt).
	// Older comment: Always update the tray icon in case the paused state of the subroutine
	// we're about to resume is different from our previous paused state.  Do this even
	// when the macro is used by CheckScriptTimers(), which although it might not technically
	// need it, lends maintainability and peace of mind.
	g_script.UpdateTrayIcon();

	if (!g_nThreads)
	{
		// If this was the last running thread and the script has nothing keeping it open (hotkeys, Gui,
		// message monitors, etc.) then it should terminate now:
		if (!g_OnExitIsRunning)
			g_script.ExitIfNotPersistent(EXIT_EXIT);
		// Since above didn't ExitApp(), reset the exit code for when we do eventually ExitApp():
		g_script.mPendingExitCode = 0;
	}
}



BOOL IsInterruptible()
// Generally, anything that checks the value of g->AllowThreadToBeInterrupted should instead call this.
// Indicates whether the current thread should be allowed to be interrupted by a new thread.
// Managing uninterruptibility this way vs. SetTimer()+KillTimer() makes thread-creation speed
// at least 10 times faster (perhaps even 100x if only the creation itself is measured).
// This is because SetTimer() and/or KillTimer() are relatively slow calls, at least on XP.
{
	// Having the two variables g_AllowInterruption and g->AllowThreadToBeInterrupted supports
	// the case where an uninterruptible operation such as SendKeys() happens to occur while
	// g->AllowThreadToBeInterrupted is true, which avoids having to backup and restore
	// g->AllowThreadToBeInterrupted in several places.
	if (!INTERRUPTIBLE_IN_EMERGENCY)
		return FALSE; // Since it's not even emergency-interruptible, it can't be ordinary-interruptible.
	if (!g_nThreads) // Except in the above case, we're always interruptible if no threads are running.
		return TRUE; // This is necessary to allow g_array[0].ThreadIsCritical to hold the default
					 // setting without it preventing threads from launching after auto-exec finishes.
	// Otherwise, update g->AllowThreadToBeInterrupted (if necessary) and use that to determine the final answer.
	// Below are the reasons for having g->UninterruptibleDuration as a field separate from g->ThreadStartTime
	// (i.e. this is why they aren't merged into one called TimeThreadWillBecomeInterruptible):
	// (1) Need some way of indicating "never time out", which is currently done via
	//     UninterruptibleDuration<0. Although that could be done as a bool to reduce global_struct size,
	//     there's also the below.
	// (2) Suspending or hibernating the computer while a thread is uninterruptible could cause the thread to
	//    become semi-permanently uninterruptible.  For example:
	//     - If a thread's uninterruptibility timeout is 60 seconds (or even 60 milliseconds);
	//     - And the computer is suspended/hibernated before the uninterruptibility can time out;
	//     - And the computer is then resumed 25 days later;
	//     - Thread would be wrongly uninterruptible for ~24 days because the single-field method
	//       (TimeThreadWillBecomeInterruptible) would have to subtract that field from GetTickCount()
	//       and view the result as a signed integer.
	// (3) Somewhat similar to #2 above, there may be ways now or in the future for a script to have an
	//     active/current thread but nothing the script does ever requires a call to IsInterruptible()
	//     (currently unlikely since MSG_FILTER_MAX calls it).  This could cause a TickCount to get too
	//     stale and have the same consequences as described in #2 above.
	// OVERALL: Although the method used here can wrongly extend the uninterruptibility of a thread by as
	// much as 100%, since g_script.mUninterruptible time defaults to 17 milliseconds (and in practice is
	// rarely raised beyond 1000) it doesn't matter much in those cases. Even when g_script.mUninterruptible
	// is large such as 20 days, this method can be off by no more than 20 days, which isn't too bad
	// in percentage terms compared to the alternative, which could cause a timeout of 15 milliseconds to
	// increase to 24 days.  Windows Vista and beyond have a 64-bit tickcount available, so that may be of
	// use in future versions (hopefully it performs nearly as well as GetTickCount()).
	// v2.0: g->UninterruptedLineCount is checked to ensure that each thread is allowed to execute at least
	// one line before being interrupted.  This was proven necessary for g->UninterruptibleDuration <= 16
	// because GetTickCount() updates in increments of 15 or 16 and therefore 16 can be virtually no time
	// at all.  It might also be needed for larger values if the system is busy.
	if (   !g->AllowThreadToBeInterrupted // Those who check whether g->AllowThreadToBeInterrupted==false should then check whether it should be made true.
		&& g->UninterruptibleDuration > -1 // Must take precedence over the below.  g_script.mUninterruptibleTime is not checked because it's supposed to go into effect during thread creation, not after the thread is running and has possibly changed the timeout via 'Thread "Interrupt"'.
		&& (DWORD)(GetTickCount()- g->ThreadStartTime) >= (DWORD)g->UninterruptibleDuration // See big comment section above.
		&& g->UninterruptedLineCount // In case of "Critical" on the first line.  See v2.0 comment above.
		)
		// Once the thread becomes interruptible by any means, g->ThreadStartTime/UninterruptibleDuration
		// can never matter anymore because only Critical (never "Thread Interrupt") can turn off the
		// interruptibility again, and it resets g->UninterruptibleDuration.
		g->AllowThreadToBeInterrupted = true; // Avoids issues with 49.7 day limit of 32-bit TickCount, and also helps performance future callers of this function (they can skip most of the checking above).
	//else g->AllowThreadToBeInterrupted is already up-to-date.
	return (BOOL)g->AllowThreadToBeInterrupted;
}



VOID CALLBACK MsgBoxTimeout(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	// Unfortunately, it appears that MessageBox() will return zero rather
	// than AHK_TIMEOUT, specified below -- at least under WinXP.  This
	// makes it impossible to distinguish between a MessageBox() that's been
	// timed out (destroyed) by this function and one that couldn't be
	// created in the first place due to some other error.  But since
	// MessageBox() errors are rare, we assume that they timed out if
	// the MessageBox() returns 0.  UPDATE: Due to the fact that TimerProc()'s
	// are called via WM_TIMER messages in our msg queue, make sure that the
	// window really exists before calling EndDialog(), because if it doesn't,
	// chances are that EndDialog() was already called with another value.
	// UPDATE #2: Actually that isn't strictly needed because MessageBox()
	// ignores the AHK_TIMEOUT value we send here.  But it feels safer:
	if (IsWindow(hWnd))
		EndDialog(hWnd, AHK_TIMEOUT);
	KillTimer(hWnd, idEvent);
	// v1.0.33: The following was added to fix the fact that a MsgBox with only an OK button
	// does not actually send back the code sent by EndDialog() above.  The HWND is checked
	// in case "g" is no longer the original thread due to another thread having interrupted it.
	// v1.1.30.01: The loop was added so that the timeout can be detected even if the thread
	// which owns the dialog was interrupted.
	for (auto *dialog_g = g; dialog_g >= g_array; --dialog_g)
		if (dialog_g->DialogHWND == hWnd) // Regardless of whether IsWindow() is true.
		{
			dialog_g->MsgBoxTimedOut = true;
			break;
		}
}



VOID CALLBACK InputTimeout(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	int new_timer_period = 0;
	for (auto *input = g_input; input; input = input->Prev)
	{
		if (input->Timeout && input->InProgress())
		{
			int time_left = int(input->TimeoutAt - dwTime);
			if (time_left <= 0)
				input->EndByTimeout();
			else if (time_left < new_timer_period || !new_timer_period)
				new_timer_period = time_left;
		}
	}
	if (new_timer_period != 0)
		SET_INPUT_TIMER(new_timer_period, dwTime + new_timer_period)
	else
		KILL_INPUT_TIMER
}



VOID CALLBACK RefreshInterruptibility(HWND hWnd, UINT uMsg, UINT_PTR idEvent, DWORD dwTime)
{
	IsInterruptible(); // Search on RefreshInterruptibility for comments.
}
