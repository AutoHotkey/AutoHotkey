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
#include "keyboard_mouse.h"
#include "globaldata.h" // for g->KeyDelay
#include "application.h" // for MsgSleep()
#include "util.h"  // for strlicmp()
#include "window.h" // for IsWindowHung()


// Added for v1.0.25.  Search on sPrevEventType for more comments:
static KeyEventTypes sPrevEventType;
static vk_type sPrevVK = 0;
// For v1.0.25, the below is static to track it in between sends, so that the below will continue
// to work:
// Send {LWinDown}
// Send {LWinUp}  ; Should still open the Start Menu even though it's a separate Send.
static vk_type sPrevEventModifierDown = 0;
static modLR_type sModifiersLR_persistent = 0; // Tracks this script's own lifetime/persistent modifiers (the ones it caused to be persistent and thus is responsible for tracking).

// v1.0.44.03: Below supports multiple keyboard layouts better by having script adapt to active window's layout.
#define MAX_CACHED_LAYOUTS 10  // Hard to imagine anyone using more languages/layouts than this, but even if they do it will still work; performance would just be a little worse due to being uncached.
static CachedLayoutType sCachedLayout[MAX_CACHED_LAYOUTS] = {{0}};
static HKL sTargetKeybdLayout;           // Set by SendKeys() for use by the functions it calls directly and indirectly.
static ResultType sTargetLayoutHasAltGr; //

// v1.0.43: Support for SendInput() and journal-playback hook:
#define MAX_INITIAL_EVENTS_SI 500UL  // sizeof(INPUT) == 28 as of 2006. Since Send is called so often, and since most Sends are short, reducing the load on the stack is also a deciding factor for these.
#define MAX_INITIAL_EVENTS_PB 1500UL // sizeof(PlaybackEvent) == 8, so more events are justified before resorting to malloc().
static LPINPUT sEventSI;        // No init necessary.  An array that's allocated/deallocated by SendKeys().
static PlaybackEvent *&sEventPB = (PlaybackEvent *&)sEventSI;
static UINT sEventCount, sMaxEvents; // Number of items in the above arrays and the current array capacity.
static UINT sCurrentEvent;
static modLR_type sEventModifiersLR; // Tracks the modifier state to following the progress/building of the SendInput array.
static POINT sSendInputCursorPos;    // Tracks/predicts cursor position as SendInput array is built.
static HookType sHooksToRemoveDuringSendInput;
static SendModes sSendMode = SM_EVENT; // Whether a SendInput or Hook array is currently being constructed.
static bool sAbortArraySend;         // No init needed.
static bool sFirstCallForThisEvent;  //
static bool sInBlindMode;            //
static DWORD sThisEventTime;         //

// Dynamically resolve SendInput() because otherwise the app won't launch at all on Windows 95/NT-pre-SP3:
typedef UINT (WINAPI *MySendInputType)(UINT, LPINPUT, int);
static MySendInputType sMySendInput = (MySendInputType)GetProcAddress(GetModuleHandle(_T("user32")), "SendInput");
// Above will be NULL for Win95/NT-pre-SP3.


void DisguiseWinAltIfNeeded(vk_type aVK)
// For v1.0.25, the following situation is fixed by the code below: If LWin or LAlt
// becomes a persistent modifier (e.g. via Send {LWin down}) and the user physically
// releases LWin immediately before: 1) the {LWin up} is scheduled; and 2) SendKey()
// returns.  Then SendKey() will push the modifier back down so that it is in effect
// for other things done by its caller (SendKeys) and also so that if the Send
// operation ends, the key will still be down as the user intended (to modify future
// keystrokes, physical or simulated).  However, since that down-event is followed
// immediately by an up-event, the Start Menu appears for WIN-key or the active
// window's menu bar is activated for ALT-key.  SOLUTION: Disguise Win-up and Alt-up
// events in these cases.  This workaround has been successfully tested.  It's also
// limited is scope so that a script can still explicitly invoke the Start Menu with
// "Send {LWin}", or activate the menu bar with "Send {Alt}".
// The check of sPrevEventModifierDown allows "Send {LWinDown}{LWinUp}" etc., to
// continue to work.
// v1.0.40: For maximum flexibility and minimum interference while in blind mode,
// don't disguise Win and Alt keystrokes then.
{
	// Caller has ensured that aVK is about to have a key-up event, so if the event immediately
	// prior to this one is a key-down of the same type of modifier key, it's our job here
	// to send the disguising keystrokes now (if appropriate).
	if (sPrevEventType == KEYDOWN && sPrevEventModifierDown != aVK && !sInBlindMode
		// SendPlay mode can't display Start Menu, so no need for disguise keystrokes (such keystrokes might cause
		// unwanted effects in certain games):
		&& ((aVK == VK_LWIN || aVK == VK_RWIN) && (sPrevVK == VK_LWIN || sPrevVK == VK_RWIN) && sSendMode != SM_PLAY
			|| (aVK == VK_LMENU || (aVK == VK_RMENU && sTargetLayoutHasAltGr != CONDITION_TRUE)) && (sPrevVK == VK_LMENU || sPrevVK == VK_RMENU)))
		KeyEvent(KEYDOWNANDUP, g_MenuMaskKey); // Disguise it to suppress Start Menu or prevent activation of active window's menu bar.
}



// moved from SendKeys
void SendUnicodeChar(wchar_t aChar, int aModifiers = -1)
{
	// Set modifier keystate for consistent results. If not specified by caller, default to releasing
	// Alt/Ctrl/Shift since these are known to interfere in some cases, and because SendAsc() does it
	// (except for LAlt). Leave LWin/RWin as they are, for consistency with SendAsc().
	if (aModifiers == -1)
	{
		aModifiers = sSendMode ? sEventModifiersLR : GetModifierLRState();
		aModifiers &= ~(MOD_LALT | MOD_RALT | MOD_LCONTROL | MOD_RCONTROL | MOD_LSHIFT | MOD_RSHIFT);
	}
	SetModifierLRState((modLR_type)aModifiers, sSendMode ? sEventModifiersLR : GetModifierLRState(), NULL, false, true, KEY_IGNORE);

	if (sSendMode == SM_INPUT)
	{
		// Calling SendInput() now would cause characters to appear out of sequence.
		// Instead, put them into the array and allow them to be sent in sequence.
		PutKeybdEventIntoArray(0, 0, aChar, KEYEVENTF_UNICODE, KEY_IGNORE_LEVEL(g->SendLevel));
		PutKeybdEventIntoArray(0, 0, aChar, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP, KEY_IGNORE_LEVEL(g->SendLevel));
		return;
	}
	//else caller has ensured sSendMode is SM_EVENT. In that mode, events are sent one at a time,
	// so it is safe to immediately call SendInput(). SM_PLAY is not supported; for simplicity,
	// SendASC() is called instead of this function. Although this means Unicode chars probably
	// won't work, it seems better than sending chars out of order. One possible alternative could
	// be to "flush" the event array, but since SendInput and SendEvent are probably much more common,
	// this is left for a future version.

	INPUT u_input[2];

	u_input[0].type = INPUT_KEYBOARD;
	u_input[0].ki.wVk = 0;
	u_input[0].ki.wScan = aChar;
	u_input[0].ki.dwFlags = KEYEVENTF_UNICODE;
	u_input[0].ki.time = 0;
	// L25: Set dwExtraInfo to ensure AutoHotkey ignores the event; otherwise it may trigger a SCxxx hotkey (where xxx is u_code).
	u_input[0].ki.dwExtraInfo = KEY_IGNORE_LEVEL(g->SendLevel);
	
	u_input[1].type = INPUT_KEYBOARD;
	u_input[1].ki.wVk = 0;
	u_input[1].ki.wScan = aChar;
	u_input[1].ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
	u_input[1].ki.time = 0;
	u_input[1].ki.dwExtraInfo = KEY_IGNORE_LEVEL(g->SendLevel);

	SendInput(2, u_input, sizeof(INPUT));
}



void SendKeys(LPTSTR aKeys, bool aSendRaw, SendModes aSendModeOrig, HWND aTargetWindow)
// The aKeys string must be modifiable (not constant), since for performance reasons,
// it's allowed to be temporarily altered by this function.  mThisHotkeyModifiersLR, if non-zero,
// should be the set of modifiers used to trigger the hotkey that called the subroutine
// containing the Send that got us here.  If any of those modifiers are still down,
// they will be released prior to sending the batch of keys specified in <aKeys>.
// v1.0.43: aSendModeOrig was added.
{
	if (!*aKeys)
		return;
	global_struct &g = *::g; // Reduces code size and may improve performance.

	// For performance and also to reserve future flexibility, recognize {Blind} only when it's the first item
	// in the string.
	if (sInBlindMode = !aSendRaw && !_tcsnicmp(aKeys, _T("{Blind}"), 7)) // Don't allow {Blind} while in raw mode due to slight chance {Blind} is intended to be sent as a literal string.
		// Blind Mode (since this seems too obscure to document, it's mentioned here):  Blind Mode relies
		// on modifiers already down for something like ^c because ^c is saying "manifest a ^c", which will
		// happen if ctrl is already down.  By contrast, Blind does not release shift to produce lowercase
		// letters because avoiding that adds flexibility that couldn't be achieved otherwise.
		// Thus, ^c::Send {Blind}c produces the same result when ^c is substituted for the final c.
		// But Send {Blind}{LControl down} will generate the extra events even if ctrl already down.
		aKeys += 7; // Remove "{Blind}" from further consideration.

	int orig_key_delay = g.KeyDelay;
	int orig_press_duration = g.PressDuration;
	if (aSendModeOrig == SM_INPUT || aSendModeOrig == SM_INPUT_FALLBACK_TO_PLAY) // Caller has ensured aTargetWindow==NULL for SendInput and SendPlay modes.
	{
		// Both of these modes fall back to a different mode depending on whether some other script
		// is running with a keyboard/mouse hook active.  Of course, the detection of this isn't foolproof
		// because older versions of AHK may be running and/or other apps with LL keyboard hooks. It's
		// just designed to add a lot of value for typical usage because SendInput is preferred due to it
		// being considerably faster than SendPlay, especially for long replacements when the CPU is under
		// heavy load.
		if (   !sMySendInput // Win95/NT-pre-SP3 don't support SendInput, so fall back to the specified mode.
			|| SystemHasAnotherKeybdHook() // This function has been benchmarked to ensure it doesn't yield our timeslice, etc.  200 calls take 0ms according to tick-count, even when CPU is maxed.
			|| !aSendRaw && SystemHasAnotherMouseHook() && tcscasestr(aKeys, _T("{Click"))   ) // Ordered for short-circuit boolean performance.  v1.0.43.09: Fixed to be strcasestr vs. !strcasestr
		{
			// Need to detect in advance what type of array to build (for performance and code size).  That's why
			// it done this way, and here are the comments about it:
			// strcasestr() above has an unwanted amount of overhead if aKeys is huge, but it seems acceptable
			// because it's called only when system has another mouse hook but *not* another keybd hook (very rare).
			// Also, for performance reasons, {LButton and such are not checked for, which is documented and seems
			// justified because the new {Click} method is expected to become prevalent, especially since this
			// whole section only applies when the new SendInput mode is in effect.
			// Finally, checking aSendRaw isn't foolproof because the string might contain {Raw} prior to {Click,
			// but the complexity and performance of checking for that seems unjustified given the rarity,
			// especially since there are almost never any consequences to reverting to hook mode vs. SendInput.
			if (aSendModeOrig == SM_INPUT_FALLBACK_TO_PLAY)
				aSendModeOrig = SM_PLAY;
			else // aSendModeOrig == SM_INPUT, so fall back to EVENT.
			{
				aSendModeOrig = SM_EVENT;
				// v1.0.43.08: When SendInput reverts to SendEvent mode, the majority of users would want
				// a fast sending rate that is more comparable to SendInput's speed that the default KeyDelay
				// of 10ms.  PressDuration may be generally superior to KeyDelay because it does a delay after
				// each changing of modifier state (which tends to improve reliability for certain apps).
				// The following rules seem likely to be the best benefit in terms of speed and reliability:
				// KeyDelay 0+,-1+ --> -1, 0
				// KeyDelay -1, 0+ --> -1, 0
				// KeyDelay -1,-1 --> -1, -1
				g.PressDuration = (g.KeyDelay < 0 && g.PressDuration < 0) ? -1 : 0;
				g.KeyDelay = -1; // Above line must be done before this one.
			}
		}
		else // SendInput is available and no other impacting hooks are obviously present on the system, so use SendInput unconditionally.
			aSendModeOrig = SM_INPUT; // Resolve early so that other sections don't have to consider SM_INPUT_FALLBACK_TO_PLAY a valid value.
	}

	// Might be better to do this prior to changing capslock state.  UPDATE: In v1.0.44.03, the following section
	// has been moved to the top of the function because:
	// 1) For ControlSend, GetModifierLRState() might be more accurate if the threads are attached beforehand.
	// 2) Determines sTargetKeybdLayout and sTargetLayoutHasAltGr early (for maintainability).
	bool threads_are_attached = false; // Set default.
	DWORD keybd_layout_thread = 0;     //
	DWORD target_thread; // Doesn't need init.
	if (aTargetWindow) // Caller has ensured this is NULL for SendInput and SendPlay modes.
	{
		if ((target_thread = GetWindowThreadProcessId(aTargetWindow, NULL)) // Assign.
			&& target_thread != g_MainThreadID && !IsWindowHung(aTargetWindow))
		{
			threads_are_attached = AttachThreadInput(g_MainThreadID, target_thread, TRUE) != 0;
			keybd_layout_thread = target_thread; // Testing shows that ControlSend benefits from the adapt-to-layout technique too.
		}
		//else no target thread, or it's our thread, or it's hung; so keep keybd_layout_thread at its default.
	}
	else
	{
		// v1.0.48.01: On Vista or later, work around the fact that an "L" keystroke (physical or artificial) will
		// lock the computer whenever either Windows key is physically pressed down (artificially releasing the
		// Windows key isn't enough to solve it because Win+L is apparently detected aggressively like
		// Ctrl-Alt-Delete.  Unlike the handling of SM_INPUT in another section, this one here goes into
		// effect for all Sends because waiting for an "L" keystroke to be sent would be too late since the
		// Windows would have already been artificially released by then, so IsKeyDownAsync() wouldn't be
		// able to detect when the user physically releases the key.
		if (   (g_script.mThisHotkeyModifiersLR & (MOD_LWIN|MOD_RWIN)) // Limit the scope to only those hotkeys that have a Win modifier, since anything outside that scope hasn't been fully analyzed.
			&& (GetTickCount() - g_script.mThisHotkeyStartTime) < (DWORD)50 // Ensure g_script.mThisHotkeyModifiersLR is up-to-date enough to be reliable.
			&& aSendModeOrig != SM_PLAY // SM_PLAY is reported to be incapable of locking the computer.
			&& !sInBlindMode // The philosophy of blind-mode is that the script should have full control, so don't do any waiting during blind mode.
			&& g_os.IsWinVistaOrLater() // Only Vista (and presumably later OSes) check the physical state of the Windows key for Win+L.
			&& GetCurrentThreadId() == g_MainThreadID // Exclude the hook thread because it isn't allowed to call anything like MsgSleep, nor are any calls from the hook thread within the understood/analyzed scope of this workaround.
			)
		{
			bool wait_for_win_key_release;
			if (aSendRaw)
				wait_for_win_key_release = StrChrAny(aKeys, _T("Ll")) != NULL;
			else
			{
				// It seems worthwhile to scan for any "L" characters to avoid waiting for the release
				// of the Windows key when there are no L's.  For performance and code size, the check
				// below isn't comprehensive (e.g. it fails to consider things like {L} and #L).
				// Although RegExMatch() could be used instead of the below, that would use up one of
				// the RegEx cache entries, plus it would probably perform worse.  So scan manually.
				LPTSTR L_pos, brace_pos;
				for (wait_for_win_key_release = false, brace_pos = aKeys; L_pos = StrChrAny(brace_pos, _T("Ll"));)
				{
					// Encountering a #L seems too rare, and the consequences too mild (or nonexistent), to
					// justify the following commented-out section:
					//if (L_pos > aKeys && L_pos[-1] == '#') // A simple check; it won't detect things like #+L.
					//	brace_pos = L_pos + 1;
					//else
					if (!(brace_pos = StrChrAny(L_pos + 1, _T("{}"))) || *brace_pos == '{') // See comment below.
					{
						wait_for_win_key_release = true;
						break;
					}
					//else it found a '}' without a preceding '{', which means this "L" is inside braces.
					// For simplicity, ignore such L's (probably not a perfect check, but seems worthwhile anyway).
				}
			}
			if (wait_for_win_key_release)
				while (IsKeyDownAsync(VK_LWIN) || IsKeyDownAsync(VK_RWIN)) // Even if the keyboard hook is installed, it seems best to use IsKeyDownAsync() vs. g_PhysicalKeyState[] because it's more likely to produce consistent behavior.
					SLEEP_WITHOUT_INTERRUPTION(INTERVAL_UNSPECIFIED); // Seems best not to allow other threads to launch, for maintainability and because SendKeys() isn't designed to be interruptible.
		}

		// v1.0.44.03: The following change is meaningful only to people who use more than one keyboard layout.
		// It seems that the vast majority of them would want the Send command (as well as other features like
		// Hotstrings and the Input command) to adapt to the keyboard layout of the active window (or target window
		// in the case of ControlSend) rather than sticking with the script's own keyboard layout.  In addition,
		// testing shows that this adapt-to-layout method costs almost nothing in performance, especially since
		// the active window, its thread, and its layout are retrieved only once for each Send rather than once
		// for each keystroke.
		HWND active_window;
		if (active_window = GetForegroundWindow())
			keybd_layout_thread = GetWindowThreadProcessId(active_window, NULL);
		//else no foreground window, so keep keybd_layout_thread at default.
	}
	sTargetKeybdLayout = GetKeyboardLayout(keybd_layout_thread); // If keybd_layout_thread==0, this will get our thread's own layout, which seems like the best/safest default.
	sTargetLayoutHasAltGr = LayoutHasAltGr(sTargetKeybdLayout);  // Note that WM_INPUTLANGCHANGEREQUEST is not monitored by MsgSleep for the purpose of caching our thread's keyboard layout.  This is because it would be unreliable if another msg pump such as MsgBox is running.  Plus it hardly helps perf. at all, and hurts maintainability.

	// Below is now called with "true" so that the hook's modifier state will be corrected (if necessary)
	// prior to every send.
	modLR_type mods_current = GetModifierLRState(true); // Current "logical" modifier state.

	// Make a best guess of what the physical state of the keys is prior to starting (there's no way
	// to be certain without the keyboard hook). Note: We only want those physical
	// keys that are also logically down (it's possible for a key to be down physically
	// but not logically such as when R-control, for example, is a suffix hotkey and the
	// user is physically holding it down):
	modLR_type mods_down_physically_orig, mods_down_physically_and_logically
		, mods_down_physically_but_not_logically_orig;
	if (g_KeybdHook)
	{
		// Since hook is installed, use its more reliable tracking to determine which
		// modifiers are down.
		mods_down_physically_orig = g_modifiersLR_physical;
		mods_down_physically_and_logically = g_modifiersLR_physical & g_modifiersLR_logical; // intersect
		mods_down_physically_but_not_logically_orig = g_modifiersLR_physical & ~g_modifiersLR_logical;
	}
	else // Use best-guess instead.
	{
		// Even if TickCount has wrapped due to system being up more than about 49 days,
		// DWORD subtraction still gives the right answer as long as g_script.mThisHotkeyStartTime
		// itself isn't more than about 49 days ago:
		if ((GetTickCount() - g_script.mThisHotkeyStartTime) < (DWORD)g_HotkeyModifierTimeout) // Elapsed time < timeout-value
			mods_down_physically_orig = mods_current & g_script.mThisHotkeyModifiersLR; // Bitwise AND is set intersection.
		else
			// Since too much time as passed since the user pressed the hotkey, it seems best,
			// based on the action that will occur below, to assume that no hotkey modifiers
			// are physically down:
			mods_down_physically_orig = 0;
		mods_down_physically_and_logically = mods_down_physically_orig;
		mods_down_physically_but_not_logically_orig = 0; // There's no way of knowing, so assume none.
	}

	// Any of the external modifiers that are down but NOT due to the hotkey are probably
	// logically down rather than physically (perhaps from a prior command such as
	// "Send, {CtrlDown}".  Since there's no way to be sure without the keyboard hook or some
	// driver-level monitoring, it seems best to assume that
	// they are logically vs. physically down.  This value contains the modifiers that
	// we will not attempt to change (e.g. "Send, A" will not release the LWin
	// before sending "A" if this value indicates that LWin is down).  The below sets
	// the value to be all the down-keys in mods_current except any that are physically
	// down due to the hotkey itself.  UPDATE: To improve the above, we now exclude from
	// the set of persistent modifiers any that weren't made persistent by this script.
	// Such a policy seems likely to do more good than harm as there have been cases where
	// a modifier was detected as persistent just because #HotkeyModifier had timed out
	// while the user was still holding down the key, but then when the user released it,
	// this logic here would think it's still persistent and push it back down again
	// to enforce it as "always-down" during the send operation.  Thus, the key would
	// basically get stuck down even after the send was over:
	sModifiersLR_persistent &= mods_current & ~mods_down_physically_and_logically;
	modLR_type persistent_modifiers_for_this_SendKeys, extra_persistent_modifiers_for_blind_mode;
	if (sInBlindMode)
	{
		// The following value is usually zero unless the user is currently holding down
		// some modifiers as part of a hotkey. These extra modifiers are the ones that
		// this send operation (along with all its calls to SendKey and similar) should
		// consider to be down for the duration of the Send (unless they go up via an
		// explicit {LWin up}, etc.)
		extra_persistent_modifiers_for_blind_mode = mods_current & ~sModifiersLR_persistent;
		persistent_modifiers_for_this_SendKeys = mods_current;
	}
	else
	{
		extra_persistent_modifiers_for_blind_mode = 0;
		persistent_modifiers_for_this_SendKeys = sModifiersLR_persistent;
	}
	// Above:
	// Keep sModifiersLR_persistent and persistent_modifiers_for_this_SendKeys in sync with each other from now on.
	// By contrast to persistent_modifiers_for_this_SendKeys, sModifiersLR_persistent is the lifetime modifiers for
	// this script that stay in effect between sends.  For example, "Send {LAlt down}" leaves the alt key down
	// even after the Send ends, by design.
	//
	// It seems best not to change persistent_modifiers_for_this_SendKeys in response to the user making physical
	// modifier changes during the course of the Send.  This is because it seems more often desirable that a constant
	// state of modifiers be kept in effect for the entire Send rather than having the user's release of a hotkey
	// modifier key, which typically occurs at some unpredictable time during the Send, to suddenly alter the nature
	// of the Send in mid-stride.  Another reason is to make the behavior of Send consistent with that of SendInput.

	// The default behavior is to turn the capslock key off prior to sending any keys
	// because otherwise lowercase letters would come through as uppercase and vice versa.
	ToggleValueType prior_capslock_state;
	if (threads_are_attached || !g_os.IsWin9x())
	{
		// Only under either of the above conditions can the state of Capslock be reliably
		// retrieved and changed.  Remember that apps like MS Word have an auto-correct feature that
		// might make it wrongly seem that the turning off of Capslock below needs a Sleep(0) to take effect.
		prior_capslock_state = g.StoreCapslockMode && !sInBlindMode
			? ToggleKeyState(VK_CAPITAL, TOGGLED_OFF)
			: TOGGLE_INVALID; // In blind mode, don't do store capslock (helps remapping and also adds flexibility).
	}
	else // OS is Win9x and threads are not attached.
	{
		// Attempt to turn off capslock, but never attempt to turn it back on because we can't
		// reliably detect whether it was on beforehand.  Update: This didn't do any good, so
		// it's disabled for now:
		//CapslockOffWin9x();
		prior_capslock_state = TOGGLE_INVALID;
	}

	// sSendMode must be set only after setting Capslock state above, because the hook method
	// is incapable of changing the on/off state of toggleable keys like Capslock.
	// However, it can change Capslock state as seen in the window to which playback events are being
	// sent; but the behavior seems inconsistent and might vary depending on OS type, so it seems best
	// not to rely on it.
	sSendMode = aSendModeOrig;
	if (sSendMode) // Build an array.  We're also responsible for setting sSendMode to SM_EVENT prior to returning.
	{
		size_t mem_size;
		if (sSendMode == SM_INPUT)
		{
			mem_size = MAX_INITIAL_EVENTS_SI * sizeof(INPUT);
			sMaxEvents = MAX_INITIAL_EVENTS_SI;
		}
		else // Playback type.
		{
			mem_size = MAX_INITIAL_EVENTS_PB * sizeof(PlaybackEvent);
			sMaxEvents = MAX_INITIAL_EVENTS_PB;
		}
		// _alloca() is used to avoid the overhead of malloc/free (99% of Sends will thus fit in stack memory).
		// _alloca() never returns a failure code, it just raises an exception (e.g. stack overflow).
		InitEventArray(_alloca(mem_size), sMaxEvents, mods_current);
	}

	bool blockinput_prev = g_BlockInput;
	bool do_selective_blockinput = (g_BlockInputMode == TOGGLE_SEND || g_BlockInputMode == TOGGLE_SENDANDMOUSE)
		&& !sSendMode && !aTargetWindow && g_os.IsWinNT4orLater();
	if (do_selective_blockinput)
		Line::ScriptBlockInput(true); // Turn it on unconditionally even if it was on, since Ctrl-Alt-Del might have disabled it.

	vk_type vk;
	sc_type sc;
	modLR_type key_as_modifiersLR = 0;
	modLR_type mods_for_next_key = 0;
	// Above: For v1.0.35, it was changed to modLR vs. mod so that AltGr keys such as backslash and '{'
	// are supported on layouts such as German when sending to apps such as Putty that are fussy about
	// which ALT key is held down to produce the character.
	vk_type this_event_modifier_down;
	size_t key_text_length, key_name_length;
	TCHAR *end_pos, *space_pos, *next_word, old_char, single_char_string[2];
	KeyEventTypes event_type;
	int repeat_count, click_x, click_y;
	bool move_offset, key_down_is_persistent;
	DWORD placeholder;
	single_char_string[1] = '\0'; // Terminate in advance.

	LONG_OPERATION_INIT  // Needed even for SendInput/Play.

	for (; *aKeys; ++aKeys, sPrevEventModifierDown = this_event_modifier_down)
	{
		this_event_modifier_down = 0; // Set default for this iteration, overridden selectively below.
		if (!sSendMode)
			LONG_OPERATION_UPDATE_FOR_SENDKEYS // This does not measurably affect the performance of SendPlay/Event.

		if (!aSendRaw && _tcschr(_T("^+!#{}"), *aKeys))
		{
			switch (*aKeys)
			{
			case '^':
				if (!(persistent_modifiers_for_this_SendKeys & (MOD_LCONTROL|MOD_RCONTROL)))
					mods_for_next_key |= MOD_LCONTROL;
				// else don't add it, because the value of mods_for_next_key may also used to determine
				// which keys to release after the key to which this modifier applies is sent.
				// We don't want persistent modifiers to ever be released because that's how
				// AutoIt2 behaves and it seems like a reasonable standard.
				continue;
			case '+':
				if (!(persistent_modifiers_for_this_SendKeys & (MOD_LSHIFT|MOD_RSHIFT)))
					mods_for_next_key |= MOD_LSHIFT;
				continue;
			case '!':
				if (!(persistent_modifiers_for_this_SendKeys & (MOD_LALT|MOD_RALT)))
					mods_for_next_key |= MOD_LALT;
				continue;
			case '#':
				if (!(persistent_modifiers_for_this_SendKeys & (MOD_LWIN|MOD_RWIN)))
					mods_for_next_key |= MOD_LWIN;
				continue;
			case '}': continue;  // Important that these be ignored.  Be very careful about changing this, see below.
			case '{':
			{
				if (   !(end_pos = _tcschr(aKeys + 1, '}'))   ) // Ignore it and due to rarity, don't reset mods_for_next_key.
					continue; // This check is relied upon by some things below that assume a '}' is present prior to the terminator.
				aKeys = omit_leading_whitespace(aKeys + 1); // v1.0.43: Skip leading whitespace inside the braces to be more flexible.
				if (   !(key_text_length = end_pos - aKeys)   )
				{
					if (end_pos[1] == '}')
					{
						// The literal string "{}}" has been encountered, which is interpreted as a single "}".
						++end_pos;
						key_text_length = 1;
					}
					else if (IS_SPACE_OR_TAB(end_pos[1])) // v1.0.48: Support "{} down}", "{} downtemp}" and "{} up}".
					{
						next_word = omit_leading_whitespace(end_pos + 1);
						if (   !_tcsnicmp(next_word, _T("Down"), 4) // "Down" or "DownTemp" (or likely enough).
							|| !_tcsnicmp(next_word, _T("Up"), 2)   )
						{
							if (   !(end_pos = _tcschr(next_word, '}'))   ) // See comments at similar section above.
								continue;
							key_text_length = end_pos - aKeys; // This result must be non-zero due to the checks above.
						}
						else
							goto brace_case_end;  // The loop's ++aKeys will now skip over the '}', ignoring it.
					}
					else // Empty braces {} were encountered (or all whitespace, but literal whitespace isn't sent).
						goto brace_case_end;  // The loop's ++aKeys will now skip over the '}', ignoring it.
				}

				if (!_tcsnicmp(aKeys, _T("Click"), 5))
				{
					*end_pos = '\0';  // Temporarily terminate the string here to omit the closing brace from consideration below.
					ParseClickOptions(omit_leading_whitespace(aKeys + 5), click_x, click_y, vk
						, event_type, repeat_count, move_offset);
					*end_pos = '}';  // Undo temp termination.
					if (repeat_count < 1) // Allow {Click 100, 100, 0} to do a mouse-move vs. click (but modifiers like ^{Click..} aren't supported in this case.
						MouseMove(click_x, click_y, placeholder, g.DefaultMouseSpeed, move_offset);
					else // Use SendKey because it supports modifiers (e.g. ^{Click}) SendKey requires repeat_count>=1.
						SendKey(vk, 0, mods_for_next_key, persistent_modifiers_for_this_SendKeys
							, repeat_count, event_type, 0, aTargetWindow, click_x, click_y, move_offset);
					goto brace_case_end; // This {} item completely handled, so move on to next.
				}
				else if (!_tcsnicmp(aKeys, _T("Raw"), 3)) // This is used by auto-replace hotstrings too.
				{
					// As documented, there's no way to switch back to non-raw mode afterward since there's no
					// correct way to support special (non-literal) strings such as {Raw Off} while in raw mode.
					aSendRaw = true;
					goto brace_case_end; // This {} item completely handled, so move on to next.
				}

				// Since above didn't "goto", this item isn't {Click}.
				event_type = KEYDOWNANDUP;         // Set defaults.
				repeat_count = 1;                  //
				key_name_length = key_text_length; //
				*end_pos = '\0';  // Temporarily terminate the string here to omit the closing brace from consideration below.

				if (space_pos = StrChrAny(aKeys, _T(" \t"))) // Assign. Also, it relies on the fact that {} key names contain no spaces.
				{
					old_char = *space_pos;
					*space_pos = '\0';  // Temporarily terminate here so that TextToVK() can properly resolve a single char.
					key_name_length = space_pos - aKeys; // Override the default value set above.
					next_word = omit_leading_whitespace(space_pos + 1);
					UINT next_word_length = (UINT)(end_pos - next_word);
					if (next_word_length > 0)
					{
						if (!_tcsnicmp(next_word, _T("Down"), 4))
						{
							event_type = KEYDOWN;
							// v1.0.44.05: Added key_down_is_persistent (which is not initialized except here because
							// it's only applicable when event_type==KEYDOWN).  It avoids the following problem:
							// When a key is remapped to become a modifier (such as F1::Control), launching one of
							// the script's own hotkeys via F1 would lead to bad side-effects if that hotkey uses
							// the Send command. This is because the Send command assumes that any modifiers pressed
							// down by the script itself (such as Control) are intended to stay down during all
							// keystrokes generated by that script. To work around this, something like KeyWait F1
							// would otherwise be needed. within any hotkey triggered by the F1 key.
							key_down_is_persistent = _tcsnicmp(next_word + 4, _T("Temp"), 4); // "DownTemp" means non-persistent.
						}
						else if (!_tcsicmp(next_word, _T("Up")))
							event_type = KEYUP;
						else
							repeat_count = ATOI(next_word);
							// Above: If negative or zero, that is handled further below.
							// There is no complaint for values <1 to support scripts that want to conditionally send
							// zero keystrokes, e.g. Send {a %Count%}
					}
				}

				vk = TextToVK(aKeys, &mods_for_next_key, true, false, sTargetKeybdLayout); // false must be passed due to below.
				sc = vk ? 0 : TextToSC(aKeys);  // If sc is 0, it will be resolved by KeyEvent() later.
				if (!vk && !sc && ctoupper(aKeys[0]) == 'V' && ctoupper(aKeys[1]) == 'K')
				{
					LPTSTR sc_string = StrChrAny(aKeys + 2, _T("Ss")); // Look for the "SC" that demarks the scan code.
					if (sc_string && ctoupper(sc_string[1]) == 'C')
						sc = (sc_type)_tcstol(sc_string + 2, NULL, 16);  // Convert from hex.
					// else leave sc set to zero and just get the specified VK.  This supports Send {VKnn}.
					vk = (vk_type)_tcstol(aKeys + 2, NULL, 16);  // Convert from hex.
				}

				if (space_pos)  // undo the temporary termination
					*space_pos = old_char;
				*end_pos = '}';  // undo the temporary termination
				if (repeat_count < 1)
					goto brace_case_end; // Gets rid of one level of indentation. Well worth it.

				if (vk || sc)
				{
					if (key_as_modifiersLR = KeyToModifiersLR(vk, sc)) // Assign
					{
						if (!aTargetWindow)
						{
							if (event_type == KEYDOWN) // i.e. make {Shift down} have the same effect {ShiftDown}
							{
								this_event_modifier_down = vk;
								if (key_down_is_persistent) // v1.0.44.05.
									sModifiersLR_persistent |= key_as_modifiersLR;
								persistent_modifiers_for_this_SendKeys |= key_as_modifiersLR; // v1.0.44.06: Added this line to fix the fact that "DownTemp" should keep the key pressed down after the send.
							}
							else if (event_type == KEYUP) // *not* KEYDOWNANDUP, since that would be an intentional activation of the Start Menu or menu bar.
							{
								DisguiseWinAltIfNeeded(vk);
								sModifiersLR_persistent &= ~key_as_modifiersLR;
								// By contrast with KEYDOWN, KEYUP should also remove this modifier
								// from extra_persistent_modifiers_for_blind_mode if it happens to be
								// in there.  For example, if "#i::Send {LWin Up}" is a hotkey,
								// LWin should become persistently up in every respect.
								extra_persistent_modifiers_for_blind_mode &= ~key_as_modifiersLR;
								// Fix for v1.0.43: Also remove LControl if this key happens to be AltGr.
								if (vk == VK_RMENU && sTargetLayoutHasAltGr == CONDITION_TRUE) // It is AltGr.
									extra_persistent_modifiers_for_blind_mode &= ~MOD_LCONTROL;
								// Since key_as_modifiersLR isn't 0, update to reflect any changes made above:
								persistent_modifiers_for_this_SendKeys = sModifiersLR_persistent | extra_persistent_modifiers_for_blind_mode;
							}
							// else must never change sModifiersLR_persistent in response to KEYDOWNANDUP
							// because that would break existing scripts.  This is because that same
							// modifier key may have been pushed down via {ShiftDown} rather than "{Shift Down}".
							// In other words, {Shift} should never undo the effects of a prior {ShiftDown}
							// or {Shift down}.
						}
						//else don't add this event to sModifiersLR_persistent because it will not be
						// manifest via keybd_event.  Instead, it will done via less intrusively
						// (less interference with foreground window) via SetKeyboardState() and
						// PostMessage().  This change is for ControlSend in v1.0.21 and has been
						// documented.
					}
					// Below: sModifiersLR_persistent stays in effect (pressed down) even if the key
					// being sent includes that same modifier.  Surprisingly, this is how AutoIt2
					// behaves also, which is good.  Example: Send, {AltDown}!f  ; this will cause
					// Alt to still be down after the command is over, even though F is modified
					// by Alt.
					SendKey(vk, sc, mods_for_next_key, persistent_modifiers_for_this_SendKeys
						, repeat_count, event_type, key_as_modifiersLR, aTargetWindow);
				}

				else if (key_name_length == 1) // No vk/sc means a char of length one is sent via special method.
				{
					// v1.0.40: SendKeySpecial sends only keybd_event keystrokes, not ControlSend style
					// keystrokes.
					// v1.0.43.07: Added check of event_type!=KEYUP, which causes something like Send {ð up} to
					// do nothing if the curr. keyboard layout lacks such a key.  This is relied upon by remappings
					// such as F1::ð (i.e. a destination key that doesn't have a VK, at least in English).
					if (event_type != KEYUP) // In this mode, mods_for_next_key and event_type are ignored due to being unsupported.
					{
						if (aTargetWindow)
							// Although MSDN says WM_CHAR uses UTF-16, it seems to really do automatic
							// translation between ANSI and UTF-16; we rely on this for correct results:
							PostMessage(aTargetWindow, WM_CHAR, aKeys[0], 0);
						else
							SendKeySpecial(aKeys[0], repeat_count);
					}
				}

				// See comment "else must never change sModifiersLR_persistent" above about why
				// !aTargetWindow is used below:
				else if (vk = TextToSpecial(aKeys, key_text_length, event_type
					, persistent_modifiers_for_this_SendKeys, !aTargetWindow)) // Assign.
				{
					if (!aTargetWindow)
					{
						if (event_type == KEYDOWN)
							this_event_modifier_down = vk;
						else // It must be KEYUP because TextToSpecial() never returns KEYDOWNANDUP.
							DisguiseWinAltIfNeeded(vk);
					}
					// Since we're here, repeat_count > 0.
					// v1.0.42.04: A previous call to SendKey() or SendKeySpecial() might have left modifiers
					// in the wrong state (e.g. Send +{F1}{ControlDown}).  Since modifiers can sometimes affect
					// each other, make sure they're in the state intended by the user before beginning:
					SetModifierLRState(persistent_modifiers_for_this_SendKeys
						, sSendMode ? sEventModifiersLR : GetModifierLRState()
						, aTargetWindow, false, false); // It also does DoKeyDelay(g->PressDuration).
					for (int i = 0; i < repeat_count; ++i)
					{
						// Don't tell it to save & restore modifiers because special keys like this one
						// should have maximum flexibility (i.e. nothing extra should be done so that the
						// user can have more control):
						KeyEvent(event_type, vk, 0, aTargetWindow, true);
						if (!sSendMode)
							LONG_OPERATION_UPDATE_FOR_SENDKEYS
					}
				}

				else if (key_text_length > 4 && !_tcsnicmp(aKeys, _T("ASC "), 4) && !aTargetWindow) // {ASC nnnnn}
				{
					// Include the trailing space in "ASC " to increase uniqueness (selectivity).
					// Also, sending the ASC sequence to window doesn't work, so don't even try:
					SendASC(omit_leading_whitespace(aKeys + 3));
					// Do this only once at the end of the sequence:
					DoKeyDelay(); // It knows not to do the delay for SM_INPUT.
				}

				else if (key_text_length > 2 && !_tcsnicmp(aKeys, _T("U+"), 2))
				{
					// L24: Send a unicode value as shown by Character Map.
					wchar_t u_code = (wchar_t) _tcstol(aKeys + 2, NULL, 16);

					if (aTargetWindow)
					{
						// Although MSDN says WM_CHAR uses UTF-16, PostMessageA appears to truncate it to 8-bit.
						// This probably means it does automatic translation between ANSI and UTF-16.  Since we
						// specifically want to send a Unicode character value, use PostMessageW:
						PostMessageW(aTargetWindow, WM_CHAR, u_code, 0);
					}
					else
					{
						// Use SendInput in unicode mode if available, otherwise fall back to SendASC.
						// To know why the following requires sSendMode != SM_PLAY, see SendUnicodeChar.
						if (sSendMode != SM_PLAY && g_os.IsWin2000orLater())
						{
							SendUnicodeChar(u_code, mods_for_next_key | persistent_modifiers_for_this_SendKeys);
						}
						else // Note that this method generally won't work with Unicode characters except
						{	 // with specific controls which support it, such as RichEdit (tested on WordPad).
							TCHAR asc[8];
							*asc = '0';
							_itot(u_code, asc + 1, 10);
							SendASC(asc);
						}
					}
					DoKeyDelay();
				}

				//else do nothing since it isn't recognized as any of the above "else if" cases (see below).

				// If what's between {} is unrecognized, such as {Bogus}, it's safest not to send
				// the contents literally since that's almost certainly not what the user intended.
				// In addition, reset the modifiers, since they were intended to apply only to
				// the key inside {}.  Also, the below is done even if repeat-count is zero.

brace_case_end: // This label is used to simplify the code without sacrificing performance.
				aKeys = end_pos;  // In prep for aKeys++ done by the loop.
				mods_for_next_key = 0;
				continue;
			} // case '{'
			} // switch()
		} // if (!aSendRaw && strchr("^+!#{}", *aKeys))

		else // Encountered a character other than ^+!#{} ... or we're in raw mode.
		{
			// Best to call this separately, rather than as first arg in SendKey, since it changes the
			// value of modifiers and the updated value is *not* guaranteed to be passed.
			// In other words, SendKey(TextToVK(...), modifiers, ...) would often send the old
			// value for modifiers.
			single_char_string[0] = *aKeys; // String was pre-terminated earlier.
			if (vk = TextToVK(single_char_string, &mods_for_next_key, true, true, sTargetKeybdLayout))
				// TextToVK() takes no measurable time compared to the amount of time SendKey takes.
				SendKey(vk, 0, mods_for_next_key, persistent_modifiers_for_this_SendKeys, 1, KEYDOWNANDUP
					, 0, aTargetWindow);
			else // Try to send it by alternate means.
			{
				// In this mode, mods_for_next_key is ignored due to being unsupported.
				if (aTargetWindow) 
					// Although MSDN says WM_CHAR uses UTF-16, it seems to really do automatic
					// translation between ANSI and UTF-16; we rely on this for correct results:
					PostMessage(aTargetWindow, WM_CHAR, *aKeys, 0);
				else
					SendKeySpecial(*aKeys, 1);
			}
			mods_for_next_key = 0;  // Safest to reset this regardless of whether a key was sent.
		}
	} // for()

	modLR_type mods_to_set;
	if (sSendMode)
	{
		int final_key_delay = -1;  // Set default.
		if (!sAbortArraySend && sEventCount > 0) // Check for zero events for performance, but more importantly because playback hook will not operate correctly with zero.
		{
			// Add more events to the array (prior to sending) to support the following:
			// Restore the modifiers to match those the user is physically holding down, but do it as *part*
			// of the single SendInput/Play call.  The reasons it's done here as part of the array are:
			// 1) It avoids the need for #HotkeyModifierTimeout (and it's superior to it) for both SendInput
			//    and SendPlay.
			// 2) The hook will not be present during the SendInput, nor can it be reinstalled in time to
			//    catch any physical events generated by the user during the Send. Consequently, there is no
			//    known way to reliably detect physical keystate changes.
			// 3) Changes made to modifier state by SendPlay are seen only by the active window's thread.
			//    Thus, it would be inconsistent and possibly incorrect to adjust global modifier state
			//    after (or during) a SendPlay.
			// So rather than resorting to #HotkeyModifierTimeout, we can restore the modifiers within the
			// protection of SendInput/Play's uninterruptibility, allowing the user's buffered keystrokes
			// (if any) to hit against the correct modifier state when the SendInput/Play completes.
			// For example, if #c:: is a hotkey and the user releases Win during the SendInput/Play, that
			// release would hit after SendInput/Play restores Win to the down position, and thus Win would
			// not be stuck down.  Furthermore, if the user didn't release Win, Win would be in the
			// correct/intended position.
			// This approach has a few weaknesses (but the strengths appear to outweigh them):
			// 1) Hitting SendInput's 5000 char limit would omit the tail-end keystrokes, which would mess up
			//    all the assumptions here.  But hitting that limit should be very rare, especially since it's
			//    documented and thus scripts will avoid it.
			// 2) SendInput's assumed uninterruptibility is false if any other app or script has an LL hook
			//    installed.  This too is documented, so scripts should generally avoid using SendInput when
			//    they know there are other LL hooks in the system.  In any case, there's no known solution
			//    for it, so nothing can be done.
			mods_to_set = persistent_modifiers_for_this_SendKeys
				| (sInBlindMode ? 0 : (mods_down_physically_orig & ~mods_down_physically_but_not_logically_orig)); // The last item is usually 0.
			// Above: When in blind mode, don't restore physical modifiers.  This is done to allow a hotkey
			// such as the following to release Shift:
			//    +space::SendInput/Play {Blind}{Shift up}
			// Note that SendPlay can make such a change only from the POV of the target window; i.e. it can
			// release shift as seen by the target window, but not by any other thread; so the shift key would
			// still be considered to be down for the purpose of firing hotkeys (it can't change global key state
			// as seen by GetAsyncKeyState).
			// For more explanation of above, see a similar section for the non-array/old Send below.
			SetModifierLRState(mods_to_set, sEventModifiersLR, NULL, true, true); // Disguise in case user released or pressed Win/Alt during the Send (seems best to do it even for SendPlay, though it probably needs only Alt, not Win).
			// mods_to_set is used further below as the set of modifiers that were explicitly put into effect at the tail end of SendInput.
			SendEventArray(final_key_delay, mods_to_set);
		}
		CleanupEventArray(final_key_delay);
	}
	else // A non-array send is in effect, so a more elaborate adjustment to logical modifiers is called for.
	{
		// Determine (or use best-guess, if necessary) which modifiers are down physically now as opposed
		// to right before the Send began.
		modLR_type mods_down_physically; // As compared to mods_down_physically_orig.
		if (g_KeybdHook)
			mods_down_physically = g_modifiersLR_physical;
		else // No hook, so consult g_HotkeyModifierTimeout to make the determination.
			// Assume that the same modifiers that were phys+logically down before the Send are still
			// physically down (though not necessarily logically, since the Send may have released them),
			// but do this only if the timeout period didn't expire (or the user specified that it never
			// times out; i.e. elapsed time < timeout-value; DWORD subtraction gives the right answer even if
			// tick-count has wrapped around).
			mods_down_physically = (g_HotkeyModifierTimeout < 0 // It never times out or...
				|| (GetTickCount() - g_script.mThisHotkeyStartTime) < (DWORD)g_HotkeyModifierTimeout) // It didn't time out.
				? mods_down_physically_orig : 0;

		// Restore the state of the modifiers to be those the user is physically holding down right now.
		// Any modifiers that are logically "persistent", as detected upon entrance to this function
		// (e.g. due to something such as a prior "Send, {LWinDown}"), are also pushed down if they're not already.
		// Don't press back down the modifiers that were used to trigger this hotkey if there's
		// any doubt that they're still down, since doing so when they're not physically down
		// would cause them to be stuck down, which might cause unwanted behavior when the unsuspecting
		// user resumes typing.
		// v1.0.42.04: Now that SendKey() is lazy about releasing Ctrl and/or Shift (but not Win/Alt),
		// the section below also releases Ctrl/Shift if appropriate.  See SendKey() for more details.
		mods_to_set = persistent_modifiers_for_this_SendKeys; // Set default.
		if (sInBlindMode) // This section is not needed for the array-sending modes because they exploit uninterruptibility to perform a more reliable restoration.
		{
			// At the end of a blind-mode send, modifiers are restored differently than normal. One
			// reason for this is to support the explicit ability for a Send to turn off a hotkey's
			// modifiers even if the user is still physically holding them down.  For example:
			//   #space::Send {LWin up}  ; Fails to release it, by design and for backward compatibility.
			//   #space::Send {Blind}{LWin up}  ; Succeeds, allowing LWin to be logically up even though it's physically down.
			modLR_type mods_changed_physically_during_send = mods_down_physically_orig ^ mods_down_physically;
			// Fix for v1.0.42.04: To prevent keys from getting stuck down, compensate for any modifiers
			// the user physically pressed or released during the Send (especially those released).
			// Remove any modifiers physically released during the send so that they don't get pushed back down:
			mods_to_set &= ~(mods_changed_physically_during_send & mods_down_physically_orig); // Remove those that changed from down to up.
			// Conversely, add any modifiers newly, physically pressed down during the Send, because in
			// most cases the user would want such modifiers to be logically down after the Send.
			// Obsolete comment from v1.0.40: For maximum flexibility and minimum interference while
			// in blind mode, never restore modifiers to the down position then.
			mods_to_set |= mods_changed_physically_during_send & mods_down_physically; // Add those that changed from up to down.
		}
		else // Regardless of whether the keyboard hook is present, the following formula applies.
			mods_to_set |= mods_down_physically & ~mods_down_physically_but_not_logically_orig; // The second item is usually 0.
			// Above takes into account the fact that the user may have pressed and/or released some modifiers
			// during the Send.
			// So it includes all keys that are physically down except those that were down physically but not
			// logically at the *start* of the send operation (since the send operation may have changed the
			// logical state).  In other words, we want to restore the keys to their former logical-down
			// position to match the fact that the user is still holding them down physically.  The
			// previously-down keys we don't do this for are those that were physically but not logically down,
			// such as a naked Control key that's used as a suffix without being a prefix.  More details:
			// mods_down_physically_but_not_logically_orig is used to distinguish between the following two cases,
			// allowing modifiers to be properly restored to the down position when the hook is installed:
			// 1) A naked modifier key used only as suffix: when the user phys. presses it, it isn't
			//    logically down because the hook suppressed it.
			// 2) A modifier that is a prefix, that triggers a hotkey via a suffix, and that hotkey sends
			//    that modifier.  The modifier will go back up after the SEND, so the key will be physically
			//    down but not logically.

		// Use KEY_IGNORE_ALL_EXCEPT_MODIFIER to tell the hook to adjust g_modifiersLR_logical_non_ignored
		// because these keys being put back down match the physical pressing of those same keys by the
		// user, and we want such modifiers to be taken into account for the purpose of deciding whether
		// other hotkeys should fire (or the same one again if auto-repeating):
		// v1.0.42.04: A previous call to SendKey() might have left Shift/Ctrl in the down position
		// because by procrastinating, extraneous keystrokes in examples such as "Send ABCD" are
		// eliminated (previously, such that example released the shift key after sending each key,
		// only to have to press it down again for the next one.  For this reason, some modifiers
		// might get released here in addition to any that need to get pressed down.  That's why
		// SetModifierLRState() is called rather than the old method of pushing keys down only,
		// never releasing them.
		// Put the modifiers in mods_to_set into effect.  Although "true" is passed to disguise up-events,
		// there generally shouldn't be any up-events for Alt or Win because SendKey() would have already
		// released them.  One possible exception to this is when the user physically released Alt or Win
		// during the send (perhaps only during specific sensitive/vulnerable moments).
		SetModifierLRState(mods_to_set, GetModifierLRState(), aTargetWindow, true, true); // It also does DoKeyDelay(g->PressDuration).
	} // End of non-array Send.

	// For peace of mind and because that's how it was tested originally, the following is done
	// only after adjusting the modifier state above (since that adjustment might be able to
	// affect the global variables used below in a meaningful way).
	if (g_KeybdHook)
	{
		// Ensure that g_modifiersLR_logical_non_ignored does not contain any down-modifiers
		// that aren't down in g_modifiersLR_logical.  This is done mostly for peace-of-mind,
		// since there might be ways, via combinations of physical user input and the Send
		// commands own input (overlap and interference) for one to get out of sync with the
		// other.  The below uses ^ to find the differences between the two, then uses & to
		// find which are down in non_ignored that aren't in logical, then inverts those bits
		// in g_modifiersLR_logical_non_ignored, which sets those keys to be in the up position:
		g_modifiersLR_logical_non_ignored &= ~((g_modifiersLR_logical ^ g_modifiersLR_logical_non_ignored)
			& g_modifiersLR_logical_non_ignored);
	}

	if (prior_capslock_state == TOGGLED_ON) // The current user setting requires us to turn it back on.
		ToggleKeyState(VK_CAPITAL, TOGGLED_ON);

	// Might be better to do this after changing capslock state, since having the threads attached
	// tends to help with updating the global state of keys (perhaps only under Win9x in this case):
	if (threads_are_attached)
		AttachThreadInput(g_MainThreadID, target_thread, FALSE);

	if (do_selective_blockinput && !blockinput_prev) // Turn it back off only if it was off before we started.
		Line::ScriptBlockInput(false);

	// v1.0.43.03: Someone reported that when a non-autoreplace hotstring calls us to do its backspacing, the
	// hotstring's subroutine can execute a command that activates another window owned by the script before
	// the original window finished receiving its backspaces.  Although I can't reproduce it, this behavior
	// fits with expectations since our thread won't necessarily have a chance to process the incoming
	// keystrokes before executing the command that comes after SendInput.  If those command(s) activate
	// another of this thread's windows, that window will most likely intercept the keystrokes (assuming
	// that the message pump dispatches buffered keystrokes to whichever window is active at the time the
	// message is processed).
	// This fix does not apply to the SendPlay or SendEvent modes, the former due to the fact that it sleeps
	// a lot while the playback is running, and the latter due to key-delay and because testing has never shown
	// a need for it.
	if (aSendModeOrig == SM_INPUT && GetWindowThreadProcessId(GetForegroundWindow(), NULL) == g_MainThreadID) // GetWindowThreadProcessId() tolerates a NULL hwnd.
		SLEEP_WITHOUT_INTERRUPTION(-1);

	// v1.0.43.08: Restore the original thread key-delay values in case above temporarily overrode them.
	g.KeyDelay = orig_key_delay;
	g.PressDuration = orig_press_duration;
}



void SendKey(vk_type aVK, sc_type aSC, modLR_type aModifiersLR, modLR_type aModifiersLRPersistent
	, int aRepeatCount, KeyEventTypes aEventType, modLR_type aKeyAsModifiersLR, HWND aTargetWindow
	, int aX, int aY, bool aMoveOffset)
// Caller has ensured that: 1) vk or sc may be zero, but not both; 2) aRepeatCount > 0.
// This function is responsible for first setting the correct state of the modifier keys
// (as specified by the caller) before sending the key.  After sending, it should put the
// modifier keys  back to the way they were originally (UPDATE: It does this only for Win/Alt
// for the reasons described near the end of this function).
{
	// Caller is now responsible for verifying this:
	// Avoid changing modifier states and other things if there is nothing to be sent.
	// Otherwise, menu bar might activated due to ALT keystrokes that don't modify any key,
	// the Start Menu might appear due to WIN keystrokes that don't modify anything, etc:
	//if ((!aVK && !aSC) || aRepeatCount < 1)
	//	return;

	// I thought maybe it might be best not to release unwanted modifier keys that are already down
	// (perhaps via something like "Send, {altdown}{esc}{altup}"), but that harms the case where
	// modifier keys are down somehow, unintentionally: The send command wouldn't behave as expected.
	// e.g. "Send, abc" while the control key is held down by other means, would send ^a^b^c,
	// possibly dangerous.  So it seems best to default to making sure all modifiers are in the
	// proper down/up position prior to sending any Keybd events.  UPDATE: This has been changed
	// so that only modifiers that were actually used to trigger that hotkey are released during
	// the send.  Other modifiers that are down may be down intentionally, e.g. due to a previous
	// call to Send such as: Send {ShiftDown}.
	// UPDATE: It seems best to save the initial state only once, prior to sending the key-group,
	// because only at the beginning can the original state be determined without having to
	// save and restore it in each loop iteration.
	// UPDATE: Not saving and restoring at all anymore, due to interference (side-effects)
	// caused by the extra keybd events.

	// The combination of aModifiersLR and aModifiersLRPersistent are the modifier keys that
	// should be down prior to sending the specified aVK/aSC. aModifiersLR are the modifiers
	// for this particular aVK keystroke, but aModifiersLRPersistent are the ones that will stay
	// in pressed down even after it's sent.
	modLR_type modifiersLR_specified = aModifiersLR | aModifiersLRPersistent;
	bool vk_is_mouse = IsMouseVK(aVK); // Caller has ensured that VK is non-zero when it wants a mouse click.

	LONG_OPERATION_INIT
	for (int i = 0; i < aRepeatCount; ++i)
	{
		if (!sSendMode)
			LONG_OPERATION_UPDATE_FOR_SENDKEYS  // This does not measurably affect the performance of SendPlay/Event.
		// These modifiers above stay in effect for each of these keypresses.
		// Always on the first iteration, and thereafter only if the send won't be essentially
		// instantaneous.  The modifiers are checked before every key is sent because
		// if a high repeat-count was specified, the user may have time to release one or more
		// of the modifier keys that were used to trigger a hotkey.  That physical release
		// will cause a key-up event which will cause the state of the modifiers, as seen
		// by the system, to change.  For example, if user releases control-key during the operation,
		// some of the D's won't be control-D's:
		// ^c::Send,^{d 15}
		// Also: Seems best to do SetModifierLRState() even if Keydelay < 0:
		// Update: If this key is itself a modifier, don't change the state of the other
		// modifier keys just for it, since most of the time that is unnecessary and in
		// some cases, the extra generated keystrokes would cause complications/side-effects.
		if (!aKeyAsModifiersLR)
		{
			// DISGUISE UP: Pass "true" to disguise UP-events on WIN and ALT due to hotkeys such as:
			// !a::Send test
			// !a::Send {LButton}
			// v1.0.40: It seems okay to tell SetModifierLRState to disguise Win/Alt regardless of
			// whether our caller is in blind mode.  This is because our caller already put any extra
			// blind-mode modifiers into modifiersLR_specified, which prevents any actual need to
			// disguise anything (only the release of Win/Alt is ever disguised).
			// DISGUISE DOWN: Pass "false" to avoid disguising DOWN-events on Win and Alt because Win/Alt
			// will be immediately followed by some key for them to "modify".  The exceptions to this are
			// when aVK is a mouse button (e.g. sending !{LButton} or #{LButton}).  But both of those are
			// so rare that the flexibility of doing exactly what the script specifies seems better than
			// a possibly unwanted disguising.  Also note that hotkeys such as #LButton automatically use
			// both hooks so that the Start Menu doesn't appear when the Win key is released, so we're
			// not responsible for that type of disguising here.
			SetModifierLRState(modifiersLR_specified, sSendMode ? sEventModifiersLR : GetModifierLRState()
				, aTargetWindow, false, true, KEY_IGNORE_LEVEL(g->SendLevel)); // See keyboard_mouse.h for explanation of KEY_IGNORE.
			// SetModifierLRState() also does DoKeyDelay(g->PressDuration).
		}

		// v1.0.42.04: Mouse clicks are now handled here in the same loop as keystrokes so that the modifiers
		// will be readjusted (above) if the user presses/releases modifier keys during the mouse clicks.
		if (vk_is_mouse && !aTargetWindow)
			MouseClick(aVK, aX, aY, 1, g->DefaultMouseSpeed, aEventType, aMoveOffset);
			// Above: Since it's rare to send more than one click, it seems best to simplify and reduce code size
			// by not doing more than one click at a time event when mode is SendInput/Play.
		else
			// Sending mouse clicks via ControlSend is not supported, so in that case fall back to the
			// old method of sending the VK directly (which probably has no effect 99% of the time):
			KeyEvent(aEventType, aVK, aSC, aTargetWindow, true, KEY_IGNORE_LEVEL(g->SendLevel));
	} // for() [aRepeatCount]

	// The final iteration by the above loop does a key or mouse delay (KeyEvent and MouseClick do it internally)
	// prior to us changing the modifiers below.  This is a good thing because otherwise the modifiers would
	// sometimes be released so soon after the keys they modify that the modifiers are not in effect.
	// This can be seen sometimes when/ ctrl-shift-tabbing back through a multi-tabbed dialog:
	// The last ^+{tab} might otherwise not take effect because the CTRL key would be released too quickly.

	// Release any modifiers that were pressed down just for the sake of the above
	// event (i.e. leave any persistent modifiers pressed down).  The caller should
	// already have verified that aModifiersLR does not contain any of the modifiers
	// in aModifiersLRPersistent.  Also, call GetModifierLRState() again explicitly
	// rather than trying to use a saved value from above, in case the above itself
	// changed the value of the modifiers (i.e. aVk/aSC is a modifier).  Admittedly,
	// that would be pretty strange but it seems the most correct thing to do (another
	// reason is that the user may have pressed or released modifier keys during the
	// final mouse/key delay that was done above).
	if (!aKeyAsModifiersLR) // See prior use of this var for explanation.
	{
		// It seems best not to use KEY_IGNORE_ALL_EXCEPT_MODIFIER in this case, though there's
		// a slight chance that a script or two might be broken by not doing so.  The chance
		// is very slight because the only thing KEY_IGNORE_ALL_EXCEPT_MODIFIER would allow is
		// something like the following example.  Note that the hotkey below must be a hook
		// hotkey (even more rare) because registered hotkeys will still see the logical modifier
		// state and thus fire regardless of whether g_modifiersLR_logical_non_ignored says that
		// they shouldn't:
		// #b::Send, {CtrlDown}{AltDown}
		// $^!a::MsgBox You pressed the A key after pressing the B key.
		// In the above, making ^!a a hook hotkey prevents it from working in conjunction with #b.
		// UPDATE: It seems slightly better to have it be KEY_IGNORE_ALL_EXCEPT_MODIFIER for these reasons:
		// 1) Persistent modifiers are fairly rare.  When they're in effect, it's usually for a reason
		//    and probably a pretty good one and from a user who knows what they're doing.
		// 2) The condition that g_modifiersLR_logical_non_ignored was added to fix occurs only when
		//    the user physically presses a suffix key (or auto-repeats one by holding it down)
		//    during the course of a SendKeys() operation.  Since the persistent modifiers were
		//    (by definition) already in effect prior to the Send, putting them back down for the
		//    purpose of firing hook hotkeys does not seem unreasonable, and may in fact add value.
		// DISGUISE DOWN: When SetModifierLRState() is called below, it should only release keys, not press
		// any down (except if the user's physical keystrokes interfered).  Therefore, passing true or false
		// for the disguise-down-events parameter doesn't matter much (but pass "true" in case the user's
		// keystrokes did interfere in a way that requires a Alt or Win to be pressed back down, because
		// disguising it seems best).
		// DISGUISE UP: When SetModifierLRState() is called below, it is passed "false" for disguise-up
		// to avoid generating unnecessary disguise-keystrokes.  They are not needed because if our keystrokes
		// were modified by either WIN or ALT, the release of the WIN or ALT key will already be disguised due to
		// its having modified something while it was down.  The exceptions to this are when aVK is a mouse button
		// (e.g. sending !{LButton} or #{LButton}).  But both of those are so rare that the flexibility of doing
		// exactly what the script specifies seems better than a possibly unwanted disguising.
		// UPDATE for v1.0.42.04: Only release Win and Alt (if appropriate), not Ctrl and Shift, since we know
		// Win/Alt don't have to be disguised but our caller would have trouble tracking that info or making that
		// determination.  This avoids extra keystrokes, while still procrastinating the release of Ctrl/Shift so
		// that those can be left down if the caller's next keystroke happens to need them.
		modLR_type state_now = sSendMode ? sEventModifiersLR : GetModifierLRState();
		modLR_type win_alt_to_be_released = ((state_now ^ aModifiersLRPersistent) & state_now) // The modifiers to be released...
			& (MOD_LWIN|MOD_RWIN|MOD_LALT|MOD_RALT); // ... but restrict them to only Win/Alt.
		if (win_alt_to_be_released)
			SetModifierLRState(state_now & ~win_alt_to_be_released
				, state_now, aTargetWindow, true, false); // It also does DoKeyDelay(g->PressDuration).
	}
}



void SendKeySpecial(TCHAR aChar, int aRepeatCount)
// Caller must be aware that keystrokes are sent directly (i.e. never to a target window via ControlSend mode).
// It must also be aware that the event type KEYDOWNANDUP is always what's used since there's no way
// to support anything else.  Furthermore, there's no way to support "modifiersLR_for_next_key" such as ^€
// (assuming € is a character for which SendKeySpecial() is required in the current layout).
// This function uses some of the same code as SendKey() above, so maintain them together.
{
	// Caller must verify that aRepeatCount > 1.
	// Avoid changing modifier states and other things if there is nothing to be sent.
	// Otherwise, menu bar might activated due to ALT keystrokes that don't modify any key,
	// the Start Menu might appear due to WIN keystrokes that don't modify anything, etc:
	//if (aRepeatCount < 1)
	//	return;

	// v1.0.40: This function was heavily simplified because the old method of simulating
	// characters via dead keys apparently never executed under any keyboard layout.  It never
	// got past the following on the layouts I tested (Russian, German, Danish, Spanish):
	//		if (!send1 && !send2) // Can't simulate aChar.
	//			return;
	// This might be partially explained by the fact that the following old code always exceeded
	// the bounds of the array (because aChar was always between 0 and 127), so it was never valid
	// in the first place:
	//		asc_int = cAnsiToAscii[(int)((aChar - 128) & 0xff)] & 0xff;

	// Producing ANSI characters via Alt+Numpad and a leading zero appears standard on most languages
	// and layouts (at least those whose active code page is 1252/Latin 1 US/Western Europe).  However,
	// Russian (code page 1251 Cyrillic) is apparently one exception as shown by the fact that sending
	// all of the characters above Chr(127) while under Russian layout produces Cyrillic characters
	// if the active window's focused control is an Edit control (even if its an ANSI app).
	// I don't know the difference between how such characters are actually displayed as opposed to how
	// they're stored in memory (in notepad at least, there appears to be some kind of on-the-fly
	// translation to Unicode as shown when you try to save such a file).  But for now it doesn't matter
	// because for backward compatibility, it seems best not to change it until some alternative is
	// discovered that's high enough in value to justify breaking existing scripts that run under Russian
	// and other non-code-page-1252 layouts.
	//
	// Production of ANSI characters above 127 has been tested on both Windows XP and 98se (but not the
	// Win98 command prompt).

	TCHAR asc_string[16], *cp = asc_string;

	// The following range isn't checked because this function appears never to be called for such
	// characters (tested in English and Russian so far), probably because VkKeyScan() finds a way to
	// manifest them via Control+VK combinations:
	//if (aChar > -1 && aChar < 32)
	//	return;
	if (aChar & ~127)    // Try using ANSI.
		*cp++ = '0';  // ANSI mode is achieved via leading zero in the Alt+Numpad keystrokes.
	//else use Alt+Numpad without the leading zero, which allows the characters a-z, A-Z, and quite
	// a few others to be produced in Russian and perhaps other layouts, which was impossible in versions
	// prior to 1.0.40.
	_itot((TBYTE)aChar, cp, 10); // Convert to UCHAR in case aChar < 0.

	LONG_OPERATION_INIT
	for (int i = 0; i < aRepeatCount; ++i)
	{
		if (!sSendMode)
			LONG_OPERATION_UPDATE_FOR_SENDKEYS
#ifdef UNICODE
		if (sSendMode != SM_PLAY) // See SendUnicodeChar for comments.
			SendUnicodeChar(aChar);
		else
#endif
		SendASC(asc_string);
		DoKeyDelay(); // It knows not to do the delay for SM_INPUT.
	}

	// It is not necessary to do SetModifierLRState() to put a caller-specified set of persistent modifier
	// keys back into effect because:
	// 1) Our call to SendASC above (if any) at most would have released some of the modifiers (though never
	//    WIN because it isn't necessary); but never pushed any new modifiers down (it even releases ALT
	//    prior to returning).
	// 2) Our callers, if they need to push ALT back down because we didn't do it, will either disguise it
	//    or avoid doing so because they're about to send a keystroke (just about anything) that ALT will
	//    modify and thus not need to be disguised.
}



void SendASC(LPCTSTR aAscii)
// Caller must be aware that keystrokes are sent directly (i.e. never to a target window via ControlSend mode).
// aAscii is a string to support explicit leading zeros because sending 216, for example, is not the same as
// sending 0216.  The caller is also responsible for restoring any desired modifier keys to the down position
// (this function needs to release some of them if they're down).
{
	// UPDATE: In v1.0.42.04, the left Alt key is always used below because:
	// 1) It might be required on Win95/NT (though testing shows that RALT works okay on Windows 98se).
	// 2) It improves maintainability because if the keyboard layout has AltGr, and the Control portion
	//    of AltGr is released without releasing the RAlt portion, anything that expects LControl to
	//    be down whenever RControl is down would be broken.
	// The following test demonstrates that on previous versions under German layout, the right-Alt key
	// portion of AltGr could be used to manifest Alt+Numpad combinations:
	//   Send {RAlt down}{Asc 67}{RAlt up}  ; Should create a C character even when both the active window an AHK are set to German layout.
	//   KeyHistory  ; Shows that the right-Alt key was successfully used rather than the left.
	// Changing the modifier state via SetModifierLRState() (rather than some more error-prone multi-step method)
	// also ensures that the ALT key is pressed down only after releasing any shift key that needed it above.
	// Otherwise, the OS's switch-keyboard-layout hotkey would be triggered accidentally; e.g. the following
	// in English layout: Send ~~âÂ{^}.
	//
	// Make sure modifier state is correct: ALT pressed down and other modifiers UP
	// because CTRL and SHIFT seem to interfere with this technique if they are down,
	// at least under WinXP (though the Windows key doesn't seem to be a problem).
	// Specify KEY_IGNORE so that this action does not affect the modifiers that the
	// hook uses to determine which hotkey should be triggered for a suffix key that
	// has more than one set of triggering modifiers (for when the user is holding down
	// that suffix to auto-repeat it -- see keyboard_mouse.h for details).
	modLR_type modifiersLR_now = sSendMode ? sEventModifiersLR : GetModifierLRState();
	SetModifierLRState((modifiersLR_now | MOD_LALT) & ~(MOD_RALT | MOD_LCONTROL | MOD_RCONTROL | MOD_LSHIFT | MOD_RSHIFT)
		, modifiersLR_now, NULL, false // Pass false because there's no need to disguise the down-event of LALT.
		, true, KEY_IGNORE); // Pass true so that any release of RALT is disguised (Win is never released here).
	// Note: It seems best never to press back down any key released above because the
	// act of doing so may do more harm than good (i.e. the keystrokes may caused
	// unexpected side-effects.

	// Known limitation (but obscure): There appears to be some OS limitation that prevents the following
	// AltGr hotkey from working more than once in a row:
	// <^>!i::Send {ASC 97}
	// Key history indicates it's doing what it should, but it doesn't actually work.  You have to press the
	// left-Alt key (not RAlt) once to get the hotkey working again.

	// This is not correct because it is possible to generate unicode characters by typing
	// Alt+256 and beyond:
	// int value = ATOI(aAscii);
	// if (value < 0 || value > 255) return 0; // Sanity check.

	// Known issue: If the hotkey that triggers this Send command is CONTROL-ALT
	// (and maybe either CTRL or ALT separately, as well), the {ASC nnnn} method
	// might not work reliably due to strangeness with that OS feature, at least on
	// WinXP.  I already tried adding delays between the keystrokes and it didn't help.

	// Caller relies upon us to stop upon reaching the first non-digit character:
	for (LPCTSTR cp = aAscii; *cp >= '0' && *cp <= '9'; ++cp)
		// A comment from AutoIt3: ASCII 0 is 48, NUMPAD0 is 96, add on 48 to the ASCII.
		// Also, don't do WinDelay after each keypress in this case because it would make
		// such keys take up to 3 or 4 times as long to send (AutoIt3 avoids doing the
		// delay also).  Note that strings longer than 4 digits are allowed because
		// some or all OSes support Unicode characters 0 through 65535.
		KeyEvent(KEYDOWNANDUP, *cp + 48);

	// Must release the key regardless of whether it was already down, so that the sequence will take effect
	// immediately.  Otherwise, our caller might not release the Alt key (since it might need to stay down for
	// other purposes), in which case Alt+Numpad character would never appear and the caller's subsequent
	// keystrokes might get absorbed by the OS's special state of "waiting for Alt+Numpad sequence to complete".
	// Another reason is that the user may be physically holding down Alt, in which case the caller might never
	// release it.  In that case, we want the Alt+Numpad character to appear immediately rather than waiting for
	// the user to release Alt (in the meantime, the caller will likely press Alt back down to match the physical
	// state).
	KeyEvent(KEYUP, VK_MENU);
}



LRESULT CALLBACK PlaybackProc(int aCode, WPARAM wParam, LPARAM lParam)
// Journal playback hook.
{
	static bool sThisEventHasBeenLogged, sThisEventIsScreenCoord;

	switch (aCode)
	{
	case HC_GETNEXT:
	{
		if (sFirstCallForThisEvent)
		{
			// Gather the delay(s) for this event, if any, and calculate the time the keystroke should be sent.
			// NOTE: It must be done this way because testing shows that simply returning the desired delay
			// for the first call of each event is not reliable, at least not for the first few events (they
			// tend to get sent much more quickly than specified).  More details:
			// MSDN says, "When the system ...calls the hook procedure [after the first time] with code set to
			// HC_GETNEXT to retrieve the same message... the return value... should be zero."
			// Apparently the above is overly cautious wording with the intent to warn people not to write code
			// that gets stuck in infinite playback due to never returning 0, because returning non-zero on
			// calls after the first works fine as long as 0 is eventually returned.  Furthermore, I've seen
			// other professional code examples that uses this "countdown" approach, so it seems valid.
			sFirstCallForThisEvent = false;
			sThisEventHasBeenLogged = false;
			sThisEventIsScreenCoord = false;
			for (sThisEventTime = GetTickCount()
				; !sEventPB[sCurrentEvent].message // HC_SKIP has ensured there is a non-delay event, so no need to check sCurrentEvent < sEventCount.
				; sThisEventTime += sEventPB[sCurrentEvent++].time_to_wait); // Overflow is okay.
		}
		// Above has ensured that sThisEventTime is valid regardless of whether this is the first call
		// for this event.  It has also incremented sCurrentEvent, if needed, for use below.

		// Copy the current mouse/keyboard event to the EVENTMSG structure (lParam).
		// MSDN says that HC_GETNEXT can be received multiple times consecutively, in which case the
		// same event should be copied into the structure each time.
		PlaybackEvent &source_event = sEventPB[sCurrentEvent];
		EVENTMSG &event = *(PEVENTMSG)lParam;  // For convenience, maintainability, and possibly performance.
		// Currently, the following isn't documented entirely accurately at MSDN, but other sources confirm
		// the below are the proper values to store.  In addition, the extended flag as set below has been
		// confirmed to work properly by monitoring the resulting WM_KEYDOWN message in a main message loop.
		//
		// Strip off extra bits early for maintainability.  It must be stripped off the source event itself
		// because if HC_GETNEXT is called again for this same event, don't want to apply the offset again.
		bool has_coord_offset;
		if (has_coord_offset = source_event.message & MSG_OFFSET_MOUSE_MOVE)
			source_event.message &= ~MSG_OFFSET_MOUSE_MOVE;
		event.message = source_event.message;
		// The following members are not set because testing confirms that they're ignored:
		// event.hwnd: ignored even if assigned the HWND of an existing window or control.
		// event.time: Apparently ignored in favor of this playback proc's return value.  Furthermore,
		// testing shows that the posted keystroke message (e.g. WM_KEYDOWN) has the correct timestamp
		// even when event.time is left as a random time, which shows that the member is completely
		// ignored during playback, at least on XP.
		bool is_keyboard_not_mouse;
		if (is_keyboard_not_mouse = (source_event.message >= WM_KEYFIRST && source_event.message <= WM_KEYLAST)) // Keyboard event.
		{
			event.paramL = (source_event.sc << 8) | source_event.vk;
			event.paramH = source_event.sc & 0xFF; // 0xFF omits the extended-key-bit, if present.
			if (source_event.sc & 0x100) // It's an extended key.
				event.paramH |= 0x8000; // So mark it that way using EVENTMSG's convention.
			// Notes about inability of playback to simulate LWin and RWin in a way that performs their native function:
			// For the following reasons, it seems best not to send LWin/RWin via keybd_event inside the playback hook:
			// 1) Complexities such as having to check for an array that consists entirely of LWin/RWin events,
			//    in which case the playback hook mustn't be activated because it requires that we send
			//    at least one event through it.  Another complexity is that all keys modified by Win would
			//    have to be flagged in the array as needing to be sent via keybd_event.
			// 2) It might preserve some flexibility to be able to send LWin/RWin events directly to a window,
			//    similar to ControlSend (perhaps for shells other than Explorer, who might allow apps to make
			//    use of LWin/RWin internally). The window should receive LWIN/RWIN as WM_KEYDOWN messages when
			//    sent via playback.  Note: unlike the neutral SHIFT/ALT/CTRL keys, which are detectible via the
			//    target thread's call to GetKeyState(), LWin and RWin aren't detectible that way.
			// 3) Code size and complexity.
			//
			// Related: LWin and RWin are released and pressed down during playback for simplicity and also
			// on the off-chance the target window takes note of the incoming WM_KEYDOWN on VK_LWIN/RWIN and
			// changes state until the up-event is received (however, the target thread's call of GetKeyState
			// can't see a any effect for hook-sent LWin/RWin).
			//
			// Related: If LWin or RWin is logically down at start of SendPlay, SendPlay's events won't be
			// able to release it from the POV of the target thread's calls to GetKeyState().  That might mess
			// things up for apps that check the logical state of the Win keys.  But due to rarity: in those
			// cases, a workaround would be to do an explicit old-style Send {Blind} (as the first line of the
			// hotkey) to release the modifier logically prior to SendPlay commands.
			//
			// Related: Although some apps might not like receiving SendPlay's LWin/RWin if shell==Explorer
			// (since there may be no normal way for such keystrokes to arrive as WM_KEYDOWN events) maybe it's
			// best not to omit/ignore LWin/RWin if it is possible in other shells, or adds flexibility.
			// After all, sending {LWin/RWin} via hook should be rare, especially if it has no effect (except
			// for cases where a Win hotkey releases LWin as part of SendPlay, but even that can be worked
			// around via an explicit Send {Blind}{LWin up} beforehand).
		}
		else // MOUSE EVENT.
		{
			// Unlike keybd_event() and SendInput(), explicit coordinates must be specified for each mouse event.
			// The builder of this array must ensure that coordinates are valid or set to COORD_UNSPECIFIED_SHORT.
			if (source_event.x == COORD_UNSPECIFIED_SHORT || has_coord_offset)
			{
				// For simplicity with calls such as CoordToScreen(), the one who set up this array has ensured
				// that both X and Y are either COORD_UNSPECIFIED_SHORT or not so (i.e. not a combination).
				// Since the user nor anything else can move the cursor during our playback, GetCursorPos()
				// should accurately reflect the position set by any previous mouse-move done by this playback.
				// This seems likely to be true even for DirectInput games, though hasn't been tested yet.
				POINT cursor;
				GetCursorPos(&cursor);
				event.paramL = cursor.x;
				event.paramH = cursor.y;
				if (has_coord_offset) // The specified coordinates are offsets to be applied to the cursor's current position.
				{
					event.paramL += source_event.x;
					event.paramH += source_event.y;
					// Update source array in case HC_GETNEXT is called again for this same event, in which case
					// don't want to apply the offset again (the has-offset flag has already been removed from the
					// source event higher above).
					source_event.x = event.paramL;
					source_event.y = event.paramH;
					sThisEventIsScreenCoord = true; // Mark the above as absolute vs. relative in case HC_GETNEXT is called again for this event.
				}
			}
			else
			{
				event.paramL = source_event.x;
				event.paramH = source_event.y;
				if (!sThisEventIsScreenCoord) // Coordinates are relative to the window that is active now (during realtime playback).
					CoordToScreen((int &)event.paramL, (int &)event.paramH, COORD_MODE_MOUSE); // Playback uses screen coords.
			}
		}
		LRESULT time_until_event = (int)(sThisEventTime - GetTickCount()); // Cast to int to avoid loss of negatives from DWORD subtraction.
		if (time_until_event > 0)
			return time_until_event;
		// Otherwise, the event is scheduled to occur immediately (or is overdue).  In case HC_GETNEXT can be
		// called multiple times even when we previously returned 0, ensure the event is logged only once.
		if (!sThisEventHasBeenLogged && is_keyboard_not_mouse) // Mouse events aren't currently logged for consistency with other send methods.
		{
			// The event is logged here rather than higher above so that its timestamp is accurate.
			// It's also so that events aren't logged if the user cancel's the operation in the middle
			// (by pressing Ctrl-Alt-Del or Ctrl-Esc).
			UpdateKeyEventHistory(source_event.message == WM_KEYUP || source_event.message == WM_SYSKEYUP
				, source_event.vk, source_event.sc);
			sThisEventHasBeenLogged = true;
		}
		return 0; // No CallNextHookEx(). See comments further below.
	} // case HC_GETNEXT.

	case HC_SKIP: // Advance to the next mouse/keyboard event, if any.
		// Advance to the next item, which is either a delay or an event (preps for next HC_GETNEXT).
		++sCurrentEvent;
		// Although caller knows it has to do the tail-end delay (if any) since there's no way to
		// do a trailing delay at the end of playback, it may have put a delay at the end of the
		// array anyway for code simplicity.  For that reason and maintainability:
		// Skip over any delays that are present to discover if there is a next event.
		UINT u;
		for (u = sCurrentEvent; u < sEventCount && !sEventPB[u].message; ++u);
		if (u == sEventCount) // No more events.
		{
			// MSDN implies in the following statement that it's acceptable (and perhaps preferable in
			// the case of a playback hook) for the hook to unhook itself: "The hook procedure can be in the
			// state of being called by another thread even after UnhookWindowsHookEx returns."
			UnhookWindowsHookEx(g_PlaybackHook);
			g_PlaybackHook = NULL; // Signal the installer of the hook that it's gone now.
			// The following is an obsolete method from pre-v1.0.44.  Do not reinstate it without adding handling
			// to MainWindowProc() to do "g_PlaybackHook = NULL" upon receipt of WM_CANCELJOURNAL.
			// PostMessage(g_hWnd, WM_CANCELJOURNAL, 0, 0); // v1.0.44: Post it to g_hWnd vs. NULL because it's a little safer (SEE COMMENTS in MsgSleep's WM_CANCELJOURNAL for why it's almost completely safe with NULL).
			// Above: By using WM_CANCELJOURNAL instead of a custom message, the creator of this hook can easily
			// use a message filter to watch for both a system-generated removal of the hook (via the user
			// pressing Ctrl-Esc. or Ctrl-Alt-Del) or one we generate here (though it's currently not implemented
			// that way because it would prevent journal playback to one of our thread's own windows).
		}
		else
			sFirstCallForThisEvent = true; // Reset to prepare for next HC_GETNEXT.
		return 0; // MSDN: The return value is used only if the hook code is HC_GETNEXT; otherwise, it is ignored.

	default:
		// Covers the following cases:
		//case HC_NOREMOVE: // MSDN: An application has called the PeekMessage function with wRemoveMsg set to PM_NOREMOVE, indicating that the message is not removed from the message queue after PeekMessage processing.
		//case HC_SYSMODALON:  // MSDN: A system-modal dialog box is being displayed. Until the dialog box is destroyed, the hook procedure must stop playing back messages.
		//case HC_SYSMODALOFF: // MSDN: A system-modal dialog box has been destroyed. The hook procedure must resume playing back the messages.
		//case(...aCode < 0...): MSDN docs specify that the hook should return in this case.
		//
		// MS gives some sample code at http://support.microsoft.com/default.aspx?scid=KB;EN-US;124835
		// about the proper values to return to avoid hangs on NT (it seems likely that this implementation
		// is compliant enough if you read between the lines).  Their sample code indicates that
		// "return CallNextHook()"  should be done for basically everything except HC_SKIP/HC_GETNEXT, so
		// as of 1.0.43.08, that is what is done here.
		// Testing shows that when a so-called system modal dialog is displayed (even if it isn't the
		// active window) playback stops automatically, probably because the system doesn't call the hook
		// during such times (only a "MsgBox 4096" has been tested so far).
		//
		// The first parameter uses g_PlaybackHook rather than NULL because MSDN says it's merely
		// "currently ignored", but in the older "Win32 hooks" article, it says that the behavior
		// may change in the future.
		return CallNextHookEx(g_PlaybackHook, aCode, wParam, lParam);
		// Except for the cases above, CallNextHookEx() is not called for performance and also because from
		// what I can tell from the MSDN docs and other examples, it is neither required nor desirable to do so
		// during playback's SKIP/GETNEXT.
		// MSDN: The return value is used only if the hook code is HC_GETNEXT; otherwise, it is ignored.
	} // switch().

	// Execution should never reach since all cases do their own custom return above.
}



#ifdef JOURNAL_RECORD_MODE
LRESULT CALLBACK RecordProc(int aCode, WPARAM wParam, LPARAM lParam)
{
	switch (aCode)
	{
	case HC_ACTION:
	{
		EVENTMSG &event = *(PEVENTMSG)lParam;
		PlaybackEvent &dest_event = sEventPB[sEventCount];
		dest_event.message = event.message;
		if (event.message >= WM_MOUSEFIRST && event.message <= WM_MOUSELAST) // Mouse event, including wheel.
		{
			if (event.message != WM_MOUSEMOVE)
			{
				// WHEEL: No info comes in about which direction the wheel was turned (nor by how many notches).
				// In addition, it appears impossible to specify such info when playing back the event.
				// Therefore, playback usually produces downward wheel movement (but upward in some apps like
				// Visual Studio).
				dest_event.x = event.paramL;
				dest_event.y = event.paramH;
				++sEventCount;
			}
		}
		else // Keyboard event.
		{
			dest_event.vk = event.paramL & 0x00FF;
			dest_event.sc = (event.paramL & 0xFF00) >> 8;
			if (event.paramH & 0x8000) // Extended key.
				dest_event.sc |= 0x100;
			if (dest_event.vk == VK_CANCEL) // Ctrl+Break.
			{
				UnhookWindowsHookEx(g_PlaybackHook);
				g_PlaybackHook = NULL; // Signal the installer of the hook that it's gone now.
				// Obsolete method, pre-v1.0.44:
				//PostMessage(g_hWnd, WM_CANCELJOURNAL, 0, 0); // v1.0.44: Post it to g_hWnd vs. NULL so that it isn't lost when script is displaying a MsgBox or other dialog.
			}
			++sEventCount;
		}
		break;
	}

	//case HC_SYSMODALON:  // A system-modal dialog box is being displayed. Until the dialog box is destroyed, the hook procedure must stop playing back messages.
	//case HC_SYSMODALOFF: // A system-modal dialog box has been destroyed. The hook procedure must resume playing back the messages.
	//	break;
	}

	// Unlike the playback hook, it seems more correct to call CallNextHookEx() unconditionally so that
	// any other journal record hooks can also record the event.  But MSDN is quite vague about this.
	return CallNextHookEx(g_PlaybackHook, aCode, wParam, lParam);
	// Return value is ignored, except possibly when aCode < 0 (MSDN is unclear).
}
#endif



void KeyEvent(KeyEventTypes aEventType, vk_type aVK, sc_type aSC, HWND aTargetWindow
	, bool aDoKeyDelay, DWORD aExtraInfo)
// aSC or aVK (but not both), can be zero to cause the default to be used.
// For keys like NumpadEnter -- that have a unique scancode but a non-unique virtual key --
// caller can just specify the sc.  In addition, the scan code should be specified for keys
// like NumpadPgUp and PgUp.  In that example, the caller would send the same scan code for
// both except that PgUp would be extended.   sc_to_vk() would map both of them to the same
// virtual key, which is fine since it's the scan code that matters to apps that can
// differentiate between keys with the same vk.
//
// Thread-safe: This function is not fully thread-safe because keystrokes can get interleaved,
// but that's always a risk with anything short of SendInput.  In fact,
// when the hook ISN'T installed, keystrokes can occur in another thread, causing the key state to
// change in the middle of KeyEvent, which is the same effect as not having thread-safe key-states
// such as GetKeyboardState in here.  Also, the odds of both our threads being in here simultaneously
// is greatly reduced by the fact that the hook thread never passes "true" for aDoKeyDelay, thus
// its calls should always be very fast.  Even if a collision does occur, there are thread-safety
// things done in here that should reduce the impact to nearly as low as having a dedicated
// KeyEvent function solely for calling by the hook thread (which might have other problems of its own).
{
	if (!(aVK | aSC)) // MUST USE BITWISE-OR (see comment below).
		return;
	// The above implements the rule "if neither VK nor SC was specified, return".  But they must be done as
	// bitwise-OR rather than logical-AND/OR, otherwise MSVC++ 7.1 generates 16KB of extra OBJ size for some reason.
	// That results in an 2 KB increase in compressed EXE size, and I think about 6 KB uncompressed.
	// I tried all kids of variations and reconfigurations, but the above is the only simple one that works.
	// Strangely, the same logic above (but with the preferred logical-AND/OR operator) appears elsewhere in the
	// code but doesn't bloat there.  Examples:
	//   !aVK && !aSC
	//   !vk && !sc
	//   !(aVK || aSC)

	if (!aExtraInfo) // Shouldn't be called this way because 0 is considered false in some places below (search on " = aExtraInfo" to find them).
		aExtraInfo = KEY_IGNORE_ALL_EXCEPT_MODIFIER; // Seems slightly better to use a standard default rather than something arbitrary like 1.

	// Since calls from the hook thread could come in even while the SendInput array is being constructed,
	// don't let those events get interspersed with the script's explicit use of SendInput.
	bool caller_is_keybd_hook = (GetCurrentThreadId() == g_HookThreadID);
	bool put_event_into_array = sSendMode && !caller_is_keybd_hook;
	if (sSendMode == SM_INPUT || caller_is_keybd_hook) // First check is necessary but second is just for maintainability.
		aDoKeyDelay = false;

	// Even if the sc_to_vk() mapping results in a zero-value vk, don't return.
	// I think it may be valid to send keybd_events	that have a zero vk.
	// In any case, it's unlikely to hurt anything:
	if (!aVK)
		aVK = sc_to_vk(aSC);
	else
		if (!aSC)
			// In spite of what the MSDN docs say, the scan code parameter *is* used, as evidenced by
			// the fact that the hook receives the proper scan code as sent by the below, rather than
			// zero like it normally would.  Even though the hook would try to use MapVirtualKey() to
			// convert zero-value scan codes, it's much better to send it here also for full compatibility
			// with any apps that may rely on scan code (and such would be the case if the hook isn't
			// active because the user doesn't need it; also for some games maybe).  In addition, if the
			// current OS is Win9x, we must map it here manually (above) because otherwise the hook
			// wouldn't be able to differentiate left/right on keys such as RControl, which is detected
			// via its scan code.
			aSC = vk_to_sc(aVK);

	BYTE aSC_lobyte = LOBYTE(aSC);
	DWORD event_flags = HIBYTE(aSC) ? KEYEVENTF_EXTENDEDKEY : 0;

	// Do this only after the above, so that the SC is left/right specific if the VK was such,
	// even on Win9x (though it's probably never called that way for Win9x; it's probably always
	// called with either just the proper left/right SC or that plus the neutral VK).
	// Under WinNT/2k/XP, sending VK_LCONTROL and such result in the high-level (but not low-level
	// I think) hook receiving VK_CONTROL.  So somewhere internally it's being translated (probably
	// by keybd_event itself).  In light of this, translate the keys here manually to ensure full
	// support under Win9x (which might not support this internal translation).  The scan code
	// looked up above should still be correct for left-right centric keys even under Win9x.
	// v1.0.43: Apparently, the journal playback hook also requires neutral modifier keystrokes
	// rather than left/right ones.  Otherwise, the Shift key can't capitalize letters, etc.
	if (sSendMode == SM_PLAY || g_os.IsWin9x())
	{
		// Convert any non-neutral VK's to neutral for these OSes, since apps and the OS itself
		// can't be expected to support left/right specific VKs while running under Win9x:
		switch(aVK)
		{
		case VK_LCONTROL:
		case VK_RCONTROL: aVK = VK_CONTROL; break; // But leave scan code set to a left/right specific value rather than converting it to "left" unconditionally.
		case VK_LSHIFT:
		case VK_RSHIFT: aVK = VK_SHIFT; break;
		case VK_LMENU:
		case VK_RMENU: aVK = VK_MENU; break;
		}
	}

	// aTargetWindow is almost always passed in as NULL by our caller, even if the overall command
	// being executed is ControlSend.  This is because of the following reasons:
	// 1) Modifiers need to be logically changed via keybd_event() when using ControlSend against
	//    a cmd-prompt, console, or possibly other types of windows.
	// 2) If a hotkey triggered the ControlSend that got us here and that hotkey is a naked modifier
	//    such as RAlt:: or modified modifier such as ^#LShift, that modifier would otherwise auto-repeat
	//    an possibly interfere with the send operation.  This auto-repeat occurs because unlike a normal
	//    send, there are no calls to keybd_event() (keybd_event() stop the auto-repeat as a side-effect).
	// One exception to this is something like "ControlSend, Edit1, {Control down}", which explicitly
	// calls us with a target window.  This exception is by design and has been bug-fixed and documented
	// in ControlSend for v1.0.21:
	if (aTargetWindow) // This block shouldn't affect overall thread-safety because hook thread never calls it in this mode.
	{
		if (KeyToModifiersLR(aVK, aSC))
		{
			// When sending modifier keystrokes directly to a window, use the AutoIt3 SetKeyboardState()
			// technique to improve the reliability of changes to modifier states.  If this is not done,
			// sometimes the state of the SHIFT key (and perhaps other modifiers) will get out-of-sync
			// with what's intended, resulting in uppercase vs. lowercase problems (and that's probably
			// just the tip of the iceberg).  For this to be helpful, our caller must have ensured that
			// our thread is attached to aTargetWindow's (but it seems harmless to do the below even if
			// that wasn't done for any reason).  Doing this here in this function rather than at a
			// higher level probably isn't best in terms of performance (e.g. in the case where more
			// than one modifier is being changed, the multiple calls to Get/SetKeyboardState() could
			// be consolidated into one call), but it is much easier to code and maintain this way
			// since many different functions might call us to change the modifier state:
			BYTE state[256];
			GetKeyboardState((PBYTE)&state);
			if (aEventType == KEYDOWN)
				state[aVK] |= 0x80;
			else if (aEventType == KEYUP)
				state[aVK] &= ~0x80;
			// else KEYDOWNANDUP, in which case it seems best (for now) not to change the state at all.
			// It's rarely if ever called that way anyway.

			// If aVK is a left/right specific key, be sure to also update the state of the neutral key:
			switch(aVK)
			{
			case VK_LCONTROL: 
			case VK_RCONTROL:
				if ((state[VK_LCONTROL] & 0x80) || (state[VK_RCONTROL] & 0x80))
					state[VK_CONTROL] |= 0x80;
				else
					state[VK_CONTROL] &= ~0x80;
				break;
			case VK_LSHIFT:
			case VK_RSHIFT:
				if ((state[VK_LSHIFT] & 0x80) || (state[VK_RSHIFT] & 0x80))
					state[VK_SHIFT] |= 0x80;
				else
					state[VK_SHIFT] &= ~0x80;
				break;
			case VK_LMENU:
			case VK_RMENU:
				if ((state[VK_LMENU] & 0x80) || (state[VK_RMENU] & 0x80))
					state[VK_MENU] |= 0x80;
				else
					state[VK_MENU] &= ~0x80;
				break;
			}

			SetKeyboardState((PBYTE)&state);
			// Even after doing the above, we still continue on to send the keystrokes
			// themselves to the window, for greater reliability (same as AutoIt3).
		}

		// lowest 16 bits: repeat count: always 1 for up events, probably 1 for down in our case.
		// highest order bits: 11000000 (0xC0) for keyup, usually 00000000 (0x00) for keydown.
		LPARAM lParam = (LPARAM)(aSC << 16);
		if (aEventType != KEYUP)  // i.e. always do it for KEYDOWNANDUP
			PostMessage(aTargetWindow, WM_KEYDOWN, aVK, lParam | 0x00000001);
		// The press-duration delay is done only when this is a down-and-up because otherwise,
		// the normal g->KeyDelay will be in effect.  In other words, it seems undesirable in
		// most cases to do both delays for only "one half" of a keystroke:
		if (aDoKeyDelay && aEventType == KEYDOWNANDUP)
			DoKeyDelay(g->PressDuration); // Since aTargetWindow!=NULL, sSendMode!=SM_PLAY, so no need for to ever use the SendPlay press-duration.
		if (aEventType != KEYDOWN)  // i.e. always do it for KEYDOWNANDUP
			PostMessage(aTargetWindow, WM_KEYUP, aVK, lParam | 0xC0000001);
	}
	else // Keystrokes are to be sent with keybd_event() or the event array rather than PostMessage().
	{
		// The following static variables are intentionally NOT thread-safe because their whole purpose
		// is to watch the combined stream of keystrokes from all our threads.  Due to our threads'
		// keystrokes getting interleaved with the user's and those of other threads, this kind of
		// monitoring is never 100% reliable.  All we can do is aim for an astronomically low chance
		// of failure.
		// Users of the below want them updated only for keybd_event() keystrokes (not PostMessage ones):
		sPrevEventType = aEventType;
		sPrevVK = aVK;
		// Turn off BlockInput momentarily to support sending of the ALT key.  This is not done for
		// Win9x because input cannot be simulated during BlockInput on that platform anyway; thus
		// it seems best (due to backward compatibility) not to turn off BlockInput then.
		// Jon Bennett noted: "As many of you are aware BlockInput was "broken" by a SP1 hotfix under
		// Windows XP so that the ALT key could not be sent. I just tried it under XP SP2 and it seems
		// to work again."  In light of this, it seems best to unconditionally and momentarily disable
		// input blocking regardless of which OS is being used (except Win9x, since no simulated input
		// is even possible for those OSes).
		// For thread safety, allow block-input modification only by the main thread.  This should avoid
		// any chance that block-input will get stuck on due to two threads simultaneously reading+changing
		// g_BlockInput (changes occur via calls to ScriptBlockInput).
		bool we_turned_blockinput_off = g_BlockInput && (aVK == VK_MENU || aVK == VK_LMENU || aVK == VK_RMENU)
			&& !caller_is_keybd_hook && g_os.IsWinNT4orLater(); // Ordered for short-circuit performance.
		if (we_turned_blockinput_off)
			Line::ScriptBlockInput(false);

		vk_type control_vk;      // When not set somewhere below, these are left uninitialized to help catch bugs.
		HKL target_keybd_layout; //
		ResultType r_mem, &target_layout_has_altgr = caller_is_keybd_hook ? r_mem : sTargetLayoutHasAltGr; // Same as above.
		bool hookable_ralt, lcontrol_was_down, do_detect_altgr;
		if (do_detect_altgr = hookable_ralt = (aVK == VK_RMENU && !put_event_into_array && g_KeybdHook)) // This is an RALT that the hook will be able to monitor. Using VK_RMENU vs. VK_MENU should be safe since this logic is only needed for the hook, which is never in effect on Win9x.
		{
			if (!caller_is_keybd_hook) // sTargetKeybdLayout is set/valid only by SendKeys().
				target_keybd_layout = sTargetKeybdLayout;
			else
			{
				// Below is similar to the macro "Get_active_window_keybd_layout":
				HWND active_window;
				target_keybd_layout = GetKeyboardLayout((active_window = GetForegroundWindow())\
					? GetWindowThreadProcessId(active_window, NULL) : 0); // When no foreground window, the script's own layout seems like the safest default.
				target_layout_has_altgr = LayoutHasAltGr(target_keybd_layout); // In the case of this else's "if", target_layout_has_altgr was already set properly higher above.
			}
			if (target_layout_has_altgr != LAYOUT_UNDETERMINED) // This layout's AltGr status is already known with certainty.
				do_detect_altgr = false; // So don't go through the detection steps here and other places later below.
			else
			{
				control_vk = g_os.IsWin2000orLater() ? VK_LCONTROL : VK_CONTROL;
				lcontrol_was_down = IsKeyDownAsync(control_vk);
				// Add extra detection of AltGr if hook is installed, which has been show to be useful for some
				// scripts where the other AltGr detection methods don't occur in a timely enough fashion.
				// The following method relies upon the fact that it's impossible for the hook to receive
				// events from the user while it's processing our keybd_event() here.  This is because
				// any physical keystrokes that happen to occur at the exact moment of our keybd_event()
				// will stay queued until the main event loop routes them to the hook via GetMessage().
				g_HookReceiptOfLControlMeansAltGr = aExtraInfo;
				// Thread-safe: g_HookReceiptOfLControlMeansAltGr isn't thread-safe, but by its very nature it probably
				// shouldn't be (ways to do it might introduce an unwarranted amount of complexity and performance loss
				// given that the odds of collision might be astronomically low in this case, and the consequences too
				// mild).  The whole point of g_HookReceiptOfLControlMeansAltGr and related altgr things below is to
				// watch what keystrokes the hook receives in response to simulating a press of the right-alt key.
				// Due to their global/system nature, keystrokes are never thread-safe in the sense that any process
				// in the entire system can be sending keystrokes simultaneously with ours.
			}
		}

		// Calculated only once for performance (and avoided entirely if not needed):
		modLR_type key_as_modifiersLR = put_event_into_array ? KeyToModifiersLR(aVK, aSC) : 0;

		bool do_key_history = !caller_is_keybd_hook // If caller is hook, don't log because it does logs it.
			&& sSendMode != SM_PLAY  // In playback mode, the journal hook logs so that timestamps are accurate.
			&& (!g_KeybdHook || sSendMode == SM_INPUT); // In the remaining cases, log only when the hook isn't installed or won't be at the time of the event.

		if (aEventType != KEYUP)  // i.e. always do it for KEYDOWNANDUP
		{
			if (put_event_into_array)
				PutKeybdEventIntoArray(key_as_modifiersLR, aVK, aSC, event_flags, aExtraInfo);
			else
			{
				// In v1.0.35.07, g_IgnoreNextLControlDown/Up was added.
				// The following global is used to flag as our own the keyboard driver's LControl-down keystroke
				// that is triggered by RAlt-down (AltGr).  This prevents it from triggering hotkeys such as
				// "*Control::".  It probably fixes other obscure side-effects and bugs also, since the
				// event should be considered script-generated even though indirect.  Note: The problem with
				// having the hook detect AltGr's automatic LControl-down is that the keyboard driver seems
				// to generate the LControl-down *before* notifying the system of the RAlt-down.  That makes
				// it impossible for the hook to flag the LControl keystroke in advance, so it would have to
				// retroactively undo the effects.  But that is impossible because the previous keystroke might
				// already have wrongly fired a hotkey.
				if (hookable_ralt && target_layout_has_altgr == CONDITION_TRUE)
					g_IgnoreNextLControlDown = aExtraInfo; // Must be set prior to keybd_event() to be effective.
				keybd_event(aVK, aSC_lobyte // naked scan code (the 0xE0 prefix, if any, is omitted)
					, event_flags, aExtraInfo);
				// The following is done by us rather than by the hook to avoid problems where:
				// 1) The hook is removed at a critical point during the operation, preventing the variable from
				//    being reset to false.
				// 2) For some reason this AltGr keystroke done above did not cause LControl to go down (perhaps
				//    because the current keyboard layout doesn't have AltGr as we thought), which would be a bug
				//    because some other Ctrl keystroke would then be wrongly ignored.
				g_IgnoreNextLControlDown = 0; // Unconditional reset.
				if (do_detect_altgr)
				{
					do_detect_altgr = false; // Indicate to the KEYUP section later below that detection has already been done.
					if (g_HookReceiptOfLControlMeansAltGr)
					{
						g_HookReceiptOfLControlMeansAltGr = 0; // Must reset promptly in case key-delay below routes physical keystrokes to hook.
						// The following line is multipurpose:
						// 1) Retrieves an updated value of target_layout_has_altgr in case the hook just changed it.
						// 2) If the hook didn't change it, the target keyboard layout doesn't have an AltGr key.
						//    Only in that case will the following line set it to FALSE (because LayoutHasAltGr only
						//    changes the value if it's currently undetermined).
						// Finally, this also updates sTargetLayoutHasAltGr in cases where target_layout_has_altgr
						// is an alias/reference for it.
						target_layout_has_altgr = LayoutHasAltGr(target_keybd_layout, CONDITION_FALSE);
					}
					else if (!lcontrol_was_down) // i.e. if LControl was already down, this detection method isn't possible.
						// Called this way, it updates the specified layout as long as keybd_event's call to the hook didn't already determine it to be FALSE or TRUE:
						target_layout_has_altgr = LayoutHasAltGr(target_keybd_layout, IsKeyDownAsync(control_vk) ? CONDITION_TRUE : CONDITION_FALSE);
						// Above also updates sTargetLayoutHasAltGr in cases where target_layout_has_altgr is an alias/reference for it.
				}
			}

#ifdef CONFIG_WIN9X
			if (aVK == VK_NUMLOCK && g_os.IsWin9x()) // Under Win9x, Numlock needs special treatment.
				ToggleNumlockWin9x();
#endif

			if (do_key_history)
				UpdateKeyEventHistory(false, aVK, aSC); // Should be thread-safe since if no hook means only one thread ever sends keystrokes (with possible exception of mouse hook, but that seems too rare).
		}
		// The press-duration delay is done only when this is a down-and-up because otherwise,
		// the normal g->KeyDelay will be in effect.  In other words, it seems undesirable in
		// most cases to do both delays for only "one half" of a keystroke:
		if (aDoKeyDelay && aEventType == KEYDOWNANDUP) // Hook should never specify a delay, so no need to check if caller is hook.
			DoKeyDelay(sSendMode == SM_PLAY ? g->PressDurationPlay : g->PressDuration); // DoKeyDelay() is not thread safe but since the hook thread should never pass true for aKeyDelay, it shouldn't be an issue.
		if (aEventType != KEYDOWN)  // i.e. always do it for KEYDOWNANDUP
		{
			event_flags |= KEYEVENTF_KEYUP;
			if (put_event_into_array)
				PutKeybdEventIntoArray(key_as_modifiersLR, aVK, aSC, event_flags, aExtraInfo);
			else
			{
				if (hookable_ralt && target_layout_has_altgr == CONDITION_TRUE) // See comments in similar section above for details.
					g_IgnoreNextLControlUp = aExtraInfo; // Must be set prior to keybd_event() to be effective.
				keybd_event(aVK, aSC_lobyte, event_flags, aExtraInfo);
				g_IgnoreNextLControlUp = 0; // Unconditional reset (see similar section above).
				if (do_detect_altgr) // This should be true only when aEventType==KEYUP because otherwise the KEYDOWN event above would have set it to false.
				{
					if (g_HookReceiptOfLControlMeansAltGr)
					{
						g_HookReceiptOfLControlMeansAltGr = 0; // Must reset promptly in case key-delay below routes physical keystrokes to hook.
						target_layout_has_altgr = LayoutHasAltGr(target_keybd_layout, CONDITION_FALSE); // See similar section above for comments.
					}
					else if (lcontrol_was_down) // i.e. if LControl was already up, this detection method isn't possible.
						// See similar section above for comments:
						target_layout_has_altgr = LayoutHasAltGr(target_keybd_layout, IsKeyDownAsync(control_vk) ? CONDITION_FALSE : CONDITION_TRUE);
				}
			}
			// The following is done to avoid an extraneous artificial {LCtrl Up} later on,
			// since the keyboard driver should insert one in response to this {RAlt Up}:
			if (target_layout_has_altgr && aSC == SC_RALT)
				sEventModifiersLR &= ~MOD_LCONTROL;

			if (do_key_history)
				UpdateKeyEventHistory(true, aVK, aSC);
		}

		if (we_turned_blockinput_off)  // Already made thread-safe by action higher above.
			Line::ScriptBlockInput(true);  // Turn BlockInput back on.
	}

	if (aDoKeyDelay) // SM_PLAY also uses DoKeyDelay(): it stores the delay item in the event array.
		DoKeyDelay(); // Thread-safe because only called by main thread in this mode.  See notes above.
}


///////////////////
// Mouse related //
///////////////////

ResultType PerformClick(LPTSTR aOptions)
{
	int x, y;
	vk_type vk;
	KeyEventTypes event_type;
	int repeat_count;
	bool move_offset;

	ParseClickOptions(aOptions, x, y, vk, event_type, repeat_count, move_offset);
	PerformMouseCommon(repeat_count < 1 ? ACT_MOUSEMOVE : ACT_MOUSECLICK // Treat repeat-count<1 as a move (like {click}).
		, vk, x, y, 0, 0, repeat_count, event_type, g->DefaultMouseSpeed, move_offset);

	return OK; // For caller convenience.
}



void ParseClickOptions(LPTSTR aOptions, int &aX, int &aY, vk_type &aVK, KeyEventTypes &aEventType
	, int &aRepeatCount, bool &aMoveOffset)
// Caller has trimmed leading whitespace from aOptions, but not necessarily the trailing whitespace.
// aOptions must be a modifiable string because this function temporarily alters it.
{
	// Set defaults for all output parameters for caller.
	aX = COORD_UNSPECIFIED;
	aY = COORD_UNSPECIFIED;
	aVK = VK_LBUTTON_LOGICAL; // v1.0.43: Logical vs. physical for {Click} and Click-cmd, in case user has buttons swapped via control panel.
	aEventType = KEYDOWNANDUP;
	aRepeatCount = 1;
	aMoveOffset = false;

	TCHAR *next_option, *option_end, orig_char;
	vk_type temp_vk;

	for (next_option = aOptions; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		// Allow optional commas to make scripts more readable.  Don't support g_delimiter because StrChrAny
		// below doesn't.
		while (*next_option == ',') // Ignore all commas.
			if (!*(next_option = omit_leading_whitespace(next_option + 1)))
				goto break_both; // Entire option string ends in a comma.
		// Find the end of this option item:
		if (   !(option_end = StrChrAny(next_option, _T(" \t,")))   )  // Space, tab, comma.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.

		// Temp termination for IsPureNumeric(), ConvertMouseButton(), and peace-of-mind.
		orig_char = *option_end;
		*option_end = '\0';

		// Parameters can occur in almost any order to enhance usability (at the cost of
		// slightly diminishing the ability unambiguously add more parameters in the future).
		// Seems okay to support floats because ATOI() will just omit the decimal portion.
		if (IsPureNumeric(next_option, true, false, true))
		{
			// Any numbers present must appear in the order: X, Y, RepeatCount
			// (optionally with other options between them).
			if (aX == COORD_UNSPECIFIED) // This will be converted into repeat-count if it is later discovered there's no Y coordinate.
				aX = ATOI(next_option);
			else if (aY == COORD_UNSPECIFIED)
				aY = ATOI(next_option);
			else // Third number is the repeat-count (but if there's only one number total, that's repeat count too, see further below).
				aRepeatCount = ATOI(next_option); // If negative or zero, caller knows to handle it as a MouseMove vs. Click.
		}
		else // Mouse button/name and/or Down/Up/Repeat-count is present.
		{
			if (temp_vk = Line::ConvertMouseButton(next_option, true, true))
				aVK = temp_vk;
			else
			{
				switch (ctoupper(*next_option))
				{
				case 'D': aEventType = KEYDOWN; break;
				case 'U': aEventType = KEYUP; break;
				case 'R': aMoveOffset = true; break; // Since it wasn't recognized as the right mouse button, it must have other letters after it, e.g. Rel/Relative.
				// default: Ignore anything else to reserve them for future use.
				}
			}
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	} // for() each item in option list

break_both:
	if (aX != COORD_UNSPECIFIED && aY == COORD_UNSPECIFIED)
	{
		// When only one number is present (e.g. {Click 2}, it's assumed to be the repeat count.
		aRepeatCount = aX;
		aX = COORD_UNSPECIFIED;
	}
}



ResultType PerformMouse(ActionTypeType aActionType, LPTSTR aButton, LPTSTR aX1, LPTSTR aY1, LPTSTR aX2, LPTSTR aY2
	, LPTSTR aSpeed, LPTSTR aOffsetMode, LPTSTR aRepeatCount, LPTSTR aDownUp)
{
	vk_type vk;
	if (aActionType == ACT_MOUSEMOVE)
		vk = 0;
	else
		// ConvertMouseButton() treats blank as "Left":
		if (   !(vk = Line::ConvertMouseButton(aButton, aActionType == ACT_MOUSECLICK))   )
			vk = VK_LBUTTON; // See below.
			// v1.0.43: Seems harmless (due to rarity) to treat invalid button names as "Left" (keeping in
			// mind that due to loadtime validation, invalid buttons are possible only when the button name is
			// contained in a variable, e.g. MouseClick %ButtonName%.

	KeyEventTypes event_type = KEYDOWNANDUP;  // Set defaults.
	int repeat_count = 1;                     //

	if (aActionType == ACT_MOUSECLICK)
	{
		if (*aRepeatCount)
			repeat_count = ATOI(aRepeatCount);
		switch(*aDownUp)
		{
		case 'u':
		case 'U':
			event_type = KEYUP;
			break;
		case 'd':
		case 'D':
			event_type = KEYDOWN;
			break;
		// Otherwise, leave it set to the default.
		}
	}

	PerformMouseCommon(aActionType, vk
		, *aX1 ? ATOI(aX1) : COORD_UNSPECIFIED  // If no starting coords are specified, mark it as "use the
		, *aY1 ? ATOI(aY1) : COORD_UNSPECIFIED  // current mouse position":
		, *aX2 ? ATOI(aX2) : COORD_UNSPECIFIED  // These two are blank except for dragging.
		, *aY2 ? ATOI(aY2) : COORD_UNSPECIFIED  //
		, repeat_count, event_type
		, *aSpeed ? ATOI(aSpeed) : g->DefaultMouseSpeed
		, ctoupper(*aOffsetMode) == 'R'); // aMoveOffset.

	return OK; // For caller convenience.
}



void PerformMouseCommon(ActionTypeType aActionType, vk_type aVK, int aX1, int aY1, int aX2, int aY2
	, int aRepeatCount, KeyEventTypes aEventType, int aSpeed, bool aMoveOffset)
{
	// The maximum number of events, which in this case would be from a MouseClickDrag.  To be conservative
	// (even though INPUT is a much larger struct than PlaybackEvent and SendInput doesn't use mouse-delays),
	// include room for the maximum number of mouse delays too.
	// Drag consists of at most:
	// 1) Move; 2) Delay; 3) Down; 4) Delay; 5) Move; 6) Delay; 7) Delay (dupe); 8) Up; 9) Delay.
	#define MAX_PERFORM_MOUSE_EVENTS 10
	INPUT event_array[MAX_PERFORM_MOUSE_EVENTS]; // Use type INPUT vs. PlaybackEvent since the former is larger (i.e. enough to hold either one).

	sSendMode = g->SendMode;
	if (sSendMode == SM_INPUT || sSendMode == SM_INPUT_FALLBACK_TO_PLAY)
	{
		if (!sMySendInput || SystemHasAnotherMouseHook()) // See similar section in SendKeys() for details.
			sSendMode = (sSendMode == SM_INPUT) ? SM_EVENT : SM_PLAY;
		else
			sSendMode = SM_INPUT; // Resolve early so that other sections don't have to consider SM_INPUT_FALLBACK_TO_PLAY a valid value.
	}
	if (sSendMode) // We're also responsible for setting sSendMode to SM_EVENT prior to returning.
		InitEventArray(event_array, MAX_PERFORM_MOUSE_EVENTS, 0);

	// Turn it on unconditionally even if it was on, since Ctrl-Alt-Del might have disabled it.
	// Turn it back off only if it wasn't ON before we started.
	bool blockinput_prev = g_BlockInput;
	bool do_selective_blockinput = (g_BlockInputMode == TOGGLE_MOUSE || g_BlockInputMode == TOGGLE_SENDANDMOUSE)
		&& !sSendMode && g_os.IsWinNT4orLater();
	if (do_selective_blockinput) // It seems best NOT to use g_BlockMouseMove for this, since often times the user would want keyboard input to be disabled too, until after the mouse event is done.
		Line::ScriptBlockInput(true); // Turn it on unconditionally even if it was on, since Ctrl-Alt-Del might have disabled it.

	switch (aActionType)
	{
	case ACT_MOUSEMOVE:
		DWORD unused;
		MouseMove(aX1, aY1, unused, aSpeed, aMoveOffset); // Does nothing if coords are invalid.
		break;
	case ACT_MOUSECLICK:
		MouseClick(aVK, aX1, aY1, aRepeatCount, aSpeed, aEventType, aMoveOffset); // Does nothing if coords are invalid.
		break;
	case ACT_MOUSECLICKDRAG:
		MouseClickDrag(aVK, aX1, aY1, aX2, aY2, aSpeed, aMoveOffset); // Does nothing if coords are invalid.
		break;
	} // switch()

	if (sSendMode)
	{
		int final_key_delay = -1; // Set default.
		if (!sAbortArraySend && sEventCount > 0)
			SendEventArray(final_key_delay, 0); // Last parameter is ignored because keybd hook isn't removed for a pure-mouse SendInput.
		CleanupEventArray(final_key_delay);
	}

	if (do_selective_blockinput && !blockinput_prev)  // Turn it back off only if it was off before we started.
		Line::ScriptBlockInput(false);
}



void MouseClickDrag(vk_type aVK, int aX1, int aY1, int aX2, int aY2, int aSpeed, bool aMoveOffset)
{
	// Check if one of the coordinates is missing, which can happen in cases where this was called from
	// a source that didn't already validate it. Can't call Line::ValidateMouseCoords() because that accepts strings.
	if (   (aX1 == COORD_UNSPECIFIED && aY1 != COORD_UNSPECIFIED) || (aX1 != COORD_UNSPECIFIED && aY1 == COORD_UNSPECIFIED)
		|| (aX2 == COORD_UNSPECIFIED && aY2 != COORD_UNSPECIFIED) || (aX2 != COORD_UNSPECIFIED && aY2 == COORD_UNSPECIFIED)   )
		return;

	// I asked Jon, "Have you discovered that insta-drags almost always fail?" and he said
	// "Yeah, it was weird, absolute lack of drag... Don't know if it was my config or what."
	// However, testing reveals "insta-drags" work ok, at least on my system, so leaving them enabled.
	// User can easily increase the speed if there's any problem:
	//if (aSpeed < 2)
	//	aSpeed = 2;

	// v1.0.43: Translate logical buttons into physical ones.  Which physical button it becomes depends
	// on whether the mouse buttons are swapped via the Control Panel.  Note that journal playback doesn't
	// need the swap because every aspect of it is "logical".
	if (aVK == VK_LBUTTON_LOGICAL)
		aVK = sSendMode != SM_PLAY && GetSystemMetrics(SM_SWAPBUTTON) ? VK_RBUTTON : VK_LBUTTON;
	else if (aVK == VK_RBUTTON_LOGICAL)
		aVK = sSendMode != SM_PLAY && GetSystemMetrics(SM_SWAPBUTTON) ? VK_LBUTTON : VK_RBUTTON;

	// MSDN: If [event_flags] is not MOUSEEVENTF_WHEEL, [MOUSEEVENTF_HWHEEL,] MOUSEEVENTF_XDOWN,
	// or MOUSEEVENTF_XUP, then [event_data] should be zero. 
	DWORD event_down, event_up, event_flags = 0, event_data = 0; // Set defaults for some.
	switch (aVK)
	{
	case VK_LBUTTON:
		event_down = MOUSEEVENTF_LEFTDOWN;
		event_up = MOUSEEVENTF_LEFTUP;
		break;
	case VK_RBUTTON:
		event_down = MOUSEEVENTF_RIGHTDOWN;
		event_up = MOUSEEVENTF_RIGHTUP;
		break;
	case VK_MBUTTON:
		event_down = MOUSEEVENTF_MIDDLEDOWN;
		event_up = MOUSEEVENTF_MIDDLEUP;
		break;
	case VK_XBUTTON1:
	case VK_XBUTTON2:
		event_down = MOUSEEVENTF_XDOWN;
		event_up = MOUSEEVENTF_XUP;
		event_data = (aVK == VK_XBUTTON1) ? XBUTTON1 : XBUTTON2;
		break;
	}

	// If the drag isn't starting at the mouse's current position, move the mouse to the specified position:
	if (aX1 != COORD_UNSPECIFIED && aY1 != COORD_UNSPECIFIED)
	{
		// The movement must be a separate event from the click, otherwise it's completely unreliable with
		// SendInput() and probably keybd_event() too.  SendPlay is unknown, but it seems best for
		// compatibility and peace-of-mind to do it for that too.  For example, some apps may be designed
		// to expect mouse movement prior to a click at a *new* position, which is not unreasonable given
		// that this would be the case 99.999% of the time if the user were moving the mouse physically.
		MouseMove(aX1, aY1, event_flags, aSpeed, aMoveOffset); // It calls DoMouseDelay() and also converts aX1 and aY1 to MOUSEEVENTF_ABSOLUTE coordinates.
		// v1.0.43: event_flags was added to improve reliability.  Explanation: Since the mouse was just moved to an
		// explicitly specified set of coordinates, use those coordinates with subsequent clicks.  This has been
		// shown to significantly improve reliability in cases where the user is moving the mouse during the
		// MouseClick/Drag commands.
	}
	MouseEvent(event_flags | event_down, event_data, aX1, aY1); // It ignores aX and aY when MOUSEEVENTF_MOVE is absent.
	DoMouseDelay(); // Inserts delay for all modes except SendInput, for which it does nothing.
	// Now that the mouse button has been pushed down, move the mouse to perform the drag:
	MouseMove(aX2, aY2, event_flags, aSpeed, aMoveOffset); // It calls DoMouseDelay() and also converts aX2 and aY2 to MOUSEEVENTF_ABSOLUTE coordinates.
	DoMouseDelay(); // Duplicate, see below.
	// Above is a *duplicate* delay because MouseMove() already did one. But it seems best to keep it because:
	// 1) MouseClickDrag can be a CPU intensive operation for the target window depending on what it does
	//    during the drag (selecting objects, etc.)  Thus, and extra delay might help a lot of things.
	// 2) It would probably break some existing scripts to remove the delay, due to timing issues.
	// 3) Dragging is pretty rarely used, so the added performance of removing the delay wouldn't be
	//    a big benefit.
	MouseEvent(event_flags | event_up, event_data, aX2, aY2); // It ignores aX and aY when MOUSEEVENTF_MOVE is absent.
	DoMouseDelay();
	// Above line: It seems best to always do this delay too in case the script line that
	// caused us to be called here is followed immediately by another script line which
	// is either another mouse click or something that relies upon this mouse drag
	// having been completed:
}



void MouseClick(vk_type aVK, int aX, int aY, int aRepeatCount, int aSpeed, KeyEventTypes aEventType
	, bool aMoveOffset)
{
	// Check if one of the coordinates is missing, which can happen in cases where this was called from
	// a source that didn't already validate it (such as MouseClick, %x%, %BlankVar%).
	// Allow aRepeatCount<1 to simply "do nothing", because it increases flexibility in the case where
	// the number of clicks is a dereferenced script variable that may sometimes (by intent) resolve to
	// zero or negative.  For backward compatibility, a RepeatCount <1 does not move the mouse (unlike
	// the Click command and Send {Click}).
	if (   (aX == COORD_UNSPECIFIED && aY != COORD_UNSPECIFIED) || (aX != COORD_UNSPECIFIED && aY == COORD_UNSPECIFIED)
		|| (aRepeatCount < 1)   )
		return;

	DWORD event_flags = 0; // Set default.

	if (!(aX == COORD_UNSPECIFIED || aY == COORD_UNSPECIFIED)) // Both coordinates were specified.
	{
		// The movement must be a separate event from the click, otherwise it's completely unreliable with
		// SendInput() and probably keybd_event() too.  SendPlay is unknown, but it seems best for
		// compatibility and peace-of-mind to do it for that too.  For example, some apps may be designed
		// to expect mouse movement prior to a click at a *new* position, which is not unreasonable given
		// that this would be the case 99.999% of the time if the user were moving the mouse physically.
		MouseMove(aX, aY, event_flags, aSpeed, aMoveOffset); // It calls DoMouseDelay() and also converts aX and aY to MOUSEEVENTF_ABSOLUTE coordinates.
		// v1.0.43: event_flags was added to improve reliability.  Explanation: Since the mouse was just moved to an
		// explicitly specified set of coordinates, use those coordinates with subsequent clicks.  This has been
		// shown to significantly improve reliability in cases where the user is moving the mouse during the
		// MouseClick/Drag commands.
	}
	// Above must be done prior to below because initial mouse-move is supported even for wheel turning.

	// For wheel turning, if the user activated this command via a hotkey, and that hotkey
	// has a modifier such as CTRL, the user is probably still holding down the CTRL key
	// at this point.  Therefore, there's some merit to the fact that we should release
	// those modifier keys prior to turning the mouse wheel (since some apps disable the
	// wheel or give it different behavior when the CTRL key is down -- for example, MSIE
	// changes the font size when you use the wheel while CTRL is down).  However, if that
	// were to be done, there would be no way to ever hold down the CTRL key explicitly
	// (via Send, {CtrlDown}) unless the hook were installed.  The same argument could probably
	// be made for mouse button clicks: modifier keys can often affect their behavior.  But
	// changing this function to adjust modifiers for all types of events would probably break
	// some existing scripts.  Maybe it can be a script option in the future.  In the meantime,
	// it seems best not to adjust the modifiers for any mouse events and just document that
	// behavior in the MouseClick command.
	switch (aVK)
	{
	case VK_WHEEL_UP:
		MouseEvent(event_flags | MOUSEEVENTF_WHEEL, aRepeatCount * WHEEL_DELTA, aX, aY);  // It ignores aX and aY when MOUSEEVENTF_MOVE is absent.
		return;
	case VK_WHEEL_DOWN:
		MouseEvent(event_flags | MOUSEEVENTF_WHEEL, -(aRepeatCount * WHEEL_DELTA), aX, aY);
		return;
	// v1.0.48: Lexikos: Support horizontal scrolling in Windows Vista and later.
	case VK_WHEEL_LEFT:
		MouseEvent(event_flags | MOUSEEVENTF_HWHEEL, -(aRepeatCount * WHEEL_DELTA), aX, aY);
		return;
	case VK_WHEEL_RIGHT:
		MouseEvent(event_flags | MOUSEEVENTF_HWHEEL, aRepeatCount * WHEEL_DELTA, aX, aY);
		return;
	}
	// Since above didn't return:

	// Although not thread-safe, the following static vars seem okay because:
	// 1) This function is currently only called by the main thread.
	// 2) Even if that isn't true, the serialized nature of simulated mouse clicks makes it likely that
	//    the statics will produce the correct behavior anyway.
	// 3) Even if that isn't true, the consequences of incorrect behavior seem minimal in this case.
	static vk_type sWorkaroundVK = 0;
	static LRESULT sWorkaroundHitTest; // Not initialized because the above will be the sole signal of whether the workaround is in progress.
	DWORD event_down, event_up, event_data = 0; // Set default.
	// MSDN: If [event_flags] is not MOUSEEVENTF_WHEEL, MOUSEEVENTF_XDOWN, or MOUSEEVENTF_XUP, then [event_data]
	// should be zero. 

	// v1.0.43: Translate logical buttons into physical ones.  Which physical button it becomes depends
	// on whether the mouse buttons are swapped via the Control Panel.
	if (aVK == VK_LBUTTON_LOGICAL)
		aVK = sSendMode != SM_PLAY && GetSystemMetrics(SM_SWAPBUTTON) ? VK_RBUTTON : VK_LBUTTON;
	else if (aVK == VK_RBUTTON_LOGICAL)
		aVK = sSendMode != SM_PLAY && GetSystemMetrics(SM_SWAPBUTTON) ? VK_LBUTTON : VK_RBUTTON;

	switch (aVK)
	{
	case VK_LBUTTON:
	case VK_RBUTTON:
		// v1.0.43 The first line below means: We're not in SendInput/Play mode or we are but this
		// will be the first event inside the array.  The latter case also implies that no initial
		// mouse-move was done above (otherwise there would already be a MouseMove event in the array,
		// and thus the click here wouldn't be the first item).  It doesn't seem necessary to support
		// the MouseMove case above because the workaround generally isn't needed in such situations
		// (see detailed comments below).  Furthermore, if the MouseMove were supported in array-mode,
		// it would require that GetCursorPos() below be conditionally replaced with something like
		// the following (since when in array-mode, the cursor hasn't actually moved *yet*):
		//		CoordToScreen(aX_orig, aY_orig, COORD_MODE_MOUSE);  // Moving mouse relative to the active window.
		// Known limitation: the work-around described below isn't as complete for SendPlay as it is
		// for the other modes: because dragging the title bar of one of this thread's windows with a
		// remap such as F1::LButton doesn't work if that remap uses SendPlay internally (the window
		// gets stuck to the mouse cursor).
		if (   (!sSendMode || !sEventCount) // See above.
			&& (aEventType == KEYDOWN || (aEventType == KEYUP && sWorkaroundVK))   ) // i.e. this is a down-only event or up-only event.
		{
			// v1.0.40.01: The following section corrects misbehavior caused by a thread sending
			// simulated mouse clicks to one of its own windows.  A script consisting only of the
			// following two lines can reproduce this issue:
			// F1::LButton
			// F2::RButton
			// The problems came about from the following sequence of events:
			// 1) Script simulates a left-click-down in the title bar's close, minimize, or maximize button.
			// 2) WM_NCLBUTTONDOWN is sent to the window's window proc, which then passes it on to
			//    DefWindowProc or DefDlgProc, which then apparently enters a loop in which no messages
			//    (or a very limited subset) are pumped.
			// 3) Thus, if the user presses a hotkey while the thread is in this state, that hotkey is
			//    queued/buffered until DefWindowProc/DefDlgProc exits its loop.
			// 4) But the buffered hotkey is the very thing that's supposed to exit the loop via sending a
			//    simulated left-click-up event.
			// 5) Thus, a deadlock occurs.
			// 6) A similar situation arises when a right-click-down is sent to the title bar or sys-menu-icon.
			//
			// The following workaround operates by suppressing qualified click-down events until the
			// corresponding click-up occurs, at which time the click-up is transformed into a down+up if the
			// click-up is still in the same cursor position as the down. It seems preferable to fix this here
			// rather than changing each window proc. to always respond to click-down rather vs. click-up
			// because that would make all of the script's windows behave in a non-standard way, possibly
			// producing side-effects and defeating other programs' attempts to interact with them.
			// (Thanks to Shimanov for this solution.)
			//
			// Remaining known limitations:
			// 1) Title bar buttons are not visibly in a pressed down state when a simulated click-down is sent
			//    to them.
			// 2) A window that should not be activated, such as AlwaysOnTop+Disabled, is activated anyway
			//    by SetForegroundWindowEx().  Not yet fixed due to its rarity and minimal consequences.
			// 3) A related problem for which no solution has been discovered (and perhaps it's too obscure
			//    an issue to justify any added code size): If a remapping such as "F1::LButton" is in effect,
			//    pressing and releasing F1 while the cursor is over a script window's title bar will cause the
			//    window to move slightly the next time the mouse is moved.
			// 4) Clicking one of the script's window's title bar with a key/button that has been remapped to
			//    become the left mouse button sometimes causes the button to get stuck down from the window's
			//    point of view.  The reasons are related to those in #1 above.  In both #1 and #2, the workaround
			//    is not at fault because it's not in effect then.  Instead, the issue is that DefWindowProc enters
			//    a non-msg-pumping loop while it waits for the user to drag-move the window.  If instead the user
			//    releases the button without dragging, the loop exits on its own after a 500ms delay or so.
			// 5) Obscure behavior caused by keyboard's auto-repeat feature: Use a key that's been remapped to
			//    become the left mouse button to click and hold the minimize button of one of the script's windows.
			//    Drag to the left.  The window starts moving.  This is caused by the fact that the down-click is
			//    suppressed, thus the remap's hotkey subroutine thinks the mouse button is down, thus its
			//    auto-repeat suppression doesn't work and it sends another click.
			POINT point;
			GetCursorPos(&point); // Assuming success seems harmless.
			// Despite what MSDN says, WindowFromPoint() appears to fetch a non-NULL value even when the
			// mouse is hovering over a disabled control (at least on XP).
			HWND child_under_cursor, parent_under_cursor;
			if (   (child_under_cursor = WindowFromPoint(point))
				&& (parent_under_cursor = GetNonChildParent(child_under_cursor)) // WM_NCHITTEST below probably requires parent vs. child.
				&& GetWindowThreadProcessId(parent_under_cursor, NULL) == g_MainThreadID   ) // It's one of our thread's windows.
			{
				LRESULT hit_test = SendMessage(parent_under_cursor, WM_NCHITTEST, 0, MAKELPARAM(point.x, point.y));
				if (   aVK == VK_LBUTTON && (hit_test == HTCLOSE || hit_test == HTMAXBUTTON // Title bar buttons: Close, Maximize.
						|| hit_test == HTMINBUTTON || hit_test == HTHELP) // Title bar buttons: Minimize, Help.
					|| aVK == VK_RBUTTON && (hit_test == HTCAPTION || hit_test == HTSYSMENU)   )
				{
					if (aEventType == KEYDOWN)
					{
						// Ignore this event and substitute for it: Activate the window when one
						// of its title bar buttons is down-clicked.
						sWorkaroundVK = aVK;
						sWorkaroundHitTest = hit_test;
						SetForegroundWindowEx(parent_under_cursor); // Try to reproduce customary behavior.
						// For simplicity, aRepeatCount>1 is ignored and DoMouseDelay() is not done.
						return;
					}
					else // KEYUP
					{
						if (sWorkaroundHitTest == hit_test) // To weed out cases where user clicked down on a button then released somewhere other than the button.
							aEventType = KEYDOWNANDUP; // Translate this click-up into down+up to make up for the fact that the down was previously suppressed.
						//else let the click-up occur in case it does something or user wants it.
					}
				}
			} // Work-around for sending mouse clicks to one of our thread's own windows.
		}
		// sWorkaroundVK is reset later below.

		// Since above didn't return, the work-around isn't in effect and normal click(s) will be sent:
		if (aVK == VK_LBUTTON)
		{
			event_down = MOUSEEVENTF_LEFTDOWN;
			event_up = MOUSEEVENTF_LEFTUP;
		}
		else // aVK == VK_RBUTTON
		{
			event_down = MOUSEEVENTF_RIGHTDOWN;
			event_up = MOUSEEVENTF_RIGHTUP;
		}
		break;

	case VK_MBUTTON:
		event_down = MOUSEEVENTF_MIDDLEDOWN;
		event_up = MOUSEEVENTF_MIDDLEUP;
		break;

	case VK_XBUTTON1:
	case VK_XBUTTON2:
		event_down = MOUSEEVENTF_XDOWN;
		event_up = MOUSEEVENTF_XUP;
		event_data = (aVK == VK_XBUTTON1) ? XBUTTON1 : XBUTTON2;
		break;
	} // switch()

	// For simplicity and possibly backward compatibility, LONG_OPERATION_INIT/UPDATE isn't done.
	// In addition, some callers might do it for themselves, at least when aRepeatCount==1.
	for (int i = 0; i < aRepeatCount; ++i)
	{
		if (aEventType != KEYUP) // It's either KEYDOWN or KEYDOWNANDUP.
		{
			// v1.0.43: Reliability is significantly improved by specifying the coordinates with the event (if
			// caller gave us coordinates).  This is mostly because of SetMouseDelay: In previously versions,
			// the delay between a MouseClick's move and its click allowed time for the user to move the mouse
			// away from the target position before the click was sent.
			MouseEvent(event_flags | event_down, event_data, aX, aY); // It ignores aX and aY when MOUSEEVENTF_MOVE is absent.
			// It seems best to always Sleep a certain minimum time between events
			// because the click-down event may cause the target app to do something which
			// changes the context or nature of the click-up event.  AutoIt3 has also been
			// revised to do this. v1.0.40.02: Avoid doing the Sleep between the down and up
			// events when the workaround is in effect because any MouseDelay greater than 10
			// would cause DoMouseDelay() to pump messages, which would defeat the workaround:
			if (!sWorkaroundVK)
				DoMouseDelay(); // Inserts delay for all modes except SendInput, for which it does nothing.
		}
		if (aEventType != KEYDOWN) // It's either KEYUP or KEYDOWNANDUP.
		{
			MouseEvent(event_flags | event_up, event_data, aX, aY); // It ignores aX and aY when MOUSEEVENTF_MOVE is absent.
			// It seems best to always do this one too in case the script line that caused
			// us to be called here is followed immediately by another script line which
			// is either another mouse click or something that relies upon the mouse click
			// having been completed:
			DoMouseDelay(); // Inserts delay for all modes except SendInput, for which it does nothing.
		}
	} // for()

	sWorkaroundVK = 0; // Reset this indicator in all cases except those for which above already returned.
}



void MouseMove(int &aX, int &aY, DWORD &aEventFlags, int aSpeed, bool aMoveOffset)
// This function also does DoMouseDelay() for the caller.
// This function converts aX and aY for the caller into MOUSEEVENTF_ABSOLUTE coordinates.
// The exception is when the playback mode is in effect, in which case such a conversion would be undesirable
// both here and by the caller.
// It also puts appropriate bit-flags into aEventFlags.
{
	// Most callers have already validated this, but some haven't.  Since it doesn't take too long to check,
	// do it here rather than requiring all callers to do (helps maintainability).
	if (aX == COORD_UNSPECIFIED || aY == COORD_UNSPECIFIED)
		return;

	if (sSendMode == SM_PLAY) // Journal playback mode.
	{
		// Mouse speed (aSpeed) is ignored for SendInput/Play mode: the mouse always moves instantaneously
		// (though in the case of playback-mode, MouseDelay still applies between each movement and click).
		// Playback-mode ignores mouse speed because the cases where the user would want to move the mouse more
		// slowly (such as a demo) seem too rare to justify the code size and complexity, especially since the
		// incremental-move code would have to be implemented in the hook itself to ensure reliability.  This is
		// because calling GetCursorPos() before playback begins could cause the mouse to wind up at the wrong
		// destination, especially if our thread is preempted while building the array (which would give the user
		// a chance to physically move the mouse before uninterruptibility begins).
		// For creating demonstrations that user slower mouse movement, the older MouseMove command can be used
		// in conjunction with BlockInput. This also applies to SendInput because it's conceivable that mouse
		// speed could be supported there (though it seems useless both visually and to improve reliability
		// because no mouse delays are possible within SendInput).
		//
		// MSG_OFFSET_MOUSE_MOVE is used to have the playback hook apply the offset (rather than doing it here
		// like is done for SendInput mode).  This adds flexibility in cases where a new window becomes active
		// during playback, or the active window changes position -- if that were to happen, the offset would
		// otherwise be wrong while CoordMode is Relative because the changes can only be observed and
		// compensated for during playback.
		PutMouseEventIntoArray(MOUSEEVENTF_MOVE | (aMoveOffset ? MSG_OFFSET_MOUSE_MOVE : 0)
			, 0, aX, aY); // The playback hook uses normal vs. MOUSEEVENTF_ABSOLUTE coordinates.
		DoMouseDelay();
		if (aMoveOffset)
		{
			// Now that we're done using the old values of aX and aY above, reset them to COORD_UNSPECIFIED
			// for the caller so that any subsequent clicks it does will be marked as "at current coordinates".
			aX = COORD_UNSPECIFIED;
			aY = COORD_UNSPECIFIED;
		}
		return; // Other parts below rely on this returning early to avoid converting aX/aY into MOUSEEVENTF_ABSOLUTE.
	}

	// The playback mode returned from above doesn't need these flags added because they're ignored for clicks:
	aEventFlags |= MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE; // Done here for caller, for easier maintenance.

	POINT cursor_pos;
	if (aMoveOffset)  // We're moving the mouse cursor relative to its current position.
	{
		if (sSendMode == SM_INPUT)
		{
			// Since GetCursorPos() can't be called to find out a future cursor position, use the position
			// tracked for SendInput (facilitates MouseClickDrag's R-option as well as Send{Click}'s).
			if (sSendInputCursorPos.x == COORD_UNSPECIFIED) // Initial/starting value hasn't yet been set.
				GetCursorPos(&sSendInputCursorPos); // Start it off where the cursor is now.
			aX += sSendInputCursorPos.x;
			aY += sSendInputCursorPos.y;
		}
		else
		{
			GetCursorPos(&cursor_pos); // None of this is done for playback mode since that mode already returned higher above.
			aX += cursor_pos.x;
			aY += cursor_pos.y;
		}
	}
	else
	{
		// Convert relative coords to screen coords if necessary (depends on CoordMode).
		// None of this is done for playback mode since that mode already returned higher above.
		CoordToScreen(aX, aY, COORD_MODE_MOUSE);
	}

	if (sSendMode == SM_INPUT) // Track predicted cursor position for use by subsequent events put into the array.
	{
		sSendInputCursorPos.x = aX; // Always stores normal coords (non-MOUSEEVENTF_ABSOLUTE).
		sSendInputCursorPos.y = aY; // 
	}

	// Find dimensions of primary monitor.
	// Without the MOUSEEVENTF_VIRTUALDESK flag (supported only by SendInput, and then only on
	// Windows 2000/XP or later), MOUSEEVENTF_ABSOLUTE coordinates are relative to the primary monitor.
	int screen_width = GetSystemMetrics(SM_CXSCREEN);
	int screen_height = GetSystemMetrics(SM_CYSCREEN);

	// Convert the specified screen coordinates to mouse event coordinates (MOUSEEVENTF_ABSOLUTE).
	// MSDN: "In a multimonitor system, [MOUSEEVENTF_ABSOLUTE] coordinates map to the primary monitor."
	// The above implies that values greater than 65535 or less than 0 are appropriate, but only on
	// multi-monitor systems.  For simplicity, performance, and backward compatibility, no check for
	// multi-monitor is done here. Instead, the system's default handling for out-of-bounds coordinates
	// is used; for example, mouse_event() stops the cursor at the edge of the screen.
	// UPDATE: In v1.0.43, the following formula was fixed (actually broken, see below) to always yield an
	// in-range value. The previous formula had a +1 at the end:
	// aX|Y = ((65535 * aX|Y) / (screen_width|height - 1)) + 1;
	// The extra +1 would produce 65536 (an out-of-range value for a single-monitor system) if the maximum
	// X or Y coordinate was input (e.g. x-position 1023 on a 1024x768 screen).  Although this correction
	// seems inconsequential on single-monitor systems, it may fix certain misbehaviors that have been reported
	// on multi-monitor systems. Update: according to someone I asked to test it, it didn't fix anything on
	// multimonitor systems, at least those whose monitors vary in size to each other.  In such cases, he said
	// that only SendPlay or DllCall("SetCursorPos") make mouse movement work properly.
	// FIX for v1.0.44: Although there's no explanation yet, the v1.0.43 formula is wrong and the one prior
	// to it was correct; i.e. unless +1 is present below, a mouse movement to coords near the upper-left corner of
	// the screen is typically off by one pixel (only the y-coordinate is affected in 1024x768 resolution, but
	// in other resolutions both might be affected).
	// v1.0.44.07: The following old formula has been replaced:
	// (((65535 * coord) / (width_or_height - 1)) + 1)
	// ... with the new one below.  This is based on numEric's research, which indicates that mouse_event()
	// uses the following inverse formula internally:
	// x_or_y_coord = (x_or_y_abs_coord * screen_width_or_height) / 65536
	#define MOUSE_COORD_TO_ABS(coord, width_or_height) (((65536 * coord) / width_or_height) + (coord < 0 ? -1 : 1))
	aX = MOUSE_COORD_TO_ABS(aX, screen_width);
	aY = MOUSE_COORD_TO_ABS(aY, screen_height);
	// aX and aY MUST BE SET UNCONDITIONALLY because the output parameters must be updated for caller.
	// The incremental-move section further below also needs it.

	if (aSpeed < 0)  // This can happen during script's runtime due to something like: MouseMove, X, Y, %VarContainingNegative%
		aSpeed = 0;  // 0 is the fastest.
	else
		if (aSpeed > MAX_MOUSE_SPEED)
			aSpeed = MAX_MOUSE_SPEED;
	if (aSpeed == 0 || sSendMode == SM_INPUT) // Instantaneous move to destination coordinates with no incremental positions in between.
	{
		// See the comments in the playback-mode section at the top of this function for why SM_INPUT ignores aSpeed.
		MouseEvent(MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE, 0, aX, aY);
		DoMouseDelay(); // Inserts delay for all modes except SendInput, for which it does nothing.
		return;
	}

	// Since above didn't return, use the incremental mouse move to gradually move the cursor until
	// it arrives at the destination coordinates.
	// Convert the cursor's current position to mouse event coordinates (MOUSEEVENTF_ABSOLUTE).
	GetCursorPos(&cursor_pos);
	DoIncrementalMouseMove(
		  MOUSE_COORD_TO_ABS(cursor_pos.x, screen_width)  // Source/starting coords.
		, MOUSE_COORD_TO_ABS(cursor_pos.y, screen_height) //
		, aX, aY, aSpeed);                                // Destination/ending coords.
}



void MouseEvent(DWORD aEventFlags, DWORD aData, DWORD aX, DWORD aY)
// Having this part outsourced to a function helps remember to use KEY_IGNORE so that our own mouse
// events won't be falsely detected as hotkeys by the hooks (if they are installed).
{
	if (sSendMode)
		PutMouseEventIntoArray(aEventFlags, aData, aX, aY);
	else
		mouse_event(aEventFlags
			, aX == COORD_UNSPECIFIED ? 0 : aX // v1.0.43.01: Must be zero if no change in position is desired
			, aY == COORD_UNSPECIFIED ? 0 : aY // (fixes compatibility with certain apps/games).
			, aData, KEY_IGNORE_LEVEL(g->SendLevel));
}


///////////////////////
// SUPPORT FUNCTIONS //
///////////////////////

void PutKeybdEventIntoArray(modLR_type aKeyAsModifiersLR, vk_type aVK, sc_type aSC, DWORD aEventFlags, DWORD aExtraInfo)
// This function is designed to be called from only one thread (the main thread) since it's not thread-safe.
// Playback hook only supports sending neutral modifiers.  Caller must ensure that any left/right modifiers
// such as VK_RCONTROL are translated into neutral (e.g. VK_CONTROL).
{
	bool key_up = aEventFlags & KEYEVENTF_KEYUP;
	// To make the SendPlay method identical in output to the other keystroke methods, have it generate
	// a leading down/up LControl event immediately prior to each RAlt event (with no key-delay).
	// This avoids having to add special handling to places like SetModifierLRState() to do AltGr things
	// differently when sending via playback vs. other methods.  The event order recorded by the journal
	// record hook is a little different than what the low-level keyboard hook sees, but I don't think
	// the order should matter in this case:
	//   sc  vk key	 msg
	//   138 12 Alt	 syskeydown (right vs. left scan code)
	//   01d 11 Ctrl keydown (left scan code) <-- In keyboard hook, normally this precedes Alt, not follows it. Seems inconsequential (testing confirms).
	//   01d 11 Ctrl keyup	(left scan code)
	//   138 12 Alt	 syskeyup (right vs. left scan code)
	// Check for VK_MENU not VK_RMENU because caller should have translated it to neutral:
	if (aVK == VK_MENU && aSC == SC_RALT && sTargetLayoutHasAltGr == CONDITION_TRUE && sSendMode == SM_PLAY)
		// Must pass VK_CONTROL rather than VK_LCONTROL because playback hook requires neutral modifiers.
		PutKeybdEventIntoArray(MOD_LCONTROL, VK_CONTROL, SC_LCONTROL, aEventFlags, aExtraInfo); // Recursive call to self.

	// Above must be done prior to the capacity check below because above might add a new array item.
	if (sEventCount == sMaxEvents) // Array's capacity needs expanding.
		if (!ExpandEventArray())
			return;

	// Keep track of the predicted modifier state for use in other places:
	if (key_up)
		sEventModifiersLR &= ~aKeyAsModifiersLR;
	else
		sEventModifiersLR |= aKeyAsModifiersLR;

	if (sSendMode == SM_INPUT)
	{
		INPUT &this_event = sEventSI[sEventCount]; // For performance and convenience.
		this_event.type = INPUT_KEYBOARD;
		this_event.ki.wVk = aVK;
		this_event.ki.wScan = (aEventFlags & KEYEVENTF_UNICODE) ? aSC : LOBYTE(aSC);
		this_event.ki.dwFlags = aEventFlags;
		this_event.ki.dwExtraInfo = aExtraInfo; // Although our hook won't be installed (or won't detect, in the case of playback), that of other scripts might be, so set this for them.
		this_event.ki.time = 0; // Let the system provide its own timestamp, which might be more accurate for individual events if this will be a very long SendInput.
		sHooksToRemoveDuringSendInput |= HOOK_KEYBD; // Presence of keyboard hook defeats uninterruptibility of keystrokes.
	}
	else // Playback hook.
	{
		PlaybackEvent &this_event = sEventPB[sEventCount]; // For performance and convenience.
		if (!(aVK || aSC)) // Caller is signaling that aExtraInfo contains a delay/sleep event.
		{
			// Although delays at the tail end of the playback array can't be implemented by the playback
			// itself, caller wants them put in too.
			this_event.message = 0; // Message number zero flags it as a delay rather than an actual event.
			this_event.time_to_wait = aExtraInfo;
		}
		else // A normal (non-delay) event for playback.
		{
			// By monitoring incoming events in a message/event loop, the following key combinations were
			// confirmed to be WM_SYSKEYDOWN vs. WM_KEYDOWN (up events weren't tested, so are assumed to
			// be the same as down-events):
			// Alt+Win
			// Alt+Shift
			// Alt+Capslock/Numlock/Scrolllock
			// Alt+AppsKey
			// Alt+F2/Delete/Home/End/Arrow/BS
			// Alt+Space/Enter
			// Alt+Numpad (tested all digits & most other keys, with/without Numlock ON)
			// F10 (by itself) / Win+F10 / Alt+F10 / Shift+F10 (but not Ctrl+F10)
			// By contrast, the following are not SYS: Alt+Ctrl, Alt+Esc, Alt+Tab (the latter two
			// are never received by msg/event loop probably because the system intercepts them).
			// So the rule appears to be: It's a normal (non-sys) key if Alt isn't down and the key
			// isn't F10, or if Ctrl is down. Though a press of the Alt key itself is a syskey unless Ctrl is down.
			// Update: The release of ALT is WM_KEYUP vs. WM_SYSKEYUP when it modified at least one key while it was down.
			if (sEventModifiersLR & (MOD_LCONTROL | MOD_RCONTROL) // Control is down...
				|| !(sEventModifiersLR & (MOD_LALT | MOD_RALT))   // ... or: Alt isn't down and this key isn't Alt or F10...
					&& aVK != VK_F10 && !(aKeyAsModifiersLR & (MOD_LALT | MOD_RALT))
				|| (sEventModifiersLR & (MOD_LALT | MOD_RALT)) && key_up) // ... or this is the release of Alt (for simplicity, assume that Alt modified something while it was down).
				this_event.message = key_up ? WM_KEYUP : WM_KEYDOWN;
			else
				this_event.message = key_up ? WM_SYSKEYUP : WM_SYSKEYDOWN;
			this_event.vk = aVK;
			this_event.sc = aSC; // Don't omit the extended-key-bit because it is used later on.
		}
	}
	++sEventCount;
}



void PutMouseEventIntoArray(DWORD aEventFlags, DWORD aData, DWORD aX, DWORD aY)
// This function is designed to be called from only one thread (the main thread) since it's not thread-safe.
// If the array-type is journal playback, caller should include MOUSEEVENTF_ABSOLUTE in aEventFlags if the
// the mouse coordinates aX and aY are relative to the screen rather than the active window.
{
	if (sEventCount == sMaxEvents) // Array's capacity needs expanding.
		if (!ExpandEventArray())
			return;

	if (sSendMode == SM_INPUT)
	{
		INPUT &this_event = sEventSI[sEventCount]; // For performance and convenience.
		this_event.type = INPUT_MOUSE;
		this_event.mi.dx = (aX == COORD_UNSPECIFIED) ? 0 : aX; // v1.0.43.01: Must be zero if no change in position is
		this_event.mi.dy = (aY == COORD_UNSPECIFIED) ? 0 : aY; // desired (fixes compatibility with certain apps/games).
		this_event.mi.dwFlags = aEventFlags;
		this_event.mi.mouseData = aData;
		this_event.mi.dwExtraInfo = KEY_IGNORE_LEVEL(g->SendLevel); // Although our hook won't be installed (or won't detect, in the case of playback), that of other scripts might be, so set this for them.
		this_event.mi.time = 0; // Let the system provide its own timestamp, which might be more accurate for individual events if this will be a very long SendInput.
		sHooksToRemoveDuringSendInput |= HOOK_MOUSE; // Presence of mouse hook defeats uninterruptibility of mouse clicks/moves.
	}
	else // Playback hook.
	{
		// Note: Delay events (sleeps), which are supported in playback mode but not SendInput, are always inserted
		// via PutKeybdEventIntoArray() rather than this function.
		PlaybackEvent &this_event = sEventPB[sEventCount]; // For performance and convenience.
		// Determine the type of event specified by caller, but also omit MOUSEEVENTF_MOVE so that the
		// follow variations can be differentiated:
		// 1) MOUSEEVENTF_MOVE by itself.
		// 2) MOUSEEVENTF_MOVE with a click event or wheel turn (in this case MOUSEEVENTF_MOVE is permitted but
		//    not required, since all mouse events in playback mode must have explicit coordinates at the
		//    time they're played back).
		// 3) A click event or wheel turn by itself (same remark as above).
		// Bits are isolated in what should be a future-proof way (also omits MSG_OFFSET_MOUSE_MOVE bit).
		switch (aEventFlags & (0x1FFF & ~MOUSEEVENTF_MOVE)) // v1.0.48: 0x1FFF vs. 0xFFF to support MOUSEEVENTF_HWHEEL.
		{
		case 0:                      this_event.message = WM_MOUSEMOVE; break; // It's a movement without a click.
		// In cases other than the above, it's a click or wheel turn with optional WM_MOUSEMOVE too.
		case MOUSEEVENTF_LEFTDOWN:   this_event.message = WM_LBUTTONDOWN; break;
		case MOUSEEVENTF_LEFTUP:     this_event.message = WM_LBUTTONUP; break;
		case MOUSEEVENTF_RIGHTDOWN:  this_event.message = WM_RBUTTONDOWN; break;
		case MOUSEEVENTF_RIGHTUP:    this_event.message = WM_RBUTTONUP; break;
		case MOUSEEVENTF_MIDDLEDOWN: this_event.message = WM_MBUTTONDOWN; break;
		case MOUSEEVENTF_MIDDLEUP:   this_event.message = WM_MBUTTONUP; break;
		case MOUSEEVENTF_XDOWN:      this_event.message = WM_XBUTTONDOWN; break;
		case MOUSEEVENTF_XUP:        this_event.message = WM_XBUTTONUP; break;
		case MOUSEEVENTF_WHEEL:      this_event.message = WM_MOUSEWHEEL; break;
		case MOUSEEVENTF_HWHEEL:     this_event.message = WM_MOUSEHWHEEL; break; // v1.0.48
		// WHEEL: No info comes into journal-record about which direction the wheel was turned (nor by how many
		// notches).  In addition, it appears impossible to specify such info when playing back the event.
		// Therefore, playback usually produces downward wheel movement (but upward in some apps like
		// Visual Studio).
		}
		// COORD_UNSPECIFIED_SHORT is used so that the very first event can be a click with unspecified
		// coordinates: it seems best to have the cursor's position fetched during playback rather than
		// here because if done here, there might be time for the cursor to move physically before
		// playback begins (especially if our thread is preempted while building the array).
		this_event.x = (aX == COORD_UNSPECIFIED) ? COORD_UNSPECIFIED_SHORT : (WORD)aX;
		this_event.y = (aY == COORD_UNSPECIFIED) ? COORD_UNSPECIFIED_SHORT : (WORD)aY;
		if (aEventFlags & MSG_OFFSET_MOUSE_MOVE) // Caller wants this event marked as a movement relative to cursor's current position.
			this_event.message |= MSG_OFFSET_MOUSE_MOVE;
	}
	++sEventCount;
}



ResultType ExpandEventArray()
// Returns OK or FAIL.
{
	#define EVENT_EXPANSION_MULTIPLIER 2  // Should be very rare for array to need to expand more than a few times.
	size_t event_size = (sSendMode == SM_INPUT) ? sizeof(INPUT) : sizeof(PlaybackEvent);
	void *new_mem;
	// SendInput() appears to be limited to 5000 chars (10000 events in array), at least on XP.  This is
	// either an undocumented SendInput limit or perhaps it's due to the system setting that determines
	// how many messages can get backlogged in each thread's msg queue before some start to get dropped.
	// Note that SendInput()'s return value always seems to indicate that all the characters were sent
	// even when the ones beyond the limit were clearly never received by the target window.
	// In any case, it seems best not to restrict to 5000 here in case the limit can vary for any reason.
	// The 5000 limit is documented in the help file.
	if (   !(new_mem = malloc(EVENT_EXPANSION_MULTIPLIER * sMaxEvents * event_size))   )
		sAbortArraySend = true; // Usually better to send nothing rather than partial.
		// And continue on below to free the old block, if appropriate.
	else // Copy old array into new memory area (note that sEventSI and sEventPB are different views of the same variable).
		memcpy(new_mem, sEventSI, sEventCount * event_size);
	if (sMaxEvents > (sSendMode == SM_INPUT ? MAX_INITIAL_EVENTS_SI : MAX_INITIAL_EVENTS_PB))
		free(sEventSI); // Previous block was malloc'd vs. _alloc'd, so free it.
	if (sAbortArraySend)
		return FAIL;
	sEventSI = (LPINPUT)new_mem; // Note that sEventSI and sEventPB are different views of the same variable.
	sMaxEvents *= EVENT_EXPANSION_MULTIPLIER;
	return OK;
}



void InitEventArray(void *aMem, UINT aMaxEvents, modLR_type aModifiersLR)
{
	sEventPB = (PlaybackEvent *)aMem; // Sets sEventSI too, since both are the same.
	sMaxEvents = aMaxEvents;
	sEventModifiersLR = aModifiersLR;
	sSendInputCursorPos.x = COORD_UNSPECIFIED;
	sSendInputCursorPos.y = COORD_UNSPECIFIED;
	sHooksToRemoveDuringSendInput = 0;
	sEventCount = 0;
	sAbortArraySend = false; // If KeyEvent() ever sets it to true, that allows us to send nothing at all rather than a partial send.
	sFirstCallForThisEvent = true;
	// The above isn't a local static inside PlaybackProc because PlaybackProc might get aborted in the
	// middle of a NEXT/SKIP pair by user pressing Ctrl-Esc, etc, which would make it unreliable.
}



void SendEventArray(int &aFinalKeyDelay, modLR_type aModsDuringSend)
// Caller must avoid calling this function if sMySendInput is NULL.
// aFinalKeyDelay (which the caller should have initialized to -1 prior to calling) may be changed here
// to the desired/final delay.  Caller must not act upon it until changing sTypeOfHookToBuild to something
// that will allow DoKeyDelay() to do a real delay.
{
	if (sSendMode == SM_INPUT)
	{
		// Remove hook(s) temporarily because the presence of low-level (LL) keybd hook completely disables
		// the uninterruptibility of SendInput's keystrokes (but the mouse hook doesn't affect them).
		// The converse is also true.  This was tested via:
		//	#space::
		//	SendInput {Click 400, 400, 100}
		//	MsgBox
		//	ExitApp
		// ... and also with BurnK6 running, a CPU maxing utility.  The mouse clicks were sent directly to the
		// BurnK6 window, and were pretty slow, and also I could physically move the mouse cursor a little
		// between each of sendinput's mouse clicks.  But removing the mouse-hook during SendInputs solves all that.
		// Rather than removing both hooks unconditionally, it's better to
		// remove only those that actually have corresponding events in the array.  This avoids temporarily
		// losing visibility of physical key states (especially when the keyboard hook is removed).
		HookType active_hooks;
		if (active_hooks = GetActiveHooks())
			AddRemoveHooks(active_hooks & ~sHooksToRemoveDuringSendInput, true);

		// Caller has ensured that sMySendInput isn't NULL.
		sMySendInput(sEventCount, sEventSI, sizeof(INPUT)); // Must call dynamically-resolved version for Win95/NT compatibility.
		// The return value is ignored because it never seems to be anything other than sEventCount, even if
		// the Send seems to partially fail (e.g. due to hitting 5000 event maximum).
		// Typical speed of SendInput: 10ms or less for short sends (under 100 events).
		// Typically 30ms for 500 events; and typically no more than 200ms for 5000 events (which is
		// the apparent max).
		// Testing shows that when SendInput returns to its caller, all of its key states are in effect
		// even if the target window hasn't yet had time to receive them all.  For example, the
		// following reports that LShift is down:
		//   SendInput {a 4900}{LShift down}
		//   MsgBox % GetKeyState("LShift")
		// Furthermore, if the user manages to physically press or release a key during the call to
		// SendInput, testing shows that such events are in effect immediately when SendInput returns
		// to its caller, perhaps because SendInput clears out any backlog of physical keystrokes prior to
		// returning, or perhaps because the part of the OS that updates key states is a very high priority.

		if (active_hooks)
		{
			if (active_hooks & sHooksToRemoveDuringSendInput & HOOK_KEYBD) // Keyboard hook was actually removed during SendInput.
			{
				// The above call to SendInput() has not only sent its own events, it has also emptied
				// the buffer of any events generated outside but during the SendInput.  Since such
				// events are almost always physical events rather than simulated ones, it seems to do
				// more good than harm on average to consider any such changes to be physical.
				// The g_PhysicalKeyState array is also updated by GetModifierLRState(true), but only
				// for the modifier keys, not for all keys on the keyboard.  Even if adjust all keys
				// is possible, it seems overly complex and it might impact performance more than it's
				// worth given the rarity of the user changing physical key states during a SendInput
				// and then wanting to explicitly retrieve that state via GetKeyState(Key, "P").
				modLR_type mods_current = GetModifierLRState(true); // This also serves to correct the hook's logical modifiers, since hook was absent during the SendInput.
				modLR_type mods_changed_physically_during_send = aModsDuringSend ^ mods_current;
				g_modifiersLR_physical &= ~(mods_changed_physically_during_send & aModsDuringSend); // Remove those that changed from down to up.
				g_modifiersLR_physical |= mods_changed_physically_during_send & mods_current; // Add those that changed from up to down.
				g_HShwnd = GetForegroundWindow(); // An item done by ResetHook() that seems worthwhile here.
				// Most other things done by ResetHook() seem like they would do more harm than good to reset here
				// because of the the time the hook is typically missing is very short, usually under 30ms.
			}
			AddRemoveHooks(active_hooks, true); // Restore the hooks that were active before the SendInput.
		}
		return;
	}

	// Since above didn't return, sSendMode == SM_PLAY.
	// It seems best not to call IsWindowHung() here because:
	// 1) It might improve script reliability to allow playback to a hung window because even though
	//    the entire system would appear frozen, if the window becomes unhung, the keystrokes would
	//    eventually get sent to it as intended (and the script may be designed to rely on this).
	//    Furthermore, the user can press Ctrl-Alt-Del or Ctrl-Esc to unfreeze the system.
	// 2) It might hurt performance.
	//
	// During journal playback, it appears that LL hook receives events in realtime; its just that
	// keystrokes the hook passes through (or generates itself) don't actually hit the active window
	// until after the playback is done.  Preliminary testing shows that the hook's disguise of Alt/Win
	// still function properly for Win/Alt hotkeys that use the playback method.
	sCurrentEvent = 0; // Reset for use by the hook below.  Should be done BEFORE the hook is installed in the next line.
#ifdef JOURNAL_RECORD_MODE
// To record and analyze events via the above:
// - Uncomment the line that defines this in the header file.
// - Put breakpoint after the hook removes itself (a few lines below).  Don't try to put breakpoint in RECORD hook
//   itself because it tends to freeze keyboard input (must press Ctrl-Alt-Del or Ctrl-Esc to unfreeze).
// - Have the script send a keystroke (best to use non-character keystroke such as SendPlay {Shift}).
// - It is now recording, so press the desired keys.
// - Press Ctrl+Break, Ctrl-Esc, or Ctrl-Alt-Del to stop recording (which should then hit breakpoint below).
// - Study contents of the sEventPB array, which contains the keystrokes just recorded.
	sEventCount = 0; // Used by RecordProc().
	if (   !(g_PlaybackHook = SetWindowsHookEx(WH_JOURNALRECORD, RecordProc, g_hInstance, 0))   )
		return;
#else
	if (   !(g_PlaybackHook = SetWindowsHookEx(WH_JOURNALPLAYBACK, PlaybackProc, g_hInstance, 0))   )
		return;
	// During playback, have the keybd hook (if it's installed) block presses of the Windows key.
	// This is done because the Windows key is about the only key (other than Ctrl-Esc or Ctrl-Alt-Del)
	// that takes effect immediately rather than getting buffered/postponed until after the playback.
	// It should be okay to set this after the playback hook is installed because playback shouldn't
	// actually begin until we have our thread do its first MsgSleep later below.
	g_BlockWinKeys = true;
#endif

	// Otherwise, hook is installed, so:
	// Wait for the hook to remove itself because the script should not be allowed to continue
	// until the Send finishes.
	// GetMessage(single_msg_filter) can't be used because then our thread couldn't playback
	// keystrokes to one of its own windows.  In addition, testing shows that it wouldn't
	// measurably improve performance anyway.
	// Note: User can't activate tray icon with mouse (because mouse is blocked), so there's
	// no need to call our main event loop merely to have the tray menu responsive.
	// Sleeping for 0 performs at least 15% worse than INTERVAL_UNSPECIFIED. I think this is
	// because the journal playback hook can operate only when this thread is in a message-pumping
	// state, and message pumping is far more efficient with GetMessage than PeekMessage.
	// Also note that both registered and hook hotkeys are noticed/caught during journal playback
	// (confirmed through testing).  However, they are kept buffered until the Send finishes
	// because ACT_SEND and such are designed to be uninterruptible by other script threads;
	// also, it would be undesirable in almost any conceivable case.
	//
	// Use a loop rather than a single call to MsgSleep(WAIT_FOR_MESSAGES) because
	// WAIT_FOR_MESSAGES is designed only for use by WinMain().  The loop doesn't measurably
	// affect performance because there used to be the following here in place of the loop,
	// and it didn't perform any better:
	// GetMessage(&msg, NULL, WM_CANCELJOURNAL, WM_CANCELJOURNAL);
	while (g_PlaybackHook)
		SLEEP_WITHOUT_INTERRUPTION(INTERVAL_UNSPECIFIED); // For maintainability, macro is used rather than optimizing/splitting the code it contains.
	g_BlockWinKeys = false;
	// Either the hook unhooked itself or the OS did due to Ctrl-Esc or Ctrl-Alt-Del.
	// MSDN: When an application sees a [system-generated] WM_CANCELJOURNAL message, it can assume
	// two things: the user has intentionally cancelled the journal record or playback mode,
	// and the system has already unhooked any journal record or playback hook procedures.
	if (!sEventPB[sEventCount - 1].message) // Playback hook can't do the final delay, so we do it here.
		// Don't do delay right here because the delay would be put into the array instead.
		aFinalKeyDelay = sEventPB[sEventCount - 1].time_to_wait;
	// GetModifierLRState(true) is not done because keystrokes generated by the playback hook
	// aren't really keystrokes in the sense that they affect global key state or modifier state.
	// They affect only the keystate retrieved when the target thread calls GetKeyState()
	// (GetAsyncKeyState can never see such changes, even if called from the target thread).
	// Furthermore, the hook (if present) continues to operate during journal playback, so it
	// will keep its own modifiers up-to-date if any physical or simulate keystrokes happen to
	// come in during playback (such keystrokes arrive in the hook in real time, but they don't
	// actually hit the active window until the playback finishes).
}



void CleanupEventArray(int aFinalKeyDelay)
{
	if (sMaxEvents > (sSendMode == SM_INPUT ? MAX_INITIAL_EVENTS_SI : MAX_INITIAL_EVENTS_PB))
		free(sEventSI); // Previous block was malloc'd vs. _alloc'd, so free it.  Note that sEventSI and sEventPB are different views of the same variable.
	// The following must be done only after functions called above are done using it.  But it must also be done
	// prior to our caller toggling capslock back on , to avoid the capslock keystroke from going into the array.
	sSendMode = SM_EVENT;
	DoKeyDelay(aFinalKeyDelay); // Do this only after resetting sSendMode above.  Should be okay for mouse events too.
}



/////////////////////////////////


void DoKeyDelay(int aDelay)
// Doesn't need to be thread safe because it should only ever be called from main thread.
{
	if (aDelay < 0) // To support user-specified KeyDelay of -1 (fastest send rate).
		return;
	if (sSendMode)
	{
		if (sSendMode == SM_PLAY && aDelay > 0) // Zero itself isn't supported by playback hook, so no point in inserting such delays into the array.
			PutKeybdEventIntoArray(0, 0, 0, 0, aDelay); // Passing zero for vk and sc signals it that aExtraInfo contains the delay.
		//else for other types of arrays, never insert a delay or do one now.
		return;
	}
	if (g_os.IsWin9x())
	{
		// Do a true sleep on Win9x because the MsgSleep() method is very inaccurate on Win9x
		// for some reason (a MsgSleep(1) causes a sleep between 10 and 55ms, for example):
		Sleep(aDelay);
		return;
	}
	SLEEP_WITHOUT_INTERRUPTION(aDelay);
}



void DoMouseDelay() // Helper function for the mouse functions below.
{
	int mouse_delay = sSendMode == SM_PLAY ? g->MouseDelayPlay : g->MouseDelay;
	if (mouse_delay < 0) // To support user-specified KeyDelay of -1 (fastest send rate).
		return;
	if (sSendMode)
	{
		if (sSendMode == SM_PLAY && mouse_delay > 0) // Zero itself isn't supported by playback hook, so no point in inserting such delays into the array.
			PutKeybdEventIntoArray(0, 0, 0, 0, mouse_delay); // Passing zero for vk and sc signals it that aExtraInfo contains the delay.
		//else for other types of arrays, never insert a delay or do one now (caller should have already
		// checked that, so it's written this way here only for maintainability).
		return;
	}
	// I believe the varying sleep methods below were put in place to avoid issues when simulating
	// clicks on the script's own windows.  There are extensive comments in MouseClick() and the
	// hook about these issues.  Also, Sleep() is more accurate on Win9x than MsgSleep, which is
	// why it's used in that case.  Here are more details from an older comment:
	// Always sleep a certain minimum amount of time between events to improve reliability,
	// but allow the user to specify a higher time if desired.  Note that for Win9x,
	// a true Sleep() is done because it is much more accurate than the MsgSleep() method,
	// at least on Win98SE when short sleeps are done.  UPDATE: A true sleep is now done
	// unconditionally if the delay period is small.  This fixes a small issue where if
	// LButton is a hotkey that includes "MouseClick left" somewhere in its subroutine,
	// the script's own main window's title bar buttons for min/max/close would not
	// properly respond to left-clicks.  By contrast, the following is no longer an issue
	// due to the dedicated thread in v1.0.39 (or more likely, due to an older change that
	// causes the tray menu to open upon RButton-up rather than down):
	// RButton is a hotkey that includes "MouseClick right" somewhere in its subroutine,
	// the user would not be able to correctly open the script's own tray menu via
	// right-click (note that this issue affected only the one script itself, not others).
	if (mouse_delay < 11 || (mouse_delay < 25 && g_os.IsWin9x()))
		Sleep(mouse_delay);
	else
		SLEEP_WITHOUT_INTERRUPTION(mouse_delay)
}



void UpdateKeyEventHistory(bool aKeyUp, vk_type aVK, sc_type aSC)
{
	if (!g_KeyHistory) // Don't access the array if it doesn't exist (i.e. key history is disabled).
		return;
	g_KeyHistory[g_KeyHistoryNext].key_up = aKeyUp;
	g_KeyHistory[g_KeyHistoryNext].vk = aVK;
	g_KeyHistory[g_KeyHistoryNext].sc = aSC;
	g_KeyHistory[g_KeyHistoryNext].event_type = 'i'; // Callers all want this.
	g_HistoryTickNow = GetTickCount();
	g_KeyHistory[g_KeyHistoryNext].elapsed_time = (g_HistoryTickNow - g_HistoryTickPrev) / (float)1000;
	g_HistoryTickPrev = g_HistoryTickNow;
	HWND fore_win = GetForegroundWindow();
	if (fore_win)
	{
		if (fore_win != g_HistoryHwndPrev)
			GetWindowText(fore_win, g_KeyHistory[g_KeyHistoryNext].target_window, _countof(g_KeyHistory[g_KeyHistoryNext].target_window));
		else // i.e. avoid the call to GetWindowText() if possible.
			*g_KeyHistory[g_KeyHistoryNext].target_window = '\0';
	}
	else
		_tcscpy(g_KeyHistory[g_KeyHistoryNext].target_window, _T("N/A"));
	g_HistoryHwndPrev = fore_win; // Update unconditionally in case it's NULL.
	if (++g_KeyHistoryNext >= g_MaxHistoryKeys)
		g_KeyHistoryNext = 0;
}



ToggleValueType ToggleKeyState(vk_type aVK, ToggleValueType aToggleValue)
// Toggle the given aVK into another state.  For performance, it is the caller's responsibility to
// ensure that aVK is a toggleable key such as capslock, numlock, insert, or scrolllock.
// Returns the state the key was in before it was changed (but it's only a best-guess under Win9x).
{
	// Can't use IsKeyDownAsync/GetAsyncKeyState() because it doesn't have this info:
	ToggleValueType starting_state = IsKeyToggledOn(aVK) ? TOGGLED_ON : TOGGLED_OFF;
	if (aToggleValue != TOGGLED_ON && aToggleValue != TOGGLED_OFF) // Shouldn't be called this way.
		return starting_state;
	if (starting_state == aToggleValue) // It's already in the desired state, so just return the state.
		return starting_state;
	if (aVK == VK_NUMLOCK)
	{
#ifdef CONFIG_WIN9X
		if (g_os.IsWin9x())
		{
			// For Win9x, we want to set the state unconditionally to be sure it's right.  This is because
			// the retrieval of the Capslock state, for example, is unreliable, at least under Win98se
			// (probably due to lack of an AttachThreadInput() having been done).  Although the
			// SetKeyboardState() method used by ToggleNumlockWin9x is not required for caps & scroll lock keys,
			// it is required for Numlock:
			ToggleNumlockWin9x();
			return starting_state;  // Best guess, but might be wrong.
		}
#endif
		// Otherwise, NT/2k/XP:
		// Sending an extra up-event first seems to prevent the problem where the Numlock
		// key's indicator light doesn't change to reflect its true state (and maybe its
		// true state doesn't change either).  This problem tends to happen when the key
		// is pressed while the hook is forcing it to be either ON or OFF (or it suppresses
		// it because it's a hotkey).  Needs more testing on diff. keyboards & OSes:
		KeyEvent(KEYUP, aVK);
	}
	// Since it's not already in the desired state, toggle it:
	KeyEvent(KEYDOWNANDUP, aVK);
	// Fix for v1.0.40: IsKeyToggledOn()'s call to GetKeyState() relies on our thread having
	// processed messages.  Confirmed necessary 100% of the time if our thread owns the active window.
	// v1.0.43: Extended the above fix to include all toggleable keys (not just Capslock) and to apply
	// to both directions (ON and OFF) since it seems likely to be needed for them all.
	bool our_thread_is_foreground;
	if (our_thread_is_foreground = (GetWindowThreadProcessId(GetForegroundWindow(), NULL) == g_MainThreadID)) // GetWindowThreadProcessId() tolerates a NULL hwnd.
		SLEEP_WITHOUT_INTERRUPTION(-1);
	if (aVK == VK_CAPITAL && aToggleValue == TOGGLED_OFF && IsKeyToggledOn(aVK))
	{
		// Fix for v1.0.36.06: Since it's Capslock and it didn't turn off as attempted, it's probably because
		// the OS is configured to turn Capslock off only in response to pressing the SHIFT key (via Ctrl Panel's
		// Regional settings).  So send shift to do it instead:
	 	KeyEvent(KEYDOWNANDUP, VK_SHIFT);
		if (our_thread_is_foreground) // v1.0.43: Added to try to achieve 100% reliability in all situations.
			SLEEP_WITHOUT_INTERRUPTION(-1); // Check msg queue to put SHIFT's turning off of Capslock into effect from our thread's POV.
	}
	return starting_state;
}


#ifdef CONFIG_WIN9X

void ToggleNumlockWin9x()
// Numlock requires a special method to toggle the state and its indicator light under Win9x.
// Capslock and Scrolllock do not need this method, since keybd_event() works for them.
{
	BYTE state[256];
	GetKeyboardState((PBYTE)&state);
	state[VK_NUMLOCK] ^= 0x01;  // Toggle the low-order bit to the opposite state.
	SetKeyboardState((PBYTE)&state);
}

//void CapslockOffWin9x()
//{
//	BYTE state[256];
//	GetKeyboardState((PBYTE)&state);
//	state[VK_CAPITAL] &= ~0x01;
//	SetKeyboardState((PBYTE)&state);
//}

#endif


/*
void SetKeyState (vk_type vk, int aKeyUp)
// Later need to adapt this to support Win9x by using SetKeyboardState for those OSs.
{
	if (!vk) return;
	int key_already_up = !(GetKeyState(vk) & 0x8000);
	if ((key_already_up && aKeyUp) || (!key_already_up && !aKeyUp))
		return;
	KeyEvent(aKeyUp, vk);
}
*/



void SetModifierLRState(modLR_type aModifiersLRnew, modLR_type aModifiersLRnow, HWND aTargetWindow
	, bool aDisguiseDownWinAlt, bool aDisguiseUpWinAlt, DWORD aExtraInfo)
// This function is designed to be called from only the main thread; it's probably not thread-safe.
// Puts modifiers into the specified state, releasing or pressing down keys as needed.
// The modifiers are released and pressed down in a very delicate order due to their interactions with
// each other and their ability to show the Start Menu, activate the menu bar, or trigger the OS's language
// bar hotkeys.  Side-effects like these would occur if a more simple approach were used, such as releasing
// all modifiers that are going up prior to pushing down the ones that are going down.
// When the target layout has an altgr key, it is tempting to try to simplify things by removing MOD_LCONTROL
// from aModifiersLRnew whenever aModifiersLRnew contains MOD_RALT.  However, this a careful review how that
// would impact various places below where sTargetLayoutHasAltGr is checked indicates that it wouldn't help.
// Note that by design and as documented for ControlSend, aTargetWindow is not used as the target for the
// various calls to KeyEvent() here.  It is only used as a workaround for the GUI window issue described
// at the bottom.
{
	if (aModifiersLRnow == aModifiersLRnew) // They're already in the right state, so avoid doing all the checks.
		return; // Especially avoids the aTargetWindow check at the bottom.

	// Notes about modifier key behavior on Windows XP (these probably apply to NT/2k also, and has also
	// been tested to be true on Win98): The WIN and ALT keys are the problem keys, because if either is
	// released without having modified something (even another modifier), the WIN key will cause the
	// Start Menu to appear, and the ALT key will activate the menu bar of the active window (if it has one).
	// For example, a hook hotkey such as "$#c::Send text" (text must start with a lowercase letter
	// to reproduce the issue, because otherwise WIN would be auto-disguised as a side effect of the SHIFT
	// keystroke) would cause the Start Menu to appear if the disguise method below weren't used.
	//
	// Here are more comments formerly in SetModifierLRStateSpecific(), which has since been eliminated
	// because this function is sufficient:
	// To prevent it from activating the menu bar, the release of the ALT key should be disguised
	// unless a CTRL key is currently down.  This is because CTRL always seems to avoid the
	// activation of the menu bar (unlike SHIFT, which sometimes allows the menu to be activated,
	// though this is hard to reproduce on XP).  Another reason not to use SHIFT is that the OS
	// uses LAlt+Shift as a hotkey to switch languages.  Such a hotkey would be triggered if SHIFT
	// were pressed down to disguise the release of LALT.
	//
	// Alt-down events are also disguised whenever they won't be accompanied by a Ctrl-down.
	// This is necessary whenever our caller does not plan to disguise the key itself.  For example,
	// if "!a::Send Test" is a registered hotkey, two things must be done to avoid complications:
	// 1) Prior to sending the word test, ALT must be released in a way that does not activate the
	//    menu bar.  This is done by sandwiching it between a CTRL-down and a CTRL-up.
	// 2) After the send is complete, SendKeys() will restore the ALT key to the down position if
	//    the user is still physically holding ALT down (this is done to make the logical state of
	//    the key match its physical state, which allows the same hotkey to be fired twice in a row
	//    without the user having to release and press down the ALT key physically).
	// The #2 case above is the one handled below by ctrl_wont_be_down.  It is especially necessary
	// when the user releases the ALT key prior to releasing the hotkey suffix, which would otherwise
	// cause the menu bar (if any) of the active window to be activated.
	//
	// Some of the same comments above for ALT key apply to the WIN key.  More about this issue:
	// Although the disguise of the down-event is usually not needed, it is needed in the rare case
	// where the user releases the WIN or ALT key prior to releasing the hotkey's suffix.
	// Although the hook could be told to disguise the physical release of ALT or WIN in these
	// cases, it's best not to rely on the hook since it is not always installed.
	//
	// Registered WIN and ALT hotkeys that don't use the Send command work okay except ALT hotkeys,
	// which if the user releases ALT prior the hotkey's suffix key, cause the menu bar to be activated.
	// Since it is unusual for users to do this and because it is standard behavior for  ALT hotkeys
	// registered in the OS, fixing it via the hook seems like a low priority, and perhaps isn't worth
	// the added code complexity/size.  But if there is ever a need to do so, the following note applies:
	// If the hook is installed, could tell it to disguise any need-to-be-disguised Alt-up that occurs
	// after receipt of the registered ALT hotkey.  But what if that hotkey uses the send command:
	// there might be interference?  Doesn't seem so, because the hook only disguises non-ignored events.

	// Set up some conditions so that the keystrokes that disguise the release of Win or Alt
	// are only sent when necessary (which helps avoid complications caused by keystroke interaction,
	// while improving performance):
	modLR_type aModifiersLRunion = aModifiersLRnow | aModifiersLRnew; // The set of keys that were or will be down.
	bool ctrl_not_down = !(aModifiersLRnow & (MOD_LCONTROL | MOD_RCONTROL)); // Neither CTRL key is down now.
	bool ctrl_will_not_be_down = !(aModifiersLRnew & (MOD_LCONTROL | MOD_RCONTROL)) // Nor will it be.
		&& !(sTargetLayoutHasAltGr == CONDITION_TRUE && (aModifiersLRnew & MOD_RALT)); // Nor will it be pushed down indirectly due to AltGr.

	bool ctrl_nor_shift_nor_alt_down = ctrl_not_down                             // Neither CTRL key is down now.
		&& !(aModifiersLRnow & (MOD_LSHIFT | MOD_RSHIFT | MOD_LALT | MOD_RALT)); // Nor is any SHIFT/ALT key.

	bool ctrl_or_shift_or_alt_will_be_down = !ctrl_will_not_be_down             // CTRL will be down.
		|| (aModifiersLRnew & (MOD_LSHIFT | MOD_RSHIFT | MOD_LALT | MOD_RALT)); // or SHIFT or ALT will be.

	// If the required disguise keys aren't down now but will be, defer the release of Win and/or Alt
	// until after the disguise keys are in place (since in that case, the caller wanted them down
	// as part of the normal operation here):
	bool defer_win_release = ctrl_nor_shift_nor_alt_down && ctrl_or_shift_or_alt_will_be_down;
	bool defer_alt_release = ctrl_not_down && !ctrl_will_not_be_down;  // i.e. Ctrl not down but it will be.
	bool release_shift_before_alt_ctrl = defer_alt_release // i.e. Control is moving into the down position or...
		|| !(aModifiersLRnow & (MOD_LALT | MOD_RALT)) && (aModifiersLRnew & (MOD_LALT | MOD_RALT)); // ...Alt is moving into the down position.
	// Concerning "release_shift_before_alt_ctrl" above: Its purpose is to prevent unwanted firing of the OS's
	// language bar hotkey.  See the bottom of this function for more explanation.

	// ALT:
	bool disguise_alt_down = aDisguiseDownWinAlt && ctrl_not_down && ctrl_will_not_be_down; // Since this applies to both Left and Right Alt, don't take sTargetLayoutHasAltGr into account here. That is done later below.

	// WIN: The WIN key is successfully disguised under a greater number of conditions than ALT.
	// Since SendPlay can't display Start Menu, there's no need to send the disguise-keystrokes (such
	// keystrokes might cause unwanted effects in certain games):
	bool disguise_win_down = aDisguiseDownWinAlt && sSendMode != SM_PLAY
		&& ctrl_not_down && ctrl_will_not_be_down
		&& !(aModifiersLRunion & (MOD_LSHIFT | MOD_RSHIFT)) // And neither SHIFT key is down, nor will it be.
		&& !(aModifiersLRunion & (MOD_LALT | MOD_RALT));    // And neither ALT key is down, nor will it be.

	bool release_lwin = (aModifiersLRnow & MOD_LWIN) && !(aModifiersLRnew & MOD_LWIN);
	bool release_rwin = (aModifiersLRnow & MOD_RWIN) && !(aModifiersLRnew & MOD_RWIN);
	bool release_lalt = (aModifiersLRnow & MOD_LALT) && !(aModifiersLRnew & MOD_LALT);
	bool release_ralt = (aModifiersLRnow & MOD_RALT) && !(aModifiersLRnew & MOD_RALT);
	bool release_lshift = (aModifiersLRnow & MOD_LSHIFT) && !(aModifiersLRnew & MOD_LSHIFT);
	bool release_rshift = (aModifiersLRnow & MOD_RSHIFT) && !(aModifiersLRnew & MOD_RSHIFT);

	// Handle ALT and WIN prior to the other modifiers because the "disguise" methods below are
	// only needed upon release of ALT or WIN.  This is because such releases tend to have a better
	// chance of being "disguised" if SHIFT or CTRL is down at the time of the release.  Thus, the
	// release of SHIFT or CTRL (if called for) is deferred until afterward.

	// ** WIN
	// Must be done before ALT in case it is relying on ALT being down to disguise the release WIN.
	// If ALT is going to be pushed down further below, defer_win_release should be true, which will make sure
	// the WIN key isn't released until after the ALT key is pushed down here at the top.
	// Also, WIN is a little more troublesome than ALT, so it is done first in case the ALT key
	// is down but will be going up, since the ALT key being down might help the WIN key.
	// For example, if you hold down CTRL, then hold down LWIN long enough for it to auto-repeat,
	// then release CTRL before releasing LWIN, the Start Menu would appear, at least on XP.
	// But it does not appear if CTRL is released after LWIN.
	// Also note that the ALT key can disguise the WIN key, but not vice versa.
	if (release_lwin)
	{
		if (!defer_win_release)
		{
			// Fixed for v1.0.25: To avoid triggering the system's LAlt+Shift language hotkey, the
			// Control key is now used to suppress LWIN/RWIN (preventing the Start Menu from appearing)
			// rather than the Shift key.  This is definitely needed for ALT, but is done here for
			// WIN also in case ALT is down, which might cause the use of SHIFT as the disguise key
			// to trigger the language switch.
			if (ctrl_nor_shift_nor_alt_down && aDisguiseUpWinAlt // Nor will they be pushed down later below, otherwise defer_win_release would have been true and we couldn't get to this point.
				&& sSendMode != SM_PLAY) // SendPlay can't display Start Menu, so disguise not needed (also, disguise might mess up some games).
				KeyEvent(KEYDOWNANDUP, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Disguise key release to suppress Start Menu.
				// The above event is safe because if we're here, it means VK_CONTROL will not be
				// pressed down further below.  In other words, we're not defeating the job
				// of this function by sending these disguise keystrokes.
			KeyEvent(KEYUP, VK_LWIN, 0, NULL, false, aExtraInfo);
		}
		// else release it only after the normal operation of the function pushes down the disguise keys.
	}
	else if (!(aModifiersLRnow & MOD_LWIN) && (aModifiersLRnew & MOD_LWIN)) // Press down LWin.
	{
		if (disguise_win_down)
			KeyEvent(KEYDOWN, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Ensures that the Start Menu does not appear.
		KeyEvent(KEYDOWN, VK_LWIN, 0, NULL, false, aExtraInfo);
		if (disguise_win_down)
			KeyEvent(KEYUP, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Ensures that the Start Menu does not appear.
	}

	if (release_rwin)
	{
		if (!defer_win_release)
		{
			if (ctrl_nor_shift_nor_alt_down && MOD_RWIN && sSendMode != SM_PLAY)
				KeyEvent(KEYDOWNANDUP, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Disguise key release to suppress Start Menu.
			KeyEvent(KEYUP, VK_RWIN, 0, NULL, false, aExtraInfo);
		}
		// else release it only after the normal operation of the function pushes down the disguise keys.
	}
	else if (!(aModifiersLRnow & MOD_RWIN) && (aModifiersLRnew & MOD_RWIN)) // Press down RWin.
	{
		if (disguise_win_down)
			KeyEvent(KEYDOWN, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Ensures that the Start Menu does not appear.
		KeyEvent(KEYDOWN, VK_RWIN, 0, NULL, false, aExtraInfo);
		if (disguise_win_down)
			KeyEvent(KEYUP, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Ensures that the Start Menu does not appear.
	}

	// ** SHIFT (PART 1 OF 2)
	if (release_shift_before_alt_ctrl)
	{
		if (release_lshift)
			KeyEvent(KEYUP, VK_LSHIFT, 0, NULL, false, aExtraInfo);
		if (release_rshift)
			KeyEvent(KEYUP, VK_RSHIFT, 0, NULL, false, aExtraInfo);
	}

	// ** ALT
	if (release_lalt)
	{
		if (!defer_alt_release)
		{
			if (ctrl_not_down && aDisguiseUpWinAlt)
				KeyEvent(KEYDOWNANDUP, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Disguise key release to suppress menu activation.
			KeyEvent(KEYUP, VK_LMENU, 0, NULL, false, aExtraInfo);
		}
	}
	else if (!(aModifiersLRnow & MOD_LALT) && (aModifiersLRnew & MOD_LALT))
	{
		if (disguise_alt_down)
			KeyEvent(KEYDOWN, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Ensures that menu bar is not activated.
		KeyEvent(KEYDOWN, VK_LMENU, 0, NULL, false, aExtraInfo);
		if (disguise_alt_down)
			KeyEvent(KEYUP, g_MenuMaskKey, 0, NULL, false, aExtraInfo);
	}

	if (release_ralt)
	{
		if (!defer_alt_release || sTargetLayoutHasAltGr == CONDITION_TRUE) // No need to defer if RAlt==AltGr. But don't change the value of defer_alt_release because LAlt uses it too.
		{
			if (sTargetLayoutHasAltGr == CONDITION_TRUE)
			{
				// Indicate that control is up now, since the release of AltGr will cause that indirectly.
				// Fix for v1.0.43: Unlike the pressing down of AltGr in a later section, which callers want
				// to automatically press down LControl too (by the very nature of AltGr), callers do not want
				// the release of AltGr to release LControl unless they specifically asked for LControl to be
				// released too.  This is because the caller may need LControl down to manifest something
				// like ^c. So don't do: aModifiersLRnew &= ~MOD_LCONTROL.
				// Without this fix, a hotkey like <^>!m::Send ^c would send "c" vs. "^c" on the German layout.
				// See similar section below for more details.
				aModifiersLRnow &= ~MOD_LCONTROL; // To reflect what KeyEvent(KEYUP, VK_RMENU) below will do.
			}
			else // No AltGr, so check if disguise is necessary (AltGr itself never needs disguise).
				if (ctrl_not_down && aDisguiseUpWinAlt)
					KeyEvent(KEYDOWNANDUP, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Disguise key release to suppress menu activation.
			KeyEvent(KEYUP, VK_RMENU, 0, NULL, false, aExtraInfo);
		}
	}
	else if (!(aModifiersLRnow & MOD_RALT) && (aModifiersLRnew & MOD_RALT)) // Press down RALT.
	{
		// For the below: There should never be a need to disguise AltGr.  Doing so would likely cause unwanted
		// side-effects. Also, disguise_alt_key does not take sTargetLayoutHasAltGr into account because
		// disguise_alt_key also applies to the left alt key.
		if (disguise_alt_down && sTargetLayoutHasAltGr != CONDITION_TRUE)
		{
			KeyEvent(KEYDOWN, g_MenuMaskKey, 0, NULL, false, aExtraInfo); // Ensures that menu bar is not activated.
			KeyEvent(KEYDOWN, VK_RMENU, 0, NULL, false, aExtraInfo);
			KeyEvent(KEYUP, g_MenuMaskKey, 0, NULL, false, aExtraInfo);
		}
		else // No disguise needed.
		{
			// v1.0.43: The following check was added to complement the other .43 fix higher above.
			// It may also fix other things independently of that other fix.
			// The following two lines release LControl before pushing down AltGr because otherwise,
			// the next time RAlt is released (such as by the user), some quirk of the OS or driver
			// prevents it from automatically releasing LControl like it normally does (perhaps
			// because the OS is designed to leave LControl down if it was down before AltGr went down).
			// This would cause LControl to get stuck down for hotkeys in German layout such as:
			//   <^>!a::SendRaw, {
			//   <^>!m::Send ^c
			if (sTargetLayoutHasAltGr == CONDITION_TRUE && (aModifiersLRnow & MOD_LCONTROL))
				KeyEvent(KEYUP, VK_LCONTROL, 0, NULL, false, aExtraInfo);
			KeyEvent(KEYDOWN, VK_RMENU, 0, NULL, false, aExtraInfo);
			if (sTargetLayoutHasAltGr == CONDITION_TRUE) // Note that KeyEvent() might have just changed the value of sTargetLayoutHasAltGr.
			{
				// Indicate that control is both down and required down so that the section after this one won't
				// release it.  Without this fix, a hotkey that sends an AltGr char such as "^ä:: SendRaw, {"
				// would fail to work under German layout because left-alt would be released after right-alt
				// goes down.
				aModifiersLRnow |= MOD_LCONTROL; // To reflect what KeyEvent() did above.
				aModifiersLRnew |= MOD_LCONTROL; // All callers want LControl to be down if they wanted AltGr to be down.
			}
		}
	}

	// CONTROL and SHIFT are done only after the above because the above might rely on them
	// being down before for certain early operations.

	// ** CONTROL
	if (   (aModifiersLRnow & MOD_LCONTROL) && !(aModifiersLRnew & MOD_LCONTROL) // Release LControl.
		// v1.0.41.01: The following line was added to fix the fact that callers do not want LControl
		// released when the new modifier state includes AltGr.  This solves a hotkey such as the following and
		// probably several other circumstances:
		// <^>!a::send \  ; Backslash is solved by this fix; it's manifest via AltGr+Dash on German layout.
		&& !((aModifiersLRnew & MOD_RALT) && sTargetLayoutHasAltGr == CONDITION_TRUE)   )
		KeyEvent(KEYUP, VK_LCONTROL, 0, NULL, false, aExtraInfo);
	else if (!(aModifiersLRnow & MOD_LCONTROL) && (aModifiersLRnew & MOD_LCONTROL)) // Press down LControl.
		KeyEvent(KEYDOWN, VK_LCONTROL, 0, NULL, false, aExtraInfo);
	if ((aModifiersLRnow & MOD_RCONTROL) && !(aModifiersLRnew & MOD_RCONTROL)) // Release RControl
		KeyEvent(KEYUP, VK_RCONTROL, 0, NULL, false, aExtraInfo);
	else if (!(aModifiersLRnow & MOD_RCONTROL) && (aModifiersLRnew & MOD_RCONTROL)) // Press down RControl.
		KeyEvent(KEYDOWN, VK_RCONTROL, 0, NULL, false, aExtraInfo);
	
	// ** SHIFT (PART 2 OF 2)
	// Must follow CTRL and ALT because a release of SHIFT while ALT/CTRL is down-but-soon-to-be-up
	// would switch languages via the OS hotkey.  It's okay if defer_alt_release==true because in that case,
	// CTRL just went down above (by definition of defer_alt_release), which will prevent the language hotkey
	// from firing.
	if (release_lshift && !release_shift_before_alt_ctrl) // Release LShift.
		KeyEvent(KEYUP, VK_LSHIFT, 0, NULL, false, aExtraInfo);
	else if (!(aModifiersLRnow & MOD_LSHIFT) && (aModifiersLRnew & MOD_LSHIFT)) // Press down LShift.
		KeyEvent(KEYDOWN, VK_LSHIFT, 0, NULL, false, aExtraInfo);
	if (release_rshift && !release_shift_before_alt_ctrl) // Release RShift.
		KeyEvent(KEYUP, VK_RSHIFT, 0, NULL, false, aExtraInfo);
	else if (!(aModifiersLRnow & MOD_RSHIFT) && (aModifiersLRnew & MOD_RSHIFT)) // Press down RShift.
		KeyEvent(KEYDOWN, VK_RSHIFT, 0, NULL, false, aExtraInfo);

	// ** KEYS DEFERRED FROM EARLIER
	if (defer_win_release) // Must be done before ALT because it might rely on ALT being down to disguise release of WIN key.
	{
		if (release_lwin)
			KeyEvent(KEYUP, VK_LWIN, 0, NULL, false, aExtraInfo);
		if (release_rwin)
			KeyEvent(KEYUP, VK_RWIN, 0, NULL, false, aExtraInfo);
	}
	if (defer_alt_release)
	{
		if (release_lalt)
			KeyEvent(KEYUP, VK_LMENU, 0, NULL, false, aExtraInfo);
		if (release_ralt && sTargetLayoutHasAltGr != CONDITION_TRUE) // If AltGr is present, RAlt would already have been released earlier since defer_alt_release would have been ignored for it.
			KeyEvent(KEYUP, VK_RMENU, 0, NULL, false, aExtraInfo);
	}

	// When calling KeyEvent(), probably best not to specify a scan code unless
	// absolutely necessary, since some keyboards may have non-standard scan codes
	// which KeyEvent() will resolve into the proper vk translations for us.
	// Decided not to Sleep() between keystrokes, even zero, out of concern that this
	// would result in a significant delay (perhaps more than 10ms) while the system
	// is under load.

	// Since the above didn't return early, keybd_event() has been used to change the state
	// of at least one modifier.  As a result, if the caller gave a non-NULL aTargetWindow,
	// it wants us to check if that window belongs to our thread.  If it does, we should do
	// a short msg queue check to prevent an apparent synchronization problem when using
	// ControlSend against the script's own GUI or other windows.  Here is an example of a
	// line whose modifier would not be in effect in time for its keystroke to be modified
	// by it:
	// ControlSend, Edit1, ^{end}, Test Window
	// Update: Another bug-fix for v1.0.21, as was the above: If the keyboard hook is installed,
	// the modifier keystrokes must have a way to get routed through the hook BEFORE the
	// keystrokes get sent via PostMessage().  If not, the correct modifier state will usually
	// not be in effect (or at least not be in sync) for the keys sent via PostMessage() afterward.
	// Notes about the macro below:
	// aTargetWindow!=NULL means ControlSend mode is in effect.
	// The g_KeybdHook check must come first (it should take precedence if both conditions are true).
	// -1 has been verified to be insufficient, at least for the very first letter sent if it is
	// supposed to be capitalized.
	// g_MainThreadID is the only thread of our process that owns any windows.

	int press_duration = (sSendMode == SM_PLAY) ? g->PressDurationPlay : g->PressDuration;
	if (press_duration > -1) // SM_PLAY does use DoKeyDelay() to store a delay item in the event array.
		// Since modifiers were changed by the above, do a key-delay if the special intra-keystroke
		// delay is in effect.
		// Since there normally isn't a delay between a change in modifiers and the first keystroke,
		// if a PressDuration is in effect, also do it here to improve reliability (I have observed
		// cases where modifiers need to be left alone for a short time in order for the keystrokes
		// that follow to be be modified by the intended set of modifiers).
		DoKeyDelay(press_duration); // It knows not to do the delay for SM_INPUT.
	else // Since no key-delay was done, check if a a delay is needed for any other reason.
	{
		// IMPORTANT UPDATE for v1.0.39: Now that the hooks are in a separate thread from the part
		// of the program that sends keystrokes for the script, you might think synchronization of
		// keystrokes would become problematic or at least change.  However, this is apparently not
		// the case.  MSDN doesn't spell this out, but testing shows that what happens with a low-level
		// hook is that the moment a keystroke comes into a thread (either physical or simulated), the OS
		// immediately calls something similar to SendMessage() from that thread to notify the hook
		// thread that a keystroke has arrived.  However, if the hook thread's priority is lower than
		// some other thread next in line for a timeslice, it might take some time for the hook thread
		// to get a timeslice (that's why the hook thread is given a high priority).
		// The SendMessage() call doesn't return until its timeout expires (as set in the registry for
		// hooks) or the hook thread processes the keystroke (which requires that it call something like
		// GetMessage/PeekMessage followed by a HookProc "return").  This is good news because it serializes
		// keyboard and mouse input to make the presence of the hook transparent to other threads (unless
		// the hook does something to reveal itself, such as suppressing keystrokes). Serialization avoids
		// any chance of synchronization problems such as a program that changes the state of a key then
		// immediately checks the state of that same key via GetAsyncKeyState().  Another way to look at
		// all of this is that in essence, a single-threaded hook program that simulates keystrokes or
		// mouse clicks should behave the same when the hook is moved into a separate thread because from
		// the program's point-of-view, keystrokes & mouse clicks result in a calling the hook almost
		// exactly as if the hook were in the same thread.
		if (aTargetWindow)
		{
			if (g_KeybdHook)
				SLEEP_WITHOUT_INTERRUPTION(0) // Don't use ternary operator to combine this with next due to "else if".
			else if (GetWindowThreadProcessId(aTargetWindow, NULL) == g_MainThreadID)
				SLEEP_WITHOUT_INTERRUPTION(-1)
		}
	}

	// Commented out because a return value is no longer needed by callers (since we do the key-delay here,
	// if appropriate).
	//return aModifiersLRnow ^ aModifiersLRnew; // Calculate the set of modifiers that changed (currently excludes AltGr's change of LControl's state).


	// NOTES about "release_shift_before_alt_ctrl":
	// If going down on alt or control (but not both, though it might not matter), and shift is to be released:
	//	Release shift first.
	// If going down on shift, and control or alt (but not both) is to be released:
	//	Release ctrl/alt first (this is already the case so nothing needs to be done).
	//
	// Release both shifts before going down on lalt/ralt or lctrl/rctrl (but not necessary if going down on
	// *both* alt+ctrl?
	// Release alt and both controls before going down on lshift/rshift.
	// Rather than the below, do the above (for the reason below).
	// But if do this, don't want to prevent a legit/intentional language switch such as:
	//    Send {LAlt down}{Shift}{LAlt up}.
	// If both Alt and Shift are down, Win or Ctrl (or any other key for that matter) must be pressed before either
	// is released.
	// If both Ctrl and Shift are down, Win or Alt (or any other key) must be pressed before either is released.
	// remind: Despite what the Regional Settings window says, RAlt+Shift (and Shift+RAlt) is also a language hotkey (i.e. not just LAlt), at least when RAlt isn't AltGr!
	// remind: Control being down suppresses language switch only once.  After that, control being down doesn't help if lalt is re-pressed prior to re-pressing shift.
	//
	// Language switch occurs when:
	// alt+shift (upon release of shift)
	// shift+alt (upon release of lalt)
	// ctrl+shift (upon release of shift)
	// shift+ctrl (upon release of ctrl)
	// Because language hotkey only takes effect upon release of Shift, it can be disguised via a Control keystroke if that is ever needed.

	// NOTES: More details about disguising ALT and WIN:
	// Registered Alt hotkeys don't quite work if the Alt key is released prior to the suffix.
	// Key history for Alt-B hotkey released this way, which undesirably activates the menu bar:
	// A4  038	 	d	0.03	Alt            	
	// 42  030	 	d	0.03	B              	
	// A4  038	 	u	0.24	Alt            	
	// 42  030	 	u	0.19	B              	
	// Testing shows that the above does not happen for a normal (non-hotkey) alt keystroke such as Alt-8,
	// so the above behavior is probably caused by the fact that B-down is suppressed by the OS's hotkey
	// routine, but not B-up.
	// The above also happens with registered WIN hotkeys, but only if the Send cmd resulted in the WIN
	// modifier being pushed back down afterward to match the fact that the user is still holding it down.
	// This behavior applies to ALT hotkeys also.
	// One solution: if the hook is installed, have it keep track of when the start menu or menu bar
	// *would* be activated.  These tracking vars can be consulted by the Send command, and the hook
	// can also be told when to use them after a registered hotkey has been pressed, so that the Alt-up
	// or Win-up keystroke that belongs to it can be disguised.

	// The following are important ways in which other methods of disguise might not be sufficient:
	// Sequence: shift-down win-down shift-up win-up: invokes Start Menu when WIN is held down long enough
	// to auto-repeat.  Same when Ctrl or Alt is used in lieu of Shift.
	// Sequence: shift-down alt-down alt-up shift-up: invokes menu bar.  However, as long as another key,
	// even Shift, is pressed down *after* alt is pressed down, menu bar is not activated, e.g. alt-down
	// shift-down shift-up alt-up.  In addition, CTRL always prevents ALT from activating the menu bar,
	// even with the following sequences:
	// ctrl-down alt-down alt-up ctrl-up
	// alt-down ctrl-down ctrl-up alt-up
	// (also seems true for all other permutations of Ctrl/Alt)
}



modLR_type GetModifierLRState(bool aExplicitlyGet)
// Try to report a more reliable state of the modifier keys than GetKeyboardState alone could.
// Fix for v1.0.42.01: On Windows 2000/XP or later, GetAsyncKeyState() should be called rather than
// GetKeyState().  This is because our callers always want the current state of the modifier keys
// rather than their state at the time of the currently-in-process message was posted.  For example,
// if the control key wasn't down at the time our thread's current message was posted, but it's logically
// down according to the system, we would want to release that control key before sending non-control
// keystrokes, even if one of our thread's own windows has keyboard focus (because if it does, the
// control-up keystroke should wind up getting processed after our thread realizes control is down).
// This applies even when the keyboard/mouse hook call use because keystrokes routed to the hook via
// the hook's message pump aren't messages per se, and thus GetKeyState and GetAsyncKeyState probably
// return the exact same thing whenever there are no messages in the hook's thread-queue (which is almost
// always the case).
{
	// If the hook is active, rely only on its tracked value rather than calling Get():
	if (g_KeybdHook && !aExplicitlyGet)
		return g_modifiersLR_logical;

	// Very old comment:
	// Use GetKeyState() rather than GetKeyboardState() because it's the only way to get
	// accurate key state when a console window is active, it seems.  I've also seen other
	// cases where GetKeyboardState() is incorrect (at least under WinXP) when GetKeyState(),
	// in its place, yields the correct info.  Very strange.

	modLR_type modifiersLR = 0;  // Allows all to default to up/off to simplify the below.
	if (g_os.IsWin9x() || g_os.IsWinNT4())
	{
		// Assume it's the left key since there's no way to tell which of the pair it
		// is? (unless the hook is installed, in which case it's value would have already
		// been returned, above).
		if (IsKeyDown9xNT(VK_SHIFT))   modifiersLR |= MOD_LSHIFT;
		if (IsKeyDown9xNT(VK_CONTROL)) modifiersLR |= MOD_LCONTROL;
		if (IsKeyDown9xNT(VK_MENU))    modifiersLR |= MOD_LALT;
		if (IsKeyDown9xNT(VK_LWIN))    modifiersLR |= MOD_LWIN;
		if (IsKeyDown9xNT(VK_RWIN))    modifiersLR |= MOD_RWIN;
	}
	else
	{
		if (IsKeyDownAsync(VK_LSHIFT))   modifiersLR |= MOD_LSHIFT;
		if (IsKeyDownAsync(VK_RSHIFT))   modifiersLR |= MOD_RSHIFT;
		if (IsKeyDownAsync(VK_LCONTROL)) modifiersLR |= MOD_LCONTROL;
		if (IsKeyDownAsync(VK_RCONTROL)) modifiersLR |= MOD_RCONTROL;
		if (IsKeyDownAsync(VK_LMENU))    modifiersLR |= MOD_LALT;
		if (IsKeyDownAsync(VK_RMENU))    modifiersLR |= MOD_RALT;
		if (IsKeyDownAsync(VK_LWIN))     modifiersLR |= MOD_LWIN;
		if (IsKeyDownAsync(VK_RWIN))     modifiersLR |= MOD_RWIN;
	}

	// Thread-safe: The following section isn't thread-safe because either the hook thread
	// or the main thread can be calling it.  However, given that anything dealing with
	// keystrokes isn't thread-safe in the sense that keystrokes can be coming in simultaneously
	// from multiple sources, it seems acceptable to keep it this way (especially since
	// the consequences of a thread collision seem very mild in this case).
	if (g_KeybdHook)
	{
		// Since hook is installed, fix any modifiers that it incorrectly thinks are down.
		// Though rare, this situation does arise during periods when the hook cannot track
		// the user's keystrokes, such as when the OS is doing something with the hardware,
		// e.g. switching to TV-out or changing video resolutions.  There are probably other
		// situations where this happens -- never fully explored and identified -- so it
		// seems best to do this, at least when the caller specified aExplicitlyGet = true.
		// To limit the scope of this workaround, only change the state of the hook modifiers
		// to be "up" for those keys the hook thinks are logically down but which the OS thinks
		// are logically up.  Note that it IS possible for a key to be physically down without
		// being logically down (i.e. during a Send command where the user is physically holding
		// down a modifier, but the Send command needs to put it up temporarily), so do not
		// change the hook's physical state for such keys in that case.
		// UPDATE: The following adjustment is now also relied upon by the SendInput method
		// to correct physical modifier state during periods when the hook was temporarily removed
		// to allow a SendInput to be uninterruptible.
		modLR_type modifiers_wrongly_down = g_modifiersLR_logical & ~modifiersLR;
		if (modifiers_wrongly_down)
		{
			// Adjust the physical and logical hook state to release the keys that are wrongly down.
			// If a key is wrongly logically down, it seems best to release it both physically and
			// logically, since the hook's failure to see the up-event probably makes its physical
			// state wrong in most such cases.
			g_modifiersLR_physical &= ~modifiers_wrongly_down;
			g_modifiersLR_logical &= ~modifiers_wrongly_down;
			g_modifiersLR_logical_non_ignored &= ~modifiers_wrongly_down;
			// Also adjust physical state so that the GetKeyState command will retrieve the correct values:
			AdjustKeyState(g_PhysicalKeyState, g_modifiersLR_physical);
		}
	}

	return modifiersLR;

	// Only consider a modifier key to be really down if both the hook's tracking of it
	// and GetKeyboardState() agree that it should be down.  The should minimize the impact
	// of the inherent unreliability present in each method (and each method is unreliable in
	// ways different from the other).  I have verified through testing that this eliminates
	// many misfires of hotkeys.  UPDATE: Both methods are fairly reliable now due to starting
	// to send scan codes with keybd_event(), using MapVirtualKey to resolve zero-value scan
	// codes in the keyboardproc(), and using GetKeyState() rather than GetKeyboardState().
	// There are still a few cases when they don't agree, so return the bitwise-and of both
	// if the keyboard hook is active.  Bitwise and is used because generally it's safer
	// to assume a modifier key is up, when in doubt (e.g. to avoid firing unwanted hotkeys):
//	return g_KeybdHook ? (g_modifiersLR_logical & g_modifiersLR_get) : g_modifiersLR_get;
}



void AdjustKeyState(BYTE aKeyState[], modLR_type aModifiersLR)
// Caller has ensured that aKeyState is a 256-BYTE array of key states, in the same format used
// by GetKeyboardState() and ToAsciiEx().
{
	aKeyState[VK_LSHIFT] = (aModifiersLR & MOD_LSHIFT) ? STATE_DOWN : 0;
	aKeyState[VK_RSHIFT] = (aModifiersLR & MOD_RSHIFT) ? STATE_DOWN : 0;
	aKeyState[VK_LCONTROL] = (aModifiersLR & MOD_LCONTROL) ? STATE_DOWN : 0;
	aKeyState[VK_RCONTROL] = (aModifiersLR & MOD_RCONTROL) ? STATE_DOWN : 0;
	aKeyState[VK_LMENU] = (aModifiersLR & MOD_LALT) ? STATE_DOWN : 0;
	aKeyState[VK_RMENU] = (aModifiersLR & MOD_RALT) ? STATE_DOWN : 0;
	aKeyState[VK_LWIN] = (aModifiersLR & MOD_LWIN) ? STATE_DOWN : 0;
	aKeyState[VK_RWIN] = (aModifiersLR & MOD_RWIN) ? STATE_DOWN : 0;
	// Update the state of neutral keys only after the above, in case both keys of the pair were wrongly down:
	aKeyState[VK_SHIFT] = (aKeyState[VK_LSHIFT] || aKeyState[VK_RSHIFT]) ? STATE_DOWN : 0;
	aKeyState[VK_CONTROL] = (aKeyState[VK_LCONTROL] || aKeyState[VK_RCONTROL]) ? STATE_DOWN : 0;
	aKeyState[VK_MENU] = (aKeyState[VK_LMENU] || aKeyState[VK_RMENU]) ? STATE_DOWN : 0;
}



modLR_type KeyToModifiersLR(vk_type aVK, sc_type aSC, bool *pIsNeutral)
// Convert the given virtual key / scan code to its equivalent bitwise modLR value.
// Callers rely upon the fact that we convert a neutral key such as VK_SHIFT into MOD_LSHIFT,
// not the bitwise combo of MOD_LSHIFT|MOD_RSHIFT.
// v1.0.43: VK_SHIFT should yield MOD_RSHIFT if the caller explicitly passed the right vs. left scan code.
// The SendPlay method relies on this to properly release AltGr, such as after "SendPlay @" in German.
// Other things may also rely on it because it is more correct.
{
	bool placeholder;
	bool &is_neutral = pIsNeutral ? *pIsNeutral : placeholder; // Simplifies other things below.
	is_neutral = false; // Set default for output parameter for caller.

	if (!(aVK || aSC))
		return 0;

	if (aVK) // Have vk take precedence over any non-zero sc.
		switch(aVK)
		{
		case VK_SHIFT:
			if (aSC == SC_RSHIFT)
				return MOD_RSHIFT;
			//else aSC is omitted (0) or SC_LSHIFT.  Either way, most callers would probably want that considered "neutral".
			is_neutral = true;
			return MOD_LSHIFT;
		case VK_LSHIFT: return MOD_LSHIFT;
		case VK_RSHIFT:	return MOD_RSHIFT;

		case VK_CONTROL:
			if (aSC == SC_RCONTROL)
				return MOD_RCONTROL;
			//else aSC is omitted (0) or SC_LCONTROL.  Either way, most callers would probably want that considered "neutral".
			is_neutral = true;
			return MOD_LCONTROL;
		case VK_LCONTROL: return MOD_LCONTROL;
		case VK_RCONTROL: return MOD_RCONTROL;

		case VK_MENU:
			if (aSC == SC_RALT)
				return MOD_RALT;
			//else aSC is omitted (0) or SC_LALT.  Either way, most callers would probably want that considered "neutral".
			is_neutral = true;
			return MOD_LALT;
		case VK_LMENU: return MOD_LALT;
		case VK_RMENU: return MOD_RALT;

		case VK_LWIN: return MOD_LWIN;
		case VK_RWIN: return MOD_RWIN;

		default:
			return 0;
		}

	// If above didn't return, rely on the scan code instead, which is now known to be non-zero.
	switch(aSC)
	{
	case SC_LSHIFT: return MOD_LSHIFT;
	case SC_RSHIFT:	return MOD_RSHIFT;
	case SC_LCONTROL: return MOD_LCONTROL;
	case SC_RCONTROL: return MOD_RCONTROL;
	case SC_LALT: return MOD_LALT;
	case SC_RALT: return MOD_RALT;
	case SC_LWIN: return MOD_LWIN;
	case SC_RWIN: return MOD_RWIN;
	}
	return 0;
}



modLR_type ConvertModifiers(mod_type aModifiers)
// Convert the input param to a modifiersLR value and return it.
{
	modLR_type modifiersLR = 0;
	if (aModifiers & MOD_WIN) modifiersLR |= (MOD_LWIN | MOD_RWIN);
	if (aModifiers & MOD_ALT) modifiersLR |= (MOD_LALT | MOD_RALT);
	if (aModifiers & MOD_CONTROL) modifiersLR |= (MOD_LCONTROL | MOD_RCONTROL);
	if (aModifiers & MOD_SHIFT) modifiersLR |= (MOD_LSHIFT | MOD_RSHIFT);
	return modifiersLR;
}



mod_type ConvertModifiersLR(modLR_type aModifiersLR)
// Convert the input param to a normal modifiers value and return it.
{
	mod_type modifiers = 0;
	if (aModifiersLR & (MOD_LWIN | MOD_RWIN)) modifiers |= MOD_WIN;
	if (aModifiersLR & (MOD_LALT | MOD_RALT)) modifiers |= MOD_ALT;
	if (aModifiersLR & (MOD_LSHIFT | MOD_RSHIFT)) modifiers |= MOD_SHIFT;
	if (aModifiersLR & (MOD_LCONTROL | MOD_RCONTROL)) modifiers |= MOD_CONTROL;
	return modifiers;
}



LPTSTR ModifiersLRToText(modLR_type aModifiersLR, LPTSTR aBuf)
// Caller has ensured that aBuf is not NULL.
{
	*aBuf = '\0';
	if (aModifiersLR & MOD_LWIN) _tcscat(aBuf, _T("LWin "));
	if (aModifiersLR & MOD_RWIN) _tcscat(aBuf, _T("RWin "));
	if (aModifiersLR & MOD_LSHIFT) _tcscat(aBuf, _T("LShift "));
	if (aModifiersLR & MOD_RSHIFT) _tcscat(aBuf, _T("RShift "));
	if (aModifiersLR & MOD_LCONTROL) _tcscat(aBuf, _T("LCtrl "));
	if (aModifiersLR & MOD_RCONTROL) _tcscat(aBuf, _T("RCtrl "));
	if (aModifiersLR & MOD_LALT) _tcscat(aBuf, _T("LAlt "));
	if (aModifiersLR & MOD_RALT) _tcscat(aBuf, _T("RAlt "));
	return aBuf;
}



bool ActiveWindowLayoutHasAltGr()
// Thread-safety: See comments in LayoutHasAltGr() below.
{
	Get_active_window_keybd_layout // Defines the variable active_window_keybd_layout for use below.
	return LayoutHasAltGr(active_window_keybd_layout) == CONDITION_TRUE; // i.e caller wants both CONDITION_FALSE and LAYOUT_UNDETERMINED to be considered non-AltGr.
}



ResultType LayoutHasAltGr(HKL aLayout, ResultType aHasAltGr)
// Thread-safety: While not thoroughly thread-safe, due to the extreme simplicity of the cache array, even if
// a collision occurs it should be inconsequential.
// Caller must ensure that aLayout is a valid layout (special values like 0 aren't supported here).
// If aHasAltGr is not at its default of LAYOUT_UNDETERMINED, the specified layout's has_altgr property is
// updated to the new value, but only if it is currently undetermined (callers can rely on this).
{
	// Layouts are cached for performance (to avoid the discovery loop later below).
	int i;
	for (i = 0; i < MAX_CACHED_LAYOUTS && sCachedLayout[i].hkl; ++i)
		if (sCachedLayout[i].hkl == aLayout) // Match Found.
		{
			if (aHasAltGr != LAYOUT_UNDETERMINED && sCachedLayout[i].has_altgr == LAYOUT_UNDETERMINED) // Caller relies on this.
				sCachedLayout[i].has_altgr = aHasAltGr;
			return sCachedLayout[i].has_altgr;
		}

	// Since above didn't return, this layout isn't cached yet.  So create a new cache entry for it and
	// determine whether this layout has an AltGr key.  If i<MAX_CACHED_LAYOUTS (which it almost always will be),
	// there's room in the array for a new cache entry.  In the very unlikely event that there isn't room,
	// overwrite an arbitrary item in the array.  An LRU/MRU algorithm (timestamp) isn't used because running out
	// of slots seems too unlikely, and the consequences of running out are merely a slight degradation in performance.
	CachedLayoutType &cl = sCachedLayout[(i < MAX_CACHED_LAYOUTS) ? i : MAX_CACHED_LAYOUTS-1];
	if (aHasAltGr != LAYOUT_UNDETERMINED) // Caller determined it for us.  See top of function for explanation.
	{
		cl.hkl = aLayout;
		return cl.has_altgr = aHasAltGr;
	}

	// Otherwise, do AltGr detection on this newly cached layout so that we can return the AltGr state to caller.
	// This detection is probably not 100% reliable because there may be some layouts (especially custom ones)
	// that have an AltGr key yet none of its characters actually require AltGr to manifest.  A more reliable
	// way to detect AltGr would be to simulate an RALT keystroke (maybe only an up event, not a down) and have
	// a keyboard hook catch and block it.  If the layout has altgr, the hook would see a driver-generated LCtrl
	// keystroke immediately prior to RAlt.
	// Performance: This loop is quite fast. Doing this section 1000 times only takes about 160ms
	// on a 2gHz system (0.16ms per call).
	SHORT s;
	for (cl.has_altgr = LAYOUT_UNDETERMINED, i = 32; i <= UorA(WCHAR_MAX,UCHAR_MAX); ++i) // Include Spacebar up through final ANSI character (i.e. include 255 but not 256).
	{
		s = VkKeyScanEx((TCHAR)i, aLayout);
		// Check for presence of Ctrl+Alt but allow other modifiers like Shift to be present because
		// I believe there are some layouts that manifest characters via Shift+AltGr.
		if (s != -1 && (s & 0x600) == 0x600) // In this context, Ctrl+Alt means AltGr.
		{
			cl.has_altgr = CONDITION_TRUE;
			break;
		}
	}

	// If loop didn't break, leave cl.has_altgr as LAYOUT_UNDETERMINED because we can't be sure whether AltGr is
	// present (see other comments for details).
	cl.hkl = aLayout; // This is done here (immediately after has_altgr was set in the loop above) rather than earlier to minimize the consequences of not being fully thread-safe.
	return cl.has_altgr;
}



LPTSTR SCtoKeyName(sc_type aSC, LPTSTR aBuf, int aBufSize, bool aUseFallback)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Always produces a non-empty string.
{
	for (int i = 0; i < g_key_to_sc_count; ++i)
	{
		if (g_key_to_sc[i].sc == aSC)
		{
			tcslcpy(aBuf, g_key_to_sc[i].key_name, aBufSize);
			return aBuf;
		}
	}
	// Since above didn't return, no match was found.  Use the default format for an unknown scan code:
	if (aUseFallback)
		sntprintf(aBuf, aBufSize, _T("sc%03X"), aSC);
	else
		*aBuf = '\0';
	return aBuf;
}



LPTSTR VKtoKeyName(vk_type aVK, LPTSTR aBuf, int aBufSize, bool aUseFallback)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Caller may omit aSC and it will be derived if needed.
{
	for (int i = 0; i < g_key_to_vk_count; ++i)
	{
		if (g_key_to_vk[i].vk == aVK)
		{
			tcslcpy(aBuf, g_key_to_vk[i].key_name, aBufSize);
			return aBuf;
		}
	}
	// Since above didn't return, no match was found.  Try to map it to
	// a character or use the default format for an unknown key code:
	if (*aBuf = VKtoChar(aVK))
		aBuf[1] = '\0';
	else if (aUseFallback && aVK)
		sntprintf(aBuf, aBufSize, _T("vk%02X"), aVK);
	else
		*aBuf = '\0';
	return aBuf;
}


TCHAR VKtoChar(vk_type aVK, HKL aKeybdLayout)
// Given a VK code, returns the character that an unmodified keypress would produce
// on the given keyboard layout.  Defaults to the script's own layout if omitted.
// Using this rather than MapVirtualKey() fixes some inconsistency that used to
// exist between 'A'-'Z' and every other key.
{
	if (!aKeybdLayout)
		aKeybdLayout = GetKeyboardLayout(0);
	
	// MapVirtualKeyEx() always produces 'A'-'Z' for those keys regardless of keyboard layout,
	// but for any other keys it produces the correct results, so we'll use it:
	if (aVK > 'Z' || aVK < 'A')
		return (TCHAR)MapVirtualKeyEx(aVK, MAPVK_VK_TO_CHAR, aKeybdLayout);

	// For any other keys, 
	TCHAR ch[3], ch_not_used[2];
	BYTE key_state[256];
	ZeroMemory(key_state, sizeof(key_state));
	TCHAR dead_char = 0;
	int n;

	// If there's a pending dead-key char in aKeybdLayout's buffer, it would modify the result.
	// We don't want that to happen, so as a workaround we pass a key-code which doesn't combine
	// with any dead chars, and will therefore pull it out.  VK_DECIMAL is used because it is
	// almost always valid; see http://www.siao2.com/2007/10/27/5717859.aspx
	if (ToUnicodeOrAsciiEx(VK_DECIMAL, 0, key_state, ch, 0, aKeybdLayout) == 2)
	{
		// Save the char to be later re-injected.
		dead_char = ch[0];
	}
	// Retrieve the character that corresponds to aVK, if any.
	n = ToUnicodeOrAsciiEx(aVK, 0, key_state, ch, 0, aKeybdLayout);
	if (n < 0) // aVK is a dead key, and we've just placed it into aKeybdLayout's buffer.
	{
		// Flush it out in the same manner as before (see above).
		ToUnicodeOrAsciiEx(VK_DECIMAL, 0, key_state, ch_not_used, 0, aKeybdLayout);
	}
	if (dead_char)
	{
		// Re-inject the dead-key char so that user input is not interrupted.
		// To do this, we need to find the right VK and modifier key combination:
		modLR_type modLR;
		vk_type dead_vk = CharToVKAndModifiers(dead_char, &modLR, aKeybdLayout);
		if (dead_vk)
		{
			AdjustKeyState(key_state, modLR);
			ToUnicodeOrAsciiEx(dead_vk, 0, key_state, ch_not_used, 0, aKeybdLayout);
		}
		//else: can't do it.
	}
	// ch[0] is set even for n < 0, but might not be for n == 0.
	return n ? ch[0] : 0;
}



sc_type TextToSC(LPTSTR aText)
{
	if (!*aText) return 0;
	for (int i = 0; i < g_key_to_sc_count; ++i)
		if (!_tcsicmp(g_key_to_sc[i].key_name, aText))
			return g_key_to_sc[i].sc;
	// Do this only after the above, in case any valid key names ever start with SC:
	if (ctoupper(*aText) == 'S' && ctoupper(*(aText + 1)) == 'C')
		return (sc_type)_tcstol(aText + 2, NULL, 16);  // Convert from hex.
	return 0; // Indicate "not found".
}



vk_type TextToVK(LPTSTR aText, modLR_type *pModifiersLR, bool aExcludeThoseHandledByScanCode, bool aAllowExplicitVK
	, HKL aKeybdLayout)
// If modifiers_p is non-NULL, place the modifiers that are needed to realize the key in there.
// e.g. M is really +m (shift-m), # is really shift-3.
// HOWEVER, this function does not completely overwrite the contents of pModifiersLR; instead, it just
// adds the required modifiers into whatever is already there.
{
	if (!*aText) return 0;

	// Don't trim() aText or modify it because that will mess up the caller who expects it to be unchanged.
	// Instead, for now, just check it as-is.  The only extra whitespace that should exist, due to trimming
	// of text during load, is that on either side of the COMPOSITE_DELIMITER (e.g. " then ").

	if (!aText[1]) // _tcslen(aText) == 1
		return CharToVKAndModifiers(*aText, pModifiersLR, aKeybdLayout); // Making this a function simplifies things because it can do early return, etc.

	if (aAllowExplicitVK && ctoupper(aText[0]) == 'V' && ctoupper(aText[1]) == 'K')
		return (vk_type)_tcstol(aText + 2, NULL, 16);  // Convert from hex.

	for (int i = 0; i < g_key_to_vk_count; ++i)
		if (!_tcsicmp(g_key_to_vk[i].key_name, aText))
			return g_key_to_vk[i].vk;

	if (aExcludeThoseHandledByScanCode)
		return 0; // Zero is not a valid virtual key, so it should be a safe failure indicator.

	// Otherwise check if aText is the name of a key handled by scan code and if so, map that
	// scan code to its corresponding virtual key:
	sc_type sc = TextToSC(aText);
	return sc ? sc_to_vk(sc) : 0;
}



vk_type CharToVKAndModifiers(TCHAR aChar, modLR_type *pModifiersLR, HKL aKeybdLayout)
// If non-NULL, pModifiersLR contains the initial set of modifiers provided by the caller, to which
// we add any extra modifiers required to realize aChar.
{
	// For v1.0.25.12, it seems best to avoid the many recent problems with linefeed (`n) being sent
	// as Ctrl+Enter by changing it to always send a plain Enter, just like carriage return (`r).
	if (aChar == '\n')
		return VK_RETURN;

	// Otherwise:
	SHORT mod_plus_vk = VkKeyScanEx(aChar, aKeybdLayout); // v1.0.44.03: Benchmark shows that VkKeyScanEx() is the same speed as VkKeyScan() when the layout has been pre-fetched.
	vk_type vk = LOBYTE(mod_plus_vk);
	char keyscan_modifiers = HIBYTE(mod_plus_vk);
	if (keyscan_modifiers == -1 && vk == (UCHAR)-1) // No translation could be made.
		return 0;
	if (keyscan_modifiers & 0x38) // "The Hankaku key is pressed" or either of the "Reserved" state bits (for instance, used by Neo2 keyboard layout).
		// Callers expect failure in this case so that a fallback method can be used.
		return 0;

	// For v1.0.35, pModifiersLR was changed to modLR vs. mod so that AltGr keys such as backslash and
	// '{' are supported on layouts such as German when sending to apps such as Putty that are fussy about
	// which ALT key is held down to produce the character.  The following section detects AltGr by the
	// assuming that any character that requires both CTRL and ALT (with optional SHIFT) to be held
	// down is in fact an AltGr key (I don't think there are any that aren't AltGr in this case, but
	// confirmation would be nice).  Also, this is not done for Win9x because the distinction between
	// right and left-alt is not well-supported and it might do more harm than good (testing is
	// needed on fussy apps like Putty on Win9x).  UPDATE: Windows NT4 is now excluded from this
	// change because apparently it wants the left Alt key's virtual key and not the right's (though
	// perhaps it would prefer the right scan code vs. the left in apps such as Putty, but until that
	// is proven, the complexity is not added here).  Otherwise, on French and other layouts on NT4,
	// AltGr-produced characters such as backslash do not get sent properly.  In hindsight, this is
	// not surprising because the keyboard hook also receives neutral modifier keys on NT4 rather than
	// a more specific left/right key.

	// The win docs for VkKeyScan() are a bit confusing, referring to flag "bits" when it should really
	// say flag "values".  In addition, it seems that these flag values are incompatible with
	// MOD_ALT, MOD_SHIFT, and MOD_CONTROL, so they must be translated:
	if (pModifiersLR) // The caller wants this info added to the output param.
	{
		// Best not to reset this value because some callers want to retain what was in it before,
		// merely merging these new values into it:
		//*pModifiers = 0;
		if ((keyscan_modifiers & 0x06) == 0x06 && g_os.IsWin2000orLater()) // 0x06 means "requires/includes AltGr".
		{
			// v1.0.35: The critical difference below is right vs. left ALT.  Must not include MOD_LCONTROL
			// because simulating the RAlt keystroke on these keyboard layouts will automatically
			// press LControl down.
			*pModifiersLR |= MOD_RALT;
		}
		else // Do normal/default translation.
		{
			// v1.0.40: If caller-supplied modifiers already include the right-side key, no need to
			// add the left-side key (avoids unnecessary keystrokes).
			if (   (keyscan_modifiers & 0x02) && !(*pModifiersLR & (MOD_LCONTROL|MOD_RCONTROL))   )
				*pModifiersLR |= MOD_LCONTROL; // Must not be done if requires_altgr==true, see above.
			if (   (keyscan_modifiers & 0x04) && !(*pModifiersLR & (MOD_LALT|MOD_RALT))   )
				*pModifiersLR |= MOD_LALT;
		}
		// v1.0.36.06: Done unconditionally because presence of AltGr should not preclude the presence of Shift.
		// v1.0.40: If caller-supplied modifiers already contains MOD_RSHIFT, no need to add LSHIFT (avoids
		// unnecessary keystrokes).
		if (   (keyscan_modifiers & 0x01) && !(*pModifiersLR & (MOD_LSHIFT|MOD_RSHIFT))   )
			*pModifiersLR |= MOD_LSHIFT;
	}
	return vk;
}



vk_type TextToSpecial(LPTSTR aText, size_t aTextLength, KeyEventTypes &aEventType, modLR_type &aModifiersLR
	, bool aUpdatePersistent)
// Returns vk for key-down, negative vk for key-up, or zero if no translation.
// We also update whatever's in *pModifiers and *pModifiersLR to reflect the type of key-action
// specified in <aText>.  This makes it so that {altdown}{esc}{altup} behaves the same as !{esc}.
// Note that things like LShiftDown are not supported because: 1) they are rarely needed; and 2)
// they can be down via "lshift down".
{
	if (!tcslicmp(aText, _T("ALTDOWN"), aTextLength))
	{
		if (aUpdatePersistent)
			if (!(aModifiersLR & (MOD_LALT | MOD_RALT))) // i.e. do nothing if either left or right is already present.
				aModifiersLR |= MOD_LALT; // If neither is down, use the left one because it's more compatible.
		aEventType = KEYDOWN;
		return VK_MENU;
	}
	if (!tcslicmp(aText, _T("ALTUP"), aTextLength))
	{
		// Unlike for Lwin/Rwin, it seems best to have these neutral keys (e.g. ALT vs. LALT or RALT)
		// restore either or both of the ALT keys into the up position.  The user can use {LAlt Up}
		// to be more specific and avoid this behavior:
		if (aUpdatePersistent)
			aModifiersLR &= ~(MOD_LALT | MOD_RALT);
		aEventType = KEYUP;
		return VK_MENU;
	}
	if (!tcslicmp(aText, _T("SHIFTDOWN"), aTextLength))
	{
		if (aUpdatePersistent)
			if (!(aModifiersLR & (MOD_LSHIFT | MOD_RSHIFT))) // i.e. do nothing if either left or right is already present.
				aModifiersLR |= MOD_LSHIFT; // If neither is down, use the left one because it's more compatible.
		aEventType = KEYDOWN;
		return VK_SHIFT;
	}
	if (!tcslicmp(aText, _T("SHIFTUP"), aTextLength))
	{
		if (aUpdatePersistent)
			aModifiersLR &= ~(MOD_LSHIFT | MOD_RSHIFT); // See "ALTUP" for explanation.
		aEventType = KEYUP;
		return VK_SHIFT;
	}
	if (!tcslicmp(aText, _T("CTRLDOWN"), aTextLength) || !tcslicmp(aText, _T("CONTROLDOWN"), aTextLength))
	{
		if (aUpdatePersistent)
			if (!(aModifiersLR & (MOD_LCONTROL | MOD_RCONTROL))) // i.e. do nothing if either left or right is already present.
				aModifiersLR |= MOD_LCONTROL; // If neither is down, use the left one because it's more compatible.
		aEventType = KEYDOWN;
		return VK_CONTROL;
	}
	if (!tcslicmp(aText, _T("CTRLUP"), aTextLength) || !tcslicmp(aText, _T("CONTROLUP"), aTextLength))
	{
		if (aUpdatePersistent)
			aModifiersLR &= ~(MOD_LCONTROL | MOD_RCONTROL); // See "ALTUP" for explanation.
		aEventType = KEYUP;
		return VK_CONTROL;
	}
	if (!tcslicmp(aText, _T("LWINDOWN"), aTextLength))
	{
		if (aUpdatePersistent)
			aModifiersLR |= MOD_LWIN;
		aEventType = KEYDOWN;
		return VK_LWIN;
	}
	if (!tcslicmp(aText, _T("LWINUP"), aTextLength))
	{
		if (aUpdatePersistent)
			aModifiersLR &= ~MOD_LWIN;
		aEventType = KEYUP;
		return VK_LWIN;
	}
	if (!tcslicmp(aText, _T("RWINDOWN"), aTextLength))
	{
		if (aUpdatePersistent)
			aModifiersLR |= MOD_RWIN;
		aEventType = KEYDOWN;
		return VK_RWIN;
	}
	if (!tcslicmp(aText, _T("RWINUP"), aTextLength))
	{
		if (aUpdatePersistent)
			aModifiersLR &= ~MOD_RWIN;
		aEventType = KEYUP;
		return VK_RWIN;
	}
	// Otherwise, leave aEventType unchanged and return zero to indicate failure:
	return 0;
}



#ifdef ENABLE_KEY_HISTORY_FILE
ResultType KeyHistoryToFile(LPTSTR aFilespec, char aType, bool aKeyUp, vk_type aVK, sc_type aSC)
{
	static TCHAR sTargetFilespec[MAX_PATH] = _T("");
	static FILE *fp = NULL;
	static HWND last_foreground_window = NULL;
	static DWORD last_tickcount = GetTickCount();

	if (!g_KeyHistory) // Since key history is disabled, keys are not being tracked by the hook, so there's nothing to log.
		return OK;     // Files should not need to be closed since they would never have been opened in the first place.

	if (!aFilespec && !aVK && !aSC) // Caller is signaling to close the file if it's open.
	{
		if (fp)
		{
			fclose(fp);
			fp = NULL;
		}
		return OK;
	}

	if (aFilespec && *aFilespec && lstrcmpi(aFilespec, sTargetFilespec)) // Target filename has changed.
	{
		if (fp)
		{
			fclose(fp);
			fp = NULL;  // To indicate to future calls to this function that it's closed.
		}
		tcslcpy(sTargetFilespec, aFilespec, _countof(sTargetFilespec));
	}

	if (!aVK && !aSC) // Caller didn't want us to log anything this time.
		return OK;
	if (!*sTargetFilespec)
		return OK; // No target filename has ever been specified, so don't even attempt to open the file.

	if (!aVK)
		aVK = sc_to_vk(aSC);
	else
		if (!aSC)
			aSC = vk_to_sc(aVK);

	TCHAR buf[2048] = _T(""), win_title[1024] = _T("<Init>"), key_name[128] = _T("");
	HWND curr_foreground_window = GetForegroundWindow();
	DWORD curr_tickcount = GetTickCount();
	bool log_changed_window = (curr_foreground_window != last_foreground_window);
	if (log_changed_window)
	{
		if (curr_foreground_window)
			GetWindowText(curr_foreground_window, win_title, _countof(win_title));
		else
			tcslcpy(win_title, _T("<None>"), _countof(win_title));
		last_foreground_window = curr_foreground_window;
	}

	sntprintf(buf, _countof(buf), _T("%02X") _T("\t%03X") _T("\t%0.2f") _T("\t%c") _T("\t%c") _T("\t%s") _T("%s%s\n")
		, aVK, aSC
		, (float)(curr_tickcount - last_tickcount) / (float)1000
		, aType
		, aKeyUp ? 'u' : 'd'
		, GetKeyName(aVK, aSC, key_name, sizeof(key_name))
		, log_changed_window ? _T("\t") : _T("")
		, log_changed_window ? win_title : _T("")
		);
	last_tickcount = curr_tickcount;
	if (!fp)
		if (   !(fp = _tfopen(sTargetFilespec, _T("a")))   )
			return OK;
	_fputts(buf, fp);
	return OK;
}
#endif



LPTSTR GetKeyName(vk_type aVK, sc_type aSC, LPTSTR aBuf, int aBufSize, LPTSTR aDefault)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Caller has ensured that aBuf isn't NULL.
{
	if (aBufSize < 3)
		return aBuf;

	*aBuf = '\0'; // Set default.
	if (!aVK && !aSC)
		return aBuf;

	if (!aVK)
		aVK = sc_to_vk(aSC);
	else
		if (!aSC)
			aSC = vk_to_sc(aVK);

	// Check SC first to properly differentiate between Home/NumpadHome, End/NumpadEnd, etc.
	// v1.0.43: WheelDown/Up store the notch/turn count in SC, so don't consider that to be a valid SC.
	if (aSC && !IS_WHEEL_VK(aVK))
	{
		if (*SCtoKeyName(aSC, aBuf, aBufSize, false))
			return aBuf;
		// Otherwise this key is probably one we can handle by VK.
	}
	if (*VKtoKeyName(aVK, aBuf, aBufSize, false))
		return aBuf;
	// Since this key is unrecognized, return the caller-supplied default value.
	return aDefault;
}



sc_type vk_to_sc(vk_type aVK, bool aReturnSecondary)
// For v1.0.37.03, vk_to_sc() was converted into a function rather than being an array because if the
// script's keyboard layout changes while it's running, the array would get out-of-date.
// If caller passes true for aReturnSecondary, the non-primary scan code will be returned for
// virtual keys that two scan codes (if there's only one scan code, callers rely on zero being returned).
{
	// Try to minimize the number mappings done manually because MapVirtualKey is a more reliable
	// way to get the mapping if user has non-standard or custom keyboard layout.

	sc_type sc = 0;

	switch (aVK)
	{
	// Yield a manually translation for virtual keys that MapVirtualKey() doesn't support or for which it
	// doesn't yield consistent result (such as Win9x supporting only SHIFT rather than VK_LSHIFT/VK_RSHIFT).
	case VK_LSHIFT:   sc = SC_LSHIFT; break; // Modifiers are listed first for performance.
	case VK_RSHIFT:   sc = SC_RSHIFT; break;
	case VK_LCONTROL: sc = SC_LCONTROL; break;
	case VK_RCONTROL: sc = SC_RCONTROL; break;
	case VK_LMENU:    sc = SC_LALT; break;
	case VK_RMENU:    sc = SC_RALT; break;
	case VK_LWIN:     sc = SC_LWIN; break; // Earliest versions of Win95/NT might not support these, so map them manually.
	case VK_RWIN:     sc = SC_RWIN; break; //

	// According to http://support.microsoft.com/default.aspx?scid=kb;en-us;72583
	// most or all numeric keypad keys cannot be mapped reliably under any OS. The article is
	// a little unclear about which direction, if any, that MapVirtualKey() does work in for
	// the numpad keys, so for peace-of-mind map them all manually for now:
	case VK_NUMPAD0:  sc = SC_NUMPAD0; break;
	case VK_NUMPAD1:  sc = SC_NUMPAD1; break;
	case VK_NUMPAD2:  sc = SC_NUMPAD2; break;
	case VK_NUMPAD3:  sc = SC_NUMPAD3; break;
	case VK_NUMPAD4:  sc = SC_NUMPAD4; break;
	case VK_NUMPAD5:  sc = SC_NUMPAD5; break;
	case VK_NUMPAD6:  sc = SC_NUMPAD6; break;
	case VK_NUMPAD7:  sc = SC_NUMPAD7; break;
	case VK_NUMPAD8:  sc = SC_NUMPAD8; break;
	case VK_NUMPAD9:  sc = SC_NUMPAD9; break;
	case VK_DECIMAL:  sc = SC_NUMPADDOT; break;
	case VK_NUMLOCK:  sc = SC_NUMLOCK; break;
	case VK_DIVIDE:   sc = SC_NUMPADDIV; break;
	case VK_MULTIPLY: sc = SC_NUMPADMULT; break;
	case VK_SUBTRACT: sc = SC_NUMPADSUB; break;
	case VK_ADD:      sc = SC_NUMPADADD; break;
	}

	if (sc) // Above found a match.
		return aReturnSecondary ? 0 : sc; // Callers rely on zero being returned for VKs that don't have secondary SCs.

	// Use the OS API's MapVirtualKey() to resolve any not manually done above:
	if (   !(sc = MapVirtualKey(aVK, 0))   )
		return 0; // Indicate "no mapping".

	// Turn on the extended flag for those that need it.
	// Because MapVirtualKey can only accept (and return) naked scan codes (the low-order byte),
	// handle extended scan codes that have a non-extended counterpart manually.
	// Older comment: MapVirtualKey() should include 0xE0 in HIBYTE if key is extended, BUT IT DOESN'T.
	// There doesn't appear to be any built-in function to determine whether a vk's scan code
	// is extended or not.  See MSDN topic "keyboard input" for the below list.
	// Note: NumpadEnter is probably the only extended key that doesn't have a unique VK of its own.
	// So in that case, probably safest not to set the extended flag.  To send a true NumpadEnter,
	// as well as a true NumPadDown and any other key that shares the same VK with another, the
	// caller should specify the sc param to circumvent the need for KeyEvent() to use the below:
	switch (aVK)
	{
	case VK_APPS:     // Application key on keyboards with LWIN/RWIN/Apps.  Not listed in MSDN as "extended"?
	case VK_CANCEL:   // Ctrl-break
	case VK_SNAPSHOT: // PrintScreen
	case VK_DIVIDE:   // NumpadDivide (slash)
	case VK_NUMLOCK:
	// Below are extended but were already handled and returned from higher above:
	//case VK_LWIN:
	//case VK_RWIN:
	//case VK_RMENU:
	//case VK_RCONTROL:
	//case VK_RSHIFT: // WinXP needs this to be extended for keybd_event() to work properly.
		sc |= 0x0100;
		break;

	// The following virtual keys have more than one physical key, and thus more than one scan code.
	// If the caller passed true for aReturnSecondary, the extended version of the scan code will be
	// returned (all of the following VKs have two SCs):
	case VK_RETURN:
	case VK_INSERT:
	case VK_DELETE:
	case VK_PRIOR: // PgUp
	case VK_NEXT:  // PgDn
	case VK_HOME:
	case VK_END:
	case VK_UP:
	case VK_DOWN:
	case VK_LEFT:
	case VK_RIGHT:
		return aReturnSecondary ? (sc | 0x0100) : sc; // Below relies on the fact that these cases return early.
	}

	// Since above didn't return, if aReturnSecondary==true, return 0 to indicate "no secondary SC for this VK".
	return aReturnSecondary ? 0 : sc; // Callers rely on zero being returned for VKs that don't have secondary SCs.
}



vk_type sc_to_vk(sc_type aSC)
{
	// These are mapped manually because MapVirtualKey() doesn't support them correctly, at least
	// on some -- if not all -- OSs.  The main app also relies upon the values assigned below to
	// determine which keys should be handled by scan code rather than vk:
	switch (aSC)
	{
	// Even though neither of the SHIFT keys are extended -- and thus could be mapped with MapVirtualKey()
	// -- it seems better to define them explicitly because under Win9x (maybe just Win95).
	// I'm pretty sure MapVirtualKey() would return VK_SHIFT instead of the left/right VK.
	case SC_LSHIFT:      return VK_LSHIFT; // Modifiers are listed first for performance.
	case SC_RSHIFT:      return VK_RSHIFT;
	case SC_LCONTROL:    return VK_LCONTROL;
	case SC_RCONTROL:    return VK_RCONTROL;
	case SC_LALT:        return VK_LMENU;
	case SC_RALT:        return VK_RMENU;

	// Numpad keys require explicit mapping because MapVirtualKey() doesn't support them on all OSes.
	// See comments in vk_to_sc() for details.
	case SC_NUMLOCK:     return VK_NUMLOCK;
	case SC_NUMPADDIV:   return VK_DIVIDE;
	case SC_NUMPADMULT:  return VK_MULTIPLY;
	case SC_NUMPADSUB:   return VK_SUBTRACT;
	case SC_NUMPADADD:   return VK_ADD;
	case SC_NUMPADENTER: return VK_RETURN;

	// The following are ambiguous because each maps to more than one VK.  But be careful
	// changing the value to the other choice because some callers rely upon the values
	// assigned below to determine which keys should be handled by scan code rather than vk:
	case SC_NUMPADDEL:   return VK_DELETE;
	case SC_NUMPADCLEAR: return VK_CLEAR;
	case SC_NUMPADINS:   return VK_INSERT;
	case SC_NUMPADUP:    return VK_UP;
	case SC_NUMPADDOWN:  return VK_DOWN;
	case SC_NUMPADLEFT:  return VK_LEFT;
	case SC_NUMPADRIGHT: return VK_RIGHT;
	case SC_NUMPADHOME:  return VK_HOME;
	case SC_NUMPADEND:   return VK_END;
	case SC_NUMPADPGUP:  return VK_PRIOR;
	case SC_NUMPADPGDN:  return VK_NEXT;

	// No callers currently need the following alternate virtual key mappings.  If it is ever needed,
	// could have a new aReturnSecondary parameter that if true, causes these to be returned rather
	// than the above:
	//case SC_NUMPADDEL:   return VK_DECIMAL;
	//case SC_NUMPADCLEAR: return VK_NUMPAD5; // Same key as Numpad5 on most keyboards?
	//case SC_NUMPADINS:   return VK_NUMPAD0;
	//case SC_NUMPADUP:    return VK_NUMPAD8;
	//case SC_NUMPADDOWN:  return VK_NUMPAD2;
	//case SC_NUMPADLEFT:  return VK_NUMPAD4;
	//case SC_NUMPADRIGHT: return VK_NUMPAD6;
	//case SC_NUMPADHOME:  return VK_NUMPAD7;
	//case SC_NUMPADEND:   return VK_NUMPAD1;
	//case SC_NUMPADPGUP:  return VK_NUMPAD9;
	//case SC_NUMPADPGDN:  return VK_NUMPAD3;	

	case SC_APPSKEY:	return VK_APPS; // Added in v1.1.17.00.
	}

	// Use the OS API call to resolve any not manually set above.  This should correctly
	// resolve even elements such as SC_INSERT, which is an extended scan code, because
	// it passes in only the low-order byte which is SC_NUMPADINS.  In the case of SC_INSERT
	// and similar ones, MapVirtualKey() will return the same vk for both, which is correct.
	// Only pass the LOBYTE because I think it fails to work properly otherwise.
	// Also, DO NOT pass 3 for the 2nd param of MapVirtualKey() because apparently
	// that is not compatible with Win9x so it winds up returning zero for keys
	// such as UP, LEFT, HOME, and PGUP (maybe other sorts of keys too).  This
	// should be okay even on XP because the left/right specific keys have already
	// been resolved above so don't need to be looked up here (LWIN and RWIN
	// each have their own VK's so shouldn't be problem for the below call to resolve):
	return MapVirtualKey((BYTE)aSC, 1);
}
