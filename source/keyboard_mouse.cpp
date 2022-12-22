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
#include "abi.h"


// Added for v1.0.25.  Search on sPrevEventType for more comments:
static KeyEventTypes sPrevEventType;
static vk_type sPrevVK = 0;
// For v1.0.25, the below is static to track it in between sends, so that the below will continue
// to work:
// Send {LWinDown}
// Send {LWinUp}  ; Should still open the Start Menu even though it's a separate Send.
static vk_type sPrevEventModifierDown = 0;
static modLR_type sModifiersLR_persistent = 0; // Tracks this script's own lifetime/persistent modifiers (the ones it caused to be persistent and thus is responsible for tracking).
static modLR_type sModifiersLR_remapped = 0;

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
		KeyEventMenuMask(KEYDOWNANDUP); // Disguise it to suppress Start Menu or prevent activation of active window's menu bar.
}



// moved from SendKeys
void SendUnicodeChar(wchar_t aChar, modLR_type aModifiers)
{
	// Set modifier keystate as specified by caller.  Generally this will be 0, since
	// key combinations with Unicode packets either do nothing at all or do the same as
	// without the modifiers.  All modifiers are known to interfere in some applications.
	SetModifierLRState(aModifiers, sSendMode ? sEventModifiersLR : GetModifierLRState(), NULL, false, true, KEY_IGNORE);

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



void SendKeys(LPCTSTR aKeys, SendRawModes aSendRaw, SendModes aSendModeOrig, HWND aTargetWindow)
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

	DWORD orig_last_peek_time = g_script.mLastPeekTime;

	// For performance and also to reserve future flexibility, recognize {Blind} only when it's the first item
	// in the string.
	modLR_type mods_excluded_from_blind = 0;
	if (sInBlindMode = !aSendRaw && !_tcsnicmp(aKeys, _T("{Blind"), 6)) // Don't allow {Blind} while in raw mode due to slight chance {Blind} is intended to be sent as a literal string.
	{
		// Blind Mode (since this seems too obscure to document, it's mentioned here):  Blind Mode relies
		// on modifiers already down for something like ^c because ^c is saying "manifest a ^c", which will
		// happen if ctrl is already down.  By contrast, Blind does not release shift to produce lowercase
		// letters because avoiding that adds flexibility that couldn't be achieved otherwise.
		// Thus, ^c::Send {Blind}c produces the same result when ^c is substituted for the final c.
		// But Send {Blind}{LControl down} will generate the extra events even if ctrl already down.
		for (aKeys += 6; *aKeys != '}'; ++aKeys)
		{
			switch (*aKeys)
			{
			case '^': mods_excluded_from_blind |= MOD_LCONTROL|MOD_RCONTROL; break;
			case '+': mods_excluded_from_blind |= MOD_LSHIFT|MOD_RSHIFT; break;
			case '!': mods_excluded_from_blind |= MOD_LALT|MOD_RALT; break;
			case '#': mods_excluded_from_blind |= MOD_LWIN|MOD_RWIN; break;
			case '\0': return; // Just ignore the error.
			}
		}
	}

	if (!aSendRaw && !_tcsnicmp(aKeys, _T("{Text}"), 6))
	{
		// Setting this early allows CapsLock and the Win+L workaround to be skipped:
		aSendRaw = SCM_RAW_TEXT;
		aKeys += 6;
	}

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
		if (   SystemHasAnotherKeybdHook() // This function has been benchmarked to ensure it doesn't yield our timeslice, etc.  200 calls take 0ms according to tick-count, even when CPU is maxed.
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
			&& aSendRaw != SCM_RAW_TEXT // {Text} mode does not trigger Win+L.
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
				LPCTSTR L_pos, brace_pos;
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
		// v1.1.27.01: Use the thread of the focused control, which may differ from the active window.
		keybd_layout_thread = GetFocusedCtrlThread();
	}
	sTargetKeybdLayout = GetKeyboardLayout(keybd_layout_thread); // If keybd_layout_thread==0, this will get our thread's own layout, which seems like the best/safest default.
	sTargetLayoutHasAltGr = LayoutHasAltGr(sTargetKeybdLayout);  // Note that WM_INPUTLANGCHANGEREQUEST is not monitored by MsgSleep for the purpose of caching our thread's keyboard layout.  This is because it would be unreliable if another msg pump such as MsgBox is running.  Plus it hardly helps perf. at all, and hurts maintainability.

	// Below is now called with "true" so that the hook's modifier state will be corrected (if necessary)
	// prior to every send.
	modLR_type mods_current = GetModifierLRState(true); // Current "logical" modifier state.

	// For any modifiers put in the "down" state by {xxx DownR}, keep only those which
	// are still logically down before each Send starts.  Otherwise each Send would reset
	// the modifier to "down" even after the user "releases" it by some other means.
	sModifiersLR_remapped &= mods_current;

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
	// a modifier was detected as persistent just because A_HotkeyModifierTimeout expired
	// while the user was still holding down the key, but then when the user released it,
	// this logic here would think it's still persistent and push it back down again
	// to enforce it as "always-down" during the send operation.  Thus, the key would
	// basically get stuck down even after the send was over:
	sModifiersLR_persistent &= mods_current & ~mods_down_physically_and_logically;
	modLR_type persistent_modifiers_for_this_SendKeys;
	modLR_type mods_released_for_selective_blind = 0;
	if (sInBlindMode)
	{
		// The following value is usually zero unless the user is currently holding down
		// some modifiers as part of a hotkey. These extra modifiers are the ones that
		// this send operation (along with all its calls to SendKey and similar) should
		// consider to be down for the duration of the Send (unless they go up via an
		// explicit {LWin up}, etc.)
		persistent_modifiers_for_this_SendKeys = mods_current;
		if (mods_excluded_from_blind) // Caller specified modifiers to exclude from Blind treatment.
		{
			persistent_modifiers_for_this_SendKeys &= ~mods_excluded_from_blind;
			mods_released_for_selective_blind = mods_current ^ persistent_modifiers_for_this_SendKeys;
		}
	}
	else
	{
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
	// Remember that apps like MS Word have an auto-correct feature that might make it
	// wrongly seem that the turning off of Capslock below needs a Sleep(0) to take effect.
	prior_capslock_state = g.StoreCapslockMode && !sInBlindMode && aSendRaw != SCM_RAW_TEXT
		? ToggleKeyState(VK_CAPITAL, TOGGLED_OFF)
		: TOGGLE_INVALID; // In blind mode, don't do store capslock (helps remapping and also adds flexibility).

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
		&& !sSendMode && !aTargetWindow;
	if (do_selective_blockinput)
		OurBlockInput(true); // Turn it on unconditionally even if it was on, since Ctrl-Alt-Del might have disabled it.

	vk_type vk;
	sc_type sc;
	modLR_type key_as_modifiersLR = 0;
	modLR_type mods_for_next_key = 0;
	// Above: For v1.0.35, it was changed to modLR vs. mod so that AltGr keys such as backslash and '{'
	// are supported on layouts such as German when sending to apps such as Putty that are fussy about
	// which ALT key is held down to produce the character.
	vk_type this_event_modifier_down;
	size_t key_text_length, key_name_length;
	LPCTSTR end_pos, next_word;
	TCHAR key_text[1024]; // Make it reasonably large to support any conceivable {Click ...} usage.
	KeyEventTypes event_type;
	int repeat_count, click_x, click_y;
	bool move_offset;
	enum { KEYDOWN_TEMP = 0, KEYDOWN_PERSISTENT, KEYDOWN_REMAP } key_down_type;
	DWORD placeholder;

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

				// Make a modifiable null-terminated copy to simplify comparisons etc.
				if (key_text_length >= _countof(key_text))
					goto brace_case_end; // Skip this unreasonably long (probably invalid) item.
				tmemcpy(key_text, aKeys, key_text_length);
				key_text[key_text_length] = '\0';

				if (!_tcsnicmp(key_text, _T("Click"), 5))
				{
					ParseClickOptions(omit_leading_whitespace(key_text + 5), click_x, click_y, vk
						, event_type, repeat_count, move_offset);
					if (repeat_count < 1) // Allow {Click 100, 100, 0} to do a mouse-move vs. click (but modifiers like ^{Click..} aren't supported in this case.
						MouseMove(click_x, click_y, placeholder, g.DefaultMouseSpeed, move_offset);
					else // Use SendKey because it supports modifiers (e.g. ^{Click}) SendKey requires repeat_count>=1.
						SendKey(vk, 0, mods_for_next_key, persistent_modifiers_for_this_SendKeys
							, repeat_count, event_type, 0, aTargetWindow, click_x, click_y, move_offset);
					goto brace_case_end; // This {} item completely handled, so move on to next.
				}
				else if (!_tcsicmp(key_text, _T("Raw"))) // This is used by auto-replace hotstrings too.
				{
					// As documented, there's no way to switch back to non-raw mode afterward since there's no
					// correct way to support special (non-literal) strings such as {Raw Off} while in raw mode.
					aSendRaw = SCM_RAW;
					goto brace_case_end; // This {} item completely handled, so move on to next.
				}
				else if (!_tcsicmp(key_text, _T("Text"))) // Added in v1.1.27
				{
					aSendRaw = SCM_RAW_TEXT;
					goto brace_case_end; // This {} item completely handled, so move on to next.
				}

				// Since above didn't "goto", this item isn't {Click}.
				event_type = KEYDOWNANDUP;         // Set defaults.
				repeat_count = 1;                  //
				key_name_length = key_text_length; //

				if (auto space_pos = StrChrAny(key_text, _T(" \t"))) // Assign. Also, it relies on the fact that {} key names contain no spaces.
				{
					*space_pos = '\0';  // Terminate here so that TextToVK() can properly resolve a single char.
					key_name_length = space_pos - key_text; // Override the default value set above.
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
							if (!_tcsnicmp(next_word + 4, _T("Temp"), 4)) // "DownTemp" means non-persistent.
								key_down_type = KEYDOWN_TEMP;
							else if (toupper(next_word[4] == 'R')) // "DownR" means treated as a physical modifier (R = remap); i.e. not kept down during Send, but restored after Send (unlike Temp).
								key_down_type = KEYDOWN_REMAP;
							else
								key_down_type = KEYDOWN_PERSISTENT;
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

				TextToVKandSC(key_text, vk, sc, &mods_for_next_key, sTargetKeybdLayout);

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
								if (key_down_type == KEYDOWN_PERSISTENT) // v1.0.44.05.
									sModifiersLR_persistent |= key_as_modifiersLR;
								else if (key_down_type == KEYDOWN_REMAP) // v1.1.27.00
									sModifiersLR_remapped |= key_as_modifiersLR;
								persistent_modifiers_for_this_SendKeys |= key_as_modifiersLR; // v1.0.44.06: Added this line to fix the fact that "DownTemp" should keep the key pressed down after the send.
							}
							else if (event_type == KEYUP) // *not* KEYDOWNANDUP, since that would be an intentional activation of the Start Menu or menu bar.
							{
								DisguiseWinAltIfNeeded(vk);
								sModifiersLR_persistent &= ~key_as_modifiersLR;
								sModifiersLR_remapped &= ~key_as_modifiersLR;
								persistent_modifiers_for_this_SendKeys &= ~key_as_modifiersLR;
								// Fix for v1.0.43: Also remove LControl if this key happens to be AltGr.
								if (vk == VK_RMENU && sTargetLayoutHasAltGr == CONDITION_TRUE) // It is AltGr.
									persistent_modifiers_for_this_SendKeys &= ~MOD_LCONTROL;
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
						{
							// Although MSDN says WM_CHAR uses UTF-16, it seems to really do automatic
							// translation between ANSI and UTF-16; we rely on this for correct results:
							for (int i = 0; i < repeat_count; ++i)
								PostMessage(aTargetWindow, WM_CHAR, key_text[0], 0);
						}
						else
							SendKeySpecial(key_text[0], repeat_count, mods_for_next_key | persistent_modifiers_for_this_SendKeys);
					}
				}

				// See comment "else must never change sModifiersLR_persistent" above about why
				// !aTargetWindow is used below:
				else if (vk = TextToSpecial(key_text, key_text_length, event_type
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

				else if (key_text_length > 4 && !_tcsnicmp(key_text, _T("ASC "), 4) && !aTargetWindow) // {ASC nnnnn}
				{
					// Include the trailing space in "ASC " to increase uniqueness (selectivity).
					// Also, sending the ASC sequence to window doesn't work, so don't even try:
					SendASC(omit_leading_whitespace(key_text + 3));
					// Do this only once at the end of the sequence:
					DoKeyDelay(); // It knows not to do the delay for SM_INPUT.
				}

				else if (key_text_length > 2 && !_tcsnicmp(key_text, _T("U+"), 2))
				{
					// L24: Send a unicode value as shown by Character Map.
					UINT u_code = (UINT) _tcstol(key_text + 2, NULL, 16);
					wchar_t wc1, wc2;
					if (u_code >= 0x10000)
					{
						// Supplementary characters are encoded as UTF-16 and split into two messages.
						u_code -= 0x10000;
						wc1 = 0xd800 + ((u_code >> 10) & 0x3ff);
						wc2 = 0xdc00 + (u_code & 0x3ff);
					}
					else
					{
						wc1 = (wchar_t) u_code;
						wc2 = 0;
					}
					if (aTargetWindow)
					{
						// Although MSDN says WM_CHAR uses UTF-16, PostMessageA appears to truncate it to 8-bit.
						// This probably means it does automatic translation between ANSI and UTF-16.  Since we
						// specifically want to send a Unicode character value, use PostMessageW:
						PostMessageW(aTargetWindow, WM_CHAR, wc1, 0);
						if (wc2)
							PostMessageW(aTargetWindow, WM_CHAR, wc2, 0);
					}
					else
					{
						// Use SendInput in unicode mode if available, otherwise fall back to SendASC.
						// To know why the following requires sSendMode != SM_PLAY, see SendUnicodeChar.
						if (sSendMode != SM_PLAY)
						{
							SendUnicodeChar(wc1, mods_for_next_key | persistent_modifiers_for_this_SendKeys);
							if (wc2)
								SendUnicodeChar(wc2, mods_for_next_key | persistent_modifiers_for_this_SendKeys);
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
			if (aSendRaw == SCM_RAW_TEXT)
			{
				// \b needs to produce VK_BACK for auto-replace hotstrings to work (this is more useful anyway).
				// \r and \n need to produce VK_RETURN for decent compatibility.  SendKeySpecial('\n') works for
				// some controls (such as Scintilla) but has no effect in other common applications.
				// \t has more utility if translated to VK_TAB.  SendKeySpecial('\t') has no effect in many
				// common cases, and seems to only work in cases where {tab} would work just as well.
				switch (*aKeys)
				{
				case '\r': // Translate \r but ignore any trailing \n, since \r\n -> {Enter 2} is counter-intuitive.
					if (aKeys[1] == '\n')
						++aKeys;
					// Fall through:
				case '\n': vk = VK_RETURN; break;
				case '\b': vk = VK_BACK; break;
				case '\t': vk = VK_TAB; break;
				default: vk = 0; break; // Send all other characters via SendKeySpecial()/WM_CHAR.
				}
			}
			else
			{
				// Best to call this separately, rather than as first arg in SendKey, since it changes the
				// value of modifiers and the updated value is *not* guaranteed to be passed.
				// In other words, SendKey(TextToVK(...), modifiers, ...) would often send the old
				// value for modifiers.
				vk = CharToVKAndModifiers(*aKeys, &mods_for_next_key, sTargetKeybdLayout
					, (mods_for_next_key | persistent_modifiers_for_this_SendKeys) != 0 && !aSendRaw); // v1.1.27.00: Disable the a-z to vk41-vk5A fallback translation when modifiers are present since it would produce the wrong printable characters.
				// CharToVKAndModifiers() takes no measurable time compared to the amount of time SendKey takes.
			}
			if (vk)
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
					SendKeySpecial(*aKeys, 1, mods_for_next_key | persistent_modifiers_for_this_SendKeys);
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
			// 1) It avoids the need for A_HotkeyModifierTimeout (and it's superior to it) for both SendInput
			//    and SendPlay.
			// 2) The hook will not be present during the SendInput, nor can it be reinstalled in time to
			//    catch any physical events generated by the user during the Send. Consequently, there is no
			//    known way to reliably detect physical keystate changes.
			// 3) Changes made to modifier state by SendPlay are seen only by the active window's thread.
			//    Thus, it would be inconsistent and possibly incorrect to adjust global modifier state
			//    after (or during) a SendPlay.
			// So rather than resorting to A_HotkeyModifierTimeout, we can restore the modifiers within the
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
				| sModifiersLR_remapped // Restore any modifiers which were put in the down state by remappings or {key DownR} prior to this Send.
				| (sInBlindMode ? mods_released_for_selective_blind
					: (mods_down_physically_orig & ~mods_down_physically_but_not_logically_orig)); // The last item is usually 0.
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

		// Put any modifiers in sModifiersLR_remapped back into effect, as if they were physically down.
		mods_down_physically |= sModifiersLR_remapped;

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
			// For selective {Blind!#^+}, restore any modifiers that were automatically released at the
			// start, such as for *^1::Send "{Blind^}2" when Ctrl+Alt+1 is pressed (Ctrl is released).
			// But do this before the below so that if the key was physically down at the start and has
			// since been released, it won't be pushed back down.
			mods_to_set |= mods_released_for_selective_blind;
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
		// g_modifiersLR_numpad_mask is used to work around an issue where our changes to shift-key state
		// trigger the system's shift-numpad handling (usually in combination with actual user input),
		// which in turn causes the Shift key to stick down.  If non-zero, the Shift key is currently "up"
		// but should be "released" anyway, since the system will inject Shift-down either before the next
		// keyboard event or after the Numpad key is released.  Find "fake shift" for more details.
		SetModifierLRState(mods_to_set, GetModifierLRState() | g_modifiersLR_numpad_mask, aTargetWindow, true, true); // It also does DoKeyDelay(g->PressDuration).
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
		OurBlockInput(false);

	// The following MsgSleep(-1) solves unwanted buffering of hotkey activations while SendKeys is in progress
	// in a non-Critical thread.  Because SLEEP_WITHOUT_INTERRUPTION is used to perform key delays, any incoming
	// hotkey messages would be left in the queue.  It is not until the next interruptible sleep that hotkey
	// messages may be processed, and potentially discarded due to #MaxThreadsPerHotkey (even #MaxThreadsBuffer
	// should only allow one buffered activation).  But if the hotkey thread just calls Send in a loop and then
	// returns, it never performs an interruptible sleep, so the hotkey messages are processed one by one after
	// each new hotkey thread returns, even though Critical was not used.  Also note SLEEP_WITHOUT_INTERRUPTION
	// causes g_script.mLastScriptRest to be reset, so it's unlikely that a sleep would occur between Send calls.
	// To solve this, call MsgSleep(-1) now (unless no delays were performed, or the thread is uninterruptible):
	if (aSendModeOrig == SM_EVENT && g_script.mLastPeekTime != orig_last_peek_time && IsInterruptible())
		MsgSleep(-1);

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
				, aTargetWindow, false, true, g->SendLevel ? KEY_IGNORE_LEVEL(g->SendLevel) : KEY_IGNORE); // See keyboard_mouse.h for explanation of KEY_IGNORE.
			// Above: Fixed for v1.1.27 to use KEY_IGNORE except when SendLevel is non-zero (since that
			// would indicate that the script probably wants to trigger a hotkey).  KEY_IGNORE is used
			// (and was prior to v1.1.06.00) to prevent the temporary modifier state changes here from
			// interfering with the use of hotkeys while a Send is in progress.
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
		modLR_type win_alt_to_be_released = (state_now & ~aModifiersLRPersistent) // The modifiers to be released...
			& (MOD_LWIN|MOD_RWIN|MOD_LALT|MOD_RALT); // ... but restrict them to only Win/Alt.
		if (win_alt_to_be_released)
		{
			// Originally used the following for mods new/now: state_now & ~win_alt_to_be_released, state_now
			// When AltGr is to be released, the above formula passes LCtrl+RAlt as the current state and just
			// LCtrl as the new state, which results in LCtrl being pushed back down after it is released via
			// AltGr.  Although our caller releases LCtrl if needed, it usually uses KEY_IGNORE, so if we put
			// LCtrl down here, it would be wrongly stuck down in g_modifiersLR_logical_non_ignored, which
			// causes ^-modified hotkeys to fire when they shouldn't and prevents non-^ hotkeys from firing.
			// By ignoring the current modifier state and only specifying the modifiers we want released,
			// we avoid any chance of sending any unwanted key-down:
			SetModifierLRState(0, win_alt_to_be_released, aTargetWindow, true, false); // It also does DoKeyDelay(g->PressDuration).
		}
	}
}



void SendKeySpecial(TCHAR aChar, int aRepeatCount, modLR_type aModifiersLR)
// Caller must be aware that keystrokes are sent directly (i.e. never to a target window via ControlSend mode).
// It must also be aware that the event type KEYDOWNANDUP is always what's used since there's no way
// to support anything else.  Furthermore, there's no way to support "modifiersLR_for_next_key" such as ^€
// (assuming € is a character for which SendKeySpecial() is required in the current layout) with ASC mode.
// This function uses some of the same code as SendKey() above, so maintain them together.
{
	// Caller must verify that aRepeatCount >= 1.
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

	bool use_sendasc = sSendMode == SM_PLAY; // See SendUnicodeChar for why it isn't called for SM_PLAY.
	TCHAR asc_string[16];
	WCHAR wc;

	if (use_sendasc)
	{
		// The following range isn't checked because this function appears never to be called for such
		// characters (tested in English and Russian so far), probably because VkKeyScan() finds a way to
		// manifest them via Control+VK combinations:
		//if (aChar > -1 && aChar < 32)
		//	return;
		TCHAR *cp = asc_string;
		if (aChar & ~127)    // Try using ANSI.
			*cp++ = '0';  // ANSI mode is achieved via leading zero in the Alt+Numpad keystrokes.
		//else use Alt+Numpad without the leading zero, which allows the characters a-z, A-Z, and quite
		// a few others to be produced in Russian and perhaps other layouts, which was impossible in versions
		// prior to 1.0.40.
		_itot((TBYTE)aChar, cp, 10); // Convert to UCHAR in case aChar < 0.
	}
	else
	{
#ifdef UNICODE
		wc = aChar;
#else
		// Convert our ANSI character to Unicode for sending.
		if (!MultiByteToWideChar(CP_ACP, MB_ERR_INVALID_CHARS, &aChar, 1, &wc, 1))
			return;
#endif
	}

	LONG_OPERATION_INIT
	for (int i = 0; i < aRepeatCount; ++i)
	{
		if (!sSendMode)
			LONG_OPERATION_UPDATE_FOR_SENDKEYS
		if (use_sendasc)
			SendASC(asc_string);
		else
			SendUnicodeChar(wc, aModifiersLR);
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
			// active because the user doesn't need it; also for some games maybe).
			aSC = vk_to_sc(aVK);

	BYTE aSC_lobyte = LOBYTE(aSC);
	DWORD event_flags = HIBYTE(aSC) ? KEYEVENTF_EXTENDEDKEY : 0;

	// v1.0.43: Apparently, the journal playback hook requires neutral modifier keystrokes
	// rather than left/right ones.  Otherwise, the Shift key can't capitalize letters, etc.
	if (sSendMode == SM_PLAY)
	{
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

		ResultType target_layout_has_altgr = caller_is_keybd_hook ? LayoutHasAltGr(GetFocusedKeybdLayout())
			: sTargetLayoutHasAltGr;
		bool hookable_altgr = aVK == VK_RMENU && target_layout_has_altgr && !put_event_into_array && g_KeybdHook;

		// Calculated only once for performance (and avoided entirely if not needed):
		modLR_type key_as_modifiersLR = put_event_into_array ? KeyToModifiersLR(aVK, aSC) : 0;

		bool do_key_history = !caller_is_keybd_hook // If caller is hook, don't log because it does.
			&& sSendMode != SM_PLAY  // In playback mode, the journal hook logs so that timestamps are accurate.
			&& (!g_KeybdHook || sSendMode == SM_INPUT); // In the remaining cases, log only when the hook isn't installed or won't be at the time of the event.

		if (aEventType != KEYUP)  // i.e. always do it for KEYDOWNANDUP
		{
			if (put_event_into_array)
				PutKeybdEventIntoArray(key_as_modifiersLR, aVK, aSC, event_flags, aExtraInfo);
			else
			{
				// The following global is used to flag as our own the keyboard driver's LCtrl-down keystroke
				// that is triggered by RAlt-down (AltGr).  The hook prevents those keystrokes from triggering
				// hotkeys such as "*Control::" anyway, but this ensures the LCtrl-down is marked as "ignored"
				// and given the correct SendLevel.  It may fix other obscure side-effects and bugs, since the
				// event should be considered script-generated even though indirect.  Note: The problem with
				// having the hook detect AltGr's automatic LControl-down is that the keyboard driver seems
				// to generate the LControl-down *before* notifying the system of the RAlt-down.  That makes
				// it impossible for the hook to automatically adjust its SendLevel based on the RAlt-down.
				if (hookable_altgr)
					g_AltGrExtraInfo = aExtraInfo;
				keybd_event(aVK, aSC_lobyte // naked scan code (the 0xE0 prefix, if any, is omitted)
					, event_flags, aExtraInfo);
				g_AltGrExtraInfo = 0; // Unconditional reset.
			}

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
				if (hookable_altgr) // See comments in similar section above for details.
					g_AltGrExtraInfo = aExtraInfo;
				keybd_event(aVK, aSC_lobyte, event_flags, aExtraInfo);
				g_AltGrExtraInfo = 0; // Unconditional reset.
			}
			// The following is done to avoid an extraneous artificial {LCtrl Up} later on,
			// since the keyboard driver should insert one in response to this {RAlt Up}:
			if (target_layout_has_altgr && aSC == SC_RALT)
				sEventModifiersLR &= ~MOD_LCONTROL;

			if (do_key_history)
				UpdateKeyEventHistory(true, aVK, aSC);
		}
	}

	if (aDoKeyDelay) // SM_PLAY also uses DoKeyDelay(): it stores the delay item in the event array.
		DoKeyDelay(); // Thread-safe because only called by main thread in this mode.  See notes above.
}



void KeyEventMenuMask(KeyEventTypes aEventType, DWORD aExtraInfo)
// Send a menu masking key event (use of this function reduces code size).
{
	KeyEvent(aEventType, g_MenuMaskKeyVK, g_MenuMaskKeySC, NULL, false, aExtraInfo);
}



///////////////////
// Mouse related //
///////////////////


BIF_DECL(BIF_Click)
{
	TCHAR args[1024]; // Arbitrary size, significantly larger than anything likely to be valid.
	*args = '\0';
	for (int i = 0; i < aParamCount; ++i)
	{
		if (TokenToObject(*aParam[i]))
			_f_throw_param(i);
		LPTSTR arg = TokenToString(*aParam[i], _f_number_buf);
		sntprintfcat(args, _countof(args), _T("%s,"), arg);
	}
	PerformClick(args);
}



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
	aVK = VK_LBUTTON;
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

		// Temp termination for IsNumeric(), ConvertMouseButton(), and peace-of-mind.
		orig_char = *option_end;
		*option_end = '\0';

		// Parameters can occur in almost any order to enhance usability (at the cost of
		// slightly diminishing the ability unambiguously add more parameters in the future).
		// Seems okay to support floats because ATOI() will just omit the decimal portion.
		if (IsNumeric(next_option, true, false, true))
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
			if (temp_vk = Line::ConvertMouseButton(next_option, true))
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



FResult PerformMouse(ActionTypeType aActionType, optl<StrArg> aButton
	, optl<int> aX1, optl<int> aY1, optl<int> aX2, optl<int> aY2
	, optl<int> aSpeed, optl<StrArg> aOffsetMode, optl<int> aRepeatCount, optl<StrArg> aDownUp)
{
	vk_type vk;
	if (aActionType == ACT_MOUSEMOVE)
		vk = 0;
	else
		// ConvertMouseButton() treats omitted/blank as "Left":
		if (   !(vk = Line::ConvertMouseButton(aButton.value_or_null(), aActionType == ACT_MOUSECLICK)))
			return FR_E_ARG(0); // of MouseClick/MouseClickDrag

	KeyEventTypes event_type = KEYDOWNANDUP;  // Set defaults.
	int repeat_count = aRepeatCount.has_value() ? *aRepeatCount : 1;
	if (repeat_count < 0)
		return FR_E_ARG(2); // of MouseClick

	if (aDownUp.has_value())
	{
		switch (*aDownUp.value())
		{
		case 'u':
		case 'U':
			event_type = KEYUP;
			break;
		case 'd':
		case 'D':
			event_type = KEYDOWN;
			break;
		case '\0':
			break;
		default:
			return FR_E_ARG(5); // of MouseClick
		}
	}

	bool move_offset = false;
	if (aOffsetMode.has_nonempty_value())
	{
		if (ctoupper(*aOffsetMode.value()) == 'R')
			move_offset = true;
		else
			return FR_E_ARG(aActionType == ACT_MOUSEMOVE ? 3 : 6);
	}

	PerformMouseCommon(aActionType, vk
		, aX1.value_or(COORD_UNSPECIFIED)  // If no starting coords are specified, mark it as "use the
		, aY1.value_or(COORD_UNSPECIFIED)  // current mouse position":
		, aX2.value_or(COORD_UNSPECIFIED)  // These two are non-null only for MouseClickDrag.
		, aY2.value_or(COORD_UNSPECIFIED)  //
		, repeat_count, event_type
		, aSpeed.value_or(g->DefaultMouseSpeed)
		, move_offset);

	return OK;
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
		if (SystemHasAnotherMouseHook()) // See similar section in SendKeys() for details.
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
		&& !sSendMode;
	if (do_selective_blockinput) // It seems best NOT to use g_BlockMouseMove for this, since often times the user would want keyboard input to be disabled too, until after the mouse event is done.
		OurBlockInput(true); // Turn it on unconditionally even if it was on, since Ctrl-Alt-Del might have disabled it.

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
		OurBlockInput(false);
}



void MouseClickDrag(vk_type aVK, int aX1, int aY1, int aX2, int aY2, int aSpeed, bool aMoveOffset)
{
	// Check if one of the coordinates is missing, which can happen in cases where this was called from
	// a source that didn't already validate it.
	if (   (aX1 == COORD_UNSPECIFIED && aY1 != COORD_UNSPECIFIED) || (aX1 != COORD_UNSPECIFIED && aY1 == COORD_UNSPECIFIED)
		|| (aX2 == COORD_UNSPECIFIED && aY2 != COORD_UNSPECIFIED) || (aX2 != COORD_UNSPECIFIED && aY2 == COORD_UNSPECIFIED)   )
		return;

	// I asked Jon, "Have you discovered that insta-drags almost always fail?" and he said
	// "Yeah, it was weird, absolute lack of drag... Don't know if it was my config or what."
	// However, testing reveals "insta-drags" work ok, at least on my system, so leaving them enabled.
	// User can easily increase the speed if there's any problem:
	//if (aSpeed < 2)
	//	aSpeed = 2;

	// v2.0: Always translate logical buttons into physical ones.  Which physical button it becomes depends
	// on whether the mouse buttons are swapped via the Control Panel.  Note that journal playback doesn't
	// need the swap because every aspect of it is "logical".
	if ((aVK == VK_LBUTTON || aVK == VK_RBUTTON) && sSendMode != SM_PLAY && GetSystemMetrics(SM_SWAPBUTTON))
		aVK = (aVK == VK_LBUTTON) ? VK_RBUTTON : VK_LBUTTON;

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

	// v2.0: Always translate logical buttons into physical ones.  Which physical button it becomes depends
	// on whether the mouse buttons are swapped via the Control Panel.
	if ((aVK == VK_LBUTTON || aVK == VK_RBUTTON) && sSendMode != SM_PLAY && GetSystemMetrics(SM_SWAPBUTTON))
		aVK = (aVK == VK_LBUTTON) ? VK_RBUTTON : VK_LBUTTON;

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
	if (sAbortArraySend) // A prior call failed (might be impossible).  Avoid malloc() in this case.
		return FAIL;
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
	{
		sAbortArraySend = true; // Usually better to send nothing rather than partial.
		// Leave sEventSI and sMaxEvents in their current valid state, to be freed by CleanupEventArray().
		return FAIL;
	}
	else // Copy old array into new memory area (note that sEventSI and sEventPB are different views of the same variable).
		memcpy(new_mem, sEventSI, sEventCount * event_size);
	if (sMaxEvents > (sSendMode == SM_INPUT ? MAX_INITIAL_EVENTS_SI : MAX_INITIAL_EVENTS_PB))
		free(sEventSI); // Previous block was malloc'd vs. _alloc'd, so free it.
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

		SendInput(sEventCount, sEventSI, sizeof(INPUT)); // Must call dynamically-resolved version for Win95/NT compatibility.
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
	if (   !(g_PlaybackHook = SetWindowsHookEx(WH_JOURNALPLAYBACK, PlaybackProc, g_hInstance, 0))   )
		return;
	// During playback, have the keybd hook (if it's installed) block presses of the Windows key.
	// This is done because the Windows key is about the only key (other than Ctrl-Esc or Ctrl-Alt-Del)
	// that takes effect immediately rather than getting buffered/postponed until after the playback.
	// It should be okay to set this after the playback hook is installed because playback shouldn't
	// actually begin until we have our thread do its first MsgSleep later below.
	g_BlockWinKeys = true;

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
	// because SendKeys() is designed to be uninterruptible by other script threads;
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
	// hook about these issues.  Here are more details from an older comment:
	// Always sleep a certain minimum amount of time between events to improve reliability,
	// but allow the user to specify a higher time if desired.  A true sleep is done if the
	// delay period is small.  This fixes a small issue where if LButton is a hotkey that
	// includes "MouseClick left" somewhere in its subroutine, the script's own main window's
	// title bar buttons for min/max/close would not properly respond to left-clicks.
	if (mouse_delay < 11)
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



// Allocates or resizes and resets g_KeyHistory.
// Must be called from the hook thread if it's running, otherwise it can be called from the main thread.
void SetKeyHistoryMax(int aMax)
{
	free(g_KeyHistory);
	g_KeyHistory = aMax ? (KeyHistoryItem *)malloc(aMax * sizeof(KeyHistoryItem)) : nullptr;
	if (g_KeyHistory)
	{
		ZeroMemory(g_KeyHistory, aMax * sizeof(KeyHistoryItem)); // Must be zeroed.
		g_MaxHistoryKeys = aMax;
		g_HistoryTickPrev = GetTickCount();
		g_HistoryHwndPrev = NULL;
	}
	else
		g_MaxHistoryKeys = 0;
	g_KeyHistoryNext = 0;
}



ToggleValueType ToggleKeyState(vk_type aVK, ToggleValueType aToggleValue)
// Toggle the given aVK into another state.  For performance, it is the caller's responsibility to
// ensure that aVK is a toggleable key such as capslock, numlock, insert, or scrolllock.
// Returns the state the key was in before it was changed.
{
	// Can't use IsKeyDownAsync/GetAsyncKeyState() because it doesn't have this info:
	ToggleValueType starting_state = IsKeyToggledOn(aVK) ? TOGGLED_ON : TOGGLED_OFF;
	if (aToggleValue != TOGGLED_ON && aToggleValue != TOGGLED_OFF) // Shouldn't be called this way.
		return starting_state;
	if (starting_state == aToggleValue) // It's already in the desired state, so just return the state.
		return starting_state;
	//if (aVK == VK_NUMLOCK) // v1.1.22.05: This also applies to CapsLock and ScrollLock.
	{
		// If the key is being held down, sending a KEYDOWNANDUP won't change its toggle
		// state unless the key is "released" first.  This has been confirmed for NumLock,
		// CapsLock and ScrollLock on Windows 2000 (in a VM) and Windows 10.
		// Examples where problems previously occurred:
		//   ~CapsLock & x::Send abc  ; Produced "ABC"
		//   ~CapsLock::Send abc  ; Alternated between "abc" and "ABC", even without {Blind}
		//   ~ScrollLock::SetScrollLockState Off  ; Failed to change state
		// The behaviour can still be observed by sending the keystrokes manually:
		//   ~NumLock::Send {NumLock}  ; No effect
		//   ~NumLock::Send {NumLock up}{NumLock}  ; OK
		// OLD COMMENTS:
		// Sending an extra up-event first seems to prevent the problem where the Numlock
		// key's indicator light doesn't change to reflect its true state (and maybe its
		// true state doesn't change either).  This problem tends to happen when the key
		// is pressed while the hook is forcing it to be either ON or OFF (or it suppresses
		// it because it's a hotkey).  Needs more testing on diff. keyboards & OSes:
		if (IsKeyDown(aVK))
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


/*
void SetKeyState (vk_type vk, int aKeyUp)
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
				KeyEventMenuMask(KEYDOWNANDUP, aExtraInfo); // Disguise key release to suppress Start Menu.
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
			KeyEventMenuMask(KEYDOWN, aExtraInfo); // Ensures that the Start Menu does not appear.
		KeyEvent(KEYDOWN, VK_LWIN, 0, NULL, false, aExtraInfo);
		if (disguise_win_down)
			KeyEventMenuMask(KEYUP, aExtraInfo); // Ensures that the Start Menu does not appear.
	}

	if (release_rwin)
	{
		if (!defer_win_release)
		{
			if (ctrl_nor_shift_nor_alt_down && aDisguiseUpWinAlt && sSendMode != SM_PLAY)
				KeyEventMenuMask(KEYDOWNANDUP, aExtraInfo); // Disguise key release to suppress Start Menu.
			KeyEvent(KEYUP, VK_RWIN, 0, NULL, false, aExtraInfo);
		}
		// else release it only after the normal operation of the function pushes down the disguise keys.
	}
	else if (!(aModifiersLRnow & MOD_RWIN) && (aModifiersLRnew & MOD_RWIN)) // Press down RWin.
	{
		if (disguise_win_down)
			KeyEventMenuMask(KEYDOWN, aExtraInfo); // Ensures that the Start Menu does not appear.
		KeyEvent(KEYDOWN, VK_RWIN, 0, NULL, false, aExtraInfo);
		if (disguise_win_down)
			KeyEventMenuMask(KEYUP, aExtraInfo); // Ensures that the Start Menu does not appear.
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
				KeyEventMenuMask(KEYDOWNANDUP, aExtraInfo); // Disguise key release to suppress menu activation.
			KeyEvent(KEYUP, VK_LMENU, 0, NULL, false, aExtraInfo);
		}
	}
	else if (!(aModifiersLRnow & MOD_LALT) && (aModifiersLRnew & MOD_LALT))
	{
		if (disguise_alt_down)
			KeyEventMenuMask(KEYDOWN, aExtraInfo); // Ensures that menu bar is not activated.
		KeyEvent(KEYDOWN, VK_LMENU, 0, NULL, false, aExtraInfo);
		if (disguise_alt_down)
			KeyEventMenuMask(KEYUP, aExtraInfo);
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
					KeyEventMenuMask(KEYDOWNANDUP, aExtraInfo); // Disguise key release to suppress menu activation.
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
			KeyEventMenuMask(KEYDOWN, aExtraInfo); // Ensures that menu bar is not activated.
			KeyEvent(KEYDOWN, VK_RMENU, 0, NULL, false, aExtraInfo);
			KeyEventMenuMask(KEYUP, aExtraInfo);
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
			if (sTargetLayoutHasAltGr == CONDITION_TRUE)
			{
				if (aModifiersLRnow & MOD_LCONTROL)
					KeyEvent(KEYUP, VK_LCONTROL, 0, NULL, false, aExtraInfo);
				if (aModifiersLRnow & MOD_RCONTROL)
				{
					// Release RCtrl before pressing AltGr, because otherwise the system will not put
					// LCtrl into effect, but it will still inject LCtrl-up when AltGr is released.
					// With LCtrl not in effect and RCtrl being released below, AltGr would instead
					// act as pure RAlt, which would not have the right effect.
					// RCtrl will be put back into effect below if aModifiersLRnew & MOD_RCONTROL.
					KeyEvent(KEYUP, VK_RCONTROL, 0, NULL, false, aExtraInfo);
					aModifiersLRnow &= ~MOD_RCONTROL;
				}
			}
			KeyEvent(KEYDOWN, VK_RMENU, 0, NULL, false, aExtraInfo);
			if (sTargetLayoutHasAltGr == CONDITION_TRUE) // Note that KeyEvent() might have just changed the value of sTargetLayoutHasAltGr.
			{
				// Indicate that control is both down and required down so that the section after this one won't
				// release it.  Without this fix, a hotkey that sends an AltGr char such as "^ä:: Send '{Raw}{'"
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
	if (IsKeyDownAsync(VK_LSHIFT))   modifiersLR |= MOD_LSHIFT;
	if (IsKeyDownAsync(VK_RSHIFT))   modifiersLR |= MOD_RSHIFT;
	if (IsKeyDownAsync(VK_LCONTROL)) modifiersLR |= MOD_LCONTROL;
	if (IsKeyDownAsync(VK_RCONTROL)) modifiersLR |= MOD_RCONTROL;
	if (IsKeyDownAsync(VK_LMENU))    modifiersLR |= MOD_LALT;
	if (IsKeyDownAsync(VK_RMENU))    modifiersLR |= MOD_RALT;
	if (IsKeyDownAsync(VK_LWIN))     modifiersLR |= MOD_LWIN;
	if (IsKeyDownAsync(VK_RWIN))     modifiersLR |= MOD_RWIN;

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
		// UPDATE: The modifier state might also become incorrect due to keyboard events which
		// are missed due to User Interface Privelege Isolation; i.e. because a window belonging
		// to a process with higher integrity level than our own became active while the key was
		// down, so we saw the down event but not the up event.
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
			// Also reset pPrefixKey if it is one of the wrongly-down modifiers.
			if (pPrefixKey && (pPrefixKey->as_modifiersLR & modifiers_wrongly_down))
				pPrefixKey = NULL;
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



DWORD GetFocusedCtrlThread(HWND *apControl, HWND aWindow)
{
	// Determine the thread for which we want the keyboard layout.
	// When no foreground window, the script's own layout seems like the safest default.
	DWORD thread_id = 0;
	if (aWindow)
	{
		// Get thread of aWindow (which should be the foreground window).
		thread_id = GetWindowThreadProcessId(aWindow, NULL);
		// Get focus.  Benchmarks showed this additional step added only 6% to the time,
		// and the total was only around 4µs per iteration anyway (on a Core i5-4460).
		// It is necessary for UWP apps such as Microsoft Edge, and any others where
		// the top-level window belongs to a different thread than the focused control.
		GUITHREADINFO thread_info;
		thread_info.cbSize = sizeof(thread_info);
		if (GetGUIThreadInfo(thread_id, &thread_info))
		{
			if (thread_info.hwndFocus)
			{
				// Use the focused control's thread.
				thread_id = GetWindowThreadProcessId(thread_info.hwndFocus, NULL);
				if (apControl)
					*apControl = thread_info.hwndFocus;
			}
		}
	}
	return thread_id;
}



HKL GetFocusedKeybdLayout(HWND aWindow)
{
	return GetKeyboardLayout(GetFocusedCtrlThread(NULL, aWindow));
}



bool ActiveWindowLayoutHasAltGr()
// Thread-safety: See comments in LayoutHasAltGr() below.
{
	Get_active_window_keybd_layout // Defines the variable active_window_keybd_layout for use below.
	return LayoutHasAltGr(active_window_keybd_layout) == CONDITION_TRUE; // i.e caller wants both CONDITION_FALSE and LAYOUT_UNDETERMINED to be considered non-AltGr.
}



HMODULE LoadKeyboardLayoutModule(HKL aLayout)
// Loads a keyboard layout DLL and returns its handle.
// Activates the layout as a side-effect, but reverts it if !aSideEffectsOK.
{
	HMODULE hmod = NULL;
	// Unfortunately activating the layout seems to be the only way to retrieve it's name.
	// This may have side-effects in general (such as the language selector flickering),
	// but shouldn't have any in our case since we're only changing layouts for our thread,
	// and only if some other window is active (because if our window was active, aLayout
	// is already the current layout).
	if (HKL old_layout = ActivateKeyboardLayout(aLayout, 0))
	{
		#define KEYBOARD_LAYOUTS_REG_KEY _T("SYSTEM\\CurrentControlSet\\Control\\Keyboard Layouts\\")
		const size_t prefix_length = _countof(KEYBOARD_LAYOUTS_REG_KEY) - 1;
		TCHAR keyname[prefix_length + KL_NAMELENGTH];
		_tcscpy(keyname, KEYBOARD_LAYOUTS_REG_KEY);
		if (GetKeyboardLayoutName(keyname + prefix_length))
		{
			TCHAR layout_file[MAX_PATH]; // It's probably much smaller (like "KBDUSX.dll"), but who knows what whacky custom layouts exist?
			if (ReadRegString(HKEY_LOCAL_MACHINE, keyname, _T("Layout File"), layout_file, _countof(layout_file)))
			{
				hmod = LoadLibrary(layout_file);
			}
		}
		if (aLayout != old_layout)
			ActivateKeyboardLayout(old_layout, 0); // Nothing we can do if it fails.
	}
	return hmod;
}



ResultType LayoutHasAltGrDirect(HKL aLayout)
// Loads and reads the keyboard layout DLL to determine if it has AltGr.
// Activates the layout as a side-effect, but reverts it if !aSideEffectsOK.
// This is fast enough that there's no need to cache these values on startup.
{
	// This abbreviated definition is based on the actual definition in kbd.h (Windows DDK):
	// Updated to use two separate struct definitions, since the pointers are qualified with
	// KBD_LONG_POINTER, which is 64-bit when building for Wow64 (32-bit dll on 64-bit system).
	typedef UINT64 KLP64;
	typedef UINT KLP32;

	// Struct used on 64-bit systems (by both 32-bit and 64-bit programs):
	struct KBDTABLES64 {
		KLP64 pCharModifiers;
		KLP64 pVkToWcharTable;
		KLP64 pDeadKey;
		KLP64 pKeyNames;
		KLP64 pKeyNamesExt;
		KLP64 pKeyNamesDead;
		KLP64 pusVSCtoVK;
		BYTE  bMaxVSCtoVK;
		KLP64 pVSCtoVK_E0;
		KLP64 pVSCtoVK_E1;
		// This is the one we want:
		DWORD fLocaleFlags;
		// Struct definition truncated.
	};

	// Struct used on 32-bit systems:
	struct KBDTABLES32 {
		KLP32 pCharModifiers;
		KLP32 pVkToWcharTable;
		KLP32 pDeadKey;
		KLP32 pKeyNames;
		KLP32 pKeyNamesExt;
		KLP32 pKeyNamesDead;
		KLP32 pusVSCtoVK;
		BYTE  bMaxVSCtoVK;
		KLP32 pVSCtoVK_E0;
		KLP32 pVSCtoVK_E1;
		// This is the one we want:
		DWORD fLocaleFlags;
		// Struct definition truncated.
	};
	
	#define KLLF_ALTGR 0x0001 // Also defined in kbd.h.
	typedef PVOID (* KbdLayerDescriptorType)();

	ResultType result = FAIL;

	if (HMODULE hmod = LoadKeyboardLayoutModule(aLayout))
	{
		KbdLayerDescriptorType kbdLayerDescriptor = (KbdLayerDescriptorType)GetProcAddress(hmod, "KbdLayerDescriptor");
		if (kbdLayerDescriptor)
		{
			PVOID kl = kbdLayerDescriptor();
			DWORD flags = IsOS64Bit() ? ((KBDTABLES64 *)kl)->fLocaleFlags : ((KBDTABLES32 *)kl)->fLocaleFlags;
			result = (flags & KLLF_ALTGR) ? CONDITION_TRUE : CONDITION_FALSE;
		}
		FreeLibrary(hmod);
	}
	return result;
}



ResultType LayoutHasAltGr(HKL aLayout)
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
			return sCachedLayout[i].has_altgr;

	// Since above didn't return, this layout isn't cached yet.  So create a new cache entry for it and
	// determine whether this layout has an AltGr key.  If i<MAX_CACHED_LAYOUTS (which it almost always will be),
	// there's room in the array for a new cache entry.  In the very unlikely event that there isn't room,
	// overwrite an arbitrary item in the array.  An LRU/MRU algorithm (timestamp) isn't used because running out
	// of slots seems too unlikely, and the consequences of running out are merely a slight degradation in performance.
	CachedLayoutType &cl = sCachedLayout[(i < MAX_CACHED_LAYOUTS) ? i : MAX_CACHED_LAYOUTS-1];

	// The old approach here was to call VkKeyScanEx for each character code and find any that require
	// AltGr.  However, that was unacceptably slow for the wider character range of the Unicode build.
	// It was also unreliable (as noted below), so required additional logic in Send and the hook to
	// compensate.  Instead, read the AltGr value directly from the keyboard layout DLL.
	// This method has been compared to the VkKeyScanEx method and another one using Send and hotkeys,
	// and was found to have 100% accuracy for the 203 standard layouts on Windows 10, whereas the
	// VkKeyScanEx method failed for two layouts:
	//   - N'Ko has AltGr but does not appear to use it for anything.
	//   - Ukrainian has AltGr but only uses it for one character, which is also assigned to a naked
	//     VK (so VkKeyScanEx returns that one).  Likely the key in question is absent from some keyboards.
	cl.has_altgr = LayoutHasAltGrDirect(aLayout);
	cl.hkl = aLayout; // This is done here (immediately after has_altgr is set) rather than earlier to minimize the consequences of not being fully thread-safe.
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



sc_type TextToSC(LPCTSTR aText, bool *aSpecifiedByNumber)
{
	if (!*aText) return 0;
	for (int i = 0; i < g_key_to_sc_count; ++i)
		if (!_tcsicmp(g_key_to_sc[i].key_name, aText))
			return g_key_to_sc[i].sc;
	// Do this only after the above, in case any valid key names ever start with SC:
	if (ctoupper(*aText) == 'S' && ctoupper(*(aText + 1)) == 'C')
	{
		LPTSTR endptr;
		sc_type sc = (sc_type)_tcstol(aText + 2, &endptr, 16);  // Convert from hex.
		if (*endptr)
			return 0; // Fixed for v1.1.27: Disallow any invalid suffix so that hotkeys like a::scb() are not misinterpreted as remappings.
		if (aSpecifiedByNumber)
			*aSpecifiedByNumber = true; // Override caller-set default.
		return sc;
	}
	return 0; // Indicate "not found".
}



vk_type TextToVK(LPCTSTR aText, modLR_type *pModifiersLR, bool aExcludeThoseHandledByScanCode, bool aAllowExplicitVK
	, HKL aKeybdLayout)
// If pModifiersLR is non-NULL, place the modifiers that are needed to realize the key in there.
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
	{
		LPTSTR endptr;
		vk_type vk = (vk_type)_tcstol(aText + 2, &endptr, 16);  // Convert from hex.
		return *endptr ? 0 : vk; // Fixed for v1.1.27: Disallow any invalid suffix so that hotkeys like a::vkb() are not misinterpreted as remappings.
	}

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



vk_type CharToVKAndModifiers(TCHAR aChar, modLR_type *pModifiersLR, HKL aKeybdLayout, bool aEnableAZFallback)
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
	{
		if (  !(aEnableAZFallback && cisalpha(aChar))  )
			return 0;
		// v1.1.27.00: Use the A-Z fallback; assume the user means vk41-vk5A, since these letters
		// are commonly used to describe keyboard shortcuts even when these vk codes are actually
		// mapped to other characters.  Our callers should pass false for aEnableAZFallback if
		// they require a strict printable character->keycode mapping, such as for sending text.
		vk = (vk_type)ctoupper(aChar);
		keyscan_modifiers = cisupper(aChar) ? 0x01 : 0; // It's debatable whether the user intends this to be Shift+letter; this at least makes `Send ^A` consistent across (most?) layouts.
	}
	if (keyscan_modifiers & 0x38) // "The Hankaku key is pressed" or either of the "Reserved" state bits (for instance, used by Neo2 keyboard layout).
		// Callers expect failure in this case so that a fallback method can be used.
		return 0;

	// For v1.0.35, pModifiersLR was changed to modLR vs. mod so that AltGr keys such as backslash and
	// '{' are supported on layouts such as German when sending to apps such as Putty that are fussy about
	// which ALT key is held down to produce the character.  The following section detects AltGr by the
	// assuming that any character that requires both CTRL and ALT (with optional SHIFT) to be held
	// down is in fact an AltGr key (I don't think there are any that aren't AltGr in this case, but
	// confirmation would be nice).

	// The win docs for VkKeyScan() are a bit confusing, referring to flag "bits" when it should really
	// say flag "values".  In addition, it seems that these flag values are incompatible with
	// MOD_ALT, MOD_SHIFT, and MOD_CONTROL, so they must be translated:
	if (pModifiersLR) // The caller wants this info added to the output param.
	{
		// Best not to reset this value because some callers want to retain what was in it before,
		// merely merging these new values into it:
		//*pModifiers = 0;
		if ((keyscan_modifiers & 0x06) == 0x06) // 0x06 means "requires/includes AltGr".
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



bool TextToVKandSC(LPCTSTR aText, vk_type &aVK, sc_type &aSC, modLR_type *pModifiersLR, HKL aKeybdLayout)
{
	if (aVK = TextToVK(aText, pModifiersLR, true, true, aKeybdLayout))
	{
		aSC = 0; // Caller should call vk_to_sc(aVK) if needed.
		return true;
	}
	if (aSC = TextToSC(aText))
	{
		// Leave aVK set to 0.  Caller should call sc_to_vk(aSC) if needed.
		return true;
	}
	if (!_tcsnicmp(aText, _T("VK"), 2)) // Could be vkXXscXXX, which TextToVK() does not permit in v1.1.27+.
	{
		LPTSTR cp;
		vk_type vk = (vk_type)_tcstoul(aText + 2, &cp, 16);
		if (!_tcsnicmp(cp, _T("SC"), 2))
		{
			sc_type sc = (sc_type)_tcstoul(cp + 2, &cp, 16);
			if (!*cp) // No invalid suffix.
			{
				aVK = vk;
				aSC = sc;
				return true;
			}
		}
		else // Invalid suffix.  *cp!=0 is implied as vkXX would be handled by TextToVK().
			return false;
	}
	return false;
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
		if (!aSC && (aVK == VK_RETURN || !(aSC = vk_to_sc(aVK, true)))) // Prefer the non-Numpad name.
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
// If caller passes true for aReturnSecondary, the "extended" scan code will be returned for
// virtual keys that have two scan codes and two names (if there's only one, callers rely on
// zero being returned).  In those cases, the caller may want to know:
//  a) Whether the hook needs to be used to identify a hotkey defined by name.
//  b) Whether InputHook should handle the keys by SC in order to differentiate.
//  c) Whether to retrieve the key's name by SC rather than VK.
// In all of those cases, only keys that we've given multiple names matter.
// Custom layouts could assign some other VK to multiple SCs, but there would
// be no reason (or way) to differentiate them in this context.
{
	// Try to minimize the number mappings done manually because MapVirtualKey is a more reliable
	// way to get the mapping if user has non-standard or custom keyboard layout.

	sc_type sc = 0;

	switch (aVK)
	{
	// MapVirtualKey() returns 0xE11D, but we want the code normally received by the
	// hook (sc045).  See sc_to_vk() for more comments.
	case VK_PAUSE:    sc = SC_PAUSE; break;

	// PrintScreen: MapVirtualKey() returns 0x54, which is SysReq (produced by pressing
	// Alt+PrintScreen, but still maps to VK_SNAPSHOT).  Use sc137 for consistency with
	// what the hook reports for the naked keypress (and therefore what a hotkey is
	// likely to need).
	case VK_SNAPSHOT: sc = SC_PRINTSCREEN; break;

	// See comments in sc_to_vk().
	case VK_NUMLOCK:  sc = SC_NUMLOCK; break;
	}

	if (sc) // Above found a match.
		return aReturnSecondary ? 0 : sc; // Callers rely on zero being returned for VKs that don't have secondary SCs.

	if (   !(sc = MapVirtualKey(aVK, MAPVK_VK_TO_VSC_EX))   )
		return 0; // Indicate "no mapping".

	if (sc & 0xE000) // Prefix byte E0 or E1 (but E1 should only be possible for Pause/Break, which was already handled above).
		sc = 0x0100 | (sc & 0xFF);

	switch (aVK)
	{
	// The following virtual keys have more than one physical key, and thus more than one scan code.
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
		// This is likely to be incorrect for custom layouts where aVK is mapped to two SCs
		// that differ in the low byte.  There seems to be no simple way to fix that;
		// the complex ways would be:
		//  - Build our own conversion table by mapping all SCs to VKs (taking care to detect
		//    changes to the current keyboard layout).  Find the second SC that maps to aVK,
		//    or the first one with the 0xE000 flag.  However, there's no guarantee that it
		//    would correspond to NumpadEnter vs. Enter, or Insert vs. NumpadIns, for example.
		//  - Load the keyboard layout dll manually and search the SC-to-VK conversion tables.
		//    What we actually want is to differentiate Numpad keys from their non-Numpad
		//    counter-parts, and we can do that by checking for the KBDNUMPAD flag.
		// Custom layouts might cause these issues:
		//  - If the scan code of the secondary key is changed, the Hotkey control (and
		//    other sections that don't call this function) may return either "scXXX"
		//    or a name inconsistent with the key's current VK (but if it's a custom
		//    layout + standard keyboard, it should match the key's original function).
		//  - The Hotkey control assumes that the HOTKEYF_EXT flag corresponds to the
		//    secondary key, but either/both/neither could be extended on a custom layout.
		//    If it's both/neither, the control would give no way to distinguish.
		return aReturnSecondary ? (sc | 0x0100) : sc; // Below relies on the fact that these cases return early.

	// See "case SC_RSHIFT:" in sc_to_vk() for comments.
	case VK_RSHIFT:
		sc |= 0x0100;
		break;
	}

	// Since above didn't return, if aReturnSecondary==true, return 0 to indicate "no secondary SC for this VK".
	return aReturnSecondary ? 0 : sc; // Callers rely on zero being returned for VKs that don't have secondary SCs.
}



vk_type sc_to_vk(sc_type aSC)
{
	// aSC is actually a combination of the last byte of the keyboard make code combined with
	// 0x100 for the extended-key flag.  Although in most cases the flag corresponds to a prefix
	// byte of 0xE0, it seems it's actually set by the KBDEXT flag in the keyboard layout dll
	// (it's hard to find documentation).  A few keys have the KBDEXT flag inverted, which means
	// we can't tell reliably which scan codes really need the 0xE0 prefix, so just handle them
	// as special cases and hope that the flag never varies between layouts.
	// If this approach ever fails for custom layouts, some alternatives are:
	//  - Load the keyboard layout dll manually and check the scan code conversion tables for
	//    the presence of the KBDEXT flag.
	//  - Convert aSC and (aSC ^ 0x100), check the conversion of VK back to SC, and if it
	//    round-trips use that VK instead.
	// However, it seems that neither MSKLC nor KbdEdit provide a means to change the KBDEXT flag.
	// US layout: https://github.com/microsoft/Windows-driver-samples/blob/master/input/layout/kbdus/kbdus.c
	// Keyboard make codes: http://stanislavs.org/helppc/make_codes.html
	// More low-level keyboard details: https://www.win.tue.nl/~aeb/linux/kbd/scancodes-1.html#ss1.5
	switch (aSC)
	{
	// RShift doesn't have the 0xE0 prefix but has KBDEXT.  The US layout sample says
	// "Right-hand Shift key must have KBDEXT bit set", so it's probably always set.
	// KbdEdit seems to follow this rule when VK_RSHIFT is assigned to a non-ext key.
	// It's definitely possible to assign RShift a different VK, but 1) it can't be
	// done with MSKLC, and 2) KbdEdit clears the ext flag (so aSC != SC_RSHIFT).
	case SC_RSHIFT:
	// NumLock doesn't have the 0xE0 prefix but has KBDEXT.  Actually pressing the key
	// will produce VK_PAUSE if CTRL is down, but with SC_NUMLOCK rather than SC_PAUSE.
	case SC_NUMLOCK:
		// These cases can be handled by adjusting aSC to reflect the fact that these
		// keys don't really have the 0xE0 prefix, and allowing MapVirtualKey() to be
		// called below in case they have been remapped.
		aSC &= 0xFF;
		break;

	// Pause actually generates 0xE1,0x1D,0x45, or in other words, E1,LCtrl,NumLock.
	// kbd.h says "We must convert the E1+LCtrl to BREAK, then ignore the Numlock".
	// So 0xE11D maps to and from VK_PAUSE, and 0x45 is "ignored".  However, the hook
	// receives only 0x45, not 0xE11D (which I guess would be truncated to 0x1D/ctrl).
	// The documentation for KbdEdit also indicates the mapping of Pause is "hard-wired":
	// http://www.kbdedit.com/manual/low_level_edit_vk_mappings.html
	case SC_PAUSE:
		return VK_PAUSE;
	}

	if (aSC & 0x100) // Our extended-key flag.
	{
		// Since it wasn't handled above, assume the extended-key flag corresponds to the 0xE0
		// prefix byte.  Passing 0xE000 should work on Vista and up, though it appears to be
		// documented only for MapVirtualKeyEx() as at 2019-10-26.  Details can be found in
		// archives of Michael Kaplan's blog (the original blog has been taken down):
		// https://web.archive.org/web/20070219075710/http://blogs.msdn.com/michkap/archive/2006/08/29/729476.aspx
		aSC = 0xE000 | (aSC & 0xFF);
	}
	return MapVirtualKey(aSC, MAPVK_VSC_TO_VK_EX);
}



#pragma region Top-level Functions

bif_impl void Send(StrArg aKeys)
{
	SendKeys(aKeys, SCM_NOT_RAW, g->SendMode);
}

bif_impl void SendText(StrArg aText)
{
	SendKeys(aText, SCM_RAW_TEXT, g->SendMode);
}

bif_impl void SendInput(StrArg aKeys)
{
	SendKeys(aKeys, SCM_NOT_RAW, g->SendMode == SM_INPUT_FALLBACK_TO_PLAY ? SM_INPUT_FALLBACK_TO_PLAY : SM_INPUT);
}

bif_impl void SendPlay(StrArg aKeys)
{
	SendKeys(aKeys, SCM_NOT_RAW, SM_PLAY);
}

bif_impl void SendEvent(StrArg aKeys)
{
	SendKeys(aKeys, SCM_NOT_RAW, SM_EVENT);
}

bif_impl FResult SetNumLockState(optl<StrArg> aState)
{
	return SetToggleState(VK_NUMLOCK, g_ForceNumLock, aState);
}

bif_impl FResult SetCapsLockState(optl<StrArg> aState)
{
	return SetToggleState(VK_CAPITAL, g_ForceCapsLock, aState);
}

bif_impl FResult SetScrollLockState(optl<StrArg> aState)
{
	return SetToggleState(VK_SCROLL, g_ForceScrollLock, aState);
}

bif_impl FResult MouseClick(optl<StrArg> aButton, optl<int> aX, optl<int> aY, optl<int> aClickCount, optl<int> aSpeed, optl<StrArg> aDownOrUp, optl<StrArg> aRelative)
{
	return PerformMouse(ACT_MOUSECLICK, aButton, aX, aY, nullptr, nullptr, aSpeed, aRelative, aClickCount, aDownOrUp);
}

bif_impl FResult MouseClickDrag(optl<StrArg> aButton, int aX1, int aY1, int aX2, int aY2, optl<int> aSpeed, optl<StrArg> aRelative)
{
	return PerformMouse(ACT_MOUSECLICKDRAG, aButton, aX1, aY1, aX2, aY2, aSpeed, aRelative, nullptr, nullptr);
}

bif_impl FResult MouseMove(int aX, int aY, optl<int> aSpeed, optl<StrArg> aRelative)
{
	return PerformMouse(ACT_MOUSEMOVE, nullptr, aX, aY, nullptr, nullptr, aSpeed, aRelative, nullptr, nullptr);
}

#pragma endregion



#pragma region BlockInput

void OurBlockInput(bool aEnable)
{
	// Always turn input ON/OFF even if g_BlockInput says its already in the right state.  This is because
	// BlockInput can be externally and undetectably disabled, e.g. if the user presses Ctrl-Alt-Del:
	BlockInput(aEnable ? TRUE : FALSE);
	g_BlockInput = aEnable;
}

bif_impl FResult ScriptBlockInput(StrArg aMode)
{
	switch (auto toggle = Line::ConvertBlockInput(aMode))
	{
	case TOGGLED_ON:
		OurBlockInput(true);
		break;
	case TOGGLED_OFF:
		OurBlockInput(false);
		break;
	case TOGGLE_SEND:
	case TOGGLE_MOUSE:
	case TOGGLE_SENDANDMOUSE:
	case TOGGLE_DEFAULT:
		g_BlockInputMode = toggle;
		break;
	case TOGGLE_MOUSEMOVE:
		g_BlockMouseMove = true;
		Hotkey::InstallMouseHook();
		break;
	case TOGGLE_MOUSEMOVEOFF:
		g_BlockMouseMove = false; // But the mouse hook is left installed because it might be needed by other things. This approach is similar to that used by the Input command.
		break;
	default: // NEUTRAL or TOGGLE_INVALID
		return FR_E_ARG(0);
	}
	return OK;
}

#pragma endregion
