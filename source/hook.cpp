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
#include "hook.h"
#include "globaldata.h"  // for access to several global vars
#include "util.h" // for snprintfcat()
#include "window.h" // for MsgBox()
#include "application.h" // For MsgSleep().

// Declare static variables (global to only this file/module, i.e. no external linkage):
static HANDLE sKeybdMutex = NULL;
static HANDLE sMouseMutex = NULL;
#define KEYBD_MUTEX_NAME "AHK Keybd"
#define MOUSE_MUTEX_NAME "AHK Mouse"

// Whether to disguise the next up-event for lwin/rwin to suppress Start Menu.
// These are made global, rather than static inside the hook function, so that
// we can ensure they are initialized by the keyboard init function every
// time it's called (currently it can be only called once):
static bool sDisguiseNextLWinUp;        // Initialized by ResetHook().
static bool sDisguiseNextRWinUp;        //
static bool sDisguiseNextLAltUp;        //
static bool sDisguiseNextRAltUp;        //
static bool sAltTabMenuIsVisible;       //
static vk_type sVKtoIgnoreNextTimeDown; //

// The prefix key that's currently down (i.e. in effect).
// It's tracked this way, rather than as a count of the number of prefixes currently down, out of
// concern that such a count might accidentally wind up above zero (due to a key-up being missed somehow)
// and never come back down, thus penalizing performance until the program is restarted:
static key_type *pPrefixKey;  // Initialized by ResetHook().

// Less memory overhead (space and performance) to allocate a solid block for multidimensional arrays:
// These store all the valid modifier+suffix combinations (those that result in hotkey actions) except
// those with a ModifierVK/SC.  Doing it this way should cut the CPU overhead caused by having many
// hotkeys handled by the hook, down to a fraction of what it would be otherwise.  Indeed, doing it
// this way makes the performance impact of adding many additional hotkeys of this type exactly zero
// once the program has started up and initialized.  The main alternative is a binary search on an
// array of keyboard-hook hotkeys (similar to how the mouse is done):
static HotkeyIDType *kvkm = NULL;
static HotkeyIDType *kscm = NULL;
static HotkeyIDType *hotkey_up = NULL;
static key_type *kvk = NULL;
static key_type *ksc = NULL;
// Macros for convenience in accessing the above arrays as multidimensional objects.
// When using them, be sure to consistently access the first index as ModLR (i.e. the rows)
// and the second as VK or SC (i.e. the columns):
#define Kvkm(i,j) kvkm[(i)*(MODLR_MAX + 1) + (j)]
#define Kscm(i,j) kscm[(i)*(MODLR_MAX + 1) + (j)]
#define KVKM_SIZE ((MODLR_MAX + 1)*(VK_ARRAY_COUNT))
#define KSCM_SIZE ((MODLR_MAX + 1)*(SC_ARRAY_COUNT))

// Notes about the following variables:
// Used to help make a workaround for the way the keyboard driver generates physical
// shift-key events to release the shift key whenever it is physically down during
// the press or release of a dual-state Numpad key. These keyboard driver generated
// shift-key events only seem to happen when Numlock is ON, the shift key is logically
// or physically down, and a dual-state numpad key is pressed or released (i.e. the shift
// key might not have been down for the press, but if it's down for the release, the driver
// will suddenly start generating shift events).  I think the purpose of these events is to
// allow the shift keyto temporarily alter the state of the Numlock key for the purpose of
// sending that one key, without the shift key actually being "seen" as down while the key
// itself is sent (since some apps may have special behavior when they detect the shift key
// is down).

// Note: numlock, numpaddiv/mult/sub/add/enter are not affected by this because they have only
// a single state (i.e. they are unaffected by the state of the Numlock key).  Also, these
// driver-generated events occur at a level lower than the hook, so it doesn't matter whether
// the hook suppresses the keys involved (i.e. the shift events still happen anyway).

// So which keys are not physical even though they're non-injected?:
// 1) The shift-up that precedes a down of a dual-state numpad key (only happens when shift key is logically down).
// 2) The shift-down that precedes a pressing down (or releasing in certain very rare cases caused by the
//    exact sequence of keys received) of a key WHILE the numpad key in question is still down.
//    Although this case may seem rare, it's happened to both Robert Yaklin and myself when doing various
//    sorts of hotkeys.
// 3) The shift-up that precedes an up of a dual-state numpad key.  This only happens if the shift key is
//    logically down for any reason at this exact moment, which can be achieved via the send command.
// 4) The shift-down that follows the up of a dual-state numpad key (i.e. the driver is restoring the shift state
//    to what it was before).  This can be either immediate or "lazy".  It's lazy whenever the user had pressed
//    another key while a numpad key was being held down (i.e. case #2 above), in which case the driver waits
//    indefinitely for the user to press any other key and then immediately sneaks in the shift key-down event
//    right before it in the input stream (insertion).
// 5) Similar to #4, but if the driver needs to generate a shift-up for an unexpected Numpad-up event,
//    the restoration of the shift key will be "lazy".  This case was added in response to the below
//    example, wherein the shift key got stuck physically down (incorrectly) by the hook:
// 68  048	 	d	0.00	Num 8          	
// 6B  04E	 	d	0.09	Num +          	
// 68  048	i	d	0.00	Num 8          	
// 68  048	i	u	0.00	Num 8          	
// A0  02A	i	d	0.02	Shift          	part of the macro
// 01  000	i	d	0.03	LButton        	
// A0  02A	 	u	0.00	Shift          	driver, for the next key
// 26  048	 	u	0.00	Num 8          	
// A0  02A	 	d	0.49	Shift          	driver lazy down (but not detected as non-physical)
// 6B  04E	 	d	0.00	Num +          	

// The below timeout is for the subset of driver-generated shift-events that occur immediately
// before or after some other keyboard event.  The elapsed time is usually zero, but using 22ms
// here just in case slower systems or systems under load have longer delays between keystrokes:
#define SHIFT_KEY_WORKAROUND_TIMEOUT 22
static bool sNextPhysShiftDownIsNotPhys; // All of these are initialized by ResetHook().
static vk_type sPriorVK;
static sc_type sPriorSC;
static bool sPriorEventWasKeyUp;
static bool sPriorEventWasPhysical;
static DWORD sPriorEventTickCount;
static modLR_type sPriorModifiersLR_physical;
static BYTE sPriorShiftState;
static BYTE sPriorLShiftState;

enum DualNumpadKeys	{PAD_DECIMAL, PAD_NUMPAD0, PAD_NUMPAD1, PAD_NUMPAD2, PAD_NUMPAD3
, PAD_NUMPAD4, PAD_NUMPAD5, PAD_NUMPAD6, PAD_NUMPAD7, PAD_NUMPAD8, PAD_NUMPAD9
, PAD_DELETE, PAD_INSERT, PAD_END, PAD_DOWN, PAD_NEXT, PAD_LEFT, PAD_CLEAR
, PAD_RIGHT, PAD_HOME, PAD_UP, PAD_PRIOR, PAD_TOTAL_COUNT};
static bool sPadState[PAD_TOTAL_COUNT];  // Initialized by ChangeHookState()

/////////////////////////////////////////////////////////////////////////////////////////////

/*
One of the main objectives of a the hooks is to minimize the amount of CPU overhead caused by every
input event being handled by the procedure.  One way this is done is to return immediately on
simple conditions that are relatively frequent (such as receiving a key that's not involved in any
hotkey combination).

Another way is to avoid API or system calls that might have a high overhead.  That's why the state of
every prefix key is tracked independently, rather than calling the WinAPI to determine if the
key is actually down at the moment of consideration.
*/

inline bool IsIgnored(ULONG_PTR aExtraInfo)
// KEY_PHYS_IGNORE events must be mostly ignored because currently there is no way for a given
// hook instance to detect if it sent the event or some other instance.  Therefore, to treat
// such events as true physical events might cause infinite loops or other side-effects in
// the instance that generated the event.  More review of this is needed if KEY_PHYS_IGNORE
// events ever need to be treated as true physical events by the instances of the hook that
// didn't originate them:
{
	return aExtraInfo == KEY_IGNORE || aExtraInfo == KEY_PHYS_IGNORE || aExtraInfo == KEY_IGNORE_ALL_EXCEPT_MODIFIER;
}



LRESULT CALLBACK LowLevelKeybdProc(int aCode, WPARAM wParam, LPARAM lParam)
{
	if (aCode != HC_ACTION)  // MSDN docs specify that both LL keybd & mouse hook should return in this case.
		return CallNextHookEx(g_KeybdHook, aCode, wParam, lParam);

	KBDLLHOOKSTRUCT &event = *(PKBDLLHOOKSTRUCT)lParam;  // For convenience, maintainability, and possibly performance.

	// Change the event to be physical if that is indicated in its dwExtraInfo attribute.
	// This is done for cases when the hook is installed multiple times and one instance of
	// it wants to inform the others that this event should be considered physical for the
	// purpose of updating modifier and key states:
	if (event.dwExtraInfo == KEY_PHYS_IGNORE)
		event.flags &= ~LLKHF_INJECTED;

	// Make all keybd events physical to try to fool the system into accepting CTRL-ALT-DELETE.
	// This didn't work, which implies that Ctrl-Alt-Delete is trapped at a lower level than
	// this hook (folks have said that it's trapped in the keyboard driver itself):
	//event.flags &= ~LLKHF_INJECTED;

	// Note: Some scan codes are shared by more than one key (e.g. Numpad7 and NumpadHome).  This is why
	// the keyboard hook must be able to handle hotkeys by either their virtual key or their scan code.
	// i.e. if sc were always used in preference to vk, we wouldn't be able to distinguish between such keys.

	bool key_up = (wParam == WM_KEYUP || wParam == WM_SYSKEYUP);
	vk_type vk = (vk_type)event.vkCode;
	sc_type sc = (sc_type)event.scanCode;
	if (vk && !sc) // Might happen if another app calls keybd_event with a zero scan code.
		sc = vk_to_sc(vk);
	// MapVirtualKey() does *not* include 0xE0 in HIBYTE if key is extended.  In case it ever
	// does in the future (or if event.scanCode ever does), force sc to be an 8-bit value
	// so that it's guaranteed consistent and to ensure it won't exceed SC_MAX (which might cause
	// array indexes to be out-of-bounds).  The 9th bit is later set to 1 if the key is extended:
	sc &= 0xFF;
	// Change sc to be extended if indicated.  But avoid doing so for VK_RSHIFT, which is
	// apparently considered extended by the API when it shouldn't be.  Update: Well, it looks like
	// VK_RSHIFT really is an extended key, at least on WinXP (and probably be extension on the other
	// NT based OSes as well).  What little info I could find on the 'net about this is contradictory,
	// but it's clear that some things just don't work right if the non-extended scan code is sent.  For
	// example, the shift key will appear to get stuck down in the foreground app if the non-extended
	// scan code is sent with VK_RSHIFT key-up event:
	if ((event.flags & LLKHF_EXTENDED)) // && vk != VK_RSHIFT)
		sc |= 0x100;

	// The below must be done prior to any returns that indirectly call UpdateKeybdState() to update
	// modifier state.
	// Update: It seems best to do the below unconditionally, even if the OS is Win2k or WinXP,
	// since it seems like this translation will add value even in those cases:
	// To help ensure consistency with Windows XP and 2k, for which this hook has been primarily
	// designed and tested, translate neutral modifier keys into their left/right specific VKs,
	// since beardboy's testing shows that NT4 receives the neutral keys like Win9x does:
	switch (vk)
	{
	case VK_SHIFT:   vk = (sc == SC_RSHIFT)   ? VK_RSHIFT   : VK_LSHIFT; break;
	case VK_CONTROL: vk = (sc == SC_RCONTROL) ? VK_RCONTROL : VK_LCONTROL; break;
	case VK_MENU:    vk = (sc == SC_RALT)     ? VK_RMENU    : VK_LMENU; break;
	}

	// Now that the above has translated VK_CONTROL to VK_LCONTROL (if necessary):
	if (vk == VK_LCONTROL)
	{
		// The following helps hasten AltGr detection after script startup.  It's kept to supplement
		// LayoutHasAltGr() because that function isn't 100% reliable for the reasons described there.
		// It shouldn't be necessary to check what type of LControl event (up or down) is received, since
		// it should be impossible for any unrelated keystrokes to be received while g_HookReceiptOfLControlMeansAltGr
		// is true.  This is because all unrelated keystrokes stay buffered until the main thread calls GetMessage().
		// UPDATE for v1.0.39: Now that the hook has a dedicated thread, the above is no longer 100% certain.
		// However, I think confidence is very high that this AltGr detection method is okay as it is because:
		// 1) Hook thread has high priority, which means it generally shouldn't get backlogged with buffered keystrokes.
		// 2) When the main thread calls keybd_event(), there is strong evidence that the OS immediately preempts
		//    the main thread (or executes a SendMessage(), which is the same as preemption) and gives the next
		//    timeslice to the hook thread, which then immediately processes the incoming keystroke as long
		//    as there are no keystrokes in the queue ahead of it (see #1).
		// 3) Even if #1 and #2 fail to hold, the probability of misclassifying an LControl event seems very low.
		//    If there is ever concern, adding a call to IsIgnored() below would help weed out physical keystrokes
		//    (or those of other programs) that arrive during a vulnerable time.
		if (g_HookReceiptOfLControlMeansAltGr)
		{
			// But don't reset g_HookReceiptOfLControlMeansAltGr here to avoid timing problems where the hook
			// is installed at a time when g_HookReceiptOfLControlMeansAltGr is wrongly true because the
			// inactive hook never made it false.  Let KeyEvent() do that.
			Get_active_window_keybd_layout // Defines the variable active_window_keybd_layout for use below.
			LayoutHasAltGr(active_window_keybd_layout, CONDITION_TRUE);
			// The following must be done; otherwise, if LayoutHasAltGr hasn't yet been autodetected by the
			// time the first AltGr keystroke comes through, that keystroke would cause LControl to get stuck down
			// as seen in g_modifiersLR_physical.
			event.flags |= LLKHF_INJECTED; // Flag it as artificial for any other instances of the hook that may be running.
			event.dwExtraInfo = g_HookReceiptOfLControlMeansAltGr; // The one who set this variable put the desired ExtraInfo in it.
		}
		else // Since AltGr wasn't detected above, see if any other means is ready to detect it.
		{
			// v1.0.42.04: This section was moved out of IsIgnored() to here because:
			// 1) Immediately correcting the incoming event simplifies other parts of the hook.
			// 2) It allows this instance of the hook to communicate with other instances of the hook by
			//    correcting the bogus values directly inside the event structure.  This is something those
			//    other hooks can't do easily if the keystrokes were generated/simulated inside our instance
			//    (due to our instance's KeyEvent() function communicating corrections via
			//    g_HookReceiptOfLControlMeansAltGr and g_IgnoreNextLControlDown/Up).
			//
			// This new approach solves an AltGr keystroke's disruption of A_TimeIdlePhysical and related
			// problems that were caused by AltGr's incoming keystroke being marked by the driver or OS as a
			// physical event (even when the AltGr keystroke that caused it was artificial).  It might not
			// be a perfect solution, but it's pretty complete. For example, with the exception of artificial
			// AltGr keystrokes from non-AHK sources, it completely solves the A_TimeIdlePhysical issue because
			// by definition, any script that *uses* A_TimeIdlePhysical in a way that the fix applies to also
			// has the keyboard hook installed (if it only has the mouse hook installed, the fix isn't in effect,
			// but it also doesn't matter because that script detects only mouse events as true physical events,
			// as described in the documentation for A_TimeIdlePhysical).
			if (key_up)
			{
				if (g_IgnoreNextLControlUp)
				{
					event.flags |= LLKHF_INJECTED; // Flag it as artificial for any other instances of the hook that may be running.
					event.dwExtraInfo = g_IgnoreNextLControlUp; // The one who set this variable put the desired ExtraInfo in here.
				}
			}
			else // key-down event
			{
				if (g_IgnoreNextLControlDown)
				{
					event.flags |= LLKHF_INJECTED; // Flag it as artificial for any other instances of the hook that may be running.
					event.dwExtraInfo = g_IgnoreNextLControlDown; // The one who set this variable put the desired ExtraInfo in here.
				}
			}
		}
	} // if (vk == VK_LCONTROL)

	return LowLevelCommon(g_KeybdHook, aCode, wParam, lParam, vk, sc, key_up, event.dwExtraInfo, event.flags);
}



LRESULT CALLBACK LowLevelMouseProc(int aCode, WPARAM wParam, LPARAM lParam)
{
	// code != HC_ACTION should be evaluated PRIOR to considering the values
	// of wParam and lParam, because those values may be invalid or untrustworthy
	// whenever code < 0.
	if (aCode != HC_ACTION)
		return CallNextHookEx(g_MouseHook, aCode, wParam, lParam);

	MSLLHOOKSTRUCT &event = *(PMSLLHOOKSTRUCT)lParam;  // For convenience, maintainability, and possibly performance.

	// Make all mouse events physical to try to simulate mouse clicks in games that normally ignore
	// artificial input.
	//event.flags &= ~LLMHF_INJECTED;

	if (!(event.flags & LLMHF_INJECTED)) // Physical mouse movement or button action (uses LLMHF vs. LLKHF).
		g_TimeLastInputPhysical = GetTickCount();
		// Above: Don't use event.time, mostly because SendInput can produce invalid timestamps on such events
		// (though in truth, that concern isn't valid because SendInput's input isn't marked as physical).
		// Another concern is the comments at the other update of "g_TimeLastInputPhysical" elsewhere in this file.
		// A final concern is that some drivers might be faulty and might not generate an accurate timestamp.

	if (wParam == WM_MOUSEMOVE) // Only after updating for physical input, above, is this checked.
		return (g_BlockMouseMove && !(event.flags & LLMHF_INJECTED)) ? 1 : CallNextHookEx(g_MouseHook, aCode, wParam, lParam);
		// Above: In v1.0.43.11, a new mode was added to block mouse movement only since it's more flexible than
		// BlockInput (which keybd too, and blocks all mouse buttons too).  However, this mode blocks only
		// physical mouse movement because it seems most flexible (and simplest) to allow all artificial
		// movement, even if that movement came from a source other than an AHK script (such as some other
		// macro program).

	// MSDN: WM_LBUTTONDOWN, WM_LBUTTONUP, WM_MOUSEMOVE, WM_MOUSEWHEEL [, WM_MOUSEHWHEEL], WM_RBUTTONDOWN, or WM_RBUTTONUP.
	// But what about the middle button?  It's undocumented, but it is received.
	// What about doubleclicks (e.g. WM_LBUTTONDBLCLK): I checked: They are NOT received.
	// This is expected because each click in a doubleclick could be separately suppressed by
	// the hook, which would make it become a non-doubleclick.
	vk_type vk = 0;
	sc_type sc = 0; // To be overriden if this even is a wheel turn.
	short wheel_delta;
	bool key_up = true;  // Set default to safest value.

	switch (wParam)
	{
		case WM_MOUSEWHEEL:
		case WM_MOUSEHWHEEL: // v1.0.48: Lexikos: Support horizontal scrolling in Windows Vista and later.
			// MSDN: "A positive value indicates that the wheel was rotated forward, away from the user;
			// a negative value indicates that the wheel was rotated backward, toward the user. One wheel
			// click is defined as WHEEL_DELTA, which is 120."  Testing shows that on XP at least, the
			// abs(delta) is greater than 120 when the user turns the wheel quickly (also depends on
			// granularity of wheel hardware); i.e. the system combines multiple turns into a single event.
			wheel_delta = GET_WHEEL_DELTA_WPARAM(event.mouseData); // Must typecast to short (not int) via macro, otherwise the conversion to negative/positive number won't be correct.
			if (wParam == WM_MOUSEWHEEL)
				vk = wheel_delta < 0 ? VK_WHEEL_DOWN : VK_WHEEL_UP;
			else
				vk = wheel_delta < 0 ? VK_WHEEL_LEFT : VK_WHEEL_RIGHT;
			// Dividing by WHEEL_DELTA was a mistake because some mice can yield detas less than 120.
			// However, this behavior is kept for backward compatibility because some scripts may rely
			// on A_EventInfo==0 meaning "delta is between 1 and 119".  WheelLeft/Right were also done
			// that way because consistency may be more important than correctness.  In the future, perhaps
			// an A_EventInfo2 can be added, or some hotkey aliases like "FineWheelXXX".
			sc = (wheel_delta > 0 ? wheel_delta : -wheel_delta) / WHEEL_DELTA; // See above. Note that sc is unsigned.
			key_up = false; // Always consider wheel movements to be "key down" events.
			break;
		case WM_LBUTTONUP: vk = VK_LBUTTON;	break;
		case WM_RBUTTONUP: vk = VK_RBUTTON; break;
		case WM_MBUTTONUP: vk = VK_MBUTTON; break;
		case WM_NCXBUTTONUP:  // NC means non-client.
		case WM_XBUTTONUP: vk = (HIWORD(event.mouseData) == XBUTTON1) ? VK_XBUTTON1 : VK_XBUTTON2; break;
		case WM_LBUTTONDOWN: vk = VK_LBUTTON; key_up = false; break;
		case WM_RBUTTONDOWN: vk = VK_RBUTTON; key_up = false; break;
		case WM_MBUTTONDOWN: vk = VK_MBUTTON; key_up = false; break;
		case WM_NCXBUTTONDOWN:
		case WM_XBUTTONDOWN: vk = (HIWORD(event.mouseData) == XBUTTON1) ? VK_XBUTTON1 : VK_XBUTTON2; key_up = false; break;
	}

	return LowLevelCommon(g_MouseHook, aCode, wParam, lParam, vk, sc, key_up, event.dwExtraInfo, event.flags);
}



LRESULT LowLevelCommon(const HHOOK aHook, int aCode, WPARAM wParam, LPARAM lParam, const vk_type aVK
	, sc_type aSC, bool aKeyUp, ULONG_PTR aExtraInfo, DWORD aEventFlags)
// v1.0.38.06: The keyboard and mouse hooks now call this common function to reduce code size and improve
// maintainability.  The code size savings as of v1.0.38.06 is 3.5 KB of uncompressed code, but that
// savings will grow larger if more complexity is ever added to the hooks.
{
	HotkeyIDType hotkey_id_to_post = HOTKEY_ID_INVALID; // Set default.
	bool is_ignored = IsIgnored(aExtraInfo);

	// The following is done for more than just convenience.  It solves problems that would otherwise arise
	// due to the value of a global var such as KeyHistoryNext changing due to the reentrancy of
	// this procedure.  For example, a call to KeyEvent() in here would alter the value of
	// KeyHistoryNext, in most cases before we had a chance to finish using the old value.  In other
	// words, we use an automatic variable so that every instance of this function will get its
	// own copy of the variable whose value will stays constant until that instance returns:
	KeyHistoryItem *pKeyHistoryCurr, khi_temp; // Must not be static (see above).  Serves as a storage spot for a single keystroke in case key history is disabled.
	if (!g_KeyHistory)
		pKeyHistoryCurr = &khi_temp;  // Having a non-NULL pKeyHistoryCurr simplifies the code in other places.
	else
	{
		pKeyHistoryCurr = g_KeyHistory + g_KeyHistoryNext;
		if (++g_KeyHistoryNext >= g_MaxHistoryKeys)
			g_KeyHistoryNext = 0;
		pKeyHistoryCurr->vk = aVK; // aSC is done later below.
		pKeyHistoryCurr->key_up = aKeyUp;
		g_HistoryTickNow = GetTickCount();
		pKeyHistoryCurr->elapsed_time = (g_HistoryTickNow - g_HistoryTickPrev) / (float)1000;
		g_HistoryTickPrev = g_HistoryTickNow;
		HWND fore_win = GetForegroundWindow();
		if (fore_win)
		{
			if (fore_win != g_HistoryHwndPrev)
			{
				// The following line is commented out in favor of the one beneath it (seem below comment):
				//GetWindowText(fore_win, pKeyHistoryCurr->target_window, sizeof(pKeyHistoryCurr->target_window));
				PostMessage(g_hWnd, AHK_GETWINDOWTEXT, (WPARAM)pKeyHistoryCurr->target_window, (LPARAM)fore_win);
				// v1.0.44.12: The reason for the above is that clicking a window's close or minimize button
				// (and possibly other types of title bar clicks) causes a delay for the following window, at least
				// when XP Theme (but not classic theme) is in effect:
				//#InstallMouseHook
				//Gui, +AlwaysOnTop
				//Gui, Show, w200 h100
				//return
				// The problem came about from the following sequence of events:
				// 1) User clicks the one of the script's window's title bar's close, minimize, or maximize button.
				// 2) WM_NCLBUTTONDOWN is sent to the window's window proc, which then passes it on to
				//    DefWindowProc or DefDlgProc, which then apparently enters a loop in which no messages
				//    (or a very limited subset) are pumped.
				// 3) If anyone sends a message to that window (such as GetWindowText(), which sends a message
				//    in cases where it doesn't have the title pre-cached), the message will not receive a reply
				//    until after the mouse button is released.
				// 4) But the hook is the very thing that's supposed to release the mouse button, and it can't
				//    until a reply is received.
				// 5) Thus, a deadlock occurs.  So after a short but noticeable delay, the OS sees the hook as
				//    unresponsive and bypasses it, sending the click through normally, which breaks the deadlock.
				// 6) A similar situation might arise when a right-click-down is sent to the title bar or
				//    sys-menu-icon.
				//
				// SOLUTION:
				// Post the message to our main thread to have it do the GetWindowText call.  That way, if
				// the target window is one of the main thread's own window's, there's no chance it can be
				// in an unreponsive state like the deadlock described above.  In addition, do this for ALL
				// windows because its simpler, more maintainable, and especially might sovle other hook
				// performance problems if GetWindowText() has other situations where it is slow to return
				// (which seems likely).
				// Although the above solution could create rare situations where there's a lag before window text
				// is updated, that seems unlikely to be common or have signficant consequences.  Furthermore,
				// it has the advantage of improving hook performance by avoiding the call to GetWindowText (which
				// incidentally might solve hotkey lag problems that have been observed while the active window
				// is momentarily busy/unresponsive -- but maybe not because the main thread would then be lagged
				// instead of the hook thread, which is effectively the same result from user's POV).
				// Note: It seems best not to post the message to the hook thread because if LButton is down,
				// the hook's main event loop would be sending a message to an unresponsive thread (our main thread),
				// which would create the same deadlock.
				// ALTERNATE SOLUTIONS:
				// - #1: Avoid calling GetWindowText at all when LButton or RButton is in a logically-down state.
				// - Same as #1, but do so only if one of the main thread's target windows is known to be in a tight loop (might be too unreliable to detect all such cases).
				// - Same as #1 but less rigorous and more catch-all, such as by checking if the active window belongs to our thread.
				// - Avoid calling GetWindowText at all upon release of LButton.
				// - Same, but only if the window to have text retrieved belongs to our process.
				// - Same, but only if the mouse is inside the close/minimize/etc. buttons of the active window.
			}
			else // i.e. where possible, avoid the overhead of the call to GetWindowText().
				*pKeyHistoryCurr->target_window = '\0';
		}
		else
			strcpy(pKeyHistoryCurr->target_window, "N/A"); // Due to AHK_GETWINDOWTEXT, this could collide with main thread's writing to same string; but in addition to being extremely rare, it would likely be inconsequential.
		g_HistoryHwndPrev = fore_win;  // Updated unconditionally in case fore_win is NULL.
	}
	// Keep the following flush with the above to indicate that they're related.
	// The following is done even if key history is disabled because firing a wheel hotkey via PostMessage gets
	// the notch count from pKeyHistoryCurr->sc.
	pKeyHistoryCurr->sc = aSC; // Will be zero if our caller is the mouse hook (except for wheel notch count).
	// After logging the wheel notch count (above), purify aSC for readability and maintainability.
	if (IS_WHEEL_VK(aVK)) // Lexikos: Added checks for VK_WHEEL_LEFT and VK_WHEEL_RIGHT to support horizontal scrolling on Vista.
		aSC = 0; // Also relied upon by by sc_takes_precedence below.

	bool is_artificial;
	if (aHook == g_MouseHook)
	{
		if (   !(is_artificial = (aEventFlags & LLMHF_INJECTED))   ) // It's a physical mouse event.
			g_PhysicalKeyState[aVK] = aKeyUp ? 0 : STATE_DOWN;
	}
	else // Keybd hook.
	{
		// Even if the below is set to false, the event might be reclassified as artificial later (though it
		// won't be logged as such).  See comments in KeybdEventIsPhysical() for details.
		is_artificial = aEventFlags & LLKHF_INJECTED; // LLKHF vs. LLMHF
		// If the scan code is extended, the key that was pressed is not a dual-state numpad key,
		// i.e. it could be the counterpart key, such as End vs. NumpadEnd, located elsewhere on
		// the keyboard, but we're not interested in those.  Also, Numlock must be ON because
		// otherwise the driver will not generate those false-physical shift key events:
		if (!(aSC & 0x100) && (IsKeyToggledOn(VK_NUMLOCK)))
		{
			switch (aVK)
			{
			case VK_DELETE: case VK_DECIMAL: sPadState[PAD_DECIMAL] = !aKeyUp; break;
			case VK_INSERT: case VK_NUMPAD0: sPadState[PAD_NUMPAD0] = !aKeyUp; break;
			case VK_END:    case VK_NUMPAD1: sPadState[PAD_NUMPAD1] = !aKeyUp; break;
			case VK_DOWN:   case VK_NUMPAD2: sPadState[PAD_NUMPAD2] = !aKeyUp; break;
			case VK_NEXT:   case VK_NUMPAD3: sPadState[PAD_NUMPAD3] = !aKeyUp; break;
			case VK_LEFT:   case VK_NUMPAD4: sPadState[PAD_NUMPAD4] = !aKeyUp; break;
			case VK_CLEAR:  case VK_NUMPAD5: sPadState[PAD_NUMPAD5] = !aKeyUp; break;
			case VK_RIGHT:  case VK_NUMPAD6: sPadState[PAD_NUMPAD6] = !aKeyUp; break;
			case VK_HOME:   case VK_NUMPAD7: sPadState[PAD_NUMPAD7] = !aKeyUp; break;
			case VK_UP:     case VK_NUMPAD8: sPadState[PAD_NUMPAD8] = !aKeyUp; break;
			case VK_PRIOR:  case VK_NUMPAD9: sPadState[PAD_NUMPAD9] = !aKeyUp; break;
			}
		}

		// Track physical state of keyboard & mouse buttons. Also, if it's a modifier, let another section
		// handle it because it's not as simple as just setting the value to true or false (e.g. if LShift
		// goes up, the state of VK_SHIFT should stay down if VK_RSHIFT is down, or up otherwise).
		// Also, even if this input event will wind up being suppressed (usually because of being
		// a hotkey), still update the physical state anyway, because we want the physical state to
		// be entirely independent of the logical state (i.e. we want the key to be reported as
		// physically down even if it isn't logically down):
		if (!kvk[aVK].as_modifiersLR && KeybdEventIsPhysical(aEventFlags, aVK, aKeyUp))
			g_PhysicalKeyState[aVK] = aKeyUp ? 0 : STATE_DOWN;

		// Pointer to the key record for the current key event.  Establishes this_key as an alias
		// for the array element in kvk or ksc that corresponds to the vk or sc, respectively.
		// I think the compiler can optimize the performance of reference variables better than
		// pointers because the pointer indirection step is avoided.  In any case, this must be
		// a true alias to the object, not a copy of it, because it's address (&this_key) is compared
		// to other addresses for equality further below.
	}

	// The following is done even if key history is disabled because sAltTabMenuIsVisible relies on it:
	pKeyHistoryCurr->event_type = is_ignored ? 'i' : (is_artificial ? 'a' : ' '); // v1.0.42.04: 'a' was added, but 'i' takes precedence over 'a'.

	// v1.0.43: Block the Win keys during journal playback to prevent keystrokes hitting the Start Menu
	// if the user accidentally presses one of those keys during playback.  Note: Keys other than Win
	// don't need to be blocked because the playback hook defers them until after playback.
	// Only block the down-events in case the user is physically holding down the key at the start
	// of playback but releases it during the Send (avoids Win key becoming logically stuck down).
	// This also serves to block Win shortcuts such as Win+R and Win+E during playback.
	// Also, it seems best to block artificial LWIN keystrokes too, in case some other script or
	// program tries to display the Start Menu during playback.
	if (g_BlockWinKeys && (aVK == VK_LWIN || aVK == VK_RWIN) && !aKeyUp)
		return SuppressThisKey;

	// v1.0.37.07: Cancel the alt-tab menu upon receipt of Escape so that it behaves like the OS's native Alt-Tab.
	// Even if is_ignored==true, it seems more flexible/useful to cancel the Alt-Tab menu upon receiving
	// an Escape keystroke of any kind.
	// Update: Must not handle Alt-Up here in a way similar to Esc-down in case the hook sent Alt-up to
	// dismiss its own menu. Otherwise, the shift key might get stuck down if Shift-Alt-Tab was in effect.
	// Instead, the release-of-prefix-key section should handle it via its checks of this_key.it_put_shift_down, etc.
	if (sAltTabMenuIsVisible && aVK == VK_ESCAPE && !aKeyUp)
	{
		// When the alt-tab window is owned by the script (it is owned by csrss.exe unless the script
		// is the process that invoked the alt-tab window), testing shows that the script must be the
		// originator of the Escape keystroke.  Therefore, substitute a simulated keystroke for the
		// user's physical keystroke. It might be necessary to do this even if is_ignored==true because
		// a keystroke from some other script/process might not qualify as a valid means to cancel it.
		// UPDATE for v1.0.39: The escape handler below works only if the hook's thread invoked the
		// alt-tab window, not if the script's thread did via something like "Send {Alt down}{tab down}".
		// This is true even if the process ID is checked instead of the thread ID below.  I think this
		// behavior is due to the window obeying escape only when its own thread sends it.  This
		// is probably done to allow a program to automate the alt-tab menu without interference
		// from Escape keystrokes typed by the user.  Although this could probably be fixed by
		// sending a message to the main thread and having it send the Escape keystroke, it seems
		// best not to do this because:
		// 1) The ability to dismiss a script-invoked alt-tab menu with escape would vary depending on
		//    whether the keyboard hook is installed (i.e. it's inconsistent).
		// 2) It's more flexible to preserve the ability to protect the alt-tab menu from physical
		//    escape keystrokes typed by the user.  The script can simulate an escape key to explicitly
		//    close an alt-tab window it invoked (a simulated escape keystroke can apparently also close
		//    any alt-tab menu, even one invoked by physical keystrokes; but the converse isn't true).
		// 3) Lesser reason: Reduces code size and complexity.
		HWND alt_tab_window;
		if ((alt_tab_window = FindWindow("#32771", NULL)) // There is an alt-tab window...
			&& GetWindowThreadProcessId(alt_tab_window, NULL) == GetCurrentThreadId()) // ...and it's owned by the hook thread (not the main thread).
		{
			KeyEvent(KEYDOWN, VK_ESCAPE);
			// By definition, an Alt key should be logically down if the alt-tab menu is visible (even if it
			// isn't, sending an extra up-event seems harmless).  Releasing that Alt key seems best because:
			// 1) If the prefix key that pushed down the alt key is still physically held down and the user
			//    presses a new (non-alt-tab) suffix key to form a hotkey, it avoids any alt-key disruption
			//    of things such as MouseClick that that subroutine might due.
			// 2) If the user holds down the prefix, presses Escape to dismiss the menu, then presses an
			//    alt-tab suffix, testing shows that the existing alt-tab logic here in the hook will put
			//    alt or shift-alt back down if it needs to.
			KeyEvent(KEYUP, (g_modifiersLR_logical & MOD_RALT) ? VK_RMENU : VK_LMENU);
			return SuppressThisKey; // Testing shows that by contrast, the upcoming key-up on Escape doesn't require this logic.
		}
		// Otherwise, the alt-tab window doesn't exist or (more likely) it's owned by some other process
		// such as crss.exe.  Do nothing extra to avoid inteferring with the native function of Escape or
		// any remappings or hotkeys assigned to Escape.  Also, do not set sAltTabMenuIsVisible to false
		// in any of the cases here because there is logic elsewhere in the hook that does that more
		// reliably; it takes into account things such as whether the Escape keystroke will be suppressed
		// due to being a hotkey).
	}

	bool sc_takes_precedence = ksc[aSC].sc_takes_precedence;
	// Check hook type too in case a script every explicitly specifies scan code zero as a hotkey:
	key_type &this_key = *((aHook == g_KeybdHook && sc_takes_precedence) ? (ksc + aSC) : (kvk + aVK));

	// Do this after above since AllowKeyToGoToSystem requires that sc be properly determined.
	// Another reason to do it after the above is due to the fact that KEY_PHYS_IGNORE permits
	// an ignored key to be considered physical input, which is handled above:
	if (is_ignored)
	{
		// This is a key sent by our own app that we want to ignore.
		// It's important never to change this to call the SuppressKey function because
		// that function would cause an infinite loop when the Numlock key is pressed,
		// which would likely hang the entire system:
		// UPDATE: This next part is for cases where more than one script is using the hook
		// simultaneously.  In such cases, it desirable for the KEYEVENT_PHYS() of one
		// instance to affect the down-state of the current prefix-key in the other
		// instances.  This check is done here -- even though there may be a better way to
		// implement it -- to minimize the chance of side-effects that a more fundamental
		// change might cause (i.e. a more fundamental change would require a lot more
		// testing, though it might also fix more things):
		if (aExtraInfo == KEY_PHYS_IGNORE && aKeyUp && pPrefixKey == &this_key)
		{
			this_key.is_down = false;
			this_key.down_performed_action = false;  // Seems best, but only for PHYS_IGNORE.
			pPrefixKey = NULL;
		}
		return AllowKeyToGoToSystem;
	}

	if (!aKeyUp) // Set defaults for this down event.
	{
		this_key.hotkey_down_was_suppressed = false;
		this_key.hotkey_to_fire_upon_release = HOTKEY_ID_INVALID;
		// Don't do the following because of the keyboard key-repeat feature.  In other words,
		// the NO_SUPPRESS_NEXT_UP_EVENT should stay pending even in the face of consecutive
		// down-events.  Even if it's possible for the flag to never be cleared due to never
		// reaching any of the parts that clear it (which currently seems impossible), it seems
		// inconsequential since by its very nature, this_key never consults the flag.
		// this_key.no_suppress &= ~NO_SUPPRESS_NEXT_UP_EVENT;
	}

	if (aHook == g_KeybdHook)
	{
		// The below DISGUISE events are done only after ignored events are returned from, above.
		// In other words, only non-ignored events (usually physical) are disguised.  The Send {Blind}
		// method is designed with this in mind, since it's more concerned that simulated keystrokes
		// don't get disguised (i.e. it seems best to disguise physical keystrokes even during {Blind} mode).
		// Do this only after the above because the SuppressThisKey macro relies
		// on the vk variable being available.  It also relies upon the fact that sc has
		// already been properly determined. Also, in rare cases it may be necessary to disguise
		// both left and right, which is why it's not done as a generic windows key:
		if (   aKeyUp && ((sDisguiseNextLWinUp && aVK == VK_LWIN) || (sDisguiseNextRWinUp && aVK == VK_RWIN)
			|| (sDisguiseNextLAltUp && aVK == VK_LMENU) || (sDisguiseNextRAltUp && aVK == VK_RMENU))   )
		{
			// Do this first to avoid problems with reentrancy triggered by the KeyEvent() calls further below.
			switch (aVK)
			{
			case VK_LWIN: sDisguiseNextLWinUp = false; break;
			case VK_RWIN: sDisguiseNextRWinUp = false; break;
			// UPDATE: The comment below is no longer a concern since neutral keys are translated higher above
			// into their left/right-specific counterpart:
			// For now, assume VK_MENU the left alt key.  This neutral key is probably never received anyway
			// due to the nature of this type of hook on NT/2k/XP and beyond.  Later, this can be further
			// optimized to check the scan code and such (what's being done here isn't that essential to
			// start with, so it's not a high priority -- but when it is done, be sure to review the
			// above IF statement also).
			case VK_LMENU: sDisguiseNextLAltUp = false; break;
			case VK_RMENU: sDisguiseNextRAltUp = false; break;
			}
			// Send our own up-event to replace this one.  But since ours has the Shift key
			// held down for it, the Start Menu or foreground window's menu bar won't be invoked.
			// It's necessary to send an up-event so that it's state, as seen by the system,
			// is put back into the up position, which would be needed if its previous
			// down-event wasn't suppressed (probably due to the fact that this win or alt
			// key is a prefix but not a suffix).
			// Fix for v1.0.25: Use CTRL vs. Shift to avoid triggering the LAlt+Shift language-change hotkey.
			// This is definitely needed for ALT, but is done here for WIN also in case ALT is down,
			// which might cause the use of SHIFT as the disguise key to trigger the language switch.
			if (!(g_modifiersLR_logical & (MOD_LCONTROL | MOD_RCONTROL))) // Neither CTRL key is down.
				KeyEvent(KEYDOWNANDUP, g_MenuMaskKey);
			// Since the above call to KeyEvent() calls the keybd hook recursively, a quick down-and-up
			// on Control is all that is necessary to disguise the key.  This is because the OS will see
			// that the Control keystroke occurred while ALT or WIN is still down because we haven't
			// done CallNextHookEx() yet.
			// Fix for v1.0.36.07: Don't return here because this release might also be a hotkey such as
			// "~LWin Up::".
		}
	}
	else // Mouse hook
	{
		// If no vk, there's no mapping for this key, so currently there's no way to process it.
		if (!aVK)
			return AllowKeyToGoToSystem;
		// Also, if the script is displaying a menu (tray, main, or custom popup menu), always
		// pass left-button events through -- even if LButton is defined as a hotkey -- so
		// that menu items can be properly selected.  This is necessary because if LButton is
		// a hotkey, it can't launch now anyway due to the script being uninterruptible while
		// a menu is visible.  And since it can't launch, it can't do its typical "MouseClick
		// left" to send a true mouse-click through as a replacement for the suppressed
		// button-down and button-up events caused by the hotkey.  Also, for simplicity this
		// is done regardless of which modifier keys the user is holding down since the desire
		// to fire mouse hotkeys while a context or popup menu is displayed seems too rare.
		//
		// Update for 1.0.37.05: The below has been extended to look for menus beyond those
		// supported by g_MenuIsVisible, namely the context menus of a MonthCal or Edit control
		// (even the script's main window's edit control's context menu).  It has also been
		// extended to include RButton because:
		// 1) Right and left buttons may have been swapped via control panel to take on each others' functions.
		// 2) Right-click is a valid way to select a context menu items (but apparently not popup or menu bar items).
		// 3) Right-click should invoke another instance of the context menu (or dismiss existing menu, depending
		//    on where the click occurs) if user clicks outside of our thread's existing context menu.
		HWND menu_hwnd;
		if (   (aVK == VK_LBUTTON || aVK == VK_RBUTTON) && (g_MenuIsVisible // Ordered for short-circuit performance.
				|| ((menu_hwnd = FindWindow("#32768", NULL))
					&& GetWindowThreadProcessId(menu_hwnd, NULL) == g_MainThreadID))   ) // Don't call GetCurrentThreadId() because our thread is different than main's.
		{
			// Bug-fix for v1.0.22: If "LControl & LButton::" (and perhaps similar combinations)
			// is a hotkey, the foreground window would think that the mouse is stuck down, at least
			// if the user clicked outside the menu to dismiss it.  Specifically, this comes about
			// as follows:
			// The wrong up-event is suppressed:
			// ...because down_performed_action was true when it should have been false
			// ...because the while-menu-was-displayed up-event never set it to false
			// ...because it returned too early here before it could get to that part further below.
			this_key.down_performed_action = false; // Seems ok in this case to do this for both aKeyUp and !aKeyUp.
			this_key.is_down = !aKeyUp;
			return AllowKeyToGoToSystem;
		}
	} // Mouse hook.

	// Any key-down event (other than those already ignored and returned from,
	// above) should probably be considered an attempt by the user to use the
	// prefix key that's currently being held down as a "modifier".  That way,
	// if pPrefixKey happens to also be a suffix, its suffix action won't fire
	// when the key is released, which is probably the correct thing to do 90%
	// or more of the time.  But don't consider the modifiers themselves to have
	// been modified by a prefix key, since that is almost never desirable:
	if (pPrefixKey && pPrefixKey != &this_key && !aKeyUp) // There is a prefix key being held down and the user has now pressed some other key.
		if (   (aHook == g_KeybdHook) ? !this_key.as_modifiersLR : pPrefixKey->as_modifiersLR   )
			pPrefixKey->was_just_used = AS_PREFIX; // Indicate that currently-down prefix key has been "used".
	// Formerly, the above was done only for keyboard hook, not the mouse.  This was because
	// most people probably would not want a prefix key's suffix-action to be stopped
	// from firing just because a non-hotkey mouse button was pressed while the key
	// was held down (i.e. for games).  But now a small exception to this has been made:
	// Prefix keys that are also modifiers (ALT/SHIFT/CTRL/WIN) will now not fire their
	// suffix action on key-up if they modified a mouse button event (since Ctrl-LeftClick,
	// for example, is a valid native action and we don't want to give up that flexibility).

	// WinAPI docs state that for both virtual keys and scan codes:
	// "If there is no translation, the return value is zero."
	// Therefore, zero is never a key that can be validly configured (and likely it's never received here anyway).
	// UPDATE: For performance reasons, this check isn't even done.  Even if sc and vk are both zero, both kvk[0]
	// and ksc[0] should have all their attributes initialized to FALSE so nothing should happen for that key
	// anyway.
	//if (!vk && !sc)
	//	return AllowKeyToGoToSystem;

	if (!this_key.used_as_prefix && !this_key.used_as_suffix)
		return AllowKeyToGoToSystem;

	bool is_explicit_key_up_hotkey = false;                // Set default.
	HotkeyIDType hotkey_id_with_flags = HOTKEY_ID_INVALID; //
	bool firing_is_certain = false;                        //
	HotkeyIDType hotkey_id_temp; // For informal/temp storage of the ID-without-flags.

	bool down_performed_action, was_down_before_up;
	if (aKeyUp)
	{
		// Save prior to reset.  These var's should only be used further below in conjunction with aKeyUp
		// being TRUE.  Otherwise, their values will be unreliable (refer to some other key, probably).
		was_down_before_up = this_key.is_down;
		down_performed_action = this_key.down_performed_action;  // Save prior to reset below.
		// Reset these values in preparation for the next call to this procedure that involves this key:
		this_key.down_performed_action = false;
		if (this_key.hotkey_to_fire_upon_release != HOTKEY_ID_INVALID)
		{
			hotkey_id_with_flags = this_key.hotkey_to_fire_upon_release;
			is_explicit_key_up_hotkey = true; // Can't rely on (hotkey_id_with_flags & HOTKEY_KEY_UP) because some key-up hotkeys (such as the hotkey_up array) might not be flagged that way.
			// The line below is done even though the down-event also resets it in case it is ever
			// possible for keys to generate mulitple consecutive key-up events (faulty or unusual keyboards?)
			this_key.hotkey_to_fire_upon_release = HOTKEY_ID_INVALID;
		}
	}
	this_key.is_down = !aKeyUp;
	bool modifiers_were_corrected = false;

	if (aHook == g_KeybdHook)
	{
		// The below was added to fix hotkeys that have a neutral suffix such as "Control & LShift".
		// It may also fix other things and help future enhancements:
		if (this_key.as_modifiersLR)
		{
			// The neutral modifier "Win" is not currently supported.
			kvk[VK_CONTROL].is_down = kvk[VK_LCONTROL].is_down || kvk[VK_RCONTROL].is_down;
			kvk[VK_MENU].is_down = kvk[VK_LMENU].is_down || kvk[VK_RMENU].is_down;
			kvk[VK_SHIFT].is_down = kvk[VK_LSHIFT].is_down || kvk[VK_RSHIFT].is_down;
			// No longer possible because vk is translated early on from neutral to left-right specific:
			// I don't think these ever happen with physical keyboard input, but it might with artificial input:
			//case VK_CONTROL: kvk[sc == SC_RCONTROL ? VK_RCONTROL : VK_LCONTROL].is_down = !aKeyUp; break;
			//case VK_MENU: kvk[sc == SC_RALT ? VK_RMENU : VK_LMENU].is_down = !aKeyUp; break;
			//case VK_SHIFT: kvk[sc == SC_RSHIFT ? VK_RSHIFT : VK_LSHIFT].is_down = !aKeyUp; break;
		}
	}
	else // Mouse hook
	{
		// If the mouse hook is installed without the keyboard hook, update g_modifiersLR_logical
		// manually so that it can be referred to by the mouse hook after this point:
		if (!g_KeybdHook)
		{
			g_modifiersLR_logical = g_modifiersLR_logical_non_ignored = GetModifierLRState(true);
			modifiers_were_corrected = true;
		}
	}

	modLR_type modifiersLRnew;

	bool this_toggle_key_can_be_toggled = this_key.pForceToggle && *this_key.pForceToggle == NEUTRAL; // Relies on short-circuit boolean order.

	///////////////////////////////////////////////////////////////////////////////////////
	// CASE #1 of 4: PREFIX key has been pressed down.  But use it in this capacity only if
	// no other prefix is already in effect or if this key isn't a suffix.  Update: Or if
	// this key-down is the same as the prefix already down, since we want to be able to
	// a prefix when it's being used in its role as a modified suffix (see below comments).
	///////////////////////////////////////////////////////////////////////////////////////
	if (this_key.used_as_prefix && !aKeyUp && (!pPrefixKey || !this_key.used_as_suffix || &this_key == pPrefixKey))
	{
		// v1.0.41: Even if this prefix key is non-suppressed (passes through to active window),
		// still call PrefixHasNoEnabledSuffixes() because don't want to overwrite the old value of
		// pPrefixKey (see comments in "else" later below).
		// v1.0.44: Added check for PREFIX_ACTUAL so that a PREFIX_FORCED prefix will be considered
		// a prefix even if it has no suffixes.  This fixes an unintentional change in v1.0.41 where
		// naked, neutral modifier hotkeys Control::, Alt::, and Shift:: started firing on press-down
		// rather than release as intended.  The PREFIX_FORCED facility may also provide the means to
		// introduce a new hotkey modifier such as an "up2" keyword that makes any key into a prefix
		// key even if it never acts as a prefix for other keys, which in turn has the benefit of firing
		// on key-up, but only if the no other key was pressed while the user was holding it down.
		bool has_no_enabled_suffixes;
		if (   !(has_no_enabled_suffixes = (this_key.used_as_prefix == PREFIX_ACTUAL)
			&& Hotkey::PrefixHasNoEnabledSuffixes(sc_takes_precedence ? aSC : aVK, sc_takes_precedence))   )
		{
			// This check is necessary in cases such as the following, in which the "A" key continues
			// to repeat becauses pressing a mouse button (unlike pressing a keyboard key) does not
			// stop the prefix key from repeating:
			// $a::send, a
			// a & lbutton::
			if (&this_key != pPrefixKey)
			{
				// Override any other prefix key that might be in effect with this one, in case the
				// prior one, due to be old for example, was invalid somehow.  UPDATE: It seems better
				// to leave the old one in effect to support the case where one prefix key is modifying
				// a second one in its role as a suffix.  In other words, if key1 is a prefix and
				// key2 is both a prefix and a suffix, we want to leave key1 in effect as a prefix,
				// rather than key2.  Hence, a null-check was added in the above if-stmt:
				pPrefixKey = &this_key;
				// It should be safe to init this because even if the current key is repeating,
				// it should be impossible to receive here the key-downs that occurred after
				// the first, because there's a return-on-repeat check farther above (update: that check
				// is gone now).  Even if that check weren't done, it's safe to reinitialize this to zero
				// because on most (all?) keyboards & OSs, the moment the user presses another key while
				// this one is held down, the key-repeating ceases and does not resume for
				// this key (though the second key will begin to repeat if it too is held down).
				// In other words, the fear that this would be wrongly initialized and thus cause
				// this prefix's suffix-action to fire upon key-release seems unfounded.
				// It seems easier (and may perform better than alternative ways) to init this
				// here rather than say, upon the release of the prefix key:
				this_key.was_just_used = 0; // Init to indicate it hasn't yet been used in its role as a prefix.
			}
		}
		//else this prefix has no enabled suffixes, so its role as prefix is also disabled.
		// Therefore, don't set pPrefixKey to this_key because don't want the following line
		// (in another section) to execute when a suffix comes in (there may be other reasons too,
		// such as not wanting to lose track of the previous prefix key in cases where the user is
		// holding down more than one prefix):
		// pPrefixKey->was_just_used = AS_PREFIX

		if (this_key.used_as_suffix) // v1.0.41: Added this check to avoid doing all of the below when unnecessary.
		{
			// This new section was added May 30, 2004, to fix scenarios such as the following example:
			// a & b::Msgbox a & b
			// $^a::MsgBox a
			// Previously, the ^a hotkey would only fire on key-up (unless it was registered, in which
			// case it worked as intended on the down-event).  When the user presses A, it's okay (and
			// probably desirable) to have recorded that event as a prefix-key-down event (above).
			// But in addition to that, we now check if this is a normal, modified hotkey that should
			// fire now rather than waiting for the key-up event.  This is done because it makes sense,
			// it's more correct, and also it makes the behavior of a hooked ^a hotkey consistent with
			// that of a registered ^a.

			// Prior to considering whether to fire a hotkey, correct the hook's modifier state.
			// Although this is rarely needed, there are times when the OS disables the hook, thus
			// it is possible for it to miss keystrokes.  See comments in GetModifierLRState()
			// for more info:
			if (!modifiers_were_corrected)
			{
				modifiers_were_corrected = true;
				GetModifierLRState(true);
			}

			// non_ignored is always used when considering whether a key combination is in place to
			// trigger a hotkey:
			modifiersLRnew = g_modifiersLR_logical_non_ignored;
			if (this_key.as_modifiersLR) // This will always be false if our caller is the mouse hook.
				// Hotkeys are not defined to modify themselves, so look for a match accordingly.
				modifiersLRnew &= ~this_key.as_modifiersLR;
			// For this case to be checked, there must be at least one modifier key currently down (other
			// than this key itself if it's a modifier), because if there isn't and this prefix is also
			// a suffix, its suffix action should only fire on key-up (i.e. not here, but later on).
			// UPDATE: In v1.0.41, an exception to the above is when a prefix is disabled via
			// has_no_enabled_suffixes, in which case it seems desirable for most uses to have its
			// suffix action fire on key-down rather than key-up.
			if (modifiersLRnew || has_no_enabled_suffixes)
			{
				// Check hook type too in case a script every explicitly specifies scan code zero as a hotkey:
				hotkey_id_with_flags = (aHook == g_KeybdHook && sc_takes_precedence)
					? Kscm(modifiersLRnew, aSC) : Kvkm(modifiersLRnew, aVK);
				if (hotkey_id_with_flags & HOTKEY_KEY_UP) // And it's okay even if it's is HOTKEY_ID_INVALID.
				{
					// Queue it for later, which is done here rather than upon release of the key so that
					// the user can release the key's modifiers before releasing the key itself, which
					// is likely to happen pretty often. v1.0.41: This is done even if the hotkey is subject
					// to #IfWin because it seems more correct to check those criteria at the actual time
					// the key is released rather than now:
					this_key.hotkey_to_fire_upon_release = hotkey_id_with_flags;
					hotkey_id_with_flags = HOTKEY_ID_INVALID;
				}
				else // hotkey_id_with_flags is either HOTKEY_ID_INVALID or a valid key-down hotkey.
				{
					hotkey_id_temp = hotkey_id_with_flags & HOTKEY_ID_MASK;
					if (hotkey_id_temp < Hotkey::sHotkeyCount)
						this_key.hotkey_to_fire_upon_release = hotkey_up[hotkey_id_temp]; // Might assign HOTKEY_ID_INVALID.
					// Since this prefix key is being used in its capacity as a suffix instead,
					// hotkey_id_with_flags now contains a hotkey ready for firing later below.
					// v1.0.41: Above is done even if the hotkey is subject to #IfWin because:
					// 1) The down-hotkey's #IfWin criteria might be different from that of the up's.
					// 2) It seems more correct to check those criteria at the actual time the key is
					// released rather than now (and also probably reduces code size).
				}
			}
			// Alt-tab need not be checked here (like it is in the similar section below) because all
			// such hotkeys use (or were converted at load-time to use) a modifier_vk, not a set of
			// modifiers or modifierlr's.
		} // if (this_key.used_as_suffix)

		if (hotkey_id_with_flags == HOTKEY_ID_INVALID)
		{
			if (has_no_enabled_suffixes)
			{
				this_key.no_suppress |= NO_SUPPRESS_NEXT_UP_EVENT; // Since the "down" is non-suppressed, so should the "up".
				pKeyHistoryCurr->event_type = '#'; // '#' to indicate this prefix key is disabled due to #IfWin criterion.
			}
			// In this case, a key-down event can't trigger a suffix, so return immediately.
			// If our caller is the mouse hook, both of the following will always be false:
			// this_key.as_modifiersLR
			// this_toggle_key_can_be_toggled
			return (this_key.as_modifiersLR || (this_key.no_suppress & NO_SUPPRESS_PREFIX)
				|| this_toggle_key_can_be_toggled || has_no_enabled_suffixes)
				? AllowKeyToGoToSystem : SuppressThisKey;
		}
		//else valid suffix hotkey has been found; this will now fall through to Case #4 by virtue of aKeyUp==false.
	}

	//////////////////////////////////////////////////////////////////////////////////
	// CASE #2 of 4: SUFFIX key (that's not a prefix, or is one but has just been used
	// in its capacity as a suffix instead) has been released.
	// This is done before Case #3 for performance reasons.
	//////////////////////////////////////////////////////////////////////////////////
	// v1.0.37.05: Added "|| down_performed_action" to the final check below because otherwise a
	// script such as the following would send two M's for +b, one upon down and one upon up:
	// +b::Send, M
	// b & z::return
	// I don't remember exactly what the "pPrefixKey != &this_key" check is for below, but it is kept
	// to minimize the chance of breaking other things:
	bool fell_through_from_case2 = false; // Set default.
	if (this_key.used_as_suffix && aKeyUp && (pPrefixKey != &this_key || down_performed_action)) // Note: hotkey_id_with_flags might be already valid due to this_key.hotkey_to_fire_upon_release.
	{
		if (pPrefixKey == &this_key) // v1.0.37.05: Added so that scripts such as the example above don't leave pPrefixKey wrongly non-NULL.
			pPrefixKey = NULL;       // Also, it seems unnecessary to check this_key.it_put_alt_down and such like is done in Case #3.
		// If it did perform an action, suppress this key-up event.  Do this even
		// if this key is a modifier because it's previous key-down would have
		// already been suppressed (since this case is for suffixes that aren't
		// also prefixes), thus the key-up can be safely suppressed as well.
		// It's especially important to do this for keys whose up-events are
		// special actions within the OS, such as AppsKey, Lwin, and Rwin.
		// Toggleable keys are also suppressed here on key-up because their
		// previous key-down event would have been suppressed in order for
		// down_performed_action to be true.  UPDATE: Added handling for
		// NO_SUPPRESS_NEXT_UP_EVENT and also applied this next part to both
		// mouse and keyboard.
		// v1.0.40.01: It was observed that a hotkey that consists of a mouse button as a prefix and
		// a keyboard key as a suffix can cause sticking keys in rare cases.  For example, when
		// "MButton & LShift" is a hotkey, if you hold down LShift long enough for it to begin
		// auto-repeating then press MButton, the hotkey fires the next time LShift auto-repeats (since
		// pressing a mouse button doesn't stop a keyboard key from auto-repeating).  Fixing that type
		// of firing seems likely to break more things than it fixes.  But since it isn't fixed, when
		// the user releases LShift, the up-event is suppressed here, which causes the key to get
		// stuck down.  That could be fixed in the following ways, but all of them seem likely to break
		// more things than they fix, especially given the rarity that both a hotkey of this type would
		// exist and its mirror image does something useful that isn't a hotkey (for example, Shift+MButton
		// is a meaningful action in few if any applications):
		// 1) Don't suppress the physical release of a suffix key if that key is logically down (as reported
		//    by GetKeyState/GetAsyncKeyState): Seems too broad in scope because there might be cases where
		//    the script or user wants the key to stay logically down (e.g. Send {Shift down}{a down}).
		// 2) Same as #1 but limit the non-suppression to only times when the suffix key was logically down
		//    when its first qualified physical down-event came in.  This is definitely better but like
		//    #1, the uncertainty of breaking existing scripts and/or causing more harm than good seems too
		//    high.
		// 3) Same as #2 but limit it only to cases where the suffix key is a keyboard key and its prefix
		//    is a mouse key.  Although very selective, it doesn't mitigate the fact it might still do more
		//    harm than good and/or break existing scripts.
		// In light of the above, it seems best to keep this documented here as a known limitation for now.
		//
		// v1.0.28: The following check is done to support certain keyboards whose keys or scroll wheels
		// generate up events without first having generated any down-event for the key.  UPDATE: I think
		// this check is now also needed to allow fall-through in cases like "b" and "b up" both existing.
		if (!this_key.used_as_key_up)
		{
			bool suppress_up_event;
			if (this_key.no_suppress & NO_SUPPRESS_NEXT_UP_EVENT)
			{
				suppress_up_event = false;
				this_key.no_suppress &= ~NO_SUPPRESS_NEXT_UP_EVENT;  // This ticket has been used up.
			}
			else // the default is to suppress the up-event.
				suppress_up_event = true;
			return (down_performed_action && suppress_up_event) ? SuppressThisKey : AllowKeyToGoToSystem;
		}
		//else continue checking to see if the right modifiers are down to trigger one of this
		// suffix key's key-up hotkeys.
		fell_through_from_case2 = true;
	}

	//////////////////////////////////////////////
	// CASE #3 of 4: PREFIX key has been released.
	//////////////////////////////////////////////
	if (this_key.used_as_prefix && aKeyUp) // If these are true, hotkey_id_with_flags should be valid only by means of this_key.hotkey_to_fire_upon_release.
	{
		if (pPrefixKey == &this_key)
			pPrefixKey = NULL;
		// Else it seems best to keep the old one in effect.  This could happen, for example,
		// if the user holds down prefix1, holds down prefix2, then releases prefix1.
		// In that case, we would want to keep the most recent prefix (prefix2) in effect.
		// This logic would fail to work properly in a case like this if the user releases
		// prefix2 but still has prefix1 held down.  The user would then have to release
		// prefix1 and press it down again to get the hook to realize that it's in effect.
		// This seems very unlikely to be something commonly done by anyone, so for now
		// it's just documented here as a limitation.

		if (this_key.it_put_alt_down) // key pushed ALT down, or relied upon it already being down, so go up:
		{
			this_key.it_put_alt_down = false;
			KeyEvent(KEYUP, VK_MENU);
		}
		if (this_key.it_put_shift_down) // similar to above
		{
			this_key.it_put_shift_down = false;
			KeyEvent(KEYUP, VK_SHIFT);
		}

		// Section added in v1.0.41:
		// Fix for v1.0.44.04: Defer use of the ticket and avoid returning here if hotkey_id_with_flags is valid,
		// which only happens by means of this_key.hotkey_to_fire_upon_release.  This fixes custom combination
		// hotkeys whose composite hotkey is also present such as:
		//LShift & ~LAlt::
		//LAlt & ~LShift::
		//LShift & ~LAlt up::
		//LAlt & ~LShift up::
		//ToolTip %A_ThisHotkey%
		//return
		if (hotkey_id_with_flags == HOTKEY_ID_INVALID && this_key.no_suppress & NO_SUPPRESS_NEXT_UP_EVENT)
		{
			this_key.no_suppress &= ~NO_SUPPRESS_NEXT_UP_EVENT;  // This ticket has been used up.
			return AllowKeyToGoToSystem; // This should handle pForceToggle for us, suppressing if necessary.
		}

		if (this_toggle_key_can_be_toggled) // Always false if our caller is the mouse hook.
		{
			// It's done this way because CapsLock, for example, is a key users often
			// press quickly while typing.  I suspect many users are like me in that
			// they're in the habit of not having releasing the CapsLock key quite yet
			// before they resume typing, expecting it's new mode to be in effect.
			// This resolves that problem by always toggling the state of a toggleable
			// key upon key-down.  If this key has just acted in its role of a prefix
			// to trigger a suffix action, toggle its state back to what it was before
			// because the firing of a hotkey should not have the side-effect of also
			// toggling the key:
			// Toggle the key by replacing this key-up event with a new sequence
			// of our own.  This entire-replacement is done so that the system
			// will see all three events in the right order:
			if (this_key.was_just_used == AS_PREFIX_FOR_HOTKEY) // If this is true, it's probably impossible for hotkey_id_with_flags to be valid by means of this_key.hotkey_to_fire_upon_release.
			{
				KEYEVENT_PHYS(KEYUP, aVK, aSC); // Mark it as physical for any other hook instances.
				KeyEvent(KEYDOWNANDUP, aVK, aSC);
				return SuppressThisKey;
			}

			// Otherwise, if it was used to modify a non-suffix key, or it was just
			// pressed and released without any keys in between, don't suppress its up-event
			// at all.  UPDATE: Don't return here if it didn't modify anything because
			// this prefix might also be a suffix. Let later sections handle it then.
			if (this_key.was_just_used == AS_PREFIX)
				return AllowKeyToGoToSystem;
		}
		else // It's not a toggleable key, or it is but it's being kept forcibly on or off.
			// Seems safest to suppress this key if the user pressed any non-modifier key while it
			// was held down.  As a side-effect of this, if the user holds down numlock, for
			// example, and then presses another key that isn't actionable (i.e. not a suffix),
			// the numlock state won't be toggled even it's normally configured to do so.
			// This is probably the right thing to do in most cases.
			// Older note:
			// In addition, this suppression is relied upon to prevent toggleable keys from toggling
			// when they are used to modify other keys.  For example, if "Capslock & A" is a hotkey,
			// the state of the Capslock key should not be changed when the hotkey is pressed.
			// Do this check prior to the below check (give it precedence).
			if (this_key.was_just_used  // AS_PREFIX or AS_PREFIX_FOR_HOTKEY.
				&& hotkey_id_with_flags == HOTKEY_ID_INVALID) // v1.0.44.04: Must check this because this prefix might be being used in its role as a suffix instead.
			{
				if (this_key.as_modifiersLR) // Always false if our caller is the mouse hook.
					return (this_key.was_just_used == AS_PREFIX_FOR_HOTKEY) ? AllowKeyToGoToSystemButDisguiseWinAlt
						: AllowKeyToGoToSystem;  // i.e. don't disguise Win or Alt key if it didn't fire a hotkey.
				// Otherwise:
				return (this_key.no_suppress & NO_SUPPRESS_PREFIX) ? AllowKeyToGoToSystem : SuppressThisKey;
			}

		// v1.0.41: This spot cannot be reached when a disabled prefix key's up-action fires on
		// key-down instead (via Case #1).  This is because upon release, that prefix key would be
		// returned from in Case #2 (by virtue of its check of down_performed_action).

		// Since above didn't return, this key-up for this prefix key wasn't used in it's role
		// as a prefix.  If it's not a suffix, we're done, so just return.  Don't do
		// "DisguiseWinAlt" because we want the key's native key-up function to take effect.
		// Also, allow key-ups for toggleable keys that the user wants to be toggleable to
		// go through to the system, because the prior key-down for this prefix key
		// wouldn't have been suppressed and thus this up-event goes with it (and this
		// up-event is also needed by the OS, at least WinXP, to properly set the indicator
		// light and toggle state):
		if (!this_key.used_as_suffix)
			// If our caller is the mouse hook, both of the following will always be false:
			// this_key.as_modifiersLR
			// this_toggle_key_can_be_toggled
			return (this_key.as_modifiersLR || (this_key.no_suppress & NO_SUPPRESS_PREFIX)
				// The order on this line important; it relies on short-circuit boolean:
				|| this_toggle_key_can_be_toggled) ? AllowKeyToGoToSystem : SuppressThisKey;

		// Since the above didn't return, this key is both a prefix and a suffix, but
		// is currently operating in its capacity as a suffix.
		// If this key wasn't thought to be down prior to this up-event, it's probably because
		// it is registered with another prefix by RegisterHotkey().  In this case, the keyup
		// should be passed back to the system rather than performing it's key-up suffix
		// action.  UPDATE: This can't happen with a low-level hook.  But if there's another
		// low-level hook installed that receives events before us, and it's not
		// well-implemented (i.e. it sometimes sends ups without downs), this check
		// may help prevent unexpected behavior.  UPDATE: The check "!this_key.used_as_key_up"
		// is now done too so that an explicit key-up hotkey can operate even if the key wasn't
		// thought to be down before. One thing this helps with is certain keyboards (e.g. some
		// Dells) that generate only up events for some of their special keys but no down events,
		// even when *no* keyboard management software is installed). Some keyboards also have
		// scroll wheels that generate a stream of up events in one direction and down in the other.
		if (!(was_down_before_up || this_key.used_as_key_up)) // Verified correct.
			return AllowKeyToGoToSystem;
		//else v1.0.37.05: Since no suffix action was triggered while it was held down, fall through
		// rather than returning so that the key's own unmodified/naked suffix action will be considered.
		// For example:
		// a & b::
		// a::   // This fires upon release of "a".
	}

	////////////////////////////////////////////////////////////////////////////////////////////////////
	// CASE #4 of 4: SUFFIX key has been pressed down (or released if it's a key-up event, in which case
	// it fell through from CASE #3 or #2 above).  This case can also happen if it fell through from
	// case #1 (i.e. it already determined the value of hotkey_id_with_flags).
	////////////////////////////////////////////////////////////////////////////////////////////////////
	// First correct modifiers, because at this late a state, the likelihood of firing a hotkey is high.
	// For details, see comments for "modifiers_were_corrected" above:
	if (!modifiers_were_corrected && hotkey_id_with_flags == HOTKEY_ID_INVALID)
	{
		modifiers_were_corrected = true;
		GetModifierLRState(true);
	}
	bool fire_with_no_suppress = false; // Set default.

	if (pPrefixKey && (!aKeyUp || this_key.used_as_key_up) && hotkey_id_with_flags == HOTKEY_ID_INVALID) // Helps performance by avoiding all the below checking.
	{
		// Action here is considered first, and takes precedence since a suffix's ModifierVK/SC should
		// take effect regardless of whether any win/ctrl/alt/shift modifiers are currently down, even if
		// those modifiers themselves form another valid hotkey with this suffix.  In other words,
		// ModifierVK/SC combos take precedence over normally-modified combos:
		int i;
		vk_type modifier_vk;
		sc_type modifier_sc;
		for (i = 0; i < this_key.nModifierVK; ++i)
		{
			vk_hotkey &this_modifier_vk = this_key.ModifierVK[i]; // For performance and convenience.
			// The following check supports the prefix+suffix pairs that have both an up hotkey and a down,
			// such as:
			//a & b::     ; Down.
			//a & b up::  ; Up.
			//MsgBox %A_ThisHotkey%
			//return
			if (kvk[this_modifier_vk.vk].is_down) // A prefix key qualified to trigger this suffix is down.
			{
				if (this_modifier_vk.id_with_flags & HOTKEY_KEY_UP)
				{
					if (!aKeyUp) // Key-up hotkey but the event is a down-event.
					{
						// Queue the up-hotkey for later so that the user is free to release the
						// prefix key prior to releasing the suffix (which seems quite common and
						// thus desirable).  v1.0.41: This is done even if the hotkey is subject
						// to #IfWin because it seems more correct to check those criteria at the actual time
						// the key is released rather than now:
						this_key.hotkey_to_fire_upon_release = this_modifier_vk.id_with_flags;
						if (hotkey_id_with_flags != HOTKEY_ID_INVALID) // i.e. a previous iteration already found the down-event to fire.
							break;
						//else continue searching for the down hotkey that goes with this up (if any).
					}
					else // this hotkey is qualified to fire.
					{
						hotkey_id_with_flags = this_modifier_vk.id_with_flags;
						modifier_vk = this_modifier_vk.vk;
						break;
					}
				}
				else // This is a normal hotkey that fires on suffix-key-down.
				{
					if (!aKeyUp)
					{
						hotkey_id_with_flags = this_modifier_vk.id_with_flags;
						modifier_vk = this_modifier_vk.vk; // Set this now in case loop ends on its own (not via break).
						// and continue searching for the up hotkey (if any) to queue up for firing upon the key's release).
					}
					//else this key-down hotkey can't fire because the current event is a up-event.
					// But continue searching for an up-hotkey in case this key is of the type that never
					// generates down-events (e.g. certain Dell keyboards).
				}
			} // qualified prefix is down
		} // for each prefix of this suffix
		if (hotkey_id_with_flags != HOTKEY_ID_INVALID)
		{
			// Update pPrefixKey, even though it was probably already done close to the top of the function,
			// just in case this for-loop changed the value pPrefixKey (perhaps because there
			// is currently more than one prefix being held down).
			// Since the hook is now designed to receive only left/right specific modifier keys
			// -- never the neutral keys -- never indicate that a neutral prefix key is down because
			// then it would never be released properly by the other main prefix/suffix handling
			// cases of the hook.  Instead, always identify which prefix key (left or right) is
			// in effect:
			switch (modifier_vk)
			{
			case VK_SHIFT: pPrefixKey = kvk + (kvk[VK_RSHIFT].is_down ? VK_RSHIFT : VK_LSHIFT); break;
			case VK_CONTROL: pPrefixKey = kvk + (kvk[VK_RCONTROL].is_down ? VK_RCONTROL : VK_LCONTROL); break;
			case VK_MENU: pPrefixKey = kvk + (kvk[VK_RMENU].is_down ? VK_RMENU : VK_LMENU); break;
			default: pPrefixKey = kvk + modifier_vk;
			}
		}
		else // Now check scan codes since above didn't find a valid hotkey.
		{
			for (i = 0; i < this_key.nModifierSC; ++i)
			{
				sc_hotkey &this_modifier_sc = this_key.ModifierSC[i]; // For performance and convenience.
				if (ksc[this_modifier_sc.sc].is_down)
				{
					// See similar section above for comments about the section below:
					if (this_modifier_sc.id_with_flags & HOTKEY_KEY_UP)
					{
						if (!aKeyUp)
						{
							this_key.hotkey_to_fire_upon_release = this_modifier_sc.id_with_flags;
							if (hotkey_id_with_flags != HOTKEY_ID_INVALID) // i.e. a previous iteration already found the down-event to fire.
								break;
						}
						else // this hotkey is qualified to fire.
						{
							hotkey_id_with_flags = this_modifier_sc.id_with_flags;
							modifier_sc = this_modifier_sc.sc;
							break;
						}
					}
					else // This is a normal hotkey that fires on suffix key-down.
					{
						if (!aKeyUp)
						{
							hotkey_id_with_flags = this_modifier_sc.id_with_flags;
							modifier_sc = this_modifier_sc.sc; // Set this now in case loop ends on its own (not via break).
						}
					}
				}
			} // for()
			if (hotkey_id_with_flags != HOTKEY_ID_INVALID)
				// Update pPrefixKey, even though it was probably already done close to the top of the function,
				// just in case this for-loop changed the value pPrefixKey (perhaps because there
				// is currently more than one prefix being held down).
				pPrefixKey = ksc + modifier_sc;
		} // The check for scan code vs. VK.
		if (hotkey_id_with_flags == HOTKEY_ID_INVALID)
		{
			// Search again, but this time do it with this_key translated into its neutral counterpart.
			// This avoids the need to display a warning dialog for an example such as the following,
			// which was previously unsupported:
			// AppsKey & Control::MsgBox %A_ThisHotkey%
			// Note: If vk was a neutral modifier when it first came in (e.g. due to NT4), it was already
			// translated early on (above) to be non-neutral.
			vk_type vk_neutral = 0;  // Set default.  Note that VK_LWIN/VK_RWIN have no neutral VK.
			switch (aVK)
			{
			case VK_LCONTROL:
			case VK_RCONTROL: vk_neutral = VK_CONTROL; break;
			case VK_LMENU:
			case VK_RMENU:    vk_neutral = VK_MENU; break;
			case VK_LSHIFT:
			case VK_RSHIFT:   vk_neutral = VK_SHIFT; break;
			}
			if (vk_neutral)
			{
				// These next two for() loops are nearly the same as the ones above, so see comments there
				// and maintain them together:
				int max = kvk[vk_neutral].nModifierVK;
				for (i = 0; i < max; ++i)
				{
					vk_hotkey &this_modifier_vk = kvk[vk_neutral].ModifierVK[i]; // For performance and convenience.
					if (kvk[this_modifier_vk.vk].is_down)
					{
						// See similar section above for comments about the section below:
						if (this_modifier_vk.id_with_flags & HOTKEY_KEY_UP)
						{
							if (!aKeyUp)
							{
								this_key.hotkey_to_fire_upon_release = this_modifier_vk.id_with_flags;
								if (hotkey_id_with_flags != HOTKEY_ID_INVALID) // i.e. a previous iteration already found the down-event to fire.
									break;
							}
							else // this hotkey is qualified to fire.
							{
								hotkey_id_with_flags = this_modifier_vk.id_with_flags;
								modifier_vk = this_modifier_vk.vk;
								break;
							}
						}
						else // This is a normal hotkey that fires on suffix key-down.
						{
							if (!aKeyUp)
							{
								hotkey_id_with_flags = this_modifier_vk.id_with_flags;
								modifier_vk = this_modifier_vk.vk; // Set this now in case loop ends on its own (not via break).
							}
						}
					}
				}
				if (hotkey_id_with_flags != HOTKEY_ID_INVALID)
				{
					// See the nearly identical section above for comments on the below:
					switch (modifier_vk)
					{
					case VK_SHIFT: pPrefixKey = kvk + (kvk[VK_RSHIFT].is_down ? VK_RSHIFT : VK_LSHIFT); break;
					case VK_CONTROL: pPrefixKey = kvk + (kvk[VK_RCONTROL].is_down ? VK_RCONTROL : VK_LCONTROL); break;
					case VK_MENU: pPrefixKey = kvk + (kvk[VK_RMENU].is_down ? VK_RMENU : VK_LMENU); break;
					default: pPrefixKey = kvk + modifier_vk;
					}
				}
				else  // Now check scan codes since above didn't find one.
				{
					for (max = kvk[vk_neutral].nModifierSC, i = 0; i < max; ++i)
					{
						sc_hotkey &this_modifier_sc = kvk[vk_neutral].ModifierSC[i]; // For performance and convenience.
						if (ksc[this_modifier_sc.sc].is_down)
						{
							// See similar section above for comments about the section below:
							if (this_modifier_sc.id_with_flags & HOTKEY_KEY_UP)
							{
								if (!aKeyUp)
								{
									this_key.hotkey_to_fire_upon_release = this_modifier_sc.id_with_flags;
									if (hotkey_id_with_flags != HOTKEY_ID_INVALID) // i.e. a previous iteration already found the down-event to fire.
										break;
								}
								else // this hotkey is qualified to fire.
								{
									hotkey_id_with_flags = this_modifier_sc.id_with_flags;
									modifier_sc = this_modifier_sc.sc;
									break;
								}
							}
							else // This is a normal hotkey that fires on suffix key-down.
							{
								if (!aKeyUp)
								{
									hotkey_id_with_flags = this_modifier_sc.id_with_flags;
									modifier_sc = this_modifier_sc.sc; // Set this now in case loop ends on its own (not via break).
								}
							}
						}
					} // for()
					if (hotkey_id_with_flags != HOTKEY_ID_INVALID)
						pPrefixKey = ksc + modifier_sc; // See the nearly identical section above for comments.
				} // SC vs. VK
			} // If this_key has a counterpart neutral modifier.
		} // Above block searched for match using neutral modifier.

		hotkey_id_temp = hotkey_id_with_flags & HOTKEY_ID_MASK; // For use both inside and outside the block below.
		if (hotkey_id_with_flags != HOTKEY_ID_INVALID)
		{
			// v1.0.41:
			if (hotkey_id_temp < Hotkey::sHotkeyCount) // Don't call the below for Alt-tab hotkeys and similar.
				if (   !(firing_is_certain = Hotkey::CriterionFiringIsCertain(hotkey_id_with_flags
					, aKeyUp, this_key.no_suppress, fire_with_no_suppress, &pKeyHistoryCurr->event_type))   )
					return AllowKeyToGoToSystem; // This should handle pForceToggle for us, suppressing if necessary.
				else // The naked hotkey ID may have changed, so update it (flags currently don't matter in this case).
					hotkey_id_temp = hotkey_id_with_flags & HOTKEY_ID_MASK; // Update in case CriterionFiringIsCertain() changed it.
			pPrefixKey->was_just_used = AS_PREFIX_FOR_HOTKEY;
		}
		//else: No "else if" here to avoid an extra indentation for the whole section below (it's not needed anyway).

		// Alt-tab: Alt-tab actions that require a prefix key are handled directly here rather than via
		// posting a message back to the main window.  In part, this is because it would be difficult
		// to design a way to tell the main window when to release the alt-key.
		if (hotkey_id_temp == HOTKEY_ID_ALT_TAB || hotkey_id_temp == HOTKEY_ID_ALT_TAB_SHIFT)
		{
			// Not sure if it's necessary to set this in this case.  Review.
			if (!aKeyUp)
				this_key.down_performed_action = true; // aKeyUp is known to be false due to an earlier check.
		
			if (   !(g_modifiersLR_logical & (MOD_LALT | MOD_RALT))   )  // Neither ALT key is down.
				// Note: Don't set the ignore-flag in this case because we want the hook to notice it.
				// UPDATE: It might be best, after all, to have the hook ignore these keys.  That's because
				// we want to avoid any possibility that other hotkeys will fire off while the user is
				// alt-tabbing (though we can't stop that from happening if they were registered with
				// RegisterHotkey).  In other words, since the
				// alt-tab window is in the foreground until the user released the substitute-alt key,
				// don't allow other hotkeys to be activated.  One good example that this helps is the case
				// where <key1> & rshift is defined as alt-tab but <key1> & <key2> is defined as shift-alt-tab.
				// In that case, if we didn't ignore these events, one hotkey might unintentionally trigger
				// the other.
				KeyEvent(KEYDOWN, VK_MENU);
				// And leave it down until a key-up event on the prefix key occurs.

			if ((aVK == VK_LCONTROL || aVK == VK_RCONTROL) && !aKeyUp)
				// Even though this suffix key would have been suppressed, it seems that the
				// OS's alt-tab functionality sees that it's down somehow and thus this is necessary
				// to allow the alt-tab menu to appear.  This doesn't need to be done for any other
				// modifier than Control, nor any normal key since I don't think normal keys
				// being in a down-state causes any problems with alt-tab:
				KeyEvent(KEYUP, aVK, aSC);

			// Update the prefix key's
			// flag to indicate that it was this key that originally caused the alt-key to go down,
			// so that we know to set it back up again when the key is released.  UPDATE: Actually,
			// it's probably better if this flag is set regardless of whether ALT is already down.
			// That way, in case it's state go stuck down somehow, it will be reset by an Alt-TAB
			// (i.e. alt-tab will always behave as expected even if ALT was down before starting).
			// Note: pPrefixKey must already be non-NULL or this couldn't be an alt-tab event:
			pPrefixKey->it_put_alt_down = true;
			if (hotkey_id_temp == HOTKEY_ID_ALT_TAB_SHIFT)
			{
				if (   !(g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT))   ) // Neither SHIFT key is down.
					KeyEvent(KEYDOWN, VK_SHIFT);  // Same notes apply to this key.
				pPrefixKey->it_put_shift_down = true;
			}
			// And this may do weird things if VK_TAB itself is already assigned a as a naked hotkey, since
			// it will recursively call the hook, resulting in the launch of some other action.  But it's hard
			// to imagine someone ever reassigning the naked VK_TAB key (i.e. with no modifiers).
			// UPDATE: The new "ignore" method should prevent that.  Or in the case of low-level hook:
			// keystrokes sent by our own app by default will not fire hotkeys.  UPDATE: Even though
			// the LL hook will have suppressed this key, it seems that the OS's alt-tab menu uses
			// some weird method (apparently not GetAsyncState(), because then our attempt to put
			// it up would fail) to determine whether the shift-key is down, so we need to still do this:
			else if (hotkey_id_temp == HOTKEY_ID_ALT_TAB) // i.e. it's not shift-alt-tab
			{
				// Force it to be alt-tab as the user intended.
				if ((aVK == VK_LSHIFT || aVK == VK_RSHIFT) && !aKeyUp)  // Needed.  See above comments. vk == VK_SHIFT not needed.
					// If a shift key is the suffix key, this must be done every time,
					// not just the first:
					KeyEvent(KEYUP, aVK, aSC);
				// UPDATE: Don't do "else" because sometimes the opposite key may be down, so the
				// below needs to be unconditional:
				//else

				// In the below cases, it's not necessary to put the shift key back down because
				// the alt-tab menu only disappears after the prefix key has been released (and
				// it's not realistic that a user would try to trigger another hotkey while the
				// alt-tab menu is visible).  In other words, the user will be releasing the
				// shift key anyway as part of the alt-tab process, so it's not necessary to put
				// it back down for the user here (the shift stays in effect as a prefix for us
				// here because it's sent as an ignore event -- but the prefix will be correctly
				// canceled when the user releases the shift key).
				if (g_modifiersLR_logical & MOD_LSHIFT)
					KeyEvent(KEYUP, VK_LSHIFT);
				if (g_modifiersLR_logical & MOD_RSHIFT)
					KeyEvent(KEYUP, VK_RSHIFT);
			}
			// Any down control key prevents alt-tab from working.  This is similar to
			// what's done for the shift-key above, so see those comments for details.
			// Note: Since this is the low-level hook, the current OS must be something
			// beyond other than Win9x, so there's no need to conditionally send
			// VK_CONTROL instead of the left/right specific key of the pair:
			if (g_modifiersLR_logical & MOD_LCONTROL)
				KeyEvent(KEYUP, VK_LCONTROL);
			if (g_modifiersLR_logical & MOD_RCONTROL)
				KeyEvent(KEYUP, VK_RCONTROL);

			KeyEvent(KEYDOWNANDUP, VK_TAB);

			if (hotkey_id_temp == HOTKEY_ID_ALT_TAB_SHIFT && pPrefixKey->it_put_shift_down
				&& ((aVK >= VK_NUMPAD0 && aVK <= VK_NUMPAD9) || aVK == VK_DECIMAL)) // dual-state numpad key.
			{
				// In this case, if there is a numpad key involved, it's best to put the shift key
				// back up in between every alt-tab to avoid problems caused due to the fact that
				// the shift key being down would CHANGE the VK being received when the key is
				// released (due to the fact that SHIFT temporarily disables numlock).
				KeyEvent(KEYUP, VK_SHIFT);
				pPrefixKey->it_put_shift_down = false;  // Reset for next time since we put it back up already.
			}
			pKeyHistoryCurr->event_type = 'h'; // h = hook hotkey (not one registered with RegisterHotkey)
			if (!aKeyUp)
				this_key.hotkey_down_was_suppressed = true;
			return SuppressThisKey;
		} // end of alt-tab section.
		// Since abov didn't return, this isn't a prefix-triggered alt-tab action (though it might be
		// a non-prefix alt-tab action, which is handled later below).
	} // end of section that searches for a suffix modified by the prefix that's currently held down.

	if (hotkey_id_with_flags == HOTKEY_ID_INVALID)  // Since above didn't find a hotkey, check if modifiers+this_key qualifies a firing.
	{
		modifiersLRnew = g_modifiersLR_logical_non_ignored;
		if (this_key.as_modifiersLR)
			// Hotkeys are not defined to modify themselves, so look for a match accordingly.
			modifiersLRnew &= ~this_key.as_modifiersLR;
		// Check hook type too in case a script every explicitly specifies scan code zero as a hotkey:
		hotkey_id_with_flags = (aHook == g_KeybdHook && sc_takes_precedence)
			? Kscm(modifiersLRnew, aSC) : Kvkm(modifiersLRnew, aVK);
		// Bug fix for v1.0.20: The below second attempt is no longer made if the current keystroke
		// is a tab-down/up  This is because doing so causes any naked TAB that has been defined as
		// a hook hotkey to incorrectly fire when the user holds down ALT and presses tab two or more
		// times to advance through the alt-tab menu.  Here is the sequence:
		// $TAB is defined as a hotkey in the script.
		// User holds down ALT and presses TAB two or more times.
		// The Alt-tab menu becomes visible on the first TAB keystroke.
		// The $TAB hotkey fires on the second keystroke because of the below (now-fixed) logic.
		// By the way, the overall idea behind the below might be considered faulty because
		// you could argue that non-modified hotkeys should never be allowed to fire while ALT is
		// down just because the alt-tab menu is visible.  However, it seems justified because
		// the benefit (which I believe was originally and particularly that an unmodified mouse button
		// or wheel hotkey could be used to advance through the menu even though ALT is artificially
		// down due to support displaying the menu) outweighs the cost, which seems low since
		// it would be rare that anyone would press another hotkey while they are navigating through
		// the Alt-Tab menu.
		if (hotkey_id_with_flags == HOTKEY_ID_INVALID && sAltTabMenuIsVisible && aVK != VK_TAB)
		{
			// Try again, this time without the ALT key in case the user is trying to
			// activate an alt-tab related key (i.e. a special hotkey action such as AltTab
			// that relies on the Alt key being logically but not physically down).
			modifiersLRnew &= ~(MOD_LALT | MOD_RALT);
			hotkey_id_with_flags = (aHook == g_KeybdHook && sc_takes_precedence)
				? Kscm(modifiersLRnew, aSC) : Kvkm(modifiersLRnew, aVK);
			// Fix for v1.0.28: If the ID isn't an alt-tab type, don't consider it to be valid.
			// Someone pointed out that pressing Alt-Tab and then pressing ESC while still holding
			// down ALT fired the ~Esc hotkey even when it should just dismiss the alt-tab menu.
			// Note: Both of the below checks must be done because the high-order bits of the
			// hotkey_id_with_flags might be set to indicate no-suppress, etc:
			hotkey_id_temp = hotkey_id_with_flags & HOTKEY_ID_MASK;
			if (!IS_ALT_TAB(hotkey_id_temp))
				hotkey_id_with_flags = HOTKEY_ID_INVALID; // Since it's not an Alt-tab action, don't fire this hotkey.
		}

		if (hotkey_id_with_flags & HOTKEY_KEY_UP)
		{
			if (!aKeyUp) // Key-up hotkey but the event is a down-event.
			{
				this_key.hotkey_to_fire_upon_release = hotkey_id_with_flags; // Seem comments above in other occurrences of this line.
				hotkey_id_with_flags = HOTKEY_ID_INVALID;
			}
			//else hotkey_id_with_flags contains the up-hotkey that is now eligible for firing.
		}
		else if (hotkey_id_with_flags != HOTKEY_ID_INVALID) // hotkey_id_with_flags is a valid key-down hotkey.
		{
			hotkey_id_temp = hotkey_id_with_flags & HOTKEY_ID_MASK;
			if (aKeyUp)
			{
				// Even though the key is being released, a hotkey should fire unconditionally because
				// the only way we can reach this exact point for a non-key-up hotkey is when it fell
				// through from Case #3, in which case this hotkey_id_with_flags is implicitly a key-up
				// hotkey if there is no actual explicit key-up hotkey for it.  UPDATE: It is now possible
				// to fall through from Case #2, so that is checked below.
				if (hotkey_id_temp < Hotkey::sHotkeyCount && hotkey_up[hotkey_id_temp] != HOTKEY_ID_INVALID) // Relies on short-circuit boolean order.
					hotkey_id_with_flags = hotkey_up[hotkey_id_temp];
				else // Leave it at its former value unless case#2.  See comments above and below.
					// Fix for v1.0.44.09: Since no key-up counterpart was found above (either in hotkey_up[]
					// or via the HOTKEY_KEY_UP flag), don't fire this hotkey when it fell through from Case #2.
					// This prevents a hotkey like $^b from firing TWICE (once on down and again on up) when a
					// key-up hotkey with different modifiers also exists, such as "#b" and "#b up" existing with $^b.
					if (fell_through_from_case2)
						hotkey_id_with_flags = HOTKEY_ID_INVALID;
			}
			else // hotkey_id_with_flags contains the down-hotkey that is now eligible for firing. But check if there's an up-event to queue up for later.
				if (hotkey_id_temp < Hotkey::sHotkeyCount)
					this_key.hotkey_to_fire_upon_release = hotkey_up[hotkey_id_temp];
		}

		// Check hotkey_id_with_flags again now that the above possibly changed it:
		if (hotkey_id_with_flags == HOTKEY_ID_INVALID)
		{
			// Even though at this point this_key is a valid suffix, no actionable ModifierVK/SC
			// or modifiers were pressed down, so just let the system process this normally
			// (except if it's a toggleable key).  This case occurs whenever a suffix key (which
			// is also a prefix) is released but the key isn't configured to perform any action
			// upon key-release.  Currently, I think the only way a key-up event will result
			// in a hotkey action is for the release of a naked/modifierless prefix key.
			// Example of a configuration that would result in this case whenever Rshift alone
			// is pressed then released:
			// RControl & RShift = Alt-Tab
			// RShift & RControl = Shift-Alt-Tab
			if (aKeyUp)
				// These sequence is basically the same as the one used in Case #3
				// when a prefix key that isn't a suffix failed to modify anything
				// and was then released, so consider any modifications made here
				// or there for inclusion in the other one.  UPDATE: Since
				// the previous sentence is a bit obsolete, describe this better:
				// If it's a toggleable key that the user wants to allow to be
				// toggled, just allow this up-event to go through because the
				// previous down-event for it (in its role as a prefix) would not
				// have been suppressed:
				// NO_SUPPRESS_PREFIX can occur if it fell through from Case #3 but the right
				// modifier keys aren't down to have triggered a key-up hotkey:
				return (this_key.as_modifiersLR || (this_key.no_suppress & NO_SUPPRESS_PREFIX)
					// The following line was added for v1.0.37.02 to take into account key-up hotkeys,
					// the release of which should never be suppressed if it didn't actually fire the
					// up-hotkey (due to the wrong modifiers being down):
					|| !this_key.used_as_prefix
					// The order on this line important; it relies on short-circuit boolean:
					|| this_toggle_key_can_be_toggled) ? AllowKeyToGoToSystem : SuppressThisKey;
				// v1.0.37.02: Added !this_key.used_as_prefix for mouse hook too (see comment above).

			// For execution to have reached this point, the following must be true:
			// 1) aKeyUp==false
			// 2) this_key must be both a prefix and suffix, but be acting in its capacity as a suffix.
			// 3) No hotkey is eligible to fire.
			// Since no hotkey action will fire, and since this_key wasn't used as a prefix, I think that
			// must mean that not all of the required modifiers aren't present.  For example:
			// a & b::Run calc
			// LShift & a:: Run Notepad
			// In that case, if the 'a' key is pressed and released by itself, perhaps its native
			// function should be performed by suppressing this key-up event, replacing it with a
			// down and up of our own.  However, it seems better not to do this, for now, since this
			// is really just a subset of allowing all prefixes to perform their native functions
			// upon key-release their value of was_just_used is false, which is probably
			// a bad idea in many cases (e.g. if user configures VK_VOLUME_MUTE button to be a
			// prefix, it might be undesirable for the volume to be muted if the button is pressed
			// but the user changes his mind and doesn't use it to modify anything, so just releases
			// it (at least it seems that I do this).  In any case, this default behavior can be
			// changed by explicitly configuring 'a', in the example above, to be "Send, a".
			// Here's a more complete example:
			// a & b:: Run Notepad
			// LControl & a:: Run Calc
			// a::Send a
			// So in summary, by default a prefix key's native function is always suppressed except if it's
			// a toggleable key such as num/caps/scroll-lock.
			if (this_key.hotkey_to_fire_upon_release == HOTKEY_ID_INVALID)
				return AllowKeyToGoToSystem;
			// Otherwise (v1.0.44): Since there is a hotkey to fire upon release (somewhat rare under these conditions),
			// check if any of its criteria will allow it to fire, and if so whether that variant is non-suppressed.
			// If it is, this down-even should be non-suppressed too (for symmetry).  This check isn't 100% reliable
			// because the active/existing windows checked by the criteria might change before the user actually
			// releases the key, but there doesn't seem any way around that.
			Hotkey::CriterionFiringIsCertain(this_key.hotkey_to_fire_upon_release // firing_is_certain==false under these conditions, so no need to check it.
				, true  // Always a key-up since it's will fire upon release.
				, this_key.no_suppress // Unused and won't be altered because above is "true".
				, fire_with_no_suppress, NULL); // fire_with_no_suppress is the value we really need to get back from it.
			return fire_with_no_suppress ? AllowKeyToGoToSystem : SuppressThisKey;
		}
		//else an eligible hotkey was found.
	} // Final attempt to find hotkey based on suffix have the right combo of modifiers.

	// Since above didn't return, hotkey_id_with_flags is now a valid hotkey.  The only thing that can
	// stop it from firing now is CriterionFiringIsCertain().

	// v1.0.41: Below should be done prior to the next section's "return AllowKeyToGoToSystem" so that the
	// NO_SUPPRESS_NEXT_UP_EVENT ticket is used up rather than staying around to possibly take effect for
	// a future key-up for which it wasn't intended.
	// Handling for NO_SUPPRESS_NEXT_UP_EVENT was added because it seems more correct that key-up
	// hotkeys should obey NO_SUPPRESS_NEXT_UP_EVENT too.  The absence of this might have been inconsequential
	// due to other safety/redundancies; but it seems more maintainable this way.
	if ((this_key.no_suppress & NO_SUPPRESS_NEXT_UP_EVENT) && aKeyUp)
	{
		fire_with_no_suppress = true; // In spite of this being a key-up, there may be circumstances in which this was already true due to action above.
		this_key.no_suppress &= ~NO_SUPPRESS_NEXT_UP_EVENT; // This ticket has been used up, so remove it.
	}

	// v1.0.41: This must be done prior to the setting of sDisguiseNextLWinUp and similar items below.
	hotkey_id_temp = hotkey_id_with_flags & HOTKEY_ID_MASK;
	if (hotkey_id_temp < Hotkey::sHotkeyCount // i.e. don't call the below for Alt-tab hotkeys and similar.
		&& !firing_is_certain  // i.e. CriterionFiringIsCertain() wasn't already called earlier.
		&& !Hotkey::CriterionFiringIsCertain(hotkey_id_with_flags, aKeyUp, this_key.no_suppress, fire_with_no_suppress, &pKeyHistoryCurr->event_type))
		return AllowKeyToGoToSystem;
	hotkey_id_temp = hotkey_id_with_flags & HOTKEY_ID_MASK; // Update in case CriterionFiringIsCertain() changed the naked/raw ID.

	// Now above has ensured that everything is in place for an action to be performed.
	// Determine the final ID at this late stage to improve maintainability:
	HotkeyIDType hotkey_id_to_fire = hotkey_id_temp;

	// If only a windows key was held down (and no other modifiers) to activate this hotkey,
	// suppress the next win-up event so that the Start Menu won't appear (if other modifiers are
	// present, there's no need to do this because the Start Menu doesn't appear, at least on WinXP).
	// The appearance of the Start Menu would be caused by the fact that the hotkey's suffix key
	// was suppressed, therefore the OS doesn't see that the WIN key "modified" anything while
	// it was held down.  Note that if the WIN key is auto-repeating due to the user having held
	// it down long enough, when the user presses the hotkey's suffix key, the auto-repeating
	// stops, probably because auto-repeat is a very low-level feature.  It's also interesting
	// that unlike non-modifier keys such as letters, the auto-repeat does not resume after
	// the user releases the suffix, even if the WIN key is kept held down for a long time.
	// When the user finally releases the WIN key, that release will be disguised if called
	// for by the logic below.
	if (aHook == g_KeybdHook)
	{
		if (!(g_modifiersLR_logical & ~(MOD_LWIN | MOD_RWIN))) // Only lwin, rwin, both, or neither are currently down.
		{
			// If it's used as a prefix, there's no need (and it would probably break something)
			// to disguise the key this way since the prefix-handling logic already does that
			// whenever necessary:
			if ((g_modifiersLR_logical & MOD_LWIN) && !kvk[VK_LWIN].used_as_prefix)
				sDisguiseNextLWinUp = true;
			if ((g_modifiersLR_logical & MOD_RWIN) && !kvk[VK_RWIN].used_as_prefix)
				sDisguiseNextRWinUp = true;
		}
	}
	else // Mouse hook
	{
		// The mouse hook requires suppression in more situations than the keyboard because the
		// OS does not consider LWin/RWin to have modified a mouse button, only a keyboard key.
		// Thus, the Start Menu appears in the following cases if the user releases LWin/RWin
		// *after* the other modifier:
		// #+MButton::return
		// #^MButton::return
		if (!(g_modifiersLR_logical & (MOD_LALT|MOD_RALT)))
		{
			if ((g_modifiersLR_logical & MOD_LWIN) && !kvk[VK_LWIN].used_as_prefix)
				sDisguiseNextLWinUp = true;
			else if ((g_modifiersLR_logical & MOD_RWIN) && !kvk[VK_RWIN].used_as_prefix)
				sDisguiseNextRWinUp = true;
		}
		// An earlier stage has ensured that the keyboard hook is installed for the above, because the sending
		// of CTRL directly (here) would otherwise not suppress the Start Menu for LWin/RWin (though it does
		// supress menu bar activation for ALT hotkeys, as described below).
	}

	// For maximum reliability on the maximum range of systems, it seems best to do the above
	// for ALT keys also, to prevent them from invoking the icon menu or menu bar of the
	// foreground window (rarer than the Start Menu problem, above, I think).
	// Update for v1.0.25: This is usually only necessary for hotkeys whose only modifier is ALT.
	// For example, Shift-Alt hotkeys do not need it if Shift is pressed after Alt because Alt
	// "modified" the shift so the OS knows it's not a naked ALT press to activate the menu bar.
	// Conversely, if Shift is pressed prior to Alt, but released before Alt, I think the shift-up
	// counts as a "modification" and the same rule applies.  However, if shift is released after Alt,
	// that would activate the menu bar unless the ALT key is disguised below.  This issue does
	// not apply to the WIN key above because apparently it is disguised automatically
	// whenever some other modifier was involved with it in any way and at any time during the
	// keystrokes that comprise the hotkey.
	if ((g_modifiersLR_logical & MOD_LALT) && !kvk[VK_LMENU].used_as_prefix)
	{
		if (g_KeybdHook)
			sDisguiseNextLAltUp = true;
		else
			// Since no keyboard hook, no point in setting the variable because it would never be acted up.
			// Instead, disguise the key now with a CTRL keystroke. Note that this is not done for
			// mouse buttons that use the WIN key as a prefix because it does not work reliably for them
			// (i.e. sometimes the Start Menu appears, even if two CTRL keystrokes are sent rather than one).
			// Therefore, as of v1.0.25.05, mouse button hotkeys that use only the WIN key as a modifier cause
			// the keyboard hook to be installed.  This determination is made during the hotkey loading stage.
			KeyEvent(KEYDOWNANDUP, g_MenuMaskKey);
	}
	else if ((g_modifiersLR_logical & MOD_RALT) && !kvk[VK_RMENU].used_as_prefix && !ActiveWindowLayoutHasAltGr()) // If RAlt==AltGr, it should never need disguising.
	{
		// The two else if's above: If it's used as a prefix, there's no need (and it would probably break something)
		// to disguise the key this way since the prefix-handling logic already does that whenever necessary.
		if (g_KeybdHook)
			sDisguiseNextRAltUp = true;
		else
			KeyEvent(KEYDOWNANDUP, g_MenuMaskKey);
	}

	switch (hotkey_id_to_fire)
	{
		case HOTKEY_ID_ALT_TAB_MENU_DISMISS: // This case must occur before HOTKEY_ID_ALT_TAB_MENU due to non-break.
			if (!sAltTabMenuIsVisible)
				// Even if the menu really is displayed by other means, we can't easily detect it
				// because it's not a real window?
				return AllowKeyToGoToSystem;  // Let the key do its native function.
			// else fall through to the next case.
		case HOTKEY_ID_ALT_TAB_MENU:  // These cases must occur before the Alt-tab ones due to conditional break.
		case HOTKEY_ID_ALT_TAB_AND_MENU:
		{
			vk_type which_alt_down = 0;
			if (g_modifiersLR_logical & MOD_LALT)
				which_alt_down = VK_LMENU;
			else if (g_modifiersLR_logical & MOD_RALT)
				which_alt_down = VK_RMENU;

			if (sAltTabMenuIsVisible)  // Can be true even if which_alt_down is zero.
			{
				if (hotkey_id_to_fire != HOTKEY_ID_ALT_TAB_AND_MENU) // then it is MENU or DISMISS.
				{
					// Since it is possible for the menu to be visible when neither ALT
					// key is down, always send an alt-up event if one isn't down
					// so that the menu is dismissed as intended:
					KeyEvent(KEYUP, which_alt_down ? which_alt_down : VK_MENU);
					if (this_key.as_modifiersLR && aVK != VK_LWIN && aVK != VK_RWIN && !aKeyUp)
						// Something strange seems to happen with the foreground app
						// thinking the modifier is still down (even though it was suppressed
						// entirely [confirmed!]).  For example, if the script contains
						// the line "lshift::AltTabMenu", pressing lshift twice would
						// otherwise cause the newly-activated app to think the shift
						// key is down.  Sending an extra UP here seems to fix that,
						// hopefully without breaking anything else.  Note: It's not
						// done for Lwin/Rwin because most (all?) apps don't care whether
						// LWin/RWin is down, and sending an up event might risk triggering
						// the start menu in certain hotkey configurations.  This policy
						// might not be the right one for everyone, however:
						KeyEvent(KEYUP, aVK); // Can't send sc here since it's not defined for the mouse hook.
					sAltTabMenuIsVisible = false;
					break;
				}
				// else HOTKEY_ID_ALT_TAB_AND_MENU, do nothing (don't break) because we want
				// the switch to fall through to the Alt-Tab case.
			}
			else // alt-tab menu is not visible
			{
				// Unlike CONTROL, SHIFT, AND ALT, the LWIN/RWIN keys don't seem to need any
				// special handling to make them work with the alt-tab features.

				bool vk_is_alt = aVK == VK_LMENU || aVK == VK_RMENU;  // Tranlated & no longer needed: || vk == VK_MENU;
				bool vk_is_shift = aVK == VK_LSHIFT || aVK == VK_RSHIFT;  // || vk == VK_SHIFT;
				bool vk_is_control = aVK == VK_LCONTROL || aVK == VK_RCONTROL;  // || vk == VK_CONTROL;

				vk_type which_shift_down = 0;
				if (g_modifiersLR_logical & MOD_LSHIFT)
					which_shift_down = VK_LSHIFT;
				else if (g_modifiersLR_logical & MOD_RSHIFT)
					which_shift_down = VK_RSHIFT;
				else if (!aKeyUp && vk_is_shift)
					which_shift_down = aVK;

				vk_type which_control_down = 0;
				if (g_modifiersLR_logical & MOD_LCONTROL)
					which_control_down = VK_LCONTROL;
				else if (g_modifiersLR_logical & MOD_RCONTROL)
					which_control_down = VK_RCONTROL;
				else if (!aKeyUp && vk_is_control)
					which_control_down = aVK;

				bool shift_put_up = false;
				if (which_shift_down)
				{
					KeyEvent(KEYUP, which_shift_down);
					shift_put_up = true;
				}

				bool control_put_up = false;
				if (which_control_down)
				{
					// In this case, the control key must be put up because the OS, at least
					// WinXP, knows the control key is down even though the down event was
					// suppressed by the hook.  So put it up and leave it up, because putting
					// it back down would cause it to be down even after the user releases
					// it (since the up-event of a hotkey is also suppressed):
					KeyEvent(KEYUP, which_control_down);
					control_put_up = true;
				}

				// Alt-tab menu is not visible, or was not made visible by us.  In either case,
				// try to make sure it's displayed:
				// Don't put alt down if it's already down, it might mess up cases where the
				// ALT key itself is assigned to be one of the alt-tab actions:
				if (vk_is_alt)
					if (aKeyUp)
						// The system won't see it as down for the purpose of alt-tab, so remove this
						// modifier from consideration.  This is necessary to allow something like this
						// to work:
						// LAlt & WheelDown::AltTab
						// LAlt::AltTabMenu   ; Since LAlt is a prefix key above, it will be a key-up hotkey here.
						which_alt_down = 0;
					else // Because there hasn't been a chance to update g_modifiersLR_logical yet:
						which_alt_down = aVK;
				if (!which_alt_down)
					KeyEvent(KEYDOWN, VK_MENU); // Use the generic/neutral ALT key so it works with Win9x.

				KeyEvent(KEYDOWNANDUP, VK_TAB); // v1.0.28: KEYDOWNANDUP vs. KEYDOWN.

				// Only put it put it back down if it wasn't the hotkey itself, because
				// the system would never have known it was down because the down-event
				// on the hotkey would have been suppressed.  And since the up-event
				// will also be suppressed, putting it down like this would result in
				// it being permanently down even after the user releases the key!:
				if (shift_put_up && !vk_is_shift) // Must do this regardless of the value of aKeyUp.
					KeyEvent(KEYDOWN, which_shift_down);
				
				// Update: Can't do this one because going down on control will instantly
				// dismiss the alt-tab menu, which we don't want if we're here.
				//if (control_put_up && !vk_is_control) // Must do this regardless of the value of aKeyUp.
				//	KeyEvent(KEYDOWN, which_control_down);

				// At this point, the alt-tab menu has displayed and advanced by one icon
				// (to the next window in the z-order).  Rather than sending a shift-tab to
				// go back to the first icon in the menu, it seems best to leave it where
				// it is because usually the user will want to go forward at least one item.
				// Going backward through the menu is a lot more rare for most people.
				sAltTabMenuIsVisible = true;
				break;
			}
		}
		case HOTKEY_ID_ALT_TAB:
		case HOTKEY_ID_ALT_TAB_SHIFT:
		{
			// Since we're here, this ALT-TAB hotkey didn't have a prefix or it would have
			// already been handled and we would have returned above.  Therefore, this
			// hotkey is defined as taking effect only if the alt-tab menu is currently
			// displayed, otherwise it will just be passed through to perform it's native
			// function.  Example:
			// MButton::AltTabMenu
			// WheelDown::AltTab     ; But if the menu is displayed, the wheel will function normally.
			// WheelUp::ShiftAltTab  ; But if the menu is displayed, the wheel will function normally.
			if (!sAltTabMenuIsVisible)
				// Even if the menu really is displayed by other means, we can't easily detect it
				// because it's not a real window?
				return AllowKeyToGoToSystem;

			// Unlike CONTROL, SHIFT, AND ALT, the LWIN/RWIN keys don't seem to need any
			// special handling to make them work with the alt-tab features.

			// Must do this to prevent interference with Alt-tab when these keys
			// are used to do the navigation.  Don't put any of these back down
			// after putting them up since that would probably cause them to become
			// stuck down due to the fact that the user's physical release of the
			// key will be suppressed (since it's a hotkey):
			if (!aKeyUp && (aVK == VK_LCONTROL || aVK == VK_RCONTROL || aVK == VK_LSHIFT || aVK == VK_RSHIFT))
				// Don't do the ALT key because it causes more problems than it solves
				// (in fact, it might not solve any at all).
				KeyEvent(KEYUP, aVK); // Can't send sc here since it's not defined for the mouse hook.

			// Even when the menu is visible, it's possible that neither of the ALT keys
			// is down, at least under XP (probably NT and 2k also).  Not sure about Win9x:
			if (   !(g_modifiersLR_logical & (MOD_LALT | MOD_RALT))   // Neither ALT key is down 
				|| (aKeyUp && (aVK == VK_LMENU || aVK == VK_RMENU))   ) // Or the suffix key for Alt-tab *is* an ALT key and it's being released: must push ALT down for upcoming TAB to work.
				KeyEvent(KEYDOWN, VK_MENU);
				// And never put it back up because that would dismiss the menu.
			// Otherwise, use keystrokes to navigate through the menu:
			bool shift_put_down = false;
			if (hotkey_id_to_fire == HOTKEY_ID_ALT_TAB_SHIFT && !(g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT))) // Neither SHIFT key is down.
			{
				KeyEvent(KEYDOWN, VK_SHIFT);
				shift_put_down = true;
			}
			KeyEvent(KEYDOWNANDUP, VK_TAB);
			if (shift_put_down)
				KeyEvent(KEYUP, VK_SHIFT);
			break;
		}
		default:
			// Notify the main thread (via its main window) of which hotkey has been pressed.
			// Post the message rather than sending it, because Send would need
			// SendMessageTimeout(), which is undesirable because the whole point of
			// making this hook thread separate from the main thread is to have it be
			// maximally responsive (especially to prevent mouse cursor lag).
			// v1.0.42: The hotkey variant is not passed via the message below because
			// upon receipt of the message, the variant is recalculated in case conditions
			// have changed between msg-post and arrival.  See comments in the message loop for details.
			// v1.0.42.01: the message is now posted at the latest possible moment to avoid
			// situations in which the message arrives and is processed by the main thread
			// before we finish processing the hotkey's final keystroke here.  This avoids
			// problems with a script calling GetKeyState() and getting an inaccurate value
			// because the hook thread is either pre-empted or is running in parallel
			// (multiprocessor) and hasn't yet returned 1 or 0 to determine whether the final
			// keystroke is suppressed or passed through to the active window.  Similarly, this solves
			// the fact that previously, g_PhysicalKeyState was not updated for modifier keys until after
			// the hotkey message was posted, which on some PCs caused the hotkey subroutine to see
			// the wrong key state via KeyWait (which defaults to detecting the physical key state).
			// For example, the following hotkeys would be a problem on certain PCs, presumably due to
			// split-second timing where the hook thread gets preempted and the main thread gets a
			// timeslice that allows it to launch a script subroutine before the hook can get
			// another timeslice to finish up:
			//$LAlt::
			//if not GetKeyState("LAlt", "P")
			//	ToolTip `nProblem 1`n
			//return
			//
			//~LControl::
			//if not (DllCall("GetAsyncKeyState", int, 0xA2) & 0x8000)
			//    ToolTip `nProblem 2`n
			//return
			hotkey_id_to_post = hotkey_id_to_fire; // Set this only when it is certain that this ID should be sent to the main thread via msg.
	}

	pKeyHistoryCurr->event_type = 'h'; // h = hook hotkey (not one registered with RegisterHotkey)

	if (this_toggle_key_can_be_toggled && aKeyUp && this_key.used_as_prefix)
	{
		// In this case, since all the above conditions are true, the key-down
		// event for this key-up (which fired a hotkey) would not have been
		// suppressed.  Thus, we should toggle the state of the key back
		// the what it was before the user pressed it (due to the policy that
		// the natural function of a key should never take effect when that
		// key is used as a hotkey suffix).  You could argue that instead
		// of doing this, we should change *pForceToggle's value to make the
		// key untoggleable whenever it's both a prefix and a naked
		// (key-up triggered) suffix.  However, this isn't too much harder
		// and has the added benefit of allowing the key to be toggled if
		// a modifier is held down before it (e.g. alt-CapsLock would then
		// be able to toggle the CapsLock key):
		KEYEVENT_PHYS(KEYUP, aVK, aSC); // Mark it as physical for any other hook instances.
		KeyEvent(KEYDOWNANDUP, aVK, aSC);
		return SuppressThisKey;
	}

	if (aKeyUp)
	{
		if (this_key.as_modifiersLR)
			// Since this hotkey is fired on a key-up event, and since it's a modifier, must
			// not suppress the key because otherwise the system's state for this modifier
			// key would be stuck down due to the fact that the previous down-event for this
			// key (which is presumably a prefix *and* a suffix) was not suppressed. UPDATE:
			// For v1.0.28, if the new field hotkey_down_was_suppressed is true, also suppress
			// this up event, one purpose of which is to allow a pair of remappings such
			// as the following to display the Start Menu (because otherwise the non-suppressed
			// Alt key events would prevent it):
			// *LAlt up::Send {LWin up}
			// *LAlt::Send {LWin down}
			return this_key.hotkey_down_was_suppressed ? SuppressThisKey : AllowKeyToGoToSystemButDisguiseWinAlt;

		// Adding the check of this_key.no_suppress fixes a pair of hotkeys such as the following by forcing
		// tilde in front of the other:
		// ~a & b::
		// a::
		// However, the converse is not done because there is another workaround for that. Search
		// on "send a down event to make up" to find that section.
		// Fixed in v1.0.41 to use this_key rather than referring to both kvk[aVK] and ksc[aSC].
		// Although this bug had no known consequences, I suspect there were some obscure ones.
		if (fire_with_no_suppress || (this_key.no_suppress & NO_SUPPRESS_PREFIX)) // Plus we know it's not a modifier since otherwise it would've returned above.
		{
			// Currently not supporting the mouse buttons for the above method, because KeyEvent()
			// doesn't support the translation of a mouse-VK into a mouse_event() call.
			// Such a thing might not work anyway because the hook probably received extra
			// info such as the location where the mouse click should occur and other things.
			// That info plus anything else relevant in MSLLHOOKSTRUCT would have to be
			// translated into the correct info for a call to mouse_event().
			if (aHook == g_MouseHook)
				return AllowKeyToGoToSystem;
			// Otherwise, our caller is the keyboard hook.
			// Since this hotkey is firing on key-up but the user specified not to suppress its native
			// function, send a down event to make up for the fact that the original down event was
			// suppressed (since key-up hotkeys' down events are always suppressed because they
			// are also prefix keys by definition).  UPDATE: Now that it is possible for a prefix key
			// to be non-suppressed, this is done only if the prior down event wasn't suppressed.
			// Note that for a pair of hotkeys such as:
			// *capslock::Send {Ctrl Down}
			// *~capslock up:: Send {Ctrl Up}  ; Specify tilde to allow caps lock to be toggled upon release.
			// ... the following key history is produced (see note):
			//14  03A	h	d	3.46	Caps Lock   	
			//A2  01D	i	d	0.00	Ctrl        	
			//14  03A	h	u	0.10	Caps Lock   	
			//14  03A	i	d	0.00	Caps Lock    <<< This actually came before the prior due to re-entrancy.
			//A2  01D	i	u	0.00	Ctrl        	
			// Can't use this_toggle_key_can_be_toggled in this case. Relies on short-circuit boolean order:
			bool suppress_to_prevent_toggle = this_key.pForceToggle && *this_key.pForceToggle != NEUTRAL;
			// The following isn't checked as part of the above because this_key.was_just_used would
			// never be true with hotkeys such as the Capslock pair shown above. That's because
			// Caplock isn't a prefix in that case, it's just a suffix. Even if it were a prefix, it would
			// never reach this point in the execution because places higher above return early if the value of
			// this_key.was_just_used is AS_PREFIX/AS_PREFIX_FOR_HOTKEY.
			// Used as either a prefix for a hotkey or just a plain modifier for another key.
			// ... && (*this_key.pForceToggle != NEUTRAL || this_key.was_just_used);
			if (this_key.hotkey_down_was_suppressed // Down was suppressed.
				&& !is_explicit_key_up_hotkey // v1.0.36.02: Prevents a hotkey such as "~5 up::" from generating double characters, regardless of whether it's paired with a "~5::" hotkey.
				&& !suppress_to_prevent_toggle) // Mouse vs. keybd hook was already checked higher above.
				KeyEvent(KEYDOWN, aVK, aSC); // Substitute this to make up for the suppression (a check higher above has already determined that no_supress==true).
				// Now allow the up-event to go through.  The DOWN should always wind up taking effect
				// before the UP because the above should already have "finished" by now, since
				// it resulted in a recursive call to this function (using our hook-thread
				// rather than our main thread or some other thread):
			return suppress_to_prevent_toggle ? SuppressThisKey : AllowKeyToGoToSystem;
		} // No suppression.
	}
	else // Key Down
	{
		// Do this only for DOWN (not UP) events that triggered an action:
		this_key.down_performed_action = true;

		// Fix for v1.0.26: The below is now done only if pPrefixKey is a modifier.
		// This is because otherwise, a prefix key would not perform its normal function when it
		// modifies a hook suffix that doesn't belong to it.  For example, if "Capslock & A"
		// and "MButton" are both hotkeys, clicking the middle button while holding down Capslock
		// should turn CapsLock On or Off as expected.  This should be okay because
		// AS_PREFIX_FOR_HOTKEY is set in other places it needs to be except the following,
		// which are still correctly handled here:
		// 1) Fall through from Case #1, in which case this hotkey is one that uses normal
		//    modifiers (i.e. not something like "Capslock & A").
		// 2) Any hotkey similar to Case #1 that isn't actually handled by Case #1.
		// In both of the above situations, if pPrefixKey is not a modifier, it can't be
		// part of the reason for this hotkey firing (because both of the above do not
		// consider the state of any prefix keys other than those that happen to be modifiers).
		// In other words, pPrefixKey is down only incidentally and has nothing to do with
		// triggering this hotkey.
		// Update pPrefixKey in case the currently-down prefix key is both a modifier
		// and a normal prefix key (in which case it isn't stored in this_key's array
		// of VK and SC prefixes, so this value wouldn't have yet been set).
		// Update: The below is done even if pPrefixKey != &this_key, which happens
		// when we reached this point after having fallen through from Case #1 above.
		// The reason for this is that we just fired a hotkey action for this key,
		// so we don't want it's action to fire again upon key-up:
		if (pPrefixKey && pPrefixKey->as_modifiersLR)
			pPrefixKey->was_just_used = AS_PREFIX_FOR_HOTKEY;

		if (fire_with_no_suppress)
		{
			// Since this hotkey is firing on key-down but the user specified not to suppress its native
			// function, substitute an DOWN+UP pair of events for this event, since we want the
			// DOWN to precede the UP.  It's necessary to send the UP because the user's physical UP
			// will be suppressed automatically when this function is called for that event.
			// UPDATE: The below method causes side-effects due to the fact that it is simulated
			// input vs. physical input, e.g. when used with the Input command, which distinguishes
			// between "ignored" and physical input.  Therefore, let this down event pass through
			// and set things up so that the corresponding up-event is also not suppressed:
			//KeyEvent(KEYDOWNANDUP, aVK, aSC);
			// No longer relevant due to the above change:
			// Now let it just fall through to suppress this down event, because we can't use it
			// since doing so would result in the UP event having preceded the DOWN, which would
			// be the wrong order.
			this_key.no_suppress |= NO_SUPPRESS_NEXT_UP_EVENT;
			return AllowKeyToGoToSystem;
		}
		else if (aVK == VK_LMENU || aVK == VK_RMENU)
		{
			// Fix for v1.0.34: For some reason, the release of the ALT key here causes the Start Menu
			// to appear instantly for the hotkey #LAlt (and probably #RAlt), even when the hotkey does
			// nothing other than return.  This seems like an OS quirk since it doesn't conform to any
			// known Start Menu activation sequence.  This happens only when neither Shift nor Control is
			// down.  To work around it, send the menu-suppressing Control keystroke here.  Another one
			// will probably be sent later when the WIN key is physically released, but it seems best
			// for simplicity and avoidance of side-effects not to make this one prevent that one.
			if (   (g_modifiersLR_logical & (MOD_LWIN | MOD_RWIN))   // At least one WIN key is down.
				&& !(g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT | MOD_LCONTROL | MOD_RCONTROL))   ) // But no SHIFT or CONTROL key is down to help us.
				KeyEvent(KEYDOWNANDUP, g_MenuMaskKey);
			// Since this is a hotkey that fires on ALT-DOWN and it's a normal (suppressed) hotkey,
			// send an up-event to "turn off" the OS's low-level handling for the alt key with
			// respect to having it modify keypresses.  For example, the following hotkeys would
			// fail to work properly without this workaround because the OS apparently sees that
			// the ALT key is physically down even though it is not logically down:
			// RAlt::Send f  ; Actually triggers !f, which activates the FILE menu if the active window has one.
			// RAlt::Send {PgDn}  ; Fails to work because ALT-PgDn usually does nothing.
			KeyEvent(KEYUP, aVK, aSC);
		}
	}
	
	// Otherwise:
	if (!aKeyUp)
		this_key.hotkey_down_was_suppressed = true;
	return SuppressThisKey;
}



LRESULT SuppressThisKeyFunc(const HHOOK aHook, LPARAM lParam, const vk_type aVK, const sc_type aSC, bool aKeyUp
	, KeyHistoryItem *pKeyHistoryCurr, HotkeyIDType aHotkeyIDToPost, WPARAM aHSwParamToPost, LPARAM aHSlParamToPost)
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
	if (pKeyHistoryCurr->event_type == ' ') // then it hasn't been already set somewhere else
		pKeyHistoryCurr->event_type = 's';
	// This handles the troublesome Numlock key, which on some (most/all?) keyboards
	// will change state independent of the keyboard's indicator light even if its
	// keydown and up events are suppressed.  This is certainly true on the
	// MS Natural Elite keyboard using default drivers on WinXP.  SetKeyboardState()
	// doesn't resolve this, so the only alternative to the below is to use the
	// Win9x method of setting the Numlock state explicitly whenever the key is released.
	// That might be complicated by the fact that the unexpected state change described
	// here can't be detected by GetKeyboardState() and such (it sees the state indicated
	// by the numlock light on the keyboard, which is wrong).  In addition, doing it this
	// way allows Numlock to be a prefix key for something like Numpad7, which would
	// otherwise be impossible because Numpad7 would become NumpadHome the moment
	// Numlock was pressed down.  Note: this problem doesn't appear to affect Capslock
	// or Scrolllock for some reason, possibly hardware or driver related.
	// Note: the check for KEY_IGNORE isn't strictly necessary, but here just for safety
	// in case this is ever called for a key that should be ignored.  If that were
	// to happen and we didn't check for it, and endless loop of keyboard events
	// might be caused due to the keybd events sent below.
	if (aHook == g_KeybdHook)
	{
		KBDLLHOOKSTRUCT &event = *(PKBDLLHOOKSTRUCT)lParam;
		if (aVK == VK_NUMLOCK && !aKeyUp && !IsIgnored(event.dwExtraInfo))
		{
			// This seems to undo the faulty indicator light problem and toggle
			// the key back to the state it was in prior to when the user pressed it.
			// Originally, I had two keydowns and before that some keyups too, but
			// testing reveals that only a single key-down is needed.  UPDATE:
			// It appears that all 4 of these key events are needed to make it work
			// in every situation, especially the case when ForceNumlock is on but
			// numlock isn't used for any hotkeys.
			// Note: The only side-effect I've discovered of this method is that the
			// indicator light can't be toggled after the program is exitted unless the
			// key is pressed twice:
			KeyEvent(KEYUP, VK_NUMLOCK);
			KeyEvent(KEYDOWNANDUP, VK_NUMLOCK);
			KeyEvent(KEYDOWN, VK_NUMLOCK);
		}
		UpdateKeybdState(event, aVK, aSC, aKeyUp, true);
	}

#ifdef ENABLE_KEY_HISTORY_FILE
	if (g_KeyHistoryToFile)
		KeyHistoryToFile(NULL, pKeyHistoryCurr->event_type, pKeyHistoryCurr->key_up
			, pKeyHistoryCurr->vk, pKeyHistoryCurr->sc);  // A fairly low overhead operation.
#endif

	// These should be posted only at the last possible moment before returning in order to
	// minimize the chance that the main thread will receive and process the message before
	// our thread can finish updating key states and other maintenance.  This has been proven
	// to be a problem on single-processor systems when the hook thread gets preempted
	// before it can return.  Apparently, the fact that the hook thread is much higher in priority
	// than the main thread is not enough to prevent the main thread from getting a timeslice
	// before the hook thread gets back another (at least on some systems, perhaps due to their
	// system settings of the same ilk as "favor background processes").
	if (aHotkeyIDToPost != HOTKEY_ID_INVALID)
		PostMessage(g_hWnd, AHK_HOOK_HOTKEY, aHotkeyIDToPost, pKeyHistoryCurr->sc); // v1.0.43.03: sc is posted currently only to support the number of wheel turns (to store in A_EventInfo).
	if (aHSwParamToPost != HOTSTRING_INDEX_INVALID)
		PostMessage(g_hWnd, AHK_HOTSTRING, aHSwParamToPost, aHSlParamToPost);
	return 1;
}



LRESULT AllowIt(const HHOOK aHook, int aCode, WPARAM wParam, LPARAM lParam, const vk_type aVK, const sc_type aSC
	, bool aKeyUp, KeyHistoryItem *pKeyHistoryCurr, HotkeyIDType aHotkeyIDToPost, bool aDisguiseWinAlt)
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
	WPARAM hs_wparam_to_post = HOTSTRING_INDEX_INVALID; // Set default.
	LPARAM hs_lparam_to_post; // Not initialized because the above is the sole indicator of whether its contents should even be examined.

	// Prevent toggleable keys from being toggled (if the user wanted that) by suppressing it.
	// Seems best to suppress key-up events as well as key-down, since a key-up by itself,
	// if seen by the system, doesn't make much sense and might have unwanted side-effects
	// in rare cases (e.g. if the foreground app takes note of these types of key events).
	// Don't do this for ignored keys because that could cause an endless loop of
	// numlock events due to the keybd events that SuppressThisKey sends.
	// It's a little more readable and comfortable not to rely on short-circuit
	// booleans and instead do these conditions as separate IF statements.
	if (aHook == g_MouseHook)
	{
		// Since a mouse button that is physically down is not necessarily logically down -- such as
		// when the mouse button is a suppressed hotkey -- only update the logical state (which is the
		// state the OS believes the key to be in) when this even is non-supressed (i.e. allowed to
		// go to the system):
#ifdef FUTURE_USE_MOUSE_BUTTONS_LOGICAL
		// THIS ENTIRE SECTION might never be necessary if it's true that GetAsyncKeyState() and
		// GetKeyState() can retrieve the logical mouse button state on Windows NT/2000/XP, which are
		// the only OSes that matter for this purpose because the hooks aren't supported on Win9x.
		KBDLLHOOKSTRUCT &event = *(PMSDLLHOOKSTRUCT)lParam;  // For convenience, maintainability, and possibly performance.
		switch (wParam)
		{
			case WM_LBUTTONUP: g_mouse_buttons_logical &= ~MK_LBUTTON; break;
			case WM_RBUTTONUP: g_mouse_buttons_logical &= ~MK_RBUTTON; break;
			case WM_MBUTTONUP: g_mouse_buttons_logical &= ~MK_MBUTTON; break;
			// WM_NCXBUTTONUP is a click in the non-client area of a window.  MSDN implies this message can be
			// received by the mouse hook  but it seems doubtful because its counterparts, such as WM_NCLBUTTONUP,
			// are apparently never received:
			case WM_NCXBUTTONUP:
			case WM_XBUTTONUP:
				g_mouse_buttons_logical &= ~(   (HIWORD(event.mouseData)) == XBUTTON1 ? MK_XBUTTON1 : MK_XBUTTON2   );
				break;
			case WM_LBUTTONDOWN: g_mouse_buttons_logical |= MK_LBUTTON; break;
			case WM_RBUTTONDOWN: g_mouse_buttons_logical |= MK_RBUTTON; break;
			case WM_MBUTTONDOWN: g_mouse_buttons_logical |= MK_MBUTTON; break;
			case WM_NCXBUTTONDOWN:
			case WM_XBUTTONDOWN:
				g_mouse_buttons_logical |= (HIWORD(event.mouseData) == XBUTTON1) ? MK_XBUTTON1 : MK_XBUTTON2;
				break;
		}
#endif
#ifdef ENABLE_KEY_HISTORY_FILE
		if (g_KeyHistoryToFile)
			KeyHistoryToFile(NULL, pKeyHistoryCurr->event_type, pKeyHistoryCurr->key_up
				, pKeyHistoryCurr->vk, pKeyHistoryCurr->sc);  // A fairly low overhead operation.
#endif
	}
	else // Our caller is the keyboard hook.
	{
		KBDLLHOOKSTRUCT &event = *(PKBDLLHOOKSTRUCT)lParam;

		bool is_ignored = IsIgnored(event.dwExtraInfo);
		if (!is_ignored)
		{
			if (kvk[aVK].pForceToggle) // Key is a toggleable key.
			{
				// Dereference to get the global var's value:
				if (*(kvk[aVK].pForceToggle) != NEUTRAL) // Prevent toggle.
					return SuppressThisKeyFunc(aHook, lParam, aVK, aSC, aKeyUp, pKeyHistoryCurr, aHotkeyIDToPost);
			}
		}

		// This is done unconditionally so that even if a qualified Input is not in progress, the
		// variable will be correctly reset anyway:
		if (sVKtoIgnoreNextTimeDown && sVKtoIgnoreNextTimeDown == aVK && !aKeyUp)
			sVKtoIgnoreNextTimeDown = 0;  // i.e. this ignore-for-the-sake-of-CollectInput() ticket has now been used.
		else if ((Hotstring::mAtLeastOneEnabled && !is_ignored) || (g_input.status == INPUT_IN_PROGRESS && !(g_input.IgnoreAHKInput && is_ignored)))
			if (!CollectInput(event, aVK, aSC, aKeyUp, is_ignored, hs_wparam_to_post, hs_lparam_to_post)) // Key should be invisible (suppressed).
				return SuppressThisKeyFunc(aHook, lParam, aVK, aSC, aKeyUp, pKeyHistoryCurr, aHotkeyIDToPost, hs_wparam_to_post, hs_lparam_to_post);

		// Do these here since the above "return SuppressThisKey" will have already done it in that case.
#ifdef ENABLE_KEY_HISTORY_FILE
		if (g_KeyHistoryToFile)
			KeyHistoryToFile(NULL, pKeyHistoryCurr->event_type, pKeyHistoryCurr->key_up
				, pKeyHistoryCurr->vk, pKeyHistoryCurr->sc);  // A fairly low overhead operation.
#endif

		UpdateKeybdState(event, aVK, aSC, aKeyUp, false);

		// UPDATE: The Win-L and Ctrl-Alt-Del workarounds below are still kept in effect in spite of the
		// anti-stick workaround done via GetModifierLRState().  This is because ResetHook() resets more
		// than just the modifiers and physical key state, which seems appropriate since the user might
		// be away for a long period of time while the computer is locked or the security screen is displayed.
		// Win-L uses logical keys, unlike Ctrl-Alt-Del which uses physical keys (i.e. Win-L can be simulated,
		// but Ctrl-Alt-Del must be physically pressed by the user):
		if (   aVK == 'L' && !aKeyUp && (g_modifiersLR_logical == MOD_LWIN  // i.e. *no* other keys but WIN.
			|| g_modifiersLR_logical == MOD_RWIN || g_modifiersLR_logical == (MOD_LWIN | MOD_RWIN))
			&& g_os.IsWinXPorLater())
		{
			// Since the user has pressed Win-L with *no* other modifier keys held down, and since
			// this key isn't being suppressed (since we're here in this function), the computer
			// is about to be locked.  When that happens, the hook is apparently disabled or
			// deinstalled until the user logs back in.  Because it is disabled, it will not be
			// notified when the user releases the LWIN or RWIN key, so we should assume that
			// it's now not in the down position.  This avoids it being thought to be down when the
			// user logs back in, which might cause hook hotkeys to accidentally fire.
			// Update: I've received an indication from a single Win2k user (unconfirmed from anyone
			// else) that the Win-L hotkey doesn't work on Win2k.  AutoIt3 docs confirm this.
			// Thus, it probably doesn't work on NT either.  So it's been changed to happen only on XP:
			ResetHook(true); // We already know that *only* the WIN key is down.
			// Above will reset g_PhysicalKeyState, especially for the windows keys and the 'L' key
			// (in our case), in preparation for re-logon:
		}

		// Although the delete key itself can be simulated (logical or physical), the user must be physically
		// (not logically) holding down CTRL and ALT for the ctrl-alt-del sequence to take effect,
		// which is why g_modifiersLR_physical is used vs. g_modifiersLR_logical (which is used above since
		// it's different).  Also, this is now done for XP -- in addition to NT4 & Win2k -- in case XP is
		// configured to display the NT/2k style security window instead of the task manager.  This is
		// probably very common because whenever the welcome screen is diabled, that's the default behavior?:
		// Control Panel > User Accounts > Use the welcome screen for fast and easy logon
		if (   (aVK == VK_DELETE || aVK == VK_DECIMAL) && !aKeyUp         // Both of these qualify, see notes.
			&& (g_modifiersLR_physical & (MOD_LCONTROL | MOD_RCONTROL)) // At least one CTRL key is physically down.
			&& (g_modifiersLR_physical & (MOD_LALT | MOD_RALT))         // At least one ALT key is physically down.
			&& !(g_modifiersLR_physical & (MOD_LSHIFT | MOD_RSHIFT))    // Neither shift key is phys. down (WIN is ok).
			&& g_os.IsWinNT4orLater()   )
		{
			// Similar to the above case except for Windows 2000.  I suspect it also applies to NT,
			// but I'm not sure.  It seems safer to apply it to NT until confirmed otherwise.
			// Note that Ctrl-Alt-Delete works with *either* delete key, and it works regardless
			// of the state of Numlock (at least on XP, so it's probably that way on Win2k/NT also,
			// though it would be nice if this too is someday confirmed).  Here's the key history
			// someone for when the pressed ctrl-alt-del and then pressed esc to dismiss the dialog
			// on Win2k (Win2k invokes a 6-button dialog, with choices such as task manager and lock
			// workstation, if I recall correctly -- unlike XP which invokes task mgr by default):
			// A4  038	 	d	21.24	Alt            	
			// A2  01D	 	d	0.00	Ctrl           	
			// A2  01D	 	d	0.52	Ctrl           	
			// 2E  053	 	d	0.02	Num Del        	<-- notice how there's no following up event
			// 1B  001	 	u	2.80	Esc             <-- notice how there's no preceding down event
			// Other notes: On XP at least, shift key must not be down, otherwise Ctrl-Alt-Delete does
			// not take effect.  Windows key can be down, however.
			// Since the user will be gone for an unknown amount of time, it seems best just to reset
			// all hook tracking of the modifiers to the "up" position.  The user can always press them
			// down again upon return.  It also seems best to reset both logical and physical, just for
			// peace of mind and simplicity:
			ResetHook(true);
			// The above will also reset g_PhysicalKeyState so that especially the following will not
			// be thought to be physically down:CTRL, ALT, and DEL keys.  This is done in preparation
			// for returning from the security screen.  The neutral keys (VK_MENU and VK_CONTROL)
			// must also be reset -- not just because it's correct but because CollectInput() relies on it.
		}

		// Bug-fix for v1.0.20: The below section was moved out of LowLevelKeybdProc() to here because
		// sAltTabMenuIsVisible should not be set to true prior to knowing whether the current tab-down
		// event will be suppressed.  This is because if it is suppressed, the menu will not become visible
		// after all since the system will never see the tab-down event.
		// Having this extra check here, in addition to the other(s) that set sAltTabMenuIsVisible to be
		// true, allows AltTab and ShiftAltTab hotkeys to function even when the AltTab menu was invoked by
		// means other than an AltTabMenu or AltTabAndMenu hotkey.  The alt-tab menu becomes visible only
		// under these exact conditions, at least under WinXP:
		if (aVK == VK_TAB && !aKeyUp && !sAltTabMenuIsVisible
			&& (g_modifiersLR_logical & (MOD_LALT | MOD_RALT)) // At least one ALT key is down.
			&& !(g_modifiersLR_logical & (MOD_LCONTROL | MOD_RCONTROL))) // Neither CTRL key is down.
			sAltTabMenuIsVisible = true;

		if (kvk[aVK].as_modifiersLR) // It's a modifier key.
		{
			// Don't do it this way because then the alt key itself can't be reliable used as "AltTabMenu"
			// (due to ShiftAltTab causing sAltTabMenuIsVisible to become false):
			//if (   sAltTabMenuIsVisible && !((g_modifiersLR_logical & MOD_LALT) || (g_modifiersLR_logical & MOD_RALT))
			//	&& !(aKeyUp && pKeyHistoryCurr->event_type == 'h')   )  // In case the alt key itself is "AltTabMenu"
			if (   sAltTabMenuIsVisible && // Release of Alt key or press down of Escape:
				(aKeyUp && (aVK == VK_LMENU || aVK == VK_RMENU || aVK == VK_MENU)
					|| !aKeyUp && aVK == VK_ESCAPE)
				// In case the alt key itself is "AltTabMenu":
				&& pKeyHistoryCurr->event_type != 'h' && pKeyHistoryCurr->event_type != 's'   )
				// It's important to reset in this case because if sAltTabMenuIsVisible were to
				// stay true and the user presses ALT in the future for a purpose other than to
				// display the Alt-tab menu, we would incorrectly believe the menu to be displayed:
				sAltTabMenuIsVisible = false;

			bool vk_is_win = aVK == VK_LWIN || aVK == VK_RWIN;
			if (aDisguiseWinAlt && aKeyUp && (vk_is_win || aVK == VK_MENU || aVK == VK_LMENU
				|| aVK == VK_RMENU && !ActiveWindowLayoutHasAltGr())) // AltGr should never need disguising, and avoiding it may help avoid unwanted side-effects.
			{
				// I think the best way to do this is to suppress the given key-event and substitute
				// some new events to replace it.  This is because otherwise we would probably have to
				// Sleep() or wait for the shift key-down event to take effect before calling
				// CallNextHookEx(), so that the shift key will be in effect in time for the win
				// key-up event to be disguised properly.  UPDATE: Currently, this doesn't check
				// to see if a shift key is already down for some other reason; that would be
				// pretty rare anyway, and I have more confidence in the reliability of putting
				// the shift key down every time.  UPDATE #2: Ctrl vs. Shift is now used to avoid
				// issues with the system's language-switch hotkey.  See detailed comments in
				// SetModifierLRState() about this.
				// Also, check the current logical state of CTRL to see if it's already down, for these reasons:
				// 1) There is no need to push it down again, since the release of ALT or WIN will be
				//    successfully disguised as long as it's down currently.
				// 2) If it's already down, the up-event part of the disguise keystroke would put it back
				//    up, which might mess up other things that rely upon it being down.
				bool disguise_it = true;  // Starting default.
				if (g_modifiersLR_logical & (MOD_LCONTROL | MOD_RCONTROL))
					disguise_it = false; // LCTRL or RCTRL is already down, so disguise is already in effect.
				else if (   vk_is_win && (g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT | MOD_LALT | MOD_RALT))   )
					disguise_it = false; // WIN key disguise is easier to satisfy, so don't need it in these cases either.
				// Since the below call to KeyEvent() calls the keybd hook recursively, a quick down-and-up
				// on Control is all that is necessary to disguise the key.  This is because the OS will see
				// that the Control keystroke occurred while ALT or WIN is still down because we haven't
				// done CallNextHookEx() yet.
				if (disguise_it)
					KeyEvent(KEYDOWNANDUP, g_MenuMaskKey); // Fix for v1.0.25: Use Ctrl vs. Shift to avoid triggering the LAlt+Shift language-change hotkey.
			}
		} // It's a modifier key.
	} // Keyboard vs. mouse hook.

	// Since above didn't return, this keystroke is being passed through rather than suppressed.
	if (g_HSResetUponMouseClick && (aVK == VK_LBUTTON || aVK == VK_RBUTTON)) // v1.0.42.03
	{
		*g_HSBuf = '\0';
		g_HSBufLength = 0;
	}
	// In case CallNextHookEx() is high overhead or can sometimes take a long time to return,
	// call it before posting the messages.  This solves conditions in which the main thread is
	// able to launch a script subroutine before the hook thread can finish updating its key state.
	// Search on AHK_HOOK_HOTKEY in this file for more comments.
	LRESULT result_to_return = CallNextHookEx(aHook, aCode, wParam, lParam);
	if (aHotkeyIDToPost != HOTKEY_ID_INVALID)
		PostMessage(g_hWnd, AHK_HOOK_HOTKEY, aHotkeyIDToPost, pKeyHistoryCurr->sc); // v1.0.43.03: sc is posted currently only to support the number of wheel turns (to store in A_EventInfo).
	if (hs_wparam_to_post != HOTSTRING_INDEX_INVALID)
		PostMessage(g_hWnd, AHK_HOTSTRING, hs_wparam_to_post, hs_lparam_to_post);
	return result_to_return;
}



bool CollectInput(KBDLLHOOKSTRUCT &aEvent, const vk_type aVK, const sc_type aSC, bool aKeyUp, bool aIsIgnored
	, WPARAM &aHotstringWparamToPost, LPARAM &aHotstringLparamToPost)
// Caller is reponsible for having initialized aHotstringWparamToPost to HOTSTRING_INDEX_INVALID.
// Returns true if the caller should treat the key as visible (non-suppressed).
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
#define shs Hotstring::shs  // For convenience.
	// Generally, we return this value to our caller so that it will treat the event as visible
	// if either there's no input in progress or if there is but it's visible.  Below relies on
	// boolean evaluation order:
	bool treat_as_visible = g_input.status != INPUT_IN_PROGRESS || g_input.Visible
		|| kvk[aVK].pForceToggle;  // Never suppress toggleable keys such as CapsLock.

	if (aKeyUp)
		// Always pass modifier-up events through unaltered.  At the very least, this is needed for
		// cases where a user presses a #z hotkey, for example, to initiate an Input.  When the user
		// releases the LWIN/RWIN key during the input, that up-event should not be suppressed
		// otherwise the modifier key would get "stuck down".  
		return kvk[aVK].as_modifiersLR ? true : treat_as_visible;

	// Hotstrings monitor neither ignored input nor input that is invisible due to suppression by
	// the Input command.  One reason for not monitoring ignored input is to avoid any chance of
	// an infinite loop of keystrokes caused by one hotstring triggering itself directly or
	// indirectly via a different hotstring:
	bool do_monitor_hotstring = shs && !aIsIgnored && treat_as_visible;
	bool do_input = g_input.status == INPUT_IN_PROGRESS && !(g_input.IgnoreAHKInput && aIsIgnored);

	UCHAR end_key_attributes;
	if (do_input)
	{
		end_key_attributes = g_input.EndVK[aVK];
		if (!end_key_attributes)
			end_key_attributes = g_input.EndSC[aSC];
		if (end_key_attributes) // A terminating keystroke has now occurred unless the shift state isn't right.
		{
			// Caller has ensured that only one of the flags below is set (if any):
			bool shift_must_be_down = end_key_attributes & END_KEY_WITH_SHIFT;
			bool shift_must_not_be_down = end_key_attributes & END_KEY_WITHOUT_SHIFT;
			bool shift_state_matters = shift_must_be_down && !shift_must_not_be_down
				|| !shift_must_be_down && shift_must_not_be_down; // i.e. exactly one of them.
			if (    !shift_state_matters
				|| (shift_must_be_down && (g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT)))
				|| (shift_must_not_be_down && !(g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT)))   )
			{
				// The shift state is correct to produce the desired end-key.
				g_input.status = INPUT_TERMINATED_BY_ENDKEY;
				g_input.EndedBySC = g_input.EndSC[aSC];
				g_input.EndingVK = aVK;
				g_input.EndingSC = aSC;
				// Don't change this line:
				g_input.EndingRequiredShift = shift_must_be_down && (g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT));
				if (!do_monitor_hotstring)
					return treat_as_visible;
				// else need to return only after the input is collected for the hotstring.
			}
		}
	}

	// Fix for v1.0.29: Among other things, this resolves the problem where an Input command without
	// the Visible option in effect would capture uppercase and other shifted characters as unshifted:
	switch (aVK)
	{
	case VK_LSHIFT:
	case VK_RSHIFT:
	case VK_LCONTROL:
	case VK_RCONTROL:
	case VK_LMENU:
	case VK_RMENU:
	case VK_LWIN:
	case VK_RWIN:
		return true; // "true" because the above should never be suppressed, nor do they need further processing below.

	case VK_LEFT:
	case VK_RIGHT:
	case VK_DOWN:
	case VK_UP:
	case VK_NEXT:
	case VK_PRIOR:
	case VK_HOME:
	case VK_END:
		// Reset hotstring detection if the user seems to be navigating within an editor.  This is done
		// so that hotstrings do not fire in unexpected places.
		if (do_monitor_hotstring && g_HSBufLength)
		{
			*g_HSBuf = '\0';
			g_HSBufLength = 0;
		}
		break;
	}

	// Don't unconditionally transcribe modified keys such as Ctrl-C because calling ToAsciiEx() on
	// some such keys (e.g. Ctrl-LeftArrow or RightArrow if I recall correctly), disrupts the native
	// function of those keys.  That is the reason for the existence of the
	// g_input.TranscribeModifiedKeys option.
	// Fix for v1.0.38: Below now uses g_modifiersLR_logical vs. g_modifiersLR_physical because
	// it's the logical state that determines what will actually be produced on the screen and
	// by ToAsciiEx() below.  This fixes the Input command to properly capture simulated
	// keystrokes even when they were sent via hotkey such #c or a hotstring for which the user
	// might still be holding down a modifier, such as :*:<t>::Test (if '>' requires shift key).
	// It might also fix other issues.
	if (g_modifiersLR_logical
		&& !(g_input.status == INPUT_IN_PROGRESS && g_input.TranscribeModifiedKeys)
		&& g_modifiersLR_logical != MOD_LSHIFT && g_modifiersLR_logical != MOD_RSHIFT
		&& g_modifiersLR_logical != (MOD_LSHIFT & MOD_RSHIFT)
		&& !((g_modifiersLR_logical & (MOD_LALT | MOD_RALT)) && (g_modifiersLR_logical & (MOD_LCONTROL | MOD_RCONTROL))))
		// Since in some keybd layouts, AltGr (Ctrl+Alt) will produce valid characters (such as the @ symbol,
		// which is Ctrl+Alt+Q in the German/IBM layout and Ctrl+Alt+2 in the Spanish layout), an attempt
		// will now be made to transcribe all of the following modifier combinations:
		// - Anything with no modifiers at all.
		// - Anything that uses ONLY the shift key.
		// - Anything with Ctrl+Alt together in it, including Ctrl+Alt+Shift, etc. -- but don't do
		//   "anything containing the Alt key" because that causes weird side-effects with
		//   Alt+LeftArrow/RightArrow and maybe other keys too).
		// Older comment: If any modifiers except SHIFT are physically down, don't transcribe the key since
		// most users wouldn't want that.  An additional benefit of this policy is that registered hotkeys will
		// normally be excluded from the input (except those rare ones that have only SHIFT as a modifier).
		// Note that ToAsciiEx() will translate ^i to a tab character, !i to plain i, and many other modified
		// letters as just the plain letter key, which we don't want.
		return treat_as_visible;

	static vk_type sPendingDeadKeyVK = 0;
	static sc_type sPendingDeadKeySC = 0; // Need to track this separately because sometimes default VK-to-SC mapping isn't correct.
	static bool sPendingDeadKeyUsedShift = false;
	static bool sPendingDeadKeyUsedAltGr = false;

	// v1.0.21: Only true (unmodified) backspaces are recognized by the below.  Another reason to do
	// this is that ^backspace has a native function (delete word) different than backspace in many editors.
	// Fix for v1.0.38: Below now uses g_modifiersLR_logical vs. physical because it's the logical state
	// that determines whether the backspace behaves like an unmodified backspace.  This solves the issue
	// of the Input command collecting simulated backspaces as real characters rather than recognizing
	// them as a means to erase the previous character in the buffer.
	if (aVK == VK_BACK && !g_modifiersLR_logical) // Backspace
	{
		// Note that it might have been in progress upon entry to this function but now isn't due to
		// INPUT_TERMINATED_BY_ENDKEY above:
		if (do_input && g_input.status == INPUT_IN_PROGRESS && g_input.BackspaceIsUndo) // Backspace is being used as an Undo key.
			if (g_input.BufferLength)
				g_input.buffer[--g_input.BufferLength] = '\0';
		if (do_monitor_hotstring && g_HSBufLength)
			g_HSBuf[--g_HSBufLength] = '\0';
		if (sPendingDeadKeyVK) // Doing this produces the expected behavior when a backspace occurs immediately after a dead key.
			sPendingDeadKeyVK = 0;
		return treat_as_visible;
	}


	BYTE ch[3], key_state[256];
	memcpy(key_state, g_PhysicalKeyState, 256);
	// As of v1.0.25.10, the below fixes the Input command so that when it is capturing artificial input,
	// such as from the Send command or a hotstring's replacement text, the captured input will reflect
	// any modifiers that are logically but not physically down (or vice versa):
	AdjustKeyState(key_state, g_modifiersLR_logical);
	// Make the state of capslock accurate so that ToAsciiEx() will return upper vs. lower if appropriate:
	if (IsKeyToggledOn(VK_CAPITAL))
		key_state[VK_CAPITAL] |= STATE_ON;
	else
		key_state[VK_CAPITAL] &= ~STATE_ON;

	// Use ToAsciiEx() vs. ToAscii() because there is evidence from Putty author that ToAsciiEx() works better
	// with more keyboard layouts under 2k/XP than ToAscii() does (though if true, there is no MSDN explanation). 
	// UPDATE: In v1.0.44.03, need to use ToAsciiEx() anyway because of the adapt-to-active-window-layout feature.
	Get_active_window_keybd_layout // Defines the variables active_window and active_window_keybd_layout for use below.
	int byte_count = ToAsciiEx(aVK, aEvent.scanCode  // Uses the original scan code, not the adjusted "sc" one.
		, key_state, (LPWORD)ch, g_MenuIsVisible ? 1 : 0, active_window_keybd_layout);
	if (!byte_count) // No translation for this key.
		return treat_as_visible;

	// More notes about dead keys: The dead key behavior of Enter/Space/Backspace is already properly
	// maintained when an Input or hotstring monitoring is in effect.  In addition, keys such as the
	// following already work okay (i.e. the user can press them in between the pressing of a dead
	// key and it's finishing/base/trigger key without disrupting the production of diacritical letters)
	// because ToAsciiEx() finds no translation-to-char for them:
	// pgup/dn/home/end/ins/del/arrowkeys/f1-f24/etc.
	// Note that if a pending dead key is followed by the press of another dead key (including itself),
	// the sequence should be triggered and both keystrokes should appear in the active window.
	// That case has been tested too, and works okay with the layouts tested so far.
	// I've only discovered two keys which need special handling, and they are handled below:
	// VK_TAB & VK_ESCAPE
	// These keys have an ascii translation but should not trigger/complete a pending dead key,
	// at least not on the Spanish and Danish layouts, which are the two I've tested so far.

	// Dead keys in Danish layout as they appear on a US English keyboard: Equals and Plus /
	// Right bracket & Brace / probably others.

	// SUMMARY OF DEAD KEY ISSUE:
	// Calling ToAsciiEx() on the dead key itself doesn't disrupt anything. The disruption occurs on the next key
	// (for which the dead key is pending): ToAsciiEx() consumes previous/pending dead key, which causes the
	// active window's call of ToAsciiEx() to fail to see a dead key. So unless the program reinserts the dead key
	// after the call to ToAsciiEx() but before allowing the dead key's successor key to pass through to the
	// active window, that window would see a non-diacritic like "u" instead of û.  In other words, the program
	// "uses up" the dead key to populate its own hotstring buffer, depriving the active window of the dead key.
	//
	// JAVA ISSUE: Hotstrings are known to disrupt dead keys in Java apps on some systems (though not my XP one).
	// I spent several hours on it but was unable to solve it using anything other than a Sleep(20) after the
	// reinsertion of the dead key (and PhiLho reports that even that didn't fully solve it).  A Sleep here in the
	// hook would probably do more harm than good, so is avoided for now.  Other approaches:
	// 1) Send a simulated substitute for the successor key rather than allowing the hook to pass it through.
	//    Maybe that would somehow put things in a better order for Java.  However, there might be side-effects to
	//    that, such as in DirectX games.
	// 2) Have main thread (rather than hook thread) reinsert the dead key and its successor key (hook would have
	//    suppressed both), which allows the main thread to do a Sleep or MsgSleep.  Such a Sleep be more effective
	//    because the main thread's priority is lower than that of the hook's, allowing better round-robin.
	// 
	// If this key isn't a dead key but there's a dead key pending and this incoming key is capable of
	// completing/triggering it, do a workaround for the side-effects of ToAsciiEx().  This workaround
	// allows dead keys to continue to operate properly in the user's foreground window, while still
	// being capturable by the Input command and recognizable by any defined hotstrings whose
	// abbreviations use diacritical letters:
	bool dead_key_sequence_complete = sPendingDeadKeyVK && aVK != VK_TAB && aVK != VK_ESCAPE;
	if (byte_count < 0 && !dead_key_sequence_complete) // It's a dead key and it doesn't complete a sequence (i.e. there is no pending dead key before it).
	{
		if (treat_as_visible)
		{
			sPendingDeadKeyVK = aVK;
			sPendingDeadKeySC = aSC;
			sPendingDeadKeyUsedShift = g_modifiersLR_logical & (MOD_LSHIFT | MOD_RSHIFT);
			// Detect AltGr as fully and completely as possible in case the current keyboard layout
			// doesn't even have an AltGr key.  The section above which references sPendingDeadKeyUsedAltGr
			// relies on this check having been done here.  UPDATE:
			// v1.0.35.10: Allow Ctrl+Alt to be seen as AltGr too, which allows users to press Ctrl+Alt+Deadkey
			// rather than AltGr+Deadkey.  It might also resolve other issues.  This change seems okay since
			// the mere fact that this IS a dead key (as checked above) should mean that it's a deadkey made
			// manifest through AltGr.
			// Previous method:
			//sPendingDeadKeyUsedAltGr = (g_modifiersLR_logical & (MOD_LCONTROL | MOD_RALT)) == (MOD_LCONTROL | MOD_RALT);
			sPendingDeadKeyUsedAltGr = (g_modifiersLR_logical & (MOD_LCONTROL | MOD_RCONTROL))
				&& (g_modifiersLR_logical & (MOD_LALT | MOD_RALT));
		}
		// Dead keys must always be hidden, otherwise they would be shown twice literally due to
		// having been "damaged" by ToAsciiEx():
		return false;
	}

	if (ch[0] == '\r')  // Translate \r to \n since \n is more typical and useful in Windows.
		ch[0] = '\n';
	if (ch[1] == '\r')  // But it's never referred to if byte_count < 2
		ch[1] = '\n';

	bool suppress_hotstring_final_char = false; // Set default.

	if (do_monitor_hotstring)
	{
		if (active_window != g_HShwnd)
		{
			// Since the buffer tends to correspond to the text to the left of the caret in the
			// active window, if the active window changes, it seems best to reset the buffer
			// to avoid misfires.
			g_HShwnd = active_window;
			*g_HSBuf = '\0';
			g_HSBufLength = 0;
		}
		else if (HS_BUF_SIZE - g_HSBufLength < 3) // Not enough room for up-to-2 chars.
		{
			// Make room in buffer by removing chars from the front that are no longer needed for HS detection:
			// Bug-fixed the below for v1.0.21:
			g_HSBufLength = (int)strlen(g_HSBuf + HS_BUF_DELETE_COUNT);  // The new length.
			memmove(g_HSBuf, g_HSBuf + HS_BUF_DELETE_COUNT, g_HSBufLength + 1); // +1 to include the zero terminator.
		}

		g_HSBuf[g_HSBufLength++] = ch[0];
		if (byte_count > 1)
			// MSDN: "This usually happens when a dead-key character (accent or diacritic) stored in the
			// keyboard layout cannot be composed with the specified virtual key to form a single character."
			g_HSBuf[g_HSBufLength++] = ch[1];
		g_HSBuf[g_HSBufLength] = '\0';

		if (g_HSBufLength)
		{
			char *cphs, *cpbuf, *cpcase_start, *cpcase_end;
			int case_capable_characters;
			bool first_char_with_case_is_upper, first_char_with_case_has_gone_by;
			CaseConformModes case_conform_mode;

			// Searching through the hot strings in the original, physical order is the documented
			// way in which precedence is determined, i.e. the first match is the only one that will
			// be triggered.
			for (HotstringIDType u = 0; u < Hotstring::sHotstringCount; ++u)
			{
				Hotstring &hs = *shs[u];  // For performance and convenience.
				if (hs.mSuspended)
					continue;
				if (hs.mEndCharRequired)
				{
					if (g_HSBufLength <= hs.mStringLength) // Ensure the string is long enough for loop below.
						continue;
					if (!strchr(g_EndChars, g_HSBuf[g_HSBufLength - 1])) // It's not an end-char, so no match.
						continue;
					cpbuf = g_HSBuf + g_HSBufLength - 2; // Init once for both loops. -2 to omit end-char.
				}
				else // No ending char required.
				{
					if (g_HSBufLength < hs.mStringLength) // Ensure the string is long enough for loop below.
						continue;
					cpbuf = g_HSBuf + g_HSBufLength - 1; // Init once for both loops.
				}
				cphs = hs.mString + hs.mStringLength - 1; // Init once for both loops.
				// Check if this item is a match:
				if (hs.mCaseSensitive)
				{
					for (; cphs >= hs.mString; --cpbuf, --cphs)
						if (*cpbuf != *cphs)
							break;
				}
				else // case insensitive
					// v1.0.43.03: Using CharLower vs. tolower seems the best default behavior (even though slower)
					// so that languages in which the higher ANSI characters are common will see "Ä" == "ä", etc.
					for (; cphs >= hs.mString; --cpbuf, --cphs)
						if (ltolower(*cpbuf) != ltolower(*cphs)) // v1.0.43.04: Fixed crash by properly casting to UCHAR (via macro).
							break;

				// Check if one of the loops above found a matching hotstring (relies heavily on
				// short-circuit boolean order):
				if (   cphs >= hs.mString // One of the loops above stopped early due discovering "no match"...
					// ... or it did but the "?" option is not present to protect from the fact that
					// what lies to the left of this hotstring abbreviation is an alphanumberic character:
					|| !hs.mDetectWhenInsideWord && cpbuf >= g_HSBuf && IsCharAlphaNumeric(*cpbuf)
					// ... v1.0.41: Or it's a perfect match but the right window isn't active or doesn't exist.
					// In that case, continue searching for other matches in case the script contains
					// hotstrings that would trigger simultaneously were it not for the "only one" rule.
					// L4: Added hs.mHotExprLine for #if (expression).
					|| !HotCriterionAllowsFiring(hs.mHotCriterion, hs.mHotWinTitle, hs.mHotWinText, hs.mHotExprIndex)   )
					continue; // No match or not eligible to fire.
					// v1.0.42: The following scenario defeats the ability to give criterion hotstrings
					// precedence over non-criterion:
					// A global/non-criterion hotstring is higher up in the file than some criterion hotstring,
					// but both are eligible to fire at the same instant.  In v1.0.41, the global one would
					// take precedence because it's higher up (and this behavior is preserved not just for
					// backward compatibility, but also because it might be more flexible -- this is because
					// unlike hotkeys, variants aren't stored under a parent hotstring, so we don't know which
					// ones are exact dupes of each other (same options+abbreviation).  Thus, it would take
					// extra code to determine this at runtime; and even if it were added, it might be
					// more flexible not to do it; instead, to let the script determine (even by resorting to
					// #IfWinNOTActive) what precedence hotstrings have with respect to each other.

				// MATCHING HOTSTRING WAS FOUND (since above didn't continue).
				// Since default KeyDelay is 0, and since that is expected to be typical, it seems
				// best to unconditionally post a message rather than trying to handle the backspacing
				// and replacing here.  This is because a KeyDelay of 0 might be fairly slow at
				// sending keystrokes if the system is under heavy load, in which case we would
				// not be returning to our caller in a timely fashion, which would case the OS to
				// think the hook is unreponsive, which in turn would cause it to timeout and
				// route the key through anyway (testing confirms this).
				if (!hs.mConformToCase)
					case_conform_mode = CASE_CONFORM_NONE;
				else
				{
					// Find out what case the user typed the string in so that we can have the
					// replacement produced in similar case:
					cpcase_end = g_HSBuf + g_HSBufLength;
					if (hs.mEndCharRequired)
						--cpcase_end;
					// Bug-fix for v1.0.19: First find out how many of the characters in the abbreviation
					// have upper and lowercase versions (i.e. exclude digits, punctuation, etc):
					for (case_capable_characters = 0, first_char_with_case_is_upper = first_char_with_case_has_gone_by = false
						, cpcase_start = cpcase_end - hs.mStringLength
						; cpcase_start < cpcase_end; ++cpcase_start)
						if (IsCharLower(*cpcase_start) || IsCharUpper(*cpcase_start)) // A case-capable char.
						{
							if (!first_char_with_case_has_gone_by)
							{
								first_char_with_case_has_gone_by = true;
								if (IsCharUpper(*cpcase_start))
									first_char_with_case_is_upper = true; // Override default.
							}
							++case_capable_characters;
						}
					if (!case_capable_characters) // All characters in the abbreviation are caseless.
						case_conform_mode = CASE_CONFORM_NONE;
					else if (case_capable_characters == 1)
						// Since there is only a single character with case potential, it seems best as
						// a default behavior to capitalize the first letter of the replacment whenever
						// that character was typed in uppercase.  The behavior can be overridden by
						// turning off the case-conform mode.
						case_conform_mode = first_char_with_case_is_upper ? CASE_CONFORM_FIRST_CAP : CASE_CONFORM_NONE;
					else // At least two characters have case potential. If all of them are upper, use ALL_CAPS.
					{
						if (!first_char_with_case_is_upper) // It can't be either FIRST_CAP or ALL_CAPS.
							case_conform_mode = CASE_CONFORM_NONE;
						else // First char is uppercase, and if all the others are too, this will be ALL_CAPS.
						{
							case_conform_mode = CASE_CONFORM_FIRST_CAP; // Set default.
							// Bug-fix for v1.0.19: Changed !IsCharUpper() below to IsCharLower() so that
							// caseless characters such as the @ symbol do not disqualify an abbreviation
							// from being considered "all uppercase":
							for (cpcase_start = cpcase_end - hs.mStringLength; cpcase_start < cpcase_end; ++cpcase_start)
								if (IsCharLower(*cpcase_start)) // Use IsCharLower to better support chars from non-English languages.
									break; // Any lowercase char disqualifies CASE_CONFORM_ALL_CAPS.
							if (cpcase_start == cpcase_end) // All case-possible characters are uppercase.
								case_conform_mode = CASE_CONFORM_ALL_CAPS;
							//else leave it at the default set above.
						}
					}
				}

				if (hs.mDoBackspace || hs.mOmitEndChar) // Fix for v1.0.37.07: Added hs.mOmitEndChar so that B0+O will omit the ending character.
				{
					// Have caller suppress this final key pressed by the user, since it would have
					// to be backspaced over anyway.  Even if there is a visible Input command in
					// progress, this should still be okay since the input will still see the key,
					// it's just that the active window won't see it, which is okay since once again
					// it would have to be backspaced over anyway.  UPDATE: If an Input is in progress,
					// it should not receive this final key because otherwise the hotstring's backspacing
					// would backspace one too few times from the Input's point of view, thus the input
					// would have one extra, unwanted character left over (namely the first character
					// of the hotstring's abbreviation).  However, this method is not a complete
					// solution because it fails to work under a situation such as the following:
					// A hotstring script is started, followed by a separate script that uses the
					// Input command.  The Input script's hook will take precedence (since it was
					// started most recently), thus when the Hotstring's script's hook does sends
					// its replacement text, the Input script's hook will get a hold of it first
					// before the Hotstring's script has a chance to suppress it.  In other words,
					// The Input command will capture the ending character and then there will
					// be insufficient backspaces sent to clear the abbrevation out of it.  This
					// situation is quite rare so for now it's just mentioned here as a known limitation.
					treat_as_visible = false; // It might already have been false due to an invisible-input in progress, etc.
					suppress_hotstring_final_char = true; // This var probably must be separate from treat_as_visible to support invisible inputs.
				}

				// Post the message rather than sending it, because Send would need
				// SendMessageTimeout(), which is undesirable because the whole point of
				// making this hook thread separate from the main thread is to have it be
				// maximally responsive (especially to prevent mouse cursor lag).
				// Put the end char in the LOWORD and the case_conform_mode in the HIWORD.
				// Casting to UCHAR might be necessary to avoid problems when MAKELONG
				// casts a signed char to an unsigned WORD.
				// UPDATE: In v1.0.42.01, the message is posted later (by our caller) to avoid
				// situations in which the message arrives and is processed by the main thread
				// before we finish processing the hotstring's final keystroke here.  This avoids
				// problems with a script calling GetKeyState() and getting an inaccurate value
				// because the hook thread is either pre-empted or is running in parallel
				// (multiprocessor) and hasn't yet returned 1 or 0 to determine whether the final
				// keystroke is suppressed or passed through to the active window.
				// UPDATE: In v1.0.43, the ending character is not put into the Lparam when
				// hs.mDoBackspace is false.  This is because:
				// 1) When not backspacing, it's more correct that the ending character appear where the
				//    user typed it rather than appearing at the end of the replacement.
				// 2) Two ending characters would appear in pre-1.0.43 versions: one where the user typed
				//    it and one at the end, which is clearly incorrect.
				aHotstringWparamToPost = u; // Override the default set by caller.
				aHotstringLparamToPost = MAKELONG(
					hs.mEndCharRequired  // v1.0.48.04: Fixed to omit "&& hs.mDoBackspace" so that A_EndChar is set properly even for option "B0" (no backspacing).
						? (UCHAR)g_HSBuf[g_HSBufLength - 1]  // Used by A_EndChar and Hotstring::DoReplace().
						: (dead_key_sequence_complete && suppress_hotstring_final_char) // v1.0.44.09: See comments below.
					, case_conform_mode);
				// v1.0.44.09: dead_key_sequence_complete was added above to tell DoReplace() to do one fewer
				// backspaces in cases where the final/triggering key of a hotstring is the second key of
				// a dead key sequence (such as a tilde in Portuguese followed by virtually any character).
				// What happens in that case is that the dead key is suppressed (for the reasons described in
				// the dead keys handler), but so is the key that follows it when suppress_hotstring_final_char
				// is true.  In addition to being suppressed, no substitute ever needs to be sent for the dead key
				// because it will never appear on the screen (due to being a true auto-replace hotstring).
				// Note: To enhance maintainability and understandability, above also checks
				// suppress_hotstring_final_char (even though probably not strictly necessary).

				// Clean up.
				// The keystrokes to be sent by the other thread upon receiving the message prepared above
				// will not be received by this function because:
				// 1) CollectInput() is not called for simulated keystrokes.
				// 2) The keyboard hook is absent during a SendInput hotstring.
				// 3) The keyboard hook does not receive SendPlay keystrokes (if hotstring is of that type).
				// Consequently, the buffer should be adjusted below to ensure it's in the right state to work
				// in situations such as the user typing two hotstrings consecutively where the ending
				// character of the first is used as a valid starting character (non-alphanumeric) for the next.
				if (*hs.mReplacement)
				{
					// Since the buffer no longer reflects what is actually on screen to the left
					// of the caret position (since a replacement is about to be done), reset the
					// buffer, except for any end-char (since that might legitimately form part
					// of another hot string adjacent to the one just typed).  The end-char
					// sent by DoReplace() won't be captured (since it's "ignored input", which
					// is why it's put into the buffer manually here):
					if (hs.mEndCharRequired)
					{
						*g_HSBuf = g_HSBuf[g_HSBufLength - 1];
						g_HSBufLength = 1; // The buffer will be terminated to reflect this length later below.
					}
					else
						g_HSBufLength = 0; // The buffer will be terminated to reflect this length later below.
				}
				else if (hs.mDoBackspace)
				{
					// It's *not* a replacement, but we're doing backspaces, so adjust buf for backspaces
					// and the fact that the final char of the HS (if no end char) or the end char
					// (if end char required) will have been suppressed and never made it to the
					// active window.  A simpler way to understand is to realize that the buffer now
					// contains (for recognition purposes, in its right side) the hotstring and its
					// end char (if applicable), so remove both:
					g_HSBufLength -= hs.mStringLength;
					if (hs.mEndCharRequired)
						--g_HSBufLength; // The buffer will be terminated to reflect this length later below.
				}

				// v1.0.38.04: Fixed the following mDoReset section by moving it beneath the above because
				// the above relies on the fact that the buffer has not yet been reset.
				// v1.0.30: mDoReset was added to prevent hotstrings such as the following
				// from firing twice in a row, if you type 11 followed by another 1 afterward:
				//:*?B0:11::
				//MsgBox,0,test,%A_ThisHotkey%,1 ; Show which key was pressed and close the window after a second.
				//return
				// There are probably many other uses for the reset option (albeit obscure, but they have
				// been brought up in the forum at least twice).
				if (hs.mDoReset)
					g_HSBufLength = 0; // Further below, the buffer will be terminated to reflect this change.

				// In case the above changed the value of g_HSBufLength, terminate the buffer at that position:
				g_HSBuf[g_HSBufLength] = '\0';

				break; // Somewhere above would have done "continue" if a match wasn't found.
			} // for()
		} // if buf not empty
	} // Yes, collect hotstring input.

	// Fix for v1.0.37.06: The following section was moved beneath the hotstring section so that
	// the hotstring section has a chance to set treat_as_visible to false for use below. This fixes
	// wildcard hotstrings whose final character is diacritic, which would otherwise have the
	// dead key reinserted below, which in turn would cause the hotstring's first backspace to fire
	// the dead key (which kills the backspace, turning it into the dead key character itself).
	// For example:
	// :*:jsá::jsmith@somedomain.com
	// On the Spanish (Mexico) keyboard layout, one would type accent (English left bracket) followed by
	// the letter "a" to produce á.
	if (dead_key_sequence_complete)
	{
		vk_type vk_to_send = sPendingDeadKeyVK; // To facilitate early reset below.
		sPendingDeadKeyVK = 0; // First reset this because below results in a recursive call to keyboard hook.
		// If there's an Input in progress and it's invisible, the foreground app won't see the keystrokes,
		// thus no need to re-insert the dead key into the keyboard buffer.  Note that the Input might have
		// been in progress upon entry to this function but now isn't due to INPUT_TERMINATED_BY_ENDKEY above.
		if (treat_as_visible)
		{
			// Tell the recursively called next instance of the keyboard hook not do the following for
			// the below KEYEVENT_PHYS: Do not call ToAsciiEx() on it and do not capture it as part of
			// the Input itself.  Although this is only needed for the case where the statement
			// "(do_input && g_input.status == INPUT_IN_PROGRESS && !g_input.IgnoreAHKInput)" is true
			// (since hotstrings don't capture/monitor AHK-generated input), it's simpler and about the
			// same in performance to do it unconditonally:
			sVKtoIgnoreNextTimeDown = vk_to_send;
			// Ensure the correct shift-state is set for the below event.  The correct shift key (left or
			// right) must be used to prevent sticking keys and other side-effects:
			vk_type which_shift_down = 0;
			if (g_modifiersLR_logical & MOD_LSHIFT)
				which_shift_down = VK_LSHIFT;
			else if (g_modifiersLR_logical & MOD_RSHIFT)
				which_shift_down = VK_RSHIFT;
			vk_type which_shift_to_send = which_shift_down ? which_shift_down : VK_LSHIFT;
			if (sPendingDeadKeyUsedShift != (bool)which_shift_down)
				KeyEvent(sPendingDeadKeyUsedShift ? KEYDOWN : KEYUP, which_shift_to_send);
			// v1.0.25.14: Apply AltGr too, if necessary.  This is necessary because some keyboard
			// layouts have dead keys that are manifest only by holding down AltGr and pressing
			// another key.  If this weren't done, a hotstring script running on Belgian/French
			// layout (and probably many others that have AltGr dead keys) would disrupt the user's
			// ability to use the tilde dead key.  For example, pressing AltGr+Slash (equals sign
			// on Belgian keyboard) followed by the letter o should produce the tilde-over-o
			// character, but it would not if the following AltGr fix isn't in effect.
			// If sPendingDeadKeyUsedAltGr is true, the current keyboard layout has an AltGr key.
			// That plus the fact that VK_RMENU is not down should mean definitively that AltGr is not
			// down. Also, it might be necessary to assign the below to a variable more than just for
			// performance/readability: KeyEvent() results in a recursive call to this hook function,
			// which causes g_modifiersLR_logical to be different after the call.
			bool apply_altgr = sPendingDeadKeyUsedAltGr && !(g_modifiersLR_logical & MOD_RALT);
			if (apply_altgr) // Push down RAlt even if the dead key was achieved via Ctrl+Alt: 1) For code simplicity; 2) It might improve compatibility with Putty and other apps that demand that AltGr be RAlt (not Ctrl+Alt).
				KeyEvent(KEYDOWN, VK_RMENU); // This will also push down LCTRL as an intrinsic part of AltGr's functionality.
			// Since it's a substitute for the previously suppressed physical dead key event, mark it as physical:
			KEYEVENT_PHYS(KEYDOWNANDUP, vk_to_send, sPendingDeadKeySC);
			if (apply_altgr)
				KeyEvent(KEYUP, VK_RMENU); // This will also release LCTRL as an intrinsic part of AltGr's functionality.
			if (sPendingDeadKeyUsedShift != (bool)which_shift_down) // Restore the original shift state.
				KeyEvent(sPendingDeadKeyUsedShift ? KEYUP : KEYDOWN, which_shift_to_send);
		}
	}

	// Note that it might have been in progress upon entry to this function but now isn't due to
	// INPUT_TERMINATED_BY_ENDKEY above:
	if (!do_input || g_input.status != INPUT_IN_PROGRESS || suppress_hotstring_final_char)
		return treat_as_visible; // Returns "false" in cases such as suppress_hotstring_final_char==true.

	// Since above didn't return, the only thing remaining to do below is handle the input that's
	// in progress (which we know is the case otherwise other opportunities to return above would
	// have done so).  Hotstrings (if any) have already been fully handled by the above.

	#define ADD_INPUT_CHAR(ch) \
		if (g_input.BufferLength < g_input.BufferLengthMax)\
		{\
			g_input.buffer[g_input.BufferLength++] = ch;\
			g_input.buffer[g_input.BufferLength] = '\0';\
		}
	ADD_INPUT_CHAR(ch[0])
	if (byte_count > 1)
		// MSDN: "This usually happens when a dead-key character (accent or diacritic) stored in the
		// keyboard layout cannot be composed with the specified virtual key to form a single character."
		ADD_INPUT_CHAR(ch[1])
	if (!g_input.MatchCount) // The match list is empty.
	{
		if (g_input.BufferLength >= g_input.BufferLengthMax)
			g_input.status = INPUT_LIMIT_REACHED;
		return treat_as_visible;
	}
	// else even if BufferLengthMax has been reached, check if there's a match because a match should take
	// precedence over the length limit.

	// Otherwise, check if the buffer now matches any of the key phrases:
	if (g_input.FindAnywhere)
	{
		if (g_input.CaseSensitive)
		{
			for (UINT i = 0; i < g_input.MatchCount; ++i)
			{
				if (strstr(g_input.buffer, g_input.match[i]))
				{
					g_input.status = INPUT_TERMINATED_BY_MATCH;
					return treat_as_visible;
				}
			}
		}
		else // Not case sensitive.
		{
			for (UINT i = 0; i < g_input.MatchCount; ++i)
			{
				// v1.0.43.03: Changed lstrcasestr to strcasestr because it seems unlikely to break any existing
				// scripts and is also more useful given that that Input with match-list is pretty rarely used,
				// and even when it is used, match lists are usually short (so performance isn't impacted much
				// by this change).
				if (lstrcasestr(g_input.buffer, g_input.match[i]))
				{
					g_input.status = INPUT_TERMINATED_BY_MATCH;
					return treat_as_visible;
				}
			}
		}
	}
	else // Exact match is required
	{
		if (g_input.CaseSensitive)
		{
			for (UINT i = 0; i < g_input.MatchCount; ++i)
			{
				if (!strcmp(g_input.buffer, g_input.match[i]))
				{
					g_input.status = INPUT_TERMINATED_BY_MATCH;
					return treat_as_visible;
				}
			}
		}
		else // Not case sensitive.
		{
			for (UINT i = 0; i < g_input.MatchCount; ++i)
			{
				// v1.0.43.03: Changed to locale-insensitive search.  See similar v1.0.43.03 comment above for more details.
				if (!lstrcmpi(g_input.buffer, g_input.match[i]))
				{
					g_input.status = INPUT_TERMINATED_BY_MATCH;
					return treat_as_visible;
				}
			}
		}
	}

	// Otherwise, no match found.
	if (g_input.BufferLength >= g_input.BufferLengthMax)
		g_input.status = INPUT_LIMIT_REACHED;
	return treat_as_visible;
#undef shs  // To avoid naming conflicts
}



void UpdateKeybdState(KBDLLHOOKSTRUCT &aEvent, const vk_type aVK, const sc_type aSC, bool aKeyUp, bool aIsSuppressed)
// Caller has ensured that vk has been translated from neutral to left/right if necessary.
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
	// See above notes near the first mention of SHIFT_KEY_WORKAROUND_TIMEOUT for details.
	// This part of the workaround can be tested via "NumpadEnd::KeyHistory".  Turn on numlock,
	// hold down shift, and press numpad1. The hotkey will fire and the status should display
	// that the shift key is physically, but not logically down at that exact moment:
	if (sPriorEventWasPhysical && (sPriorVK == VK_LSHIFT || sPriorVK == VK_SHIFT)  // But not RSHIFT.
		&& (DWORD)(GetTickCount() - sPriorEventTickCount) < (DWORD)SHIFT_KEY_WORKAROUND_TIMEOUT)
	{
		bool current_is_dual_state = IsDualStateNumpadKey(aVK, aSC);
		// Verified: Both down and up events for the *current* (not prior) key qualify for this:
		bool fix_it = (!sPriorEventWasKeyUp && DualStateNumpadKeyIsDown())  // Case #4 of the workaround.
			|| (sPriorEventWasKeyUp && aKeyUp && current_is_dual_state); // Case #5
		if (fix_it)
			sNextPhysShiftDownIsNotPhys = true;
		// In the first case, both the numpad key-up and down events are eligible:
		if (   fix_it || (sPriorEventWasKeyUp && current_is_dual_state)   )
		{
			// Since the prior event (the shift key) already happened (took effect) and since only
			// now is it known that it shouldn't have been physical, undo the effects of it having
			// been physical:
			g_modifiersLR_physical = sPriorModifiersLR_physical;
			g_PhysicalKeyState[VK_SHIFT] = sPriorShiftState;
			g_PhysicalKeyState[VK_LSHIFT] = sPriorLShiftState;
		}
	}


	// Must do this part prior to updating modifier state because we want to store the values
	// as they are prior to the potentially-erroneously-physical shift key event takes effect.
	// The state of these is also saved because we can't assume that a shift-down, for
	// example, CHANGED the state to down, because it may have been already down before that:
	sPriorModifiersLR_physical = g_modifiersLR_physical;
	sPriorShiftState = g_PhysicalKeyState[VK_SHIFT];
	sPriorLShiftState = g_PhysicalKeyState[VK_LSHIFT];

	// If this function was called from SuppressThisKey(), these comments apply:
	// Currently SuppressThisKey is only called with a modifier in the rare case
	// when sDisguiseNextLWinUp/RWinUp is in effect.  But there may be other cases in the
	// future, so we need to make sure the physical state of the modifiers is updated
	// in our tracking system even though the key is being suppressed:
	modLR_type modLR;
	if (modLR = kvk[aVK].as_modifiersLR) // Update our tracking of LWIN/RWIN/RSHIFT etc.
	{
		// Caller has ensured that vk has been translated from neutral to left/right if necessary
		// (e.g. VK_CONTROL -> VK_LCONTROL). For this reason, always use the parameter vk rather
		// than the raw event.vkCode.
		// Below excludes KEY_IGNORE_ALL_EXCEPT_MODIFIER since that type of event shouldn't be ignored by
		// this function.  UPDATE: KEY_PHYS_IGNORE is now considered to be something that shouldn't be
		// ignored in this case because if more than one instance has the hook installed, it is
		// possible for g_modifiersLR_logical_non_ignored to say that a key is down in one instance when
		// that instance's g_modifiersLR_logical doesn't say it's down, which is definitely wrong.  So it
		// is now omitted below:
		bool is_not_ignored = (aEvent.dwExtraInfo != KEY_IGNORE);
		bool event_is_physical = KeybdEventIsPhysical(aEvent.flags, aVK, aKeyUp);

		if (aKeyUp)
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical &= ~modLR;
				// Even if is_not_ignored == true, this is updated unconditionally on key-up events
				// to ensure that g_modifiersLR_logical_non_ignored never says a key is down when
				// g_modifiersLR_logical says its up, which might otherwise happen in cases such
				// as alt-tab.  See this comment further below, where the operative word is "relied":
				// "key pushed ALT down, or relied upon it already being down, so go up".  UPDATE:
				// The above is no longer a concern because KeyEvent() now defaults to the mode
				// which causes our var "is_not_ignored" to be true here.  Only the Send command
				// overrides this default, and it takes responsibility for ensuring that the older
				// comment above never happens by forcing any down-modifiers to be up if they're
				// not logically down as reflected in g_modifiersLR_logical.  There's more
				// explanation for g_modifiersLR_logical_non_ignored in keyboard_mouse.h:
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored &= ~modLR;
			}
			if (event_is_physical) // Note that ignored events can be physical via KEYEVENT_PHYS()
			{
				g_modifiersLR_physical &= ~modLR;
				g_PhysicalKeyState[aVK] = 0;
				// If a modifier with an available neutral VK has been released, update the state
				// of the neutral VK to be that of the opposite key (the one that wasn't released):
				switch (aVK)
				{
				case VK_LSHIFT:   g_PhysicalKeyState[VK_SHIFT] = g_PhysicalKeyState[VK_RSHIFT]; break;
				case VK_RSHIFT:   g_PhysicalKeyState[VK_SHIFT] = g_PhysicalKeyState[VK_LSHIFT]; break;
				case VK_LCONTROL: g_PhysicalKeyState[VK_CONTROL] = g_PhysicalKeyState[VK_RCONTROL]; break;
				case VK_RCONTROL: g_PhysicalKeyState[VK_CONTROL] = g_PhysicalKeyState[VK_LCONTROL]; break;
				case VK_LMENU:    g_PhysicalKeyState[VK_MENU] = g_PhysicalKeyState[VK_RMENU]; break;
				case VK_RMENU:    g_PhysicalKeyState[VK_MENU] = g_PhysicalKeyState[VK_LMENU]; break;
				}
			}
		}
		else // Modifier key was pressed down.
		{
			if (!aIsSuppressed)
			{
				g_modifiersLR_logical |= modLR;
				if (is_not_ignored)
					g_modifiersLR_logical_non_ignored |= modLR;
			}
			if (event_is_physical)
			{
				g_modifiersLR_physical |= modLR;
				g_PhysicalKeyState[aVK] = STATE_DOWN;
				// If a modifier with an available neutral VK has been pressed down (unlike LWIN & RWIN),
				// update the state of the neutral VK to be down also:
				switch (aVK)
				{
				case VK_LSHIFT:
				case VK_RSHIFT:   g_PhysicalKeyState[VK_SHIFT] = STATE_DOWN; break;
				case VK_LCONTROL:
				case VK_RCONTROL: g_PhysicalKeyState[VK_CONTROL] = STATE_DOWN; break;
				case VK_LMENU:
				case VK_RMENU:    g_PhysicalKeyState[VK_MENU] = STATE_DOWN; break;
				}
			}
		}
	} // vk is a modifier key.

	// Now that we're done using the old values (updating modifier state and calls to KeybdEventIsPhysical()
	// above used them, as well as a section higher above), update these to their new values then call
	// KeybdEventIsPhysical() again.
	sPriorVK = aVK;
	sPriorSC = aSC;
	sPriorEventWasKeyUp = aKeyUp;
	sPriorEventWasPhysical = KeybdEventIsPhysical(aEvent.flags, aVK, aKeyUp);
	sPriorEventTickCount = GetTickCount();
}



bool KeybdEventIsPhysical(DWORD aEventFlags, const vk_type aVK, bool aKeyUp)
// Always use the parameter vk rather than event.vkCode because the caller or caller's caller
// might have adjusted vk, namely to make it a left/right specific modifier key rather than a
// neutral one.
{
	// MSDN: "The keyboard input can come from the local keyboard driver or from calls to the keybd_event
	// function. If the input comes from a call to keybd_event, the input was "injected"".
	// My: This also applies to mouse events, so use it for them too:
	if (aEventFlags & LLKHF_INJECTED)
		return false;
	// So now we know it's a physical event.  But certain LSHIFT key-down events are driver-generated.
	// We want to be able to tell the difference because the Send command and other aspects
	// of keyboard functionality need us to be accurate about which keys the user is physically
	// holding down at any given time:
	if (   (aVK == VK_LSHIFT || aVK == VK_SHIFT) && !aKeyUp   ) // But not RSHIFT.
	{
		if (sNextPhysShiftDownIsNotPhys && !DualStateNumpadKeyIsDown())
		{
			sNextPhysShiftDownIsNotPhys = false;
			return false;
		}
		// Otherwise (see notes about SHIFT_KEY_WORKAROUND_TIMEOUT above for details):
		if (sPriorEventWasKeyUp && IsDualStateNumpadKey(sPriorVK, sPriorSC)
			&& (DWORD)(GetTickCount() - sPriorEventTickCount) < (DWORD)SHIFT_KEY_WORKAROUND_TIMEOUT   )
			return false;
	}

	// Otherwise, it's physical.
	// v1.0.42.04:
	// The time member of the incoming event struct has been observed to be wrongly zero sometimes, perhaps only
	// for AltGr keyboard layouts that generate LControl events when RAlt is pressed (but in such cases, I think
	// it's only sometimes zero, not always).  It might also occur during simultation of Alt+Numpad keystrokes
	// to support {Asc NNNN}.  In addition, SendInput() is documented to have the ability to set its own timestamps;
	// if it's callers put in a bad timestamp, it will probably arrive here that way too.  Thus, use GetTickCount().
	// More importantly, when a script or other application simulates an AltGr keystroke (either down or up),
	// the LControl event received here is marked as physical by the OS or keyboard driver.  This is undesirable
	// primarly because it makes g_TimeLastInputPhysical inaccurate, but also because falsely marked physical
	// events can impact the script's calls to GetKeyState("LControl", "P"), etc.
	g_TimeLastInputPhysical = GetTickCount();
	return true;
}



bool DualStateNumpadKeyIsDown()
{
	// Note: GetKeyState() might not agree with us that the key is physically down because
	// the hook may have suppressed it (e.g. if it's a hotkey).  Therefore, sPadState
	// is the only way to know for user if the user is physically holding down a *qualified*
	// Numpad key.  "Qualified" means that it must be a dual-state key and Numlock must have
	// been ON at the time the key was first pressed down.  This last criteria is needed because
	// physically holding down the shift-key will change VK generated by the driver to appear
	// to be that of the numpad without the numlock key on.  In other words, we can't just
	// consult the g_PhysicalKeyState array because it won't tell whether a key such as
	// NumpadEnd is truly phyiscally down:
	for (int i = 0; i < PAD_TOTAL_COUNT; ++i)
		if (sPadState[i])
			return true;
	return false;
}



bool IsDualStateNumpadKey(const vk_type aVK, const sc_type aSC)
{
	if (aSC & 0x100)  // If it's extended, it can't be a numpad key.
		return false;

	switch (aVK)
	{
	// It seems best to exclude the VK_DECIMAL and VK_NUMPAD0 through VK_NUMPAD9 from the below
	// list because the callers want to know whether this is a numpad key being *modified* by
	// the shift key (i.e. the shift key is being held down to temporarily transform the numpad
	// key into its opposite state, overriding the fact that Numlock is ON):
	case VK_DELETE: // NumpadDot (VK_DECIMAL)
	case VK_INSERT: // Numpad0
	case VK_END:    // Numpad1
	case VK_DOWN:   // Numpad2
	case VK_NEXT:   // Numpad3
	case VK_LEFT:   // Numpad4
	case VK_CLEAR:  // Numpad5 (this has been verified to be the VK that is sent, at least on my keyboard).
	case VK_RIGHT:  // Numpad6
	case VK_HOME:   // Numpad7
	case VK_UP:     // Numpad8
	case VK_PRIOR:  // Numpad9
		return true;
	}

	return false;
}


/////////////////////////////////////////////////////////////////////////////////////////


struct hk_sorted_type
{
	mod_type modifiers;
	modLR_type modifiersLR;
	// Keep sub-32-bit members contiguous to save memory without having to sacrifice performance of
	// 32-bit alignment:
	bool AllowExtraModifiers;
	vk_type vk;
	sc_type sc;
	HotkeyIDType id_with_flags;
};



int sort_most_general_before_least(const void *a1, const void *a2)
// The only items whose order are important are those with the same suffix.  For a given suffix,
// we want the most general modifiers (e.g. CTRL) to appear closer to the top of the list than
// those with more specific modifiers (e.g. CTRL-ALT).  To make qsort() perform properly, it seems
// best to sort by vk/sc then by generality.
{
	hk_sorted_type &b1 = *(hk_sorted_type *)a1; // For performance and convenience.
	hk_sorted_type &b2 = *(hk_sorted_type *)a2;
	// It's probably not necessary to be so thorough.  For example, if a1 has a vk but a2 has an sc,
	// those two are immediately non-equal.  But I'm worried about consistency: qsort() may get messed
	// up if these same two objects are ever compared, in reverse order, but a different comparison
	// result is returned.  Therefore, we compare rigorously and consistently:
	if (b1.vk && b2.vk)
		if (b1.vk != b2.vk)
			return b1.vk - b2.vk;
	if (b1.sc && b2.sc)
		if (b1.sc != b2.sc)
			return b1.sc - b2.sc;
	if (b1.vk && !b2.vk)
		return 1;
	if (!b1.vk && b2.vk)
		return -1;

	// If the above didn't return, we now know that a1 and a2 have the same vk's or sc's.  So
	// we use a tie-breaker to cause the most general keys to appear closer to the top of the
	// list than less general ones.  This should result in a given suffix being grouped together
	// after the sort.  Within each suffix group, the most general modifiers should appear first.

	// This part is basically saying that keys that don't allow extra modifiers can always be processed
	// after all other keys:
	if (b1.AllowExtraModifiers && !b2.AllowExtraModifiers)
		return -1;  // Indicate that a1 is smaller, so that it will go to the top.
	if (!b1.AllowExtraModifiers && b2.AllowExtraModifiers)
		return 1;

	// However the order of suffixes that don't allow extra modifiers, among themselves, may be important.
	// Thus we don't return a zero if both have AllowExtraModifiers = 0.
	// Example: User defines ^a, but also defines >^a.  What should probably happen is that >^a fores ^a
	// to fire only when <^a occurs.

	mod_type mod_a1_merged = b1.modifiers;
	mod_type mod_a2_merged = b2.modifiers;
	if (b1.modifiersLR)
		mod_a1_merged |= ConvertModifiersLR(b1.modifiersLR);
	if (b2.modifiersLR)
		mod_a2_merged |= ConvertModifiersLR(b2.modifiersLR);

	// Check for equality first to avoid a possible infinite loop where two identical sets are subsets of each other:
	if (mod_a1_merged == mod_a1_merged)
	{
		// Here refine it further to handle a case such as ^a and >^a.  We want ^a to be considered
		// more general so that it won't override >^a altogether:
		if (b1.modifiersLR && !b2.modifiersLR)
			return 1;  // Make a1 greater, so that it goes below a2 on the list.
		if (!b1.modifiersLR && b2.modifiersLR)
			return -1;
		// After the above, the only remaining possible-problem case in this block is that
		// a1 and a2 have non-zero modifiersLRs that are different.  e.g. >+^a and +>^a
		// I don't think I want to try to figure out which of those should take precedence,
		// and how they overlap.  Maybe another day.

		// v1.0.38.03: The following check is added to handle a script containing hotkeys
		// such as the following (in this order):
		// *MButton::
		// *Mbutton Up::
		// MButton::
		// MButton Up::
		// What would happen before is that the qsort() would sometimes cause "MButton Up" from the
		// list above to be processed prior to "MButton", which would set hotkey_up[*MButton's ID]
		// to be MButton Up's ID.  Then when "MButton" was processed, it would set its_table_entry
		// to MButton's ID, but hotkey_up[MButton's ID] would be wrongly left INVALID when it should
		// have received a copy of the asterisk hotkey ID's counterpart key-up ID.  However, even
		// giving it a copy would not be quite correct because then *MButton's ID would wrongly
		// be left associated with MButton's Up's ID rather than *MButton Up's.  By solving the
		// problem here in the sort rather than copying the ID, both bugs are resolved.
		if ((b1.id_with_flags & HOTKEY_KEY_UP) != (b2.id_with_flags & HOTKEY_KEY_UP))
			return (b1.id_with_flags & HOTKEY_KEY_UP) ? 1 : -1; // Put key-up hotkeys higher in the list than their down counterparts (see comment above).

		// Otherwise, consider them to be equal for the purpose of the sort:
		return 0;
	}

	mod_type mod_intersect = mod_a1_merged & mod_a2_merged;

	if (mod_a1_merged == mod_intersect)
		// a1's modifiers are containined entirely within a2's, thus a1 is more general and
		// should be considered smaller so that it will go closer to the top of the list:
		return -1;
	if (mod_a2_merged == mod_intersect)
		return 1;

	// Otherwise, since neither is a perfect subset of the other, report that they're equal.
	// More refinement might need to be done here later for modifiers that partially overlap:
	// e.g. At this point is it possible for a1's modifiersLR to be a perfect subset of a2's,
	// or vice versa?
	return 0;
}



void SetModifierAsPrefix(vk_type aVK, sc_type aSC, bool aAlwaysSetAsPrefix = false)
// The caller has already ensured that vk and/or sc is a modifier such as VK_CONTROL.
{
	if (aVK)
	{
		switch (aVK)
		{
		case VK_MENU:
		case VK_SHIFT:
		case VK_CONTROL:
			// Since the user is configuring both the left and right counterparts of a key to perform a suffix action,
			// it seems best to always consider those keys to be prefixes so that their suffix action will only fire
			// when the key is released.  That way, those keys can still be used as normal modifiers.
			// UPDATE for v1.0.29: But don't do it if there is a corresponding key-up hotkey for this neutral
			// modifier, which allows a remap such as the following to succeed:
			// Control::Send {LWin down}
			// Control up::Send {LWin up}
			if (!aAlwaysSetAsPrefix)
			{
				for (int i = 0; i < Hotkey::sHotkeyCount; ++i)
				{
					Hotkey &h = *Hotkey::shk[i]; // For performance and convenience.
					if (h.mVK == aVK && h.mKeyUp && !h.mModifiersConsolidatedLR && !h.mModifierVK && !h.mModifierSC
						&& !h.IsCompletelyDisabled())
						return; // Since caller didn't specify aAlwaysSetAsPrefix==true, don't make this key a prefix.
				}
			}
			switch (aVK)
			{
			case VK_MENU:
				kvk[VK_MENU].used_as_prefix = PREFIX_FORCED;
				kvk[VK_LMENU].used_as_prefix = PREFIX_FORCED;
				kvk[VK_RMENU].used_as_prefix = PREFIX_FORCED;
				ksc[SC_LALT].used_as_prefix = PREFIX_FORCED;
				ksc[SC_RALT].used_as_prefix = PREFIX_FORCED;
				break;
			case VK_SHIFT:
				kvk[VK_SHIFT].used_as_prefix = PREFIX_FORCED;
				kvk[VK_LSHIFT].used_as_prefix = PREFIX_FORCED;
				kvk[VK_RSHIFT].used_as_prefix = PREFIX_FORCED;
				ksc[SC_LSHIFT].used_as_prefix = PREFIX_FORCED;
				ksc[SC_RSHIFT].used_as_prefix = PREFIX_FORCED;
				break;
			case VK_CONTROL:
				kvk[VK_CONTROL].used_as_prefix = PREFIX_FORCED;
				kvk[VK_LCONTROL].used_as_prefix = PREFIX_FORCED;
				kvk[VK_RCONTROL].used_as_prefix = PREFIX_FORCED;
				ksc[SC_LCONTROL].used_as_prefix = PREFIX_FORCED;
				ksc[SC_RCONTROL].used_as_prefix = PREFIX_FORCED;
				break;
			}
			break;

		default:  // vk is a left/right modifier key such as VK_LCONTROL or VK_LWIN:
			if (aAlwaysSetAsPrefix)
				kvk[aVK].used_as_prefix = PREFIX_ACTUAL;
			else
				if (Hotkey::FindHotkeyContainingModLR(kvk[aVK].as_modifiersLR)) // Fixed for v1.0.35.13 (used to be aSC vs. aVK).
					kvk[aVK].used_as_prefix = PREFIX_ACTUAL;
				// else allow its suffix action to fire when key is pressed down,
				// under the fairly safe assumption that the user hasn't configured
				// the opposite key to also be a key-down suffix-action (but even
				// if the user has done this, it's an explicit override of the
				// safety checks here, so probably best to allow it).
		}
		return;
	}
	// Since above didn't return, using scan code instead of virtual key:
	if (aAlwaysSetAsPrefix)
		ksc[aSC].used_as_prefix = PREFIX_ACTUAL;
	else
		if (Hotkey::FindHotkeyContainingModLR(ksc[aSC].as_modifiersLR))
			ksc[aSC].used_as_prefix = PREFIX_ACTUAL;
}



void ChangeHookState(Hotkey *aHK[], int aHK_count, HookType aWhichHook, HookType aWhichHookAlways)
// Caller must verify that aWhichHook and aWhichHookAlways accurately reflect the hooks that should
// be active when we return.  For example, the caller must have already taken into account which
// hotkeys/hotstrings are suspended, disabled, etc.
//
// Caller should always be the main thread, never the hook thread.
// One reason is that this function isn't thread-safe.  Another is that new/delete/malloc/free
// themselves might not be thread-safe when the single-threaded CRT libraries are in effect
// (not using multi-threaded libraries due to a 3.5 KB increase in compressed code size).
//
// The input params are unnecessary because could just access directly by using Hotkey::shk[].
// But aHK is a little more concise.
// aWhichHookAlways was added to force the hooks to be installed (or stay installed) in the case
// of #InstallKeybdHook and #InstallMouseHook.  This is so that these two commands will always
// still be in effect even if hotkeys are suspended, so that key history can still be monitored via
// the hooks.
// Returns the set of hooks that are active after processing is complete.
{
	// v1.0.39: For simplicity and maintainability, don't even make the attempt on Win9x since it
	// seems too rare that they would have LL hook capability somehow (such as in an emultator).
	// NOTE: Some sections rely on the fact that no warning dialogs are displayed if the hook is
	// called for but the OS doesn't support it.  For example, ManifestAllHotkeysHotstringsHooks()
	// doesn't check the OS version in many cases when marking hotkeys as hook hotkeys.
	if (g_os.IsWin9x())
		return;

	// Determine the set of hooks that should be activated or deactivated.
	HookType hooks_to_be_active = aWhichHook | aWhichHookAlways; // Bitwise union.

	if (!hooks_to_be_active) // No need to check any further in this case.  Just remove all hooks.
	{
		AddRemoveHooks(0); // Remove all hooks.
		return;
	}

	// Even if hooks_to_be_active indicates no change to hook status, we still need to continue in case
	// this is a suspend or unsuspend operation.  In both of those cases, though the hook(s)
	// may already be active, the hotkey configuration probably needs to be updated.
	// Related: Even if aHK_count is zero, still want to install the hook(s) whenever
	// aWhichHookAlways specifies that they should be.  This is done so that the
	// #InstallKeybdHook and #InstallMouseHook directives can have the hooks installed just
	// for use with something such as the KeyHistory feature, or for Hotstrings, Numlock AlwaysOn,
	// the Input command, and possibly others.

	// Now we know that at least one of the hooks is a candidate for activation.
	// Set up the arrays process all of the hook hotkeys even if the corresponding hook won't
	// become active (which should only happen if g_IsSuspended is true
	// and it turns out there are no suspend-hotkeys that are handled by the hook).

	// These arrays are dynamically allocated so that memory is conserved in cases when
	// the user doesn't need the hook at all (i.e. just normal registered hotkeys).
	// This is a waste of memory if there are no hook hotkeys, but currently the operation
	// of the hook relies upon these being allocated, even if the arrays are all clean
	// slates with nothing in them (it could check if the arrays are NULL but then the
	// performance would be slightly worse for the "average" script).  Presumably, the
	// caller is requesting the keyboard hook with zero hotkeys to support the forcing
	// of Num/Caps/ScrollLock always on or off (a fairly rare situation, probably):
	if (!kvk)  // Since it's an initialzied global, this indicates that all 4 objects are not yet allocated.
	{
		if (   !(kvk = new key_type[VK_ARRAY_COUNT])
			|| !(ksc = new key_type[SC_ARRAY_COUNT])
			|| !(kvkm = new HotkeyIDType[KVKM_SIZE])
			|| !(kscm = new HotkeyIDType[KSCM_SIZE])
			|| !(hotkey_up = new HotkeyIDType[MAX_HOTKEYS])   )
		{
			// At least one of the allocations failed.
			// Keep all 4 objects in sync with one another (i.e. either all allocated, or all not allocated):
			FreeHookMem(); // Since none of the hooks is active, just free any of the above memory that was partially allocated.
			return;
		}

		// Done once immediately after allocation to init attributes such as pForceToggle and as_modifiersLR,
		// which are zero for most keys:
		ZeroMemory(kvk, VK_ARRAY_COUNT * sizeof(key_type));
		ZeroMemory(ksc, SC_ARRAY_COUNT * sizeof(key_type));

		// Below is also a one-time-only init:
		// This attribute is exists for performance reasons (avoids a function call in the hook
		// procedure to determine this value):
		kvk[VK_CONTROL].as_modifiersLR = MOD_LCONTROL | MOD_RCONTROL;
		kvk[VK_LCONTROL].as_modifiersLR = MOD_LCONTROL;
		kvk[VK_RCONTROL].as_modifiersLR = MOD_RCONTROL;
		kvk[VK_MENU].as_modifiersLR = MOD_LALT | MOD_RALT;
		kvk[VK_LMENU].as_modifiersLR = MOD_LALT;
		kvk[VK_RMENU].as_modifiersLR = MOD_RALT;
		kvk[VK_SHIFT].as_modifiersLR = MOD_LSHIFT | MOD_RSHIFT;
		kvk[VK_LSHIFT].as_modifiersLR = MOD_LSHIFT;
		kvk[VK_RSHIFT].as_modifiersLR = MOD_RSHIFT;
		kvk[VK_LWIN].as_modifiersLR = MOD_LWIN;
		kvk[VK_RWIN].as_modifiersLR = MOD_RWIN;

		// This is a bit iffy because it's far from certain that these particular scan codes
		// are really modifier keys on anything but a standard English keyboard.  However,
		// at the very least the Win9x version must rely on something like this because a
		// low-level hook can't be used under Win9x, and a high-level hook doesn't receive
		// the left/right VKs at all (so the scan code must be used to tell them apart).
		// However: it might be possible under Win9x to use MapVirtualKey() or some similar
		// function to verify, at runtime, that the expected scan codes really do map to the
		// expected VK.  If not, perhaps MapVirtualKey() or such can be used to search through
		// every scan code to find out which map to VKs that are modifiers.  Any such keys
		// found can then be initialized similar to below:
		ksc[SC_LCONTROL].as_modifiersLR = MOD_LCONTROL;
		ksc[SC_RCONTROL].as_modifiersLR = MOD_RCONTROL;
		ksc[SC_LALT].as_modifiersLR = MOD_LALT;
		ksc[SC_RALT].as_modifiersLR = MOD_RALT;
		ksc[SC_LSHIFT].as_modifiersLR = MOD_LSHIFT;
		ksc[SC_RSHIFT].as_modifiersLR = MOD_RSHIFT;
		ksc[SC_LWIN].as_modifiersLR = MOD_LWIN;
		ksc[SC_RWIN].as_modifiersLR = MOD_RWIN;

		// Use the address rather than the value, so that if the global var's value
		// changes during runtime, ours will too:
		kvk[VK_SCROLL].pForceToggle = &g_ForceScrollLock;
		kvk[VK_CAPITAL].pForceToggle = &g_ForceCapsLock;
		kvk[VK_NUMLOCK].pForceToggle = &g_ForceNumLock;
	}

	// Init only those attributes which reflect the hotkey's definition, not those that reflect
	// the key's current status (since those are intialized only if the hook state is changing
	// from OFF to ON (later below):
	int i;
	for (i = 0; i < VK_ARRAY_COUNT; ++i)
		RESET_KEYTYPE_ATTRIB(kvk[i])
	for (i = 0; i < SC_ARRAY_COUNT; ++i)
		RESET_KEYTYPE_ATTRIB(ksc[i]) // Note: ksc not kvk.

	// Indicate here which scan codes should override their virtual keys:
	for (i = 0; i < g_key_to_sc_count; ++i)
		if (g_key_to_sc[i].sc > 0 && g_key_to_sc[i].sc <= SC_MAX)
			ksc[g_key_to_sc[i].sc].sc_takes_precedence = true;

	// These have to be initialized with with element value INVALID.
	// Don't use FillMemory because the array elements are too big (bigger than bytes):
	for (i = 0; i < KVKM_SIZE; ++i) // Simplify by viewing 2-dimensional array as a 1-dimensional array.
		kvkm[i] = HOTKEY_ID_INVALID;
	for (i = 0; i < KSCM_SIZE; ++i) // Simplify by viewing 2-dimensional array as a 1-dimensional array.
		kscm[i] = HOTKEY_ID_INVALID;
	for (i = 0; i < MAX_HOTKEYS; ++i)
		hotkey_up[i] = HOTKEY_ID_INVALID;

	hk_sorted_type hk_sorted[MAX_HOTKEYS];
	ZeroMemory(hk_sorted, sizeof(hk_sorted));
	int hk_sorted_count = 0;
	key_type *pThisKey = NULL;
	for (i = 0; i < aHK_count; ++i)
	{
		Hotkey &hk = *aHK[i]; // For performance and convenience.

		// If it's not a hook hotkey (e.g. it was already registered with RegisterHotkey() or it's a joystick
		// hotkey) don't process it here.  Similarly, if g_IsSuspended is true, we won't include it unless it's
		// exempt from suspension:
		if (   !HK_TYPE_IS_HOOK(hk.mType)
			|| (g_IsSuspended && !hk.IsExemptFromSuspend())
			|| hk.IsCompletelyDisabled()   ) // Listed last for short-circuit performance.
			continue;

		// Rule out the possibility of obnoxious values right away, preventing array-out-of bounds, etc.:
		if ((!hk.mVK && !hk.mSC) || hk.mVK > VK_MAX || hk.mSC > SC_MAX)
			continue;

		if (!hk.mVK)
		{
			// scan codes don't need something like the switch stmt below because they can't be neutral.
			// In other words, there's no scan code equivalent for something like VK_CONTROL.
			// In addition, SC_LCONTROL, for example, doesn't also need to change the kvk array
			// for VK_LCONTROL because the hook knows to give the scan code precedence, and thus
			// look it up only in the ksc array in that case.
			pThisKey = ksc + hk.mSC;
			// For some scan codes this was already set above.  But to support explicit scan code hotkeys,
			// such as "SC102::MsgBox", make sure it's set for every hotkey that uses an explicit scan code.
			pThisKey->sc_takes_precedence = true;
		}
		else
		{
			pThisKey = kvk + hk.mVK;
			// Keys that have a neutral as well as a left/right counterpart must be
			// fully initialized since the hook can receive the left, the right, or
			// the neutral (neutral only if another app calls KeyEvent(), probably).
			// There are several other switch stmts in this function like the below
			// that serve a similar purpose.  The alternative to doing all these
			// switch stmts is to always translate left/right vk's (whose sc's don't
			// take precedence) in the KeyboardProc() itself.  But that would add
			// the overhead of a switch stmt to *every* keypress ever made on the
			// system, so it seems better to set up everything correctly here since
			// this init section is only done once.
			// Note: These switch stmts probably aren't needed under Win9x since I think
			// it might be impossible for them to receive something like VK_LCONTROL,
			// except *possibly* if keybd_event() is explicitly called with VK_LCONTROL
			// and (might want to verify that -- if true, might want to keep the switches
			// even for Win9x for safety and in case Win9x ever gets overhauled and
			// improved in some future era, or in case Win9x is running in an emulator
			// that expands its capabilities.
			switch (hk.mVK)
			{
			case VK_MENU:
				// It's not strictly necessary to init all of these, since the
				// hook currently never handles VK_RMENU, for example, by its
				// vk (it uses sc instead).  But it's safest to do all of them
				// in case future changes ever ruin that assumption:
				kvk[VK_LMENU].used_as_suffix = true;
				kvk[VK_RMENU].used_as_suffix = true;
				ksc[SC_LALT].used_as_suffix = true;
				ksc[SC_RALT].used_as_suffix = true;
				kvk[VK_LMENU].used_as_key_up = hk.mKeyUp;
				kvk[VK_RMENU].used_as_key_up = hk.mKeyUp;
				ksc[SC_LALT].used_as_key_up = hk.mKeyUp;
				ksc[SC_RALT].used_as_key_up = hk.mKeyUp;
				break;
			case VK_SHIFT:
				// The neutral key itself is also set to be a suffix further below.
				kvk[VK_LSHIFT].used_as_suffix = true;
				kvk[VK_RSHIFT].used_as_suffix = true;
				ksc[SC_LSHIFT].used_as_suffix = true;
				ksc[SC_RSHIFT].used_as_suffix = true;
				kvk[VK_LSHIFT].used_as_key_up = hk.mKeyUp;
				kvk[VK_RSHIFT].used_as_key_up = hk.mKeyUp;
				ksc[SC_LSHIFT].used_as_key_up = hk.mKeyUp;
				ksc[SC_RSHIFT].used_as_key_up = hk.mKeyUp;
				break;
			case VK_CONTROL:
				kvk[VK_LCONTROL].used_as_suffix = true;
				kvk[VK_RCONTROL].used_as_suffix = true;
				ksc[SC_LCONTROL].used_as_suffix = true;
				ksc[SC_RCONTROL].used_as_suffix = true;
				kvk[VK_LCONTROL].used_as_key_up = hk.mKeyUp;
				kvk[VK_RCONTROL].used_as_key_up = hk.mKeyUp;
				ksc[SC_LCONTROL].used_as_key_up = hk.mKeyUp;
				ksc[SC_RCONTROL].used_as_key_up = hk.mKeyUp;
				break;
			// Later might want to add cases for VK_LCONTROL and such, but for right now,
			// these keys should never come up since they're done by scan code?
			}
		}

		pThisKey->used_as_suffix = true;
		HotkeyIDType hotkey_id_with_flags = hk.mID;

		if (hk.mKeyUp)
		{
			pThisKey->used_as_key_up = true;
			hotkey_id_with_flags |= HOTKEY_KEY_UP;
		}

		// If this is a naked (unmodified) modifier key, make it a prefix if it ever modifies any
		// other hotkey.  This processing might be later combined with the hotkeys activation function
		// to eliminate redundancy / improve efficiency, but then that function would probably need to
		// init everything else here as well:
		if (pThisKey->as_modifiersLR && !hk.mModifiersConsolidatedLR && !hk.mModifierVK && !hk.mModifierSC
			&& !(hk.mNoSuppress & AT_LEAST_ONE_VARIANT_HAS_TILDE)) // v1.0.45.02: ~Alt, ~Control, etc. should fire upon press-down, not release (broken by 1.0.44's PREFIX_FORCED, but I think it was probably broken in pre-1.0.41 too).
			SetModifierAsPrefix(hk.mVK, hk.mSC);

		if (hk.mModifierVK)
		{
			if (kvk[hk.mModifierVK].as_modifiersLR)
				// The hotkey's ModifierVK is itself a modifier.
				SetModifierAsPrefix(hk.mModifierVK, 0, true);
			else
			{
				kvk[hk.mModifierVK].used_as_prefix = PREFIX_ACTUAL;
				if (hk.mNoSuppress & NO_SUPPRESS_PREFIX)
					kvk[hk.mModifierVK].no_suppress |= NO_SUPPRESS_PREFIX;
			}
			if (pThisKey->nModifierVK < MAX_MODIFIER_VKS_PER_SUFFIX)  // else currently no error-reporting.
			{
				pThisKey->ModifierVK[pThisKey->nModifierVK].vk = hk.mModifierVK;
				if (hk.mHookAction)
					pThisKey->ModifierVK[pThisKey->nModifierVK].id_with_flags = hk.mHookAction;
				else
					pThisKey->ModifierVK[pThisKey->nModifierVK].id_with_flags = hotkey_id_with_flags;
				++pThisKey->nModifierVK;
				continue;
			}
		}
		else
		{
			if (hk.mModifierSC)
			{
				if (ksc[hk.mModifierSC].as_modifiersLR)  // Fixed for v1.0.35.13 (used to be kvk vs. ksc).
					// The hotkey's ModifierSC is itself a modifier.
					SetModifierAsPrefix(0, hk.mModifierSC, true);
				else
				{
					ksc[hk.mModifierSC].used_as_prefix = PREFIX_ACTUAL;
					if (hk.mNoSuppress & NO_SUPPRESS_PREFIX)
						ksc[hk.mModifierSC].no_suppress |= NO_SUPPRESS_PREFIX;
					// For some scan codes this was already set above.  But to support explicit scan code prefixes,
					// such as "SC118 & SC122::MsgBox", make sure it's set for every prefix that uses an explicit
					// scan code:
					ksc[hk.mModifierSC].sc_takes_precedence = true;
				}
				if (pThisKey->nModifierSC < MAX_MODIFIER_SCS_PER_SUFFIX)  // else currently no error-reporting.
				{
					pThisKey->ModifierSC[pThisKey->nModifierSC].sc = hk.mModifierSC;
					if (hk.mHookAction)
						pThisKey->ModifierSC[pThisKey->nModifierSC].id_with_flags = hk.mHookAction;
					else
						pThisKey->ModifierSC[pThisKey->nModifierSC].id_with_flags = hotkey_id_with_flags;
					++pThisKey->nModifierSC;
					continue;
				}
			}
		}

		// At this point, since the above didn't "continue", this hotkey is one without a ModifierVK/SC.
		// Put it into a temporary array, which will be later sorted:
		hk_sorted[hk_sorted_count].id_with_flags = hk.mHookAction ? hk.mHookAction : hotkey_id_with_flags;
		hk_sorted[hk_sorted_count].vk = hk.mVK;
		hk_sorted[hk_sorted_count].sc = hk.mSC;
		hk_sorted[hk_sorted_count].modifiers = hk.mModifiers;
		hk_sorted[hk_sorted_count].modifiersLR = hk.mModifiersLR;
		hk_sorted[hk_sorted_count].AllowExtraModifiers = hk.mAllowExtraModifiers;
		++hk_sorted_count;
	}

	if (hk_sorted_count)
	{
		// It's necessary to get them into this order to avoid problems that would be caused by
		// AllowExtraModifiers:
		qsort((void *)hk_sorted, hk_sorted_count, sizeof(hk_sorted_type), sort_most_general_before_least);

		// For each hotkey without a ModifierVK/SC (which override normal modifiers), expand its modifiers and
		// modifiersLR into its column in the kvkm or kscm arrays.

		mod_type modifiers, i_modifiers_merged;
		int modifiersLR;  // Don't make this modLR_type to avoid integer overflow, since it's a loop-counter.
		bool prev_hk_is_key_up, this_hk_is_key_up;
		HotkeyIDType this_hk_id;

		for (i = 0; i < hk_sorted_count; ++i)
		{
			hk_sorted_type &this_hk = hk_sorted[i]; // For performance and convenience.
			this_hk_is_key_up = this_hk.id_with_flags & HOTKEY_KEY_UP;
			this_hk_id = this_hk.id_with_flags & HOTKEY_ID_MASK;

			i_modifiers_merged = this_hk.modifiers;
			if (this_hk.modifiersLR)
				i_modifiers_merged |= ConvertModifiersLR(this_hk.modifiersLR);

			for (modifiersLR = 0; modifiersLR <= MODLR_MAX; ++modifiersLR)  // For each possible LR value.
			{
				modifiers = ConvertModifiersLR(modifiersLR);
				if (this_hk.AllowExtraModifiers)
				{
					// True if modifiersLR is a superset of i's modifier value.  In other words,
					// modifiersLR has the minimum required keys but also has some
					// extraneous keys, which are allowed in this case:
					if (i_modifiers_merged != (modifiers & i_modifiers_merged))
						continue;
				}
				else
					if (i_modifiers_merged != modifiers)
						continue;

				// In addition to the above, modifiersLR must also have the *specific* left or right keys
				// found in i's modifiersLR.  In other words, i's modifiersLR must be a perfect subset
				// of modifiersLR:
				if (this_hk.modifiersLR) // make sure that any more specific left/rights are also present.
					if (this_hk.modifiersLR != (modifiersLR & this_hk.modifiersLR))
						continue;

				// If above didn't "continue", modifiersLR is a valid hotkey combination so set it as such:
				if (!this_hk.vk)
				{
					// scan codes don't need the switch() stmt below because, for example,
					// the hook knows to look up left-control by only SC_LCONTROL,
					// not VK_LCONTROL.
					HotkeyIDType &its_table_entry = Kscm(modifiersLR, this_hk.sc);
					if (its_table_entry == HOTKEY_ID_INVALID) // Since there is no ID currently in the slot, key-up/down doesn't matter.
						its_table_entry = this_hk.id_with_flags;
					else
					{
						prev_hk_is_key_up = its_table_entry & HOTKEY_KEY_UP;
						// Known limitation for a set of hotkeys such as the following:
						// *MButton::
						// *Mbutton Up::
						// MButton Up::  ; This is the key point: that this hotkey lacks a counterpart down-key.
						// Because there's no down-counterpart to the non-asterisk hotkey, the non-asterik
						// hotkey's MButton Up takes over completely and *MButton is ignored.  This is because
						// a given hotkey ID can only have one entry in the hotkey_up array.  What should
						// really happen is that every Up hotkey should have an implicit identical down hotkey
						// just for the purpose of having a unique ID in the hotkey_up array.  But that seems
						// like too much code given the rarity of doing something like this, especially since
						// it can be easily avoided simply by defining MButton:: as a hotkey in the script.
						if (this_hk_is_key_up && !prev_hk_is_key_up) // Override any existing key-up hotkey for this down hotkey ID, e.g. "LButton Up" takes precedence over "*LButton Up".
							hotkey_up[its_table_entry & HOTKEY_ID_MASK] = this_hk.id_with_flags;
						else if (!this_hk_is_key_up && prev_hk_is_key_up)
						{
							// Swap them so that the down-hotkey is in the main array and the up in the secondary:
							hotkey_up[this_hk_id] = its_table_entry;
							its_table_entry = this_hk.id_with_flags;
						}
						else // Either both are key-up hotkeys or both are key-down:
						{
							// Fix for v1.0.40.09: Also copy the previous hotkey's corresponding up-hotkey (if any)
							// so that this hotkey will have that same one.  This also solves the issue of a hotkey
							// such as "^!F1" firing twice (once for down and once for up) when "*F1" and "*F1 up"
							// are both hotkeys.  Instead, the "*F1 up" hotkey should fire upon release of "^!F1"
							// so that the behavior is consistent with the case where "*F1" isn't present as a hotkey.
							// This fix doesn't appear to break anything else, most notably it still allows a hotkey
							// such as "^!F1 up" to take precedence over "*F1 up" because in such a case, this
							// code would never have executed because prev_hk_is_key_up would be true but
							// this_hk_is_key_up would be false.  Note also that sort_most_general_before_least()
							// has put key-up hotkeys above their key-down counterparts in the list.
							hotkey_up[this_hk_id] = hotkey_up[its_table_entry & HOTKEY_ID_MASK]; // Must be done prior to next line.
							its_table_entry = this_hk.id_with_flags;
						}
					}
				}
				else // This hotkey is a virtual key (non-scan code) hotkey, which is more typical.
				{
					bool do_cascade = true;
					HotkeyIDType &its_table_entry = Kvkm(modifiersLR, this_hk.vk);
					if (its_table_entry == HOTKEY_ID_INVALID) // Since there is no ID currently in the slot, key-up/down doesn't matter.
						its_table_entry = this_hk.id_with_flags;
					else
					{
						prev_hk_is_key_up = its_table_entry & HOTKEY_KEY_UP;
						if (this_hk_is_key_up && !prev_hk_is_key_up) // Override any existing key-up hotkey for this down hotkey ID, e.g. "LButton Up" takes precedence over "*LButton Up".
						{
							hotkey_up[its_table_entry & HOTKEY_ID_MASK] = this_hk.id_with_flags;
							do_cascade = false;  // Every place the down-hotkey ID already appears, it will point to this same key-up hotkey.
						}
						else if (!this_hk_is_key_up && prev_hk_is_key_up)
						{
							// Swap them so that the down-hotkey is in the main array and the up in the secondary:
							hotkey_up[this_hk_id] = its_table_entry;
							its_table_entry = this_hk.id_with_flags;
						}
						else // Either both are key-up hotkeys or both are key-down:
						{
							hotkey_up[this_hk_id] = hotkey_up[its_table_entry & HOTKEY_ID_MASK]; // v1.0.40.09: See comments at similar section above.
							its_table_entry = this_hk.id_with_flags;
						}
					}
					
					if (do_cascade)
					{
						switch (this_hk.vk)
						{
						case VK_MENU:
						case VK_LMENU: // In case the program is ever changed to support these VKs directly.
							Kvkm(modifiersLR, VK_LMENU) = this_hk.id_with_flags;
							Kscm(modifiersLR, SC_LALT) = this_hk.id_with_flags;
							if (this_hk.vk == VK_LMENU)
								break;
							//else fall through so that VK_MENU also gets the right side set below:
						case VK_RMENU:
							Kvkm(modifiersLR, VK_RMENU) = this_hk.id_with_flags;
							Kscm(modifiersLR, SC_RALT) = this_hk.id_with_flags;
							break;
						case VK_SHIFT:
						case VK_LSHIFT:
							Kvkm(modifiersLR, VK_LSHIFT) = this_hk.id_with_flags;
							Kscm(modifiersLR, SC_LSHIFT) = this_hk.id_with_flags;
							if (this_hk.vk == VK_LSHIFT)
								break;
							//else fall through so that VK_SHIFT also gets the right side set below:
						case VK_RSHIFT:
							Kvkm(modifiersLR, VK_RSHIFT) = this_hk.id_with_flags;
							Kscm(modifiersLR, SC_RSHIFT) = this_hk.id_with_flags;
							break;
						case VK_CONTROL:
						case VK_LCONTROL:
							Kvkm(modifiersLR, VK_LCONTROL) = this_hk.id_with_flags;
							Kscm(modifiersLR, SC_LCONTROL) = this_hk.id_with_flags;
							if (this_hk.vk == VK_LCONTROL)
								break;
							//else fall through so that VK_CONTROL also gets the right side set below:
						case VK_RCONTROL:
							Kvkm(modifiersLR, VK_RCONTROL) = this_hk.id_with_flags;
							Kscm(modifiersLR, SC_RCONTROL) = this_hk.id_with_flags;
							break;
						} // switch()
					} // if (do_cascade)
				} // this hotkey is a scan code hotkey.
			}
		}
	}

	// Add or remove hooks, as needed.  No change is made if the hooks are already in the correct state.
	AddRemoveHooks(hooks_to_be_active);
}



void AddRemoveHooks(HookType aHooksToBeActive, bool aChangeIsTemporary)
// Caller has already ensured that OS isn't Win9x.
// Caller has ensured that any static memory arrays used by the hook functions have been allocated.
// Caller is always the main thread, never the hook thread because this function isn't thread-safe
// and it also calls PeekMessage() for the main thread.
{
	HookType hooks_active_orig = GetActiveHooks();
	if (aHooksToBeActive == hooks_active_orig) // It's already in the right state.
		return;

	// It's done the following way because:
	// It's unclear that zero is always an invalid thread ID (not even GetWindowThreadProcessId's
	// documentation gives any hint), so its safer to assume that a thread ID can be zero and yet still valid.
	static HANDLE sThreadHandle = NULL;

	if (!hooks_active_orig) // Neither hook is active now but at least one will be or the above would have returned.
	{
		// Assert: sThreadHandle should be NULL at this point.  The only way this isn't true is if
		// a previous call to AddRemoveHooks() timed out while waiting for the hook thread to exit,
		// which seems far too rare to add extra code for.

		// CreateThread() vs. _beginthread():
		// It's not necessary to link to the multi-threading C runtime (which bloats the code by 3.5 KB
		// compressed) as long as the new thread doesn't call any C-library functions that aren't thread-safe
		// (in addition to the C functions that obviously use static data, calls to things like malloc(),
		// new, and other memory management functions probably aren't thread-safe unless the multi-threaded
		// library is used). The memory leak described in MSDN for ExitThread() applies only to the
		// multi-threaded libraries (multiple sources confirm this), so it isn't a concern either.
		// That's true even if the program is linked against the multi-threaded DLLs (MSVCRT.dll) rather
		// than the libraries (e.g. for a minimum-sized SC.bin file), as confirmed by the following quotes:
		// "This applies only to the static-link version of the runtime. For this and other reasons, I
		// *highly* recommend using the DLL runtime, which lets you use CreateThread() without prejudice.
		// Confirmation from MSDN: "Another work around is to link the *executable* to the CRT in a *DLL*
		// instead of the static CRT."
		//
		// The hooks are designed to make miminmal use of C-library calls, currently calling only things
		// like memcpy() and strlen(), which are thread safe in the single-threaded library (according to
		// their source code).  However, the hooks may indirectly call other library functions via calls
		// to KeyEvent() and other functions, which has already been reviewed for thread-safety but needs
		// to be kept in mind as changes are made in the future.
		//
		// CreateThread's second parameter is the new thread's initial stack size. The stack will grow
		// automatically if more is needed, so it's kept small here to greatly reduce the amount of
		// memory used by the hook thread.  The XP Task Manager's "VM Size" column (which seems much
		// more accurate than "Mem Usage") indicates that a new thread consumes 28 KB + its stack size.
		if (!aChangeIsTemporary) // Caller has ensured that thread already exists when aChangeIsTemporary==true.
			if (sThreadHandle = CreateThread(NULL, 8*1024, HookThreadProc, NULL, 0, &g_HookThreadID)) // Win9x: Last parameter cannot be NULL.
				SetThreadPriority(sThreadHandle, THREAD_PRIORITY_TIME_CRITICAL); // See below for explanation.
			// The above priority level seems optimal because if some other process has high priority,
			// the keyboard and mouse hooks will still take precedence, which avoids the mouse cursor
			// and keystroke lag that would otherwise occur (confirmed through testing).  Due to their
			// return-ASAP nature, the hooks are an ideal candidate for almost-realtime priority because
			// they run only rarely and only for tiny bursts of time.
			// Note that the above must also be done in such a way that it works on NT4, which doesn't support
			// below-normal and above-normal process priorities, nor perhaps other aspects of priority.
			// So what is the actual priority given to the hooks by the OS?  Assuming that the script's
			// process is set to NORMAL_PRIORITY_CLASS (which is the default), the following applies:
			// First of all, a definition: "base priority" is the actual/net priority of the thread.
			// It determines how the OS will schedule a thread relative to all other threads on the system.
			// So in a sense, if you look only at base priority, the thread's process's priority has no
			// bearing on how the thread will get scheduled (except to the extent that it contributes
			// to the calculation of the base priority itself).  Here are some common base priorities
			// along with where the hook priority (15) fits in:
			// 7 = NORMAL_PRIORITY_CLASS process + THREAD_PRIORITY_NORMAL thread.
			// 9 = NORMAL_PRIORITY_CLASS process + THREAD_PRIORITY_HIGHEST thread.
			// 13 = HIGH_PRIORITY_CLASS process + THREAD_PRIORITY_NORMAL thread.
			// 15 = (ANY)_PRIORITY_CLASS process + THREAD_PRIORITY_TIME_CRITICAL thread. <-- Seems like the optimal compromise.
			// 15 = HIGH_PRIORITY_CLASS process + THREAD_PRIORITY_HIGHEST thread.
			// 24 = REALTIME_PRIORITY_CLASS process + THREAD_PRIORITY_NORMAL thread.
		else // Failed to create thread.  Seems to rare to justify the display of an error.
		{
			FreeHookMem(); // If everything's designed right, there should be no hooks now (even if therre is, they can't be functional because their thread is nonexistent).
			return;
		}
	}
	//else there is at least one hook already active, which guarantees that the hook thread exists (assuming
	// everything is designed right).

	// Above has ensured that the hook thread now exists, so send it the status-change message.

	// Post the AHK_CHANGE_HOOK_STATE message to the new thread to put the right hooks into effect.
	// If both hooks are to be deactivated, AHK_CHANGE_HOOK_STATE also causes the hook thread to exit.
	// PostThreadMessage() has been observed to fail, such as when a script replaces a previous instance
	// of itself via #SingleInstance.  I think this happens because the new thread hasn't yet had a
	// chance to create its message queue via GetMessage().  So rather than using something like
	// WaitForSingleObject() -- which might not be reliable due to split-second timing of when the
	// queue actually gets created -- just keep retrying until time-out or PostThreadMessage() succeeds.
	for (int i = 0; i < 50 && !PostThreadMessage(g_HookThreadID, AHK_CHANGE_HOOK_STATE, aHooksToBeActive, !aChangeIsTemporary); ++i)
		Sleep(10); // Should never execute if thread already existed before this function was called.
		// Above: Sleep(10) seems better than Sleep(0), which would max the CPU while waiting.
		// MUST USE Sleep vs. MsgSleep, otherwise an infinite recursion of ExitApp is possible.
		// This can be reproduced by running a script consisting only of the line #InstallMouseHook
		// and then exiting via the tray menu.  I tried fixing it in TerminateApp with the following,
		// but it's just not enough.  So rather than spend a long time on it, it's fixed directly here:
			// Because of the below, our callers must NOT assume that an exit will actually take place.
			//static is_running = false;
			//if (is_running)
			//	return OK;
			//is_running = true; // Since we're exiting, there should be no need to set it to false further below.

	// If it times out I think it's realistically impossible that the new thread really exists because
	// if it did, it certainly would have had time to execute GetMessage() in all but extreme/theoretical
	// cases.  Therefore, no thread check/termination attempt is done.  Alternatively, a check for
	// GetExitCodeThread() could be done followed by closing the handle and setting it to NULL, but once
	// again the code size doesn't seem worth it for a situation that is probably impossible.
	//
	// Also, a timeout itself seems too rare (perhaps even impossible) to justify a warning dialog.
	// So do nothing, which retains the current values of g_KeybdHook and g_MouseHook.

	// For safety, serialize the termination of the hook thread so that this function can't be called
	// again by the main thread before the hook thread has had a chance to exit in response to the
	// previous call.  This improves reliability, especially by ensuring a clean exit (if our caller
	// is about to exit the app via exit(), which otherwise might not cleanly close all threads).
	// UPDATE: Also serialize all changes to the hook status so that our caller can rely on the new
	// hook state being in effect immediately.  For example, the Input command installs the keyboard
	// hook and it's more maintainable if we ensure the status is correct prior to returning.
	MSG msg;
	DWORD exit_code, start_time;
	bool problem_activating_hooks;
	for (problem_activating_hooks = false, start_time = GetTickCount();;) // For our caller, wait for hook thread to update the status of the hooks.
	{
		if (aHooksToBeActive) // Wait for the hook thread to activate the specified hooks.
		{
			// In this mode, the hook thread knows we want a report of success or failure via message.
			if (PeekMessage(&msg, NULL, AHK_CHANGE_HOOK_STATE, AHK_CHANGE_HOOK_STATE, PM_REMOVE))
			{
				if (msg.wParam) // The hook thread indicated failure to activate one or both of the hooks.
				{
					// This is done so that the MsgBox warning won't be shown until after these loops finish,
					// which seems safer to prevent any parts of the script from running as a result
					// the MsgBox pumping hotkey messages and such, which could result in a script
					// subroutine launching while we're in here:
					problem_activating_hooks = true;
					if (!GetActiveHooks() && !aChangeIsTemporary) // The failure is such that no hooks are now active, and thus (due to the mode) the hook thread will exit.
					{
						// Convert this loop into the mode that waits for the hook thread to exit.
						// This allows the thead handle to be closed and the memory to be freed.
						aHooksToBeActive = 0;
						continue;
					}
					// It failed but one hook is still active, or the change is temporary.  Either way,
					// we're done waiting.  Fall through to "break" below.
				}
				//else it successfully changed the state.
				// In either case, we're done waiting:
				break;
			}
			//else no AHK_CHANGE_HOOK_STATE message has arrived yet, so keep waiting until it does or timeout occurs.
		}
		else // The hook thread has been asked to deactivate both hooks.
		{
			if (aChangeIsTemporary) // The thread will not terminate in this mode, it will just remove its hooks.
			{
				if (!GetActiveHooks()) // The hooks have been deactivated.
					break; // Don't call FreeHookMem() because caller doesn't want that when aChangeIsTemporary==true.
			}
			else // Wait for the thread to terminate.
			{
				GetExitCodeThread(sThreadHandle, &exit_code);
				if (exit_code != STILL_ACTIVE) // The hook thread is now gone.
				{
					// Do the following only if it actually exited (i.e. not if this loop timed out):
					CloseHandle(sThreadHandle); // Release our refererence to it to allow the OS to delete the thread object.
					sThreadHandle = NULL;
					FreeHookMem(); // There should be no hooks now (even if therre is, they can't be functional because their thread is nonexistent).
					break;
				}
			}
		}
		if (GetTickCount() - start_time > 500) // DWORD subtraction yields correct result even when TickCount has wrapped.
			break;
		// v1.0.43: The following sleeps for 0 rather than some longer time because:
		// 1) In nearly all cases, this loop should do only one iteration because a Sleep(0) should guaranty
		//    that the hook thread will get a timeslice before our thread gets another.  In fact, it might not
		//    do any iterations if the system preempts the main thread immediately when a message is posted to
		//    a higher priority thread (especially one in its own process).
		// 2) SendKeys()'s SendInput mode relies on fast removal of hook to prevent a 10ms or longer delay before
		//    the keystrokes get sent.  Such a delay would be quite undesirable in cases where response time is
		//    critical, such as in games.
		// Testing shows that removing the Sleep() entirely does not help performance.  The following was measured
		// when the CPU was under heavy load from a cpu-maxing utility:
		//   Loop 10  ; Keybd hook must be installed for this test to be meaningful.
		//      SendInput {Shift}
		Sleep(0); // Not MsgSleep (see the "Sleep(10)" above for why).
	}
	// If the above loop timed out without the hook thread exiting (if it was asked to exit), sThreadHandle
	// is left as non-NULL to reflect this condition.

	// In case mutex create/open/close can be a high-overhread operation, do it only when the hook isn't
	// being quickly/temporarily removed then added back again.
	if (!aChangeIsTemporary)
	{
		if (g_KeybdHook && !(hooks_active_orig & HOOK_KEYBD)) // The keyboard hook has been newly added.
			sKeybdMutex = CreateMutex(NULL, FALSE, KEYBD_MUTEX_NAME); // Create-or-open this mutex and have it be unowned.
		else if (!g_KeybdHook && (hooks_active_orig & HOOK_KEYBD))  // The keyboard hook has been newly removed.
		{
			CloseHandle(sKeybdMutex);
			sKeybdMutex = NULL;
		}
		if (g_MouseHook && !(hooks_active_orig & HOOK_MOUSE)) // The mouse hook has been newly added.
			sMouseMutex = CreateMutex(NULL, FALSE, MOUSE_MUTEX_NAME); // Create-or-open this mutex and have it be unowned.
		else if (!g_MouseHook && (hooks_active_orig & HOOK_MOUSE))  // The mouse hook has been newly removed.
		{
			CloseHandle(sMouseMutex);
			sMouseMutex = NULL;
		}
	}

	// For maintainability, it seems best to display the MsgBox only at the very end.
	if (problem_activating_hooks)
	{
		// Prevent hotkeys and other subroutines from running (which could happen via MsgBox's message pump)
		// to avoid the possibility that the script will continue to call this function recursively, resulting
		// in an infinite stack of MsgBoxes. This approach is similar to that used in Hotkey::Perform()
		// for the #MaxHotkeysPerInterval warning dialog:
		g_AllowInterruption = FALSE; 
		// Below is a generic message to reduce code size.  Failure is rare, but has been known to happen when
		// certain types of games are running).
		MsgBox("Warning: The keyboard and/or mouse hook could not be activated; "
			"some parts of the script will not function.");
		g_AllowInterruption = TRUE;
	}
}



bool SystemHasAnotherKeybdHook()
{
	if (sKeybdMutex)
		CloseHandle(sKeybdMutex); // But don't set it to NULL because we need its value below as a flag.
	HANDLE mutex = CreateMutex(NULL, FALSE, KEYBD_MUTEX_NAME); // Create() vs. Open() has enough access to open the mutex if it exists.
	DWORD last_error = GetLastError();
	// Don't check g_KeybdHook because in the case of aChangeIsTemporary, it might be NULL even though
	// we want a handle to the mutex maintained here.
	if (sKeybdMutex) // It was open originally, so update the handle the the newly opened one.
		sKeybdMutex = mutex;
	else // Keep it closed because the system tracks how many handles there are, deleting the mutex when zero.
		CloseHandle(mutex);  // This facilitates other instances of the program getting the proper last_error value.
	return last_error == ERROR_ALREADY_EXISTS;
}



bool SystemHasAnotherMouseHook()
{
	if (sMouseMutex)
		CloseHandle(sMouseMutex); // But don't set it to NULL because we need its value below as a flag.
	HANDLE mutex = CreateMutex(NULL, FALSE, MOUSE_MUTEX_NAME); // Create() vs. Open() has enough access to open the mutex if it exists.
	DWORD last_error = GetLastError();
	// Don't check g_MouseHook because in the case of aChangeIsTemporary, it might be NULL even though
	// we want a handle to the mutex maintained here.
	if (sMouseMutex) // It was open originally, so update the handle the the newly opened one.
		sMouseMutex = mutex;
	else // Keep it closed because the system tracks how many handles there are, deleting the mutex when zero.
		CloseHandle(mutex);  // This facilitates other instances of the program getting the proper last_error value.
	return last_error == ERROR_ALREADY_EXISTS;
}



DWORD WINAPI HookThreadProc(LPVOID aUnused)
// The creator of this thread relies on the fact that this function always exits its thread
// when both hooks are deactivated.
{
	MSG msg;
	bool problem_activating_hooks;

	for (;;) // Infinite loop for pumping messages in this thread. This thread will exit via any use of "return" below.
	{
		if (GetMessage(&msg, NULL, 0, 0) == -1) // -1 is an error, 0 means WM_QUIT.
			continue; // Probably happens only when bad parameters are passed to GetMessage().

		switch (msg.message)
		{
		case WM_QUIT:
			// After this message, fall through to the next case below so that the hooks will be removed before
			// exiting this thread.
			msg.wParam = 0; // Indicate to AHK_CHANGE_HOOK_STATE that both hooks should be deactivated.
		// ********
		// NO BREAK IN ABOVE, FALL INTO NEXT CASE:
		// ********
		case AHK_CHANGE_HOOK_STATE: // No blank line between this in the above to indicate fall-through.
			// In this case, wParam contains the bitwise set of hooks that should be active.
			problem_activating_hooks = false;
			if (msg.wParam & HOOK_KEYBD) // Activate the keyboard hook (if it isn't already).
			{
				if (!g_KeybdHook) // The creator of this thread has already ensured that OS isn't Win9x.
				{
					// v1.0.39: Reset *before* hook is installed to avoid any chance that events can
					// flow into the hook prior to the reset:
					if (msg.lParam) // Sender of msg. is signaling that reset should be done.
						ResetHook(false, HOOK_KEYBD, true);
					if (   !(g_KeybdHook = SetWindowsHookEx(WH_KEYBOARD_LL, LowLevelKeybdProc, g_hInstance, 0))   )
						problem_activating_hooks = true;
				}
			}
			else // Caller specified that the keyboard hook is to be deactivated (if it isn't already).
				if (g_KeybdHook)
					if (UnhookWindowsHookEx(g_KeybdHook))
						g_KeybdHook = NULL;

			if (msg.wParam & HOOK_MOUSE) // Activate the mouse hook (if it isn't already).
			{
				if (!g_MouseHook) // The creator of this thread has already ensured that OS isn't Win9x.
				{
					if (msg.lParam) // Sender of msg. is signaling that reset should be done.
						ResetHook(false, HOOK_MOUSE, true);
					if (   !(g_MouseHook = SetWindowsHookEx(WH_MOUSE_LL, LowLevelMouseProc, g_hInstance, 0))   )
						problem_activating_hooks = true;
				}
			}
			else // Caller specified that the mouse hook is to be deactivated (if it isn't already).
				if (g_MouseHook)
					if (UnhookWindowsHookEx(g_MouseHook))
						g_MouseHook = NULL;

			// Upon failure, don't display MsgBox here because although MsgBox's own message pump would
			// service the hook that didn't fail (if it's active), it's best to avoid any blocking calls
			// here so that this event loop will continue to run.  For example, the script or OS might
			// ask this thread to terminate, which it couldn't do cleanly if it was in a blocking call.
			// Instead, send a reply back to the caller.
			// It's safe to post directly to thread because the creator of this thread should be
			// explicitly waiting for this message (so there's no chance that a MsgBox msg pump
			// will discard the message unless the caller has timed out, which seems impossible
			// in this case).
			if (msg.wParam) // The caller wants a reply only when it didn't ask us to terminate via deactiving both hooks.
				PostThreadMessage(g_MainThreadID, AHK_CHANGE_HOOK_STATE, problem_activating_hooks, 0);
			//else this is WM_QUIT or the caller wanted this thread to terminate.  Send no reply.

			// If caller passes true for msg.lParam, it wants a permanent change to hook state; so in that case, terminate this
			// thread whenever neither hook is no longe present.
			if (msg.lParam && !(g_KeybdHook || g_MouseHook)) // Both hooks are inactive (for whatever reason).
				return 0; // Thread is no longer needed. The "return" automatically calls ExitThread().
				// 1) Due to this thread's non-GUI nature, there doesn't seem to be any need to call
				// the somewhat mysterious PostQuitMessage() here.
				// 2) For thread safety and maintainability, it seems best to have the caller take
				// full responsibility for freeing the hook's memory.
			break;

		} // switch (msg.message)
	} // for(;;)
}



void ResetHook(bool aAllModifiersUp, HookType aWhichHook, bool aResetKVKandKSC)
// Caller should ensure that aWhichHook indicates at least one of the hooks (not none).
{
	// Reset items common to both hooks:
	pPrefixKey = NULL;

	if (aWhichHook & HOOK_MOUSE)
	{
		// Initialize some things, a very limited subset of what is initialized when the
		// keyboard hook is installed (see its comments).  This is might not everything
		// we should initialize, so further study is justified in the future:
#ifdef FUTURE_USE_MOUSE_BUTTONS_LOGICAL
		g_mouse_buttons_logical = 0;
#endif
		g_PhysicalKeyState[VK_LBUTTON] = 0;
		g_PhysicalKeyState[VK_RBUTTON] = 0;
		g_PhysicalKeyState[VK_MBUTTON] = 0;
		g_PhysicalKeyState[VK_XBUTTON1] = 0;
		g_PhysicalKeyState[VK_XBUTTON2] = 0;
		// These are not really valid, since they can't be in a physically down state, but it's
		// probably better to have a false value in them:
		g_PhysicalKeyState[VK_WHEEL_DOWN] = 0;
		g_PhysicalKeyState[VK_WHEEL_UP] = 0;
		// Lexikos: Support horizontal scrolling in Windows Vista and later.
		g_PhysicalKeyState[VK_WHEEL_LEFT] = 0;
		g_PhysicalKeyState[VK_WHEEL_RIGHT] = 0;

		if (aResetKVKandKSC)
		{
			ResetKeyTypeState(kvk[VK_LBUTTON]);
			ResetKeyTypeState(kvk[VK_RBUTTON]);
			ResetKeyTypeState(kvk[VK_MBUTTON]);
			ResetKeyTypeState(kvk[VK_XBUTTON1]);
			ResetKeyTypeState(kvk[VK_XBUTTON2]);
			ResetKeyTypeState(kvk[VK_WHEEL_DOWN]);
			ResetKeyTypeState(kvk[VK_WHEEL_UP]);
			// Lexikos: Support horizontal scrolling in Windows Vista and later.
			ResetKeyTypeState(kvk[VK_WHEEL_LEFT]);
			ResetKeyTypeState(kvk[VK_WHEEL_RIGHT]);
		}
	}

	if (aWhichHook & HOOK_KEYBD)
	{
		// Doesn't seem necessary to ever init g_KeyHistory or g_KeyHistoryNext here, since they were
		// zero-filled on startup.  But we do want to reset the below whenever the hook is being
		// installed after a (probably long) period during which it wasn't installed.  This is
		// because we don't know the current physical state of the keyboard and such:

		g_modifiersLR_physical = 0;  // Best to make this zero, otherwise keys might get stuck down after a Send.
		g_modifiersLR_logical = g_modifiersLR_logical_non_ignored = (aAllModifiersUp ? 0 : GetModifierLRState(true));

		ZeroMemory(g_PhysicalKeyState, sizeof(g_PhysicalKeyState));

		sDisguiseNextLWinUp = false;
		sDisguiseNextRWinUp = false;
		sDisguiseNextLAltUp = false;
		sDisguiseNextRAltUp = false;
		sAltTabMenuIsVisible = (FindWindow("#32771", NULL) != NULL); // I've seen indications that MS wants this to work on all operating systems.
		sVKtoIgnoreNextTimeDown = 0;

		ZeroMemory(sPadState, sizeof(sPadState));

		*g_HSBuf = '\0';
		g_HSBufLength = 0;
		g_HShwnd = GetForegroundWindow(); // Not needed by some callers, but shouldn't hurt even then.

		// Variables for the Shift+Numpad workaround:
		sNextPhysShiftDownIsNotPhys = false;
		sPriorVK = 0;
		sPriorSC = 0;
		sPriorEventWasKeyUp = false;
		sPriorEventWasPhysical = false;
		sPriorEventTickCount = 0;
		sPriorModifiersLR_physical = 0;
		sPriorShiftState = 0;  // i.e. default to "key is up".
		sPriorLShiftState = 0; //

		if (aResetKVKandKSC)
		{
			int i;
			for (i = 0; i < VK_ARRAY_COUNT; ++i)
				if (!IsMouseVK(i))  // Don't do mouse VKs since those must be handled by the mouse section.
					ResetKeyTypeState(kvk[i]);
			for (i = 0; i < SC_ARRAY_COUNT; ++i)
				ResetKeyTypeState(ksc[i]);
		}
	}
}



HookType GetActiveHooks()
{
	HookType hooks_currently_active = 0;
	if (g_KeybdHook)
		hooks_currently_active |= HOOK_KEYBD;
	if (g_MouseHook)
		hooks_currently_active |= HOOK_MOUSE;
	return hooks_currently_active;
}



void FreeHookMem()
// For maintainability, only the main thread should ever call this function.
// This is because new/delete/malloc/free themselves might not be thread-safe
// when the single-threaded CRT libraries are in effect (not using multi-threaded
// libraries due to a 3.5 KB increase in compressed code size).
{
	if (kvk)
	{
		delete [] kvk;
		kvk = NULL;
	}
	if (ksc)
	{
		delete [] ksc;
		ksc = NULL;
	}
	if (kvkm)
	{
		delete [] kvkm;
		kvkm = NULL;
	}
	if (kscm)
	{
		delete [] kscm;
		kscm = NULL;
	}
	if (hotkey_up)
	{
		delete [] hotkey_up;
		hotkey_up = NULL;
	}
}



void ResetKeyTypeState(key_type &key)
{
	key.is_down = false;
	key.it_put_alt_down = false;
	key.it_put_shift_down = false;
	key.down_performed_action = false;
	key.was_just_used = 0;
	key.hotkey_to_fire_upon_release = HOTKEY_ID_INVALID;
	// ABOVE line was added in v1.0.48.03 to fix various ways in which the hook didn't receive the key-down
	// hotkey that goes with this key-up, resulting in hotkey_to_fire_upon_release being left at its initial
	// value of zero (which is a valid hotkey ID).  Examples include:
	// The hotkey command being used to create a key-up hotkey while that key is being held down.
	// The script being reloaded or (re)started while the key is being held down.
}



void GetHookStatus(char *aBuf, int aBufSize)
// aBufSize is an int so that any negative values passed in from caller are not lost.
{
	char LRhText[128], LRpText[128];
	snprintfcat(aBuf, aBufSize,
		"Modifiers (Hook's Logical) = %s\r\n"
		"Modifiers (Hook's Physical) = %s\r\n" // Font isn't fixed-width, so don't bother trying to line them up.
		"Prefix key is down: %s\r\n"
		, ModifiersLRToText(g_modifiersLR_logical, LRhText)
		, ModifiersLRToText(g_modifiersLR_physical, LRpText)
		, pPrefixKey ? "yes" : "no");

	if (!g_KeybdHook)
		snprintfcat(aBuf, aBufSize, "\r\n"
			"NOTE: Only the script's own keyboard events are shown\r\n"
			"(not the user's), because the keyboard hook isn't installed.\r\n");

	// Add the below even if key history is already disabled so that the column headings can be seen.
	snprintfcat(aBuf, aBufSize, 
		"\r\nNOTE: To disable the key history shown below, add the line \"#KeyHistory 0\" "
		"anywhere in the script.  The same method can be used to change the size "
		"of the history buffer.  For example: #KeyHistory 100  (Default is 40, Max is 500)"
		"\r\n\r\nThe oldest are listed first.  VK=Virtual Key, SC=Scan Code, Elapsed=Seconds since the previous event"
		".  Types: h=Hook Hotkey, s=Suppressed (blocked), i=Ignored because it was generated by an AHK script"
		", a=Artificial, #=Disabled via #IfWinActive/Exist.\r\n\r\n"
		"VK  SC\tType\tUp/Dn\tElapsed\tKey\t\tWindow\r\n"
		"-------------------------------------------------------------------------------------------------------------");

	if (g_KeyHistory)
	{
		// Start at the oldest key, which is KeyHistoryNext:
		char KeyName[128];
		int item, i;
		char *title_curr = "", *title_prev = "";
		for (item = g_KeyHistoryNext, i = 0; i < g_MaxHistoryKeys; ++i, ++item)
		{
			if (item >= g_MaxHistoryKeys)
				item = 0;
			title_prev = title_curr;
			title_curr = g_KeyHistory[item].target_window;
			if (g_KeyHistory[item].vk || g_KeyHistory[item].sc)
				snprintfcat(aBuf, aBufSize, "\r\n%02X  %03X\t%c\t%c\t%0.2f\t%-15s\t%s"
					, g_KeyHistory[item].vk, g_KeyHistory[item].sc
					// It can't be both ignored and suppressed, so display only one:
					, g_KeyHistory[item].event_type
					, g_KeyHistory[item].key_up ? 'u' : 'd'
					, g_KeyHistory[item].elapsed_time
					, GetKeyName(g_KeyHistory[item].vk, g_KeyHistory[item].sc, KeyName, sizeof(KeyName))
					, strcmp(title_curr, title_prev) ? title_curr : "" // Display title only when it changes.
					);
		}
	}
}
