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

#ifndef keyboard_h
#define keyboard_h

#include "defines.h"
#include "abi.h"
EXTERN_G;

// The max number of keystrokes to Send prior to taking a break to pump messages:
#define MAX_LUMP_KEYS 50

// Maybe define more of these later, perhaps with ifndef (since they should be in the normal header, and probably
// will be eventually):
// ALREADY DEFINED: #define VK_HELP 0x2F
// In case a compiler with a non-updated header file is used:
#ifndef VK_BROWSER_BACK
	#define VK_BROWSER_BACK        0xA6
	#define VK_BROWSER_FORWARD     0xA7
	#define VK_BROWSER_REFRESH     0xA8
	#define VK_BROWSER_STOP        0xA9
	#define VK_BROWSER_SEARCH      0xAA
	#define VK_BROWSER_FAVORITES   0xAB
	#define VK_BROWSER_HOME        0xAC
	#define VK_VOLUME_MUTE         0xAD
	#define VK_VOLUME_DOWN         0xAE
	#define VK_VOLUME_UP           0xAF
	#define VK_MEDIA_NEXT_TRACK    0xB0
	#define VK_MEDIA_PREV_TRACK    0xB1
	#define VK_MEDIA_STOP          0xB2
	#define VK_MEDIA_PLAY_PAUSE    0xB3
	#define VK_LAUNCH_MAIL         0xB4
	#define VK_LAUNCH_MEDIA_SELECT 0xB5
	#define VK_LAUNCH_APP1         0xB6
	#define VK_LAUNCH_APP2         0xB7
#endif

// Create some "fake" virtual keys to simplify sections of the code.
// According to winuser.h, the following ranges (among others)
// are considered "unassigned" rather than "reserved", so should be
// fairly safe to use for the foreseeable future.  0xFF should probably
// be avoided since it's sometimes used as a failure indictor by API
// calls.  And 0x00 should definitely be avoided because it is used
// to indicate failure by many functions that deal with virtual keys.
// 0x88 - 0x8F : unassigned
// 0x97 - 0x9F : unassigned (this range seems less likely to be used)
#define VK_NEW_MOUSE_FIRST 0x9A
#define VK_WHEEL_LEFT      0x9C // v1.0.48: Lexikos: Fake virtual keys for support for horizontal scrolling in
#define VK_WHEEL_RIGHT     0x9D // Windows Vista and later.
#define VK_WHEEL_DOWN      0x9E
#define VK_WHEEL_UP        0x9F
#define IS_WHEEL_VK(aVK) ((aVK) >= VK_WHEEL_LEFT && (aVK) <= VK_WHEEL_UP)
#define VK_NEW_MOUSE_LAST  0x9F

// These are the only keys for which another key with the same VK exists.  Therefore, use scan code for these.
// If use VK for some of these (due to them being more likely to be used as hotkeys, thus minimizing the
// use of the keyboard hook), be sure to use SC for its counterpart.
// Always use the compressed version of scancode, i.e. 0x01 for the high-order byte rather than vs. 0xE0.
#define SC_NUMPADENTER 0x11C
#define SC_INSERT 0x152
#define SC_DELETE 0x153
#define SC_HOME 0x147
#define SC_END 0x14F
#define SC_UP 0x148
#define SC_DOWN 0x150
#define SC_LEFT 0x14B
#define SC_RIGHT 0x14D
#define SC_PGUP 0x149
#define SC_PGDN 0x151

// These are the same scan codes as their counterpart except the extended flag is 0 rather than
// 1 (0xE0 uncompressed):
#define SC_ENTER 0x1C
// In addition, the below dual-state numpad keys share the same scan code (but different vk's)
// regardless of the state of numlock:
#define SC_NUMPADDEL 0x53
#define SC_NUMPADINS 0x52
#define SC_NUMPADEND 0x4F
#define SC_NUMPADHOME 0x47
#define SC_NUMPADCLEAR 0x4C
#define SC_NUMPADUP 0x48
#define SC_NUMPADDOWN 0x50
#define SC_NUMPADLEFT 0x4B
#define SC_NUMPADRIGHT 0x4D
#define SC_NUMPADPGUP 0x49
#define SC_NUMPADPGDN 0x51

#define SC_NUMPADDOT SC_NUMPADDEL
#define SC_NUMPAD0 SC_NUMPADINS
#define SC_NUMPAD1 SC_NUMPADEND
#define SC_NUMPAD2 SC_NUMPADDOWN
#define SC_NUMPAD3 SC_NUMPADPGDN
#define SC_NUMPAD4 SC_NUMPADLEFT
#define SC_NUMPAD5 SC_NUMPADCLEAR
#define SC_NUMPAD6 SC_NUMPADRIGHT
#define SC_NUMPAD7 SC_NUMPADHOME
#define SC_NUMPAD8 SC_NUMPADUP
#define SC_NUMPAD9 SC_NUMPADPGUP

#define SC_NUMLOCK 0x145
#define SC_NUMPADDIV 0x135
#define SC_NUMPADMULT 0x037
#define SC_NUMPADSUB 0x04A
#define SC_NUMPADADD 0x04E
#define SC_PAUSE 0x045

#define SC_LCONTROL 0x01D
#define SC_RCONTROL 0x11D
#define SC_LSHIFT 0x02A
#define SC_RSHIFT 0x136 // Must be extended, at least on WinXP, or there will be problems, e.g. SetModifierLRState().
#define SC_LALT 0x038
#define SC_RALT 0x138
#define SC_LWIN 0x15B
#define SC_RWIN 0x15C

#define SC_APPSKEY 0x15D
#define SC_PRINTSCREEN 0x137

// The system injects events with these scan codes:
//  - For Shift-up prior to a Numpad keydown or keyup if Numlock is on and Shift is down;
//    e.g. to translate Shift+Numpad1 to unshifted-NumpadEnd.
//  - For Shift-down prior to a non-Numpad keydown if a Numpad key is still held down
//    after the above; e.g. for Shift+Numpad1+Esc.
//  - For LCtrl generated by AltGr.
// Note that the system uses the normal scan codes (0x2A or 0x36) for Shift-down following
// the Numpad keyup if no other keys were pushed.  Our hook filters out the second byte to
// simplify the code, so these values can only be found in KBDLLHOOKSTRUCT::scanCode.
// Find "fake shift-key events" for older and more detailed comments.
// Note that 0x0200 corresponds to SCANCODE_SIMULATED in kbd.h (DDK).
#define SC_FAKE_FLAG 0x200
#define SC_FAKE_LSHIFT 0x22A
#define SC_FAKE_RSHIFT 0x236 // This is the actual scancode received by the hook, excluding the 0x100 we add for "extended" keys.
// Testing with the KbdEdit Demo preview mode indicates that AltGr will send this SC
// even if the VK assigned to 0x1D is changed.  It is a combination of SCANCODE_CTRL
// and SCANCODE_SIMULATED, which are defined in kbd.h (Windows DDK).
#define SC_FAKE_LCTRL 0x21D

// UPDATE for v1.0.39: Changed sc_type to USHORT vs. UINT to save memory in structs such as sc_hotkey.
// This saves 60K of memory in one place, and possibly there are other large savings too.
// The following older comment dates back to 2003/inception and I don't remember its exact intent,
// but there is no current storage of mouse message constants in scan code variables:
// OLD: Although only need 9 bits for compressed and 16 for uncompressed scan code, use a full 32 bits 
// so that mouse messages (WPARAM) can be stored as scan codes.  Formerly USHORT (which is always 16-bit).
typedef USHORT sc_type; // Scan code.
typedef UCHAR vk_type;  // Virtual key.
typedef UINT mod_type;  // Standard Windows modifier type for storing MOD_CONTROL, MOD_WIN, MOD_ALT, MOD_SHIFT.

// The maximum number of virtual keys and scan codes that can ever exist.
// As of WinXP, these are absolute limits, except for scan codes for which there might conceivably
// be more if any non-standard keyboard or keyboard drivers generate scan codes that don't start
// with either 0x00 or 0xE0.  UPDATE: Decided to allow all possible scancodes, rather than just 512,
// since a lookup array for the 16-bit scan code value will only consume 64K of RAM if the element
// size is one char.  UPDATE: decided to go back to 512 scan codes, because WinAPI's KeyboardProc()
// itself can only handle that many (a 9-bit value).  254 is the largest valid vk, according to the
// WinAPI docs (I think 255 is value that is sometimes returned to indicate an invalid vk).  But
// just in case something ever tries to access such arrays using the largest 8-bit value (255), add
// 1 to make it 0xFF, thus ensuring array indexes will always be in-bounds if they are 8-bit values.
#define VK_MAX 0xFF
#define SC_MAX 0x1FF

typedef UCHAR modLR_type; // Only the left-right win/alt/ctrl/shift rather than their generic counterparts.
#define MODLR_MAX 0xFF
#define MODLR_COUNT 8
#define MOD_LCONTROL 0x01
#define MOD_RCONTROL 0x02
#define MOD_LALT 0x04
#define MOD_RALT 0x08
#define MOD_LSHIFT 0x10
#define MOD_RSHIFT 0x20
#define MOD_LWIN 0x40
#define MOD_RWIN 0x80
#define MODLR_STRING _T("<^>^<!>!<+>+<#>#")


struct CachedLayoutType
{
	HKL hkl;
	ResultType has_altgr;
};

struct key_to_vk_type // Map key names to virtual keys.
{
	LPTSTR key_name;
	vk_type vk;
};

struct key_to_sc_type // Map key names to scan codes.
{
	LPTSTR key_name;
	sc_type sc;
};

enum KeyStateTypes {KEYSTATE_LOGICAL, KEYSTATE_PHYSICAL, KEYSTATE_TOGGLE}; // For use with GetKeyJoyState(), etc.
enum KeyEventTypes {KEYDOWN, KEYUP, KEYDOWNANDUP};

void SendKeys(LPCTSTR aKeys, SendRawModes aSendRaw, SendModes aSendModeOrig, HWND aTargetWindow = NULL);
void SendKey(vk_type aVK, sc_type aSC, modLR_type aModifiersLR, modLR_type aModifiersLRPersistent
	, int aRepeatCount, KeyEventTypes aEventType, modLR_type aKeyAsModifiersLR, HWND aTargetWindow
	, int aX = COORD_UNSPECIFIED, int aY = COORD_UNSPECIFIED, bool aMoveOffset = false);
void SendKeySpecial(TCHAR aChar, int aRepeatCount, modLR_type aModifiersLR);
void SendASC(LPCTSTR aAscii);

struct PlaybackEvent
{
	UINT message;
	union
	{
		struct
		{
			sc_type sc; // Placed above vk for possibly better member stacking/alignment.
			vk_type vk;
		};
		struct
		{
			// Screen coordinates, which can be negative.  SHORT vs. INT is used because the likelihood
			// have having a virtual display surface wider or taller than 32,767 seems too remote to
			// justify increasing the struct size, which would impact the stack space and dynamic memory
			// used by every script every time it uses the playback method to send keystrokes or clicks.
			// Note: WM_LBUTTONDOWN uses WORDs vs. SHORTs, but they're not really comparable because
			// journal playback/record both use screen coordinates but WM_LBUTTONDOWN et. al. use client
			// coordinates.
			SHORT x;
			SHORT y;
		};
		DWORD time_to_wait; // This member is present only when message==0; otherwise, a struct is present.
	};
};
LRESULT CALLBACK PlaybackProc(int aCode, WPARAM wParam, LPARAM lParam);


// Below uses a pseudo-random value.  It's best that this be constant so that if multiple instances
// of the app are running, they will all ignore each other's keyboard & mouse events.  Also, a value
// close to UINT_MAX might be a little better since it's might be less likely to be used as a pointer
// value by any apps that send keybd events whose ExtraInfo is really a pointer value.
#define KEY_IGNORE 0xFFC3D44F
#define KEY_PHYS_IGNORE (KEY_IGNORE - 1)  // Same as above but marked as physical for other instances of the hook.
#define KEY_IGNORE_ALL_EXCEPT_MODIFIER (KEY_IGNORE - 2)  // Non-physical and ignored only if it's not a modifier.
// Same as KEY_IGNORE_ALL_EXCEPT_MODIFIER, but only ignored by Hotkeys & Hotstrings at InputLevel LEVEL and below.
// The levels are set up to use negative offsets from KEY_IGNORE_ALL_EXCEPT_MODIFIER so that we can leave
// the values above unchanged and have KEY_IGNORE_LEVEL(0) == KEY_IGNORE_ALL_EXCEPT_MODIFIER.
//
// In general, KEY_IGNORE_LEVEL(g->SendLevel) should be used for any input that's meant to be treated as "real",
// as opposed to input generated for side effects (e.g., masking modifier keys to prevent default OS responses).
// A lot of the calls that generate input fall into the latter category, so KEY_IGNORE_ALL_EXCEPT_MODIFIER
// (aka KEY_IGNORE_LEVEL(0)) still gets used quite often.
//
// Note that there are no "level" equivalents for KEY_IGNORE or KEY_PHYS_IGNORE (only KEY_IGNORE_ALL_EXCEPT_MODIFIER).
// For the KEY_IGNORE_LEVEL use cases, there isn't a need to ignore modifiers or differentiate between physical
// and non-physical, and leaving them out keeps the code much simpler.
#define KEY_IGNORE_LEVEL(LEVEL) (KEY_IGNORE_ALL_EXCEPT_MODIFIER - LEVEL)
#define KEY_IGNORE_MIN KEY_IGNORE_LEVEL(SendLevelMax)
#define KEY_IGNORE_MAX KEY_IGNORE // There are two extra values above KEY_IGNORE_LEVEL(0)
// This is used to generate an Alt key-up event for the purpose of changing system state, but having the hook
// block it from the active window to avoid unwanted side-effects:
#define KEY_BLOCK_THIS (KEY_IGNORE + 1)


// The default in the below is KEY_IGNORE_ALL_EXCEPT_MODIFIER, which causes standard calls to
// KeyEvent() to update g_modifiersLR_logical_non_ignored the same way it updates g_modifiersLR_logical.
// This is done because only the Send command has a realistic chance of interfering with (or being
// interfered with by) hook hotkeys (namely the modifiers used to decide whether to trigger them).
// There are two types of problems:
// 1) Hotkeys not firing due to Send having temporarily released one of that hotkey's modifiers that
//    the user is still holding down.  This causes the hotkey's suffix to flow through to the system,
//    which is usually undesirable.  This happens when the user is holding down a hotkey to auto-repeat
//    it, and perhaps other times.
// 2) The wrong hotkey firing because Send has temporarily put a modifier into effect and (once again)
//    the user is holding down the hotkey to auto-repeat it.  If the Send's temp-down modifier happens
//    to make the hotkey suffix match a different set of modifiers, the wrong hotkey would fire.
void KeyEvent(KeyEventTypes aEventType, vk_type aVK, sc_type aSC = 0, HWND aTargetWindow = NULL
	, bool aDoKeyDelay = false, DWORD aExtraInfo = KEY_IGNORE_ALL_EXCEPT_MODIFIER);
void KeyEventMenuMask(KeyEventTypes aEventType, DWORD aExtraInfo = KEY_IGNORE_ALL_EXCEPT_MODIFIER);

ResultType PerformClick(LPTSTR aOptions);
void ParseClickOptions(LPTSTR aOptions, int &aX, int &aY, vk_type &aVK, KeyEventTypes &aEventType
	, int &aRepeatCount, bool &aMoveOffset);
FResult PerformMouse(ActionTypeType aActionType, optl<StrArg> aButton
	, optl<int> aX1, optl<int> aY1, optl<int> aX2, optl<int> aY2
	, optl<int> aSpeed, optl<StrArg> aOffsetMode, optl<int> aRepeatCount, optl<StrArg> aDownUp);
void PerformMouseCommon(ActionTypeType aActionType, vk_type aVK, int aX1, int aY1, int aX2, int aY2
	, int aRepeatCount, KeyEventTypes aEventType, int aSpeed, bool aMoveOffset);

void MouseClickDrag(vk_type aVK // Which button.
	, int aX1, int aY1, int aX2, int aY2, int aSpeed, bool aMoveOffset);
void MouseClick(vk_type aVK // Which button.
	, int aX, int aY, int aRepeatCount, int aSpeed, KeyEventTypes aEventType, bool aMoveOffset = false);
void MouseMove(int &aX, int &aY, DWORD &aEventFlags, int aSpeed, bool aMoveOffset);
void MouseEvent(DWORD aEventFlags, DWORD aData, DWORD aX = COORD_UNSPECIFIED, DWORD aY = COORD_UNSPECIFIED);

#define MSG_OFFSET_MOUSE_MOVE 0x80000000  // Bitwise flag, should be near/at high-order bit to avoid overlap messages.
void PutKeybdEventIntoArray(modLR_type aKeyAsModifiersLR, vk_type aVK, sc_type aSC, DWORD aEventFlags, DWORD aExtraInfo);
void PutMouseEventIntoArray(DWORD aEventFlags, DWORD aData, DWORD aX, DWORD aY);
ResultType ExpandEventArray();
void InitEventArray(void *aMem, UINT aMaxEvents, modLR_type aModifiersLR);
void SendEventArray(int &aFinalKeyDelay, modLR_type aModsDuringSend);
void CleanupEventArray(int aFinalKeyDelay);

extern SendModes sSendMode;
void DoKeyDelay(int aDelay = (sSendMode == SM_PLAY) ? g->KeyDelayPlay : g->KeyDelay);
void DoMouseDelay();
void UpdateKeyEventHistory(bool aKeyUp, vk_type aVK, sc_type aSC);
void SetKeyHistoryMax(int aMax);
#define KEYEVENT_PHYS(event_type, vk, sc) KeyEvent(event_type, vk, sc, NULL, false, KEY_PHYS_IGNORE)

ToggleValueType ToggleKeyState(vk_type aVK, ToggleValueType aToggleValue);
FResult SetToggleState(vk_type aVK, ToggleValueType &ForceLock, optl<StrArg> aToggleText);

#define STD_MODS_TO_DISGUISE (MOD_LALT|MOD_RALT|MOD_LWIN|MOD_RWIN)
void SetModifierLRState(modLR_type aModifiersLRnew, modLR_type aModifiersLRnow, HWND aTargetWindow
	, bool aDisguiseDownWinAlt, bool aDisguiseUpWinAlt, DWORD aExtraInfo = KEY_IGNORE_ALL_EXCEPT_MODIFIER);
modLR_type GetModifierLRState(bool aExplicitlyGet = false);

#define IsKeyDown(vk) (GetKeyState(vk) & 0x8000)
#define IsKeyDownAsync(vk) (GetAsyncKeyState(vk) & 0x8000)
#define IsKeyToggledOn(vk) (GetKeyState(vk) & 0x01)

void AdjustKeyState(BYTE aKeyState[], modLR_type aModifiersLR);
modLR_type KeyToModifiersLR(vk_type aVK, sc_type aSC = 0, bool *pIsNeutral = NULL);
modLR_type ConvertModifiers(mod_type aModifiers);
mod_type ConvertModifiersLR(modLR_type aModifiersLR);
LPTSTR ModifiersLRToText(modLR_type aModifiersLR, LPTSTR aBuf);

DWORD GetFocusedCtrlThread(HWND *apControl = NULL, HWND aWindow = GetForegroundWindow());
HKL GetFocusedKeybdLayout(HWND aWindow = GetForegroundWindow());

bool ActiveWindowLayoutHasAltGr();
ResultType LayoutHasAltGr(HKL aLayout);

//---------------------------------------------------------------------

LPTSTR SCtoKeyName(sc_type aSC, LPTSTR aBuf, int aBufSize, bool aUseFallback = true);
LPTSTR VKtoKeyName(vk_type aVK, LPTSTR aBuf, int aBufSize, bool aUseFallback = true);
TCHAR VKtoChar(vk_type aVK, HKL aKeybdLayout = NULL);
sc_type TextToSC(LPCTSTR aText, bool *aSpecifiedByNumber = NULL);
vk_type TextToVK(LPCTSTR aText, modLR_type *pModifiersLR = NULL, bool aExcludeThoseHandledByScanCode = false
	, bool aAllowExplicitVK = true, HKL aKeybdLayout = GetKeyboardLayout(0));
vk_type CharToVKAndModifiers(TCHAR aChar, modLR_type *pModifiersLR, HKL aKeybdLayout, bool aEnableAZFallback = true);
bool TextToVKandSC(LPCTSTR aText, vk_type &aVK, sc_type &aSC, modLR_type *pModifiersLR = NULL, HKL aKeybdLayout = GetKeyboardLayout(0));
vk_type TextToSpecial(LPTSTR aText, size_t aTextLength, KeyEventTypes &aEventTypem, modLR_type &aModifiersLR
	, bool aUpdatePersistent);

LPTSTR GetKeyName(vk_type aVK, sc_type aSC, LPTSTR aBuf, int aBufSize, LPTSTR aDefault = _T("not found"));
sc_type vk_to_sc(vk_type aVK, bool aReturnSecondary = false);
vk_type sc_to_vk(sc_type aSC);

inline bool IsMouseVK(vk_type aVK)
{
	return aVK >= VK_LBUTTON && aVK <= VK_XBUTTON2 && aVK != VK_CANCEL
		|| aVK >= VK_NEW_MOUSE_FIRST && aVK <= VK_NEW_MOUSE_LAST;
}

void OurBlockInput(bool aEnable);

#endif
