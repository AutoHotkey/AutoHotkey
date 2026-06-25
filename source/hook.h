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

#ifndef hook_h
#define hook_h

#include "hotkey.h" // Use here and also by hook.cpp for ChangeHookState(), which reads from static Hotkey class vars.

// WM_USER is the lowest number that can be a user-defined message.  Anything above that is also valid.
// NOTE: Any msg >= WM_HOTKEY will be kept buffered (unreplied-to) whenever the script is uninterruptible.
// Also, it has been announced in OnMessage() that message numbers between WM_USER and 0x1000 are earmarked
// for possible future use by the program, so don't use a message above 0x1000 without good reason.
enum UserMessages {AHK_HOOK_HOTKEY = WM_USER, AHK_HOTSTRING, AHK_USER_MENU, AHK_DIALOG, AHK_NOTIFYICON
	, AHK_UNUSED_MSG, AHK_EXIT_BY_RELOAD, AHK_EXIT_BY_SINGLEINSTANCE, AHK_CHECK_DEBUGGER
	// Allow some room here in between for more "exit" type msgs to be added in the future (see below comment).
	, AHK_GUI_ACTION = WM_USER+20 // Avoid WM_USER+100/101 and vicinity.  See below comment.
	// v1.0.43.05: On second thought, it seems better to stay close to WM_USER because the OnMessage page
	// documents, "it is best to choose a number greater than 4096 (0x1000) to the extent you have a choice.
	// This reduces the chance of interfering with messages used internally by current and future versions..."
	// v1.0.43.03: Changed above msg number because Micha reports that previous number (WM_USER+100) conflicts
	// with msgs sent by HTML control (AHK_CLIPBOARD_CHANGE) and possibly others (I think WM_USER+100 may be the
	// start of a range used by other common controls too).  So trying a higher number that's (hopefully) very
	// unlikely to be used by OS features.
	, AHK_CLIPBOARD_CHANGE, AHK_HOOK_TEST_MSG, AHK_CHANGE_HOOK_STATE, AHK_GETWINDOWTEXT
	, AHK_HOT_IF_EVAL	// HotCriterionAllowsFiring uses this to ensure expressions are evaluated only on the main thread.
	, AHK_HOOK_SYNC // For WaitHookIdle().
	, AHK_INPUT_END, AHK_INPUT_KEYDOWN, AHK_INPUT_CHAR, AHK_INPUT_KEYUP
	, AHK_HOOK_SET_KEYHISTORY
};
// NOTE: TRY NEVER TO CHANGE the specific numbers of the above messages, since some users might be
// using the Post/SendMessage commands to automate AutoHotkey itself.  Here is the original order
// that should be maintained:
// AHK_HOOK_HOTKEY = WM_USER, AHK_HOTSTRING, AHK_USER_MENU, AHK_DIALOG, AHK_NOTIFYICON, AHK_RETURN_PID

// The following macro was previously used to provide a means for bypassing the message filtering our
// message loop performs when a thread is uninterruptible.  There are two problems with relying on this:
//  1) If the script unblocks the message with ChangeWindowMessageFilter, such as to allow a new script
//     instance running at a lower IL to replace one running at a higher IL (as admin/with UI access),
//     that would defeat a large part of the security provided by the UIPI message filter, since the
//     translation performed by this macro allows arbitrary messages.
//  2) Every incoming message must be checked and translated, which is an unnecessary inefficiency
//     given that there are viable alternatives in each case.
//#define TRANSLATE_AHK_MSG(msg, wparam) \
//	if (msg == WM_COMMNOTIFY)\
//	{\
//		msg = (UINT)wparam;\
//		wparam = 0;\
//	} // In the above, wparam is made zero to help catch bugs.

// And these macros are kept here so that all this trickery is centrally located and thus more maintainable:

// v2.1: WM_CLOSE is used rather than the more generic mapping of message numbers to WM_COMMNOTIFY's
// wParam so that scripts can unblock Reload/#SingleInstance messages from processes running at a
// lower IL via ChangeWindowMessageFilter without introducing a larger security risk.
#define ASK_INSTANCE_TO_CLOSE(hwnd, reason) PostMessage(hwnd, WM_CLOSE, reason, 0);

// POST_AHK_USER_MENU: A gui_hwnd value != 0 is passed with the message if it came from a GUI's menu bar.
// This is done because it's good way to pass the info, but also so that its value will be in sync with the
// timestamp of the message (in case the message is stuck in the queue for a long time).  No pointer is
// passed in this case since they might become invalid between the time the msg is posted vs. processed.
#define POST_AHK_USER_MENU(hwnd, menu, gui_hwnd) PostMessage(hwnd, AHK_USER_MENU, gui_hwnd, menu);
#define POST_AHK_GUI_ACTION(hwnd, control_index, gui_event, event_info) PostMessage(hwnd, AHK_GUI_ACTION \
	, (WPARAM)(((control_index) << 16) | (gui_event)), (LPARAM)(event_info)); // Caller must ensure that gui_event is less than 0xFFFF.
// POST_AHK_DIALOG:
// Post a special msg that will attempt to force it to the foreground after it has been displayed,
// since the dialog often will flash in the task bar instead of becoming foreground.
// It's enough just to queue up a single message that dialog's message pump will forward to our
// main window proc once the dialog window has been displayed.  This avoids the overhead of creating
// and destroying the timer (although the timer may be needed anyway if any timed subroutines are
// enabled).  My only concern about this is that on some OS's, or on slower CPUs, the message may be
// received too soon (before the dialog window actually exists) resulting in our window proc not
// being able to ensure that it's the foreground window.  That seems unlikely, however, since
// MessageBox() and the other dialog invocating API calls (for FileSelect/DirSelect) likely
// ensures its window really exists before dispatching messages.
#define POST_AHK_DIALOG(timeout) PostMessage(g_hWnd, AHK_DIALOG, 0, (WPARAM)timeout);

// Some reasoning behind the below data structures: Could build a new array for [sc][sc] and [vk][vk]
// (since only two keys are allowed in a ModifierVK/SC combination, only 2 dimensions are needed).
// But this would be a 512x512 array of shorts just for the SC part, which is 512K.  Instead, what we
// do is check whenever a key comes in: if it's a suffix and if a non-standard modifier key of any kind
// is currently down: consider action.  Most of the time, an action be found because the user isn't
// likely to be holding down a ModifierVK/SC, while pressing another key, unless it's modifying that key.
// Nor is he likely to have more than one ModifierVK/SC held down at a time.  It's still somewhat
// inefficient because have to look up the right prefix in a loop.  But most suffixes probably won't
// have more than one ModifierVK/SC anyway, so the lookup will usually find a match on the first
// iteration.
struct vk_hotkey
{
	vk_type vk;
	HotkeyIDType id_with_flags;
};
struct sc_hotkey
{
	sc_type sc;
	HotkeyIDType id_with_flags;
};

// Style reminder: Any POD structs (those without any methods) don't use the "m" prefix
// for member variables because there's no need: the variables are always prefixed by
// the struct that owns them, so there's never any ambiguity:
struct key_type
{
	ToggleValueType *pForceToggle;  // Pointer to a global variable for toggleable keys only.  NULL for others.
	// Keep sub-32-bit members contiguous to save memory without having to sacrifice performance of
	// 32-bit alignment:
	HotkeyIDType hotkey_to_fire_upon_release; // A up-event hotkey queued by a prior down-event.
	HotkeyIDType first_hotkey; // The first hotkey using this key as a suffix.
	modLR_type as_modifiersLR; // If this key is a modifier, this will have the corresponding bit(s) for that key.
	#define PREFIX_ACTUAL 1 // Values for used_as_prefix below, for places that need to distinguish between type of prefix.
	#define PREFIX_FORCED 2 // v1.0.44: Added so that a neutral hotkey like Control can be forced to fire on key-up even though it isn't actually a prefix key.
	UCHAR used_as_prefix; // Whether a given virtual key or scan code is even used by a hotkey.
	bool used_as_suffix;  //
	bool used_as_key_up;  // Whether this suffix also has an enabled key-up hotkey.
	UCHAR no_suppress; // Contains bitwise flags such as NO_SUPPRESS_PREFIX.
	bool is_down; // this key is currently down.
	bool it_put_alt_down;  // this key resulted in ALT being pushed down (due to alt-tab).
	bool it_put_shift_down;  // this key resulted in SHIFT being pushed down (due to shift-alt-tab).
	bool down_performed_action; // the last key-down resulted in an action (modifiers matched those of a valid hotkey)
	UCHAR down_was_suppressed; // Whether the down-event for a key was suppressed (thus its up-event should be too).
	// The values for "was_just_used" (zero is the initialized default, meaning it wasn't just used):
	char was_just_used; // a non-modifier key of any kind was pressed while this prefix key was down.
	// And these are the values for the above (besides 0):
	#define AS_PREFIX 1
	#define AS_PREFIX_FOR_HOTKEY 2
	#define AS_PASSTHROUGH_PREFIX -1 // v1.1.34.02: Indicates the suffix key-up hotkey of this prefix key should fire even if a combo was activated.
	bool sc_takes_precedence; // used only by the scan code array: this scan code should take precedence over vk.
}; // Keep the macro below in sync with the above.

#define RESET_KEYTYPE_ATTRIB(item) \
{\
	item.first_hotkey = HOTKEY_ID_INVALID;\
	item.used_as_prefix = 0;\
	item.used_as_suffix = false;\
	item.used_as_key_up = false;\
	item.sc_takes_precedence = false;\
}

// Since index zero is a placeholder for the invalid virtual key or scan code, add one to each MAX value
// to compute the number of elements actually needed to accommodate 0 up to and including VK_MAX or SC_MAX:
#define VK_ARRAY_COUNT (VK_MAX + 1)
#define SC_ARRAY_COUNT (SC_MAX + 1)

#define INPUT_BUFFER_SIZE 16384 // Default buffer size for Input.  Used to be the absolute max.
#define INPUTHOOK_BUFFER_SIZE 1024 // Default buffer size for InputHook.

enum InputStatusType {INPUT_OFF, INPUT_IN_PROGRESS, INPUT_TIMED_OUT, INPUT_TERMINATED_BY_MATCH
	, INPUT_TERMINATED_BY_ENDKEY, INPUT_LIMIT_REACHED};

// Bitwise flags for the Key arrays:
#define END_KEY_WITH_SHIFT 0x01
#define END_KEY_WITHOUT_SHIFT 0x02
#define END_KEY_ENABLED (END_KEY_WITH_SHIFT | END_KEY_WITHOUT_SHIFT)
#define INPUT_KEY_SUPPRESS 0x04
#define INPUT_KEY_VISIBLE 0x08
#define INPUT_KEY_VISIBILITY_MASK (INPUT_KEY_SUPPRESS | INPUT_KEY_VISIBLE)
#define INPUT_KEY_IGNORE_TEXT 0x10
#define INPUT_KEY_NOTIFY 0x20
#define INPUT_KEY_OPTION_MASK 0x3F
#define INPUT_KEY_IS_TEXT 0x40
#define INPUT_KEY_DOWN_SUPPRESSED 0x80

class InputObject;
struct input_type
{
	InputStatusType Status = INPUT_OFF;
	input_type *Prev = nullptr;
	InputObject *ScriptObject = nullptr;
	LPTSTR Buffer = nullptr; // Stores the user's actual input.
	int BufferLength = 0; // The current length of what the user entered.
	int BufferLengthMax = INPUTHOOK_BUFFER_SIZE - 1; // The maximum allowed length of the input.
	TCHAR *EndChars = nullptr; // A string of characters that should terminate the input.
	UINT EndCharsMax = 0; // Current size of EndChars buffer.
	LPTSTR *match = nullptr; // Array of strings, each string is a match-phrase which if entered, terminates the input.
	UINT MatchCount; // The number of strings currently in the array.
	UINT MatchCountMax; // The maximum number of strings that the match array can contain.
	#define INPUT_ARRAY_BLOCK_SIZE 1024  // The increment by which the above array expands.
	LPTSTR MatchBuf = nullptr; // The is the buffer whose contents are pointed to by the match array.
	UINT MatchBufSize = 0; // The capacity of the above buffer.
	int Timeout = 0;
	DWORD TimeoutAt;
	SendLevelType MinSendLevel = 0; // The minimum SendLevel that can be captured by this input (0 allows all).
	bool BackspaceIsUndo = true;
	bool CaseSensitive = false;
	bool TranscribeModifiedKeys = false; // Whether the input command will attempt to transcribe modified keys such as ^c.
	bool VisibleText = false, VisibleNonText = true;
	bool NotifyNonText = false;
	bool FindAnywhere = false;
	bool EndCharMode = false;
	bool BeforeHotkeys = false;
	vk_type EndingVK; // The hook puts the terminating key into one of these if that's how it was terminated.
	sc_type EndingSC;
	TCHAR EndingChar;
	bool EndingBySC; // Whether the Ending key was one handled by VK or SC.
	bool EndingRequiredShift; // Whether the key that terminated the input was one that needed the SHIFT key.
	modLR_type EndingMods = 0;
	UINT EndingMatchIndex;
	UCHAR KeyVK[VK_ARRAY_COUNT] {}; // A sparse array of key flags by VK.
	UCHAR KeySC[SC_ARRAY_COUNT] {}; // A sparse array of key flags by SC.
	
	input_type::input_type() {}
	~input_type()
	{
		free(Buffer);
		free(match);
		free(MatchBuf);
		if (EndCharsMax) // If zero, EndChars may point to static memory.
			free(EndChars);
	}

	inline bool InProgress() { return Status == INPUT_IN_PROGRESS; }
	bool IsEarly() { return BeforeHotkeys; }
	bool IsInteresting(KBDLLHOOKSTRUCT &aEvent);
	ResultType Setup(LPCTSTR aOptions, LPCTSTR aEndKeys, LPCTSTR aMatchList);
	void ParseOptions(LPCTSTR aOptions);
	void SetTimeoutTimer();
	ResultType SetKeyFlags(LPCTSTR aKeys, bool aEndKeyMode = true, UCHAR aFlagsRemove = 0, UCHAR aFlagsAdd = END_KEY_ENABLED);
	ResultType SetMatchList(LPCTSTR aMatchList);
	void Start();
	void EndByMatch(UINT aMatchIndex);
	void EndByKey(vk_type aVK, sc_type aSC, bool aBySC, bool aRequiredShift);
	void EndByChar(TCHAR aChar);
	void EndByTimeout() { EndByReason(INPUT_TIMED_OUT); }
	void EndByLimit() { EndByReason(INPUT_LIMIT_REACHED); }
	void Stop() { EndByReason(INPUT_OFF); }
	void CollectChar(TCHAR *ch, int char_count);
	LPTSTR GetEndReason(LPTSTR aKeyBuf, int aKeyBufSize);
private:
	void EndByReason(InputStatusType aReason);
};

#include "input_object.h"

void InputStart(input_type &input);
input_type **InputFindLink(input_type *aInput);
input_type *InputUnlinkIfStopped(input_type *aInput);
input_type *InputRelease(input_type *aInput);
input_type *InputFind(InputObject *object);


//-------------------------------------------

struct KeyHistoryItem
{
	vk_type vk;
	sc_type sc;
	TCHAR event_type; // space=none, i=ignored, s=suppressed, h=hotkey, etc.
	bool key_up;
	float elapsed_time;  // Time since prior key or mouse button, in seconds.
	// It seems better to store the foreground window's title rather than its HWND since keystrokes
	// might result in a window closing (being destroyed), in which case the displayed key history
	// would not be able to display the title at the time the history is displayed, which would
	// be undesirable.
	// To save mem, could point this into a shared buffer instead, but if that buffer were to run
	// out of space (perhaps due to the target window changing frequently), window logging would
	// no longer be possible without adding complexity to the logging function.  Seems best
	// to keep it simple:
#define KEY_HISTORY_WINDOW_TITLE_SIZE 100
	TCHAR target_window[KEY_HISTORY_WINDOW_TITLE_SIZE];
};

struct CollectInputState
{
	bool early_collected, used_dead_key_non_destructively;
	TCHAR ch[2];
	int char_count;
	HWND active_window;
	HKL keyboard_layout;
};

//-------------------------------------------


LRESULT CALLBACK LowLevelKeybdProc(int aCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK LowLevelMouseProc(int aCode, WPARAM wParam, LPARAM lParam);
LRESULT LowLevelCommon(const HHOOK aHook, int aCode, WPARAM wParam, LPARAM lParam, const vk_type aVK
	, sc_type aSC, bool aKeyUp, ULONG_PTR aExtraInfo, DWORD aEventFlags);

#define HOTSTRING_INDEX_INVALID INT_MAX  // Use a signed maximum rather than unsigned, in case indexes ever become signed.
#define SuppressThisKey SuppressThisKeyFunc(aHook, lParam, aVK, aSC, aKeyUp, aExtraInfo, pKeyHistoryCurr, hotkey_id_to_post)
LRESULT SuppressThisKeyFunc(const HHOOK aHook, LPARAM lParam, const vk_type aVK, const sc_type aSC
	, bool aKeyUp, ULONG_PTR aExtraInfo, KeyHistoryItem *pKeyHistoryCurr, WPARAM aHotkeyIDToPost
	, WPARAM aHSwParamToPost = HOTSTRING_INDEX_INVALID, LPARAM aHSlParamToPost = 0);

#define AllowKeyToGoToSystem AllowIt(aHook, aCode, wParam, lParam, aVK, aSC, aKeyUp, aExtraInfo, collect_input_state, pKeyHistoryCurr, hotkey_id_to_post)
LRESULT AllowIt(const HHOOK aHook, int aCode, WPARAM wParam, LPARAM lParam, const vk_type aVK, const sc_type aSC
	, bool aKeyUp, ULONG_PTR aExtraInfo, CollectInputState &aState, KeyHistoryItem *pKeyHistoryCurr, WPARAM aHotkeyIDToPost);

bool EarlyCollectInput(KBDLLHOOKSTRUCT &aEvent, const vk_type aVK, const sc_type aSC, bool aKeyUp, bool aIsIgnored
	, CollectInputState &aState, KeyHistoryItem *pKeyHistoryCurr);
bool CollectInput(KBDLLHOOKSTRUCT &aEvent, const vk_type aVK, const sc_type aSC, bool aKeyUp, bool aIsIgnored
	, CollectInputState &aState, KeyHistoryItem *pKeyHistoryCurr, WPARAM &aHotstringWparamToPost, LPARAM &aHotstringLparamToPost);
bool CollectHotstring(KBDLLHOOKSTRUCT &aEvent, TCHAR aChar[], int aCharCount, HWND aActiveWindow
	, KeyHistoryItem *pKeyHistoryCurr, WPARAM &aHotstringWparamToPost, LPARAM &aHotstringLparamToPost);
bool CollectInputHook(KBDLLHOOKSTRUCT &aEvent, const vk_type aVK, const sc_type aSC, TCHAR aChar[], int aCharCount, bool aIsIgnored, bool aEarly);
bool IsHotstringWordChar(TCHAR aChar);
void UpdateKeybdState(KBDLLHOOKSTRUCT &aEvent, const vk_type aVK, const sc_type aSC, bool aKeyUp, bool aIsSuppressed);
bool KeybdEventIsPhysical(DWORD aEventFlags, const vk_type aVK, bool aKeyUp);

void ChangeHookState(Hotkey *aHK[], int aHK_count, HookType aWhichHook, HookType aWhichHookAlways);
void AddRemoveHooks(HookType aHooksToBeActive, bool aChangeIsTemporary = false);
bool HookAdjustMaxHotkeys(Hotkey **&aHK, int &aCurrentMax, int aNewMax);
bool SystemHasAnotherKeybdHook();
bool SystemHasAnotherMouseHook();
DWORD WINAPI HookThreadProc(LPVOID aUnused);

void LinkKeysForCustomCombo(vk_type aNeutral, vk_type aLeft, vk_type aRight);

void ResetHook(bool aAllModifiersUp = false, HookType aWhichHook = (HOOK_KEYBD | HOOK_MOUSE)
	, bool aResetKVKandKSC = false);
HookType GetActiveHooks();
void FreeHookMem();
void ResetKeyTypeState(key_type &key);
void GetHookStatus(LPTSTR aBuf, int aBufSize);

void WaitHookIdle();

#endif
