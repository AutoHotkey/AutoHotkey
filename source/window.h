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

#ifndef window_h
#define window_h

#include "defines.h"
#include "globaldata.h"
#include "util.h" // for strlcpy()


// Note: it is apparently possible for a hidden window to be the foreground
// window (it just looks strange).  If DetectHiddenWindows is off, set
// target_window to NULL if it's hidden.  Doing this prevents, for example,
// WinClose() from closing the hidden foreground if it's some important hidden
// window like the shell or the desktop:
#define USE_FOREGROUND_WINDOW(title, text, exclude_title, exclude_text)\
	((*title == 'A' || *title == 'a') && !*(title + 1) && !*text && !*exclude_title && !*exclude_text)
#define SET_TARGET_TO_ALLOWABLE_FOREGROUND(detect_hidden_windows) \
{\
	if (target_window = GetForegroundWindow())\
		if (!(detect_hidden_windows) && !IsWindowVisible(target_window))\
			target_window = NULL;\
}
#define IF_USE_FOREGROUND_WINDOW(detect_hidden_windows, title, text, exclude_title, exclude_text)\
if (USE_FOREGROUND_WINDOW(title, text, exclude_title, exclude_text))\
{\
	SET_TARGET_TO_ALLOWABLE_FOREGROUND(detect_hidden_windows)\
}



inline bool IsTextMatch(LPTSTR aHaystack, LPTSTR aNeedle)
// Generic helper function used by WindowSearch and other things.
// To help performance, it's the caller's responsibility to ensure that all params are not NULL.
{
	if (!*aNeedle) // The empty string is always found, regardless of mode.
		return true;
	switch(g->TitleMatchMode)
	{
	case FIND_ANYWHERE:        return _tcsstr(aHaystack, aNeedle);
	case FIND_REGEX:           return RegExMatch(aHaystack, aNeedle);
	case FIND_IN_LEADING_PART: return !_tcsncmp(aHaystack, aNeedle, _tcslen(aNeedle));
	default: // Otherwise: Exact match.
		return !_tcscmp(aHaystack, aNeedle); 
	}
}



#define SEARCH_PHRASE_SIZE 1024
// Info from AutoIt3 source: GetWindowText fails under 95 if >65535, WM_GETTEXT randomly fails if > 32767.
// My: And since 32767 is what AutoIt3 passes to the API functions as the size (not the length, i.e.
// it can only store 32766 if room is left for the zero terminator) we'll use that for the size too.
// Note: MSDN says (for functions like GetWindowText): "Specifies the maximum number of characters to
// copy to the buffer, including the NULL character. If the text exceeds this limit, it is truncated."
#define WINDOW_TEXT_SIZE 32767
#define WINDOW_CLASS_SIZE 257  // MSDN implies length can't be greater than 256: "The maximum length for [WNDCLASS] lpszClassName is 256. If lpszClassName is greater than the maximum length, the RegisterClass function will fail."

// Bitwise fields to support multiple criteria in v1.0.36.02
#define CRITERION_TITLE 0x01
#define CRITERION_ID    0x02
#define CRITERION_PID   0x04
#define CRITERION_CLASS 0x08
#define CRITERION_GROUP 0x10
#define CRITERION_PATH	0x20

class WindowSearch
{
	// One of the reasons for having this class is to avoid fetching PID, Class, and Window Text
	// when only the criteria have changed but not the candidate window.  This happens when called
	// from the WinGroup class.  Another reason is that it's probably more understandable than
	// the old way, while eliminating some redundant code as well.

public:
	DWORD mCriteria; // Which criteria are currently in effect (ID, PID, Class, Title, etc.)

	// Controlled and initialized by SetCriteria():
	global_struct *mSettings;                 // Settings such as TitleMatchMode and DetectHiddenWindows.
	TCHAR mCriterionTitle[SEARCH_PHRASE_SIZE]; // For storing the title.
	TCHAR mCriterionClass[SEARCH_PHRASE_SIZE]; // For storing the "ahk_class" class name.
	size_t mCriterionTitleLength;             // Length of mCriterionTitle.
	LPTSTR mCriterionExcludeTitle;             // ExcludeTitle.
	size_t mCriterionExcludeTitleLength;      // Length of the above.
	LPTSTR mCriterionText;                     // WinText.
	LPTSTR mCriterionExcludeText;              // ExcludeText.
	HWND mCriterionHwnd;                      // For "ahk_id".
	DWORD mCriterionPID;                      // For "ahk_pid".
	WinGroup *mCriterionGroup;                // For "ahk_group".
	TCHAR mCriterionPath[SEARCH_PHRASE_SIZE]; // For "ahk_exe".

	bool mCriterionPathIsNameOnly;
	bool mFindLastMatch; // Whether to keep searching even after a match is found, so that last one is found.
	int mFoundCount;     // Accumulates how many matches have been found (either 0 or 1 unless mFindLastMatch==true).
	HWND mFoundParent;   // Must be separate from mCandidateParent because some callers don't have access to IsMatch()'s return value.
	HWND mFoundChild;    // Needed by EnumChildFind() to store its result, and other things.

	HWND *mAlreadyVisited;      // Array of HWNDs to exclude from consideration.
	int mAlreadyVisitedCount;   // Count of items in the above.
	WindowSpec *mFirstWinSpec;  // Linked list used by the WinGroup commands.
	ActionTypeType mActionType; // Used only by WinGroup::PerformShowWindow().
	int mTimeToWaitForClose;    // Same.
	Var *mArrayStart;           // Used by WinGetList() to fetch an array of matching HWNDs.

	// Controlled by the SetCandidate() method:
	HWND mCandidateParent;
	DWORD mCandidatePID;
	TCHAR mCandidateTitle[WINDOW_TEXT_SIZE];  // For storing title or class name of the given mCandidateParent.
	TCHAR mCandidateClass[WINDOW_CLASS_SIZE]; // Must not share mem with mCandidateTitle because even if ahk_class is in effect, ExcludeTitle can also be in effect.
	TCHAR mCandidatePath[MAX_PATH];


	void SetCandidate(HWND aWnd) // Must be kept thread-safe since it may be called indirectly by the hook thread.
	{
		// For performance reasons, update the attributes only if the candidate window changed:
		if (mCandidateParent != aWnd)
		{
			mCandidateParent = aWnd;
			UpdateCandidateAttributes(); // In case mCandidateParent isn't NULL, update the PID/Class/etc. based on what was set above.
		}
	}

	ResultType SetCriteria(global_struct &aSettings, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText);
	void UpdateCandidateAttributes();
	HWND IsMatch(bool aInvert = false);

	WindowSearch() // Constructor.
		// For performance and code size, only the most essential members are initialized.
		// The others do not require it or are initialized by SetCriteria() or SetCandidate().
		: mCriteria(0), mCriterionExcludeTitle(_T("")) // ExcludeTitle is referenced often, so should be initialized.
		, mFoundCount(0), mFoundParent(NULL) // Must be initialized here since none of the member functions is allowed to do it.
		, mFoundChild(NULL) // ControlExist() relies upon this.
		, mCandidateParent(NULL)
		// The following must be initialized because it's the object user's responsibility to override
		// them in those relatively rare cases when they need to be.  WinGroup::ActUponAll() and
		// WinGroup::Deactivate() (and probably other callers) rely on these attributes being retained
		// after they were overridden even upon multiple subsequent calls to SetCriteria():
		, mFindLastMatch(false), mAlreadyVisited(NULL), mAlreadyVisitedCount(0), mFirstWinSpec(NULL), mArrayStart(NULL)
	{
	}
};



struct control_list_type
{
	// For something this simple, a macro is probably a lot less overhead that making this struct
	// non-POD and giving it a constructor:
	#define CL_INIT_CONTROL_LIST(cl) \
		cl.is_first_iteration = true;\
		cl.total_classes = 0;\
		cl.total_length = 0;\
		cl.buf_free_spot = cl.class_buf; // Points to the next available/writable place in the buf.
	bool fetch_hwnds;         // True if fetching HWND of each control rather than its ClassNN.
	bool is_first_iteration;  // Must be initialized to true by Enum's caller.
	int total_classes;        // Must be initialized to 0.
	VarSizeType total_length; // Must be initialized to 0.
	VarSizeType capacity;     // Must be initialized to size of the below buffer.
	LPTSTR target_buf;         // Caller sets it to NULL if only the total_length is to be retrieved.
	#define CL_CLASS_BUF_SIZE (32 * 1024) // Even if class names average 50 chars long, this supports 655 of them.
	TCHAR class_buf[CL_CLASS_BUF_SIZE];
	LPTSTR buf_free_spot;      // Must be initialized to point to the beginning of class_buf.
	#define CL_MAX_CLASSES 500  // The number of distinct class names that can be supported in a single window.
	LPTSTR class_name[CL_MAX_CLASSES]; // Array of distinct class names, stored consecutively in class_buf.
	int class_count[CL_MAX_CLASSES];  // The quantity found for each of the above classes.
};

struct MonitorInfoPackage // A simple struct to help with EnumDisplayMonitors().
{
	int count;
	#define COUNT_ALL_MONITORS INT_MIN  // A special value that can be assigned to the below.
	int monitor_number_to_find;  // If this is left as zero, it will find the primary monitor by default.
	MONITORINFOEX monitor_info_ex;
};

struct pid_and_hwnd_type
{
	DWORD pid;
	HWND hwnd;
};

struct length_and_buf_type
{
	size_t total_length;
	size_t capacity;
	LPTSTR buf;
};

struct class_and_hwnd_type
{
	LPTSTR class_name;
	bool is_found;
	int class_count;
	HWND hwnd;
};

struct point_and_hwnd_type
{
	POINT pt;
	RECT rect_found;
	HWND hwnd_found;
	double distance;
};


HWND WinActivate(global_struct &aSettings, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText
	, bool aFindLastMatch = false
	, HWND aAlreadyVisited[] = NULL, int aAlreadyVisitedCount = 0);
HWND SetForegroundWindowEx(HWND aTargetWindow);

// Defaulting to a non-zero wait-time solves a lot of script problems that would otherwise
// require the user to specify the last param (or use WinWaitClose):
#define DEFAULT_WINCLOSE_WAIT 20
HWND WinClose(global_struct &aSettings, LPTSTR aTitle, LPTSTR aText, int aTimeToWaitForClose = DEFAULT_WINCLOSE_WAIT
	, LPTSTR aExcludeTitle = _T(""), LPTSTR aExcludeText = _T(""), bool aKillIfHung = false);
HWND WinClose(HWND aWnd, int aTimeToWaitForClose = DEFAULT_WINCLOSE_WAIT, bool aKillIfHung = false);

HWND WinActive(global_struct &aSettings, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText
	, bool aUpdateLastUsed = false);

HWND WinExist(global_struct &aSettings, LPTSTR aTitle, LPTSTR aText, LPTSTR aExcludeTitle, LPTSTR aExcludeText
	, bool aFindLastMatch = false, bool aUpdateLastUsed = false
	, HWND aAlreadyVisited[] = NULL, int aAlreadyVisitedCount = 0);

HWND GetValidLastUsedWindow(global_struct &aSettings);

BOOL CALLBACK EnumParentFind(HWND hwnd, LPARAM lParam);
BOOL CALLBACK EnumChildFind(HWND hwnd, LPARAM lParam);


// Use a fairly long default for aCheckInterval since the contents of this function's loops
// might be somewhat high in overhead (especially SendMessageTimeout):
#define SB_DEFAULT_CHECK_INTERVAL 50
ResultType StatusBarUtil(Var *aOutputVar, HWND aBarHwnd, int aPartNumber = 1, LPTSTR aTextToWaitFor = _T("")
	, int aWaitTime = -1, int aCheckInterval = SB_DEFAULT_CHECK_INTERVAL);
HWND ControlExist(HWND aParentWindow, LPTSTR aClassNameAndNum = NULL);
BOOL CALLBACK EnumControlFind(HWND aWnd, LPARAM lParam);

#define MSGBOX_NORMAL (MB_OK | MB_SETFOREGROUND)
#define MSGBOX_TEXT_SIZE (1024 * 8)
#define DIALOG_TITLE_SIZE 1024
int MsgBox(int aValue);
int MsgBox(LPCTSTR aText = _T(""), UINT uType = MSGBOX_NORMAL, LPTSTR aTitle = NULL, double aTimeout = 0, HWND aOwner = NULL);
HWND FindOurTopDialog();
BOOL CALLBACK EnumDialog(HWND hwnd, LPARAM lParam);

HWND WindowOwnsOthers(HWND aWnd);
BOOL CALLBACK EnumParentFindOwned(HWND aWnd, LPARAM lParam);
HWND GetNonChildParent(HWND aWnd);
HWND GetTopChild(HWND aParent);
bool IsWindowHung(HWND aWnd);

// Defaults to a low timeout because a window may have hundreds of controls, and if the window
// is hung, each control might result in a delay of size aTimeout during an EnumWindows.
// It shouldn't need much time anyway since the moment the call to SendMessageTimeout()
// is made, our thread is suspended and the target thread's WindowProc called directly.
// In addition:
// Whenever using SendMessageTimeout(), our app will be unresponsive until
// the call returns, since our message loop isn't running.  In addition,
// if the keyboard or mouse hook is installed, the events will lag during
// this call.  So keep the timeout value fairly short.  UPDATE: Need a longer
// timeout because otherwise searching will be inconsistent / unreliable for the
// slow Title Match method, since some apps are lazy about keeping their
// message pumps running, such as during long disk I/O operations, and thus
// may sometimes (randomly) take a long time to respond to the WM_GETTEXT message.
// 5000 seems about the largest value that should ever be needed since this is what
// Windows uses as the cutoff for determining if a window has become "unresponsive":
int GetWindowTextTimeout(HWND aWnd, LPTSTR aBuf = NULL, INT_PTR aBufSize = 0, UINT aTimeout = 5000);
void SetForegroundLockTimeout();


// Notes about the below macro:
// Update for v1.0.40.01:
// In earlier versions, a critical thread that displayed a dialog would discard any pending events
// that were waiting to start new threads (since in most cases, the dialog message pump would
// route those events directly to a window proc, which would then repost them with a NULL hwnd
// to prevent bouncing, which in turn would cause the dialog pump to discard them).  To avoid
// this and make the behavior more useful and intuitive, this has been changed so that any
// pending threads will launch right before the dialog is displayed.  But later, when the user
// dismisses the dialog, the thread becomes critical again.
// 
// Update for v1.0.38.04: Rather than setting AllowInterruption unconditionally to
// true, make it reflect the state of g->ThreadIsCritical.  This increases flexibility by allowing
// threads to stay interrruptible even when they're displaying a dialog.  In such cases, an
// incoming thread-event such as a hotkey will get routed to our MainWindowProc by the dialog's
// message pump; and from there it will get reposted to our queue, and then get pumped again.
// This bouncing effect may impact performance slightly but seems warranted to maintain
// flexibility of the "Critical" command as well as its ability to buffer incoming events.
//
// If our thread's message queue has any message pending whose HWND member is NULL -- or even
// normal messages which would be routed back to the thread by the WindowProc() -- clean them
// out of the message queue before launching the dialog's message pump below.  That message pump
// doesn't know how to properly handle such messages (it would either lose them or dispatch them
// at times we don't want them dispatched).  But first ensure the current quasi-thread is
// interruptible, since it's about to display a dialog so there little benefit (and a high cost)
// to leaving it uninterruptible.  The "high cost" is that MsgSleep (our main message pump) would
// filter out (leave queued) certain messages if the script were uninterruptible.  Then when it
// returned, the dialog message pump below would start, and it would discard or misroute the
// messages.
// If this is not done, the following scenario would also be a problem:
// A newly launched thread (in its period of uninterruptibility) displays a dialog.  As a consequence,
// the dialog's message pump starts dispatching all messages.  If during this brief time (before the
// thread becomes interruptible) a hotkey/hotstring/custom menu item/gui thread is dispatched to one
// of our WindowProc's, and then posted to our thread via PostMessage(NULL,...), the item would be lost
// because the dialog message pump discards messages that lack an HWND (since it doesn't know how to
// dispatch them to a Window Proc).
// GetQueueStatus() is used because unlike PeekMessage() or GetMessage(), it might not yield
// our timeslice if the CPU is under heavy load, which would be good to improve performance here.
#define DIALOG_PREP bool thread_was_critical = DialogPrep();
// v1.0.40.01: Turning off critical during the dialog is relied upon by ResumeUnderlyingThread(),
// UninterruptibleTimeout(), and KILL_AUTOEXEC_TIMER.  Doing it this way also seems more maintainable
// than some other approach such as storing a new flag in the "g" struct that says whether it is currently
// displaying a dialog and waiting for it to finish.
#define DIALOG_END \
{\
	g->ThreadIsCritical = thread_was_critical;\
	g->AllowThreadToBeInterrupted = !thread_was_critical;\
}
bool DialogPrep();

#endif
