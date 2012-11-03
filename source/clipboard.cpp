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
#include "clipboard.h"
#include "globaldata.h"  // for g_script.ScriptError() and g_ClipboardTimeout
#include "application.h" // for MsgSleep()
#include "util.h" // for strlcpy()

size_t Clipboard::Get(LPTSTR aBuf)
// If aBuf is NULL, it returns the length of the text on the clipboard and leaves the
// clipboard open.  Otherwise, it copies the clipboard text into aBuf and closes
// the clipboard (UPDATE: But only if the clipboard is still open from a prior call
// to determine the length -- see later comments for details).  In both cases, the
// length of the clipboard text is returned (or the value CLIPBOARD_FAILURE if error).
// If the clipboard is still open when the next MsgSleep() is called -- presumably
// because the caller never followed up with a second call to this function, perhaps
// due to having insufficient memory -- MsgSleep() will close it so that our
// app doesn't keep the clipboard tied up.  Note: In all current cases, the caller
// will use MsgBox to display an error, which in turn calls MsgSleep(), which will
// immediately close the clipboard.
{
	// Seems best to always have done this even if we return early due to failure:
	if (aBuf)
		// It should be safe to do this even at its peak capacity, because caller
		// would have then given us the last char in the buffer, which is already
		// a zero terminator, so this would have no effect:
		*aBuf = '\0';

	UINT i, file_count = 0;
	BOOL clipboard_contains_text = IsClipboardFormatAvailable(CF_NATIVETEXT);
	BOOL clipboard_contains_files = IsClipboardFormatAvailable(CF_HDROP);
	if (!(clipboard_contains_text || clipboard_contains_files))
		return 0;

	if (!mIsOpen)
	{
		// As a precaution, don't give the caller anything from the clipboard
		// if the clipboard isn't already open from the caller's previous
		// call to determine the size of what's on the clipboard (no other app
		// can alter its size while we have it open).  The is to prevent a
		// buffer overflow from happening in a scenario such as the following:
		// Caller calls us and we return zero size, either because there's no
		// CF_TEXT on the clipboard or there was a problem opening the clipboard.
		// In these two cases, the clipboard isn't open, so by the time the
		// caller calls us again, there's a chance (vanishingly small perhaps)
		// that another app (if our thread were preempted long enough, or the
		// platform is multiprocessor) will have changed the contents of the
		// clipboard to something larger than zero.  Thus, if we copy that
		// into the caller's buffer, the buffer might overflow:
		if (aBuf)
			return 0;
		if (!Open())
		{
			// Since this should be very rare, a shorter message is now used.  Formerly, it was
			// "Could not open clipboard for reading after many timed attempts. Another program is probably holding it open."
			Close(CANT_OPEN_CLIPBOARD_READ);
			return CLIPBOARD_FAILURE;
		}
		if (   !(mClipMemNow = g_clip.GetClipboardDataTimeout(clipboard_contains_files ? CF_HDROP : CF_NATIVETEXT))   )
		{
			// v1.0.47.04: Commented out the following that had been in effect when clipboard_contains_files==false:
			//    Close("GetClipboardData"); // Short error message since so rare.
			//    return CLIPBOARD_FAILURE;
			// This was done because there are situations when GetClipboardData can fail indefinitely.
			// For example, in Firefox, pulling down the Bookmarks menu then right-clicking "Bookmarks Toolbar
			// Folder" then selecting "Copy" puts one or more formats on the clipboard that cause this problem.
			// For details, search the forum for TYMED_NULL.
			//
			// v1.0.42.03: For the fix below, GetClipboardDataTimeout() knows not to try more than once
			// for CF_HDROP.
			// Fix for v1.0.31.02: When clipboard_contains_files==true, tolerate failure, which happens
			// as a normal/expected outcome when there are files on the clipboard but either:
			// 1) zero of them;
			// 2) the CF_HDROP on the clipboard is somehow misformatted.
			// If you select the parent ".." folder in WinRar then use the following hotkey, the script
			// would previously yield a runtime error:
			//#q::
			//Send, ^c
			//ClipWait, 0.5, 1
			//msgbox %Clipboard%
			//Return
			Close();
			if (aBuf)
				*aBuf = '\0';
			return CLIPBOARD_FAILURE; // Return this because otherwise, Contents() returns mClipMemNowLocked, which is NULL.
		}
		// Although GlobalSize(mClipMemNow) can yield zero in some cases -- in which case GlobalLock() should
		// not be attempted -- it probably can't yield zero for CF_HDROP and CF_TEXT because such a thing has
		// never been reported by anyone.  Therefore, GlobalSize() is currently not called.
		if (   !(mClipMemNowLocked = (LPTSTR)GlobalLock(mClipMemNow))   )
		{
			Close(_T("GlobalLock"));  // Short error message since so rare.
			return CLIPBOARD_FAILURE;
		}
		// Otherwise: Update length after every successful new open&lock:
		// Determine the length (size - 1) of the buffer than would be
		// needed to hold what's on the clipboard:
		if (clipboard_contains_files)
		{
			if (file_count = DragQueryFile((HDROP)mClipMemNowLocked, 0xFFFFFFFF, _T(""), 0))
			{
				mLength = (file_count - 1) * 2;  // Init; -1 if don't want a newline after last file.
				for (i = 0; i < file_count; ++i)
					mLength += DragQueryFile((HDROP)mClipMemNowLocked, i, NULL, 0);
			}
			else
				mLength = 0;
		}
		else // clipboard_contains_text
			mLength = _tcslen(mClipMemNowLocked);
		if (mLength >= CLIPBOARD_FAILURE) // Can't realistically happen, so just indicate silent failure.
			return CLIPBOARD_FAILURE;
	}
	if (!aBuf)
		return mLength;
		// Above: Just return the length; don't close the clipboard because we expect
		// to be called again soon.  If for some reason we aren't called, MsgSleep()
		// will automatically close the clipboard and clean up things.  It's done this
		// way to avoid the chance that the clipboard contents (and thus its length)
		// will change while we don't have it open, possibly resulting in a buffer
		// overflow.  In addition, this approach performs better because it avoids
		// the overhead of having to close and reopen the clipboard.

	// Otherwise:
	if (clipboard_contains_files)
	{
		if (file_count = DragQueryFile((HDROP)mClipMemNowLocked, 0xFFFFFFFF, _T(""), 0))
			for (i = 0; i < file_count; ++i)
			{
				// Caller has already ensured aBuf is large enough to hold them all:
				aBuf += DragQueryFile((HDROP)mClipMemNowLocked, i, aBuf, 999);
				if (i < file_count - 1) // i.e. don't add newline after the last filename.
				{
					*aBuf++ = '\r';  // These two are the proper newline sequence that the OS prefers.
					*aBuf++ = '\n';
				}
				//else DragQueryFile() has ensured that aBuf is terminated.
			}
		// else aBuf has already been terminated upon entrance to this function.
	}
	else
		_tcscpy(aBuf, mClipMemNowLocked);  // Caller has already ensured that aBuf is large enough.
	// Fix for v1.0.37: Close() is no longer called here because it prevents the clipboard variable
	// from being referred to more than once in a line.  For example:
	// Msgbox %Clipboard%%Clipboard%
	// ToolTip % StrLen(Clipboard) . Clipboard
	// Instead, the clipboard is later closed in other places (search on CLOSE_CLIPBOARD_IF_OPEN
	// to find them).  The alternative to fixing it this way would be to let it reopen the clipboard
	// by means getting rid of the following lines above:
	//if (aBuf)
	//	return 0;
	// However, that has the risks described in the comments above those two lines.
	return mLength;
}



ResultType Clipboard::Set(LPCTSTR aBuf, UINT_PTR aLength)
// Returns OK or FAIL.
{
	// It was already open for writing from a prior call.  Return failure because callers that do this
	// are probably handling things wrong:
	if (IsReadyForWrite()) return FAIL;

	if (!aBuf)
	{
		aBuf = _T("");
		aLength = 0;
	}
	else
		if (aLength == UINT_MAX) // Caller wants us to determine the length.
			aLength = (UINT)_tcslen(aBuf);

	if (aLength)
	{
		if (!PrepareForWrite(aLength + 1))
			return FAIL;  // It already displayed the error.
		tcslcpy(mClipMemNewLocked, aBuf, aLength + 1);  // Copy only a substring, if aLength specifies such.
	}
	// else just do the below to empty the clipboard, which is different than setting
	// the clipboard equal to the empty string: it's not truly empty then, as reported
	// by IsClipboardFormatAvailable(CF_TEXT) -- and we want to be able to make it truly
	// empty for use with functions such as ClipWait:
	return Commit();  // It will display any errors.
}



LPTSTR Clipboard::PrepareForWrite(size_t aAllocSize)
{
	if (!aAllocSize) return NULL; // Caller should ensure that size is at least 1, i.e. room for the zero terminator.
	if (IsReadyForWrite())
		// It was already prepared due to a prior call.  Currently, the most useful thing to do
		// here is return the memory area that's already been reserved:
		return mClipMemNewLocked;
	// Note: I think GMEM_DDESHARE is recommended in addition to the usual GMEM_MOVEABLE:
	// UPDATE: MSDN: "The following values are obsolete, but are provided for compatibility
	// with 16-bit Windows. They are ignored.": GMEM_DDESHARE
	if (   !(mClipMemNew = GlobalAlloc(GMEM_MOVEABLE, aAllocSize * sizeof(TCHAR)))   )
	{
		g_script.ScriptError(_T("GlobalAlloc"));  // Short error message since so rare.
		return NULL;
	}
	if (   !(mClipMemNewLocked = (LPTSTR)GlobalLock(mClipMemNew))   )
	{
		mClipMemNew = GlobalFree(mClipMemNew);  // This keeps mClipMemNew in sync with its state.
		g_script.ScriptError(_T("GlobalLock")); // Short error message since so rare.
		return NULL;
	}
	mCapacity = (UINT)aAllocSize; // Keep mCapacity in sync with the state of mClipMemNewLocked.
	*mClipMemNewLocked = '\0'; // Init for caller.
	return mClipMemNewLocked;  // The caller can now write to this mem.
}



ResultType Clipboard::Commit(UINT aFormat)
// If this is called while mClipMemNew is NULL, the clipboard will be set to be truly
// empty, which is different from writing an empty string to it.  Note: If the clipboard
// was already physically open, this function will close it as part of the commit (since
// whoever had it open before can't use the prior contents, since they're invalid).
{
	if (!mIsOpen && !Open())
		// Since this should be very rare, a shorter message is now used.  Formerly, it was
		// "Could not open clipboard for writing after many timed attempts.  Another program is probably holding it open."
		return AbortWrite(CANT_OPEN_CLIPBOARD_WRITE);
	if (!EmptyClipboard())
	{
		Close();
		return AbortWrite(_T("EmptyClipboard")); // Short error message since so rare.
	}
	if (mClipMemNew)
	{
		bool new_is_empty = false;
		// Unlock prior to calling SetClipboardData:
		if (mClipMemNewLocked) // probably always true if we're here.
		{
			// Best to access the memory while it's still locked, which is why this temp var is used:
			// v1.0.40.02: The following was fixed to properly recognize 0x0000 as the Unicode string terminator,
			// which fixes problems with Transform Unicode.
			new_is_empty = aFormat == CF_UNICODETEXT ? !*(LPWSTR)mClipMemNewLocked : !*(LPSTR)mClipMemNewLocked;
			GlobalUnlock(mClipMemNew); // mClipMemNew not mClipMemNewLocked.
			mClipMemNewLocked = NULL;  // Keep this in sync with the above action.
			mCapacity = 0; // Keep mCapacity in sync with the state of mClipMemNewLocked.
		}
		if (new_is_empty)
			// Leave the clipboard truly empty rather than setting it to be the
			// empty string (i.e. these two conditions are NOT the same).
			// But be sure to free the memory since we're not giving custody
			// of it to the system:
			mClipMemNew = GlobalFree(mClipMemNew);
		else
			if (SetClipboardData(aFormat, mClipMemNew))
				// In any of the failure conditions above, Close() ensures that mClipMemNew is
				// freed if it was allocated.  But now that we're here, the memory should not be
				// freed because it is owned by the clipboard (it will free it at the appropriate time).
				// Thus, we relinquish the memory because we shouldn't be looking at it anymore:
				mClipMemNew = NULL;
			else
			{
				Close();
				return AbortWrite(_T("SetClipboardData")); // Short error message since so rare.
			}
	}
	// else we will close it after having done only the EmptyClipboard(), above.
	// Note: Decided not to update mLength for performance reasons (in case clipboard is huge).
	// Anyway, it seems rather pointless because once the clipboard is closed, our app instantly
	// loses sight of how large it is, so the the value of mLength wouldn't be reliable unless
	// the clipboard were going to be immediately opened again.
	return Close();
}



ResultType Clipboard::AbortWrite(LPTSTR aErrorMessage)
// Always returns FAIL.
{
	// Since we were called in conjunction with an aborted attempt to Commit(), always
	// ensure the clipboard is physically closed because even an attempt to Commit()
	// should physically close it:
	Close();
	if (mClipMemNewLocked)
	{
		GlobalUnlock(mClipMemNew); // mClipMemNew not mClipMemNewLocked.
		mClipMemNewLocked = NULL;
		mCapacity = 0; // Keep mCapacity in sync with the state of mClipMemNewLocked.
	}
	// Above: Unlock prior to freeing below.
	if (mClipMemNew)
		mClipMemNew = GlobalFree(mClipMemNew);
	// Caller needs us to always return FAIL:
	return *aErrorMessage ? g_script.ScriptError(aErrorMessage) : FAIL;
}



ResultType Clipboard::Close(LPTSTR aErrorMessage)
// Returns OK or FAIL (but it only returns FAIL if caller gave us a non-NULL aErrorMessage).
{
	// Always close it ASAP so that it is free for other apps to use:
	if (mIsOpen)
	{
		if (mClipMemNowLocked)
		{
			GlobalUnlock(mClipMemNow); // mClipMemNow not mClipMemNowLocked.
			mClipMemNowLocked = NULL;  // Keep this in sync with its state, since it's used as an indicator.
		}
		// Above: It's probably best to unlock prior to closing the clipboard.
		CloseClipboard();
		mIsOpen = false;  // Even if above fails (realistically impossible?), seems best to do this.
		// Must do this only after GlobalUnlock():
		mClipMemNow = NULL;
	}
	// Do this cleanup for callers that didn't make it far enough to even open the clipboard.
	// UPDATE: DO *NOT* do this because it is valid to have the clipboard in a "ReadyForWrite"
	// state even after we physically close it.  Some callers rely on that.
	//if (mClipMemNewLocked)
	//{
	//	GlobalUnlock(mClipMemNew); // mClipMemNew not mClipMemNewLocked.
	//	mClipMemNewLocked = NULL;
	//	mCapacity = 0; // Keep mCapacity in sync with the state of mClipMemNewLocked.
	//}
	//// Above: Unlock prior to freeing below.
	//if (mClipMemNew)
	//	// Commit() was never called after a call to PrepareForWrite(), so just free the memory:
	//	mClipMemNew = GlobalFree(mClipMemNew);
	if (aErrorMessage && *aErrorMessage)
		// Caller needs us to always return FAIL if an error was displayed:
		return g_script.ScriptError(aErrorMessage);

	// Seems best not to reset mLength.  But it will quickly become out of date once
	// the clipboard has been closed and other apps can use it.
	return OK;
}



HANDLE Clipboard::GetClipboardDataTimeout(UINT uFormat)
// Same as GetClipboardData() except that it doesn't give up if the first call to GetClipboardData() fails.
// Instead, it continues to retry the operation for the number of milliseconds in g_ClipboardTimeout.
// This is necessary because GetClipboardData() has been observed to fail in repeatable situations (this
// is strange because our thread already has the clipboard locked open -- presumably it happens because the
// GetClipboardData() is unable to start a data stream from the application that actually serves up the data).
// If cases where the first call to GetClipboardData() fails, a subsequent call will often succeed if you give
// the owning application (such as Excel and Word) a little time to catch up.  This is especially necessary in
// the OnClipboardChange label, where sometimes a clipboard-change notification comes in before the owning
// app has finished preparing its data for subsequent readers of the clipboard.
{
#ifdef DEBUG_BY_LOGGING_CLIPBOARD_FORMATS  // Provides a convenient log of clipboard formats for analysis.
	static FILE *fp = fopen("c:\\debug_clipboard_formats.txt", "w");
#endif

	TCHAR format_name[MAX_PATH + 1]; // MSDN's RegisterClipboardFormat() doesn't document any max length, but the ones we're interested in certainly don't exceed MAX_PATH.
	if (uFormat < 0xC000 || uFormat > 0xFFFF) // It's a registered format (you're supposed to verify in-range before calling GetClipboardFormatName()).  Also helps performance.
		*format_name = '\0'; // Don't need the name if it's a standard/CF_* format.
	else
	{
		// v1.0.42.04:
		// Probably need to call GetClipboardFormatName() rather than comparing directly to uFormat because
		// MSDN implies that OwnerLink and other registered formats might not always have the same ID under
		// all OSes (past and future).
		GetClipboardFormatName(uFormat, format_name, MAX_PATH);
		// Since RegisterClipboardFormat() is case insensitive, the case might vary.  So use stricmp() when
		// comparing format_name to anything.
		// "Link Source", "Link Source Descriptor" , and anything else starting with "Link Source" is likely
		// to be data that should not be attempted to be retrieved because:
		// 1) It causes unwanted bookmark effects in various versions of MS Word.
		// 2) Tests show that these formats are on the clipboard only if MS Word is open at the time
		//    ClipboardAll is accessed.  That implies they're transitory formats that aren't as essential
		//    or well suited to ClipboardAll as the other formats (but if it weren't for #1 above, this
		//    wouldn't be enough reason to omit it).
		// 3) Although there is hardly any documentation to be found at MSDN or elsewhere about these formats,
		//    it seems they're related to OLE, with further implications that the data is transitory.
		// Here are the formats that Word 2002 removes from the clipboard when it the app closes:
		// 0xC002 ObjectLink  >>> Causes WORD bookmarking problem.
		// 0xC003 OwnerLink
		// 0xC00D Link Source  >>> Causes WORD bookmarking problem.
		// 0xC00F Link Source Descriptor  >>> Doesn't directly cause bookmarking, but probably goes with above.
		// 0xC0DC Hyperlink
		if (   !_tcsnicmp(format_name, _T("Link Source"), 11) || !_tcsicmp(format_name, _T("ObjectLink"))
			|| !_tcsicmp(format_name, _T("OwnerLink"))
			// v1.0.44.07: The following were added to solve interference with MS Outlook's MS Word editor.
			// If a hotkey like ^F1::ClipboardSave:=ClipboardAll is pressed after pressing Ctrl-C in that
			// editor (perhaps only when copying HTML), two of the following error dialogs would otherwise
			// be displayed (this occurs in Outlook 2002 and probably later versions):
			// "An outgoing call cannot be made since the application is dispatching an input-synchronous call."
			|| !_tcsicmp(format_name, _T("Native")) || !_tcsicmp(format_name, _T("Embed Source"))   )
			return NULL;
	}

#ifdef DEBUG_BY_LOGGING_CLIPBOARD_FORMATS
	_ftprintf(fp, _T("%04X\t%s\n"), uFormat, format_name);  // Since fclose() is never called, the program has to exit to close/release the file.
#endif

	HANDLE h;
	for (DWORD start_time = GetTickCount();;)
	{
		// Known failure conditions:
		// GetClipboardData() apparently fails when the text on the clipboard is greater than a certain size
		// (Even though GetLastError() reports "Operation completed successfully").  The data size at which
		// this occurs is somewhere between 20 to 96 MB (perhaps depending on system's memory and CPU speed).
		if (h = GetClipboardData(uFormat)) // Assign
			return h;

		// It failed, so act according to the type of format and the timeout that's in effect.
		// Certain standard (numerically constant) clipboard formats are known to validly yield NULL from a
		// call to GetClipboardData().  Never retry these because it would only cause unnecessary delays
		// (i.e. a failure until timeout).
		// v1.0.42.04: More importantly, retrying them appears to cause problems with saving a Word/Excel
		// clipboard via ClipboardAll.
		if (uFormat == CF_HDROP // This format can fail "normally" for the reasons described at "clipboard_contains_files".
			|| !_tcsicmp(format_name, _T("OwnerLink"))) // Known to validly yield NULL from a call to GetClipboardData(), so don't retry it to avoid having to wait the full timeout period.
			return NULL;

		if (g_ClipboardTimeout != -1) // We were not told to wait indefinitely and...
			if (!g_ClipboardTimeout   // ...we were told to make only one attempt, or ...
				|| (int)(g_ClipboardTimeout - (GetTickCount() - start_time)) <= SLEEP_INTERVAL_HALF) //...it timed out.
				// Above must cast to int or any negative result will be lost due to DWORD type.
				return NULL;

		// Use SLEEP_WITHOUT_INTERRUPTION to prevent MainWindowProc() from accepting new hotkeys
		// during our operation, since a new hotkey subroutine might interfere with
		// what we're doing here (e.g. if it tries to use the clipboard, or perhaps overwrites
		// the deref buffer if this object's caller gave it any pointers into that memory area):
		SLEEP_WITHOUT_INTERRUPTION(INTERVAL_UNSPECIFIED)
	}
}



ResultType Clipboard::Open()
{
	if (mIsOpen)
		return OK;
	for (DWORD start_time = GetTickCount();;)
	{
		if (OpenClipboard(g_hWnd))
		{
			mIsOpen = true;
			return OK;
		}
		if (g_ClipboardTimeout != -1) // We were not told to wait indefinitely...
			if (!g_ClipboardTimeout   // ...and we were told to make only one attempt, or ...
				|| (int)(g_ClipboardTimeout - (GetTickCount() - start_time)) <= SLEEP_INTERVAL_HALF) //...it timed out.
				// Above must cast to int or any negative result will be lost due to DWORD type.
				return FAIL;
		// Use SLEEP_WITHOUT_INTERRUPTION to prevent MainWindowProc() from accepting new hotkeys
		// during our operation, since a new hotkey subroutine might interfere with
		// what we're doing here (e.g. if it tries to use the clipboard, or perhaps overwrites
		// the deref buffer if this object's caller gave it any pointers into that memory area):
		SLEEP_WITHOUT_INTERRUPTION(INTERVAL_UNSPECIFIED)
	}
}
