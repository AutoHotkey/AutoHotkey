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

#ifndef clipboard_h
#define clipboard_h

#include "defines.h"


#define CANT_OPEN_CLIPBOARD_READ _T("Can't open clipboard for reading.")
#define CANT_OPEN_CLIPBOARD_WRITE _T("Can't open clipboard for writing.")

#ifdef UNICODE
// In unicode version, always try CF_UNICODETEXT first, then CF_TEXT.
#define CF_NATIVETEXT	CF_UNICODETEXT
#define CF_OTHERTEXT	CF_TEXT
#else
#define CF_NATIVETEXT	CF_TEXT
#define CF_OTHERTEXT	CF_UNICODETEXT
#endif


class Clipboard
{
public:
	HGLOBAL mClipMemNow, mClipMemNew;
	LPTSTR mClipMemNowLocked, mClipMemNewLocked;
	// NOTE: Both mLength and mCapacity are count in characters (NOT in bytes).
	size_t mLength;  // Last-known length of the clipboard contents (for internal use only because it's valid only during certain specific times).
	UINT mCapacity;  // Capacity of mClipMemNewLocked.
	BOOL mIsOpen;  // Whether the clipboard is physically open due to action by this class.  BOOL vs. bool improves some benchmarks slightly due to this item being frequently checked.

	// It seems best to default to many attempts, because a failure
	// to open the clipboard may result in the early termination
	// of a large script due to the fear that it's generally
	// unsafe to continue in such cases.  Update: Increased default
	// number of attempts from 20 to 40 because Jason (Payam) reported
	// that he was getting an error on rare occasions (but not reproducible).
	ResultType Open();
	HANDLE GetClipboardDataTimeout(UINT uFormat, BOOL *aNullIsOkay = NULL);

	// Below: Whether the clipboard is ready to be written to.  Note that the clipboard is not
	// usually physically open even when this is true, unless the caller specifically opened
	// it in that mode also:
	bool IsReadyForWrite() {return mClipMemNewLocked != NULL;}

	#define CLIPBOARD_FAILURE UINT_MAX
	size_t Get(LPTSTR aBuf = NULL);

	ResultType Set(LPCTSTR aBuf = NULL, UINT_PTR aLength = UINT_MAX);
	LPTSTR PrepareForWrite(size_t aAllocSize);
	ResultType Commit(UINT aFormat = CF_NATIVETEXT);
	ResultType AbortWrite(LPTSTR aErrorMessage = _T(""));
	ResultType Close(LPTSTR aErrorMessage = NULL);

	Clipboard() // Constructor
		: mIsOpen(false)  // Assumes our app doesn't already have it open.
		, mClipMemNow(NULL), mClipMemNew(NULL)
		, mClipMemNowLocked(NULL), mClipMemNewLocked(NULL)
		, mLength(0), mCapacity(0)
	{}
};


#endif
