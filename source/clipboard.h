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

#include "stdafx.h" // pre-compiled headers
#include "defines.h"


#define CANT_OPEN_CLIPBOARD_READ "Can't open clipboard for reading."
#define CANT_OPEN_CLIPBOARD_WRITE "Can't open clipboard for writing."


class Clipboard
{
public:
	HGLOBAL mClipMemNow, mClipMemNew;
	char *mClipMemNowLocked, *mClipMemNewLocked;
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
	HANDLE GetClipboardDataTimeout(UINT uFormat);

	// Below: Whether the clipboard is ready to be written to.  Note that the clipboard is not
	// usually physically open even when this is true, unless the caller specifically opened
	// it in that mode also:
	bool IsReadyForWrite() {return mClipMemNewLocked != NULL;}

	#define CLIPBOARD_FAILURE UINT_MAX
	size_t Get(char *aBuf = NULL);

	ResultType Set(char *aBuf = NULL, UINT aLength = UINT_MAX); //, bool aTrimIt = false);
	char *PrepareForWrite(size_t aAllocSize);
	ResultType Commit(UINT aFormat = CF_TEXT);
	ResultType AbortWrite(char *aErrorMessage = "");
	ResultType Close(char *aErrorMessage = NULL);
	char *Contents()
	{
		if (mClipMemNewLocked)
			// Its set up for being written to, which takes precedence over the fact
			// that it may be open for read also, so return the write-buffer:
			return mClipMemNewLocked;
		if (!IsClipboardFormatAvailable(CF_TEXT))
			// We check for both CF_TEXT and CF_HDROP in case it's possible for
			// the clipboard to contain both formats simultaneously.  In this case,
			// just do this for now, to remind the caller that in these cases, it should
			// call Get(buf), providing the target buf so that we have some memory to
			// transcribe the (potentially huge) list of files into:
			return IsClipboardFormatAvailable(CF_HDROP) ? "<<>>" : "";
		else
			// Some callers may rely upon receiving empty string rather than NULL on failure:
			return (Get() == CLIPBOARD_FAILURE) ? "" : mClipMemNowLocked;
	}

	Clipboard() // Constructor
		: mIsOpen(false)  // Assumes our app doesn't already have it open.
		, mClipMemNow(NULL), mClipMemNew(NULL)
		, mClipMemNowLocked(NULL), mClipMemNewLocked(NULL)
		, mLength(0), mCapacity(0)
	{}
};


#endif
