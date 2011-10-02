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
#include "var.h"
#include "globaldata.h" // for g_script


// Init static vars:
TCHAR Var::sEmptyString[] = _T(""); // For explanation, see its declaration in .h file.


ResultType Var::AssignHWND(HWND aWnd)
{
	// For backward compatibility, tradition, and the fact that operations involving HWNDs tend not to
	// be nearly as performance-critical as pure-math expressions, HWNDs are stored as a hex string,
	// and thus UpdateBinaryInt64() isn't called here.
	// Older comment: Always assign as hex for better compatibility with Spy++ and other apps that
	// report window handles.
	TCHAR buf[MAX_INTEGER_SIZE];
	buf[0] = '0';
	buf[1] = 'x';
	Exp32or64(_ultot,_ui64tot)((size_t)aWnd, buf + 2, 16);
	// If ever decide to assign a pure integer, keep in mind the type-casting comments in BIF_WinExistActive().
	return Assign(buf);
}



ResultType Var::Assign(Var &aVar)
// Assign some other variable to the "this" variable.
// Although this->Type() can be VAR_CLIPBOARD, caller must ensure that aVar.Type()==VAR_NORMAL.
// Returns OK or FAIL.
{
	// Below relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
	Var &source_var = aVar.mType == VAR_ALIAS ? *aVar.mAliasFor : aVar;
	Var &target_var = *(mType == VAR_ALIAS ? mAliasFor : this);

	switch (source_var.mAttrib & VAR_ATTRIB_TYPES) // This switch() method should squeeze a little more performance out of it compared to doing "&" for every attribute.  Only works for attributes that are mutually-exclusive, which these are.
	{
	case VAR_ATTRIB_IS_INT64:
		return target_var.Assign(source_var.mContentsInt64);
	case VAR_ATTRIB_IS_DOUBLE:
		return target_var.Assign(source_var.mContentsDouble);
	case VAR_ATTRIB_IS_OBJECT:
		return target_var.Assign(source_var.mObject);
	case VAR_ATTRIB_BINARY_CLIP:
		return target_var.AssignBinaryClip(source_var); // Caller wants a variable with binary contents assigned (copied) to another variable (usually VAR_CLIPBOARD).
	}
	// Otherwise:
	source_var.MaybeWarnUninitialized();
	return target_var.Assign(source_var.mCharContents, source_var._CharLength()); // Pass length to improve performance. It isn't necessary to call Contents()/Length() because they must be already up-to-date because there is no binary number to update them from (if there were, the above would have returned).  Also, caller ensured Type()==VAR_NORMAL.
}



ResultType Var::Assign(ExprTokenType &aToken)
// Returns OK or FAIL.
// Writes aToken's value into aOutputVar based on the type of the token.
// Caller must ensure that aToken.symbol is an operand (not an operator or other symbol).
// Caller must ensure that if aToken.symbol==SYM_VAR, aToken.var->Type()==VAR_NORMAL, not the clipboard or
// any built-in var.  However, this->Type() can be VAR_CLIPBOARD.
{
	switch (aToken.symbol)
	{
	case SYM_INTEGER: return Assign(aToken.value_int64); // Listed first for performance because it's Likely the most common from our callers.
	case SYM_VAR:     return Assign(*aToken.var); // Caller has ensured that var->Type()==VAR_NORMAL (it's only VAR_CLIPBOARD for certain expression lvalues, which would never be assigned here because aToken is an rvalue).
	case SYM_OBJECT:  return Assign(aToken.object);
	case SYM_FLOAT:   return Assign(aToken.value_double); // Listed last because it's probably the least common.
	}
	// Since above didn't return, it can only be SYM_STRING.
	return Assign(aToken.marker);
}



ResultType Var::AssignClipboardAll()
// Caller must ensure that "this" is a normal variable or the clipboard (though if it's the clipboard, this
// function does nothing).
{
	if (mType == VAR_ALIAS)
		// For maintainability, it seems best not to use the following method:
		//    Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// If that were done, bugs would be easy to introduce in a long function like this one
		// if your forget at use the implicit "this" by accident.  So instead, just call self.
		return mAliasFor->AssignClipboardAll();
	if (mType == VAR_CLIPBOARD) // Seems pointless to do Clipboard:=ClipboardAll, and the below isn't equipped
		return OK;              // to handle it, so make this have no effect.
	if (!g_clip.Open())
		return g_script.ScriptError(CANT_OPEN_CLIPBOARD_READ);

	// Calculate the size needed:
	// EnumClipboardFormats() retrieves all formats, including synthesized formats that don't
	// actually exist on the clipboard but are instead constructed on demand.  Unfortunately,
	// there doesn't appear to be any way to reliably determine which formats are real and
	// which are synthesized (if there were such a way, a large memory savings could be
	// realized by omitting the synthesized formats from the saved version). One thing that
	// is certain is that the "real" format(s) come first and the synthesized ones afterward.
	// However, that's not quite enough because although it is recommended that apps store
	// the primary/preferred format first, the OS does not enforce this.  For example, testing
	// shows that the apps do not have to store CF_UNICODETEXT prior to storing CF_TEXT,
	// in which case the clipboard might have inaccurate CF_TEXT as the first element and
	// more accurate/complete (non-synthesized) CF_UNICODETEXT stored as the next.
	// In spite of the above, the below seems likely to be accurate 99% or more of the time,
	// which seems worth it given the large savings of memory that are achieved, especially
	// for large quantities of text or large images. Confidence is further raised by the
	// fact that MSDN says there's no advantage/reason for an app to place multiple formats
	// onto the clipboard if those formats are available through synthesis.
	// And since CF_TEXT always(?) yields synthetic CF_OEMTEXT and CF_UNICODETEXT, and
	// probably (but less certainly) vice versa: if CF_TEXT is listed first, it might certainly
	// mean that the other two do not need to be stored.  There is some slight doubt about this
	// in a situation where an app explicitly put CF_TEXT onto the clipboard and then followed
	// it with CF_UNICODETEXT that isn't synthesized, nor does it match what would have been
	// synthesized. However, that seems extremely unlikely (it would be much more likely for
	// an app to store CF_UNICODETEXT *first* followed by custom/non-synthesized CF_TEXT, but
	// even that might be unheard of in practice).  So for now -- since there is no documentation
	// to be found about this anywhere -- it seems best to omit some of the most common
	// synthesized formats:
	// CF_TEXT is the first of three text formats to appear: Omit CF_OEMTEXT and CF_UNICODETEXT.
	//    (but not vice versa since those are less certain to be synthesized)
	//    (above avoids using four times the amount of memory that would otherwise be required)
	//    UPDATE: Only the first text format is included now, since MSDN says there is no
	//    advantage/reason to having multiple non-synthesized text formats on the clipboard.
	// CF_DIB: Always omit this if CF_DIBV5 is available (which must be present on Win2k+, at least
	// as a synthesized format, whenever CF_DIB is present?) This policy seems likely to avoid
	// the issue where CF_DIB occurs first yet CF_DIBV5 that comes later is *not* synthesized,
	// perhaps simply because the app stored DIB prior to DIBV5 by mistake (though there is
	// nothing mandatory, so maybe it's not really a mistake). Note: CF_DIBV5 supports alpha
	// channel / transparency, and perhaps other things, and it is likely that when synthesized,
	// no information of the original CF_DIB is lost. Thus, when CF_DIBV5 is placed back onto
	// the clipboard, any app that needs CF_DIB will have it synthesized back to the original
	// data (hopefully). It's debatable whether to do it that way or store whichever comes first
	// under the theory that an app would never store both formats on the clipboard since MSDN
	// says: "If the system provides an automatic type conversion for a particular clipboard format,
	// there is no advantage to placing the conversion format(s) on the clipboard."
	bool format_is_text;
	HGLOBAL hglobal;
	SIZE_T size;
	UINT format;
	VarSizeType space_needed;
	UINT dib_format_to_omit = 0, /*meta_format_to_omit = 0,*/ text_format_to_include = 0;
	// Start space_needed off at 4 to allow room for guaranteed final termination of the variable's contents.
	// The termination must be of the same size as format because a single-byte terminator would
	// be read in as a format of 0x00?????? where ?????? is an access violation beyond the buffer.
	for (space_needed = sizeof(format), format = 0; format = EnumClipboardFormats(format);)
	{
		switch (format)
		{
		case CF_BITMAP:
		case CF_ENHMETAFILE:
		case CF_DSPENHMETAFILE:
			// These formats appear to be specific handle types, not always safe to call GlobalSize() for.
			continue;
		}
		// No point in calling GetLastError() since it would never be executed because the loop's
		// condition breaks on zero return value.
		format_is_text = (format == CF_NATIVETEXT || format == CF_OEMTEXT || format == CF_OTHERTEXT);
		if ((format_is_text && text_format_to_include) // The first text format has already been found and included, so exclude all other text formats.
			|| format == dib_format_to_omit) // ... or this format was marked excluded by a prior iteration.
			continue;
		// GetClipboardData() causes Task Manager to report a (sometimes large) increase in
		// memory utilization for the script, which is odd since it persists even after the
		// clipboard is closed.  However, when something new is put onto the clipboard by the
		// the user or any app, that memory seems to get freed automatically.  Also, 
		// GetClipboardData(49356) fails in MS Visual C++ when the copied text is greater than
		// about 200 KB (but GetLastError() returns ERROR_SUCCESS).  When pasting large sections
		// of colorized text into MS Word, it can't get the colorized text either (just the plain
		// text). Because of this example, it seems likely it can fail in other places or under
		// other circumstances, perhaps by design of the app. Therefore, be tolerant of failures
		// because partially saving the clipboard seems much better than aborting the operation.
		if (hglobal = g_clip.GetClipboardDataTimeout(format))
		{
			space_needed += (VarSizeType)(sizeof(format) + sizeof(size) + GlobalSize(hglobal)); // The total amount of storage space required for this item.
			if (format_is_text) // If this is true, then text_format_to_include must be 0 since above didn't "continue".
				text_format_to_include = format;
			if (!dib_format_to_omit)
			{
				if (format == CF_DIB)
					dib_format_to_omit = CF_DIBV5;
				else if (format == CF_DIBV5)
					dib_format_to_omit = CF_DIB;
			}
			// Currently CF_ENHMETAFILE isn't supported, so no need for this section:
			//if (!meta_format_to_omit) // Checked for the same reasons as dib_format_to_omit.
			//{
			//	if (format == CF_ENHMETAFILE)
			//		meta_format_to_omit = CF_METAFILEPICT;
			//	else if (format == CF_METAFILEPICT)
			//		meta_format_to_omit = CF_ENHMETAFILE;
			//}
		}
		//else omit this format from consideration.
	}

	if (space_needed == sizeof(format)) // This works because even a single empty format requires space beyond sizeof(format) for storing its format+size.
	{
		g_clip.Close();
		return Assign(); // Nothing on the clipboard, so just make the variable blank.
	}

	// Resize the output variable, if needed:
	if (!SetCapacity(space_needed, true))
	{
		g_clip.Close();
		return FAIL; // Above should have already reported the error.
	}

	// Retrieve and store all the clipboard formats.  Because failures of GetClipboardData() are now
	// tolerated, it seems safest to recalculate the actual size (actual_space_needed) of the data
	// in case it varies from that found in the estimation phase.  This is especially necessary in
	// case GlobalLock() ever fails, since that isn't even attempted during the estimation phase.
	// Otherwise, the variable's mLength member would be set to something too high (the estimate),
	// which might cause problems elsewhere.
	LPVOID hglobal_locked;
	LPVOID binary_contents = mByteContents; // mContents vs. Contents() is okay due to the call to Assign() above.
	VarSizeType added_size, actual_space_used;
	for (actual_space_used = sizeof(format), format = 0; format = EnumClipboardFormats(format);)
	{
		switch (format)
		{
		case CF_BITMAP:
		case CF_ENHMETAFILE:
		case CF_DSPENHMETAFILE:
			// These formats appear to be specific handle types, not always safe to call GlobalSize() for.
			continue;
		}
		// No point in calling GetLastError() since it would never be executed because the loop's
		// condition breaks on zero return value.
		if ((format == CF_NATIVETEXT || format == CF_OEMTEXT || format == CF_OTHERTEXT) && format != text_format_to_include
			|| format == dib_format_to_omit /*|| format == meta_format_to_omit*/)
			continue;
		// Although the GlobalSize() documentation implies that a valid HGLOBAL should not be zero in
		// size, it does happen, at least in MS Word and for CF_BITMAP.  Therefore, in order to save
		// the clipboard as accurately as possible, also save formats whose size is zero.  Note that
		// GlobalLock() fails to work on hglobals of size zero, so don't do it for them.
		if ((hglobal = g_clip.GetClipboardDataTimeout(format)) // This and the next line rely on short-circuit boolean order.
			&& (!(size = GlobalSize(hglobal)) || (hglobal_locked = GlobalLock(hglobal)))) // Size of zero or lock succeeded: Include this format.
		{
			// Any changes made to how things are stored here should also be made to the size-estimation
			// phase so that space_needed matches what is done here:
			added_size = (VarSizeType)(sizeof(format) + sizeof(size) + size);
			actual_space_used += added_size;
			if (actual_space_used > mByteCapacity) // Tolerate incorrect estimate by omitting formats that won't fit. Note that mCapacity is the granted capacity, which might be a little larger than requested.
				actual_space_used -= added_size;
			else
			{
				*(UINT *)binary_contents = format;
				binary_contents = (char *)binary_contents + sizeof(format);
				*(SIZE_T *)binary_contents = size;
				binary_contents = (char *)binary_contents + sizeof(size);
				if (size)
				{
					memcpy(binary_contents, hglobal_locked, size);
					binary_contents = (char *)binary_contents + size;
				}
				//else hglobal_locked is not valid, so don't reference it or unlock it.
			}
			if (size)
				GlobalUnlock(hglobal); // hglobal not hglobal_locked.
		}
	}
	g_clip.Close();
	*(UINT *)binary_contents = 0; // Final termination (must be UINT, see above).
	mByteLength = actual_space_used;
	mAttrib |= VAR_ATTRIB_BINARY_CLIP; // VAR_ATTRIB_CONTENTS_OUT_OF_DATE and VAR_ATTRIB_CACHE were already removed by earlier call to Assign().

	return OK;
}



ResultType Var::AssignBinaryClip(Var &aSourceVar)
// Caller must ensure that this->Type() is VAR_NORMAL or VAR_CLIPBOARD (usually via load-time validation).
// Caller must ensure that aSourceVar->Type()==VAR_NORMAL and aSourceVar->IsBinaryClip()==true.
{
	if (mType == VAR_ALIAS)
		// For maintainability, it seems best not to use the following method:
		//    Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// If that were done, bugs would be easy to introduce in a long function like this one
		// if your forget at use the implicit "this" by accident.  So instead, just call self.
		return mAliasFor->AssignBinaryClip(aSourceVar);

	// Resolve early for maintainability.
	// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
	Var &source_var = (aSourceVar.mType == VAR_ALIAS) ? *aSourceVar.mAliasFor : aSourceVar;
	source_var.UpdateContents(); // Update mContents/mLength (probably not necessary because caller is supposed to ensure that aSourceVar->IsBinaryClip()==true).

	if (mType == VAR_NORMAL) // Copy a binary variable to another variable that isn't the clipboard.
	{
		if (this == &source_var) // i.e. source == destination.  Aliases were already resolved.
			return OK;
		if (!SetCapacity(source_var.mByteLength, true)) // source_var.mLength vs. Length() is okay (see above).
			return FAIL; // Above should have already reported the error.
		memcpy(mByteContents, source_var.mByteContents, source_var.mByteLength + sizeof(TCHAR)); // Add sizeof(TCHAR) not sizeof(format). Contents() vs. a variable for the same because mContents might have just changed due Assign() above.
		mAttrib |= VAR_ATTRIB_BINARY_CLIP; // VAR_ATTRIB_CACHE and VAR_ATTRIB_CONTENTS_OUT_OF_DATE were already removed by earlier call to Assign().
		return OK; // No need to call Close() in this case.
	}

	// SINCE ABOVE DIDN'T RETURN, A VARIABLE CONTAINING BINARY CLIPBOARD DATA IS BEING COPIED BACK ONTO THE CLIPBOARD.
	if (!g_clip.Open())
		return g_script.ScriptError(CANT_OPEN_CLIPBOARD_WRITE);
	EmptyClipboard(); // Failure is not checked for since it's probably impossible under these conditions.

	// In case the variable contents are incomplete or corrupted (such as having been read in from a
	// bad file with FileRead), prevent reading beyond the end of the variable:
	LPVOID next, binary_contents = source_var.mByteContents; // Fix for v1.0.47.05: Changed aSourceVar to source_var in this line and the next.
	LPVOID binary_contents_max = (char *)binary_contents + source_var.mByteLength; // The last accessible byte, which should be the last byte of the (UINT)0 terminator.
	HGLOBAL hglobal;
	LPVOID hglobal_locked;
	UINT format;
	SIZE_T size;

	while ((next = (char *)binary_contents + sizeof(format)) <= binary_contents_max
		&& (format = *(UINT *)binary_contents)) // Get the format.  Relies on short-circuit boolean order.
	{
		binary_contents = next;
		if ((next = (char *)binary_contents + sizeof(size)) > binary_contents_max)
			break;
		size = *(UINT *)binary_contents; // Get the size of this format's data.
		binary_contents = next;
		if ((next = (char *)binary_contents + size) > binary_contents_max)
			break;
		if (   !(hglobal = GlobalAlloc(GMEM_MOVEABLE, size))   ) // size==0 is okay.
		{
			g_clip.Close();
			return g_script.ScriptError(ERR_OUTOFMEM); // Short msg since so rare.
		}
		if (size) // i.e. Don't try to lock memory of size zero.  It won't work and it's not needed.
		{
			if (   !(hglobal_locked = GlobalLock(hglobal))   )
			{
				GlobalFree(hglobal);
				g_clip.Close();
				return g_script.ScriptError(_T("GlobalLock")); // Short msg since so rare.
			}
			memcpy(hglobal_locked, binary_contents, size);
			GlobalUnlock(hglobal);
			binary_contents = next;
		}
		//else hglobal is just an empty format, but store it for completeness/accuracy (e.g. CF_BITMAP).
		SetClipboardData(format, hglobal); // The system now owns hglobal.
	}

	return g_clip.Close();
}



ResultType Var::AssignString(LPCTSTR aBuf, VarSizeType aLength, bool aExactSize)
// Returns OK or FAIL.
// If aBuf isn't NULL, caller must ensure that aLength is either VARSIZE_MAX (which tells us that the
// entire strlen() of aBuf should be used) or an explicit length (can be zero) that the caller must
// ensure is less than or equal to the total length of aBuf (if less, only a substring is copied).
// If aBuf is NULL, the variable will be set up to handle a string of at least aLength
// in length.  In addition, if the var is the clipboard, it will be prepared for writing.
// Any existing contents of this variable will be destroyed regardless of whether aBuf is NULL.
// Note that aBuf's memory can safely overlap with that of this->Contents() because in that case the
// new length of the contents will always be less than or equal to the old length, and thus no
// reallocation/expansion is needed (such an expansion would free the source before it could be
// written to the destination).  This is because callers pass in an aBuf that is either:
// 1) Between this->Contents() and its terminator.
// 2) Equal to this->Contents() but with aLength passed in as shorter than this->Length().
//
// Caller can omit both params to set a var to be empty-string, but in that case, if the variable
// is of large capacity, its memory will not be freed.  This is by design because it allows the
// caller to exploit its knowledge of whether the var's large capacity is likely to be needed
// again in the near future, thus reducing the expected amount of memory fragmentation.
// To explicitly free the memory, use Assign("").
{
	if (mType == VAR_ALIAS)
		// For maintainability, it seems best not to use the following method:
		//    Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// If that were done, bugs would be easy to introduce in a long function like this one
		// if your forget at use the implicit "this" by accident.  So instead, just call self.
		return mAliasFor->AssignString(aBuf, aLength, aExactSize);

	bool do_assign = true;        // Set defaults.
	bool free_it_if_large = true; //
	if (!aBuf)
		if (aLength == VARSIZE_MAX) // Caller omitted this param too, so it wants to assign empty string.
		{
			free_it_if_large = false;
			aLength = 0; // aBuf is set to "" further below.
		}
		else // Caller gave a NULL buffer to signal us to ensure the var is at least aLength in capacity.
			do_assign = false;
	else // Caller provided a non-NULL buffer.
		if (aLength == VARSIZE_MAX) // Caller wants us to determine its length.
			aLength = (mCharContents == aBuf) ? CharLength() : (VarSizeType)_tcslen(aBuf); // v1.0.45: Added optimization check: (mContents == aBuf).
		//else leave aLength as the caller-specified value in case it's explicitly shorter than the apparent length.
	if (!aBuf)
		aBuf = _T("");  // From here on, make sure it's the empty string for all uses (read-only empty string vs. sEmptyString seems more appropriate in this case).

	size_t space_needed = aLength + 1; // +1 for the zero terminator.
	size_t space_needed_in_bytes = space_needed * sizeof(TCHAR);

	if (mType == VAR_CLIPBOARD)
	{
		if (do_assign)
			// Just return the result of this.  Note: The clipboard var's attributes,
			// such as mLength, are not maintained because it's a variable whose
			// contents usually aren't under our control.
			return g_clip.Set(aBuf, aLength);
		else
			// We open it for write now, because some caller's don't call
			// this function to write to the contents of the var, they
			// do it themselves.  Note: Below call will have displayed
			// any error that occurred:
			return g_clip.PrepareForWrite(space_needed) ? OK : FAIL;
	}

	// Since above didn't return, this variable isn't the clipboard.

	if (space_needed < 2) // Variable is being assigned the empty string (or a deref that resolves to it).
	{
		Free(free_it_if_large ? VAR_FREE_IF_LARGE : VAR_NEVER_FREE); // This also makes the variable blank and removes VAR_ATTRIB_OFTEN_REMOVED.
		return OK;
	}

	if (mAttrib & VAR_ATTRIB_IS_OBJECT) // This attrib will be removed below.
		mObject->Release(); // But no need to set mObject = NULL.

	// The below is done regardless of whether the section that follows it fails and returns early because
	// it's the correct thing to do in all cases.
	// For simplicity, this is done unconditionally even though it should be needed only
	// when do_assign is true. It's the caller's responsibility to turn on the binary-clip
	// attribute (if appropriate) by calling Var::Close() with the right option.
	mAttrib &= ~(VAR_ATTRIB_OFTEN_REMOVED | VAR_ATTRIB_IS_OBJECT);
	// HOWEVER, other things like making mLength 0 and mContents blank are not done here for performance
	// reasons (it seems too rare that early return/failure will occur below, since it's only due to
	// out-of-memory... and even if it does happen, there are probably no consequences to leaving the variable
	// the way it is now (rather than forcing it to be blank) since the script thread that caused the error
	// will be ended.

	if (space_needed_in_bytes > mByteCapacity)
	{
		size_t new_size; // Use a new name, rather than overloading space_needed, for maintainability.
		char *new_mem;

		switch (mHowAllocated)
		{
		case ALLOC_NONE:
		case ALLOC_SIMPLE:
			if (space_needed_in_bytes <= _TSIZE(MAX_ALLOC_SIMPLE))
			{
				// v1.0.31: Conserve memory within large arrays by allowing elements of length 3 or 7, for such
				// things as the storage of boolean values, or the storage of short numbers (it's best to use
				// multiples of 4 for size due to byte alignment in SimpleHeap; e.g. lengths of 3 and 7).
				// Because the above checked that space_needed > mCapacity, the capacity will increase but
				// never decrease in this section, which prevent a memory leak by only ever wasting a maximum
				// of 4+8+MAX_ALLOC_SIMPLE for each variable (and then only in the worst case -- in the average
				// case, it saves memory by avoiding the overhead incurred for each separate malloc'd block).
				if (space_needed_in_bytes <= _TSIZE(4)) // Even for aExactSize, it seems best to prevent variables from having only a zero terminator in them because that would usually waste 3 bytes due to byte alignment in SimpleHeap.
					new_size = _TSIZE(4); // v1.0.45: Increased from 2 to 4 to exploit byte alignment in SimpleHeap.
				else if (aExactSize) // Allows VarSetCapacity() to make more flexible use of SimpleHeap.
					new_size = space_needed_in_bytes;
				else
				{
					if (space_needed_in_bytes <= _TSIZE(8))
						new_size = _TSIZE(8); // v1.0.45: Increased from 7 to 8 to exploit 32-bit alignment in SimpleHeap.
					else // space_needed <= MAX_ALLOC_SIMPLE
						new_size = _TSIZE(MAX_ALLOC_SIMPLE);
				}
				// In the case of mHowAllocated==ALLOC_SIMPLE, the following will allocate another block
				// from SimpleHeap even though the var already had one. This is by design because it can
				// happen only a limited number of times per variable. See comments further above for details.
				if (   !(new_mem = (char *) SimpleHeap::Malloc(new_size))   )
					return FAIL; // It already displayed the error. Leave all var members unchanged so that they're consistent with each other. Don't bother making the var blank and its length zero for reasons described higher above.
				mHowAllocated = ALLOC_SIMPLE;  // In case it was previously ALLOC_NONE. This step must be done only after the alloc succeeded.
				break;
			}
			// ** ELSE DON'T BREAK, JUST FALL THROUGH TO THE NEXT CASE. **
			// **
		case ALLOC_MALLOC: // Can also reach here by falling through from above.
			// This case can happen even if space_needed is less than MAX_ALLOC_SIMPLE
			// because once a var becomes ALLOC_MALLOC, it should never change to
			// one of the other alloc modes.  See comments higher above for explanation.
			new_size = space_needed_in_bytes; // Below relies on this being initialized unconditionally.
			if (!aExactSize)
			{
				// Allow a little room for future expansion to cut down on the number of
				// free's and malloc's we expect to have to do in the future for this var:
				if (new_size < _TSIZE(16)) // v1.0.45.03: Added this new size to prevent all local variables in a recursive
					new_size = _TSIZE(16); // function from having a minimum size of MAX_PATH.  16 seems like a good size because it holds nearly any number.  It seems counterproductive to go too small because each malloc, no matter how small, could have around 40 bytes of overhead.
				else if (new_size < _TSIZE(MAX_PATH))
					new_size = _TSIZE(MAX_PATH);  // An amount that will fit all standard filenames seems good.
				else if (new_size < _TSIZE(160 * 1024)) // MAX_PATH to 160 KB or less -> 10% extra.
					new_size = (size_t)(new_size * 1.1);
				else if (new_size < _TSIZE(1600 * 1024))  // 160 to 1600 KB -> 16 KB extra
					new_size += _TSIZE(16 * 1024);
				else if (new_size < _TSIZE(6400 * 1024)) // 1600 to 6400 KB -> 1% extra
					new_size = (size_t)(new_size * 1.01);
				else  // 6400 KB or more: Cap the extra margin at some reasonable compromise of speed vs. mem usage: 64 KB
					new_size += _TSIZE(64 * 1024);
			}
			//else space_needed was already verified higher above to be within bounds.

			// In case the old memory area is large, free it before allocating the new one.  This reduces
			// the peak memory load on the system and reduces the chance of an actual out-of-memory error.
			bool memory_was_freed;
			if (memory_was_freed = (mHowAllocated == ALLOC_MALLOC && mByteCapacity)) // Verified correct: 1) Both are checked because it might have fallen through from case ALLOC_SIMPLE; 2) mCapacity indicates for certain whether mContents contains the empty string.
				free(mByteContents); // The other members are left temporarily out-of-sync for performance (they're resync'd only if an error occurs).
			//else mContents contains a "" or it points to memory on SimpleHeap, so don't attempt to free it.

			if (   (ptrdiff_t)new_size < 0 || !(new_mem = (char *)malloc(new_size))   ) // v1.0.44.10: Added a sanity limit of 2 GB so that small negatives like VarSetCapacity(Var, -2) [and perhaps other callers of this function] don't crash.
			{
				if (memory_was_freed) // Resync members to reflect the fact that it was freed (it's done this way for performance).
				{
					mByteCapacity = 0;             // Invariant: Anyone setting mCapacity to 0 must also set
					mCharContents = sEmptyString;  // mContents to the empty string.
					mByteLength = 0;               // mAttrib was already updated higher above.
				}
				// IMPORTANT: else it's the empty string (a constant) or it points to memory on SimpleHeap,
				// so don't change mContents/Capacity (that would cause a memory leak for reasons described elsewhere).
				// Also, don't bother making the variable blank and its length zero.  Just leave its contents
				// untouched due to the rarity of out-of-memory and the fact that the script thread will be terminated
				// anyway, so in most cases won't care what the contents are.
				return g_script.ScriptError(ERR_OUTOFMEM); // since an error is most likely to occur at runtime.
			}

			// Below is necessary because it might have fallen through from case ALLOC_SIMPLE.
			// This step must be done only after the alloc succeeded (because otherwise, want to keep it
			// set to ALLOC_SIMPLE (fall-through), if that's what it was).
			mHowAllocated = ALLOC_MALLOC;
			break;
		} // switch()

		// Since above didn't return, the alloc succeeded.  Because that's true, all the members (except those
		// set in their sections above) are updated together so that they stay consistent with each other:
		mByteContents = new_mem;
		mByteCapacity = (VarSizeType)new_size;
	} // if (space_needed > mCapacity)

	if (do_assign)
	{
		// Above has ensured that space_needed is either strlen(aBuf)-1 or the length of some
		// substring within aBuf starting at aBuf.  However, aBuf might overlap mContents or
		// even be the same memory address (due to something like GlobalVar := YieldGlobalVar(),
		// in which case ACT_ASSIGNEXPR calls us to assign GlobalVar to GlobalVar).
		if (mCharContents != aBuf)
		{
			// Don't use strlcpy() or such because:
			// 1) Caller might have specified that only part of aBuf's total length should be copied.
			// 2) mContents and aBuf might overlap (see above comment), in which case strcpy()'s result
			//    is undefined, but memmove() is guaranteed to work (and performs about the same).
			tmemmove(mCharContents, aBuf, aLength); // Some callers such as RegEx routines might rely on this copying binary zeroes over rather than stopping at the first binary zero.
		}
		//else nothing needs to be done since source and target are identical.  Some callers probably rely on
		// this optimization.
		mCharContents[aLength] = '\0'; // v1.0.45: This is now done unconditionally in case caller wants to shorten a variable's existing contents (no known callers do this, but it helps robustness).
	}
	else // Caller only wanted the variable resized as a preparation for something it will do later.
	{
		// Init for greater robustness/safety (the ongoing conflict between robustness/redundancy and performance).
		// This has been in effect for so long that some callers probably rely on it.
		*mCharContents = '\0'; // If it's sEmptyVar, that's okay too because it's writable.
		// We've done everything except the actual assignment.  Let the caller handle that.
		// Also, the length will be set below to the expected length in case the caller
		// doesn't override this.
		// Below: Already verified that the length value will fit into VarSizeType.
	}

	// Writing to union is safe because above already ensured that "this" isn't an alias.
	mByteLength = aLength * sizeof(TCHAR); // aLength was verified accurate higher above.
	return OK;
}



VarSizeType Var::Get(LPTSTR aBuf)
// Returns the length of this var's contents.  In addition, if aBuf isn't NULL, it will copy the contents into aBuf.
{
	// Aliases: VAR_ALIAS is checked and handled further down than in most other functions.
	//
	// For v1.0.25, don't do the following because in some cases the existing contents of aBuf will not
	// be altered.  Instead, it will be set to blank as needed further below.
	//if (aBuf) *aBuf = '\0';  // Init early to get it out of the way, in case of early return.
	VarSizeType length;

	switch(mType)
	{
	case VAR_NORMAL: // Listed first for performance.
		UpdateContents();  // Update mContents and mLength, if necessary.
		length = _CharLength();
		if (!aBuf)
			return length;
		else // Caller provider buffer, so if mLength is zero, just make aBuf empty now and return early (for performance).
			if (!mByteLength)
			{
				MaybeWarnUninitialized();
				*aBuf = '\0';
				return 0;
			}
			//else continue on below.
		if (mByteLength < 100000)
		{
			// Copy the var contents into aBuf.  Although a little bit slower than CopyMemory() for large
			// variables (say, over 100K), this loop seems much faster for small ones, which is the typical
			// case.  Also of note is that this code section is the main bottleneck for scripts that manipulate
			// large variables.
			for (LPTSTR cp = mCharContents; *cp; *aBuf++ = *cp++); // UpdateContents() was already called higher above to update mContents.
			*aBuf = '\0';
		}
		else
		{
			CopyMemory(aBuf, mByteContents, mByteLength); // Faster for large vars, but large vars aren't typical.
			aBuf[length] = '\0'; // This is done as a step separate from above in case mLength is inaccurate (e.g. due to script's improper use of DllCall).
		}
		return length;

	case VAR_ALIAS:
		// For maintainability, it seems best not to use the following method:
		//    Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// If that were done, bugs would be easy to introduce in a long function like this one
		// if your forget at use the implicit "this" by accident.  So instead, just call self.
		return mAliasFor->Get(aBuf);

	// Built-in vars with volatile contents:
	case VAR_CLIPBOARD:
	{
		length = (VarSizeType)g_clip.Get(aBuf); // It will also copy into aBuf if it's non-NULL.
		if (length == CLIPBOARD_FAILURE)
		{
			// Above already displayed the error, so just return.
			// If we were called only to determine the size, the
			// next call to g_clip.Get() will not put anything into
			// aBuf (it will either fail again, or it will return
			// a length of zero due to the clipboard not already
			// being open & ready), so there's no danger of future
			// buffer overflows as a result of returning zero here.
			// Also, due to this function's return type, there's
			// no easy way to terminate the current hotkey
			// subroutine (or entire script) due to this error.
			// However, due to the fact that multiple attempts
			// are made to open the clipboard, failure should
			// be extremely rare.  And the user will be notified
			// with a MsgBox anyway, during which the subroutine
			// will be suspended:
			length = 0;
		}
		if (aBuf)
			aBuf[length] = '\0'; // Might not be necessary, but kept in case it ever is.
		return length;
	}

	case VAR_CLIPBOARDALL: // There's a slight chance this case is never executed; but even if true, it should be kept for maintainability.
		// This variable is directly handled at a higher level.  As documented, any use of ClipboardAll outside of
		// the supported modes yields an empty string.
		if (aBuf)
			*aBuf = '\0';
		return 0;

	default: // v1.0.46.16: VAR_BUILTIN: Call the function associated with this variable to retrieve its contents.  This change reduced uncompressed coded size by 6 KB.
		return mBIV(aBuf, mName);
	} // switch(mType)
}



void Var::Free(int aWhenToFree, bool aExcludeAliasesAndRequireInit)
// The name "Free" is a little misleading because this function:
// ALWAYS sets the variable to be blank (except for static variables and aExcludeAliases==true).
// BUT ONLY SOMETIMES frees the memory, depending on various factors described further below.
// Caller must be aware that ALLOC_SIMPLE (due to its nature) is never freed.
// aExcludeAliasesAndRequireInit may be split into two if any caller ever wants to pass
// true for one and not the other (currently there is only one caller who passes true).
{
	// Not checked because even if it's not VAR_NORMAL, there are few if any consequences to continuing.
	//if (mType != VAR_NORMAL) // For robustness, since callers generally shouldn't call it this way.
	//	return;

	if (mType == VAR_ALIAS) // For simplicity and reduced code size, just make a recursive call to self.
	{
		if (!aExcludeAliasesAndRequireInit)
			// For maintainability, it seems best not to use the following method:
			//    Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
			// If that were done, bugs would be easy to introduce in a long function like this one
			// if your forget at use the implicit "this" by accident.  So instead, just call self.
			mAliasFor->Free(aWhenToFree);
		//else caller didn't want the target of the alias freed, so do nothing.
		return;
	}

	// Must check this one first because caller relies not only on var not being freed in this case,
	// but also on its contents not being set to an empty string:
	if (aWhenToFree == VAR_ALWAYS_FREE_BUT_EXCLUDE_STATIC && IsStatic())
		return; // This is the only case in which the variable ISN'T made blank.

	if (mAttrib & VAR_ATTRIB_IS_OBJECT) // This attrib will be removed below.
		mObject->Release(); // But no need to set mObject = NULL.

	mByteLength = 0; // Writing to union is safe because above already ensured that "this" isn't an alias.
	// Even if it isn't free'd, variable will be made blank.  So it seems proper to always remove
	// the binary_clip attribute (since it can't be used that way after it's been made blank) and
	// the uninitialized attribute (since *we* are initializing it).  Some callers may rely on us
	// removing these attributes:
	mAttrib &= ~(VAR_ATTRIB_OFTEN_REMOVED | VAR_ATTRIB_IS_OBJECT);

	if (aExcludeAliasesAndRequireInit)
		// Caller requires this var to be considered uninitialized from now on.  This attribute may
		// have been removed above, but there was no cost involved.  It might not have been set in
		// the first place, so we must add it here anyway:
		mAttrib |= VAR_ATTRIB_UNINITIALIZED;

	switch (mHowAllocated)
	{
	// Shouldn't be necessary to check the following because by definition, ALLOC_NONE
	// means mContents==sEmptyString (this policy is enforced in other places).
	//
	//case ALLOC_NONE:
	//	mContents = sEmptyString;
	//	break;

	case ALLOC_SIMPLE:
		// Don't set to sEmptyString because then we'd have a memory leak.  i.e. once a var becomes
		// ALLOC_SIMPLE, it should never become ALLOC_NONE again (though it can become ALLOC_MALLOC).
		*mCharContents = '\0';
		break;

	case ALLOC_MALLOC:
		// Setting a var whose contents are very large to be nothing or blank is currently the
		// only way to free up the memory of that var.  Shrinking it dynamically seems like it
		// might introduce too much memory fragmentation and overhead (since in many cases,
		// it would likely need to grow back to its former size in the near future).  So we
		// only free relatively large vars:
		if (mByteCapacity)
		{
			// aWhenToFree==VAR_FREE_IF_LARGE: the memory is not freed if it is a small area because
			// it might help reduce memory fragmentation and improve performance in cases where
			// the memory will soon be needed again (in which case one free+malloc is saved).
			if (   aWhenToFree < VAR_ALWAYS_FREE_LAST  // Fixed for v1.0.40.07 to prevent memory leak in recursive script-function calls.
				|| aWhenToFree == VAR_FREE_IF_LARGE && mByteCapacity > (4 * 1024)   )
			{
				free(mByteContents);
				mByteCapacity = 0;             // Invariant: Anyone setting mCapacity to 0 must also set
				mCharContents = sEmptyString;  // mContents to the empty string.
				// BUT DON'T CHANGE mHowAllocated to ALLOC_NONE (see comments further below).
			}
			else // Don't actually free it, but make it blank (callers rely on this).
				*mCharContents = '\0';
		}
		//else mCapacity==0, so mContents is already the empty string, so don't attempt to free
		// it or assign to it. It was the responsibility of whoever set mCapacity to 0 to ensure mContents
		// was set properly (which is why it's not done here).

		// But do not change mHowAllocated to be ALLOC_NONE because it would cause a
		// a memory leak in this sequence of events:
		// var1 is assigned something short enough to make it ALLOC_SIMPLE
		// var1 is assigned something large enough to require malloc()
		// var1 is set to empty string and its mem is thus free()'d by the above.
		// var1 is once again assigned something short enough to make it ALLOC_SIMPLE
		// The last step above would be a problem because the 2nd ALLOC_SIMPLE can't
		// reclaim the spot in SimpleHeap that had been in use by the first.  In other
		// words, when a var makes the transition from ALLOC_SIMPLE to ALLOC_MALLOC,
		// its ALLOC_SIMPLE memory is lost to the system until the program exits.
		// But since this loss occurs at most once per distinct variable name,
		// it's not considered a memory leak because the loss can't exceed a fixed
		// amount regardless of how long the program runs.  The reason for all of this
		// is that allocating dynamic memory is costly: it causes system memory fragmentation,
		// (especially if a var were to be malloc'd and free'd thousands of times in a loop)
		// and small-sized mallocs have a large overhead: it's been said that every block
		// of dynamic mem, even those of size 1 or 2, incurs about 40 bytes of overhead
		// (and testing of SimpleHeap supports this).
		// UPDATE: Yes it's true that a script could create an unlimited number of new variables
		// by constantly using different dynamic names when creating them.  But it's still supportable
		// that there's no memory leak because even if the worst-case series of events mentioned
		// higher above occurs for every variable, only a fixed amount of memory is ever lost per
		// variable.  So you could think of the lost memory as part of the basic overhead of creating
		// a variable (since variables are never truly destroyed, just their contents freed).
		// The odds against all of these worst-case factors occurring simultaneously in anything other
		// than a theoretical test script seem nearly astronomical.
		break;
	} // switch()
}



ResultType Var::AppendIfRoom(LPTSTR aStr, VarSizeType aLength)
// Returns OK if there's room enough to append aStr and it succeeds.
// Returns FAIL otherwise (also returns FAIL for VAR_CLIPBOARD).
// Environment variables aren't supported here; instead, aStr is appended directly onto the actual/internal
// contents of the "this" variable.
{
	// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()):
	Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
	if (var.mType != VAR_NORMAL // e.g. VAR_CLIPBOARD. Some callers do call it this way, but even if not it should be kept for maintainability.
		|| (var.mAttrib & VAR_ATTRIB_IS_OBJECT)) // It seems best for maintainability to not handle this case here.
		return FAIL; // CHECK THIS FIRST, BEFORE BELOW, BECAUSE CALLERS ALWAYS WANT IT TO BE A FAILURE.
	if (!aLength) // Consider the appending of nothing (even onto unsupported things like clipboard) to be a success.
		return OK;
	VarSizeType var_length = var.LengthIgnoreBinaryClip(); // Get the apparent length because one caller is a concat that wants consistent behavior of the .= operator regardless of whether this shortcut succeeds or not.
	VarSizeType new_length = var_length + aLength;
	if (new_length >= var._CharCapacity()) // Not enough room.
		return FAIL;
	tmemmove(var.mCharContents + var_length, aStr, aLength);  // mContents was updated via LengthIgnoreBinaryClip() above. Use memmove() vs. memcpy() in case there's any overlap between source and dest.
	var.mCharContents[new_length] = '\0'; // Terminate it as a separate step in case caller passed a length shorter than the apparent length of aStr.
	var.mByteLength = new_length * sizeof(TCHAR);
	// If this is a binary-clip variable, appending has probably "corrupted" it; so don't allow it to ever be
	// put back onto the clipboard as binary data (the routine that does that is designed to detect corruption,
	// but it might not be perfect since corruption is so rare).  Also remove the other flags that are no longer
	// appropriate:
	var.mAttrib &= ~VAR_ATTRIB_OFTEN_REMOVED; // This also removes VAR_ATTRIB_NOT_NUMERIC because appending some digits to an empty variable would make it numeric.
	return OK;
}



void Var::AcceptNewMem(LPTSTR aNewMem, VarSizeType aLength)
// Caller provides a new malloc'd memory block (currently must be non-NULL).  That block and its
// contents are directly hung onto this variable in place of its old block, which is freed (except
// in the case of VAR_CLIPBOARD, in which case the memory is copied onto the clipboard then freed).
// Caller must ensure that mType == VAR_NORMAL or VAR_CLIPBOARD.
// This function was added in v1.0.45 to aid callers in improving performance.
{
	// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
	Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
	if (var.mType == VAR_CLIPBOARD)
	{
		var.Assign(aNewMem, aLength); // Clipboard requires GlobalAlloc memory so can't directly accept aNewMem.  So just copy it the normal way.
		free(aNewMem); // Caller gave it to us to take charge of, but we have no further use for it.
	}
	else // VAR_NORMAL
	{
		var.Free(VAR_ALWAYS_FREE); // Release the variable's old memory. This also removes flags VAR_ATTRIB_OFTEN_REMOVED.
		var.mHowAllocated = ALLOC_MALLOC; // Must always be this type to avoid complications and possible memory leaks.
		var.mByteContents = (char *) aNewMem;
		var.mByteLength = aLength * sizeof(TCHAR);
		ASSERT( var.mByteLength + sizeof(TCHAR) <= _msize(aNewMem) );
		var.mByteCapacity = (VarSizeType)_msize(aNewMem); // Get actual capacity in case it's a lot bigger than aLength+1. _msize() is only about 36 bytes of code and probably a very fast call.
		// Already done by Free() above:
		//mAttrib &= ~VAR_ATTRIB_OFTEN_REMOVED; // New memory is always non-binary-clip.  A new parameter could be added to change this if it's ever needed.

		// Shrink the memory if there's a lot of wasted space because the extra capacity is seldom utilized
		// in real-world scripts.
		// A simple rule seems best because shrinking is probably fast regardless of the sizes involved,
		// plus and there's considerable rarity to ever needing capacity beyond what's in a variable
		// (concat/append is one example).  This will leave a large percentage of extra space in small variables,
		// but in those rare cases when a script needs to create thousands of such variables, there may be a
		// current or future way to shrink an existing variable to contain only its current length, such as
		// VarShrink().
		if (var.mByteCapacity - var.mByteLength > 64)
		{
			var.mByteCapacity = var.mByteLength + sizeof(TCHAR); // This will become the new capacity.
			// _expand() is only about 75 bytes of uncompressed code size and probably performs very quickly
			// when shrinking.  Also, MSDN implies that when shrinking, failure won't happen unless something
			// is terribly wrong (e.g. corrupted heap).  But for robustness it is checked anyway:
			if (   !(var.mByteContents = (char *)_expand(var.mByteContents, var.mByteCapacity))   )
			{
				var.mByteLength = 0;
				var.mByteCapacity = 0;
			}
		}
	}
}



void Var::SetLengthFromContents()
// Function added in v1.0.43.06.  It updates the mLength member to reflect the actual current length of mContents.
// Caller must ensure that Type() is VAR_NORMAL.
{
	// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
	Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
	VarSizeType capacity = var.Capacity();
	var.UpdateContents(); // Ensure mContents and mLength are up-to-date.
	if (capacity > 0)
	{
 		var.mCharContents[capacity - 1] = '\0';  // Caller wants us to ensure it's terminated, to avoid crashing strlen() below.
		var.mByteLength = ((VarSizeType)_tcslen(var.mCharContents)) * sizeof(TCHAR);
	}
	//else it has no capacity, so do nothing (it could also be a reserved/built-in variable).
}



ResultType Var::BackupFunctionVars(Func &aFunc, VarBkp *&aVarBackup, int &aVarBackupCount)
// All parameters except the first are output parameters that are set for our caller (though caller
// is responsible for having initialized aVarBackup to NULL).
// If there is nothing to backup, only the aVarBackupCount is changed (to zero).
// Returns OK or FAIL.
{
	if (   !(aVarBackupCount = aFunc.mVarCount + aFunc.mLazyVarCount)   )  // Nothing needs to be backed up.
		return OK; // Leave aVarBackup set to NULL as set by the caller.

	// NOTES ABOUT MALLOC(): Apparently, the implementation of malloc() is quite good, at least for small blocks
	// needed to back up 50 or less variables.  It nearly as fast as alloca(), at least when the system
	// isn't under load and has the memory to spare without swapping.  Therefore, the attempt to use alloca to
	// speed up recursive script-functions didn't result in enough of a speed-up (only 1 to 5%) to be worth the
	// added complexity.
	// Since Var is not a POD struct (it contains private members, a custom constructor, etc.), the VarBkp
	// POD struct is used to hold the backup because it's probably better performance than using Var's
	// constructor to create each backup array element.
	if (   !(aVarBackup = (VarBkp *)malloc(aVarBackupCount * sizeof(VarBkp)))   ) // Caller will take care of freeing it.
		return FAIL;

	int i;
	aVarBackupCount = 0;  // Init only once prior to both loops. aVarBackupCount is being "overloaded" to track the current item in aVarBackup, BUT ALSO its being updated to an actual count in case some statics are omitted from the array.

	// Note that Backup() does not make the variable empty after backing it up because that is something
	// that must be done by our caller at a later stage.
	for (i = 0; i < aFunc.mVarCount; ++i)
		if (!aFunc.mVar[i]->IsStatic()) // Don't bother backing up statics because they won't need to be restored.
			aFunc.mVar[i]->Backup(aVarBackup[aVarBackupCount++]);
	for (i = 0; i < aFunc.mLazyVarCount; ++i)
		if (!aFunc.mLazyVar[i]->IsStatic()) // Don't bother backing up statics because they won't need to be restored.
			aFunc.mLazyVar[i]->Backup(aVarBackup[aVarBackupCount++]);
	return OK;
}



void Var::Backup(VarBkp &aVarBkp)
// Caller must not call this function for static variables because it's not equipped to deal with them
// (they don't need to be backed up or restored anyway).
// This method is used rather than struct copy (=) because it's of expected higher performance than
// using the Var::constructor to make a copy of each var.  Also note that something like memcpy()
// can't be used on Var objects since they're not POD (e.g. they have a constructor and they have
// private members).
{
	aVarBkp.mVar = this; // Allows the restoration process to always know its target without searching.
	aVarBkp.mByteContents = mByteContents;
	aVarBkp.mContentsInt64 = mContentsInt64; // This also copies the other members of the union: mContentsDouble, mObject.
	aVarBkp.mByteLength = mByteLength; // Since it's a union, it might actually be backing up mAliasFor (happens at least for recursive functions that pass parameters ByRef).
	aVarBkp.mByteCapacity = mByteCapacity;
	aVarBkp.mHowAllocated = mHowAllocated; // This might be ALLOC_SIMPLE or ALLOC_NONE if backed up variable was at the lowest layer of the call stack.
	aVarBkp.mAttrib = mAttrib;
	aVarBkp.mType = mType; // Fix for v1.0.47.06: Must also back up and restore mType in case an optional ByRef parameter is omitted by one call by specified by another thread that interrupts the first thread's call.
	// Once the backup is made, Free() is not called because the whole point of the backup is to
	// preserve the original memory/contents of each variable.  Instead, clear the variable
	// completely and set it up to become ALLOC_MALLOC in case anything actually winds up using
	// the variable prior to the restoration of the backup.  In other words, ALLOC_SIMPLE and NONE
	// retained (if present) because that would cause a memory leak when multiple layers are all
	// allowed to use ALLOC_SIMPLE yet none are ever able to free it (the bottommost layer is
	// allowed to use ALLOC_SIMPLE because that's a fixed/constant amount of memory gets freed
	// when the program exits).
	// Now reset this variable (caller has ensured it's non-static) to create a "new layer" for it, keeping
	// its backup intact but allowing this variable (or formal parameter) to be given a new value in the future:
	mByteCapacity = 0;             // Invariant: Anyone setting mCapacity to 0 must also set...
	mCharContents = sEmptyString;  // ...mContents to the empty string.
	if (mType != VAR_ALIAS) // Fix for v1.0.42.07: Don't reset mLength if the other member of the union is in effect.
		mByteLength = 0;        // Otherwise, functions that recursively pass ByRef parameters can crash because mType stays as VAR_ALIAS.
	mHowAllocated = ALLOC_MALLOC; // Never NONE because that would permit SIMPLE. See comments higher above.
	mAttrib = VAR_ATTRIB_UNINITIALIZED; // The function's new recursion layer should consider this var uninitialized, even if it was initialized by the previous layer.
}



void Var::Restore(VarBkp &aVarBkp)
{
	mByteContents = aVarBkp.mByteContents;
	mContentsInt64 = aVarBkp.mContentsInt64; // This also copies the other members of the union: mContentsDouble, mObject.
	mByteLength = aVarBkp.mByteLength; // Since it's a union, it might actually be restoring mAliasFor, which is desired.
	mByteCapacity = aVarBkp.mByteCapacity;
	mHowAllocated = aVarBkp.mHowAllocated; // This might be ALLOC_SIMPLE or ALLOC_NONE if backed-up variable was at the lowest layer of the call stack.
	mAttrib = aVarBkp.mAttrib;
	mType = aVarBkp.mType;
}



void Var::FreeAndRestoreFunctionVars(Func &aFunc, VarBkp *&aVarBackup, int &aVarBackupCount)
{
	int i;
	for (i = 0; i < aFunc.mVarCount; ++i)
		aFunc.mVar[i]->Free(VAR_ALWAYS_FREE_BUT_EXCLUDE_STATIC, true); // Pass "true" to exclude aliases, since their targets should not be freed (they don't belong to this function). Also resets the "uninitialized" attribute.
	for (i = 0; i < aFunc.mLazyVarCount; ++i)
		aFunc.mLazyVar[i]->Free(VAR_ALWAYS_FREE_BUT_EXCLUDE_STATIC, true);

	// The freeing (above) MUST be done prior to the restore-from-backup below (otherwise there would be
	// a memory leak).  Static variables are never backed up and thus do not exist in the aVarBackup array.
	// This is because by definition, the contents of statics are not freed or altered by the calling procedure
	// (regardless how recursive or multi-threaded the function is).
	if (aVarBackup) // This is the indicator that a backup was made; thus a restore is also needed.
	{
		for (i = 0; i < aVarBackupCount; ++i) // Static variables were never backed up so they won't be in this array. See comments above.
		{
			VarBkp &bkp = aVarBackup[i];
			bkp.mVar->Restore(bkp);
		}
		free(aVarBackup);
		aVarBackup = NULL; // Some callers want this reset; it's an indicator of whether the next function call in this expression (if any) will have a backup.
	}
}



ResultType Var::ValidateName(LPCTSTR aName, int aDisplayError)
// Returns OK or FAIL.
{
	if (!*aName) return FAIL;
	// Seems best to disallow variables that start with numbers for purity, to allow
	// something like 1e3 to be scientific notation, and possibly other reasons.
	if (*aName >= '0' && *aName <= '9')
	{
		if (aDisplayError)
		{
			TCHAR msg[512];
			sntprintf(msg, _countof(msg), _T("This %s name starts with a number, which is not allowed:\n\"%-1.300s\"")
				, aDisplayError == DISPLAY_VAR_ERROR ? _T("variable") : _T("function")
				, aName);
			return g_script.ScriptError(msg);
		}
		else
			return FAIL;
	}
	if (*find_identifier_end(aName) != '\0')
	{
		if (aDisplayError)
		{
			TCHAR msg[512];
			sntprintf(msg, _countof(msg), _T("The following %s name contains an illegal character:\n\"%-1.300s\"")
				, aDisplayError == DISPLAY_VAR_ERROR ? _T("variable") : _T("function")
				, aName);
			return g_script.ScriptError(msg);
		}
		return FAIL;
	}
	// Otherwise:
	return OK;
}

ResultType Var::AssignStringFromCodePage(LPCSTR aBuf, int aLength, UINT aCodePage)
{
#ifndef UNICODE
	// Not done since some callers have a more effective optimization in place:
	//if (aCodePage == CP_ACP || aCodePage == GetACP())
		// Avoid unnecessary conversion (ACP -> UTF16 -> ACP).
		//return AssignString(aBuf, aLength, true, false);
	// Convert from specified codepage to UTF-16,
	CStringWCharFromChar wide_buf(aBuf, aLength, aCodePage);
	// then back to the active codepage:
	return AssignStringToCodePage(wide_buf, wide_buf.GetLength(), CP_ACP);
#else
	int iLen = MultiByteToWideChar(aCodePage, 0, aBuf, aLength, NULL, 0);
	if (iLen > 0) {
		if (!AssignString(NULL, iLen, true))
			return FAIL;
		LPWSTR aContents = Contents(TRUE, TRUE);
		iLen = MultiByteToWideChar(aCodePage, 0, aBuf, aLength, (LPWSTR) aContents, iLen);
		aContents[iLen] = 0;
		if (!iLen)
			return FAIL;
		SetCharLength(aContents[iLen - 1] ? iLen : iLen - 1);
	}
	else
		Assign(); // Return value is ambiguous in this case: may be zero-length input or an error.  For simplicity, return OK.
	return OK;
#endif
}

ResultType Var::AssignStringToCodePage(LPCWSTR aBuf, int aLength, UINT aCodePage, DWORD aFlags, char aDefChar)
{
	char *pDefChar;
	if (aCodePage == CP_UTF8 || aCodePage == CP_UTF7) {
		pDefChar = NULL;
		aFlags = 0;
	}
	else
		pDefChar = &aDefChar;
	int iLen = WideCharToMultiByte(aCodePage, aFlags, aBuf, aLength, NULL, 0, pDefChar, NULL);
	if (iLen > 0) {
		if (!SetCapacity(iLen, true))
			return FAIL;
		LPSTR aContents = (LPSTR) Contents(TRUE, TRUE);
		iLen = WideCharToMultiByte(aCodePage, aFlags, aBuf, aLength, aContents, iLen, pDefChar, NULL);
		aContents[iLen] = 0;
		if (!iLen)
			return FAIL;
#ifndef UNICODE
		SetCharLength(aContents[iLen - 1] ? iLen : iLen - 1);
#endif
	}
	else
		Assign();
	return OK;
}

__forceinline void Var::MaybeWarnUninitialized()
{
	if (IsUninitializedNormalVar())
	{
		// The following should not be possible; if it is, there's a bug and we want to know about it:
		//if (mByteLength != 0)
		//	MarkInitialized();	// "self-correct" if we catch a var that has normal content but wasn't marked initialized
		//else
			g_script.WarnUninitializedVar(this);
	}
}
