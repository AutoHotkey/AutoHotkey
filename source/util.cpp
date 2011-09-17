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
#include <olectl.h> // for OleLoadPicture()
#include <gdiplus.h> // Used by LoadPicture().
#include "util.h"
#include "globaldata.h"


int GetYDay(int aMon, int aDay, bool aIsLeapYear)
// Returns a number between 1 and 366.
// Caller must verify that aMon is a number between 1 and 12, and aDay is a number between 1 and 31.
{
	--aMon;  // Convert to zero-based.
	if (aIsLeapYear)
	{
		int leap_offset[12] = {0, 31, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335};
		return leap_offset[aMon] + aDay;
	}
	else
	{
		int normal_offset[12] = {0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334};
		return normal_offset[aMon] + aDay;
	}
}



int GetISOWeekNumber(LPTSTR aBuf, int aYear, int aYDay, int aWDay)
// Caller must ensure that aBuf is of size 7 or greater, that aYear is a valid year (e.g. 2005),
// that aYDay is between 1 and 366, and that aWDay is between 0 and 6 (day of the week).
// Produces the week number in YYYYNN format, e.g. 200501.
// Note that year is also returned because it isn't necessarily the same as aTime's calendar year.
// Based on Linux glibc source code (GPL).
{
	--aYDay;  // Convert to zero based.
	#define ISO_WEEK_START_WDAY 1 // Monday
	#define ISO_WEEK1_WDAY 4      // Thursday
	#define ISO_WEEK_DAYS(yday, wday) (yday - (yday - wday + ISO_WEEK1_WDAY + ((366 / 7 + 2) * 7)) % 7 \
		+ ISO_WEEK1_WDAY - ISO_WEEK_START_WDAY);

	int year = aYear;
	int days = ISO_WEEK_DAYS(aYDay, aWDay);

	if (days < 0) // This ISO week belongs to the previous year.
	{
		--year;
		days = ISO_WEEK_DAYS(aYDay + (365 + IS_LEAP_YEAR(year)), aWDay);
	}
	else
	{
		int d = ISO_WEEK_DAYS(aYDay - (365 + IS_LEAP_YEAR(year)), aWDay);
		if (0 <= d) // This ISO week belongs to the next year.
		{
			++year;
			days = d;
		}
	}

	// Use sntprintf() for safety; that is, in case year contains a value longer than 4 digits.
	// This also adds the leading zeros in front of year and week number, if needed.
	return sntprintf(aBuf, 7, _T("%04d%02d"), year, (days / 7) + 1); // Return the length of the string produced.
}



ResultType YYYYMMDDToFileTime(LPTSTR aYYYYMMDD, FILETIME &aFileTime)
{
	SYSTEMTIME st;
	YYYYMMDDToSystemTime(aYYYYMMDD, st, false);  // "false" because it's validated below.
	// This will return failure if aYYYYMMDD contained any invalid elements, such as an
	// explicit zero for the day of the month.  It also reports failure if st.wYear is
	// less than 1601, which for simplicity is enforced globally throughout the program
	// since none of the Windows API calls seem to support earlier years.
	return SystemTimeToFileTime(&st, &aFileTime) ? OK : FAIL; // The st.wDayOfWeek member is ignored.
}



DWORD YYYYMMDDToSystemTime2(LPTSTR aYYYYMMDD, SYSTEMTIME *aSystemTime)
// Calls YYYYMMDDToSystemTime() to fill up to two elements of the aSystemTime array.
// Returns a GDTR bitwise combination to indicate which of the two elements, or both, are valid.
// Caller must ensure that aYYYYMMDD is a modifiable string since it's temporarily altered and restored here.
{
	DWORD gdtr = 0;
	if (!*aYYYYMMDD)
		return gdtr;
	if (*aYYYYMMDD != '-') // Since first char isn't a dash, there is a minimum present.
	{
		LPTSTR cp;
		if (cp = _tcschr(aYYYYMMDD + 1, '-'))
			*cp = '\0'; // Temporarily terminate in case only the leading part of the YYYYMMDD format is present.  Otherwise, the dash and other chars would be considered invalid fields.
		if (YYYYMMDDToSystemTime(aYYYYMMDD, aSystemTime[0], true)) // Date string is valid.
			gdtr |= GDTR_MIN; // Indicate that minimum is present.
		if (cp)
		{
			*cp = '-'; // Undo the temp. termination.
			aYYYYMMDD = cp + 1; // Set it to the maximum's position for use below.
		}
		else // No dash, so there is no maximum.  Indicate this by making aYYYYMMDD empty.
			aYYYYMMDD = _T("");
	}
	else // *aYYYYMMDD=='-', so only the maximum is present; thus there will be no minimum.
		++aYYYYMMDD; // Skip over the dash to set it to the maximum's position.
	if (*aYYYYMMDD) // There is a maximum.
	{
		if (YYYYMMDDToSystemTime(aYYYYMMDD, aSystemTime[1], true)) // Date string is valid.
			gdtr |= GDTR_MAX; // Indicate that maximum is present.
	}
	return gdtr;
}



ResultType YYYYMMDDToSystemTime(LPTSTR aYYYYMMDD, SYSTEMTIME &aSystemTime, bool aDoValidate)
// Although aYYYYMMDD need not be terminated at the end of the YYYYMMDDHH24MISS string (as long as
// the string's capacity is at least 14), it should be terminated if only the leading part
// of the YYYYMMDDHH24MISS format is present.
// Caller must ensure that aYYYYMMDD is non-NULL.  If aDoValidate is false, OK is always
// returned and aSystemTime might contain invalid elements.  Otherwise, FAIL will be returned
// if the date and time contains any invalid elements, or if the year is less than 1601
// (Windows generally does not support earlier years).
{
	// sscanf() is avoided because it adds 2 KB to the compressed EXE size.
	TCHAR temp[16];
	size_t length = _tcslen(aYYYYMMDD); // Use this rather than incrementing the pointer in case there are ever partial fields such as 20051 vs. 200501.

	tcslcpy(temp, aYYYYMMDD, 5);
	aSystemTime.wYear = _ttoi(temp);

	if (length > 4) // It has a month component.
	{
		tcslcpy(temp, aYYYYMMDD + 4, 3);
		aSystemTime.wMonth = _ttoi(temp);  // Unlike "struct tm", SYSTEMTIME uses 1 for January, not 0.
		// v1.0.48: Changed not to provide a default when month number is out-of-range.
		// This allows callers like "if var is time" to properly detect badly-formatted dates.
	}
	else
		aSystemTime.wMonth = 1;

	if (length > 6) // It has a day-of-month component.
	{
		tcslcpy(temp, aYYYYMMDD + 6, 3);
		aSystemTime.wDay = _ttoi(temp);
	}
	else
		aSystemTime.wDay = 1;

	if (length > 8) // It has an hour component.
	{
		tcslcpy(temp, aYYYYMMDD + 8, 3);
		aSystemTime.wHour = _ttoi(temp);
	}
	else
		aSystemTime.wHour = 0;   // Midnight.

	if (length > 10) // It has a minutes component.
	{
		tcslcpy(temp, aYYYYMMDD + 10, 3);
		aSystemTime.wMinute = _ttoi(temp);
	}
	else
		aSystemTime.wMinute = 0;

	if (length > 12) // It has a seconds component.
	{
		tcslcpy(temp, aYYYYMMDD + 12, 3);
		aSystemTime.wSecond = _ttoi(temp);
	}
	else
		aSystemTime.wSecond = 0;

	aSystemTime.wMilliseconds = 0;  // Always set to zero in this case.

	// v1.0.46.07: Unlike the other date/time components, which are validated further below by the call to
	// SystemTimeToFileTime(), "month" must be validated in advance because it's used to access an array
	// in the day-of-week code further below.
	// v1.0.48.04: To fix FormatTime and possibly other things, don't return FAIL when month is out-of-range unless
	// aDoValidate==false, and not even then (for maintainability) because a section further below handles it.
	if (aSystemTime.wMonth < 1 || aSystemTime.wMonth > 12) // Month must be validated prior to accessing the array further below.
		// Set an in-range default, which caller is likely to ignore if it passed true for aDoValidate
		// because the validation further below will also detect the out-of-range month and return FAIL.
		aSystemTime.wDayOfWeek = 1;
	else // Month is in-range, which is necessary for the method below to work safely.
	{
		// Day-of-week code by Tomohiko Sakamoto:
		static int t[] = {0, 3, 2, 5, 0, 3, 5, 1, 4, 6, 2, 4};
		int y = aSystemTime.wYear;
		y -= aSystemTime.wMonth < 3;
		aSystemTime.wDayOfWeek = (y + y/4 - y/100 + y/400 + t[aSystemTime.wMonth-1] + aSystemTime.wDay) % 7;
	}

	if (aDoValidate)
	{
		FILETIME ft;
		// This will return failure if aYYYYMMDD contained any invalid elements, such as an
		// explicit zero for the day of the month.  It also reports failure if st.wYear is
		// less than 1601, which for simplicity is enforced globally throughout the program
		// since none of the Windows API calls seem to support earlier years.
		return SystemTimeToFileTime(&aSystemTime, &ft) ? OK : FAIL;
		// Above: The st.wDayOfWeek member is ignored by the above (but might be used by our caller), but
		// that's okay because it shouldn't need validation.
	}
	return OK;
}



LPTSTR FileTimeToYYYYMMDD(LPTSTR aBuf, FILETIME &aTime, bool aConvertToLocalTime)
// Returns aBuf.
{
	FILETIME ft;
	if (aConvertToLocalTime)
		FileTimeToLocalFileTime(&aTime, &ft); // MSDN says that target cannot be the same var as source.
	else
		memcpy(&ft, &aTime, sizeof(FILETIME));  // memcpy() might be less code size that a struct assignment, ft = aTime.
	SYSTEMTIME st;
	if (FileTimeToSystemTime(&ft, &st))
		return SystemTimeToYYYYMMDD(aBuf, st);
	*aBuf = '\0';
	return aBuf;
}



LPTSTR SystemTimeToYYYYMMDD(LPTSTR aBuf, SYSTEMTIME &aTime)
// Returns aBuf.
// Remember not to offer a "aConvertToLocalTime" option, because calling SystemTimeToTzSpecificLocalTime()
// on Win9x apparently results in an invalid time because the function is implemented only as a stub on
// those OSes.
{
	_stprintf(aBuf, _T("%04d%02d%02d") _T("%02d%02d%02d")
		, aTime.wYear, aTime.wMonth, aTime.wDay
		, aTime.wHour, aTime.wMinute, aTime.wSecond);
	return aBuf;
}



__int64 YYYYMMDDSecondsUntil(LPTSTR aYYYYMMDDStart, LPTSTR aYYYYMMDDEnd, bool &aFailed)
// Returns the number of seconds from aYYYYMMDDStart until aYYYYMMDDEnd.
// If aYYYYMMDDStart is blank, the current time will be used in its place.
{
	aFailed = true;  // Set default for output parameter, in case of early return.
	if (!aYYYYMMDDStart || !aYYYYMMDDEnd) return 0;

	FILETIME ftStart, ftEnd, ftNowUTC;

	if (*aYYYYMMDDStart)
	{
		if (!YYYYMMDDToFileTime(aYYYYMMDDStart, ftStart))
			return 0;
	}
	else // Use the current time in its place.
	{
		GetSystemTimeAsFileTime(&ftNowUTC);
		FileTimeToLocalFileTime(&ftNowUTC, &ftStart);  // Convert UTC to local time.
	}
	if (*aYYYYMMDDEnd)
	{
		if (!YYYYMMDDToFileTime(aYYYYMMDDEnd, ftEnd))
			return 0;
	}
	else // Use the current time in its place.
	{
		GetSystemTimeAsFileTime(&ftNowUTC);
		FileTimeToLocalFileTime(&ftNowUTC, &ftEnd);  // Convert UTC to local time.
	}
	aFailed = false;  // Indicate success.
	return FileTimeSecondsUntil(&ftStart, &ftEnd);
}



__int64 FileTimeSecondsUntil(FILETIME *pftStart, FILETIME *pftEnd)
// Returns the number of seconds from pftStart until pftEnd.
{
	if (!pftStart || !pftEnd) return 0;

	// The calculation is done this way for compilers that don't support 64-bit math operations (not sure which):
	// Note: This must be LARGE vs. ULARGE because we want the calculation to be signed for cases where
	// pftStart is greater than pftEnd:
	ULARGE_INTEGER uiStart, uiEnd;
	uiStart.LowPart = pftStart->dwLowDateTime;
	uiStart.HighPart = pftStart->dwHighDateTime;
	uiEnd.LowPart = pftEnd->dwLowDateTime;
	uiEnd.HighPart = pftEnd->dwHighDateTime;
	// Must do at least the inner cast to avoid losing negative results:
	return (__int64)((__int64)(uiEnd.QuadPart - uiStart.QuadPart) / 10000000); // Convert from tenths-of-microsecond.
}



SymbolType IsNumeric(LPCTSTR aBuf, BOOL aAllowNegative, BOOL aAllowAllWhitespace
	, BOOL aAllowFloat, BOOL aAllowImpure)  // BOOL vs. bool might squeeze a little more performance out of this frequently-called function.
// String can contain whitespace.
// If aBuf doesn't contain something purely numeric, PURE_NOT_NUMERIC is returned.  The same happens if
// aBuf contains a float but aAllowFloat is false.  (The aAllowFloat parameter isn't strictly necessary
// because the caller could just check whether the return value is/isn't PURE_FLOAT to get the same effect.
// However, supporting aAllowFloat seems to greatly improve maintainability because it saves many callers
// from having to compare the return value to PURE_INTEGER [they can just interpret the return value as BOOL].
// It also improves readability due to the "Is" part of the function name.  So it seems worth keeping.)
// Otherwise, PURE_INTEGER or PURE_FLOAT is returned.
// If aAllowAllWhitespace==true and the string is blank or all whitespace, PURE_INTEGER is returned.
// Obsolete comment: Making this non-inline reduces the size of the compressed EXE by only 2K.  Since this
// function is called so often, it seems preferable to keep it inline for performance.
{
	aBuf = omit_leading_whitespace(aBuf); // i.e. caller doesn't have to have ltrimmed, only rtrimmed.
	if (!*aBuf) // The string is empty or consists entirely of whitespace.
		return aAllowAllWhitespace ? PURE_INTEGER : PURE_NOT_NUMERIC;

	if (*aBuf == '-')
	{
		if (aAllowNegative)
			++aBuf;
		else
			return PURE_NOT_NUMERIC;
	}
	else if (*aBuf == '+')
		++aBuf;

	// Relies on short circuit boolean order to prevent reading beyond the end of the string:
	BOOL is_hex = IS_HEX(aBuf); // BOOL vs. bool might squeeze a little more performance out this frequently-called function.
	if (is_hex)
		aBuf += 2;  // Skip over the 0x prefix.

	// Set defaults:
	BOOL has_decimal_point = false;
	BOOL has_exponent = false;
	BOOL has_at_least_one_digit = false; // i.e. a string consisting of only "+", "-" or "." is not considered numeric.
	int c; // int vs. char might squeeze a little more performance out of it (it does reduce code size by 5 bytes). Probably must stay signed vs. unsigned for some of the uses below.

	for (;; ++aBuf)
	{
		c = *aBuf;
		if (IS_SPACE_OR_TAB(c))
		{
			if (*omit_leading_whitespace(aBuf)) // But that space or tab is followed by something other than whitespace.
				if (!aAllowImpure) // e.g. "123 456" is not a valid pure number.
					return PURE_NOT_NUMERIC;
				// else fall through to the bottom logic.
			// else since just whitespace at the end, the number qualifies as pure, so fall through to the bottom
			// logic (it would already have returned in the loop if it was impure)
			break;
		}
		if (!c) // End of string was encountered.
			break; // The number qualifies as pure, so fall through to the logic at the bottom. (It would already have returned elsewhere in the loop if the number is impure).
		if (c == '.')
		{
			if (!aAllowFloat || has_decimal_point || is_hex) // If aAllowFloat==false, a decimal point at the very end of the number is considered non-numeric even if aAllowImpure==true.  Some callers might rely on this.
				// i.e. if aBuf contains 2 decimal points, it can't be a valid number.
				// Note that decimal points are allowed in hexadecimal strings, e.g. 0xFF.EE.
				// But since that format doesn't seem to be supported by VC++'s atof() and probably
				// related functions, and since it's extremely rare, it seems best not to support it.
				return PURE_NOT_NUMERIC;
			else
				has_decimal_point = true;
		}
		else
		{
			if (is_hex ? !_istxdigit(c) : (c < '0' || c > '9')) // And since we're here, it's not '.' either.
			{
				if (aAllowImpure) // Since aStr starts with a number (as verified above), it is considered a number.
				{
					if (has_at_least_one_digit)
						return has_decimal_point ? PURE_FLOAT : PURE_INTEGER;
					else // i.e. the strings "." and "-" are not considered to be numeric by themselves.
						return PURE_NOT_NUMERIC;
				}
				else
				{
					if (ctoupper(c) != 'E' // v1.0.46.11: Support scientific notation in floating point numbers.
						|| !has_at_least_one_digit // But it must have at least one digit to the left of the 'E'. Some callers rely on this check.
						|| has_exponent
						|| !aAllowFloat)
						return PURE_NOT_NUMERIC;
					if (aBuf[1] == '-' || aBuf[1] == '+') // The optional sign is present on the exponent.
						++aBuf; // Omit it from further consideration so that the outer loop doesn't see it as an extra/illegal sign.
					if (aBuf[1] < '0' || aBuf[1] > '9')
						// Even if it is an 'e', ensure what follows it is a valid exponent.  Some callers rely
						// on this check, such as ones that expect "0.6e" to be non-numeric (for "SetFormat Float") 
						return PURE_NOT_NUMERIC;
					has_exponent = true;
					has_decimal_point = true; // For simplicity, since a decimal point after the exponent isn't valid.
				}
			}
			else // This character is a valid digit or hex-digit.
				has_at_least_one_digit = true;
		}
	} // for()

	if (has_at_least_one_digit)
		return has_decimal_point ? PURE_FLOAT : PURE_INTEGER;
	else
		return PURE_NOT_NUMERIC; // i.e. the strings "+" "-" and "." are not numeric by themselves.
}



void strlcpy(LPSTR aDst, LPCSTR aSrc, size_t aDstSize) // Non-inline because it benches slightly faster that way.
// Caller must ensure that aDstSize is greater than 0.
// Caller must ensure that the entire capacity of aDst is writable, EVEN WHEN it knows that aSrc is much shorter
// than the aDstSize.  This is because the call to strncpy (which is used for its superior performance) zero-fills
// any unused portion of aDst.
// Description:
// Same as strncpy() but guarantees null-termination of aDst upon return.
// No more than aDstSize - 1 characters will be copied from aSrc into aDst
// (leaving room for the zero terminator, which is always inserted).
// This function is defined in some Unices but is not standard.  But unlike
// other versions, this one uses void for return value for reduced code size
// (since it's called in so many places).
{
	// Disabled for performance and reduced code size:
	//if (!aDst || !aSrc || !aDstSize) return aDstSize;  // aDstSize must not be zero due to the below method.
	// It might be worthwhile to have a custom char-copying-loop here someday so that number of characters
	// actually copied (not including the zero terminator) can be returned to callers who want it.
	--aDstSize; // Convert from size to length (caller has ensured that aDstSize > 0).
	strncpy(aDst, aSrc, aDstSize); // NOTE: In spite of its zero-filling, strncpy() benchmarks considerably faster than a custom loop, probably because it uses 32-bit memory operations vs. 8-bit.
	aDst[aDstSize] = '\0';
}



void wcslcpy(LPWSTR aDst, LPCWSTR aSrc, size_t aDstSize)
{
	--aDstSize;
	wcsncpy(aDst, aSrc, aDstSize);
	aDst[aDstSize] = '\0';
}



int sntprintf(LPTSTR aBuf, int aBufSize, LPCTSTR aFormat, ...)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// aBuf will always be terminated here except when aBufSize is <= zero (in which case the caller should
// already have terminated it).  If aBufSize is greater than zero but not large enough to hold the
// entire result, as much of the result as possible is copied and the return value is aBufSize - 1.
// Returns the exact number of characters written, not including the zero terminator.  A negative
// number is never returned, even if aBufSize is <= zero (which means there isn't even enough space left
// to write a zero terminator), under the assumption that the caller has already terminated the string
// and thus prefers to have 0 rather than -1 returned in such cases.
// MSDN says (about _snprintf(), and testing shows that it applies to _vsnprintf() too): "This function
// does not guarantee NULL termination, so ensure it is followed by sz[size - 1] = 0".
{
	// The following should probably never be changed without a full suite of tests to ensure the
	// change doesn't cause the finicky _vsnprintf() to break something.
	if (aBufSize < 1 || !aBuf || !aFormat) return 0; // It's called from so many places that the extra checks seem warranted.
	va_list ap;
	va_start(ap, aFormat);
	// Must use _vsnprintf() not _snprintf() because of the way va_list is handled:
	int result = _vsntprintf(aBuf, aBufSize, aFormat, ap); // "returns the number of characters written, not including the terminating null character, or a negative value if an output error occurs"
	aBuf[aBufSize - 1] = '\0'; // Confirmed through testing: Must terminate at this exact spot because _vsnprintf() doesn't always do it.
	// Fix for v1.0.34: If result==aBufSize, must reduce result by 1 to return an accurate result to the
	// caller.  In other words, if the line above turned the last character into a terminator, one less character
	// is now present in aBuf.
	if (result == aBufSize)
		--result;
	return result > -1 ? result : aBufSize - 1; // Never return a negative value.  See comment under function definition, above.
}



int sntprintfcat(LPTSTR aBuf, int aBufSize, LPCTSTR aFormat, ...)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// aBuf will always be terminated here except when the amount of space left in the buffer is zero or less.
// (in which case the caller should already have terminated it).  If aBufSize is greater than zero but not
// large enough to hold the entire result, as much of the result as possible is copied and the return value
// is space_remaining - 1.
// The caller must have ensured that aBuf and aFormat are non-NULL and that aBuf contains a valid string
// (i.e. that it is null-terminated somewhere within the limits of aBufSize).
// Returns the exact number of characters written, not including the zero terminator.  A negative
// number is never returned, even if aBufSize is <= zero (which means there isn't even enough space left
// to write a zero terminator), under the assumption that the caller has already terminated the string
// and thus prefers to have 0 rather than -1 returned in such cases.
{
	// The following should probably never be changed without a full suite of tests to ensure the
	// change doesn't cause the finicky _vsnprintf() to break something.
	size_t length = _tcslen(aBuf);
	int space_remaining = (int)(aBufSize - length); // Must cast to int to avoid loss of negative values.
	if (space_remaining < 1) // Can't even terminate it (no room) so just indicate that no characters were copied.
		return 0;
	aBuf += length;  // aBuf is now the spot where the new text will be written.
	va_list ap;
	va_start(ap, aFormat);
	// Must use vsnprintf() not snprintf() because of the way va_list is handled:
	int result = _vsntprintf(aBuf, (size_t)space_remaining, aFormat, ap); // "returns the number of characters written, not including the terminating null character, or a negative value if an output error occurs"
	aBuf[space_remaining - 1] = '\0'; // Confirmed through testing: Must terminate at this exact spot because _vsnprintf() doesn't always do it.
	return result > -1 ? result : space_remaining - 1; // Never return a negative value.  See comment under function definition, above.
}



// Not currently used by anything, so commented out to possibly reduce code size:
//int tcslcmp(LPTSTR aBuf1, LPTSTR aBuf2, UINT aLength1, UINT aLength2)
//// Case sensitive version.  See strlicmp() comments below.
//{
//	if (!aBuf1 || !aBuf2) return 0;
//	if (aLength1 == UINT_MAX) aLength1 = (UINT)_tcslen(aBuf1);
//	if (aLength2 == UINT_MAX) aLength2 = (UINT)_tcslen(aBuf2);
//	UINT least_length = aLength1 < aLength2 ? aLength1 : aLength2;
//	int diff;
//	for (UINT i = 0; i < least_length; ++i)
//		if (   diff = (int)((UCHAR)aBuf1[i] - (UCHAR)aBuf2[i])   ) // Use unsigned chars like strcmp().
//			return diff;
//	return (int)(aLength1 - aLength2);
//}



int tcslicmp(LPTSTR aBuf1, LPTSTR aBuf2, size_t aLength1, size_t aLength2)
// Similar to strnicmp but considers each aBuf to be a string of length aLength if aLength was
// specified.  In other words, unlike strnicmp() which would consider strnicmp("ab", "abc", 2)
// [example verified correct] to be a match, this function would consider them to be
// a mismatch.  Another way of looking at it: aBuf1 and aBuf2 will be directly
// compared to one another as though they were actually of length aLength1 and
// aLength2, respectively and then passed to stricmp() (not strnicmp) as those
// shorter strings.  This behavior is useful for cases where you don't want
// to have to bother with temporarily terminating a string so you can compare
// only a substring to something else.  The return value meaning is the
// same as strnicmp().  If either aLength param is UINT_MAX (via the default
// parameters or via explicit call), it will be assumed that the entire
// length of the respective aBuf will be used.
{
	if (!aBuf1 || !aBuf2) return 0;
	if (aLength1 == -1) aLength1 = _tcslen(aBuf1);
	if (aLength2 == -1) aLength2 = _tcslen(aBuf2);
	size_t least_length = aLength1 < aLength2 ? aLength1 : aLength2;
	int diff;
	for (size_t i = 0; i < least_length; ++i)
		if (   diff = (int)(ctoupper(aBuf1[i]) - ctoupper(aBuf2[i]))   )
			return diff;
	// Since the above didn't return, the strings are equal if they're the same length.
	// Otherwise, the longer one is considered greater than the shorter one since the
	// longer one's next character is by definition something non-zero.  I'm not completely
	// sure that this is the same policy followed by ANSI strcmp():
	return (int)(aLength1 - aLength2);
}



LPTSTR tcsrstr(LPTSTR aStr, size_t aStr_length, LPCTSTR aPattern, StringCaseSenseType aStringCaseSense, int aOccurrence)
// Returns NULL if not found, otherwise the address of the found string.
// This could probably use a faster algorithm someday.  For now it seems adequate because
// scripts rarely use it and when they do, it's usually on short haystack strings (such as
// to find the last period in a filename).
{
	if (aOccurrence < 1)
		return NULL;
	if (!*aPattern)
		// The empty string is found in every string, and since we're searching from the right, return
		// the position of the zero terminator to indicate the situation:
		return aStr + aStr_length;

	size_t aPattern_length = _tcslen(aPattern);
	TCHAR aPattern_last_char = aPattern[aPattern_length - 1];
	TCHAR aPattern_last_char_lower = (aStringCaseSense == SCS_INSENSITIVE_LOCALE)
		? (TCHAR)ltolower(aPattern_last_char)
		: ctolower(aPattern_last_char);

	int occurrence = 0;
	LPCTSTR match_starting_pos = aStr + aStr_length - 1;

	// Keep finding matches from the right until the Nth occurrence (specified by the caller) is found.
	for (;;)
	{
		if (match_starting_pos < aStr)
			return NULL;  // No further matches are possible.
		// Find (from the right) the first occurrence of aPattern's last char:
		LPCTSTR last_char_match;
		for (last_char_match = match_starting_pos; last_char_match >= aStr; --last_char_match)
		{
			if (aStringCaseSense == SCS_INSENSITIVE) // The most common mode is listed first for performance.
			{
				if (ctolower(*last_char_match) == aPattern_last_char_lower)
					break;
			}
			else if (aStringCaseSense == SCS_INSENSITIVE_LOCALE)
			{
				if (ltolower(*last_char_match) == aPattern_last_char_lower)
					break;
			}
			else // Case sensitive.
			{
				if (*last_char_match == aPattern_last_char)
					break;
			}
		}

		if (last_char_match < aStr) // No further matches are possible.
			return NULL;

		// Now that aPattern's last character has been found in aStr, ensure the rest of aPattern
		// exists in aStr to the left of last_char_match:
		LPCTSTR full_match, cp;
		bool found;
		for (found = false, cp = aPattern + aPattern_length - 2, full_match = last_char_match - 1;; --cp, --full_match)
		{
			if (cp < aPattern) // The complete pattern has been found at the position in full_match + 1.
			{
				++full_match; // Adjust for the prior iteration's decrement.
				if (++occurrence == aOccurrence)
					return (LPTSTR) full_match;
				found = true;
				break;
			}
			if (full_match < aStr) // Only after checking the above is this checked.
				break;

			if (aStringCaseSense == SCS_INSENSITIVE) // The most common mode is listed first for performance.
			{
				if (ctolower(*full_match) != ctolower(*cp))
					break;
			}
			else if (aStringCaseSense == SCS_INSENSITIVE_LOCALE)
			{
				if (ltolower(*full_match) != ltolower(*cp))
					break;
			}
			else // Case sensitive.
			{
				if (*full_match != *cp)
					break;
			}
		} // for() innermost
		if (found) // Although the above found a match, it wasn't the right one, so resume searching.
			match_starting_pos = full_match - 1;
		else // the pattern broke down, so resume searching at THIS position.
			match_starting_pos = last_char_match - 1;  // Don't go back by more than 1.
	} // while() find next match
}



LPTSTR tcscasestr(LPCTSTR phaystack, LPCTSTR pneedle)
	// To make this work with MS Visual C++, this version uses tolower/toupper() in place of
	// _tolower/_toupper(), since apparently in GNU C, the underscore macros are identical
	// to the non-underscore versions; but in MS the underscore ones do an unconditional
	// conversion (mangling non-alphabetic characters such as the zero terminator).  MSDN:
	// tolower: Converts c to lowercase if appropriate
	// _tolower: Converts c to lowercase

	// Return the offset of one string within another.
	// Copyright (C) 1994,1996,1997,1998,1999,2000 Free Software Foundation, Inc.
	// This file is part of the GNU C Library.

	// The GNU C Library is free software; you can redistribute it and/or
	// modify it under the terms of the GNU Lesser General Public
	// License as published by the Free Software Foundation; either
	// version 2.1 of the License, or (at your option) any later version.

	// The GNU C Library is distributed in the hope that it will be useful,
	// but WITHOUT ANY WARRANTY; without even the implied warranty of
	// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
	// Lesser General Public License for more details.

	// You should have received a copy of the GNU Lesser General Public
	// License along with the GNU C Library; if not, write to the Free
	// Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA
	// 02111-1307 USA.

	// My personal strstr() implementation that beats most other algorithms.
	// Until someone tells me otherwise, I assume that this is the
	// fastest implementation of strstr() in C.
	// I deliberately chose not to comment it.  You should have at least
	// as much fun trying to understand it, as I had to write it :-).
	// Stephen R. van den Berg, berg@pool.informatik.rwth-aachen.de

	// Faster looping by precalculating bl, bu, cl, cu before looping.
	// 2004 Apr 08	Jose Da Silva, digital@joescat@com
{
	register const TBYTE *haystack, *needle;
	register unsigned bl, bu, cl, cu;
	
	haystack = (const TBYTE *) phaystack;
	needle = (const TBYTE *) pneedle;

	// Since ctolower returns TCHAR (which is signed in ANSI builds), typecast to
	// TBYTE first to promote characters \x80-\xFF to unsigned 32-bit correctly:
	bl = (TBYTE)ctolower(*needle);
	if (bl != '\0')
	{
		// Scan haystack until the first character of needle is found:
		bu = (TBYTE)ctoupper(bl);
		haystack--;				/* possible ANSI violation */
		do
		{
			cl = *++haystack;
			if (cl == '\0')
				goto ret0;
		}
		while ((cl != bl) && (cl != bu));

		// See if the rest of needle is a one-for-one match with this part of haystack:
		cl = (TBYTE)ctolower(*++needle);
		if (cl == '\0')  // Since needle consists of only one character, it is already a match as found above.
			goto foundneedle;
		cu = (TBYTE)ctoupper(cl);
		++needle;
		goto jin;
		
		for (;;)
		{
			register unsigned a;
			register const TBYTE *rhaystack, *rneedle;
			do
			{
				a = *++haystack;
				if (a == '\0')
					goto ret0;
				if ((a == bl) || (a == bu))
					break;
				a = *++haystack;
				if (a == '\0')
					goto ret0;
shloop:
				;
			}
			while ((a != bl) && (a != bu));

jin:
			a = *++haystack;
			if (a == '\0')  // Remaining part of haystack is shorter than needle.  No match.
				goto ret0;

			if ((a != cl) && (a != cu)) // This promising candidate is not a complete match.
				goto shloop;            // Start looking for another match on the first char of needle.
			
			rhaystack = haystack-- + 1;
			rneedle = needle;
			a = (TBYTE)ctolower(*rneedle);
			
			if ((TBYTE)ctolower(*rhaystack) == (int) a)
			do
			{
				if (a == '\0')
					goto foundneedle;
				++rhaystack;
				a = (TBYTE)ctolower(*++needle);
				if ((TBYTE)ctolower(*rhaystack) != (int) a)
					break;
				if (a == '\0')
					goto foundneedle;
				++rhaystack;
				a = (TBYTE)ctolower(*++needle);
			}
			while ((TBYTE)ctolower(*rhaystack) == (int) a);
			
			needle = rneedle;		/* took the register-poor approach */
			
			if (a == '\0')
				break;
		} // for(;;)
	} // if (bl != '\0')
foundneedle:
	return (LPTSTR) haystack;
ret0:
	return 0;
}



LPTSTR lstrcasestr(LPCTSTR phaystack, LPCTSTR pneedle)
// This is the locale-obeying variant of strcasestr.  It uses CharUpper/Lower in place of toupper/lower,
// which sees chars like ä as the same as Ä (depending on code page/locale).  This function is about
// 1 to 8 times slower than strcasestr() depending on factors such as how many partial matches for needle
// are in haystack.
// License: GNU GPL
// Copyright (C) 1994,1996,1997,1998,1999,2000 Free Software Foundation, Inc.
// See strcasestr() for more comments.
{
	register const TBYTE *haystack, *needle;
	register unsigned bl, bu, cl, cu;
	
	haystack = (const TBYTE *) phaystack;
	needle = (const TBYTE *) pneedle;

	bl = (UINT)ltolower(*needle);
	if (bl != 0)
	{
		// Scan haystack until the first character of needle is found:
		bu = (UINT)(size_t)ltoupper(bl);
		haystack--;				/* possible ANSI violation */
		do
		{
			cl = *++haystack;
			if (cl == '\0')
				goto ret0;
		}
		while ((cl != bl) && (cl != bu));

		// See if the rest of needle is a one-for-one match with this part of haystack:
		cl = (UINT)ltolower(*++needle);
		if (cl == '\0')  // Since needle consists of only one character, it is already a match as found above.
			goto foundneedle;
		cu = (UINT)ltoupper(cl);
		++needle;
		goto jin;
		
		for (;;)
		{
			register unsigned a;
			register const TBYTE *rhaystack, *rneedle;
			do
			{
				a = *++haystack;
				if (a == '\0')
					goto ret0;
				if ((a == bl) || (a == bu))
					break;
				a = *++haystack;
				if (a == '\0')
					goto ret0;
shloop:
				;
			}
			while ((a != bl) && (a != bu));

jin:
			a = *++haystack;
			if (a == '\0')  // Remaining part of haystack is shorter than needle.  No match.
				goto ret0;

			if ((a != cl) && (a != cu)) // This promising candidate is not a complete match.
				goto shloop;            // Start looking for another match on the first char of needle.
			
			rhaystack = haystack-- + 1;
			rneedle = needle;
			a = (UINT)ltolower(*rneedle);
			
			if ((UINT)ltolower(*rhaystack) == a)
			do
			{
				if (a == '\0')
					goto foundneedle;
				++rhaystack;
				a = (UINT)ltolower(*++needle);
				if ((UINT)ltolower(*rhaystack) != a)
					break;
				if (a == '\0')
					goto foundneedle;
				++rhaystack;
				a = (UINT)ltolower(*++needle);
			}
			while ((UINT)ltolower(*rhaystack) == a);
			
			needle = rneedle;		/* took the register-poor approach */
			
			if (a == '\0')
				break;
		} // for(;;)
	} // if (bl != '\0')
foundneedle:
	return (LPTSTR) haystack;
ret0:
	return 0;
}



UINT StrReplace(LPTSTR aHaystack, LPTSTR aOld, LPTSTR aNew, StringCaseSenseType aStringCaseSense
	, UINT aLimit, size_t aSizeLimit, LPTSTR *aDest, size_t *aHaystackLength)
// Replaces all (or aLimit) occurrences of aOld with aNew in aHaystack.
// On success, it returns the number of replacements done (0 if none).  On failure (out of memory), it returns 0
// (and if aDest isn't NULL, it also sets *aDest to NULL on failure).
//
// PARAMETERS:
// - aLimit: Specify UINT_MAX to have no restriction on the number of replacements.  Otherwise, specify a number >=0.
// - aSizeLimit: Specify -1 to assume that aHaystack has enough capacity for any mode #1 replacement. Otherwise,
//   specify the size limit (in either mode 1 or 2), but it must be >= length of aHaystack (simplifies the code).
// - aDest: If NULL, the function will operate in mode #1.  Otherwise, it uses mode #2 (see further below).
// - aHaystackLength: If it isn't NULL, *aHaystackLength must be the length of aHaystack.  HOWEVER, *aHaystackLength
//   is then changed here to be the length of the result string so that caller can use it to improve performance.
//
// MODE 1 (when aDest==NULL): aHaystack is used as both the source and the destination (sometimes temporary memory
// is used for performance, but it's freed afterward and so transparent to the caller).
// When it passes in -1 for aSizeLimit (the default), caller must ensure that aHaystack has enough capacity to hold
// the new/replaced result.  When non-NULL, aSizeLimit will be enforced by limiting the number of replacements to
// the available memory (i.e. any remaining replacements are simply not done and that part of haystack is unaltered).
//
// MODE 2 (when aDest!=NULL): If zero replacements are needed, we set *aDest to be aHaystack to indicate that no
// new memory was allocated.  Otherwise, we store in *aDest the address of the new memory that holds the result.
// - The caller is responsible for any new memory in *aDest (freeing it, etc.)
// - The initial value of *aDest doesn't matter.
// - The contents of aHaystack isn't altered, not even if aOld_length==aNew_length (some callers rely on this).
//
// v1.0.45: This function was heavily revised to improve performance and flexibility.  It has also made
// two other/related StrReplace() functions obsolete.  Also, the code has been simplified to avoid doing
// a first pass through haystack to find out exactly how many replacements there are because that step
// nearly doubles the time required for the entire operation (in most cases).  Its benefit was mainly in
// memory savings and avoidance of any reallocs since the initial alloc was always exactly right; however,
// testing shows that one or two reallocs are generally much quicker than doing the size-calculation phase
// because extra alloc'ing & memcpy'ing is much faster than an extra search through haystack for all the matches.
// Furthermore, the new approach minimizes reallocs by using smart prediction.  Furthermore, the caller shrinks
// the result memory via _expand() to avoid having too much extra/overhang.  These optimizations seem to make
// the new approach better than the old one in every way, but especially performance.
{
	#define REPLACEMENT_MODE2 aDest  // For readability.

	// THINGS TO SET NOW IN CASE OF EARLY RETURN OR GOTO:
	// Set up the input/output lengths:
	size_t haystack_length = aHaystackLength ? *aHaystackLength : _tcslen(aHaystack); // For performance, use caller's length if it was provided.
	size_t length_temp; // Just a placeholder/memory location used by the alias below.
	size_t &result_length = aHaystackLength ? *aHaystackLength : length_temp; // Make an alias for convenience and maintainability (if this is an output parameter for our caller, this step takes care that in advance).
	// Set up the output buffer:
	LPTSTR result_temp; // In mode #1, holds temporary memory that is freed before we return.
	LPTSTR &result = aDest ? *aDest : result_temp; // Make an alias for convenience and maintainability (if aDest is non-NULL, it's an output parameter for our caller, and this step takes care that in advance).
	result = NULL;     // It's allocated only upon first use to avoid a potentially massive allocation that might
	result_length = 0; // be wasted and cause swapping (not to mention that we'll have better ability to estimate the correct total size after the first replacement is discovered).
	size_t result_size = 0;
	// Variables used by both replacement methods.
	LPTSTR src, match_pos;
	// END OF INITIAL SETUP.

	// From now on, result_length and result should be kept up-to-date because they may have been set up
	// as output parameters above.

	if (!(*aHaystack && *aOld))
	{
		// Nothing to do if aHaystack is blank. If aOld is blank, that is not supported because it would be an
		// infinite loop. This policy is now largely due to backward compatibility because some other policy
		// may have been better.
		result = aHaystack; // Return unaltered string to caller in its output parameter (result is an alias for *aDest).
		result_length = haystack_length; // This is an alias for an output parameter, so update it for caller.
		return 0; // Report "no replacements".
	}

	size_t aOld_length = _tcslen(aOld);
	size_t aNew_length = _tcslen(aNew);
	int length_delta = (int)(aNew_length - aOld_length); // Cast to int to avoid loss of unsigned. A negative delta means the replacement substring is smaller than what it's replacing.

	if (aSizeLimit != -1) // Caller provided a size *restriction*, so if necessary reduce aLimit to stay within bounds.  Compare directly to -1 due to unsigned.
	{
		int extra_room = (int)(aSizeLimit-1 - haystack_length); // Cast to int to preserve negatives.
		if (extra_room < 0) // Caller isn't supposed to call it this way.  To avoid having to complicate the
			aLimit = 0;     // calculations in the else-if below, allow no replacements in this case.
		else if (length_delta > 0) // New-str is bigger than old-str.  This is checked to avoid going negative or dividing by 0 below. A positive delta means length of new/replacement string is greater than that of what it's replacing.
		{
			UINT upper_limit = (UINT)(extra_room / length_delta);
			if (aLimit > upper_limit)
				aLimit = upper_limit;
		}
		//else length_delta <= 0, so there no overflow should be possible.  Leave aLimit as-is.
	}

	if (!REPLACEMENT_MODE2) // Mode #1
	{
		if (!length_delta // old_len==new_len, so use in-place method because it's just as fast in this case but it avoids the extra memory allocation.
			|| haystack_length < 5000) // ...or the in-place method will likely be faster, and an earlier stage has ensured there's no risk of overflow.
			goto in_place_method; // "Goto" to avoid annoying indentation and long IF-blocks.
		//else continue on because the extra-memory method will usually perform better than the in-place method.
		// The extra-memory method is much faster than the in-place method when many replacements are needed because
		// it avoids a memmove() to shift the remainder of the buffer up against the area of memory that
		// will be replaced (which for the in-place method is done for every replacement).  The savings
		// can be enormous if aSource is very large, assuming the system can allocate the memory without swapping.
	}
	// Otherwise:
	// Since above didn't jump to the in place method, either the extra-memory method is preferred or this is mode #2.
	// Never use the in-place method for mode #2 because caller always wants a separate memory area used (for its
	// purposes, the extra-memory method is probably just as fast or faster than in-place method).

	// Below uses a temp var. because realloc() returns NULL on failure but leaves original block allocated.
	// Note that if it's given a NULL pointer, realloc() does a malloc() instead.
	LPTSTR realloc_temp;
	#define STRREPLACE_REALLOC(size) \
	{\
		result_size = size;\
		if (   !(realloc_temp = trealloc(result, result_size))   )\
			goto out_of_mem;\
		result = realloc_temp;\
	}

	// Other variables used by the replacement loop:
	size_t haystack_portion_length, new_result_length;
	UINT replacement_count;

	// Perform the replacement:
	for (replacement_count = 0, src = aHaystack
		; aLimit && (match_pos = tcsstr2(src, aOld, aStringCaseSense));) // Relies on short-circuit boolean order.
	{
		++replacement_count;
		--aLimit;
		haystack_portion_length = match_pos - src; // The length of the haystack section between the end of the previous match and the start of the current one.

		// Using the required length calculated below, expand/realloc "result" if necessary.
		new_result_length = result_length + haystack_portion_length + aNew_length;
		if (new_result_length >= result_size) // Uses >= to allow room for terminator.
			STRREPLACE_REALLOC(PredictReplacementSize(length_delta, replacement_count, aLimit, (int)haystack_length
				, (int)new_result_length, (int)(match_pos - aHaystack))); // This will return if an alloc error occurs.

		// Now that we know "result" has enough capacity, put the new text into it.  The first step
		// is to copy over the part of haystack that appears before the match.
		if (haystack_portion_length)
		{
			tmemcpy(result + result_length, src, haystack_portion_length);
			result_length += haystack_portion_length;
		}
		// Now append the replacement string in place of the old string.
		if (aNew_length)
		{
			tmemcpy(result + result_length, aNew, aNew_length);
			result_length += aNew_length;
		}
		//else omit it altogether; i.e. replace every aOld with the empty string.

		// Set up src to be the position where the next iteration will start searching.  For consistency with
		// the in-place method, overlapping matches are not detected.  For example, the replacement
		// of all occurrences of ".." with ". ." in "..." would produce ". ..", not ". . .":
		src = match_pos + aOld_length; // This has two purposes: 1) Since match_pos is about to be altered by strstr, src serves as a placeholder for use by the next iteration; 2) it's also used further below.
	}

	if (!replacement_count) // No replacements were done, so optimize by keeping the original (avoids a malloc+memcpy).
	{
		// The following steps are appropriate for both mode #1 and #2 (for simplicity and maintainability,
		// they're all done unconditionally even though mode #1 might not require them all).
		result = aHaystack; // Return unaltered string to caller in its output parameter (result is an alias for *aDest).
		result_length = haystack_length; // This is an alias for an output parameter, so update it for caller.
		return replacement_count;
		// Since no memory was allocated, there's never anything to free.
	}
	// (Below relies only above having returned when no replacements because it assumes result!=NULL from now on.)

	// Otherwise, copy the remaining characters after the last replacement (if any) (fixed for v1.0.25.11).
	if (haystack_portion_length = haystack_length - (src - aHaystack)) // This is the remaining part of haystack that need to be copied over as-is.
	{
		new_result_length = result_length + haystack_portion_length;
		if (new_result_length >= result_size) // Uses >= to allow room for terminator.
			STRREPLACE_REALLOC(new_result_length + 1); // This will return if an alloc error occurs.
		tmemcpy(result + result_length, src, haystack_portion_length); // memcpy() usually benches a little faster than strcpy().
		result_length = new_result_length; // Remember that result_length is actually an output for our caller, so even if for no other reason, it must be kept accurate for that.
	}
	result[result_length] = '\0'; // Must terminate it unconditionally because other sections usually don't do it.

	if (!REPLACEMENT_MODE2) // Mode #1.
	{
		// Since caller didn't provide destination memory, copy the result from our temporary memory (that was used
		// for performance) back into the caller's original buf (which has already been confirmed to be large enough).
		tmemcpy(aHaystack, result, result_length + 1); // Include the zero terminator.
		free(result); // Free the temp. mem that was used for performance.
	}
	return replacement_count;  // The output parameters have already been populated properly above.

out_of_mem: // This can only happen with the extra-memory method above (which due to its nature can't fall back to the in-place method).
	if (result)
	{
		free(result); // Must be freed in mode #1.  In mode #2, it's probably a non-terminated string (not to mention being an incomplete result), so if it ever isn't freed, it should be terminated.
		result = NULL; // Indicate failure by setting output param for our caller (this also indicates that the memory was freed).
	}
	result_length = 0; // Output parameter for caller, though upon failure it shouldn't matter (just for robustness).
	return 0;

in_place_method:
	// This method is available only to mode #1.  It should help performance for short strings such as those from
	// ExpandExpression().
	// This in-place method is used when the extra-memory method wouldn't be faster enough to be worth its cost
	// for the particular strings involved here.
	//
	// Older comment:
	// The below doesn't quite work when doing a simple replacement such as ".." with ". .".
	// In the above example, "..." would be changed to ". .." rather than ". . ." as it should be.
	// Therefore, use a less efficient, but more accurate method instead.  UPDATE: But this method
	// can cause an infinite loop if the new string is a superset of the old string, so don't use
	// it after all.
	//for ( ; ptr = StrReplace(aHaystack, aOld, aNew, aStringCaseSense); ); // Note that this very different from the below.

	for (replacement_count = 0, src = aHaystack
		; aLimit && (match_pos = tcsstr2(src, aOld, aStringCaseSense)) // Relies on short-circuit boolean order.
		; --aLimit, ++replacement_count)
	{
		src = match_pos + aNew_length;  // The next search should start at this position when all is adjusted below.
		if (length_delta) // This check can greatly improve performance if old and new strings happen to be same length.
		{
			// Since new string can't fit exactly in place of old string, adjust the target area to
			// accept exactly the right length so that the rest of the string stays unaltered:
			tmemmove(src, match_pos + aOld_length
				, (haystack_length - (match_pos - aHaystack)) - aOld_length + 1); // +1 to include zero terminator.
			// Above: Calculating length vs. using strlen() makes overall speed of the operation about
			// twice as fast for some typical test cases in a 2 MB buffer such as replacing \r\n with \n.
		}
		tmemcpy(match_pos, aNew, aNew_length); // Perform the replacement.
		// Must keep haystack_length updated as we go, for use with memmove() above:
		haystack_length += length_delta; // Note that length_delta will be negative if aNew is shorter than aOld.
	}

	result_length = haystack_length; // Set for caller (it's an alias for an output parameter).
	result = aHaystack; // Not actually needed in this method, so this is just for maintainability.
	return replacement_count;
}



size_t PredictReplacementSize(ptrdiff_t aLengthDelta, int aReplacementCount, int aLimit, size_t aHaystackLength
	, size_t aCurrentLength, size_t aEndOffsetOfCurrMatch)
// Predict how much size the remainder of a replacement operation will consume, including its actual replacements
// and the parts of haystack that won't need replacement.
// PARAMETERS:
// - aLengthDelta: The estimated or actual difference between the length of the replacement and what it's replacing.
//   A negative number means the replacement is smaller, which will cause a shrinking of the result.
// - aReplacementCount: The number of replacements so far, including the one the caller is about to do.
// - aLimit: The *remaining* number of replacements *allowed* (not including the one the caller is about to do).
// - aHaystackLength: The total length of the original haystack/subject string.
// - aCurrentLength: The total length of the new/result string including the one the caller is about to do.
// - aEndOffsetOfCurrMatch: The offset of the char after the last char of the current match.  For example, if
//   the empty string is the current match and it's found at the beginning of haystack, this value would be 0.
{
	// Since realloc() is an expensive operation, especially for huge strings, make an extra
	// effort to get a good estimate based on how things have been going so far.
	// While this should definitely improve average-case memory-utilization and usually performance
	// (by avoiding costly realloc's), this estimate is crude because:
	// 1) The length of what is being replaced can vary due to wildcards in pattern, etc.
	// 2) The length of what is replacing it can vary due to backreferences.  Thus, the delta
	//    of each replacement is only a guess based on that of the current replacement.
	// 3) For code simplicity, the number of upcoming replacements isn't yet known; thus a guess
	//    is made based on how many there have been so far compared to percentage complete.

	INT_PTR total_delta; // The total increase/decrease in length from the number of predicted additional replacements.
	int repl_multiplier = aLengthDelta < 0 ? -1 : 1; // Negative is used to keep additional_replacements_expected conservative even when delta is negative.

	if (aLengthDelta == 0) // Avoid all the calculations because it will wind up being zero anyway.
		total_delta = 0;
	else
	{
		if (!aHaystackLength // aHaystackLength can be 0 if an empty haystack being replaced by something else. If so, avoid divide-by-zero in the prediction by doing something simpler.
			|| !aEndOffsetOfCurrMatch)  // i.e. don't the prediction if the current match is the empty string and it was found at the very beginning of Haystack because it would be difficult to be accurate (very rare anyway).
			total_delta = repl_multiplier * aLengthDelta; // Due to rarity, just allow room for one extra after the one we're about to do.
		else // So the above has ensured that the following won't divide by zero anywhere.
		{
			// The following doesn't take into account the original value of aStartingOffset passed in
			// from the caller because:
			// 1) It's pretty rare for it to be greater than 0.
			// 2) Even if it is, the prediction will just be too conservative at first, but that's
			//    pretty harmless; and anyway each successive realloc followed by a match makes the
			//    prediction more and more accurate in spite of aStartingOffset>0.
			// percent_complete must be a double because we need at least 9 digits of precision for cases where
			// 1 is divided by a big number like 1 GB.
			double percent_complete = aEndOffsetOfCurrMatch  // Don't subtract 1 (verified correct).
				/ (double)aHaystackLength; // percent_complete isn't actually a percentage, but a fraction of 1.  e.g. 0.5 rather than 50.
			int additional_replacements_expected = percent_complete >= 1.0 ? 0  // It's often 100% complete, in which case there's hardly ever another replacement after this one (the only possibility is to replace the final empty-string in haystack with something).
				: (int)(
				(aReplacementCount / percent_complete) // This is basically "replacements per percentage point, so far".
				* (1 - percent_complete) // This is the percentage of haystack remaining to be scanned (e.g. 0.5 for 50%).
				+ 1 * repl_multiplier // Add 1 or -1 to make it more conservative (i.e. go the opposite direction of ceil when repl_multiplier is negative).
				);
			// additional_replacements_expected is defined as the replacements expected *after* the one the caller
			// is about to do.

			if (aLimit >= 0 && aLimit < additional_replacements_expected)
			{	// A limit is currently in effect and it's less than expected replacements, so cap the expected.
				// This helps reduce memory utilization.
				additional_replacements_expected = aLimit;
			}
			else // No limit or additional_replacements_expected is within the limit.
			{
				// So now things are set up so that there's about a 50/50 chance than no more reallocs
				// will be needed.  Since realloc is costly (due to internal memcpy), try to reduce
				// the odds of it happening without going overboard on memory utilization.
				// Something a lot more complicated could be used in place of the below to improve things
				// a little, but it just doesn't seem worth it given the usage patterns expected and
				// the actual benefits.  Besides, there is some limiting logic further below that will
				// cap this if it's unreasonably large:
				additional_replacements_expected += (int)(0.20*additional_replacements_expected + 1) // +1 so that there's always at least one extra.
					* repl_multiplier; // This keeps the estimate conservative if delta < 0.
			}
			// The following is the "quality" of the estimate.  For example, if this is the very first replacement
			// and 1000 more replacements are predicted, there will often be far fewer than 1000 replacements;
			// in fact, there could well be zero.  So in the below, the quality will range from 1 to 3, where
			// 1 is the worst quality and 3 is the best.
			double quality = 1 + 2*(1-(
				(double)additional_replacements_expected / (aReplacementCount + additional_replacements_expected)
				));
			// It seems best to use whichever of the following is greater in the calculation further below:
			INT_PTR haystack_or_new_length = (aCurrentLength > aHaystackLength) ? aCurrentLength : aHaystackLength;
			// The following is a crude sanity limit to avoid going overboard with memory
			// utilization in extreme cases such as when a big string has many replacements
			// in its first half, but hardly any in its second.  It does the following:
			// 1) When Haystack-or-current length is huge, this tries to keep the portion of the memory increase
			//    that's speculative proportionate to that length, which should reduce the chance of swapping
			//    (at the expense of some performance in cases where it causes another realloc to be required).
			// 2) When Haystack-or-current length is relatively small, allow the speculative memory allocation
			//    to be many times larger than that length because the risk of swapping is low.  HOWEVER, TO
			//    AVOID WASTING MEMORY, the caller should probably call _expand() to shrink the result
			//    when it detects that far fewer replacements were needed than predicted (this is currently
			//    done by Var::AcceptNewMem()).
			INT_PTR total_delta_limit = (INT_PTR)(haystack_or_new_length < 10*1024*1024 ? quality*10*1024*1024
				: quality*haystack_or_new_length); // See comment above.
			total_delta = additional_replacements_expected
				* (aLengthDelta < 0 ? -aLengthDelta : aLengthDelta); // So actually, total_delta will be the absolute value.
			if (total_delta > total_delta_limit)
				total_delta = total_delta_limit;
			total_delta *= repl_multiplier;  // Convert back from absolute value.
		} // The current match isn't an empty string at the very beginning of haystack.
	} // aLengthDelta!=0

	// Above is responsible for having set total_delta properly.
	INT_PTR subsequent_length = aHaystackLength - aEndOffsetOfCurrMatch // This is the length of the remaining portion of haystack that might wind up going into the result exactly as-is (adjusted by the below).
		+ total_delta; // This is additional_replacements_expected times the expected delta (the length of each replacement minus what it replaces) [can be negative].
	if (subsequent_length < 0) // Must not go below zero because that would cause the next line to
		subsequent_length = 0; // create an increase that's too small to handle the current replacement.

	// Return the sum of the following:
	// 1) subsequent_length: The predicted length needed for the remainder of the operation.
	// 2) aCurrentLength: The amount we need now, which includes room for the replacement the caller is about to do.
	//    Note that aCurrentLength can be 0 (such as for an empty string replacement).
	return subsequent_length + aCurrentLength + 1; // Caller relies on +1 for the terminator.
}



LPTSTR TranslateLFtoCRLF(LPTSTR aString)
// Can't use StrReplace() for this because any CRLFs originally present in aString are not changed (i.e. they
// don't become CRCRLF) [there may be other reasons].
// Translates any naked LFs in aString to CRLF.  If there are none, the original string is returned.
// Otherwise, the translated version is copied into a malloc'd buffer, which the caller must free
// when it's done with it).  
{
	UINT naked_LF_count = 0;
	size_t length = 0;
	LPTSTR cp;

	for (cp = aString; *cp; ++cp)
	{
		++length;
		if (*cp == '\n' && (cp == aString || cp[-1] != '\r')) // Relies on short-circuit boolean order.
			++naked_LF_count;
	}

	if (!naked_LF_count)
		return aString;  // The original string is returned, which the caller must check for (vs. new string).

	// Allocate the new memory that will become the caller's responsibility:
	LPTSTR buf = (LPTSTR)tmalloc(length + naked_LF_count + 1);  // +1 for zero terminator.
	if (!buf)
		return NULL;

	// Now perform the translation.
	LPTSTR dp = buf; // Destination.
	for (cp = aString; *cp; ++cp)
	{
		if (*cp == '\n' && (cp == aString || cp[-1] != '\r')) // Relies on short-circuit boolean order.
			*dp++ = '\r';  // Insert an extra CR here, then insert the '\n' normally below.
		*dp++ = *cp;
	}
	*dp = '\0';  // Final terminator.

	return buf;  // Caller must free it when it's done with it.
}



bool DoesFilePatternExist(LPTSTR aFilePattern, DWORD *aFileAttr)
// Returns true if the file/folder exists or false otherwise.
// If non-NULL, aFileAttr's DWORD is set to the attributes of the file/folder if a match is found.
// If there is no match, its contents are undefined.
{
	if (!aFilePattern || !*aFilePattern) return false;
	// Fix for v1.0.35.12: Don't consider the question mark in "\\?\Volume{GUID}\" to be a wildcard.
	// Such volume names are obtained from GetVolumeNameForVolumeMountPoint() and perhaps other functions.
	// However, testing shows that wildcards beyond that first one should be seen as real wildcards
	// because the wildcard match-method below does work for such volume names.
	LPTSTR cp = _tcsncmp(aFilePattern, _T("\\\\?\\"), 4) ? aFilePattern : aFilePattern + 4;
	if (StrChrAny(cp, _T("?*")))
	{
		WIN32_FIND_DATA wfd;
		HANDLE hFile = FindFirstFile(aFilePattern, &wfd);
		if (hFile == INVALID_HANDLE_VALUE)
			return false;
		FindClose(hFile);
		if (aFileAttr)
			*aFileAttr = wfd.dwFileAttributes;
		return true;
	}
    else
	{
		DWORD attr = GetFileAttributes(aFilePattern);
		if (aFileAttr)
			*aFileAttr = attr;
		return attr != 0xFFFFFFFF;
	}
}



#ifdef _DEBUG
ResultType FileAppend(LPTSTR aFilespec, LPTSTR aLine, bool aAppendNewline)
{
	if (!aFilespec || !aLine) return FAIL;
	if (!*aFilespec) return FAIL;
	FILE *fp = _tfopen(aFilespec, _T("a"));
	if (fp == NULL)
		return FAIL;
	_fputts(aLine, fp);
	if (aAppendNewline)
		_puttc('\n', fp);
	fclose(fp);
	return OK;
}
#endif



LPTSTR ConvertFilespecToCorrectCase(LPTSTR aFullFileSpec)
// aFullFileSpec must be a modifiable string since it will be converted to proper case.
// Also, it should be at least MAX_PATH is size because if it contains any short (8.3)
// components, they will be converted into their long names.
// Returns aFullFileSpec, the contents of which have been converted to the case used by the
// file system.  Note: The trick of changing the current directory to be that of
// aFullFileSpec and then calling GetFullPathName() doesn't always work.  So perhaps the
// only easy way is to call FindFirstFile() on each directory that composes aFullFileSpec,
// which is what is done here.  I think there's another way involving some PIDL Explorer
// function, but that might not support UNCs correctly.
{
	if (!aFullFileSpec || !*aFullFileSpec) return aFullFileSpec;
	size_t length = _tcslen(aFullFileSpec);
	if (length < 2 || length >= MAX_PATH) return aFullFileSpec;
	// Start with something easy, the drive letter:
	if (aFullFileSpec[1] == ':')
		aFullFileSpec[0] = ctoupper(aFullFileSpec[0]);
	// else it might be a UNC that has no drive letter.
	TCHAR built_filespec[MAX_PATH], *dir_start, *dir_end;
	if (dir_start = _tcschr(aFullFileSpec, ':'))
		// MSDN: "To examine any directory other than a root directory, use an appropriate
		// path to that directory, with no trailing backslash. For example, an argument of
		// "C:\windows" will return information about the directory "C:\windows", not about
		// any directory or file in "C:\windows". An attempt to open a search with a trailing
		// backslash will always fail."
		dir_start += 2; // Skip over the first backslash that goes with the drive letter.
	else // it's probably a UNC
	{
		if (_tcsncmp(aFullFileSpec, _T("\\\\"), 2))
			// It doesn't appear to be a UNC either, so not sure how to deal with it.
			return aFullFileSpec;
		// I think MS says you can't use FindFirstFile() directly on a share name, so we
		// want to omit both that and the server name from consideration (i.e. we don't attempt
		// to find their proper case).  MSDN: "Similarly, on network shares, you can use an
		// lpFileName of the form "\\server\service\*" but you cannot use an lpFileName that
		// points to the share itself, such as "\\server\service".
		dir_start = aFullFileSpec + 2;
		LPTSTR end_of_server_name = _tcschr(dir_start, '\\');
		if (end_of_server_name)
		{
			dir_start = end_of_server_name + 1;
			LPTSTR end_of_share_name = _tcschr(dir_start, '\\');
			if (end_of_share_name)
				dir_start = end_of_share_name + 1;
		}
	}
	// Init the new string (the filespec we're building), e.g. copy just the "c:\\" part.
	tcslcpy(built_filespec, aFullFileSpec, dir_start - aFullFileSpec + 1);
	WIN32_FIND_DATA found_file;
	HANDLE file_search;
	for (dir_end = dir_start; dir_end = _tcschr(dir_end, '\\'); ++dir_end)
	{
		*dir_end = '\0';  // Temporarily terminate.
		file_search = FindFirstFile(aFullFileSpec, &found_file);
		*dir_end = '\\'; // Restore it before we do anything else.
		if (file_search == INVALID_HANDLE_VALUE)
			return aFullFileSpec;
		FindClose(file_search);
		// Append the case-corrected version of this directory name:
		sntprintfcat(built_filespec, _countof(built_filespec), _T("%s\\"), found_file.cFileName);
	}
	// Now do the filename itself:
	if (   (file_search = FindFirstFile(aFullFileSpec, &found_file)) == INVALID_HANDLE_VALUE   )
		return aFullFileSpec;
	FindClose(file_search);
	sntprintfcat(built_filespec, _countof(built_filespec), _T("%s"), found_file.cFileName);
	// It might be possible for the new one to be longer than the old, e.g. if some 8.3 short
	// names were converted to long names by the process.  Thus, the caller should ensure that
	// aFullFileSpec is large enough:
	_tcscpy(aFullFileSpec, built_filespec);
	return aFullFileSpec;
}



LPTSTR FileAttribToStr(LPTSTR aBuf, DWORD aAttr)
// Caller must ensure that aAttr is valid (i.e. that it's not 0xFFFFFFFF).
{
	if (!aBuf) return aBuf;
	int length = 0;
	if (aAttr & FILE_ATTRIBUTE_READONLY)
		aBuf[length++] = 'R';
	if (aAttr & FILE_ATTRIBUTE_ARCHIVE)
		aBuf[length++] = 'A';
	if (aAttr & FILE_ATTRIBUTE_SYSTEM)
		aBuf[length++] = 'S';
	if (aAttr & FILE_ATTRIBUTE_HIDDEN)
		aBuf[length++] = 'H';
	if (aAttr & FILE_ATTRIBUTE_NORMAL)
		aBuf[length++] = 'N';
	if (aAttr & FILE_ATTRIBUTE_DIRECTORY)
		aBuf[length++] = 'D';
	if (aAttr & FILE_ATTRIBUTE_OFFLINE)
		aBuf[length++] = 'O';
	if (aAttr & FILE_ATTRIBUTE_COMPRESSED)
		aBuf[length++] = 'C';
	if (aAttr & FILE_ATTRIBUTE_TEMPORARY)
		aBuf[length++] = 'T';
	aBuf[length] = '\0';  // Perform the final termination.
	return aBuf;
}



unsigned __int64 GetFileSize64(HANDLE aFileHandle)
// Returns ULLONG_MAX on failure.  Otherwise, it returns the actual file size.
{
    ULARGE_INTEGER ui = {0};
    ui.LowPart = GetFileSize(aFileHandle, &ui.HighPart);
    if (ui.LowPart == MAXDWORD && GetLastError() != NO_ERROR)
        return ULLONG_MAX;
    return (unsigned __int64)ui.QuadPart;
}



LPTSTR GetLastErrorText(LPTSTR aBuf, int aBufSize, bool aUpdateLastError)
// aBufSize is an int to preserve any negative values the caller might pass in.
{
	if (aBufSize < 1)
		return aBuf;
	if (aBufSize == 1)
	{
		*aBuf = '\0';
		return aBuf;
	}
	DWORD last_error = GetLastError();
	if (aUpdateLastError)
		g->LastError = last_error;
	FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, NULL, last_error, 0, aBuf, (DWORD)aBufSize - 1, NULL);
	return aBuf;
}



void AssignColor(LPTSTR aColorName, COLORREF &aColor, HBRUSH &aBrush)
// Assign the color indicated in aColorName (either a name or a hex RGB value) to both
// aColor and aBrush, deleting any prior handle in aBrush first.  If the color cannot
// be determined, it will always be set to CLR_DEFAULT (and aBrush set to NULL to match).
// It will never be set to CLR_NONE.
{
	COLORREF color;
	if (!*aColorName)
		color = CLR_DEFAULT;
	else
	{
		color = ColorNameToBGR(aColorName);
		if (color == CLR_NONE) // A matching color name was not found, so assume it's a hex color value.
			// It seems strtol() automatically handles the optional leading "0x" if present:
			color = rgb_to_bgr(_tcstol(aColorName, NULL, 16));
			// if aColorName does not contain something hex-numeric, black (0x00) will be assumed,
			// which seems okay given how rare such a problem would be.
	}
	if (color != aColor) // It's not already the right color.
	{
		aColor = color; // Set default.  v1.0.44.09: Added this line to fix the inability to change to a previously selected color after having changed to the default color.
		if (aBrush) // Free the resources of the old brush.
			DeleteObject(aBrush);
		if (color == CLR_DEFAULT) // Caller doesn't need brush for CLR_DEFAULT, assuming that's even possible.
			aBrush = NULL;
		else
			if (   !(aBrush = CreateSolidBrush(color))   ) // Failure should be very rare.
				aColor = CLR_DEFAULT; // A NULL HBRUSH should always corresponds to CLR_DEFAULT.
	}
}



COLORREF ColorNameToBGR(LPTSTR aColorName)
// These are the main HTML color names.  Returns CLR_NONE if a matching HTML color name can't be found.
// Returns CLR_DEFAULT only if aColorName is the word Default.
{
	if (!aColorName || !*aColorName) return CLR_NONE;
	if (!_tcsicmp(aColorName, _T("Black")))  return 0x000000;  // These colors are all in BGR format, not RGB.
	if (!_tcsicmp(aColorName, _T("Silver"))) return 0xC0C0C0;
	if (!_tcsicmp(aColorName, _T("Gray")))   return 0x808080;
	if (!_tcsicmp(aColorName, _T("White")))  return 0xFFFFFF;
	if (!_tcsicmp(aColorName, _T("Maroon"))) return 0x000080;
	if (!_tcsicmp(aColorName, _T("Red")))    return 0x0000FF;
	if (!_tcsicmp(aColorName, _T("Purple"))) return 0x800080;
	if (!_tcsicmp(aColorName, _T("Fuchsia")))return 0xFF00FF;
	if (!_tcsicmp(aColorName, _T("Green")))  return 0x008000;
	if (!_tcsicmp(aColorName, _T("Lime")))   return 0x00FF00;
	if (!_tcsicmp(aColorName, _T("Olive")))  return 0x008080;
	if (!_tcsicmp(aColorName, _T("Yellow"))) return 0x00FFFF;
	if (!_tcsicmp(aColorName, _T("Navy")))   return 0x800000;
	if (!_tcsicmp(aColorName, _T("Blue")))   return 0xFF0000;
	if (!_tcsicmp(aColorName, _T("Teal")))   return 0x808000;
	if (!_tcsicmp(aColorName, _T("Aqua")))   return 0xFFFF00;
	if (!_tcsicmp(aColorName, _T("Default")))return CLR_DEFAULT;
	return CLR_NONE;
}



POINT CenterWindow(int aWidth, int aHeight)
// Given a the window's width and height, calculates where to position its upper-left corner
// so that it is centered EVEN IF the task bar is on the left side or top side of the window.
// This does not currently handle multi-monitor systems explicitly, since those calculations
// require API functions that don't exist in Win95/NT (and thus would have to be loaded
// dynamically to allow the program to launch).  Therefore, windows will likely wind up
// being centered across the total dimensions of all monitors, which usually results in
// half being on one monitor and half in the other.  This doesn't seem too terrible and
// might even be what the user wants in some cases (i.e. for really big windows).
{
	RECT rect;
	SystemParametersInfo(SPI_GETWORKAREA, 0, &rect, 0);  // Get desktop rect excluding task bar.
	// Note that rect.left will NOT be zero if the taskbar is on docked on the left.
	// Similarly, rect.top will NOT be zero if the taskbar is on docked at the top of the screen.
	POINT pt;
	pt.x = rect.left + (((rect.right - rect.left) - aWidth) / 2);
	pt.y = rect.top + (((rect.bottom - rect.top) - aHeight) / 2);
	return pt;
}



bool FontExist(HDC aHdc, LPCTSTR aTypeface)
{
	LOGFONT lf;
	lf.lfCharSet = DEFAULT_CHARSET;  // Enumerate all char sets.
	lf.lfPitchAndFamily = 0;  // Must be zero.
	tcslcpy(lf.lfFaceName, aTypeface, LF_FACESIZE);
	bool font_exists = false;
	EnumFontFamiliesEx(aHdc, &lf, (FONTENUMPROC)FontEnumProc, (LPARAM)&font_exists, 0);
	return font_exists;
}



int CALLBACK FontEnumProc(ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, DWORD FontType, LPARAM lParam)
{
	*(bool *)lParam = true; // Indicate to the caller that the font exists.
	return 0;  // Stop the enumeration after the first, since even one match means the font exists.
}



void ScreenToWindow(POINT &aPoint, HWND aHwnd)
// Convert screen coordinates to window coordinates (i.e. relative to the window's upper-left corner).
{
	RECT rect;
	GetWindowRect(aHwnd, &rect);
	aPoint.x -= rect.left;
	aPoint.y -= rect.top;
}



void CoordToScreen(int &aX, int &aY, int aWhichMode)
// aX and aY are interpreted according to the current coord mode.  If necessary, they are converted to
// screen coordinates based on the position of the active window's upper-left corner (or its client area).
{
	int coord_mode = ((g->CoordMode >> aWhichMode) & COORD_MODE_MASK);

	if (coord_mode == COORD_MODE_SCREEN)
		return;

	HWND active_window = GetForegroundWindow();
	if (active_window && !IsIconic(active_window))
	{
		if (coord_mode == COORD_MODE_WINDOW)
		{
			RECT rect;
			if (GetWindowRect(active_window, &rect))
			{
				aX += rect.left;
				aY += rect.top;
			}
		}
		else // (coord_mode == COORD_MODE_CLIENT)
		{
			POINT pt = {0};
			if (ClientToScreen(active_window, &pt))
			{
				aX += pt.x;
				aY += pt.y;
			}
		}
	}
	//else no active window per se, so don't convert the coordinates.  Leave them as-is as desired
	// by the caller.  More details:
	// Revert to screen coordinates if the foreground window is minimized.  Although it might be
	// impossible for a visible window to be both foreground and minimized, it seems that hidden
	// windows -- such as the script's own main window when activated for the purpose of showing
	// a popup menu -- can be foreground while simultaneously being minimized.  This fixes an
	// issue where the mouse will move to the upper-left corner of the screen rather than the
	// intended coordinates (v1.0.17).
}



void CoordToScreen(POINT &aPoint, int aWhichMode)
// For convenience. See function above for comments.
{
	CoordToScreen((int &)aPoint.x, (int &)aPoint.y, aWhichMode);
}



void GetVirtualDesktopRect(RECT &aRect)
{
	aRect.right = GetSystemMetrics(SM_CXVIRTUALSCREEN);
	if (aRect.right) // A non-zero value indicates the OS supports multiple monitors or at least SM_CXVIRTUALSCREEN.
	{
		aRect.left = GetSystemMetrics(SM_XVIRTUALSCREEN);  // Might be negative or greater than zero.
		aRect.right += aRect.left;
		aRect.top = GetSystemMetrics(SM_YVIRTUALSCREEN);   // Might be negative or greater than zero.
		aRect.bottom = aRect.top + GetSystemMetrics(SM_CYVIRTUALSCREEN);
	}
	else // Win95/NT do not support SM_CXVIRTUALSCREEN and such, so zero was returned.
		GetWindowRect(GetDesktopWindow(), &aRect);
}



DWORD GetEnvVarReliable(LPCTSTR aEnvVarName, LPTSTR aBuf)
// Returns the length of what has been copied into aBuf.
// Caller has ensured that aBuf is large enough (though anything >=32767 is always large enough).
// This function was added in v1.0.46.08 to fix a long-standing bug that was more fully revealed
// by 1.0.46.07's reduction of DEREF_BUF_EXPAND_INCREMENT from 32 to 16K (which allowed the Win9x
// version of GetEnvironmentVariable() to detect that we were lying about the buffer being 32767
// in size).  The reason for lying is that we don't know exactly how big the buffer is, but we do
// know it's large enough. So we just pass 32767 in an attempt to make it always succeed
// (mostly because GetEnvironmentVariable() is such a slow call, so it's best to avoid calling it
// a second time just to get the length).
// UPDATE: This is now called for all versions of Windows to avoid any chance of crashing or misbehavior,
// which tends to happen whenever lying to the API (even if it's for a good cause).  See similar problems
// at ReadRegString(). See other comments below.
{
	// Windows 9x is apparently capable of detecting a lie about the buffer size.
	// For example, if the memory capacity is only 16K but we pass 32767, it sometimes/always detects that
	// and returns 0, even if the string being retrieved is short enough (which it is because the caller
	// already verified it).  To work around this, give it a genuinely large buffer with an accurate size,
	// then copy the result out into the caller's variable.  This is almost certainly much faster than
	// doing a second call to GetEnvironmentVariable() to get the length because GetEnv() is known to be a
	// very slow call.
	//
	// Don't use a size greater than 32767 because that will cause it to fail on Win95 (tested by Robert Yalkin).
	// According to MSDN, 32767 is exactly large enough to handle the largest variable plus its zero terminator.
	TCHAR buf[32767];
	DWORD length = GetEnvironmentVariable(aEnvVarName, buf, _countof(buf));
	// GetEnvironmentVariable() could be called twice, the first time to get the actual size.  But that would
	// probably perform worse since GetEnvironmentVariable() is a very slow function.  In addition, it would
	// add code complexity, so it seems best to fetch it into a large buffer then just copy it to dest-var.
	if (length) // Probably always true under the conditions in effect for our callers.
		tmemcpy(aBuf, buf, length + 1); // memcpy() usually benches a little faster than strcpy().
	else // Failure. The buffer's contents might be undefined in this case.
		*aBuf = '\0'; // Caller's buf should always have room for an empty string. So make it empty for maintainability, even if not strictly required by caller.
	return length;
}



DWORD ReadRegString(HKEY aRootKey, LPTSTR aSubkey, LPTSTR aValueName, LPTSTR aBuf, DWORD aBufSize)
// Returns the length of the string (0 if empty).
// Caller must ensure that size of aBuf is REALLY aBufSize (even when it knows aBufSize is more than
// it needs) because the API apparently reads/writes parts of the buffer beyond the string it writes!
// Caller must ensure that aBuf isn't NULL because it doesn't seem worth having a "give me the size only" mode.
// This is because the API might return a size that omits the zero terminator, and the only way to find out for
// sure is probably to actually fetch the data and check if the terminator is present.
{
	HKEY hkey;
	if (RegOpenKeyEx(aRootKey, aSubkey, 0, KEY_QUERY_VALUE, &hkey) != ERROR_SUCCESS)
	{
		*aBuf = '\0';
		return 0;
	}
	DWORD buf_size = aBufSize * sizeof(TCHAR); // Caller's value might be a constant memory area, so need a modifiable copy.
	LONG result = RegQueryValueEx(hkey, aValueName, NULL, NULL, (LPBYTE)aBuf, &buf_size);
	RegCloseKey(hkey);
	if (result != ERROR_SUCCESS || !buf_size) // Relies on short-circuit boolean order.
	{
		*aBuf = '\0'; // MSDN says the contents of the buffer is undefined after the call in some cases, so reset it.
		return 0;
	}
	buf_size /= sizeof(TCHAR); // RegQueryValueEx returns size in bytes.
	// Otherwise success and non-empty result.  This also means that buf_size is accurate and <= to what we sent in.
	// Fix for v1.0.47: ENSURE PROPER STRING TERMINATION. This is suspected to be a source of crashing.
	// MSDN: "If the data has the REG_SZ, REG_MULTI_SZ or REG_EXPAND_SZ type, the string may not have been
	// stored with the proper null-terminating characters. Therefore, even if the function returns
	// ERROR_SUCCESS, the application should ensure that the string is properly terminated before using
	// it; otherwise, it may overwrite a buffer. (Note that REG_MULTI_SZ strings should have two
	// null-terminating characters.)"
	// MSDN: "If the data has the REG_SZ, REG_MULTI_SZ or REG_EXPAND_SZ type, this size includes any
	// terminating null character or characters."
	--buf_size; // Convert to the index of the last character written.  This is safe because above already checked that buf_size>0.
	if (aBuf[buf_size] == '\0') // It wrote at least one zero terminator (it might write more than one for Unicode, or if it's simply stored with more than one in the registry).
	{
		while (buf_size && !aBuf[buf_size - 1]) // Scan leftward in case more than one consecutive terminator.
			--buf_size;
		return buf_size; // There's a tiny chance that this could the wrong length, namely when the string contains binary zero(s) to the left of non-zero characters.  But that seems too rare to be worth calling strlen() for.
	}
	// Otherwise, it didn't write a terminator, so provide one.
	++buf_size; // To reflect the fact that we're about to write an extra char.
	if (buf_size >= aBufSize) // There's no room for a terminator without truncating the data (very rare).  Seems best to indicate failure.
	{
		*aBuf = '\0';
		return 0;
	}
	// Otherwise, there's room for the terminator.
	aBuf[buf_size] = '\0';
	return buf_size; // There's a tiny chance that this could the wrong length, namely when the string contains binary zero(s) to the left of non-zero characters.  But that seems too rare to be worth calling strlen() for.
}



LPVOID AllocInterProcMem(HANDLE &aHandle, DWORD aSize, HWND aHwnd)
// aHandle is an output parameter that holds the mapping for Win9x and the process handle for NT.
// Returns NULL on failure (in which case caller should ignore the value of aHandle).
{
	// ALLOCATE APPROPRIATE TYPE OF MEMORY (depending on OS type)
	LPVOID mem;
	if (g_os.IsWin9x()) // Use file-mapping method.
	{
		if (   !(aHandle = CreateFileMapping(INVALID_HANDLE_VALUE, NULL, PAGE_READWRITE, 0, aSize, NULL))   )
			return NULL;
		mem = MapViewOfFile(aHandle, FILE_MAP_ALL_ACCESS, 0, 0, 0);
	}
	else // NT/2k/XP/2003 or later.  Use the VirtualAllocEx() so that caller can use Read/WriteProcessMemory().
	{
		DWORD pid;
		GetWindowThreadProcessId(aHwnd, &pid);
		// Even if the PID is our own, open the process anyway to simplify the code. After all, it would be
		// pretty silly for a script to access its own ListViews via this method.
		if (   !(aHandle = OpenProcess(PROCESS_VM_OPERATION | PROCESS_VM_READ | PROCESS_VM_WRITE, FALSE, pid))   )
			return NULL; // Let ErrorLevel tell the story.
		// Load function dynamically to allow program to launch on win9x:
		typedef LPVOID (WINAPI *MyVirtualAllocExType)(HANDLE, LPVOID, SIZE_T, DWORD, DWORD);
		static MyVirtualAllocExType MyVirtualAllocEx = (MyVirtualAllocExType)GetProcAddress(GetModuleHandle(_T("kernel32"))
			, "VirtualAllocEx");
		// Reason for using VirtualAllocEx(): When sending LVITEM structures to a control in a remote process, the
		// structure and its pszText buffer must both be memory inside the remote process rather than in our own.
		mem = MyVirtualAllocEx(aHandle, NULL, aSize, MEM_RESERVE | MEM_COMMIT, PAGE_READWRITE);
	}
	if (!mem)
		CloseHandle(aHandle); // Closes the mapping for Win9x and the process handle for other OSes. Caller should ignore the value of aHandle when return value is NULL.
	//else leave the handle open (required for both methods).  It's the caller's responsibility to close it.
	return mem;
}



void FreeInterProcMem(HANDLE aHandle, LPVOID aMem)
// Caller has ensured that aMem is a file-mapping for Win9x and a VirtualAllocEx block for NT/2k/XP+.
// Similarly, it has ensured that aHandle is a file-mapping handle for Win9x and a process handle for NT/2k/XP+.
{
	if (g_os.IsWin9x())
		UnmapViewOfFile(aMem);
	else
	{
		// Load function dynamically to allow program to launch on win9x:
		typedef BOOL (WINAPI *MyVirtualFreeExType)(HANDLE, LPVOID, SIZE_T, DWORD);
		static MyVirtualFreeExType MyVirtualFreeEx = (MyVirtualFreeExType)GetProcAddress(GetModuleHandle(_T("kernel32"))
			, "VirtualFreeEx");
		MyVirtualFreeEx(aHandle, aMem, 0, MEM_RELEASE); // Size 0 is used with MEM_RELEASE.
	}
	// The following closes either the mapping or the process handle, depending on OS type.
	// But close it only after the above is done using it.
	CloseHandle(aHandle);
}



HBITMAP LoadPicture(LPTSTR aFilespec, int aWidth, int aHeight, int &aImageType, int aIconNumber
	, bool aUseGDIPlusIfAvailable)
// Returns NULL on failure.
// If aIconNumber > 0, an HICON or HCURSOR is returned (both should be interchangeable), never an HBITMAP.
// However, aIconNumber==1 is treated as a special icon upon which LoadImage is given preference over ExtractIcon
// for .ico/.cur/.ani files.
// Otherwise, .ico/.cur/.ani files are normally loaded as HICON (unless aUseGDIPlusIfAvailable is true or
// something else unusual happened such as file contents not matching file's extension).  This is done to preserve
// any properties that HICONs have but HBITMAPs lack, namely the ability to be animated and perhaps other things.
//
// Loads a JPG/GIF/BMP/ICO/etc. and returns an HBITMAP or HICON to the caller (which it may call
// DeleteObject()/DestroyIcon() upon, though upon program termination all such handles are freed
// automatically).  The image is scaled to the specified width and height.  If zero is specified
// for either, the image's actual size will be used for that dimension.  If -1 is specified for one,
// that dimension will be kept proportional to the other dimension's size so that the original aspect
// ratio is retained.
{
	HBITMAP hbitmap = NULL;
	aImageType = -1; // The type of image currently inside hbitmap.  Set default value for output parameter as "unknown".

	if (!*aFilespec) // Allow blank filename to yield NULL bitmap (and currently, some callers do call it this way).
		return NULL;
	// Lexikos: Negative values now indicate an icon's integer resource ID.
	//if (aIconNumber < 0) // Allowed to be called this way by GUI and others (to avoid need for validation of user input there).
	//	aIconNumber = 0; // Use the default behavior, which is "load icon or bitmap, whichever is most appropriate".

	LPTSTR file_ext = _tcsrchr(aFilespec, '.');
	if (file_ext)
		++file_ext;

	// v1.0.43.07: If aIconNumber is zero, caller didn't specify whether it wanted an icon or bitmap.  Thus,
	// there must be some kind of detection for whether ExtractIcon is needed instead of GDIPlus/OleLoadPicture.
	// Although this could be done by attempting ExtractIcon only after GDIPlus/OleLoadPicture fails (or by
	// somehow checking the internal nature of the file), for performance and code size, it seems best to not
	// to incur this extra I/O and instead make only one attempt based on the file's extension.
	// Must use ExtractIcon() if either of the following is true:
	// 1) Caller gave an icon index of the second or higher icon in the file.  Update for v1.0.43.05: There
	//    doesn't seem to be any reason to allow a caller to explicitly specify ExtractIcon as the method of
	//    loading the *first* icon from a .ico file since LoadImage is likely always superior.  This is
	//    because unlike ExtractIcon/Ex, LoadImage: 1) Doesn't distort icons, especially 16x16 icons; 2) is
	//    capable of loading icons other than the first by means of width and height parameters.
	// 2) The target file is of type EXE/DLL/ICL/CPL/etc. (LoadImage() is documented not to work on those file types).
	//    ICL files (v1.0.43.05): Apparently ICL files are an unofficial file format. Someone on the newsgroups
	//    said that an ICL is an "ICon Library... a renamed 16-bit Windows .DLL (an NE format executable) which
	//    typically contains nothing but a resource section. The ICL extension seems to be used by convention."
	// L17: Support negative numbers to mean resource IDs. These are supported by the resource extraction method directly, and by ExtractIcon if aIconNumber < -1.
	bool ExtractIcon_was_used = aIconNumber > 1 || aIconNumber < 0 || (file_ext && (
		   !_tcsicmp(file_ext, _T("exe"))
		|| !_tcsicmp(file_ext, _T("dll"))
		|| !_tcsicmp(file_ext, _T("icl")) // Icon library: Unofficial dll container, see notes above.
		|| !_tcsicmp(file_ext, _T("cpl")) // Control panel extension/applet (ExtractIcon is said to work on these).
		|| !_tcsicmp(file_ext, _T("scr")) // Screen saver (ExtractIcon should work since these are really EXEs).
		// v1.0.44: Below are now omitted to reduce code size and improve performance. They are still supported
		// indirectly because ExtractIcon is attempted whenever LoadImage() fails further below.
		//|| !_tcsicmp(file_ext, _T("drv")) // Driver (ExtractIcon is said to work on these).
		//|| !_tcsicmp(file_ext, _T("ocx")) // OLE/ActiveX Control Extension
		//|| !_tcsicmp(file_ext, _T("vbx")) // Visual Basic Extension
		//|| !_tcsicmp(file_ext, _T("acm")) // Audio Compression Manager Driver
		//|| !_tcsicmp(file_ext, _T("bpl")) // Delphi Library (like a DLL?)
		// Not supported due to rarity, code size, performance, and uncertainty of whether ExtractIcon works on them.
		// Update for v1.0.44: The following are now supported indirectly because ExtractIcon is attempted whenever
		// LoadImage() fails further below.
		//|| !_tcsicmp(file_ext, _T("nil")) // Norton Icon Library 
		//|| !_tcsicmp(file_ext, _T("wlx")) // Total/Windows Commander Lister Plug-in
		//|| !_tcsicmp(file_ext, _T("wfx")) // Total/Windows Commander File System Plug-in
		//|| !_tcsicmp(file_ext, _T("wcx")) // Total/Windows Commander Plug-in
		//|| !_tcsicmp(file_ext, _T("wdx")) // Total/Windows Commander Plug-in
		));

	if (ExtractIcon_was_used)
	{
		aImageType = IMAGE_ICON;

		// L17: Manually extract the most appropriately sized icon resource for the best results.
		hbitmap = (HBITMAP)ExtractIconFromExecutable(aFilespec, aIconNumber, aWidth, aHeight);

		// At this time it seems unnecessary to fall back to any of the following methods:
		/*
		if (aWidth && aHeight) // Lexikos: Alternate methods that give better results than ExtractIcon.
		{
			int icon_index = aIconNumber > 0 ? aIconNumber - 1 : aIconNumber < -1 ? aIconNumber : 0;
			// Testing shows ExtractIcon retrieves the first icon in the icon group and resizes it to
			// the system large icon size. In case scripts expect this behaviour, and because it is
			// debatable which size will actually look better when scaled, use ExtractIconEX only if
			// a system size is exactly specified. Callers who specifically want the system small or
			// large icon should call GetSystemMetrics() to get the correct aWidth and aHeight.
			// For convenience and because the system icon sizes are probably always square,
			// assume aWidth == -1 is equivalent to aWidth == aHeight (and vice versa).
			if ((aWidth == GetSystemMetrics(SM_CXSMICON) || aWidth == -1) && (aHeight == GetSystemMetrics(SM_CYSMICON) || aHeight == -1))
			{
				// Retrieve icon via phiconSmall.
				ExtractIconEx(aFilespec, icon_index, NULL, (HICON*)&hbitmap, 1);
			}
			else if ((aWidth == GetSystemMetrics(SM_CXICON) || aWidth == -1) && (aHeight == GetSystemMetrics(SM_CYICON) || aHeight == -1))
			{
				// Retrieve icon via phiconLarge.
				ExtractIconEx(aFilespec, icon_index, (HICON*)&hbitmap, NULL, 1);
			}
			else
			{
				UINT icon_id;
				// This is the simplest method of loading an icon of an arbitrary size.
				// "However, this function is deprecated not intended for general use."
				// It is also not documented to support resource IDs, unlike ExtractIcon and ExtractIconEx.
				if (-1 == PrivateExtractIcons(aFilespec, icon_index, aWidth == -1 ? aHeight : aWidth, aHeight == -1 ? aWidth : aHeight, (HICON*)&hbitmap, &icon_id, 1, 0))
					hbitmap = NULL;
			}
		}
		// Lexikos: ExtractIcon supports resource IDs by specifying nIconIndex < -1, where Abs(nIconIndex)
		// is the resource ID. nIconIndex MUST NOT BE -1, which is used to indicate the total number of
		// icons should be returned. ExtractIconEx could be used to support ID -1, but it has different
		// behaviour to ExtractIcon, and the method above should work in most instances.
		if (!hbitmap) // The other methods failed or were not used.
			hbitmap = (HBITMAP)ExtractIcon(g_hInstance, aFilespec, aIconNumber > 0 ? aIconNumber - 1 : aIconNumber < -1 ? aIconNumber : 0);
		// Above: Although it isn't well documented at MSDN, apparently both ExtractIcon() and LoadIcon()
		// scale the icon to the system's large-icon size (usually 32x32) regardless of the actual size of
		// the icon inside the file.  For this reason, callers should call us in a way that allows us to
		// give preference to LoadImage() over ExtractIcon() (unless the caller needs to retain backward
		// compatibility with existing scripts that explicitly specify icon #1 to force the ExtractIcon
		// method to be used).
		*/
		if (hbitmap < (HBITMAP)2) // i.e. it's NULL or 1. Return value of 1 means "incorrect file type".
			return NULL; // v1.0.44: Fixed to return NULL vs. hbitmap, since 1 is an invalid handle (perhaps rare since no known bugs caused by it).
		//else continue on below so that the icon can be resized to the caller's specified dimensions.
	}
	else if (aIconNumber > 0) // Caller wanted HICON, never HBITMAP, so set type now to enforce that.
		aImageType = IMAGE_ICON; // Should be suitable for cursors too, since they're interchangeable for the most part.
	else if (file_ext) // Make an initial guess of the type of image if the above didn't already determine the type.
	{
		if (!_tcsicmp(file_ext, _T("ico")))
			aImageType = IMAGE_ICON;
		else if (!_tcsicmp(file_ext, _T("cur")) || !_tcsicmp(file_ext, _T("ani")))
			aImageType = IMAGE_CURSOR;
		else if (!_tcsicmp(file_ext, _T("bmp")))
			aImageType = IMAGE_BITMAP;
		//else for other extensions, leave set to "unknown" so that the below knows to use IPic or GDI+ to load it.
	}
	//else same comment as above.

	if ((aWidth == -1 || aHeight == -1) && (!aWidth || !aHeight))
		aWidth = aHeight = 0; // i.e. One dimension is zero and the other is -1, which resolves to the same as "keep original size".
	bool keep_aspect_ratio = (aWidth == -1 || aHeight == -1);

	// Caller should ensure that aUseGDIPlusIfAvailable==false when aIconNumber > 0, since it makes no sense otherwise.
	HINSTANCE hinstGDI = NULL;
	if (aUseGDIPlusIfAvailable && !(hinstGDI = LoadLibrary(_T("gdiplus")))) // Relies on short-circuit boolean order for performance.
		aUseGDIPlusIfAvailable = false; // Override any original "true" value as a signal for the section below.

	if (!hbitmap && aImageType > -1 && !aUseGDIPlusIfAvailable)
	{
		// Since image hasn't yet be loaded and since the file type appears to be one supported by
		// LoadImage() [icon/cursor/bitmap], attempt that first.  If it fails, fall back to the other
		// methods below in case the file's internal contents differ from what the file extension indicates.
		int desired_width, desired_height;
		if (keep_aspect_ratio) // Load image at its actual size.  It will be rescaled to retain aspect ratio later below.
		{
			desired_width = 0;
			desired_height = 0;
		}
		else
		{
			desired_width = aWidth;
			desired_height = aHeight;
		}
		// For LoadImage() below:
		// LR_CREATEDIBSECTION applies only when aImageType == IMAGE_BITMAP, but seems appropriate in that case.
		// Also, if width and height are non-zero, that will determine which icon of a multi-icon .ico file gets
		// loaded (though I don't know the exact rules of precedence).
		// KNOWN LIMITATIONS/BUGS:
		// LoadImage() fails when requesting a size of 1x1 for an image whose orig/actual size is small (e.g. 1x2).
		// Unlike CopyImage(), perhaps it detects that division by zero would occur and refuses to do the
		// calculation rather than providing more code to do a correct calculation that doesn't divide by zero.
		// For example:
		// LoadImage() Success:
		//   Gui, Add, Pic, h2 w2, bitmap 1x2.bmp
		//   Gui, Add, Pic, h1 w1, bitmap 4x6.bmp
		// LoadImage() Failure:
		//   Gui, Add, Pic, h1 w1, bitmap 1x2.bmp
		// LoadImage() also fails on:
		//   Gui, Add, Pic, h1, bitmap 1x2.bmp
		// And then it falls back to GDIplus, which in the particular case above appears to traumatize the
		// parent window (or its picture control), because the GUI window hangs (but not the script) after
		// doing a FileSelectFolder.  For example:
		//   Gui, Add, Button,, FileSelectFile
		//   Gui, Add, Pic, h1, bitmap 1x2.bmp  ; Causes GUI window to hang after FileSelectFolder (due to LoadImage failing then falling back to GDIplus; i.e. GDIplus is somehow triggering the problem).
		//   Gui, Show
		//   return
		//   ButtonFileSelectFile:
		//   FileSelectFile, outputvar
		//   return
		if (hbitmap = (HBITMAP)LoadImage(NULL, aFilespec, aImageType, desired_width, desired_height
			, LR_LOADFROMFILE | LR_CREATEDIBSECTION))
		{
			// The above might have loaded an HICON vs. an HBITMAP (it has been confirmed that LoadImage()
			// will return an HICON vs. HBITMAP is aImageType is IMAGE_ICON/CURSOR).  Note that HICON and
			// HCURSOR are identical for most/all Windows API uses.  Also note that LoadImage() will load
			// an icon as a bitmap if the file contains an icon but IMAGE_BITMAP was passed in (at least
			// on Windows XP).
			if (!keep_aspect_ratio) // No further resizing is needed.
				return hbitmap;
			// Otherwise, continue on so that the image can be resized via a second call to LoadImage().
		}
		// v1.0.40.10: Abort if file doesn't exist so that GDIPlus isn't even attempted. This is done because
		// loading GDIPlus apparently disrupts the color palette of certain games, at least old ones that use
		// DirectDraw in 256-color depth.
		else if (GetFileAttributes(aFilespec) == 0xFFFFFFFF) // For simplicity, we don't check if it's a directory vs. file, since that should be too rare.
			return NULL;
		// v1.0.43.07: Also abort if caller wanted an HICON (not an HBITMAP), since the other methods below
		// can't yield an HICON.
		else if (aIconNumber > 0)
		{
			// UPDATE for v1.0.44: Attempt ExtractIcon in case its some extension that's
			// was recognized as an icon container (such as AutoHotkeySC.bin) and thus wasn't handled higher above.
			//hbitmap = (HBITMAP)ExtractIcon(g_hInstance, aFilespec, aIconNumber - 1);

			// L17: Manually extract the most appropriately sized icon resource for the best results.
			hbitmap = (HBITMAP)ExtractIconFromExecutable(aFilespec, aIconNumber, aWidth, aHeight);

			if (hbitmap < (HBITMAP)2) // i.e. it's NULL or 1. Return value of 1 means "incorrect file type".
				return NULL;
			ExtractIcon_was_used = true;
		}
		//else file exists, so continue on so that the other methods are attempted in case file's contents
		// differ from what the file extension indicates, or in case the other methods can be successful
		// even when the above failed.
	}

	IPicture *pic = NULL; // Also used to detect whether IPic method was used to load the image.

	if (!hbitmap) // Above hasn't loaded the image yet, so use the fall-back methods.
	{
		// At this point, regardless of the image type being loaded (even an icon), it will
		// definitely be converted to a Bitmap below.  So set the type:
		aImageType = IMAGE_BITMAP;
		// Find out if this file type is supported by the non-GDI+ method.  This check is not foolproof
		// since all it does is look at the file's extension, not its contents.  However, it doesn't
		// need to be 100% accurate because its only purpose is to detect whether the higher-overhead
		// calls to GdiPlus can be avoided.
		if (aUseGDIPlusIfAvailable || !file_ext || (_tcsicmp(file_ext, _T("jpg"))
			&& _tcsicmp(file_ext, _T("jpeg")) && _tcsicmp(file_ext, _T("gif")))) // Non-standard file type (BMP is already handled above).
			if (!hinstGDI) // We don't yet have a handle from an earlier call to LoadLibary().
				hinstGDI = LoadLibrary(_T("gdiplus"));
		// If it is suspected that the file type isn't supported, try to use GdiPlus if available.
		// If it's not available, fall back to the old method in case the filename doesn't properly
		// reflect its true contents (i.e. in case it really is a JPG/GIF/BMP internally).
		// If the below LoadLibrary() succeeds, either the OS is XP+ or the GdiPlus extensions have been
		// installed on an older OS.
		if (hinstGDI)
		{
			// LPVOID and "int" are used to avoid compiler errors caused by... namespace issues?
			typedef int (WINAPI *GdiplusStartupType)(ULONG_PTR*, LPVOID, LPVOID);
			typedef VOID (WINAPI *GdiplusShutdownType)(ULONG_PTR);
			typedef int (WINGDIPAPI *GdipCreateBitmapFromFileType)(LPVOID, LPVOID);
			typedef int (WINGDIPAPI *GdipCreateHBITMAPFromBitmapType)(LPVOID, LPVOID, DWORD);
			typedef int (WINGDIPAPI *GdipDisposeImageType)(LPVOID);
			GdiplusStartupType DynGdiplusStartup = (GdiplusStartupType)GetProcAddress(hinstGDI, "GdiplusStartup");
  			GdiplusShutdownType DynGdiplusShutdown = (GdiplusShutdownType)GetProcAddress(hinstGDI, "GdiplusShutdown");
  			GdipCreateBitmapFromFileType DynGdipCreateBitmapFromFile = (GdipCreateBitmapFromFileType)GetProcAddress(hinstGDI, "GdipCreateBitmapFromFile");
  			GdipCreateHBITMAPFromBitmapType DynGdipCreateHBITMAPFromBitmap = (GdipCreateHBITMAPFromBitmapType)GetProcAddress(hinstGDI, "GdipCreateHBITMAPFromBitmap");
  			GdipDisposeImageType DynGdipDisposeImage = (GdipDisposeImageType)GetProcAddress(hinstGDI, "GdipDisposeImage");

			ULONG_PTR token;
			Gdiplus::GdiplusStartupInput gdi_input;
			Gdiplus::GpBitmap *pgdi_bitmap;
			if (DynGdiplusStartup && DynGdiplusStartup(&token, &gdi_input, NULL) == Gdiplus::Ok)
			{
#ifndef UNICODE
				WCHAR filespec_wide[MAX_PATH];
				ToWideChar(aFilespec, filespec_wide, MAX_PATH); // Dest. size is in wchars, not bytes.
				if (DynGdipCreateBitmapFromFile(filespec_wide, &pgdi_bitmap) == Gdiplus::Ok)
#else
				if (DynGdipCreateBitmapFromFile(aFilespec, &pgdi_bitmap) == Gdiplus::Ok)
#endif
				{
					if (DynGdipCreateHBITMAPFromBitmap(pgdi_bitmap, &hbitmap, CLR_DEFAULT) != Gdiplus::Ok)
						hbitmap = NULL; // Set to NULL to be sure.
					DynGdipDisposeImage(pgdi_bitmap); // This was tested once to make sure it really returns Gdiplus::Ok.
				}
				// The current thought is that shutting it down every time conserves resources.  If so, it
				// seems justified since it is probably called infrequently by most scripts:
				DynGdiplusShutdown(token);
			}
			FreeLibrary(hinstGDI);
		}
		else // Using old picture loading method.
		{
			// Based on code sample at http://www.codeguru.com/Cpp/G-M/bitmap/article.php/c4935/
			HANDLE hfile = CreateFile(aFilespec, GENERIC_READ, 0, NULL, OPEN_EXISTING, 0, NULL);
			if (hfile == INVALID_HANDLE_VALUE)
				return NULL;
			DWORD size = GetFileSize(hfile, NULL);
			HGLOBAL hglobal = GlobalAlloc(GMEM_MOVEABLE, size);
			if (!hglobal)
			{
				CloseHandle(hfile);
				return NULL;
			}
			LPVOID hlocked = GlobalLock(hglobal);
			if (!hlocked)
			{
				CloseHandle(hfile);
				GlobalFree(hglobal);
				return NULL;
			}
			// Read the file into memory:
			ReadFile(hfile, hlocked, size, &size, NULL);
			GlobalUnlock(hglobal);
			CloseHandle(hfile);
			LPSTREAM stream;
			if (FAILED(CreateStreamOnHGlobal(hglobal, FALSE, &stream)) || !stream)  // Relies on short-circuit boolean order.
			{
				GlobalFree(hglobal);
				return NULL;
			}
			// Specify TRUE to have it do the GlobalFree() for us.  But since the call might fail, it seems best
			// to free the mem ourselves to avoid uncertainty over what it does on failure:
			if (FAILED(OleLoadPicture(stream, 0, FALSE, IID_IPicture, (void **)&pic)))
				pic = NULL;
			stream->Release();
			GlobalFree(hglobal);
			if (!pic)
				return NULL;
			pic->get_Handle((OLE_HANDLE *)&hbitmap);
			// Above: MSDN: "The caller is responsible for this handle upon successful return. The variable is set
			// to NULL on failure."
			if (!hbitmap)
			{
				pic->Release();
				return NULL;
			}
			// Don't pic->Release() yet because that will also destroy/invalidate hbitmap handle.
		} // IPicture method was used.
	} // IPicture or GDIPlus was used to load the image, not a simple LoadImage() or ExtractIcon().

	// Above has ensured that hbitmap is now not NULL.
	// Adjust things if "keep aspect ratio" is in effect:
	if (keep_aspect_ratio)
	{
		HBITMAP hbitmap_to_analyze;
		ICONINFO ii; // Must be declared at this scope level.
		if (aImageType == IMAGE_BITMAP)
			hbitmap_to_analyze = hbitmap;
		else // icon or cursor
		{
			if (GetIconInfo((HICON)hbitmap, &ii)) // Works on cursors too.
				hbitmap_to_analyze = ii.hbmMask; // Use Mask because MSDN implies hbmColor can be NULL for monochrome cursors and such.
			else
			{
				DestroyIcon((HICON)hbitmap);
				return NULL; // No need to call pic->Release() because since it's an icon, we know IPicture wasn't used (it only loads bitmaps).
			}
		}
		// Above has ensured that hbitmap_to_analyze is now not NULL.  Find bitmap's dimensions.
		BITMAP bitmap;
		GetObject(hbitmap_to_analyze, sizeof(BITMAP), &bitmap); // Realistically shouldn't fail at this stage.
		if (aHeight == -1)
		{
			// Caller wants aHeight calculated based on the specified aWidth (keep aspect ratio).
			if (bitmap.bmWidth) // Avoid any chance of divide-by-zero.
				aHeight = (int)(((double)bitmap.bmHeight / bitmap.bmWidth) * aWidth + .5); // Round.
		}
		else
		{
			// Caller wants aWidth calculated based on the specified aHeight (keep aspect ratio).
			if (bitmap.bmHeight) // Avoid any chance of divide-by-zero.
				aWidth = (int)(((double)bitmap.bmWidth / bitmap.bmHeight) * aHeight + .5); // Round.
		}
		if (aImageType != IMAGE_BITMAP)
		{
			// It's our responsibility to delete these two when they're no longer needed:
			DeleteObject(ii.hbmColor);
			DeleteObject(ii.hbmMask);
			// If LoadImage() vs. ExtractIcon() was used originally, call LoadImage() again because
			// I haven't found any other way to retain an animated cursor's animation (and perhaps
			// other icon/cursor attributes) when resizing the icon/cursor (CopyImage() doesn't
			// retain animation):
			if (!ExtractIcon_was_used)
			{
				DestroyIcon((HICON)hbitmap); // Destroy the original HICON.
				// Load a new one, but at the size newly calculated above.
				// Due to an apparent bug in Windows 9x (at least Win98se), the below call will probably
				// crash the program with a "divide error" if the specified aWidth and/or aHeight are
				// greater than 90.  Since I don't know whether this affects all versions of Windows 9x, and
				// all animated cursors, it seems best just to document it here and in the help file rather
				// than limiting the dimensions of .ani (and maybe .cur) files for certain operating systems.
				return (HBITMAP)LoadImage(NULL, aFilespec, aImageType, aWidth, aHeight, LR_LOADFROMFILE);
			}
		}
	}

	HBITMAP hbitmap_new; // To hold the scaled image (if scaling is needed).
	if (pic) // IPicture method was used.
	{
		// The below statement is confirmed by having tested that DeleteObject(hbitmap) fails
		// if called after pic->Release():
		// "Copy the image. Necessary, because upon pic's release the handle is destroyed."
		// MSDN: CopyImage(): "[If either width or height] is zero, then the returned image will have the
		// same width/height as the original."
		// Note also that CopyImage() seems to provide better scaling quality than using MoveWindow()
		// (followed by redrawing the parent window) on the static control that contains it:
		hbitmap_new = (HBITMAP)CopyImage(hbitmap, IMAGE_BITMAP, aWidth, aHeight // We know it's IMAGE_BITMAP in this case.
			, (aWidth || aHeight) ? 0 : LR_COPYRETURNORG); // Produce original size if no scaling is needed.
		pic->Release();
		// No need to call DeleteObject(hbitmap), see above.
	}
	else // GDIPlus or a simple method such as LoadImage or ExtractIcon was used.
	{
		if (!aWidth && !aHeight) // No resizing needed.
			return hbitmap;
		// The following will also handle HICON/HCURSOR correctly if aImageType == IMAGE_ICON/CURSOR.
		// Also, LR_COPYRETURNORG|LR_COPYDELETEORG is used because it might allow the animation of
		// a cursor to be retained if the specified size happens to match the actual size of the
		// cursor.  This is because normally, it seems that CopyImage() omits cursor animation
		// from the new object.  MSDN: "LR_COPYRETURNORG returns the original hImage if it satisfies
		// the criteria for the copy—that is, correct dimensions and color depth—in which case the
		// LR_COPYDELETEORG flag is ignored. If this flag is not specified, a new object is always created."
		// KNOWN BUG: Calling CopyImage() when the source image is tiny and the destination width/height
		// is also small (e.g. 1) causes a divide-by-zero exception.
		// For example:
		//   Gui, Add, Pic, h1 w-1, bitmap 1x2.bmp  ; Crash (divide by zero)
		//   Gui, Add, Pic, h1 w-1, bitmap 2x3.bmp  ; Crash (divide by zero)
		// However, such sizes seem too rare to document or put in an exception handler for.
		hbitmap_new = (HBITMAP)CopyImage(hbitmap, aImageType, aWidth, aHeight, LR_COPYRETURNORG | LR_COPYDELETEORG);
		// Above's LR_COPYDELETEORG deletes the original to avoid cascading resource usage.  MSDN's
		// LoadImage() docs say:
		// "When you are finished using a bitmap, cursor, or icon you loaded without specifying the
		// LR_SHARED flag, you can release its associated memory by calling one of [the three functions]."
		// Therefore, it seems best to call the right function even though DeleteObject might work on
		// all of them on some or all current OSes.  UPDATE: Evidence indicates that DestroyIcon()
		// will also destroy cursors, probably because icons and cursors are literally identical in
		// every functional way.  One piece of evidence:
		//> No stack trace, but I know the exact source file and line where the call
		//> was made. But still, it is annoying when you see 'DestroyCursor' even though
		//> there is 'DestroyIcon'.
		// "Can't be helped. Icons and cursors are the same thing" (Tim Robinson (MVP, Windows SDK)).
		//
		// Finally, the reason this is important is that it eliminates one handle type
		// that we would otherwise have to track.  For example, if a gui window is destroyed and
		// and recreated multiple times, its bitmap and icon handles should all be destroyed each time.
		// Otherwise, resource usage would cascade upward until the script finally terminated, at
		// which time all such handles are freed automatically.
	}
	return hbitmap_new;
}


struct ResourceIndexToIdEnumData
{
	int find_index;
	int index;
	int result;
};

BOOL CALLBACK ResourceIndexToIdEnumProc(HMODULE hModule, LPCTSTR lpszType, LPTSTR lpszName, LONG_PTR lParam)
{
	ResourceIndexToIdEnumData &enum_data = *(ResourceIndexToIdEnumData*)lParam;
	
	if (++enum_data.index == enum_data.find_index)
	{
		enum_data.result = (int)lpszName;
		return FALSE; // Stop
	}
	return TRUE; // Continue
}

// L17: Find integer ID of resource from one-based index. i.e. IconNumber -> resource ID.
int ResourceIndexToId(HMODULE aModule, LPCTSTR aType, int aIndex)
{
	ResourceIndexToIdEnumData enum_data;
	enum_data.find_index = aIndex;
	enum_data.index = 0;
	enum_data.result = -1; // Return value of -1 indicates failure, since ID 0 may be valid.

	EnumResourceNames(aModule, aType, &ResourceIndexToIdEnumProc, (LONG_PTR)&enum_data);

	return enum_data.result;
}

// L17: Extract icon of the appropriate size from an executable (or compatible) file.
HICON ExtractIconFromExecutable(LPTSTR aFilespec, int aIconNumber, int aWidth, int aHeight)
{
	HICON hicon = NULL;

	// If the module is already loaded as an executable, LoadLibraryEx returns its handle.
	// Otherwise each call will receive its own handle to a data file mapping.
	HMODULE hdatafile = LoadLibraryEx(aFilespec, NULL, LOAD_LIBRARY_AS_DATAFILE);
	if (hdatafile)
	{
		int group_icon_id = (aIconNumber < 0 ? -aIconNumber : ResourceIndexToId(hdatafile, (LPCTSTR)RT_GROUP_ICON, aIconNumber ? aIconNumber : 1));

		HRSRC hres;
		HGLOBAL hresdata;
		LPVOID presdata;

		// MSDN indicates resources are unloaded when the *process* exits, but also states
		// that the pointer returned by LockResource is valid until the *module* containing
		// the resource is unloaded. Testing seems to indicate that unloading a module indeed
		// unloads or invalidates any resources it contains.
		if ((hres = FindResource(hdatafile, MAKEINTRESOURCE(group_icon_id), RT_GROUP_ICON))
			&& (hresdata = LoadResource(hdatafile, hres))
			&& (presdata = LockResource(hresdata)))
		{
			// LookupIconIdFromDirectoryEx seems to use whichever is larger of aWidth or aHeight,
			// so one or the other may safely be -1. However, since this behaviour is undocumented,
			// treat -1 as "same as other dimension":
			int icon_id = LookupIconIdFromDirectoryEx((PBYTE)presdata, TRUE, aWidth == -1 ? aHeight : aWidth, aHeight == -1 ? aWidth : aHeight, 0);
			if (icon_id
				&& (hres = FindResource(hdatafile, MAKEINTRESOURCE(icon_id), RT_ICON))
				&& (hresdata = LoadResource(hdatafile, hres))
				&& (presdata = LockResource(hresdata)))
			{
				hicon = CreateIconFromResourceEx((PBYTE)presdata, SizeofResource(hdatafile, hres), TRUE, 0x30000, 0, 0, 0);
			}
		}

		// Decrements the executable module's reference count or frees the data file mapping.
		FreeLibrary(hdatafile);
	}

	// L20: Fall back to ExtractIcon if the above method failed. This may work on some versions of Windows where
	// ExtractIcon supports 16-bit "executables" (such as ICL files) that cannot be loaded by LoadLibraryEx.
	// However, resource ID -1 is not supported, and if multiple icon sizes exist in the file, the first is used
	// rather than the most appropriate.
	if (!hicon)
		hicon = ExtractIcon(0, aFilespec, aIconNumber > 0 ? aIconNumber - 1 : aIconNumber < -1 ? aIconNumber : 0);

	return hicon;
}


HBITMAP IconToBitmap(HICON ahIcon, bool aDestroyIcon)
// Converts HICON to an HBITMAP that has ahIcon's actual dimensions.
// The incoming ahIcon will be destroyed if the caller passes true for aDestroyIcon.
// Returns NULL on failure, in which case aDestroyIcon will still have taken effect.
// If the icon contains any transparent pixels, they will be mapped to CLR_NONE within
// the bitmap so that the caller can detect them.
{
	if (!ahIcon)
		return NULL;

	HBITMAP hbitmap = NULL;  // Set default.  This will be the value returned.

	HDC hdc_desktop = GetDC(HWND_DESKTOP);
	HDC hdc = CreateCompatibleDC(hdc_desktop); // Don't pass NULL since I think that would result in a monochrome bitmap.
	if (hdc)
	{
		ICONINFO ii;
		if (GetIconInfo(ahIcon, &ii))
		{
			BITMAP icon_bitmap;
			// Find out how big the icon is and create a bitmap compatible with the desktop DC (not the memory DC,
			// since its bits per pixel (color depth) is probably 1.
			if (GetObject(ii.hbmColor, sizeof(BITMAP), &icon_bitmap)
				&& (hbitmap = CreateCompatibleBitmap(hdc_desktop, icon_bitmap.bmWidth, icon_bitmap.bmHeight))) // Assign
			{
				// To retain maximum quality in case caller needs to resize the bitmap we return, convert the
				// icon to a bitmap that matches the icon's actual size:
				HGDIOBJ old_object = SelectObject(hdc, hbitmap);
				if (old_object) // Above succeeded.
				{
					// Use DrawIconEx() vs. DrawIcon() because someone said DrawIcon() always draws 32x32
					// regardless of the icon's actual size.
					// If it's ever needed, this can be extended so that the caller can pass in a background
					// color to use in place of any transparent pixels within the icon (apparently, DrawIconEx()
					// skips over transparent pixels in the icon when drawing to the DC and its bitmap):
					RECT rect = {0, 0, icon_bitmap.bmWidth, icon_bitmap.bmHeight}; // Left, top, right, bottom.
					HBRUSH hbrush = CreateSolidBrush(CLR_DEFAULT);
					FillRect(hdc, &rect, hbrush);
					DeleteObject(hbrush);
					// Probably something tried and abandoned: FillRect(hdc, &rect, (HBRUSH)GetStockObject(NULL_BRUSH));
					DrawIconEx(hdc, 0, 0, ahIcon, icon_bitmap.bmWidth, icon_bitmap.bmHeight, 0, NULL, DI_NORMAL);
					// Debug: Find out properties of new bitmap.
					//BITMAP b;
					//GetObject(hbitmap, sizeof(BITMAP), &b);
					SelectObject(hdc, old_object); // Might be needed (prior to deleting hdc) to prevent memory leak.
				}
			}
			// It's our responsibility to delete these two when they're no longer needed:
			DeleteObject(ii.hbmColor);
			DeleteObject(ii.hbmMask);
		}
		DeleteDC(hdc);
	}
	ReleaseDC(HWND_DESKTOP, hdc_desktop);
	if (aDestroyIcon)
		DestroyIcon(ahIcon);
	return hbitmap;
}


// Lexikos: Used for menu icons on Windows Vista and later. Some similarities to IconToBitmap, maybe should be merged if IconToBitmap does not specifically need to create a device-dependent bitmap?
HBITMAP IconToBitmap32(HICON ahIcon, bool aDestroyIcon)
{
	// Get the icon's internal colour and mask bitmaps.
	// hbmColor is needed to measure the icon.
	// hbmMask is needed to generate an alpha channel if the icon does not have one.
	ICONINFO icon_info;
	if (!GetIconInfo(ahIcon, &icon_info))
		return NULL;

	HBITMAP hbitmap = NULL; // Set default in case of failure.

	// Get size of icon's internal bitmap.
	BITMAP icon_bitmap;
	if (GetObject(icon_info.hbmColor, sizeof(BITMAP), &icon_bitmap))
	{
		int width = icon_bitmap.bmWidth;
		int height = icon_bitmap.bmHeight;

		HDC hdc = CreateCompatibleDC(NULL);
		if (hdc)
		{
			BITMAPINFO bitmap_info = {0};
			// Set parameters for 32-bit bitmap. May also be used to retrieve bitmap data from the icon's mask bitmap.
			BITMAPINFOHEADER &bitmap_header = bitmap_info.bmiHeader;
			bitmap_header.biSize = sizeof(BITMAPINFOHEADER);
			bitmap_header.biWidth = width;
			bitmap_header.biHeight = height;
			bitmap_header.biBitCount = 32;
			bitmap_header.biPlanes = 1;

			// Create device-independent bitmap.
			UINT *bits;
			if (hbitmap = CreateDIBSection(hdc, &bitmap_info, 0, (void**)&bits, NULL, 0))
			{
				HGDIOBJ old_object = SelectObject(hdc, hbitmap);
				if (old_object) // Above succeeded.
				{
					// Draw icon onto bitmap. Use DrawIconEx because DrawIcon always draws a large (usually 32x32) icon!
					DrawIconEx(hdc, 0, 0, ahIcon, 0, 0, 0, NULL, DI_NORMAL);

					// May be necessary for bits to be in sync, according to documentation for CreateDIBSection:
					GdiFlush();

					// Calculate end of bitmap data.
					UINT *bits_end = bits + width*height;
					
					UINT *this_pixel; // Used in a few loops below.

					// Check for alpha data.
					bool has_nonzero_alpha = false;
					for (this_pixel = bits; this_pixel < bits_end; ++this_pixel)
					{
						if (*this_pixel >> 24)
						{
							has_nonzero_alpha = true;
							break;
						}
					}

					if (!has_nonzero_alpha)
					{
						// Get bitmap data from the icon's mask.
						UINT *mask_bits = (UINT*)_alloca(height*width*4);
						if (GetDIBits(hdc, icon_info.hbmMask, 0, height, (LPVOID)mask_bits, &bitmap_info, 0))
						{
							UINT *this_mask_pixel;
							// Use the icon's mask to generate alpha data.
							for (this_pixel = bits, this_mask_pixel = mask_bits; this_pixel < bits_end; ++this_pixel, ++this_mask_pixel)
							{
								if (*this_mask_pixel)
									*this_pixel = 0;
								else
									*this_pixel |= 0xff000000;
							}
						}
						else
						{
							// GetDIBits failed, simply make the bitmap opaque.
							for (this_pixel = bits; this_pixel < bits_end; ++this_pixel)
								*this_pixel |= 0xff000000;
						}
					}

					SelectObject(hdc, old_object);
				}
				else
				{
					// Probably very rare, but be on the safe side.
					DeleteObject(hbitmap);
					hbitmap = NULL;
				}
			}
			DWORD dwError = GetLastError();
			DeleteDC(hdc);
		}
	}

	// It's our responsibility to delete these two when they're no longer needed:
	DeleteObject(icon_info.hbmColor);
	DeleteObject(icon_info.hbmMask);

	if (aDestroyIcon)
		DestroyIcon(ahIcon);

	return hbitmap;
}


HRESULT MySetWindowTheme(HWND hwnd, LPCWSTR pszSubAppName, LPCWSTR pszSubIdList)
{
	// The library must be loaded dynamically, otherwise the app will not launch on OSes older than XP.
	// Theme DLL is normally available only on XP+, but an attempt to load it is made unconditionally
	// in case older OSes can ever have it.
	HRESULT hresult = !S_OK; // Set default as "failure".
	HINSTANCE hinstTheme = LoadLibrary(_T("uxtheme"));
	if (hinstTheme)
	{
		typedef HRESULT (WINAPI *MySetWindowThemeType)(HWND, LPCWSTR, LPCWSTR);
  		MySetWindowThemeType DynSetWindowTheme = (MySetWindowThemeType)GetProcAddress(hinstTheme, "SetWindowTheme");
		if (DynSetWindowTheme)
			hresult = DynSetWindowTheme(hwnd, pszSubAppName, pszSubIdList);
		FreeLibrary(hinstTheme);
	}
	return hresult;
}



//HRESULT MyEnableThemeDialogTexture(HWND hwnd, DWORD dwFlags)
//{
//	// The library must be loaded dynamically, otherwise the app will not launch on OSes older than XP.
//	// Theme DLL is normally available only on XP+, but an attempt to load it is made unconditionally
//	// in case older OSes can ever have it.
//	HRESULT hresult = !S_OK; // Set default as "failure".
//	HINSTANCE hinstTheme = LoadLibrary(_T("uxtheme"));
//	if (hinstTheme)
//	{
//		typedef HRESULT (WINAPI *MyEnableThemeDialogTextureType)(HWND, DWORD);
//  		MyEnableThemeDialogTextureType DynEnableThemeDialogTexture = (MyEnableThemeDialogTextureType)GetProcAddress(hinstTheme, "EnableThemeDialogTexture");
//		if (DynEnableThemeDialogTexture)
//			hresult = DynEnableThemeDialogTexture(hwnd, dwFlags);
//		FreeLibrary(hinstTheme);
//	}
//	return hresult;
//}



LPTSTR ConvertEscapeSequences(LPTSTR aBuf, LPTSTR aLiteralMap, bool aAllowEscapedSpace)
// Replaces any escape sequences in aBuf with their reduced equivalent.  For example, if aEscapeChar
// is accent, Each `n would become a literal linefeed.  aBuf's length should always be the same or
// lower than when the process started, so there is no chance of overflow.
{
	int i;
	for (i = 0; ; ++i)  // Increment to skip over the symbol just found by the inner for().
	{
		for (; aBuf[i] && aBuf[i] != g_EscapeChar; ++i);  // Find the next escape char.
		if (!aBuf[i]) // end of string.
			break;
		LPTSTR cp1 = aBuf + i + 1;
		switch (*cp1)
		{
			// Only lowercase is recognized for these:
			case 'a': *cp1 = '\a'; break;  // alert (bell) character
			case 'b': *cp1 = '\b'; break;  // backspace
			case 'f': *cp1 = '\f'; break;  // formfeed
			case 'n': *cp1 = '\n'; break;  // newline
			case 'r': *cp1 = '\r'; break;  // carriage return
			case 't': *cp1 = '\t'; break;  // horizontal tab
			case 'v': *cp1 = '\v'; break;  // vertical tab
			case 's': // space (not always allowed for backward compatibility reasons).
				if (aAllowEscapedSpace)
					*cp1 = ' ';
				//else do nothing extra, just let the standard action for unrecognized escape sequences.
				break;
		}
		// Replace escape-sequence with its single-char value.  This is done even if the pair isn't
		// a recognizable escape sequence (e.g. `? becomes ?), which is the Microsoft approach and
		// might not be a bad way of handling things. Below has a final +1 to include the terminator:
		tmemmove(aBuf + i, cp1, _tcslen(cp1) + 1);
		if (aLiteralMap)
			aLiteralMap[i] = 1;  // In the map, mark this char as literal.
	}
	return aBuf;
}



int FindNextDelimiter(LPCTSTR aBuf, TCHAR aDelimiter, int aStartIndex, LPCTSTR aLiteralMap)
// Returns the index of the next delimiter, taking into account quotes, parentheses, etc.
// If the delimiter is not found, returns the length of aBuf.
{
	TCHAR in_quotes = 0;
	int open_parens = 0;
	for (int mark = aStartIndex; ; ++mark)
	{
		if (aBuf[mark] == aDelimiter)
		{
			if (!in_quotes && open_parens <= 0 && !(aLiteralMap && aLiteralMap[mark]))
				// A delimiting comma other than one in a sub-statement or function.
				return mark;
			// Otherwise, its a quoted/literal comma or one in parentheses (such as function-call).
			continue;
		}
		switch (aBuf[mark])
		{
		case '"': 
		case '\'':
			if (!in_quotes)
 				in_quotes = aBuf[mark];
			else if (in_quotes == aBuf[mark] && !(aLiteralMap && aLiteralMap[mark]))
 				in_quotes = 0;
 			//else the other type of quote mark was used to begin this quoted string.
 			break;
		case '`':
			// If the caller passed non-NULL aLiteralMap, this must be a literal ` char.
			// Otherwise, caller wants it to escape quote marks inside strings, but not aDelimiter.
			if (!aLiteralMap && aBuf[mark+1] == in_quotes)
				++mark; // Skip this escape char and let the loop's increment skip the quote mark.
			//else this is an ordinary literal ` or an escape sequence we don't need to handle.
			break;
		case '(': // For our purpose, "(", "[" and "{" can be treated the same.
		case '[': // If they aren't balanced properly, a later stage will detect it.
		case '{': //
			if (!in_quotes) // Literal parentheses inside a quoted string should not be counted for this purpose.
				++open_parens;
			break;
		case ')':
		case ']':
		case '}':
			if (!in_quotes)
				--open_parens; // If this makes it negative, validation later on will catch the syntax error.
			break;
		case '\0':
			// Reached the end of the string without finding a delimiter.  Return the
			// index of the null-terminator since that's typically what the caller wants.
			return mark;
		//default: some other character; just have the loop skip over it.
		}
	}
}



bool IsStringInList(LPTSTR aStr, LPTSTR aList, bool aFindExactMatch)
// Checks if aStr exists in aList (which is a comma-separated list).
// If aStr is blank, aList must start with a delimiting comma for there to be a match.
{
	// Must use a temp. buffer because otherwise there's no easy way to properly match upon strings
	// such as the following:
	// if var in string,,with,,literal,,commas
	TCHAR buf[LINE_SIZE];
	TCHAR *next_field, *cp, *end_buf = buf + _countof(buf) - 1;

	// v1.0.48.01: Performance improved by Lexikos.
	for (TCHAR *this_field = aList; *this_field; this_field = next_field) // For each field in aList.
	{
		for (cp = buf, next_field = this_field; *next_field && cp < end_buf; ++cp, ++next_field) // For each char in the field, copy it over to temporary buffer.
		{
			if (*next_field == ',') // This is either a delimiter (,) or a literal comma (,,).
			{
				++next_field;
				if (*next_field != ',') // It's "," instead of ",," so treat it as the end of this field.
					break;
				// Otherwise it's ",," and next_field now points at the second comma; so copy that comma
				// over as a literal comma then continue copying.
			}
			*cp = *next_field;
		}

		// The end of this field has been reached (or reached the capacity of the buffer), so terminate the string
		// in the buffer.
		*cp = '\0';

		if (*buf) // It is possible for this to be blank only for the first field.  Example: if var in ,abc
		{
			if (aFindExactMatch)
			{
				if (!g_tcscmp(aStr, buf)) // Match found
					return true;
			}
			else // Substring match
				if (g_tcsstr(aStr, buf)) // Match found
					return true;
		}
		else // First item in the list is the empty string.
			if (aFindExactMatch) // In this case, this is a match if aStr is also blank.
			{
				if (!*aStr)
					return true;
			}
			else // Empty string is always found as a substring in any other string.
				return true;
	} // for()

	return false;  // No match found.
}



LPTSTR InStrAny(LPTSTR aStr, LPTSTR aNeedle[], int aNeedleCount, size_t &aFoundLen)
{
	// For each character in aStr:
	for ( ; *aStr; ++aStr)
		// For each needle:
		for (int i = 0; i < aNeedleCount; ++i)
			// For each character in this needle:
			for (LPTSTR needle_pos = aNeedle[i], str_pos = aStr; ; ++needle_pos, ++str_pos)
			{
				if (!*needle_pos)
				{
					// All characters in needle matched aStr at this position, so we've
					// found our string.  If this needle is empty, it implicitly matches
					// at the first position in the string.
					aFoundLen = needle_pos - aNeedle[i];
					return aStr;
				}
				// Otherwise, we haven't reached the end of the needle. If we've reached
				// the end of aStr, *str_pos and *needle_pos won't match, so the check
				// below will break out of the loop.
				if (*needle_pos != *str_pos)
					// Not a match: continue on to the next needle, or the next starting
					// position in aStr if this is the last needle.
					break;
			}
	// If the above loops completed without returning, no matches were found.
	return NULL;
}



#if defined(_MSC_VER) && defined(_DEBUG)
void OutputDebugStringFormat(LPCTSTR fmt, ...)
{
	CString sMsg;
	va_list ap;

	va_start(ap, fmt);
	sMsg.FormatV(fmt, ap);
	OutputDebugString(sMsg);
}
#endif
