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

#ifndef util_h
#define util_h

#include "stdafx.h" // pre-compiled headers
#include "defines.h"


#ifdef _WIN64
#define Exp32or64(a,b) (b)
#else
#define Exp32or64(a,b) (a)
#endif


#ifdef UNICODE
#define tmemcpy			wmemcpy
#define tmemmove		wmemmove
#define tmemset			wmemset
#define tmemcmp			wmemcmp
#define tmalloc(c)		((LPTSTR) malloc((c) << 1))
#define trealloc(p, c)	((LPTSTR) realloc((p), (c) << 1))
#define talloca(c)		((LPTSTR) _alloca((c) << 1))
#else
#define tmemcpy			(char*)memcpy
#define tmemmove		memmove
#define tmemset			memset
#define tmemcmp			memcmp
#define tmalloc(c)		((LPTSTR) malloc(c))
#define trealloc(p, c)	((LPTSTR) realloc((p), (c)))
#define talloca(c)		((LPTSTR) _alloca(c))
#endif

#define IS_SPACE_OR_TAB(c) (c == ' ' || c == '\t')
#ifndef UNICODE
#define IS_SPACE_OR_TAB_OR_NBSP(c) (c == ' ' || c == '\t' || c == -96) // Use a negative to support signed chars.
#else
#define IS_SPACE_OR_TAB_OR_NBSP(c) IS_SPACE_OR_TAB(c) // wchar_t is unsigned
#endif

// v1.0.43.04: The following are macros to avoid crash bugs caused by improper casting, namely a failure to cast
// a signed char to UCHAR before promoting it to LPSTR, which crashes since CharLower/Upper would interpret
// such a high unsigned value as an address rather than a single char.
#define ltolower(ch) (TBYTE)CharLower((LPTSTR)(TBYTE)(ch))  // "L" prefix stands for "locale", like lstrcpy.
#define ltoupper(ch) (TBYTE)CharUpper((LPTSTR)(TBYTE)(ch))  // For performance, some callers don't want return value cast to char.


// Locale independent ctype (applied to the ASCII characters only)
// isctype/iswctype affects the some non-ASCII characters.
inline int cisctype(TBYTE c, int type)
{
	return (c & (~0x7F)) ? 0 : _isctype(c, type);
}

#define cisalpha(c)		cisctype(c, _ALPHA)
#define cisalnum(c)		cisctype(c, _ALPHA | _DIGIT)
#define cisdigit(c)		cisctype(c, _DIGIT)
#define cisxdigit(c)	cisctype(c, _HEX)
#define cisupper(c)		cisctype(c, _UPPER)
#define cislower(c)		cisctype(c, _LOWER)
#define cisprint(c)		cisctype(c, _ALPHA | _BLANK | _DIGIT | _PUNCT)
#define cisspace(c)		cisctype(c, _SPACE)

// The results of toupper/tolower are implementations dependent (see below), though the test results are OK in VS2008's CRT.
// MDSN: In order for toupper to give the expected results, __isascii and islower must both return nonzero.
// Linux (man page): The value returned is that of the converted letter, or c if the conversion was not possible. (CONFORMING TO C89, C99, 4.3BSD.)
inline TCHAR ctoupper(TBYTE c)
{
	return cislower(c) ? (c & ~0x20) : c;
}
inline TCHAR ctolower(TBYTE c)
{
	return cisupper(c) ? (c | 0x20) : c;
}

// Runtime setting dependent. "a" prefix stand for AutoHotkey.
#define aisalpha(c)	((int)((::g->StringCaseSense == SCS_INSENSITIVE_LOCALE) ? IsCharAlpha(c) : cisalpha(c)))
#define aisalnum(c)	((int)((::g->StringCaseSense == SCS_INSENSITIVE_LOCALE) ? IsCharAlphaNumeric(c) : cisalnum(c)))
#define aisupper(c)	((int)((::g->StringCaseSense == SCS_INSENSITIVE_LOCALE) ? IsCharUpper(c) : cisupper(c)))
#define aislower(c)	((int)((::g->StringCaseSense == SCS_INSENSITIVE_LOCALE) ? IsCharLower(c) : cislower(c)))

#define atoupper(c)	((::g->StringCaseSense == SCS_INSENSITIVE_LOCALE) ? ltoupper(c) : ctoupper(c))
#define atolower(c)	((::g->StringCaseSense == SCS_INSENSITIVE_LOCALE) ? ltolower(c) : ctolower(c))

// NOTE: MOVING THINGS OUT OF THIS FILE AND INTO util.cpp can hurt benchmarks by 10% or more, so be careful
// when doing so (even when the change seems inconsequential, it can impact benchmarks due to quirks of code
// generation and caching).


inline LPTSTR StrToTitleCase(LPTSTR aStr)
// Inline functions like the following probably don't actually get put inline by the compiler (since it knows
// it's not worth it).  However, changing this to be non-inline has a significant impact on benchmarks even
// though the function is never called by those benchmarks, probably due to coincidences in how the generated
// code gets cached in the CPU.  So it seems best not to mess with anything without making sure it doesn't
// drop the benchmarks significantly.
{
	if (!aStr) return aStr;
	LPTSTR aStr_orig = aStr;	
	for (bool convert_next_alpha_char_to_upper = true; *aStr; ++aStr)
	{
		if (IsCharAlpha(*aStr)) // Use this to better support chars from non-English languages.
		{
			if (convert_next_alpha_char_to_upper)
			{
				*aStr = ltoupper(*aStr);
				convert_next_alpha_char_to_upper = false;
			}
			else
				*aStr = ltolower(*aStr);
		}
		else
			if (_istspace(*aStr))
				convert_next_alpha_char_to_upper = true;
		// Otherwise, it's a digit, punctuation mark, etc. so nothing needs to be done.
	}
	return aStr_orig;
}



inline LPTSTR StrChrAny(LPTSTR aStr, LPTSTR aCharList)
// Returns the position of the first char in aStr that is of any one of the characters listed in aCharList.
// Returns NULL if not found.
// Update: Yes, this seems identical to strpbrk().  However, since the corresponding code would
// have to be added to the EXE regardless of which was used, there doesn't seem to be much
// advantage to switching (especially since if the two differ in behavior at all, things might
// get broken).  Another reason is the name "strpbrk()" is not as easy to remember.
{
	if (aStr == NULL || aCharList == NULL) return NULL;
	if (!*aStr || !*aCharList) return NULL;
	// Don't use strchr() because that would just find the first occurrence
	// of the first search-char, which is not necessarily the first occurrence
	// of *any* search-char:
	LPTSTR look_for_this_char;
	TCHAR char_being_analyzed;
	for (; *aStr; ++aStr)
		// If *aStr is any of the search char's, we're done:
		for (char_being_analyzed = *aStr, look_for_this_char = aCharList; *look_for_this_char; ++look_for_this_char)
			if (char_being_analyzed == *look_for_this_char)
				return aStr;  // Match found.
	return NULL; // No match.
}



inline LPCTSTR omit_leading_whitespace(LPCTSTR aBuf) // 10/17/2006: __forceinline didn't help significantly.
// While aBuf points to a whitespace, moves to the right and returns the first non-whitespace
// encountered.
{
	for (; IS_SPACE_OR_TAB(*aBuf); ++aBuf);
	return aBuf;
}
inline LPTSTR omit_leading_whitespace(LPTSTR aBuf)
{
	return (LPTSTR) omit_leading_whitespace((LPCTSTR) aBuf);
}


inline LPTSTR omit_leading_any(LPTSTR aBuf, LPTSTR aOmitList, size_t aLength)
// Returns the address of the first character in aBuf that isn't a member of aOmitList.
// But no more than aLength characters of aBuf will be considered.  If aBuf is composed
// entirely of omitted characters, the address of the char after the last char in the
// string will returned (that char will be the zero terminator unless aLength explicitly
// caused only part of aBuf to be considered).
{
	LPTSTR cp;
	for (size_t i = 0; i < aLength; ++i, ++aBuf)
	{
		// Check if the current char is a member of the omitted-char list:
		for (cp = aOmitList; *cp; ++cp)
			if (*aBuf == *cp) // Match found.
				break;
		if (!*cp) // No match found, so this character is not omitted, thus we immediately return it's position.
			return aBuf;
	}
	// Since the above didn't return, aBuf is the position of the zero terminator or (if aLength
	// indicated only a substring) the position of the char after the last char in the substring.
	return aBuf;
}



inline LPTSTR omit_trailing_whitespace(LPTSTR aBuf, LPTSTR aBuf_marker)
// aBuf_marker must be a position in aBuf (to the right of it).
// Starts at aBuf_marker and keeps moving to the left until a non-whitespace
// char is encountered.  Returns the position of that char.
{
	for (; aBuf_marker > aBuf && IS_SPACE_OR_TAB(*aBuf_marker); --aBuf_marker);
	return aBuf_marker;  // Can equal aBuf.
}



inline size_t omit_trailing_any(LPTSTR aBuf, LPTSTR aOmitList, LPTSTR aBuf_marker)
// aBuf_marker must be a position in aBuf (to the right of it).
// Starts at aBuf_marker and keeps moving to the left until a char that isn't a member
// of aOmitList is found.  The length of the remaining substring is returned.
// That length will be zero if the string consists entirely of omitted characters.
{
	LPTSTR cp;
	for (; aBuf_marker > aBuf; --aBuf_marker)
	{
		// Check if the current char is a member of the omitted-char list:
		for (cp = aOmitList; *cp; ++cp)
			if (*aBuf_marker == *cp) // Match found.
				break;
		if (!*cp) // No match found, so this character is not omitted, thus we immediately return.
			return (aBuf_marker - aBuf) + 1; // The length of the string when trailing chars are omitted.
	}
	// Since the above didn't return, aBuf_marker is now equal to aBuf.  If this final character is itself
	// a member of the omitted-list, the length returned will be zero.  Otherwise it will be 1:
	for (cp = aOmitList; *cp; ++cp)
		if (*aBuf_marker == *cp) // Match found.
			return 0;
	return 1;
}



inline size_t ltrim(LPTSTR aStr, size_t aLength = -1)
// Caller must ensure that aStr is not NULL.
// v1.0.25: Returns the length if it was discovered as a result of the operation, or aLength otherwise.
// NOTE: THIS VERSION trims only tabs and spaces.  It specifically avoids
// trimming newlines because some callers want to retain those.
{
	if (!*aStr) return 0;
	LPTSTR ptr;
	// Find the first non-whitespace char (which might be the terminator):
	for (ptr = aStr; IS_SPACE_OR_TAB(*ptr); ++ptr); // Self-contained loop.
	// v1.0.25: If no trimming needed, don't do the memmove.  This seems to make a big difference
	// in the performance of critical sections of the program:
	size_t offset;
	if (offset = ptr - aStr) // Assign.
	{
		if (aLength == -1)
			aLength = _tcslen(ptr); // Set aLength as new/trimmed length, for use below and also as the return value.
		else // v1.0.25.05 bug-fix: Must adjust the length provided by caller to reflect what we did here.
			aLength -= offset;
		tmemmove(aStr, ptr, aLength + 1); // +1 to include the '\0'.  memmove() permits source & dest to overlap.
	}
	return aLength; // This will return -1 if the block above didn't execute and caller didn't specify the length.
}

inline size_t rtrim(LPTSTR aStr, size_t aLength = -1)
// Caller must ensure that aStr is not NULL.
// To improve performance, caller may specify a length (e.g. when it is already known).
// v1.0.25: Always returns the new length of the string.
// NOTE: THIS VERSION trims only tabs and spaces.  It specifically avoids trimming newlines because
// some callers want to retain those.
{
	if (!*aStr) return 0; // The below relies upon this check having been done.
	// It's done this way in case aStr just happens to be address 0x00 (probably not possible
	// on Intel & Intel-clone hardware) because otherwise --cp would decrement, causing an
	// underflow since pointers are probably considered unsigned values, which would
	// probably cause an infinite loop.  Extremely unlikely, but might as well try
	// to be thorough:
	if (aLength == -1)
		aLength = _tcslen(aStr); // Set aLength for use below and also as the return value.
	for (LPTSTR cp = aStr + aLength - 1; ; --cp, --aLength)
	{
		if (!IS_SPACE_OR_TAB(*cp))
		{
			cp[1] = '\0';
			return aLength;
		}
		// Otherwise, it is a space or tab...
		if (cp == aStr) // ... and we're now at the first character of the string...
		{
			if (IS_SPACE_OR_TAB(*cp)) // ... and that first character is also a space or tab...
			{
				*cp = '\0'; // ... so the entire string is made empty...
				return 0; // Fix for v1.0.39: Must return 0 not aLength in this case.
			}
			return aLength; // ... and we return in any case.
		}
		// else it's a space or tab, and there are still more characters to check.  Let the loop
		// do its decrements.
	}
}

inline void rtrim_literal(LPTSTR aStr, TCHAR aLiteralMap[])
// Caller must ensure that aStr is not NULL.
// NOTE: THIS VERSION trims only tabs and spaces which aren't marked as literal (so not "`t" or "` ").
// It specifically avoids trimming newlines because some callers want to retain those.
{
	if (!*aStr) return; // The below relies upon this check having been done.
	// It's done this way in case aStr just happens to be address 0x00 (probably not possible
	// on Intel & Intel-clone hardware) because otherwise --cp would decrement, causing an
	// underflow since pointers are probably considered unsigned values, which would
	// probably cause an infinite loop.  Extremely unlikely, but might as well try
	// to be thorough:
	for (size_t last = _tcslen(aStr) - 1; ; --last)
	{
		if (!IS_SPACE_OR_TAB(aStr[last]) || aLiteralMap[last]) // It's not a space or tab, or it's a literal one.
		{
			aStr[last + 1] = '\0';
			return;
		}
		// Otherwise, it is a space or tab...
		if (last == 0) // ... and we're now at the first character of the string...
		{
			if (IS_SPACE_OR_TAB(aStr[last])) // ... and that first character is also a space or tab...
				*aStr = '\0'; // ... so the entire string is made empty.
			return; // ... and we return in any case.
		}
		// else it's a space or tab, and there are still more characters to check.  Let the loop
		// do its decrements.
	}
}

inline size_t rtrim_with_nbsp(LPTSTR aStr, size_t aLength = -1)
// Returns the new length of the string.
// Caller must ensure that aStr is not NULL.
// To improve performance, caller may specify a length (e.g. when it is already known).
// Same as rtrim but also gets rid of those annoying nbsp (non breaking space) chars that sometimes
// wind up on the clipboard when copied from an HTML document, and thus get pasted into the text
// editor as part of the code (such as the sample code in some of the examples).
{
	if (!*aStr) return 0; // The below relies upon this check having been done.
	if (aLength == -1)
		aLength = _tcslen(aStr); // Set aLength for use below and also as the return value.
	for (LPTSTR cp = aStr + aLength - 1; ; --cp, --aLength)
	{
		if (!IS_SPACE_OR_TAB_OR_NBSP(*cp))
		{
			cp[1] = '\0';
			return aLength;
		}
		if (cp == aStr)
		{
			if (IS_SPACE_OR_TAB_OR_NBSP(*cp)) // ... and that first character is also a space or tab...
			{
				*cp = '\0'; // ... so the entire string is made empty...
				return 0; // Fix for v1.0.39: Must return 0 not aLength in this case.
			}
			return aLength; // ... and we return in any case.
		}
	}
}

inline size_t trim(LPTSTR aStr, size_t aLength = -1)
// Caller must ensure that aStr is not NULL.
// Returns new length of aStr.
// To improve performance, caller may specify a length (e.g. when it is already known).
// NOTE: THIS VERSION trims only tabs and spaces.  It specifically avoids
// trimming newlines because some callers want to retain those.
{
	aLength = ltrim(aStr, aLength);  // It may return -1 to indicate that it still doesn't know the length.
    return rtrim(aStr, aLength);
	// v1.0.25: rtrim() always returns the new length of the string.
}

inline size_t strip_trailing_backslash(LPTSTR aPath)
// Removes any backslash (if there is one).
// Returns length of the new string to allow some callers to avoid another strlen() call.
{
	size_t length = _tcslen(aPath);
	if (!length) // Below relies on this check having been done to prevent underflow.
		return length;
	LPTSTR cp = aPath + length - 1;
	if (*cp == _T('\\'))
	{
		*cp = '\0';
		return length - 1;
	}
	// Otherwise there no slash to remove, so return the current length.
	return length;
}



#define IS_IDENTIFIER_CHAR(c) (cisalnum(c) || (c) == '_' || ((UINT)(c) > 0x7F))
template<typename T> inline T find_identifier_end(T aBuf)
// Locates the next character which is not valid in an identifier (var, func, or obj.key name).
{
	while (IS_IDENTIFIER_CHAR(*aBuf)) ++aBuf;
	return aBuf;
}




// Transformation is the same in either direction because the end bytes are swapped
// and the middle byte is left as-is:
#define bgr_to_rgb(aBGR) rgb_to_bgr(aBGR)
inline COLORREF rgb_to_bgr(DWORD aRGB)
// Fancier methods seem prone to problems due to byte alignment or compiler issues.
{
	return RGB(GetBValue(aRGB), GetGValue(aRGB), GetRValue(aRGB));
}



inline bool IsHex(LPCTSTR aBuf) // 10/17/2006: __forceinline worsens performance, but physically ordering it near ATOI64() [via /ORDER] boosts by 3.5%.
// Note: AHK support for hex ints reduces performance by only 10% for decimal ints, even in the tightest of math loops.
{
	// For whatever reason, omit_leading_whitespace() benches consistently faster (albeit slightly) than
	// the same code put inline (confirmed again on 10/17/2006, though the difference is hardly anything):
	//for (; IS_SPACE_OR_TAB(*aBuf); ++aBuf);
	aBuf = omit_leading_whitespace(aBuf); // i.e. caller doesn't have to have ltrimmed.
	if (!*aBuf)
		return false;
	if (*aBuf == '-' || *aBuf == '+')
		++aBuf;
	// The "0x" prefix must be followed by at least one hex digit, otherwise it's not considered hex:
	#define IS_HEX(buf) (*buf == '0' && (*(buf + 1) == 'x' || *(buf + 1) == 'X') && isxdigit(*(buf + 2)))
	return IS_HEX(aBuf);
}



// As of v1.0.30, ATOI(), ITOA() and the other related functions below are no longer macros
// because there are too many places where something like ATOI(++cp) is done, which would be a
// bug if not caught since cp would be incremented more than once if the macro referred to that
// arg more than once.  In addition, a non-comprehensive, simple benchmark shows that the
// macros don't perform any better anyway, probably in part because there are many times when
// something like ArgToInt(1) is called, which forces the ARG1 macro to be expanded two or more
// times within ATOI (when it was a macro).  So for now, the below are declared as inline.
// However, it seems that the compiler chooses not to make them truly inline, which as it
// turns out is probably the right decision since a simple benchmark shows that even with
// __forceinline in effect for all of them (which is confirmed to actually force inline),
// the performance isn't any better.

inline __int64 ATOI64(LPCTSTR buf)
{
	// See ATOI() for the reason IsHex() is used.
	return IsHex(buf) ? _tcstoi64(buf, NULL, 16) : _ttoi64(buf);  // _atoi64() has superior performance, so use it when possible.
}

inline unsigned __int64 ATOU64(LPCTSTR buf)
{
	return _tcstoui64(buf, NULL, IsHex(buf) ? 16 : 10);
}

inline int ATOI(LPCTSTR buf)
{
	// Below has been updated because values with leading zeros were being interpreted as
	// octal, which is undesirable.
	// Formerly: #define ATOI(buf) strtol(buf, NULL, 0) // Use zero as last param to support both hex & dec.
	return IsHex(buf) ? _tcstol(buf, NULL, 16) : _ttoi(buf); // atoi() has superior performance, so use it when possible.
}

// v1.0.38.01: Make ATOU a macro that refers to ATOI64() to improve performance (takes advantage of _atoi64()
// being considerably faster than strtoul(), at least when the number is non-hex).  This relies on the fact
// that ATOU() and (UINT)ATOI64() produce the same result due to the way casting works.  For example:
// ATOU("-1") == (UINT)ATOI64("-1")
// ATOU("-0xFFFFFFFF") == (UINT)ATOI64("-0xFFFFFFFF")
#define ATOU(buf) (UINT)ATOI64(buf)
//inline unsigned long ATOU(char *buf)
//{
//	// As a reminder, strtoul() also handles negative numbers.  For example, ATOU("-1") is
//	// 4294967295 (0xFFFFFFFF) and ATOU("-2") is 4294967294.
//	return strtoul(buf, NULL, IsHex(buf) ? 16 : 10);
//}

inline double ATOF(LPCTSTR buf)
// Unlike some Unix versions of strtod(), the VC++ version does not seem to handle hex strings
// such as "0xFF" automatically.  So this macro must check for hex because some callers rely on that.
// Also, it uses _strtoi64() vs. strtol() so that more of a double's capacity can be utilized:
{
	return IsHex(buf) ? (double)_tcstoi64(buf, NULL, 16) : _tstof(buf);
}

#define ITOA(value, buf)	_itot(value, buf, 10)
#define ITOA64(value, buf)	_i64tot(value, buf, 10)
#define UTOA(value, buf)	_ultot(value, buf, 10)
#define UTOA64(value, buf)	_ui64tot(value, buf, 10)
#ifdef _WIN64
#define UPTRTOA UTOA64
#else
#define UPTRTOA UTOA
#endif


inline LPTSTR HwndToString(HWND aHwnd, LPTSTR aBuf)
{
	aBuf[0] = '0';
	aBuf[1] = 'x';
	// Use _ultot for performance on 32-bit systems and _ui64tot on 64-bit systems in case it's
	// possible for HWNDs to have non-zero upper 32-bits:
	Exp32or64(_ultot,_ui64tot)((size_t)aHwnd, aBuf + 2, 16);
	return aBuf;
}


//inline LPTSTR tcscatmove(LPTSTR aDst, LPCTSTR aSrc)
//// Same as strcat() but allows aSrc and aDst to overlap.
//// Unlike strcat(), it doesn't return aDst.  Instead, it returns the position
//// in aDst where aSrc was appended.
//{
//	if (!aDst || !aSrc || !*aSrc) return aDst;
//	LPTSTR aDst_end = aDst + _tcslen(aDst);
//	return (LPTSTR)memmove(aDst_end, aSrc, (_tcslen(aSrc) + 1) * sizeof(TCHAR));  // Add 1 to include aSrc's terminator.
//}



// v1.0.43.03: The following macros support the new "StringCaseSense Locale" setting.  This setting performs
// 1 to 10 times slower for most things, but has the benefit of seeing characters like ä and Ä as identical
// when insensitive.  MSDN implies that lstrcmpi() is the same as:
//     CompareString(LOCALE_USER_DEFAULT, NORM_IGNORECASE, ...)
// Note that when MSDN talks about the "word sort" vs. "string sort", it does not mean that strings like
// "co-op" and "co-op" are considered equal.  Instead, they are considered closer together than the traditional
// string sort would see them, so that they wind up together in a sorted list.
// And both of them benchmark the same, so lstrcmpi is now used here and in various other places throughout
// the program when the new locale-case-insensitive mode is in effect.
#define tcscmp2(str1, str2, string_case_sense) ((string_case_sense) == SCS_INSENSITIVE ? _tcsicmp(str1, str2) \
	: ((string_case_sense) == SCS_INSENSITIVE_LOCALE ? lstrcmpi(str1, str2) : _tcscmp(str1, str2)))
#define g_tcscmp(str1, str2) tcscmp2(str1, str2, ::g->StringCaseSense)
// The most common mode is listed first for performance:
#define tcsstr2(haystack, needle, string_case_sense) ((string_case_sense) == SCS_INSENSITIVE ? tcscasestr(haystack, needle) \
	: ((string_case_sense) == SCS_INSENSITIVE_LOCALE ? lstrcasestr(haystack, needle) : _tcsstr(haystack, needle)))
#define g_tcsstr(haystack, needle) tcsstr2(haystack, needle, ::g->StringCaseSense)
// For the following, caller must ensure that len1 and len2 aren't beyond the terminated length of the string
// because CompareString() might not stop at the terminator when a length is specified.  Also, CompareString()
// returns 0 on failure, but failure occurs only when parameter/flag is invalid, which should never happen in
// this case.
#define lstrcmpni(str1, len1, str2, len2) (CompareString(LOCALE_USER_DEFAULT, NORM_IGNORECASE, str1, (int)(len1), str2, (int)(len2)) - 2) // -2 for maintainability


// The following macros simplify and make consistent the calls to MultiByteToWideChar().
// MSDN implies that passing -1 for cbMultiByte is the most typical and secure usage because it ensures
// that the output is null-terminated: "the resulting wide character string has a null terminator, and the
// length returned by the function includes the terminating null character."
//
// I couldn't find any info on when MB_PRECOMPOSED is needed (if ever).  It's the default anyway,
// which implies that passing zero (which is quite common in many examples I've seen) is essentially
// the same as passing MB_PRECOMPOSED.  However, some modes such as CP_UTF8 should never use MB_PRECOMPOSED
// or the function will fail.
//
// #1: FROM ANSI TO UNICODE (UTF-16).  dest_size_in_wchars includes the terminator.
// From looking at the source to mbstowcs(), it might be faster when the "C" locale is in effect (which is
// the default in the absence of setlocale()) than MultiByteToWideChar() depending on how the latter is
// implemented. This is because mbstowcs() simply casts the characters to (wchar_t)(unsigned char) without
// any other translation at all.  Although that behavior is probably identical to MultiByteToWideChar(CP_ACP...),
// it's not completely certain -- so it seems best to stick with MultiByteToWideChar() for consistency
// (also, avoiding mbstowcs slightly reduces code size).  If there's ever a case where performance is
// important, create a simple casting loop (see mbstowcs.c for an example) that converts source to dest,
// and test if it performs significantly better than MultiByteToWideChar(CP_ACP...).
#define ToWideChar(source, dest, dest_size_in_wchars) MultiByteToWideChar(CP_ACP, 0, source, -1, dest, dest_size_in_wchars)
//
// #2: FROM UTF-8 TO UNICODE (UTF-16). dest_size_in_wchars includes the terminator.  MSDN: "For UTF-8, dwFlags must be set to either 0 or MB_ERR_INVALID_CHARS. Otherwise, the function fails with ERROR_INVALID_FLAGS."
#define UTF8ToWideChar(source, dest, dest_size_in_wchars) MultiByteToWideChar(CP_UTF8, 0, source, -1, dest, dest_size_in_wchars)
//
// #3: FROM UNICODE (UTF-16) TO UTF-8. dest_size_in_bytes includes the terminator.
#define WideCharToUTF8(source, dest, dest_size_in_bytes) WideCharToMultiByte(CP_UTF8, 0, source, -1, dest, dest_size_in_bytes, NULL, NULL)

#define UTF8StrLen(str, cch) MultiByteToWideChar(CP_UTF8, 0, (str), (cch), NULL, 0)
#define WideUTF8StrLen(str, cch) WideCharToMultiByte(CP_UTF8, 0, (str), (cch), NULL, 0, NULL, NULL)

#ifdef UNICODE
#define PosToUTF8Pos              WideUTF8StrLen
#define LenToUTF8Len(str,pos,len) WideUTF8StrLen(LPCWSTR(str)+int(pos),len)
#define UTF8PosToPos              UTF8StrLen
#define UTF8LenToLen(str,pos,len) UTF8StrLen(LPCSTR(str)+int(pos),len)

inline char* WideToUTF8(LPCWSTR str){
	int buf_len = WideCharToUTF8(str, NULL, 0);
	LPSTR buf = (LPSTR) malloc(buf_len);
	if (buf) WideCharToUTF8(str, buf, buf_len);
	return buf;
}
inline LPTSTR UTF8ToWide(LPCSTR str){
	int buf_len = UTF8ToWideChar(str, NULL, 0);
	LPTSTR buf = (LPTSTR) tmalloc(buf_len);
	if (buf) UTF8ToWideChar(str, buf, buf_len);
	return buf;
}
#endif

#ifdef UNICODE
#define UorA(u,a)      (u)
//#define TPosToUTF8Pos  PosToUTF8Pos
//#define TLenToUTF8Len  LenToUTF8Len
//#define UTF8PosToTPos  UTF8PosToPos
//#define UTF8LenToTLen  UTF8LenToLen
#define ToUnicodeOrAsciiEx(wVirtKey, wScanCode, lpKeyState, pszBuff, wFlags, dwhkl) \
	ToUnicodeEx((wVirtKey), (wScanCode), (lpKeyState), (LPWSTR)(pszBuff), 2, (wFlags), (dwhkl))
#else
#define UorA(u,a)            (a)
//#define TPosToUTF8Pos(a,b)   (b)
//#define TLenToUTF8Len(a,b,c) (c)
//#define UTF8PosToTPos(a,b)   (b)
//#define UTF8LenToTLen(a,b,c) (c)
#define ToUnicodeOrAsciiEx(wVirtKey, wScanCode, lpKeyState, pszBuff, wFlags, dwhkl) \
	ToAsciiEx((wVirtKey), (wScanCode), (lpKeyState), (LPWORD)(pszBuff), (wFlags), (dwhkl))
#endif

// v1.0.44.03: Callers now use the following macro rather than the old approach.  However, this change
// is meaningful only to people who use more than one keyboard layout.  In the case of hotstrings:
// It seems that the vast majority of them would want the Hotstring monitoring to adhere to the active
// window's current keyboard layout rather than the script's.  This change is somewhat less certain to
// be desirable unconditionally for the Input command (especially invisible/non-V-option Inputs); but it 
// seems best to use the same approach to avoid calling ToAsciiEx() more than once in cases where a
// script has hotstrings and also uses the Input command. Calling ToAsciiEx() twice in such a case would
// be likely to aggravate its side effects with dead keys as described at length in the hook/Input code).
#define Get_active_window_keybd_layout \
	HWND active_window;\
	HKL active_window_keybd_layout = GetKeyboardLayout((active_window = GetForegroundWindow())\
		? GetWindowThreadProcessId(active_window, NULL) : 0); // When no foreground window, the script's own layout seems like the safest default.

#define FONT_POINT(hdc, p) (-MulDiv(p, GetDeviceCaps(hdc, LOGPIXELSY), 72))
#define DATE_FORMAT_LENGTH 14 // "YYYYMMDDHHMISS"
#define IS_LEAP_YEAR(year) ((year) % 4 == 0 && ((year) % 100 != 0 || (year) % 400 == 0))

int GetYDay(int aMon, int aDay, bool aIsLeapYear);
int GetISOWeekNumber(LPTSTR aBuf, int aYear, int aYDay, int aWDay);
ResultType YYYYMMDDToFileTime(LPTSTR aYYYYMMDD, FILETIME &aFileTime);
DWORD YYYYMMDDToSystemTime2(LPTSTR aYYYYMMDD, SYSTEMTIME *aSystemTime);
ResultType YYYYMMDDToSystemTime(LPTSTR aYYYYMMDD, SYSTEMTIME &aSystemTime, bool aDoValidate);
LPTSTR FileTimeToYYYYMMDD(LPTSTR aBuf, FILETIME &aTime, bool aConvertToLocalTime = false);
LPTSTR SystemTimeToYYYYMMDD(LPTSTR aBuf, SYSTEMTIME &aTime);
__int64 YYYYMMDDSecondsUntil(LPTSTR aYYYYMMDDStart, LPTSTR aYYYYMMDDEnd, bool &aFailed);
__int64 FileTimeSecondsUntil(FILETIME *pftStart, FILETIME *pftEnd);

SymbolType IsNumeric(LPCTSTR aBuf, BOOL aAllowNegative = false // BOOL vs. bool might squeeze a little more performance out of this frequently-called function.
	, BOOL aAllowAllWhitespace = true, BOOL aAllowFloat = false, BOOL aAllowImpure = false);

void strlcpy(LPSTR aDst, LPCSTR aSrc, size_t aDstSize);
void wcslcpy(LPWSTR aDst, LPCWSTR aSrc, size_t aDstSize);
#ifdef UNICODE
#define tcslcpy wcslcpy
#else
#define tcslcpy strlcpy
#endif
int sntprintf(LPTSTR aBuf, int aBufSize, LPCTSTR aFormat, ...);
int sntprintfcat(LPTSTR aBuf, int aBufSize, LPCTSTR aFormat, ...);
// Not currently used by anything, so commented out to possibly reduce code size:
//int tcslcmp (LPTSTR aBuf1, LPTSTR aBuf2, UINT aLength1 = UINT_MAX, UINT aLength2 = UINT_MAX);
int tcslicmp(LPTSTR aBuf1, LPTSTR aBuf2, size_t aLength1 = -1, size_t aLength2 = -1);
LPTSTR tcsrstr(LPTSTR aStr, size_t aStr_length, LPCTSTR aPattern, StringCaseSenseType aStringCaseSense, int aOccurrence = 1);
LPTSTR lstrcasestr(LPCTSTR phaystack, LPCTSTR pneedle);
LPTSTR tcscasestr (LPCTSTR phaystack, LPCTSTR pneedle);
UINT StrReplace(LPTSTR aHaystack, LPTSTR aOld, LPTSTR aNew, StringCaseSenseType aStringCaseSense
	, UINT aLimit = UINT_MAX, size_t aSizeLimit = -1, LPTSTR *aDest = NULL, size_t *aHaystackLength = NULL);
size_t PredictReplacementSize(ptrdiff_t aLengthDelta, int aReplacementCount, int aLimit, size_t aHaystackLength
	, size_t aCurrentLength, size_t aEndOffsetOfCurrMatch);
LPTSTR TranslateLFtoCRLF(LPTSTR aString);
bool DoesFilePatternExist(LPTSTR aFilePattern, DWORD *aFileAttr = NULL);
#ifdef _DEBUG
	ResultType FileAppend(LPTSTR aFilespec, LPTSTR aLine, bool aAppendNewline = true);
#endif
LPTSTR ConvertFilespecToCorrectCase(LPTSTR aFullFileSpec);
LPTSTR FileAttribToStr(LPTSTR aBuf, DWORD aAttr);
unsigned __int64 GetFileSize64(HANDLE aFileHandle);
LPTSTR GetLastErrorText(LPTSTR aBuf, int aBufSize, bool aUpdateLastError = false);
void AssignColor(LPTSTR aColorName, COLORREF &aColor, HBRUSH &aBrush);
COLORREF ColorNameToBGR(LPTSTR aColorName);
HRESULT MySetWindowTheme(HWND hwnd, LPCWSTR pszSubAppName, LPCWSTR pszSubIdList);
//HRESULT MyEnableThemeDialogTexture(HWND hwnd, DWORD dwFlags);
LPTSTR ConvertEscapeSequences(LPTSTR aBuf, LPTSTR aLiteralMap, bool aAllowEscapedSpace = false);
int FindNextDelimiter(LPCTSTR aBuf, TCHAR aDelimiter = ',', int aStartIndex = 0, LPCTSTR aLiteralMap = NULL);
POINT CenterWindow(int aWidth, int aHeight);
bool FontExist(HDC aHdc, LPCTSTR aTypeface);
void ScreenToWindow(POINT &aPoint, HWND aHwnd);
void CoordToScreen(int &aX, int &aY, int aWhichMode);
void CoordToScreen(POINT &aPoint, int aWhichMode);
void GetVirtualDesktopRect(RECT &aRect);
BOOL IsProcess64Bit(HANDLE aHandle);
LPVOID AllocInterProcMem(HANDLE &aHandle, DWORD aSize, HWND aHwnd, DWORD aExtraAccess = 0);
void FreeInterProcMem(HANDLE aHandle, LPVOID aMem);

DWORD GetEnvVarReliable(LPCTSTR aEnvVarName, LPTSTR aBuf);
DWORD ReadRegString(HKEY aRootKey, LPTSTR aSubkey, LPTSTR aValueName, LPTSTR aBuf, DWORD aBufSize);

HBITMAP LoadPicture(LPTSTR aFilespec, int aWidth, int aHeight, int &aImageType, int aIconNumber
	, bool aUseGDIPlusIfAvailable);
HBITMAP IconToBitmap(HICON ahIcon, bool aDestroyIcon);
HBITMAP IconToBitmap32(HICON aIcon, bool aDestroyIcon); // Lexikos: Used for menu icons on Vista+. Creates a 32-bit (ARGB) device-independent bitmap from an icon.
int CALLBACK FontEnumProc(ENUMLOGFONTEX *lpelfe, NEWTEXTMETRICEX *lpntme, DWORD FontType, LPARAM lParam);
bool IsStringInList(LPTSTR aStr, LPTSTR aList, bool aFindExactMatch);
LPTSTR InStrAny(LPTSTR aStr, LPTSTR aNeedle[], int aNeedleCount, size_t &aFoundLen);

int ResourceIndexToId(HMODULE aModule, LPCTSTR aType, int aIndex); // L17: Find integer ID of resource from index. i.e. IconNumber -> resource ID.
HICON ExtractIconFromExecutable(LPTSTR aFilespec, int aIconNumber, int aWidth, int aHeight); // L17: Extract icon of the appropriate size from an executable (or compatible) file.

#if defined(_MSC_VER) && defined(_DEBUG)
void OutputDebugStringFormat(LPCTSTR fmt, ...); // put debug message to the "Output" panel of Visual Studio.
#endif

#endif
