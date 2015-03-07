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

#ifndef var_h
#define var_h

#include "defines.h"
#include "SimpleHeap.h"
#include "clipboard.h"
#include "util.h" // for strlcpy() & snprintf()
EXTERN_CLIPBOARD;
extern BOOL g_WriteCacheDisabledInt64;
extern BOOL g_WriteCacheDisabledDouble;

#define MAX_ALLOC_SIMPLE 64  // Do not decrease this much since it is used for the sizing of some built-in variables.
#define SMALL_STRING_LENGTH (MAX_ALLOC_SIMPLE - 1)  // The largest string that can fit in the above.
#define DEREF_BUF_EXPAND_INCREMENT (16 * 1024) // Reduced from 32 to 16 in v1.0.46.07 to reduce the memory utilization of deeply recursive UDFs.
#define ERRORLEVEL_NONE _T("0")
#define ERRORLEVEL_ERROR _T("1")
#define ERRORLEVEL_ERROR2 _T("2")

enum AllocMethod {ALLOC_NONE, ALLOC_SIMPLE, ALLOC_MALLOC};
enum VarTypes
{
  // The following must all be LOW numbers to avoid any realistic chance of them matching the address of
  // any function (namely a BIV_* function).
  VAR_ALIAS  // VAR_ALIAS must always have a non-NULL mAliasFor.  In other ways it's the same as VAR_NORMAL.  VAR_ALIAS is never seen because external users call Var::Type(), which automatically resolves ALIAS to some other type.
, VAR_NORMAL // Most variables, such as those created by the user, are this type.
, VAR_CLIPBOARD
, VAR_LAST_WRITABLE = VAR_CLIPBOARD  // Keep this in sync with any changes to the set of writable variables.
#define VAR_IS_READONLY(var) ((var).Type() > VAR_LAST_WRITABLE)
, VAR_CLIPBOARDALL // Must be read-only because it's not designed to be writable.
, VAR_BUILTIN
, VAR_LAST_TYPE = VAR_BUILTIN
};

typedef UCHAR VarTypeType;     // UCHAR vs. VarTypes to save memory.
typedef UCHAR AllocMethodType; // UCHAR vs. AllocMethod to save memory.
typedef UCHAR VarAttribType;   // Same.
typedef UINT_PTR VarSizeType;  // jackieku(2009-10-23): Change this to UINT_PTR to ensure its size is the same with a pointer.
#define VARSIZE_MAX ((VarSizeType) ~0)
#define VARSIZE_ERROR VARSIZE_MAX

class Var; // Forward declaration.
// #pragma pack(4) not used here because although it would currently save 4 bytes per VarBkp struct (28 vs. 32),
// it would probably reduce performance since VarBkp items are stored in contiguous array rather than a
// linked list (which would cause every other struct in the array to have an 8-byte member than stretches
// across two 8-byte regions in memory).
struct VarBkp // This should be kept in sync with any changes to the Var class.  See Var for comments.
{
	__int64 mContentsInt64; // 64-bit members kept at the top of the struct to reduce the chance that they'll span 2 vs. 1 64-bit regions.
	Var *mVar; // Used to save the target var to which these backed up contents will later be restored.
	char *mByteContents;
	union
	{
		VarSizeType mByteLength;
		Var *mAliasFor;
	};
	VarSizeType mByteCapacity;
	AllocMethodType mHowAllocated;
	VarAttribType mAttrib;
	VarTypeType mType;
	// Not needed in the backup:
	//bool mIsLocal;
	//TCHAR *mName;
};

#pragma warning(push)
#pragma warning(disable: 4995 4996)

// Concerning "#pragma pack" below:
// Default pack would otherwise be 8, which would cause the 64-bit mContentsInt64 member to increase the size
// of the struct from 20 to 32 (instead of 28).  Benchmarking indicates that there's no significant performance
// loss from doing this, perhaps because variables are currently stored in a linked list rather than an
// array. (In an array, having the struct size be a multiple of 8 would prevent every other struct in the array
// from having its 64-bit members span more than one 64-bit region in memory, which might reduce performance.)
#ifdef _WIN64
#pragma pack(push, 8)
#else
#pragma pack(push, 4) // 32-bit vs. 64-bit. See above.
#endif
typedef VarSizeType (* BuiltInVarType)(LPTSTR aBuf, LPTSTR aVarName);
class Var
{
private:
	// Keep VarBkp (above) in sync with any changes made to the members here.
	union // 64-bit members kept at the top of the struct to reduce the chance that they'll span 2 64-bit regions.
	{
		// Although the 8-byte members mContentsInt64 and mContentsDouble could be hung onto the struct
		// via a 4-byte-pointer, thus saving 4 bytes for each variable that never uses a binary number,
		// it doesn't seem worth it because the percentage of variables in typical scripts that will
		// acquire a cached binary number at some point seems likely to be high. A percentage of only
		// 50% would be enough to negate the savings because half the variables would consume 12 bytes
		// more than the version of AutoHotkey that has no binary-number caching, and the other half
		// would consume 4 more (due to the unused/empty pointer).  That would be an average of 8 bytes
		// extra; i.e. exactly the same as the 8 bytes used by putting the numbers directly into the struct.
		// In addition, there are the following advantages:
		// 1) Code less complicated, more maintainable, faster.
		// 2) Caching of binary numbers works even in recursive script functions.  By contrast, if the
		//    binary number were allocated on demand, recursive functions couldn't use caching because the
		//    memory from SimpleHeap could never be freed, thus producing a memory leak.
		// The main drawback is that some scripts are known to create a million variables or more, so the
		// extra 8 bytes per variable would increase memory load by 8+ MB (possibly with a boost in
		// performance if those variables are ever numeric).
		__int64 mContentsInt64;
		double mContentsDouble;
		IObject *mObject; // L31
	};
	union
	{
		char *mByteContents;
		LPTSTR mCharContents;
	};
	union
	{
		VarSizeType mByteLength;  // How much is actually stored in it currently, excluding the zero terminator.
		Var *mAliasFor;           // The variable for which this variable is an alias.
	};
	union
	{
		VarSizeType mByteCapacity; // In bytes.  Includes the space for the zero terminator.
		BuiltInVarType mBIV;
	};
	AllocMethodType mHowAllocated; // Keep adjacent/contiguous with the below to save memory.
	#define VAR_ATTRIB_BINARY_CLIP          0x01
	#define VAR_ATTRIB_OBJECT		        0x02 // mObject contains an object; mutually exclusive of the cache attribs.
	#define VAR_ATTRIB_UNINITIALIZED        0x04 // Var requires initialization before use.
	#define VAR_ATTRIB_CONTENTS_OUT_OF_DATE 0x08
	#define VAR_ATTRIB_HAS_VALID_INT64      0x10 // Cache type 1. Mutually exclusive of the other two.
	#define VAR_ATTRIB_HAS_VALID_DOUBLE     0x20 // Cache type 2. Mutually exclusive of the other two.
	#define VAR_ATTRIB_NOT_NUMERIC          0x40 // Cache type 3. Some sections might rely these being mutually exclusive.
	#define VAR_ATTRIB_CACHE_DISABLED       0x80 // If present, indicates that caching of the above 3 is disabled.
	#define VAR_ATTRIB_CACHE (VAR_ATTRIB_HAS_VALID_INT64 | VAR_ATTRIB_HAS_VALID_DOUBLE | VAR_ATTRIB_NOT_NUMERIC)
	#define VAR_ATTRIB_OFTEN_REMOVED (VAR_ATTRIB_CACHE | VAR_ATTRIB_BINARY_CLIP | VAR_ATTRIB_CONTENTS_OUT_OF_DATE)
	VarAttribType mAttrib;  // Bitwise combination of the above flags.
	#define VAR_GLOBAL			0x01
	#define VAR_LOCAL			0x02
	#define VAR_LOCAL_FUNCPARAM	0x10 // Indicates this local var is a function's parameter.  VAR_LOCAL_DECLARED should also be set.
	#define VAR_LOCAL_STATIC	0x20 // Indicates this local var retains its value between function calls.
	#define VAR_DECLARED		0x40 // Indicates this var was declared somehow, not automatic.
	#define VAR_SUPER_GLOBAL	0x80 // Indicates this global var should be visible in all functions.
	UCHAR mScope;  // Bitwise combination of the above flags.
	VarTypeType mType; // Keep adjacent/contiguous with the above due to struct alignment, to save memory.
	// Performance: Rearranging mType and the other byte-sized members with respect to each other didn't seem
	// to help performance.  However, changing VarTypeType from UCHAR to int did boost performance a few percent,
	// but even if it's not a fluke, it doesn't seem worth the increase in memory for scripts with many
	// thousands of variables.

	friend class Line; // For access to mBIV.
#ifdef CONFIG_DEBUGGER
	friend class Debugger;
#endif

	void UpdateBinaryInt64(__int64 aInt64, VarAttribType aAttrib = VAR_ATTRIB_HAS_VALID_INT64)
	// When caller doesn't include VAR_ATTRIB_CONTENTS_OUT_OF_DATE in aAttrib, CALLER MUST ENSURE THAT
	// mContents CONTAINS A PURE NUMBER; i.e. it mustn't contain something non-numeric at the end such as
	// 123abc (but trailing/leading whitespace is okay).  This is because users of the cached binary number
	// generally expect mContents to be an accurate reflection of that number.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);

		if (var.IsObject()) // L31: mObject will be overwritten below via the union, so release it now.
			var.ReleaseObject();

		var.mContentsInt64 = aInt64;
		var.mAttrib &= ~(VAR_ATTRIB_CACHE | VAR_ATTRIB_UNINITIALIZED); // But not VAR_ATTRIB_CONTENTS_OUT_OF_DATE because the caller specifies whether or not that gets added.
		var.mAttrib |= aAttrib; // Must be done prior to below. Indicate the type of binary number and whether VAR_ATTRIB_CONTENTS_OUT_OF_DATE is present.
		if (var.mAttrib & VAR_ATTRIB_CACHE_DISABLED) // Variables marked this way can't use either read or write caching.
		{
			var.UpdateContents(); // Update contents based on the new binary number just stored above. This call also removes the VAR_ATTRIB_CONTENTS_OUT_OF_DATE flag.
			var.mAttrib &= ~VAR_ATTRIB_CACHE; // Must be done after the above: Prevent the cached binary number from ever being used because this variable has been marked volatile (e.g. external changes to clipboard) and the cache can't be trusted.
		}
		else if (g_WriteCacheDisabledInt64 && (var.mAttrib & VAR_ATTRIB_HAS_VALID_INT64)
			|| g_WriteCacheDisabledDouble && (var.mAttrib & VAR_ATTRIB_HAS_VALID_DOUBLE))
		{
			if (var.mAttrib & VAR_ATTRIB_CONTENTS_OUT_OF_DATE) // For performance. See comments below.
				var.UpdateContents();
			// But don't remove VAR_ATTRIB_HAS_VALID_INT64/VAR_ATTRIB_HAS_VALID_DOUBLE because some of
			// our callers omit VAR_ATTRIB_CONTENTS_OUT_OF_DATE from aAttrib because they already KNOW
			// that var.mContents accurately represents the double or int64 in aInt64 (in such cases,
			// they also know that the precision of any floating point number in mContents matches the
			// precision/rounding that's in the double stored in aInt64).  In other words, unlike
			// VAR_ATTRIB_CACHE_DISABLED, only write-caching is disabled in the above cases (not read-caching).
			// This causes newly written numbers to be immediately written out to mContents so that the
			// SetFormat command works in realtime, for backward compatibility.  Also, even if the
			// new/incoming binary number matches the one already in the cache, MUST STILL write out
			// to mContents in case SetFormat is now different than it was before.
		}
	}

	void UpdateBinaryDouble(double aDouble, VarAttribType aAttrib = 0)
	// FOR WHAT GOES IN THIS SPOT, SEE IMPORTANT COMMENTS IN UpdateBinaryInt64().
	{
		// The type-casting below interprets the contents of aDouble as an __int64 without actually converting
		// from double to __int64.  Although the generated code isn't measurably smaller, hopefully the compiler
		// resolves it into something that performs better than a memcpy into a temporary variable.
		// Benchmarks show that the performance is at most a few percent worse than having code similar to
		// UpdateBinaryInt64() in here.
		UpdateBinaryInt64(*(__int64 *)&aDouble, aAttrib | VAR_ATTRIB_HAS_VALID_DOUBLE);
	}

	void UpdateContents() // Supports both VAR_NORMAL and VAR_CLIPBOARD.
	// Any caller who (prior to the call) stores a new cached binary number in the variable and also
	// sets VAR_ATTRIB_CONTENTS_OUT_OF_DATE must (after the call) remove VAR_ATTRIB_CACHE if the
	// variable has the VAR_ATTRIB_CACHE_DISABLED flag.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mAttrib & VAR_ATTRIB_CONTENTS_OUT_OF_DATE)
		{
			// THE FOLLOWING ISN'T NECESSARY BECAUSE THE ASSIGN() CALLS BELOW DO IT:
			//var.mAttrib &= ~VAR_ATTRIB_CONTENTS_OUT_OF_DATE;
			TCHAR value_string[MAX_NUMBER_SIZE];
			if (var.mAttrib & VAR_ATTRIB_HAS_VALID_INT64)
			{
				var.Assign(ITOA64(var.mContentsInt64, value_string)); // Return value currently not checked for this or the below.
				var.mAttrib |= VAR_ATTRIB_HAS_VALID_INT64; // Re-enable the cache because Assign() disables it (since all other callers want that).
			}
			else if (var.mAttrib & VAR_ATTRIB_HAS_VALID_DOUBLE)
			{
				// "%0.6f"; %f can handle doubles in MSVC++:
				var.Assign(value_string, sntprintf(value_string, _countof(value_string), g->FormatFloat, var.mContentsDouble));
				// In this case, read-caching should be disabled for scripts that use "SetFormat Float" because
				// they might rely on SetFormat having rounded floats off to FAR fewer decimal places (or
				// even to integers via "SetFormat, Float, 0").  Such scripts can use read-caching only when
				// mContents has been used to update the cache, not vice versa.  This restriction doesn't seem
				// to be necessary for "SetFormat Integer" because there should be no loss of precision when
				// integers are stored as hex vs. decimal:
				if (!g_WriteCacheDisabledDouble) // See comment above for why this is checked for float but not integer.
					var.mAttrib |= VAR_ATTRIB_HAS_VALID_DOUBLE; // Re-enable the cache because Assign() disables it (since all other callers want that).
			}
			//else nothing to update, which shouldn't happen in this block unless there's a flaw or bug somewhere.
		}
	}

	VarSizeType _CharLength() { return mByteLength / sizeof(TCHAR); }
	VarSizeType _CharCapacity() { return mByteCapacity / sizeof(TCHAR); }
public:
	// Testing shows that due to data alignment, keeping mType adjacent to the other less-than-4-size member
	// above it reduces size of each object by 4 bytes.
	TCHAR *mName;    // The name of the var.

	// sEmptyString is a special *writable* memory area for empty variables (those with zero capacity).
	// Although making it writable does make buffer overflows difficult to detect and analyze (since they
	// tend to corrupt the program's static memory pool), the advantages in maintainability and robustness
	// seem to far outweigh that.  For example, it avoids having to constantly think about whether
	// *Contents()='\0' is safe. The sheer number of places that's avoided is a great relief, and it also
	// cuts down on code size due to not having to always check Capacity() and/or create more functions to
	// protect from writing to read-only strings, which would hurt performance.
	// The biggest offender of buffer overflow in sEmptyString is DllCall, which happens most frequently
	// when a script forgets to call VarSetCapacity before passing a buffer to some function that writes a
	// string to it.  There is now some code there that tries to detect when that happens.
	static TCHAR sEmptyString[1]; // See above.

	VarSizeType Get(LPTSTR aBuf = NULL);
	ResultType AssignHWND(HWND aWnd);
	ResultType Assign(Var &aVar);
	ResultType Assign(ExprTokenType &aToken);
	static ResultType GetClipboardAll(Var *aOutputVar, void **aData, size_t *aDataSize);
	static ResultType SetClipboardAll(void *aData, size_t aDataSize);
	ResultType AssignClipboardAll();
	ResultType AssignBinaryClip(Var &aSourceVar);
	// Assign(char *, ...) has been break into four methods below.
	// This should prevent some mistakes, as characters and bytes are not interchangeable in the Unicode build.
	// Callers must make sure which one is the right method to call.
	ResultType AssignString(LPCTSTR aBuf = NULL, VarSizeType aLength = VARSIZE_MAX, bool aExactSize = false, bool aObeyMaxMem = true);
	inline ResultType Assign(LPCTSTR aBuf, VarSizeType aLength = VARSIZE_MAX, bool aExactSize = false, bool aObeyMaxMem = true)
	{
		ASSERT(aBuf); // aBuf shouldn't be NULL, use SetCapacity([length in bytes]) or AssignString(NULL, [length in characters]) instead.
		return AssignString(aBuf, aLength, aExactSize, aObeyMaxMem);
	}
	inline ResultType Assign()
	{
		return AssignString();
	}
	ResultType SetCapacity(VarSizeType aByteLength, bool aExactSize = false, bool aObeyMaxMem = true)
	{
#ifdef UNICODE
		return AssignString(NULL, (aByteLength >> 1) + (aByteLength & 1), aExactSize, aObeyMaxMem);
#else
		return AssignString(NULL, aByteLength, aExactSize, aObeyMaxMem);
#endif
	}

	ResultType AssignStringFromCodePage(LPCSTR aBuf, int aLength = -1, UINT aCodePage = CP_ACP);
	ResultType AssignStringFromUTF8(LPCSTR aBuf, int aLength = -1)
	{
		return AssignStringFromCodePage(aBuf, aLength, CP_UTF8);
	}
	ResultType AssignStringToCodePage(LPCWSTR aBuf, int aLength = -1, UINT aCodePage = CP_ACP, DWORD aFlags = WC_NO_BEST_FIT_CHARS, char aDefChar = '?');
	inline ResultType AssignStringW(LPCWSTR aBuf, int aLength = -1)
	{
#ifdef UNICODE
		// Pass aExactSize=true, aObeyMaxMem=false for consistency with AssignStringTo/FromCodePage/UTF8.
		// FileRead() relies on this to disobey #MaxMem:
		return AssignString(aBuf, aLength, true, false);
#else
		return AssignStringToCodePage(aBuf, aLength);
#endif
	}

	inline ResultType Assign(DWORD aValueToAssign) // For some reason, this function is actually faster when not __forceinline.
	{
		UpdateBinaryInt64(aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE|VAR_ATTRIB_HAS_VALID_INT64);
		return OK;
	}

	inline ResultType Assign(int aValueToAssign) // For some reason, this function is actually faster when not __forceinline.
	{
		UpdateBinaryInt64(aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE|VAR_ATTRIB_HAS_VALID_INT64);
		return OK;
	}

	inline ResultType Assign(__int64 aValueToAssign) // For some reason, this function is actually faster when not __forceinline.
	{
		UpdateBinaryInt64(aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE|VAR_ATTRIB_HAS_VALID_INT64);
		return OK;
	}

	inline ResultType Assign(VarSizeType aValueToAssign) // For some reason, this function is actually faster when not __forceinline.
	{
		UpdateBinaryInt64(aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE|VAR_ATTRIB_HAS_VALID_INT64);
		return OK;
	}

	inline ResultType Assign(double aValueToAssign)
	// It's best to call this method -- rather than manually converting to double -- so that the
	// digits/formatting/precision is consistent throughout the program.
	// Returns OK or FAIL.
	{
		UpdateBinaryDouble(aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE); // When not passing VAR_ATTRIB_CONTENTS_OUT_OF_DATE, all callers of UpdateBinaryDouble() must ensure that mContents is a pure number (e.g. NOT 123abc).
		return OK;
	}

	ResultType AssignSkipAddRef(IObject *aValueToAssign);

	inline ResultType Assign(IObject *aValueToAssign)
	{
		aValueToAssign->AddRef(); // Must be done before Release() in case the only other reference to this object is already in var.  Such a case seems too rare to be worth optimizing by returning early.
		return AssignSkipAddRef(aValueToAssign);
	}

	inline IObject *&Object()
	{
		return (mType == VAR_ALIAS) ? mAliasFor->mObject : mObject;
	}

	inline void ReleaseObject() // L31
	// Caller has ensured that IsObject() == true, not just HasObject().
	{
		// Remove the "this is an object" attribute and re-enable binary number caching.
		mAttrib &= ~(VAR_ATTRIB_OBJECT | VAR_ATTRIB_CACHE_DISABLED | VAR_ATTRIB_NOT_NUMERIC);
		// Release this variable's object.  MUST BE DONE AFTER THE ABOVE IN CASE IT TRIGGERS var.base.__Delete().
		mObject->Release();
	}

	void DisableCache()
	// Callers should be aware that the cache will be re-enabled (except for clipboard) whenever a the address
	// of a variable's contents changes, such as when it needs to be expanded to hold more text.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mAttrib & VAR_ATTRIB_CACHE_DISABLED) // Already marked correctly (and whoever marked it would have already done the steps further below).
			return;
		var.UpdateContents(); // Update mContents & mLength. Must be done prior to below (it also removes the VAR_ATTRIB_CONTENTS_OUT_OF_DATE flag, if present).
		var.mAttrib &= ~VAR_ATTRIB_CACHE; // Remove all cached attributes.
		var.mAttrib |= VAR_ATTRIB_CACHE_DISABLED; // Indicate that in the future, mContents should be kept up-to-date.
	}

	SymbolType IsNonBlankIntegerOrFloat(BOOL aAllowImpure = false)
	// Supports VAR_NORMAL and VAR_CLIPBOARD.  It would need review if any other types need to be supported.
	// Caller must be aware that aAllowFloat==true, aAllowNegative==true, and aAllowAllWhitespace==false
	// are in effect for this function.
	// If caller passes true for aAllowImpure, no explicit handling seems necessary here because:
	// 1) If the text number in mContents is IMPURE, it wouldn't be in the cache in the first place (other
	//    logic ensures this) and thus aAllowImpure need only be delegated to and handled by IsPureNumeric().
	// 2) If the text number in mContents is PURE, the handling below is correct regardless of whether it's
	//    already in the cache.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		switch(var.mAttrib & VAR_ATTRIB_CACHE) // This switch() method should squeeze a little more performance out of it compared to doing "&" for every attribute.  Only works for attributes that are mutually-exclusive, which these are.
		{
		case VAR_ATTRIB_HAS_VALID_INT64: return PURE_INTEGER;
		case VAR_ATTRIB_HAS_VALID_DOUBLE: return PURE_FLOAT;
		case VAR_ATTRIB_NOT_NUMERIC: return PURE_NOT_NUMERIC;
		}
		// Since above didn't return, its numeric status isn't yet known, so determine it.  To conform to
		// historical behavior (backward compatibility), the following doesn't check MAX_INEGER_LENGTH.
		// So any string of digits that is too long to be a legitimate number is still treated as a number
		// anyway (overflow).  Most of our callers are expressions anyway, in which case any unquoted
		// series of digits is always a number, never a string.
		// Below passes FALSE for aUpdateContents because we've already confirmed this var doesn't contain
		// a cached number (so can't need updating) and to suppress an "uninitialized variable" warning.
		// The majority of our callers will call ToInt64/Double() or Contents() after we return, which would
		// trigger a second warning if we didn't suppress ours and StdOut/OutputDebug warn mode is in effect.
		// IF-IS is the only caller that wouldn't cause a warning, but in that case ExpandArgs() would have
		// already caused one.
		SymbolType is_pure_numeric = IsPureNumeric(var.Contents(FALSE), true, false, true, aAllowImpure); // Contents() vs. mContents to support VAR_CLIPBOARD lvalue in a pure expression such as "clipboard:=1,clipboard+=5"
		if (is_pure_numeric == PURE_NOT_NUMERIC && !(var.mAttrib & VAR_ATTRIB_CACHE_DISABLED))
			var.mAttrib |= VAR_ATTRIB_NOT_NUMERIC;
		//else it may be a pure number, which isn't currently tracked via mAttrib (until a cached number is
		// actually stored) because the callers of this function often track it and pass the info on
		// to ToInt64() or ToDouble().
		return is_pure_numeric;
	}

	__int64 ToInt64(BOOL aIsPureInteger)
	// Caller should pass FALSE for aIsPureInteger if this variable's mContents is either:
	// 1) Not a pure number as defined by IsPureNumeric(), namely that the number has a non-numeric part
	//    at the end like 123abc (though pure numbers may have leading and trailing whitespace).
	// 2) It isn't known whether it's a pure number.
	// 3) It's pure but it's the wrong type of number (e.g. it contains decimal point yet ToInt64() vs.
	//    ToDouble() was called).
	// The reason for the above is that IsNonBlankIntegerOrFloat() relies on the state of the cache to
	// accurately report what's in mContents.
	// This function supports VAR_NORMAL and VAR_CLIPBOARD. It would need review to support any other types.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mAttrib & VAR_ATTRIB_HAS_VALID_INT64) // aIsPureInteger isn't checked here because although this caller might not know that it's pure, other logic ensures that the one who actually set it in the cache did know it was pure.
			return var.mContentsInt64;
		//else although the attribute VAR_ATTRIB_HAS_VALID_DOUBLE might be present, casting a double to an __int64
		// might produce a different result than ATOI64() in some cases.  So for backward compatibility and
		// due to rarity of such a circumstance, VAR_ATTRIB_HAS_VALID_DOUBLE isn't checked.
		__int64 int64 = ATOI64(var.Contents()); // Call Contents() vs. using mContents in case of VAR_CLIPBOARD or VAR_ATTRIB_HAS_VALID_DOUBLE, and also for maintainability.
		if (aIsPureInteger && !(var.mAttrib & VAR_ATTRIB_CACHE_DISABLED)) // This is checked to avoid the overhead of calling UpdateBinaryInt64() unconditionally because it may do a lot of things internally.
			var.UpdateBinaryInt64(int64); // Cache the binary number for future uses.
		return int64;
	}

	double ToDouble(BOOL aIsPureFloat)
	// FOR WHAT GOES IN THIS SPOT, SEE IMPORTANT COMMENTS IN ToInt64().
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mAttrib & VAR_ATTRIB_HAS_VALID_DOUBLE)  // aIsPureFloat isn't checked here because although this caller might not know that it's pure, other logic ensures that the one who actually set it in the cache did know it was pure.
			return var.mContentsDouble;
		if (var.mAttrib & VAR_ATTRIB_HAS_VALID_INT64) // If there's already a binary integer stored, don't convert the cache type to "double" because that would cause IsNonBlankIntegerOrFloat() to wrongly return PURE_FLOAT. In addition, float is rarely used and often needed only temporarily, such as x:=VarInt+VarFloat
			return (double)var.mContentsInt64; // As expected, testing shows that casting an int64 to a double is at least 100 times faster than calling ATOF() on the text version of that integer.
		// Otherwise, neither type of binary number is cached yet.
		double d = ATOF(var.Contents()); // Call Contents() vs. using mContents in case of VAR_CLIPBOARD, and also for maintainability and consistency with ToInt64().
		if (aIsPureFloat && !(var.mAttrib & VAR_ATTRIB_CACHE_DISABLED)) // This is checked to avoid the overhead of calling UpdateBinaryInt64() unconditionally because it may do a lot of things internally.
			var.UpdateBinaryDouble(d); // Cache the binary number for future uses.
		return d;
	}

	ResultType ToDoubleOrInt64(ExprTokenType &aToken)
	// aToken.var is the same as the "this" var. Converts var into a number and stores it numerically in aToken.
	// Supports VAR_NORMAL and VAR_CLIPBOARD.  It would need review if any other types need to be supported.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		switch (aToken.symbol = var.IsNonBlankIntegerOrFloat())
		{
		case PURE_INTEGER:
			aToken.value_int64 = var.ToInt64(TRUE);
			break;
		case PURE_FLOAT:
			aToken.value_double = var.ToDouble(TRUE);
			break;
		default: // Not a pure number.
			aToken.marker = _T(""); // For completeness.  Some callers such as BIF_Abs() rely on this being done.
			return FAIL;
		}
		return OK; // Since above didn't return, indicate success.
	}

	void TokenToContents(ExprTokenType &aToken) // L31: Mostly for object support.
	// See ToDoubleOrInt64 for comments.
	{
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// L33: For greater compatibility with the official release and L revisions prior to L31,
		// this section was changed to avoid converting numeric strings to SYM_INTEGER/SYM_FLOAT.
		switch(var.mAttrib & VAR_ATTRIB_CACHE)
		{
		case VAR_ATTRIB_HAS_VALID_INT64:
			aToken.symbol = SYM_INTEGER;
			aToken.value_int64 = var.mContentsInt64;
			return;
		case VAR_ATTRIB_HAS_VALID_DOUBLE:
			aToken.symbol = SYM_FLOAT;
			aToken.value_double = var.mContentsDouble;
			return;
		default:
			if (var.IsObject())
			{
				aToken.symbol = SYM_OBJECT;
				aToken.object = var.mObject;
				aToken.object->AddRef();
				return;
			}
			//else contains a regular string.
			aToken.symbol = SYM_STRING;
			aToken.marker = var.Contents();
		}
	}

	// Not an enum so that it can be global more easily:
	#define VAR_ALWAYS_FREE                    0 // This item and the next must be first and numerically adjacent to
	#define VAR_ALWAYS_FREE_BUT_EXCLUDE_STATIC 1 // each other so that VAR_ALWAYS_FREE_LAST covers only them.
	#define VAR_ALWAYS_FREE_LAST               2 // Never actually passed as a parameter, just a placeholder (see above comment).
	#define VAR_NEVER_FREE                     3
	#define VAR_FREE_IF_LARGE                  4
	void Free(int aWhenToFree = VAR_ALWAYS_FREE, bool aExcludeAliasesAndRequireInit = false);
	ResultType AppendIfRoom(LPTSTR aStr, VarSizeType aLength);
	void AcceptNewMem(LPTSTR aNewMem, VarSizeType aLength);
	void SetLengthFromContents();

	static ResultType BackupFunctionVars(Func &aFunc, VarBkp *&aVarBackup, int &aVarBackupCount);
	void Backup(VarBkp &aVarBkp);
	void Restore(VarBkp &aVarBkp);
	static void FreeAndRestoreFunctionVars(Func &aFunc, VarBkp *&aVarBackup, int &aVarBackupCount);

	#define DISPLAY_NO_ERROR   0  // Must be zero.
	#define DISPLAY_VAR_ERROR  1
	#define DISPLAY_FUNC_ERROR 2
	static ResultType ValidateName(LPCTSTR aName, int aDisplayError = DISPLAY_VAR_ERROR);

	LPTSTR ObjectToText(LPTSTR aBuf, int aBufSize);
	LPTSTR ToText(LPTSTR aBuf, int aBufSize, bool aAppendNewline)
	// Caller must ensure that Type() == VAR_NORMAL.
	// aBufSize is an int so that any negative values passed in from caller are not lost.
	// Caller has ensured that aBuf isn't NULL.
	// Translates this var into its text equivalent, putting the result into aBuf and
	// returning the position in aBuf of its new string terminator.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// v1.0.44.14: Changed it so that ByRef/Aliases report their own name rather than the target's/caller's
		// (it seems more useful and intuitive).
		var.UpdateContents(); // Update mContents and mLength for use below.
		LPTSTR aBuf_orig = aBuf;
		if (var.IsObject())
			aBuf = ObjectToText(aBuf, aBufSize);
		else
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%s[%Iu of %Iu]: %-1.60s%s"), mName // mName not var.mName (see comment above).
				, var._CharLength(), var._CharCapacity() ? (var._CharCapacity() - 1) : 0  // Use -1 since it makes more sense to exclude the terminator.
				, var.mCharContents, var._CharLength() > 60 ? _T("...") : _T(""));
		if (aAppendNewline && BUF_SPACE_REMAINING >= 2)
		{
			*aBuf++ = '\r';
			*aBuf++ = '\n';
			*aBuf = '\0';
		}
		return aBuf;
	}

	__forceinline VarTypeType Type()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		return (mType == VAR_ALIAS) ? mAliasFor->mType : mType;
	}

	__forceinline bool IsStatic()
	{
		return (mScope & VAR_LOCAL_STATIC);
	}

	__forceinline bool IsLocal()
	{
		// Since callers want to know whether this variable is local, even if it's a local alias for a
		// global, don't use the method below:
		//    return (mType == VAR_ALIAS) ? mAliasFor->mIsLocal : mIsLocal;
		return (mScope & VAR_LOCAL);
	}

	__forceinline bool IsNonStaticLocal()
	{
		// Since callers want to know whether this variable is local, even if it's a local alias for a
		// global, don't resolve VAR_ALIAS.
		// Even a ByRef local is considered local here because callers are interested in whether this
		// variable can vary from call to call to the same function (and a ByRef can vary in what it
		// points to).  Variables that vary can thus be altered by the backup/restore process.
		return (mScope & (VAR_LOCAL|VAR_LOCAL_STATIC)) == VAR_LOCAL;
	}

	//__forceinline bool IsFuncParam()
	//{
	//	return (mScope & VAR_LOCAL_FUNCPARAM);
	//}

	__forceinline bool IsDeclared()
	// Returns true if this is a declared var, such as "local var", "static var" or a func param.
	{
		return (mScope & VAR_DECLARED);
	}

	__forceinline bool IsSuperGlobal()
	{
		return (mScope & VAR_SUPER_GLOBAL);
	}

	__forceinline UCHAR &Scope()
	{
		return mScope;
	}

	__forceinline bool IsBinaryClip()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		return (mType == VAR_ALIAS ? mAliasFor->mAttrib : mAttrib) & VAR_ATTRIB_BINARY_CLIP;
	}

	__forceinline bool IsObject() // L31: Indicates this var contains an object reference which must be released if the var is emptied.
	{
		return (mAttrib & VAR_ATTRIB_OBJECT);
	}

	__forceinline bool HasObject() // L31: Indicates this var's effective value is an object reference.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		return (var.mAttrib & VAR_ATTRIB_OBJECT);
	}

	VarSizeType ByteCapacity() // __forceinline() on Capacity, Length, and/or Contents bloats the code and reduces performance.
	// Capacity includes the zero terminator (though if capacity is zero, there will also be a zero terminator in mContents due to it being "").
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// Fix for v1.0.37: Callers want the clipboard's capacity returned, if it has a capacity.  This is
		// because Capacity() is defined as being the size available in Contents(), which for the clipboard
		// would be a pointer to the clipboard-buffer-to-be-written (or zero if none).
		return var.mType == VAR_CLIPBOARD ? g_clip.mCapacity : var.mByteCapacity;
	}

	VarSizeType CharCapacity()
	{
		return ByteCapacity() / sizeof(TCHAR); 
	}

	UNICODE_CHECK VarSizeType Capacity()
	{
		return CharCapacity();
	}

	BOOL HasContents()
	// A fast alternative to Length() that avoids updating mContents.
	// Caller must ensure that Type() is VAR_NORMAL.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		return (var.mAttrib & (VAR_ATTRIB_CONTENTS_OUT_OF_DATE | VAR_ATTRIB_OBJECT)) ? TRUE : !!var.mByteLength; // i.e. the only time var.mLength isn't a valid indicator of an empty variable is when VAR_ATTRIB_CONTENTS_OUT_OF_DATE, in which case the variable is non-empty because there is a binary number in it.
	}

	BOOL HasUnflushedBinaryNumber()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		return var.mAttrib & VAR_ATTRIB_CONTENTS_OUT_OF_DATE; // VAR_ATTRIB_CONTENTS_OUT_OF_DATE implies that either VAR_ATTRIB_HAS_VALID_INT64 or VAR_ATTRIB_HAS_VALID_DOUBLE is also present.
	}

	VarSizeType &ByteLength() // __forceinline() on Capacity, Length, and/or Contents bloats the code and reduces performance.
	// This should not be called to discover a non-NORMAL var's length (nor that of an environment variable)
	// because their lengths aren't knowable without calling Get().
	// Returns a reference so that caller can use this function as an lvalue.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mType == VAR_NORMAL)
		{
			if (var.mAttrib & VAR_ATTRIB_CONTENTS_OUT_OF_DATE)
				var.UpdateContents();  // Update mContents (and indirectly, mLength).
			return var.mByteLength;
		}
		// Since the length of the clipboard isn't normally tracked, we just return a
		// temporary storage area for the caller to use.  Note: This approach is probably
		// not thread-safe, but currently there's only one thread so it's not an issue.
		// For reserved vars do the same thing as above, but this function should never
		// be called for them:
		static VarSizeType length; // Must be static so that caller can use its contents. See above.
		return length;
	}

	VarSizeType SetCharLength(VarSizeType len)
	{
		 ByteLength() = len * sizeof(TCHAR);
		 return len;
	}

	VarSizeType CharLength()
	{
		return ByteLength() / sizeof(TCHAR);
	}

	UNICODE_CHECK VarSizeType Length()
	{
		return CharLength();
	}

	VarSizeType LengthIgnoreBinaryClip()
	// Returns 0 for types other than VAR_NORMAL and VAR_CLIPBOARD.
	// IMPORTANT: Environment variables aren't supported here, so caller must either want such
	// variables treated as blank, or have already checked that they're not environment variables.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// Return the apparent length of the string (i.e. the position of its first binary zero).
		return (var.mType == VAR_NORMAL && !(var.mAttrib & VAR_ATTRIB_BINARY_CLIP))
			? var.Length() // Use Length() vs. mLength so that the length is updated if necessary.
			: _tcslen(var.Contents()); // Use Contents() vs. mContents to support VAR_CLIPBOARD.
	}

	//BYTE *ByteContents(BOOL aAllowUpdate = TRUE)
	//{
	//	return (BYTE *) CharContents(aAllowUpdate);
	//}

	TCHAR *Contents(BOOL aAllowUpdate = TRUE, BOOL aNoWarnUninitializedVar = FALSE)
	// Callers should almost always pass TRUE for aAllowUpdate because any caller who wants to READ from
	// mContents would almost always want it up-to-date.  Any caller who wants to WRITE to mContents would
	// would almost always have called Assign(NULL, ...) prior to calling Contents(), which would have
	// cleared the VAR_ATTRIB_CONTENTS_OUT_OF_DATE flag.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if ((var.mAttrib & VAR_ATTRIB_CONTENTS_OUT_OF_DATE) && aAllowUpdate) // VAR_ATTRIB_CONTENTS_OUT_OF_DATE is checked here and in the function below, for performance.
			var.UpdateContents(); // This also clears the VAR_ATTRIB_CONTENTS_OUT_OF_DATE flag.
		if (var.mType == VAR_NORMAL)
		{
			// If aAllowUpdate is FALSE, the caller just wants to compare mCharContents to another address.
			// Otherwise, the caller is probably going to use mCharContents and might want a warning:
			if (aAllowUpdate && !aNoWarnUninitializedVar)
				var.MaybeWarnUninitialized();
			return var.mCharContents;
		}
		if (var.mType == VAR_CLIPBOARD)
			// The returned value will be a writable mem area if clipboard is open for write.
			// Otherwise, the clipboard will be opened physically, if it isn't already, and
			// a pointer to its contents returned to the caller:
			return g_clip.Contents();
		return sEmptyString; // For reserved vars (but this method should probably never be called for them).
	}

	__forceinline void ConvertToNonAliasIfNecessary() // __forceinline because it's currently only called from one place.
	// When this function actually converts an alias into a normal variable, the variable's old
	// attributes (especially mContents and mCapacity) become dominant again.  This prevents a memory
	// leak in a case where a UDF is defined to provide a default value for a ByRef parameter, and is
	// called both with and without that parameter.
	{
		mAliasFor = NULL; // This also sets its counterpart in the union (mLength) to zero, which is appropriate because mContents should have been set to blank by a previous call to Free().
		mType = VAR_NORMAL; // It might already be this type, so this is just in case it's VAR_ALIAS.
	}

	__forceinline Var *ResolveAlias()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		return (mType == VAR_ALIAS) ? mAliasFor : this; // Return target if it's an alias, or itself if not.
	}

	__forceinline void UpdateAlias(Var *aTargetVar) // __forceinline because it's currently only called from one place.
	// Caller must ensure that aTargetVar isn't NULL.
	// When this function actually converts a normal variable into an alias , the variable's old
	// attributes (especially mContents and mCapacity) are hidden/suppressed by virtue of all Var:: methods
	// obeying VAR_ALIAS and resolving it to be the target variable.  This prevents a memory
	// leak in a case where a UDF is defined to provide a default value for a ByRef parameter, and is
	// called both with and without that parameter.
	{
		// BELOW IS THE MEANS BY WHICH ALIASES AREN'T ALLOWED TO POINT TO OTHER ALIASES, ONLY DIRECTLY TO
		// THE TARGET VAR.
		// Resolve aliases-to-aliases for performance and to increase the expectation of
		// reliability since a chain of aliases-to-aliases might break if an alias in
		// the middle is ever allowed to revert to a non-alias (or gets deleted).
		// A caller may ask to create an alias to an alias when a function calls another
		// function and passes to it one of its own byref-params.
		while (aTargetVar->mType == VAR_ALIAS)
			aTargetVar = aTargetVar->mAliasFor;

		// The following is done only after the above in case there's ever a way for the above
		// to circle back to become this variable.
		// Prevent potential infinite loops in other methods by refusing to change an alias
		// to point to itself.
		if (aTargetVar == this)
			return;

		mAliasFor = aTargetVar; // Should always be non-NULL due to various checks elsewhere.
		mType = VAR_ALIAS; // It might already be this type, so this is just in case it's VAR_NORMAL.
	}

	ResultType Close(bool aIsBinaryClip = false)
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mType == VAR_CLIPBOARD && g_clip.IsReadyForWrite())
			return g_clip.Commit(); // Writes the new clipboard contents to the clipboard and closes it.
		// The binary-clip attribute is also reset here for cases where a caller uses a variable without
		// having called Assign() to resize it first, which can happen if the variable's capacity is already
		// sufficient to hold the desired contents.  VAR_ATTRIB_CONTENTS_OUT_OF_DATE is also removed below
		// for maintainability; it shouldn't be necessary because any caller of Close() should have previously
		// called something that updates the flags, such as Contents().
		var.mAttrib &= ~VAR_ATTRIB_OFTEN_REMOVED;
		if (aIsBinaryClip) // If true, caller should ensure that var.mType isn't VAR_CLIPBOARD because it doesn't seem possible/valid for the clipboard to contain a binary image of the clipboard.
			var.mAttrib |= VAR_ATTRIB_BINARY_CLIP;
		//else (already done above)
		//	var.mAttrib &= ~VAR_ATTRIB_BINARY_CLIP;
		return OK; // In all other cases.
	}

	// Constructor:
	Var(LPTSTR aVarName, void *aType, UCHAR aScope)
		// The caller must ensure that aVarName is non-null.
		: mCharContents(sEmptyString) // Invariant: Anyone setting mCapacity to 0 must also set mContents to the empty string.
		// Doesn't need initialization: , mContentsInt64(NULL)
		, mByteLength(0) // This also initializes mAliasFor within the same union.
		, mHowAllocated(ALLOC_NONE)
		, mAttrib(VAR_ATTRIB_UNINITIALIZED) // Seems best not to init empty vars to VAR_ATTRIB_NOT_NUMERIC because it would reduce maintainability, plus finding out whether an empty var is numeric via IsPureNumeric() is a very fast operation.
		, mScope(aScope)
		, mName(aVarName) // Caller gave us a pointer to dynamic memory for this (or static in the case of ResolveVarOfArg()).
	{
		if (aType > (void *)VAR_LAST_TYPE) // Relies on the fact that numbers less than VAR_LAST_TYPE can never realistically match the address of any function.
		{
			mType = VAR_BUILTIN;
			mBIV = (BuiltInVarType)aType; // This also initializes mCapacity within the same union.
			mAttrib = 0; // Built-in vars are considered initialized, by definition.
		}
		else
		{
			mType = (VarTypeType)aType;
			mByteCapacity = 0; // This also initializes mBIV within the same union.
			if (mType != VAR_NORMAL)
				mAttrib = 0; // Any vars that aren't VAR_NORMAL are considered initialized, by definition.
		}
	}

	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}


	__forceinline bool IsUninitializedNormalVar()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		return (var.mAttrib & VAR_ATTRIB_UNINITIALIZED);
	}

	__forceinline void MarkInitialized()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		mAttrib &= ~VAR_ATTRIB_UNINITIALIZED;
	}

	__forceinline void MaybeWarnUninitialized();

}; // class Var
#pragma pack(pop) // Calling pack with no arguments restores the default value (which is 8, but "the alignment of a member will be on a boundary that is either a multiple of n or a multiple of the size of the member, whichever is smaller.")

#pragma warning(pop)

#endif
