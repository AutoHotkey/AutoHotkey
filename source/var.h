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

#define MAX_ALLOC_SIMPLE 64  // Do not decrease this much since it is used for the sizing of some built-in variables.
#define SMALL_STRING_LENGTH (MAX_ALLOC_SIMPLE - 1)  // The largest string that can fit in the above.
#define DEREF_BUF_EXPAND_INCREMENT (16 * 1024) // Reduced from 32 to 16 in v1.0.46.07 to reduce the memory utilization of deeply recursive UDFs.

enum AllocMethod {ALLOC_NONE, ALLOC_SIMPLE, ALLOC_MALLOC};
enum VarTypes
{
  // The following must all be LOW numbers to avoid any realistic chance of them matching the address of
  // any function (namely a BIV_* function).
  VAR_ALIAS  // VAR_ALIAS must always have a non-NULL mAliasFor.  In other ways it's the same as VAR_NORMAL.  VAR_ALIAS is never seen because external users call Var::Type(), which automatically resolves ALIAS to some other type.
, VAR_NORMAL // Most variables, such as those created by the user, are this type.
, VAR_CONSTANT // or as I like to say, not variable.
, VAR_VIRTUAL
, VAR_LAST_TYPE = VAR_VIRTUAL
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
	union
	{
		__int64 mContentsInt64; // 64-bit members kept at the top of the struct to reduce the chance that they'll span 2 vs. 1 64-bit regions.
		double mContentsDouble;
		IObject *mObject;
	};
	Var *mVar; // Used to save the target var to which these backed up contents will later be restored.
	union
	{
		char *mByteContents;
		TCHAR *mCharContents;
	};
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

	void ToToken(ExprTokenType &aValue);
};

#define BIV_DECL_R(name) void name(ResultToken &aResultToken, LPTSTR aVarName)
#define BIV_DECL_W(name) void name(ResultToken &aResultToken, LPTSTR aVarName, ExprTokenType &aValue)
#define BIV_DECL_RW(name) BIV_DECL_R(name); BIV_DECL_W(name##_Set)

struct VirtualVar
{
	typedef BIV_DECL_R((* Getter));
	typedef BIV_DECL_W((* Setter));
	Getter Get;
	Setter Set;
};

struct VarEntry
{
	LPTSTR name;
	VirtualVar type;
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
		VirtualVar *mVV; // VAR_VIRTUAL
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
	VarSizeType mByteCapacity; // In bytes.  Includes the space for the zero terminator.
	AllocMethodType mHowAllocated; // Keep adjacent/contiguous with the below to save memory.
	#define VAR_ATTRIB_CONTENTS_OUT_OF_DATE						0x01 // Combined with VAR_ATTRIB_IS_INT64/DOUBLE/OBJECT to indicate mContents is not current.
	#define VAR_ATTRIB_UNINITIALIZED							0x02 // Var requires initialization before use.
	#define VAR_ATTRIB_CONTENTS_OUT_OF_DATE_UNTIL_REASSIGNED	0x04 // Indicates that the VAR_ATTRIB_CONTENTS_OUT_OF_DATE flag should be kept until the variable is reassigned.
	#define VAR_ATTRIB_NOT_NUMERIC								0x08 // A prior call to IsNumeric() determined the var's value is PURE_NOT_NUMERIC.
	#define VAR_ATTRIB_IS_INT64									0x10 // Var's proper value is in mContentsInt64.
	#define VAR_ATTRIB_IS_DOUBLE								0x20 // Var's proper value is in mContentsDouble.
	#define VAR_ATTRIB_IS_OBJECT								0x40 // Var's proper value is in mObject.
	#define VAR_ATTRIB_VIRTUAL_OPEN								0x80 // Virtual var is open for writing.
	#define VAR_ATTRIB_CACHE (VAR_ATTRIB_IS_INT64 | VAR_ATTRIB_IS_DOUBLE | VAR_ATTRIB_NOT_NUMERIC) // These three are mutually exclusive.
	#define VAR_ATTRIB_TYPES (VAR_ATTRIB_IS_INT64 | VAR_ATTRIB_IS_DOUBLE | VAR_ATTRIB_IS_OBJECT) // These are mutually exclusive (but NOT_NUMERIC may be combined with OBJECT or BINARY_CLIP).
	#define VAR_ATTRIB_OFTEN_REMOVED (VAR_ATTRIB_CACHE | VAR_ATTRIB_CONTENTS_OUT_OF_DATE | VAR_ATTRIB_UNINITIALIZED | VAR_ATTRIB_CONTENTS_OUT_OF_DATE_UNTIL_REASSIGNED)
	VarAttribType mAttrib;  // Bitwise combination of the above flags (but many of them may be mutually exclusive).
	#define VAR_GLOBAL			0x01
	#define VAR_LOCAL			0x02
	#define VAR_FORCE_LOCAL		0x04 // Flag reserved for force-local mode in functions (not used in Var::mScope).
	#define VAR_DOWNVAR			0x08 // This var is captured by a nested function/closure (it's in Func::mDownVar).
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

#ifdef CONFIG_DEBUGGER
	friend class Debugger;
#endif

	// Caller has verified mType == VAR_VIRTUAL.
	bool HasSetter() { return mVV->Set; }
	ResultType AssignVirtual(ExprTokenType &aValue);

	// Unconditionally accepts new memory, bypassing the usual redirection to Assign() for VAR_VIRTUAL.
	void _AcceptNewMem(LPTSTR aNewMem, VarSizeType aLength);

	ResultType AssignBinaryNumber(__int64 aNumberAsInt64, VarAttribType aAttrib = VAR_ATTRIB_IS_INT64);

	void UpdateContents()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mAttrib & VAR_ATTRIB_CONTENTS_OUT_OF_DATE)
		{
			// THE FOLLOWING ISN'T NECESSARY BECAUSE THE ASSIGN() CALLS BELOW DO IT:
			//var.mAttrib &= ~VAR_ATTRIB_CONTENTS_OUT_OF_DATE;
			TCHAR value_string[MAX_NUMBER_SIZE];
			bool restore_attribute = var.mAttrib & VAR_ATTRIB_CONTENTS_OUT_OF_DATE_UNTIL_REASSIGNED; // The Assign() calls also clears this flag, hence, set this variable to indicate wether to restore this flag or not.
			if (var.mAttrib & VAR_ATTRIB_IS_INT64)
			{
				var.Assign(ITOA64(var.mContentsInt64, value_string)); // Return value currently not checked for this or the below.
				var.mAttrib |= VAR_ATTRIB_IS_INT64; // Re-enable the cache because Assign() disables it (since all other callers want that).
			}
			else if (var.mAttrib & VAR_ATTRIB_IS_DOUBLE)
			{
				var.Assign(value_string, FTOA(var.mContentsDouble, value_string, _countof(value_string)));
				var.mAttrib |= VAR_ATTRIB_IS_DOUBLE; // Re-enable the cache because Assign() disables it (since all other callers want that).
			}
			//else nothing to update, which shouldn't happen in this block unless there's a flaw or bug somewhere.
			if (restore_attribute)
				var.mAttrib |= VAR_ATTRIB_CONTENTS_OUT_OF_DATE_UNTIL_REASSIGNED | VAR_ATTRIB_CONTENTS_OUT_OF_DATE;
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
	// when a script forgets to call VarSetStrCapacity before passing a buffer to some function that writes a
	// string to it.  There is now some code there that tries to detect when that happens.
	static TCHAR sEmptyString[1]; // See above.

	VarSizeType Get(LPTSTR aBuf = NULL);
	void Get(ResultToken &aResultToken);
	ResultType AssignHWND(HWND aWnd);
	ResultType Assign(Var &aVar);
	ResultType Assign(ExprTokenType &aToken);
	static ResultType GetClipboardAll(void **aData, size_t *aDataSize);
	static ResultType SetClipboardAll(void *aData, size_t aDataSize);
	// Assign(char *, ...) has been break into four methods below.
	// This should prevent some mistakes, as characters and bytes are not interchangeable in the Unicode build.
	// Callers must make sure which one is the right method to call.
	ResultType AssignString(LPCTSTR aBuf = NULL, VarSizeType aLength = VARSIZE_MAX, bool aExactSize = false);
	inline ResultType Assign(LPCTSTR aBuf, VarSizeType aLength = VARSIZE_MAX, bool aExactSize = false)
	{
		ASSERT(aBuf); // aBuf shouldn't be NULL, use SetCapacity([length in bytes]) or AssignString(NULL, [length in characters]) instead.
		return AssignString(aBuf, aLength, aExactSize);
	}
	inline ResultType Assign()
	{
		return AssignString();
	}
	ResultType SetCapacity(VarSizeType aByteLength, bool aExactSize = false)
	{
#ifdef UNICODE
		return AssignString(NULL, (aByteLength >> 1) + (aByteLength & 1), aExactSize);
#else
		return AssignString(NULL, aByteLength, aExactSize);
#endif
	}

	ResultType AssignStringFromCodePage(LPCSTR aBuf, int aLength = -1, UINT aCodePage = CP_ACP);
	ResultType AssignStringToCodePage(LPCWSTR aBuf, int aLength = -1, UINT aCodePage = CP_ACP, DWORD aFlags = WC_NO_BEST_FIT_CHARS, char aDefChar = '?');
	inline ResultType AssignStringW(LPCWSTR aBuf, int aLength = -1)
	{
#ifdef UNICODE
		// Pass aExactSize=true for consistency with AssignStringTo/FromCodePage/UTF8.
		return AssignString(aBuf, aLength, true);
#else
		return AssignStringToCodePage(aBuf, aLength);
#endif
	}

	inline ResultType Assign(DWORD aValueToAssign)
	{
		return AssignBinaryNumber(aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE|VAR_ATTRIB_IS_INT64);
	}

	inline ResultType Assign(int aValueToAssign)
	{
		return AssignBinaryNumber(aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE|VAR_ATTRIB_IS_INT64);
	}

	inline ResultType Assign(__int64 aValueToAssign)
	{
		return AssignBinaryNumber(aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE|VAR_ATTRIB_IS_INT64);
	}

	inline ResultType Assign(VarSizeType aValueToAssign)
	{
		return AssignBinaryNumber(aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE|VAR_ATTRIB_IS_INT64);
	}

	inline ResultType Assign(double aValueToAssign)
	{
		// The type-casting below interprets the contents of aValueToAssign as an __int64 without actually
		// converting from double to __int64.  Although the generated code isn't measurably smaller, hopefully
		// the compiler resolves it into something that performs better than a memcpy into a temporary variable.
		// Benchmarks show that the performance is at most a few percent worse than having code similar to
		// AssignBinaryNumber() in here.
		return AssignBinaryNumber(*(__int64 *)&aValueToAssign, VAR_ATTRIB_CONTENTS_OUT_OF_DATE|VAR_ATTRIB_IS_DOUBLE);
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

	IObject *ToObject()
	{
		if (mType == VAR_ALIAS)
			return mAliasFor->ToObject();
		if (IsObject())
			return mObject;
		MaybeWarnUninitialized();
		return NULL;
	}

	void ReleaseObject()
	// Caller has ensured that IsObject() == true, not just HasObject().
	{
		// Remove the attributes applied by AssignSkipAddRef().
		mAttrib &= ~(VAR_ATTRIB_IS_OBJECT | VAR_ATTRIB_NOT_NUMERIC);
		// MUST BE DONE AFTER THE ABOVE IN CASE IT TRIGGERS __Delete:
		// Release this variable's object.  Setting mObject = NULL is not necessary
		// since the value of mObject, mContentsInt64 and mContentsDouble is never used
		// unless an attribute is present which indicates which one is valid.
		mObject->Release();
	}

	SymbolType IsNumeric()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		switch (var.mAttrib & VAR_ATTRIB_CACHE) // This switch() method should squeeze a little more performance out of it compared to doing "&" for every attribute.  Only works for attributes that are mutually-exclusive, which these are.
		{
		case VAR_ATTRIB_IS_INT64: return PURE_INTEGER;
		case VAR_ATTRIB_IS_DOUBLE: return PURE_FLOAT;
		case VAR_ATTRIB_NOT_NUMERIC: return PURE_NOT_NUMERIC;
		}
		// Since above didn't return, its numeric status isn't yet known, so determine it.  For simplicity
		// (and because Length() doesn't support VAR_VIRTUAL), the following doesn't check MAX_NUMBER_LENGTH.
		// So any string of digits that is too long to be a legitimate number is still treated as a number
		// anyway (overflow).  Most of our callers are expressions anyway, in which case any unquoted
		// series of digits would've been assigned as a pure number and handled above.
		// Below passes FALSE for aUpdateContents because we've already confirmed this var doesn't contain
		// a cached number (so can't need updating) and to suppress an "uninitialized variable" warning.
		// The majority of our callers will call ToInt64/Double() or Contents() after we return, which would
		// trigger a second warning if we didn't suppress ours and StdOut/OutputDebug warn mode is in effect.
		// IF-IS is the only caller that wouldn't cause a warning, but in that case ExpandArgs() would have
		// already caused one.
		SymbolType is_pure_numeric = ::IsNumeric(var.Contents(FALSE), true, false, true); // Contents() vs. mContents to support VAR_VIRTUAL lvalue in a pure expression such as "a_clipboard:=1,a_clipboard+=5"
		if (is_pure_numeric == PURE_NOT_NUMERIC && var.mType != VAR_VIRTUAL)
			var.mAttrib |= VAR_ATTRIB_NOT_NUMERIC;
		return is_pure_numeric;
	}
	
	SymbolType IsPureNumeric()
	// Unlike IsNumeric(), this is purely based on whether a pure number was assigned to this variable.
	// Implicitly supports all types of variables, since these attributes are used only by VAR_NORMAL.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		switch (var.mAttrib & VAR_ATTRIB_CACHE)
		{
		case VAR_ATTRIB_IS_INT64: return PURE_INTEGER;
		case VAR_ATTRIB_IS_DOUBLE: return PURE_FLOAT;
		}
		return PURE_NOT_NUMERIC;
	}

	inline int IsPureNumericOrObject()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		return (var.mAttrib & (VAR_ATTRIB_IS_INT64 | VAR_ATTRIB_IS_DOUBLE | VAR_ATTRIB_IS_OBJECT));
	}

	__int64 ToInt64()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mAttrib & VAR_ATTRIB_IS_INT64)
			return var.mContentsInt64;
		if (var.mAttrib & VAR_ATTRIB_IS_DOUBLE)
			// Since mContentsDouble is the true value of this var, cast it to __int64 rather than
			// calling ATOI64(var.Contents()), which might produce a different result in some cases.
			return (__int64)var.mContentsDouble;
		// Otherwise, this var does not contain a pure number.
		return ATOI64(var.Contents()); // Call Contents() vs. using mContents in case of VAR_VIRTUAL or VAR_ATTRIB_IS_DOUBLE, and also for maintainability.
	}

	double ToDouble()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mAttrib & VAR_ATTRIB_IS_DOUBLE)
			return var.mContentsDouble;
		if (var.mAttrib & VAR_ATTRIB_IS_INT64)
			return (double)var.mContentsInt64; // As expected, testing shows that casting an int64 to a double is at least 100 times faster than calling ATOF() on the text version of that integer.
		// Otherwise, this var does not contain a pure number.
		return ATOF(var.Contents()); // Call Contents() vs. using mContents in case of VAR_VIRTUAL, and also for maintainability and consistency with ToInt64().
	}

	ResultType ToDoubleOrInt64(ExprTokenType &aToken)
	// aToken.var is the same as the "this" var. Converts var into a number and stores it numerically in aToken.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		switch (aToken.symbol = var.IsNumeric())
		{
		case PURE_INTEGER:
			aToken.value_int64 = var.ToInt64();
			break;
		case PURE_FLOAT:
			aToken.value_double = var.ToDouble();
			break;
		default: // Not a pure number.
			return FAIL;
		}
		return OK; // Since above didn't return, indicate success.
	}

	void ToTokenSkipAddRef(ExprTokenType &aToken)
	// See ToDoubleOrInt64 for comments.
	{
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		switch (var.mAttrib & VAR_ATTRIB_TYPES)
		{
		case VAR_ATTRIB_IS_INT64:
			aToken.SetValue(var.mContentsInt64);
			return;
		case VAR_ATTRIB_IS_DOUBLE:
			aToken.SetValue(var.mContentsDouble);
			return;
		case VAR_ATTRIB_IS_OBJECT:
			aToken.SetValue(var.mObject);
			return;
		default:
			// VAR_ATTRIB_BINARY_CLIP or 0.
			aToken.SetValue(var.Contents(), var.Length());
		}
	}

	void ToToken(ExprTokenType &aToken)
	{
		ToTokenSkipAddRef(aToken);
		if (aToken.symbol == SYM_OBJECT)
			aToken.object->AddRef();
	}

	bool MoveMemToResultToken(ResultToken &aResultToken)
	// Caller must ensure mType == VAR_NORMAL.
	{
		if (mHowAllocated == ALLOC_MALLOC // malloc() is our allocator...
			&& ((mAttrib & (VAR_ATTRIB_IS_INT64 | VAR_ATTRIB_IS_DOUBLE | VAR_ATTRIB_IS_OBJECT | VAR_ATTRIB_UNINITIALIZED)) == 0)
			&& mByteCapacity) // ...and we actually have memory allocated.
		{
			// Caller has determined that this var's value won't be needed anymore, so avoid
			// an extra malloc and copy by moving this var's memory block into aResultToken:
			aResultToken.StealMem(this);
			return true;
		}
		return false;
	}

	bool ToReturnValue(ResultToken &aResultToken)
	{
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// Caller may have checked attrib/type, but check it anyway for maintainability:
		if ((var.mAttrib & (VAR_ATTRIB_IS_INT64 | VAR_ATTRIB_IS_DOUBLE | VAR_ATTRIB_IS_OBJECT | VAR_ATTRIB_UNINITIALIZED)) != 0
			// For static/global variables, return a direct pointer to Contents() and
			// let the caller copy it into persistent memory if needed.
			|| !var.IsNonStaticLocal())
		{
			var.ToToken(aResultToken);
			return true;
		}
		if (mType == VAR_ALIAS)
			// This var is an alias for another var.  Even though the target is local,
			// there's no quick way to determine if it's a local of this function, so
			// it's best to assume that we can't steal its memory.  Instead we rely on
			// ExpandExpression copying the string into the deref buffer.
			return false;
		// Var is local.  Since the function is returning, the var is about to be freed.
		// Instead of copying and then freeing its contents, let the caller take ownership:
		if (var.mHowAllocated == ALLOC_MALLOC && var.mByteCapacity)
		{
			// var.mCharContents was allocated with malloc(); pass it back to the caller.
			aResultToken.StealMem(&var);
		}
		else
		{
			// Copy contents into aResultToken.buf, which is always large enough because
			// MAX_ALLOC_SIMPLE < MAX_NUMBER_LENGTH.  mCharContents is used vs Contents()
			// because this isn't a number and therefore never needs UpdateContents().
			// Although Contents() should be harmless, we want to be absolutely sure
			// length isn't increased since that could cause buffer overflow.
			memcpy(aResultToken.marker = aResultToken.buf, var.mCharContents, var.mByteLength + sizeof(TCHAR));
			aResultToken.marker_length = var.mByteLength / sizeof(TCHAR);
		}
		return true;
	}

	LPTSTR StealMem()
	// Caller must ensure that mType == VAR_NORMAL and mHowAllocated == ALLOC_MALLOC.
	{
		LPTSTR mem = mCharContents;
		mCharContents = Var::sEmptyString;
		mByteCapacity = 0;
		mByteLength = 0;
		return mem;
	}

	// Not an enum so that it can be global more easily:
	#define VAR_ALWAYS_FREE                    0 // This item and the next must be first and numerically adjacent to
	#define VAR_ALWAYS_FREE_BUT_EXCLUDE_STATIC 1 // each other so that VAR_ALWAYS_FREE_LAST covers only them.
	#define VAR_ALWAYS_FREE_LAST               2 // Never actually passed as a parameter, just a placeholder (see above comment).
	#define VAR_NEVER_FREE                     3
	#define VAR_FREE_IF_LARGE                  4
	void Free(int aWhenToFree = VAR_ALWAYS_FREE, bool aExcludeAliasesAndRequireInit = false);
	ResultType Append(LPTSTR aStr, VarSizeType aLength);
	ResultType AppendIfRoom(LPTSTR aStr, VarSizeType aLength);
	void AcceptNewMem(LPTSTR aNewMem, VarSizeType aLength);
	void SetLengthFromContents();

	static ResultType BackupFunctionVars(UserFunc &aFunc, VarBkp *&aVarBackup, int &aVarBackupCount);
	void Backup(VarBkp &aVarBkp);
	void Restore(VarBkp &aVarBkp);
	static void FreeAndRestoreFunctionVars(UserFunc &aFunc, VarBkp *&aVarBackup, int &aVarBackupCount);

	#define DISPLAY_NO_ERROR    0  // Must be zero.
	#define DISPLAY_VAR_ERROR   1
	#define DISPLAY_FUNC_ERROR  2
	#define DISPLAY_CLASS_ERROR 3
	#define DISPLAY_GROUP_ERROR 4
	#define DISPLAY_METHOD_ERROR 5
	#define VALIDATENAME_SUBJECT_INDEX(n) (n-1)
	#define VALIDATENAME_SUBJECTS { _T("variable"), _T("function"), _T("class"), _T("group"), _T("method") }
	static ResultType ValidateName(LPCTSTR aName, int aDisplayError = DISPLAY_VAR_ERROR);

	LPTSTR ObjectToText(LPTSTR aName, LPTSTR aBuf, int aBufSize);
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
		switch (var.IsPureNumericOrObject())
		{
		case VAR_ATTRIB_IS_INT64:
		case VAR_ATTRIB_IS_DOUBLE:
			// Dont display [len of cap] since it's not relevant for pure numbers.
			// This also makes it obvious that the var contains a pure number:
			aBuf += sntprintf(aBuf, aBufSize, _T("%s: %s"), mName, var.mCharContents);
			break;
		case VAR_ATTRIB_IS_OBJECT:
			aBuf = var.ObjectToText(this->mName, aBuf, aBufSize);
			break;
		default:
			aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%s[%Iu of %Iu]: %-1.60s%s"), mName // mName not var.mName (see comment above).
				, var._CharLength(), var._CharCapacity() ? (var._CharCapacity() - 1) : 0  // Use -1 since it makes more sense to exclude the terminator.
				, var.mCharContents, var._CharLength() > 60 ? _T("...") : _T(""));
		}
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

	// Convert VAR_NORMAL to VAR_CONSTANT.
	void MakeReadOnly()
	{
		ASSERT(mType == VAR_NORMAL); // Should never be called on VAR_ALIAS or VAR_VIRTUAL.
		ASSERT(!(mAttrib & VAR_ATTRIB_UNINITIALIZED));
		mType = VAR_CONSTANT;
	}

#define VAR_IS_READONLY(var) ((var).IsReadOnly()) // This used to rely on var.Type(), which is no longer enough.
	bool IsReadOnly()
	{
		auto &var = *ResolveAlias();
		return var.mType == VAR_CONSTANT || (var.mType == VAR_VIRTUAL && !var.HasSetter());
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

	__forceinline bool IsFuncParam()
	{
		return (mScope & VAR_LOCAL_FUNCPARAM);
	}

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

	static LPCTSTR DeclarationType(int aDeclType)
	{
		if (aDeclType & VAR_LOCAL)
		{
			if (aDeclType & VAR_LOCAL_STATIC)
				return _T("static");
			if (aDeclType & VAR_LOCAL_FUNCPARAM)
				return _T("parameter");
			return _T("local");
		}
		return _T("global");
	}

	__forceinline bool IsObject() // L31: Indicates this var contains an object reference which must be released if the var is emptied.
	{
		return (mAttrib & VAR_ATTRIB_IS_OBJECT);
	}

	__forceinline bool HasObject() // L31: Indicates this var's effective value is an object reference.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		return (var.mAttrib & VAR_ATTRIB_IS_OBJECT);
	}

	VarSizeType ByteCapacity()
	// Capacity includes the zero terminator (though if capacity is zero, there will also be a zero terminator in mContents due to it being "").
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		return var.mByteCapacity;
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
		// mAttrib is checked because mByteLength isn't applicable when the var contains a pure number
		// or an object.  VAR_ATTRIB_BINARY_CLIP is also included, but doesn't affect the result because
		// mByteLength is always non-zero in that case.
		return (var.mAttrib & VAR_ATTRIB_TYPES) ? TRUE : var.mByteLength != 0;
	}

	VarSizeType &ByteLength()
	// This should not be called to discover a non-NORMAL var's length (nor that of an environment variable)
	// because their lengths aren't knowable without calling Get().
	// Returns a reference so that caller can use this function as an lvalue.
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		// There's no apparent reason to avoid using mByteLength for VAR_VIRTUAL,
		// so it's used temporarily despite the comments below.  Even if mByteLength
		// isn't always accurate, it's no less accurate than a static variable would be.
		//if (var.mType == VAR_NORMAL)
		{
			if (var.mAttrib & VAR_ATTRIB_CONTENTS_OUT_OF_DATE)
				var.UpdateContents();  // Update mContents (and indirectly, mByteLength).
			return var.mByteLength;
		}
		// Since the length of the clipboard isn't normally tracked, we just return a
		// temporary storage area for the caller to use.  Note: This approach is probably
		// not thread-safe, but currently there's only one thread so it's not an issue.
		// For reserved vars do the same thing as above, but this function should never
		// be called for them:
		//static VarSizeType length; // Must be static so that caller can use its contents. See above.
		//return length;
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
			var.UpdateContents(); // This also clears the VAR_ATTRIB_CONTENTS_OUT_OF_DATE flag unless VAR_ATTRIB_CONTENTS_OUT_OF_DATE_UNTIL_REASSIGNED is set.
		if (var.mType == VAR_NORMAL)
		{
			// If aAllowUpdate is FALSE, the caller just wants to compare mCharContents to another address.
			// Otherwise, the caller is probably going to use mCharContents and might want a warning:
			if (aAllowUpdate && !aNoWarnUninitializedVar)
				var.MaybeWarnUninitialized();
		}
		if (var.mType == VAR_VIRTUAL && !(var.mAttrib & VAR_ATTRIB_VIRTUAL_OPEN) && aAllowUpdate)
		{
			// This var isn't open for writing, so populate mCharContents with its current value.
			var.PopulateVirtualVar();
			// The following ensures the contents are updated each time Contents() is called:
			var.mAttrib &= ~VAR_ATTRIB_VIRTUAL_OPEN;
		}
		return var.mCharContents;
	}

	// Populate a virtual var with its current value, as a string.
	// Caller has verified aVar->mType == VAR_VIRTUAL.
	ResultType PopulateVirtualVar();

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

	ResultType Close()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		if (var.mType == VAR_VIRTUAL)
		{
			// Commit the value in our temporary buffer.
			auto result = var.AssignVirtual(ExprTokenType(var.mCharContents, var.CharLength()));
			Free(); // Free temporary memory.
			var.mAttrib &= ~VAR_ATTRIB_VIRTUAL_OPEN;
			return result;
		}
		// VAR_ATTRIB_CONTENTS_OUT_OF_DATE is removed below for maintainability; it shouldn't be
		// necessary because any caller of Close() should have previously called something that
		// updates the flags, such as Contents().
		var.mAttrib &= ~VAR_ATTRIB_OFTEN_REMOVED;
		//else (already done above)
		//	var.mAttrib &= ~VAR_ATTRIB_BINARY_CLIP;
		return OK; // In all other cases.
	}

	// Constructor:
	Var(LPTSTR aVarName, VarEntry *aBuiltIn, UCHAR aScope)
		// The caller must ensure that aVarName is non-null.
		: mCharContents(sEmptyString) // Invariant: Anyone setting mCapacity to 0 must also set mContents to the empty string.
		// Doesn't need initialization: , mContentsInt64(NULL)
		, mByteLength(0) // This also initializes mAliasFor within the same union.
		, mHowAllocated(ALLOC_NONE)
		, mAttrib(VAR_ATTRIB_UNINITIALIZED) // Seems best not to init empty vars to VAR_ATTRIB_NOT_NUMERIC because it would reduce maintainability, plus finding out whether an empty var is numeric via IsNumeric() is a very fast operation.
		, mScope(aScope)
		, mName(aVarName) // Caller gave us a pointer to dynamic memory for this.
	{
		if (!aBuiltIn)
			mType = VAR_NORMAL;
		else if ((UINT_PTR)aBuiltIn->type.Get <= VAR_LAST_TYPE) // Relies on the fact that numbers less than VAR_LAST_TYPE can never realistically match the address of any function.
			mType = (VarTypeType)(UINT_PTR)aBuiltIn->type.Get;
		else
		{
			mType = VAR_VIRTUAL;
			mVV = &aBuiltIn->type;
		}
		mByteCapacity = 0;
		if (mType != VAR_NORMAL)
			mAttrib = 0; // Any vars that aren't VAR_NORMAL are considered initialized, by definition.
	}

	Var()
		//: Var(_T(""), NULL, 0) // Not supported by Visual C++ 2010.
		// Initialized as above:
		: mCharContents(sEmptyString)
		, mByteLength(0)
		, mAttrib(VAR_ATTRIB_UNINITIALIZED)
		// For anonymous/temporary variables:
		, mScope(VAR_LOCAL)
		, mName(_T(""))
		// Normally set as a result of !aBuiltIn:
		, mType(VAR_NORMAL)
		, mByteCapacity(0)
		// Vars constructed this way are for temporary use, and therefore must have mHowAllocated set
		// as below to prevent the use of SimpleHeap::Malloc().  Otherwise, each Var could allocate
		// some memory which cannot be freed until the program exits.
		, mHowAllocated(ALLOC_MALLOC)
	{
	}

	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new(size_t aBytes, void *p) {return p;}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete(void *aPtr, void *) {}
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

	__forceinline void MarkUninitialized()
	{
		// Relies on the fact that aliases can't point to other aliases (enforced by UpdateAlias()).
		Var &var = *(mType == VAR_ALIAS ? mAliasFor : this);
		mAttrib |= VAR_ATTRIB_UNINITIALIZED;
	}

	__forceinline void MaybeWarnUninitialized();

}; // class Var
#pragma pack(pop) // Calling pack with no arguments restores the default value (which is 8, but "the alignment of a member will be on a boundary that is either a multiple of n or a multiple of the size of the member, whichever is smaller.")
#pragma warning(pop)

inline void ResultToken::StealMem(Var *aVar)
// Caller must ensure that aVar->mType == VAR_NORMAL and aVar->mHowAllocated == ALLOC_MALLOC.
{
	VarSizeType length = aVar->Length(); // Must not be combined with the line below, as the compiler is free to evaluate parameters in whatever order.
	AcceptMem(aVar->StealMem(), length);
}

#endif
