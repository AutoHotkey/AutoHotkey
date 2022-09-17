#pragma once

#include "util.h"

class StrRet
{
	LPCTSTR mValue = nullptr;
	LPTSTR mCallerBuf = nullptr, mAllocated = nullptr;
	size_t mLength = -1;

public:
	static constexpr size_t CallerBufSize = 256;

	StrRet(LPTSTR aBuf) : mCallerBuf(aBuf) {}

	// Allocate a buffer large enough for n characters plus a null-terminator,
	// and sets it as the return value.  Must be called only once.
	LPTSTR Alloc(size_t n)
	{
		ASSERT(!mAllocated && !mValue && n);
		LPTSTR buf;
		if (n < CallerBufSize)
			buf = mCallerBuf;
		else
			buf = mAllocated = tmalloc(n + 1);
		mValue = buf;
		return buf;
	}

	// Set value to a copy of the string in memory that can be returned to the caller.
	bool Copy(LPCTSTR s, size_t n)
	{
		ASSERT(!mAllocated && !mValue);
		if (!n)
			return true;
		LPTSTR buf = Alloc(n);
		if (buf)
		{
			tmemcpy(buf, s, n);
			buf[n] = '\0';
		}
		return buf;
	}
	
	// Set value to a copy of the string in memory that can be returned to the caller.
	bool Copy(LPCTSTR s)
	{
		return Copy(s, _tcslen(s));
	}

	// Returns a buffer of size StrRet::CallerBufSize allocated by the caller for use by the callee.
	LPTSTR CallerBuf()
	{
		return mCallerBuf;
	}

	bool UsedMalloc()
	{
		ASSERT(!mAllocated || mAllocated == mValue);
		return mAllocated != nullptr;
	}

	LPCTSTR Value()
	{
		ASSERT(!mAllocated || mAllocated == mValue);
		return mValue;
	}

	size_t Length()
	{
		return mLength;
	}

	// Used by callee to declare that the value should be empty, for clarity and
	// in case the default state is ever interpreted as something other than "".
	void SetEmpty()
	{
		ASSERT(!mValue);
	}

	// Set the return value.
	// s must be in static memory, such as a literal string.
	void SetStatic(LPCTSTR s)
	{
		ASSERT(!mValue);
		mValue = s;
	}

	// Set the return value and length.
	// s must be in static memory, such as a literal string.
	void SetStatic(LPCTSTR s, size_t n)
	{
		ASSERT(!mValue);
		mValue = s;
		mLength = n;
	}

	// Set the return value.
	// s must be in memory that will persist until the caller has an opportunity to
	// copy it, such as in CallerBuf().
	// Alias of SetStatic(), used to show that retaining the pointer indefinitely
	// wouldn't be safe.
	void SetTemp(LPCTSTR s)
	{
		SetStatic(s);
	}
	void SetTemp(LPCTSTR s, size_t n)
	{
		SetStatic(s, n);
	}

	// Set the length of the string which has been written into the buffer returned by Alloc().
	// Calling this is optional.
	void SetLength(size_t n)
	{
		ASSERT(mValue);
		mLength = n;
	}
};
