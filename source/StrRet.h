#pragma once

class StrRet
{
	LPTSTR mValue = nullptr, mCallerBuf = nullptr, mAllocated = nullptr;
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
	LPTSTR Copy(LPCTSTR s, size_t n)
	{
		ASSERT(!mAllocated && !mValue);
		if (!n)
			return nullptr;
		LPTSTR buf = Alloc(n);
		if (buf)
		{
			tmemcpy(buf, s, n);
			buf[n] = '\0';
		}
		return buf;
	}
	
	// Set value to a copy of the string in memory that can be returned to the caller.
	LPTSTR Copy(LPCTSTR s)
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

	LPTSTR Value()
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
	}

	// Set the return value.
	// s must be in static memory or CallerBuf().
	void SetStatic(LPTSTR s)
	{
		ASSERT(!mValue);
		mValue = s;
	}

	// Set the return value and length.
	// s must be in static memory or CallerBuf().
	void SetStatic(LPTSTR s, size_t n)
	{
		ASSERT(!mValue);
		mValue = s;
		mLength = n;
	}

	// Set the length of the string which has been written into the buffer returned by Alloc().
	// Calling this is optional.
	void SetLength(size_t n)
	{
		ASSERT(mValue);
		mLength = n;
	}
};
