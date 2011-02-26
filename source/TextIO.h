#pragma once

#define TEXT_IO_BLOCK	8192

#ifndef CP_UTF16
#define CP_UTF16		1200 // the codepage of UTF-16LE
#endif

#ifdef UNICODE
#define W_OR_A(n)	n##W
#else
#define W_OR_A(n)	n##A
#endif

#define TEOF ((TCHAR)EOF)

#include <locale.h> // For _locale_t, _create_locale and _free_locale.

extern UINT g_ACP;

// VS2005 and later come with Unicode stream IO in the C runtime library, but it doesn't work very well.
// For example, it can't read the files encoded in "system codepage" by using
// wide-char version of the functions such as fgetws(). The characters were not translated to
// UTF-16 properly. Although we can create some workarounds for it, but that will make
// the codes even hard to maintain.
class TextStream
{
public:
	enum {
		// open modes
		  READ
		, WRITE
		, APPEND
		, UPDATE
		, USEHANDLE = 0x10000000 // Used by BIF_FileOpen/FileObject. High value avoids conflict with the flags below, which can't change because it would break scripts.
		, ACCESS_MODE_MASK = READ|WRITE|APPEND|UPDATE|USEHANDLE

		// EOL translations
		, EOL_CRLF = 0x00000004 // read: CRLF to LF. write: LF to CRLF.
		, EOL_ORPHAN_CR = 0x00000008 // read: CR to LF (when the next character isn't LF)

		// write byte order mark when open for write
		, BOM_UTF8 = 0x00000010
		, BOM_UTF16 = 0x00000020

		// shared accesses
		, SHARE_READ = 0x00000100
		, SHARE_WRITE = 0x00000200
		, SHARE_DELETE = 0x00000400
		, SHARE_ALL = SHARE_READ|SHARE_WRITE|SHARE_DELETE
	};

	TextStream()
		: mFlags(0), mCodePage(-1), mLength(0), mBuffer(NULL), mPos(NULL), mEOF(true)
	{
		SetCodePage(CP_ACP);
	}
	virtual ~TextStream()
	{
		if (mBuffer)
			free(mBuffer);
		//if (mLocale)
		//	_free_locale(mLocale);
		// Close() isn't called here, it will rise a "pure virtual function call" exception.
	}

	bool Open(LPCTSTR aFileSpec, DWORD aFlags, UINT aCodePage = CP_ACP);
	void Close()
	{
		FlushWriteBuffer();
		_Close();
		mEOF = true;
	}

	DWORD Write(LPCTSTR aBuf, DWORD aBufLen = 0);
	DWORD Write(LPCVOID aBuf, DWORD aBufLen);
	DWORD Read(LPTSTR aBuf, DWORD aBufLen, int aNumLines = 0);
	DWORD Read(LPVOID aBuf, DWORD aBufLen);

	DWORD ReadLine(LPTSTR aBuf, DWORD aBufLen)
	{
		return Read(aBuf, aBufLen, 1);
	}

	INT_PTR FormatV(LPCTSTR fmt, va_list ap)
	{
		CString str;
		str.FormatV(fmt, ap);
		return Write(str, (DWORD)str.GetLength()) / sizeof(TCHAR);
	}
	INT_PTR Format(LPCTSTR fmt, ...)
	{
		va_list ap;
		va_start(ap, fmt);
		return FormatV(fmt, ap);
	}

	bool AtEOF()
	// The behavior isn't the same with feof().
	// It checks the position of the file pointer, too.
	{
		if (mEOF)
			return true;
		if (!mPos || mPos >= mBuffer + mLength) {
			__int64 pos = _Tell();
			if (pos < 0 || pos >= _Length())
				mEOF = true;
		}
		return mEOF;
	}

	void SetCodePage(UINT aCodePage)
	{
		if (aCodePage == CP_ACP)
			aCodePage = g_ACP; // Required by _create_locale.
		//if (!IsValidCodePage(aCodePage)) // Returns FALSE for UTF-16 and possibly other valid code pages, so leave it up to the user to pass a valid codepage.
			//return;

		if (mCodePage != aCodePage)
		{
			// Resist temptation to do the following as a way to avoid having an odd number of bytes in
			// the buffer, since it breaks non-seeking devices and actually isn't sufficient for cases
			// where the caller uses raw I/O in addition to text I/O (e.g. read one byte then read text).
			//RollbackFilePointer();

			mCodePage = aCodePage;
			if (!GetCPInfo(aCodePage, &mCodePageInfo))
				mCodePageInfo.LeadByte[0] = NULL;
		}
	}
	UINT GetCodePage() { return mCodePage; }
	DWORD GetFlags() { return mFlags; }

protected:
	// IO abstraction
	virtual bool    _Open(LPCTSTR aFileSpec, DWORD aFlags) = 0;
	virtual void    _Close() = 0;
	virtual DWORD   _Read(LPVOID aBuffer, DWORD aBufSize) = 0;
	virtual DWORD   _Write(LPCVOID aBuffer, DWORD aBufSize) = 0;
	virtual bool    _Seek(__int64 aDistance, int aOrigin) = 0;
	virtual __int64	_Tell() const = 0;
	virtual __int64 _Length() const = 0;
	
	void RollbackFilePointer()
	{
		if (mPos) // Buffered reading was used.
		{
			// Discard the buffer and rollback the file pointer.
			ptrdiff_t offset = (mPos - mBuffer) - mLength; // should be a value <= 0
			_Seek(offset, SEEK_CUR);
			// Callers expect the buffer to be cleared (e.g. to be reused for buffered writing), so if
			// _Seek fails, the data is simply discarded.  This can probably only happen for non-seeking
			// devices such as pipes or the console, which won't typically be both read from and written to:
			mPos = NULL;
			mLength = 0;
		}
	}
	
	void FlushWriteBuffer()
	{
		if (mLength && !mPos)
		{
			// Flush write buffer.
			_Write(mBuffer, mLength);
			mLength = 0;
		}
		mLastWriteChar = 0;
	}

	bool PrepareToWrite()
	{
		if (!mBuffer)
			mBuffer = (BYTE *) malloc(TEXT_IO_BLOCK);
		else if (mPos) // Buffered reading was used.
			RollbackFilePointer();
		return mBuffer != NULL;
	}

	bool PrepareToRead()
	{
		FlushWriteBuffer();
		return true;
	}

	template<typename TCHR>
	DWORD WriteTranslateCRLF(TCHR *aBuf, DWORD aBufLen); // Used by TextStream::Write(LPCSTR,DWORD).

	// Functions for populating the read buffer.
	DWORD Read(DWORD aReadSize = TEXT_IO_BLOCK)
	{
		ASSERT(aReadSize);
		if (mEOF)
			return 0;
		if (!mBuffer) {
			mBuffer = (BYTE *) malloc(TEXT_IO_BLOCK);
			if (!mBuffer)
				return 0;
		}
		if (mLength + aReadSize > TEXT_IO_BLOCK)
			aReadSize = TEXT_IO_BLOCK - mLength;
		DWORD dwRead = _Read(mBuffer + mLength, aReadSize);
		if (dwRead)
			mLength += dwRead;
		else
			mEOF = true;
		return dwRead;
	}
	bool ReadAtLeast(DWORD aReadSize)
	{
		if (!mPos)
			Read(TEXT_IO_BLOCK);
		else if (mPos > mBuffer + mLength - aReadSize) {
			ASSERT( (DWORD)(mPos - mBuffer) <= mLength );
			mLength -= (DWORD)(mPos - mBuffer);
			memmove(mBuffer, mPos, mLength);
			Read(TEXT_IO_BLOCK);
		}
		else
			return true;
		mPos = mBuffer;
		if (mLength < aReadSize)
			mEOF = true;
		return !mEOF;
	}

	__declspec(noinline) bool IsLeadByte(BYTE b) // noinline benchmarks slightly faster.
	{
		for (int i = 0; i < _countof(mCodePageInfo.LeadByte) && mCodePageInfo.LeadByte[i]; i += 2)
			if (b >= mCodePageInfo.LeadByte[i] && b <= mCodePageInfo.LeadByte[i+1])
				return true;
		return false;
	}

	DWORD mFlags;
	DWORD mLength;		// The length of available data in the buffer, in bytes.
	UINT  mCodePage;
	CPINFO mCodePageInfo;
	
	bool  mEOF;
	TCHAR mLastWriteChar;

	union // Pointer to the next character to read in mBuffer.
	{
		LPBYTE  mPos;
		LPSTR   mPosA;
		LPWSTR  mPosW;
	};
	union // Used by buffered/translated IO to hold raw file data. 
	{
		LPBYTE  mBuffer;
		LPSTR   mBufferA;
		LPWSTR  mBufferW;
	};
};



class TextFile : public TextStream
{
public:
	TextFile() : mFile(INVALID_HANDLE_VALUE) {}
	virtual ~TextFile() { FlushWriteBuffer(); _Close(); }

	// Text IO methods from TextStream.
	using TextStream::Read;
	using TextStream::Write;

	// These methods are exported to provide binary file IO.
	bool    Seek(__int64 aDistance, int aOrigin)
	{
		RollbackFilePointer();
		FlushWriteBuffer();
		return _Seek(aDistance, aOrigin);
	}
	__int64	Tell()
	{
		return _Tell() + (mPos ? mPos - (mBuffer + mLength) : (ptrdiff_t)mLength);
	}
	__int64 Length() { return _Length(); }
	__int64 Length(__int64 aLength)
	{
		__int64 pos = Tell();
		if (!_Seek(aLength, SEEK_SET) || !SetEndOfFile(mFile))
			return -1;
		// Make sure we do not extend the file again.
		_Seek(min(aLength, pos), SEEK_SET);
		return _Length();
	}
	HANDLE  Handle() { RollbackFilePointer(); FlushWriteBuffer(); return mFile; }
protected:
	virtual bool    _Open(LPCTSTR aFileSpec, DWORD aFlags);
	virtual void    _Close();
	virtual DWORD   _Read(LPVOID aBuffer, DWORD aBufSize);
	virtual DWORD   _Write(LPCVOID aBuffer, DWORD aBufSize);
	virtual bool    _Seek(__int64 aDistance, int aOrigin);
	virtual __int64	_Tell() const;
	virtual __int64 _Length() const;

private:
	HANDLE mFile;
};



// TextMem is intended to attach a memory block, which provides code pages and end-of-line conversions (CRLF <-> LF).
// It is used for reading the script data in compiled script.
// Note that TextMem doesn't have any ability to write and seek.
class TextMem : public TextStream
{
public:
	// TextMem tmem;
	// tmem.Open(TextMem::Buffer(buffer, length), EOL_CRLF, CP_ACP);
	struct Buffer
	{
		Buffer(LPVOID aBuf = NULL, DWORD aBufLen = 0, bool aOwned = true)
			: mBuffer(aBuf), mLength(aBufLen), mOwned(aOwned)
		{}
		operator LPCTSTR() const { return (LPCTSTR) this; }
		LPVOID mBuffer;
		DWORD mLength;
		bool mOwned;	// If true, the memory will be freed by _Close().
	};

	TextMem() : mOwned(false) {}
	virtual ~TextMem() { _Close(); }
protected:
	virtual bool    _Open(LPCTSTR aFileSpec, DWORD aFlags);
	virtual void    _Close();
	virtual DWORD   _Read(LPVOID aBuffer, DWORD aBufSize);
	virtual DWORD   _Write(LPCVOID aBuffer, DWORD aBufSize);
	virtual bool    _Seek(__int64 aDistance, int aOrigin);
	virtual __int64	_Tell() const;
	virtual __int64 _Length() const;
private:
	bool mOwned;
};
