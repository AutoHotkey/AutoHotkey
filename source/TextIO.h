#pragma once

#define TEXT_IO_BLOCK	4096

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
		: mFlags(0), mCodePage(-1), mLocale(NULL), mLength(0), mCapacity(0), mBuffer(NULL), mPos(NULL), mEOF(true)
	{
		SetCodePage(CP_ACP);
	}
	virtual ~TextStream()
	{
		if (mBuffer)
			free(mBuffer);
		if (mLocale)
			_free_locale(mLocale);
		// Close() isn't called here, it will rise a "pure virtual function call" exception.
	}

	bool Open(LPCTSTR aFileSpec, DWORD aFlags, UINT aCodePage = CP_ACP);
	void Close() { _Close(); mEOF = true; }

	DWORD ReadLine(LPTSTR aBuf, DWORD aBufLen)
	{
		LPTSTR bufEnd = aBuf + aBufLen, bufStart = aBuf;
		while (aBuf < bufEnd) {
			*aBuf = GetChar();
			if (*aBuf == TEOF)
				break;
			if (*aBuf++ == '\n')
				break;
		}
		if (aBuf < bufEnd)
			*aBuf = '\0';
		return (DWORD)(aBuf - bufStart);
	}
	DWORD Read(LPTSTR aBuf, DWORD aBufLen)
	{
		LPTSTR bufEnd = aBuf + aBufLen, bufStart = aBuf;
		while (aBuf < bufEnd) {
			*aBuf = GetChar();
			if (*aBuf == TEOF)
				break;
			aBuf++;
		}
		if (aBuf < bufEnd)
			*aBuf = '\0';
		return (DWORD)(aBuf - bufStart);
	}
	TCHAR GetChar()
	{
		TCHAR ch = ReadChar();
		if (ch == '\r' && (mFlags & (EOL_CRLF | EOL_ORPHAN_CR))) {
			TCHAR ch2 = ReadChar();
			if (ch2 != '\n') {
				mCache[1] = ch2;
				if (mFlags & EOL_ORPHAN_CR)
					return '\n';
			}
			else if (mFlags & EOL_CRLF)
				return '\n';
		}
		return ch;
	}

	DWORD Write(LPCTSTR aBuf, DWORD aBufLen = 0);

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
		if (!mCache[1] && (!mPos || mPos >= mBuffer + mLength)) {
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
			mCodePage = aCodePage;

			// mLocale is no longer relevant, so free it.
			if (mLocale)
				_free_locale(mLocale);

			if (aCodePage != CP_UTF8 && aCodePage != CP_UTF16 && aCodePage != CP_UTF7)
			{
				// Recreate locale for use with _ismbblead_l.
				char name_buf[16];
				*name_buf = '.';
				_ultoa(aCodePage, name_buf + 1, 10);
				mLocale = _create_locale(LC_ALL, name_buf);
			}
			else
				// _create_locale doesn't support Unicode; mLocale isn't needed for CP_UTF8 or CP_UTF16 anyway.
				mLocale = NULL;
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

	TCHAR ReadChar();

	DWORD Read(DWORD aReadSize = TEXT_IO_BLOCK)
	{
		ASSERT(aReadSize);
		if (mEOF)
			return 0;
		if (!mBuffer) {
			mBuffer = (BYTE *) malloc(TEXT_IO_BLOCK);
			mCapacity = TEXT_IO_BLOCK;
		}
		if (mLength + aReadSize > TEXT_IO_BLOCK)
			aReadSize = TEXT_IO_BLOCK - mLength;
		DWORD dwRead = _Read(mBuffer + mLength, (DWORD)aReadSize);
		if (dwRead) {
			mLength += dwRead;
			mPos = mBuffer;
		}
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
			memcpy(mBuffer, mPos, mLength);
			Read(TEXT_IO_BLOCK);
		}
		else
			return true;
		if (mLength < aReadSize)
			mEOF = true;
		return !mEOF;
	}

	DWORD mFlags;
	DWORD mLength;		// The length of available data in the buffer, in bytes.
	DWORD mCapacity;	// The capacity of the buffer, in bytes.
	UINT  mCodePage;
	_locale_t mLocale;
	
	bool  mEOF;
	union
	{
		CHAR  mWriteCharA;
		WCHAR mWriteCharW;
	};

	union
	{
		DWORD mCacheInt;
		TCHAR mCache[4 / sizeof(TCHAR)];
		CHAR  mCacheA[4];
#ifdef UNICODE
		WCHAR mCacheW[2];
#else
		// See ReadChar()
#endif
	};
	union // pointer to the next character to read in mBuffer
	{
		LPBYTE  mPos;
		LPSTR   mPosA;
		LPWSTR  mPosW;
	};
	union // used by buffered/translated IO. 
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
	virtual ~TextFile() { _Close(); }

	// Text IO methods from TextStream.
	using TextStream::Read;
	using TextStream::Write;

	// These methods are exported to provide binary file IO.
	DWORD   Read(LPVOID aBuffer, DWORD aBufSize)
	{
		RollbackFilePointer();
		DWORD dwRead = _Read(aBuffer, aBufSize);
		mEOF = _Tell() == _Length(); // binary IO is not buffered.
		return dwRead;
	}
	DWORD   Write(LPCVOID aBuffer, DWORD aBufSize) { RollbackFilePointer(); return _Write(aBuffer, aBufSize); }
	bool    Seek(__int64 aDistance, int aOrigin) { RollbackFilePointer(); return _Seek(aDistance, aOrigin); }
	__int64	Tell() { RollbackFilePointer(); return _Tell(); }
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
	HANDLE  Handle() { RollbackFilePointer(); return mFile; }
protected:
	virtual bool    _Open(LPCTSTR aFileSpec, DWORD aFlags);
	virtual void    _Close();
	virtual DWORD   _Read(LPVOID aBuffer, DWORD aBufSize);
	virtual DWORD   _Write(LPCVOID aBuffer, DWORD aBufSize);
	virtual bool    _Seek(__int64 aDistance, int aOrigin);
	virtual __int64	_Tell() const;
	virtual __int64 _Length() const;

	void RollbackFilePointer()
	{
		if (mPos) // Text reading was used
		{
			mCacheInt = 0; // cache is cleared to prevent unexpected results.

			// Discards the buffer and rollback the file pointer.
			ptrdiff_t offset = (mPos - mBuffer) - mLength; // should be a value <= 0
			_Seek(offset, SEEK_CUR);
			mPos = NULL;
			mLength = 0;
		}
	}
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
