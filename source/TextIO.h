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

// VS2005 and later comes with Unicode stream IO in the C runtime library, but it doesn't work very well.
// For example, it can't read the files encoded in "system codepage" by using
// wide-char version of function such as fgetws(). The characters were not translate to
// UTF-16 properly. Although we can create some workarounds for it, but that will make
// the codes even hard to maintained.
class TextStream
{
public:
	enum {
		// open modes
		  READ
		, WRITE
		, APPEND

		// EOL translations
		, EOL_CRLF = 0x00000004 // read: CRLF to LF. write: LF to CRLF.
		, EOL_ORPHAN_CR = 0x00000008 // read: CR to LF (when the next character is NOT LF)

		// write byte order mark when open for write
		, BOM_UTF8 = 0x00000010
		, BOM_UTF16 = 0x00000020

		//, OVERWRITE = 0x00000040
	};

	TextStream()
		: mFlags(0), mCodePage(CP_ACP), mLength(0), mCapacity(0), mBuffer(NULL), mPosA(NULL), mEOF(true)
	{
	}
	virtual ~TextStream()
	{
		if (mBuffer)
			free(mBuffer);
		// Close() isn't called here, it will rise a "pure virtual function call" exception.
	}

	bool Open(LPCTSTR aFileSpec, DWORD aFlags, UINT aCodePage = CP_ACP);
	void Close() { _Close(); mEOF = true; }

	DWORD ReadLine(LPWSTR aBuf, DWORD aBufLen)
	{
		LPWSTR bufEnd = aBuf + aBufLen, bufStart = aBuf;
		while (aBuf < bufEnd) {
			*aBuf = GetCharW();
			if (*aBuf == WEOF)
				break;
			if (*aBuf++ == '\n')
				break;
		}
		if (aBuf < bufEnd)
			*aBuf = '\0';
		return (DWORD)(aBuf - bufStart);
	}
	WCHAR GetCharW()
	{
		WCHAR ch = ReadCharW();
		if (ch == '\r' && (mFlags & (EOL_CRLF | EOL_ORPHAN_CR))) {
			WCHAR ch2 = ReadCharW();
			if (ch2 != '\n') {
				if (mFlags & EOL_ORPHAN_CR) {
					mCacheW[1] = ch2;
					return '\n';
				}
			}
			else if (mFlags & EOL_CRLF)
				return '\n';
		}
		return ch;
	}
	//CHAR GetCharA();
	TCHAR GetChar() { return W_OR_A(GetChar)(); }

	DWORD Write(LPCWSTR aBuf, DWORD aBufLen);
	//DWORD Write(LPCSTR aBuf, DWORD aBufLen);
	int FormatV(LPCWSTR fmt, va_list ap)
	{
		CStringW str;
		str.FormatV(fmt, ap);
		return Write(str, str.GetLength()) / sizeof(TCHAR);
	}
	int Format(LPCWSTR fmt, ...)
	{
		va_list ap;
		va_start(ap, fmt);
		return FormatV(fmt, ap);
	}
	//int FormatV(LPCSTR fmt, va_list ap);
	//int Format(LPCSTR fmt, ...);

	bool AtEOF() const { return mEOF; }
	//void SetCodePage(UINT aCodePage) { mCodePage = aCodePage; }
	//UINT GetCodePage() { return mCodePage; }
protected:
	// IO abstraction
	virtual bool    _Open(LPCTSTR aFileSpec, DWORD aFlags) = 0;
	virtual void    _Close() = 0;
	virtual DWORD   _Read(LPVOID aBuffer, DWORD aBufSize) = 0;
	virtual DWORD   _Write(LPCVOID aBuffer, DWORD aBufSize) = 0;
	virtual bool    _Seek(long aDistance, int aOrigin) = 0;
	virtual __int64 _Length() = 0;

	void Alloc(DWORD aBufSize = TEXT_IO_BLOCK)
	{
		if (!mBuffer || aBufSize > mCapacity) {
			mBuffer = (BYTE *) realloc(mBuffer, aBufSize);
			mCapacity = aBufSize;
		}
	}
	WCHAR ReadCharW();
	//CHAR ReadCharA();
	TCHAR ReadChar() { return W_OR_A(ReadChar)(); }
	DWORD Read(DWORD aReadSize = TEXT_IO_BLOCK)
	{
		ASSERT(aReadSize);
		if (mEOF)
			return 0;
		Alloc();
		if (mLength + aReadSize > TEXT_IO_BLOCK)
			aReadSize = TEXT_IO_BLOCK - mLength;
		DWORD dwRead = _Read(mBuffer + mLength, aReadSize);
		if (dwRead) {
			mLength += dwRead;
			mPosA = mBufferA;
		}
		else
			mEOF = true;
		return dwRead;
	}
	bool ReadAtLeast(DWORD aReadSize)
	{
		if (!mPosA)
			Read(TEXT_IO_BLOCK);
		else if (mPosA > mBufferA + mLength - aReadSize) {
			ASSERT( (DWORD)(mPosA - mBufferA) <= mLength );
			mLength -= (DWORD)(mPosA - mBufferA);
			memcpy(mBuffer, mPosA, mLength);
			Read(TEXT_IO_BLOCK);
		}
		else
			return true;
		if (mLength < aReadSize)
			mEOF = true;
		return !mEOF;
	}

	DWORD mFlags;
	UINT  mCodePage;
	DWORD mLength;		// The length of available data in the buffer, in bytes.
	DWORD mCapacity;	// The capacity of the buffer, in bytes.
	bool  mEOF;
	union {
		CHAR  mCacheA[4];
		WCHAR mCacheW[2];
		TCHAR mCache[4 / sizeof(TCHAR)];
	};
	union {
		LPSTR   mPosA;
		LPWSTR  mPosW;
	};
	union
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
protected:
	virtual bool    _Open(LPCTSTR aFileSpec, DWORD aFlags);
	virtual void    _Close();
	virtual DWORD   _Read(LPVOID aBuffer, DWORD aBufSize);
	virtual DWORD   _Write(LPCVOID aBuffer, DWORD aBufSize);
	virtual bool    _Seek(long aDistance, int aOrigin);
	virtual __int64 _Length();
private:
	HANDLE mFile;
};
