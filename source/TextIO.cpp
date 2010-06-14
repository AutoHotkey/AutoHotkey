#include "stdafx.h"
#include "TextIO.h"
#include "script.h"
#include "script_object.h"
#include <mbctype.h> // For _ismbblead_l.

UINT g_ACP = GetACP(); // Requires a reboot to change.
#define INVALID_CHAR UorA(0xFFFD, '?')

//
// TextStream
//
bool TextStream::Open(LPCTSTR aFileSpec, DWORD aFlags, UINT aCodePage)
{
	mLength = 0; // Set the default value here so _Open() can change it.
	if (!_Open(aFileSpec, aFlags))
		return false;

	SetCodePage(aCodePage == CP_ACP ? g_ACP : aCodePage);
	mFlags = aFlags;
	mEOF = false;
	mCacheInt = 0;

	int mode = aFlags & 3;
	if (mode != TextStream::WRITE) {
		// Detect UTF-8 and UTF-16LE BOMs
		if (mLength < 3)
			Read(3);
		if (mLength >= 2) {
			if (mBuffer[0] == 0xFF && mBuffer[1] == 0xFE) {
				SetCodePage(CP_UTF16);
				mPosW = mBufferW + 1;
			}
			else if (mBuffer[0] == 0xEF && mBuffer[1] == 0xBB) {
				if (mLength >= 3 && mBuffer[2] == 0xBF) {
					SetCodePage(CP_UTF8);
					mPosA = mBufferA + 3;
				}
			}
		}
	}
	if (mode == TextStream::WRITE || (mode == TextStream::APPEND || mode == TextStream::UPDATE) && _Length() == 0) {
		if (aFlags & BOM_UTF8)
			_Write("\xEF\xBB\xBF", 3);
		else if (aFlags & BOM_UTF16)
			_Write("\xFF\xFE", 2);
	}
	else if (mode == TextStream::APPEND)
		_Seek(0, SEEK_END);

	return true;
}



TCHAR TextStream::ReadChar()
// Fetch exactly one character from the stream, this may slow down the reading, though.
// But there are some reasons to do this:
//   1. If some invalid bytes are encountered while reading, we can detect the problem, drop those bytes and then continue to read.
//   2. It allows partial bytes of a multi-byte character at the end of the input buffer, so we don't need to load the whole file into memory.
//   3. We can also apply EOL (CR, LF, CR/LF) handling. (see GetChar())
{
	if (mCache[1]) { // surrogate pair or standalone CR
		TCHAR ret = mCache[1];
		mCache[1] = 0;
		return ret;
	}
	if (mEOF)
		return TEOF;

	if (mCodePage == CP_UTF16) {
		if (!ReadAtLeast(sizeof(WCHAR)))
			return TEOF;
#ifdef UNICODE
		return *mPosW++;
#else
		if (WideCharToMultiByte(g_ACP, 0, mPosW++, 1, mCache, 2, NULL, NULL))
			return mCache[0];
#endif
	}
	else {
		if (!ReadAtLeast(sizeof(CHAR)))
			return TEOF;

#ifndef UNICODE
		if (mCodePage == g_ACP)
			// No conversion necessary.
			return *mPosA++;
		WCHAR mCacheW[1]; // Named thusly to simplify the code, declared here as it cannot refer to the same memory as mCache.
#endif

		int iBytes;
		if (mCodePage == CP_UTF8)
		{
			if (*mPos < 0x80)
				// single byte UTF-8 character
				return (TCHAR) *mPosA++;
			// The size in bytes of UTF-8 characters.
			if ((*mPos & 0xE0) == 0xC0)
				iBytes = 2;
			else if ((*mPos & 0xF0) == 0xE0)
				iBytes = 3;
			else if ((*mPos & 0xF8) == 0xF0)
				iBytes = 4;
			else {
				// Invalid in current UTF-8 standard.
				mPosA++;
				return INVALID_CHAR;
			}
		}
		else if (_ismbblead_l(*mPos, mLocale))
			iBytes = 2;
		else
			iBytes = 1;

		if (!ReadAtLeast(iBytes))
			return TEOF;

		if (MultiByteToWideChar(mCodePage, MB_ERR_INVALID_CHARS, mPosA, iBytes, mCacheW, UorA(2,1))) // UorA usage: ANSI build currently cannot support UTF-16 surrogate pairs.
		{
			mPosA += iBytes;
#ifndef UNICODE
			// In this case mCacheW is a local variable, not a member of the union with mCache.
			if (WideCharToMultiByte(g_ACP, 0, mCacheW, 1, mCache, 2, NULL, NULL))
#endif
				return mCache[0];
		}
		else
			mPosA++; // ignore invalid byte
	}
	return INVALID_CHAR; // invalid character
}



DWORD TextStream::Write(LPCTSTR aBuf, DWORD aBufLen)
// NOTE: This method doesn't returns the number of characters are written,
// instead, it returns the number of bytes. Because this should be faster and the write operations
// should never fail unless it encounters a critical error (e.g. low disk space.).
// Therefore, the amount of characters are the same with aBufLen or wcslen(aBuf) most likely.
// The callers have knew that already.
{
	if (aBufLen == 0)
		aBufLen = _tcslen(aBuf);
	if (mCodePage == CP_UTF16)
	{
#ifdef UNICODE
		// Alias to simplify the code.  Compiler optimizations will probably make up for it.
		LPCWSTR buf_w = aBuf;
#else
		CStringWCharFromChar buf_w(aBuf, aBufLen, g_ACP);
		aBufLen = buf_w.GetLength();
#endif
		if (mFlags & EOL_CRLF)
		{
			DWORD dwWritten = 0;
			DWORD i;
			for (i = 0; i < aBufLen; i++)
			{
				if (buf_w[i] == '\n' && mWriteCharW != '\r')
					dwWritten += _Write(L"\r\n", sizeof(WCHAR) * 2);
				else
					dwWritten += _Write(&buf_w[i], sizeof(WCHAR));
				mWriteCharW = buf_w[i];
			}
			return dwWritten;
		}

		return _Write(buf_w, aBufLen * sizeof(WCHAR));
	}
	else
	{
		LPCSTR str;
		int len;
		CStringA buf_a;
#ifndef UNICODE
		// Since it's probably the most common case, optimize for the active codepage:
		if (g_ACP == mCodePage)
		{
			str = aBuf;
			len = (int)aBufLen;
		}
		else
#endif
		{
			if (mCodePage == CP_UTF8)
				StringTCharToUTF8(aBuf, buf_a, aBufLen);
			else
			{
#ifdef UNICODE
				StringWCharToChar(aBuf, buf_a, aBufLen, '?', mCodePage);
#else
				// Since mCodePage is not the active codepage, we need to convert.
				CStringWCharFromChar buf_w(aBuf, aBufLen, g_ACP);
				StringWCharToChar(buf_w.GetString(), buf_a, buf_w.GetLength(), '?', mCodePage);
#endif
			}
			str = buf_a.GetString();
			len = buf_a.GetLength();
		}
		
		if (mFlags & EOL_CRLF)
		{
			DWORD dwWritten = 0;
			int i;
			for (i = 0; i < len; i++)
			{
				if (str[i] == '\n' && mWriteCharA != '\r')
					dwWritten += _Write("\r\n", sizeof(CHAR) * 2);
				else
					dwWritten += _Write(&str[i], sizeof(CHAR) * 1);
				mWriteCharA = str[i];
			}
			return dwWritten;
		}

		return _Write(str, sizeof(CHAR) * len);
	}
}


//
// TextFile
//
bool TextFile::_Open(LPCTSTR aFileSpec, DWORD aFlags)
{
	_Close();
	DWORD dwDesiredAccess, dwShareMode, dwCreationDisposition;
	switch (aFlags & 3) {
		case TextStream::READ:
			dwDesiredAccess = GENERIC_READ;
			dwCreationDisposition = OPEN_EXISTING;
			break;
		case TextStream::WRITE:
			dwDesiredAccess = GENERIC_WRITE;
			dwCreationDisposition = CREATE_ALWAYS;
			break;
		case TextStream::APPEND:
			dwDesiredAccess = GENERIC_WRITE | GENERIC_READ;
			dwCreationDisposition = OPEN_ALWAYS;
			break;
		case TextStream::UPDATE:
			dwDesiredAccess = GENERIC_WRITE | GENERIC_READ;
			dwCreationDisposition = OPEN_EXISTING;
			break;
	}
	dwShareMode = ((aFlags >> 8) & (FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE));

	// FILE_FLAG_SEQUENTIAL_SCAN is set, as sequential accesses are quite common for text files handling.
	mFile = CreateFile(aFileSpec, dwDesiredAccess, dwShareMode, NULL, dwCreationDisposition,
		(aFlags & (EOL_CRLF | EOL_ORPHAN_CR)) ? FILE_FLAG_SEQUENTIAL_SCAN : 0, NULL);

	return mFile != INVALID_HANDLE_VALUE;
}

void TextFile::_Close()
{
	if (mFile != INVALID_HANDLE_VALUE) {
		CloseHandle(mFile);
		mFile = INVALID_HANDLE_VALUE;
	}
}

DWORD TextFile::_Read(LPVOID aBuffer, DWORD aBufSize)
{
	DWORD dwRead = 0;
	ReadFile(mFile, aBuffer, aBufSize, &dwRead, NULL);
	return dwRead;
}

DWORD TextFile::_Write(LPCVOID aBuffer, DWORD aBufSize)
{
	DWORD dwWritten = 0;
	WriteFile(mFile, aBuffer, aBufSize, &dwWritten, NULL);
	return dwWritten;
}

bool TextFile::_Seek(__int64 aDistance, int aOrigin)
{
	// see AtEOF()
	if (aOrigin != SEEK_END || aDistance != 0)
		mEOF = false;
	return !!SetFilePointerEx(mFile, *((PLARGE_INTEGER) &aDistance), NULL, aOrigin);
}

__int64 TextFile::_Tell() const
{
	LARGE_INTEGER in = {0}, out;
	SetFilePointerEx(mFile, in, &out, FILE_CURRENT);
	return out.QuadPart;
}

__int64 TextFile::_Length() const
{
	LARGE_INTEGER size;
	GetFileSizeEx(mFile, &size);
	return size.QuadPart;
}

#ifdef CONFIG_EXPERIMENTAL
// FileObject: exports TextFile interfaces to the scripts.
class FileObject : public Object
{
	FileObject() {}
	~FileObject() {}

	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
	// Reference: MetaObject::Invoke
	{
		// Allow script-defined behaviour to take precedence:
		ResultType r = Object::Invoke(aResultToken, aThisToken, aFlags, aParam, aParamCount);

		if (r == INVOKE_NOT_HANDLED && IS_INVOKE_CALL && aParam[0]->symbol == SYM_OPERAND)
		{
			aResultToken.symbol = SYM_INTEGER; // Set default return type.

			LPTSTR field = TokenToString(*aParam[0], NULL);
			if (!_tcsnicmp(field, _T("Read"), 4))
			{
				if (!field[4]) // Read
				{
					if (aParamCount == 2)
					{
						DWORD length = (DWORD) TokenToInt64(*aParam[1]);
						if (length <= MAX_NUMBER_LENGTH)
						{
							aResultToken.symbol = SYM_STRING;
							aResultToken.marker = aResultToken.buf;
							mFile.Read(aResultToken.marker, length);
							aResultToken.marker[length] = '\0';
						}
						else
						{
							if (!(aResultToken.circuit_token = (ExprTokenType *)tmalloc(length + 1))) // Out of memory.
								return FAIL;
							aResultToken.symbol = SYM_STRING;
							aResultToken.marker = (LPTSTR) aResultToken.circuit_token; // Store the address of the result for the caller.
							length = mFile.Read(aResultToken.marker, length);
							aResultToken.marker[length] = '\0';
							aResultToken.buf = (LPTSTR)(size_t) length; // MANDATORY FOR USERS OF CIRCUIT_TOKEN: "buf" is being overloaded to store the length for our caller.
						}
						return OK;
					}
				}
				else if (!_tcsicmp(field + 4, _T("Line"))) // ReadLine
				{
					if (aParamCount == 1)
					{
						if (!(aResultToken.circuit_token = (ExprTokenType *)tmalloc(READ_FILE_LINE_SIZE)))
							return FAIL;
						aResultToken.symbol = SYM_STRING;
						aResultToken.marker = (LPTSTR) aResultToken.circuit_token; // Store the address of the result for the caller.
						aResultToken.buf = (LPTSTR)(size_t) mFile.ReadLine(aResultToken.marker, READ_FILE_LINE_SIZE - 1); // MANDATORY FOR USERS OF CIRCUIT_TOKEN: "buf" is being overloaded to store the length for our caller.
						aResultToken.marker[READ_FILE_LINE_SIZE - 1] = '\0'; // Prevent buffer overrun for very long lines.
						return OK;
					}
				}
			}
			else if (!_tcsicmp(field, _T("Write"))) // Write
			{
				if (aParamCount == 2)
				{
					aResultToken.value_int64 = mFile.Write(TokenToString(*aParam[1], aResultToken.buf));
					return OK;
				}
			}
			else if (!_tcsnicmp(field, _T("Raw"), 3)) // raw mode: binary IO
			{
				if (aParamCount == 3)
				{
					int iReadWrite = 0;
					
					if (!_tcsicmp(field + 3, _T("Read"))) // RawRead
						iReadWrite = 1;
					else if (!_tcsicmp(field + 3, _T("Write"))) // RawWrite
						iReadWrite = 2;
					if (iReadWrite)
					// Reference: BIF_NumGet
					{
						size_t target; // Don't make target a pointer-type because the integer offset might not be a multiple of 4 (i.e. the below increments "target" directly by "offset" and we don't want that to use pointer math).
						ExprTokenType &target_token = *aParam[1];
						if (target_token.symbol == SYM_VAR) // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
						{
							target = (size_t)target_token.var->Contents(); // Although Contents(TRUE) will force an update of mContents if necessary, it very unlikely to be necessary here because we're about to fetch a binary number from inside mContents, not a normal/text number.
						}
						else
							target = (size_t)TokenToInt64(target_token);

						if (target < 1024) // Basic sanity check to catch incoming raw addresses that are zero or blank.
						{
							aResultToken.value_int64 = 0;
							return OK;
						}

						size_t size = (size_t) TokenToInt64(*aParam[2]);

						// The user request a size larger than the variable.
						if (target_token.symbol == SYM_VAR && size > target_token.var->ByteCapacity())
						{
							// When reading expand the target variable if needed.
							if(iReadWrite == 1)
							{
								if(!target_token.var->SetCapacity(size, false, false))
								{
									aResultToken.value_int64 = 0;
									return OK;
								}
								target = (size_t)target_token.var->Contents(); // See comment above.
							}
							else
							{
								aResultToken.value_int64 = 0;
								return OK;
							}
						}

						aResultToken.value_int64 = (iReadWrite == 1) ? mFile.Read((LPVOID) target, size) : mFile.Write((LPCVOID) target, size);
						return OK;
					}
				}
			}
			else if (!_tcsicmp(field, _T("Seek"))) // Seek
			{
				if (aParamCount == 3)
				{
					aResultToken.value_int64 = mFile.Seek(TokenToInt64(*aParam[1]), (int) TokenToInt64(*aParam[2])) ? 1 : 0;
					return OK;
				}
				else if (aParamCount == 2)
				{
					aResultToken.value_int64 = mFile.Seek(TokenToInt64(*aParam[1]), SEEK_SET) ? 1 : 0;
					return OK;
				}
			}
			else if (!_tcsicmp(field, _T("Tell"))) // Tell
			{
				if (aParamCount == 1)
				{
					aResultToken.value_int64 = mFile.Tell();
					return OK;
				}
			}
			else if (!_tcsicmp(field, _T("Length"))) // Length
			{
				if (aParamCount == 1)
				{
					aResultToken.value_int64 = mFile.Length();
					return OK;
				}
				else if (aParamCount == 2)
				{
					__int64 len = TokenToInt64(*aParam[1]);
					aResultToken.value_int64 = (len >= 0) ? mFile.Length(len) : -1;
					return OK;
				}
			}
			else if (!_tcsicmp(field, _T("AtEOF"))) // AtEOF
			{
				if (aParamCount == 1)
				{
					aResultToken.value_int64 = mFile.AtEOF() ? 1 : 0;
					return OK;
				}
			}
			else if (!_tcsicmp(field, _T("Close"))) // Close
			{
				if (aParamCount == 1)
				{
					mFile.Close();
					return OK;
				}
			}
			else if (!_tcsicmp(field, _T("__Handle"))) // __Handle, prefix with underscores because it is designed for the advanced users.
			{
				if (aParamCount == 1)
				{
					aResultToken.value_int64 = (UINT_PTR) mFile.Handle();
					return OK;
				}
			}
		}

		return r; // Should be INVOKE_NOT_HANDLED if the above didn't return
	}

	TextFile mFile;
	friend void BIF_FileOpen(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);
};

void BIF_FileOpen(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	LPTSTR aFileName = TokenToString(*aParam[0], aResultToken.buf);
	DWORD aFlags = (DWORD) TokenToInt64(*aParam[1]);
	UINT aEncoding;

	if (aParamCount > 2)
	{
		if (!TokenIsPureNumeric(*aParam[2]))
		{
			aEncoding = Line::ConvertFileEncoding(TokenToString(*aParam[2]));
			if (aEncoding == -1)
			{	// Invalid param.
				aResultToken.value_int64 = 0;
				return;
			}
		}
		else aEncoding = (UINT) TokenToInt64(*aParam[2]);
	}
	else aEncoding = g->Encoding;
	
	ASSERT( (~CP_AHKNOBOM) == CP_AHKCP );
	// aEncoding may include CP_AHKNOBOM, in which case below will not add BOM_UTFxx flag.
	if (aEncoding == CP_UTF8)
		aFlags |= TextStream::BOM_UTF8;
	else if (aEncoding == CP_UTF16)
		aFlags |= TextStream::BOM_UTF16;
	
	FileObject *fileObj = new FileObject();
	if (fileObj->mFile.Open(aFileName, aFlags, aEncoding & CP_AHKCP))
	{
		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = fileObj;
	}
	else
	{
		fileObj->Release();
		aResultToken.value_int64 = 0;
	}
}
#endif

//
// TextMem
//
bool TextMem::_Open(LPCTSTR aFileSpec, DWORD aFlags)
{
	ASSERT( (aFlags & 3) == TextStream::READ ); // Only read mode is supported.

	Buffer *buf = (Buffer *) aFileSpec;
	if (mOwned && mBuffer)
		free(mBuffer);
	mPosA = mBufferA = (LPSTR) buf->mBuffer;
	mLength = mCapacity = buf->mLength;
	mOwned = buf->mOwned;
	return true;
}

void TextMem::_Close()
{
	if (mBuffer) {
		if (mOwned)
			free(mBuffer);
		mBuffer = NULL;
	}
}

DWORD TextMem::_Read(LPVOID aBuffer, DWORD aBufSize)
{
	return 0;
}

DWORD TextMem::_Write(LPCVOID aBuffer, DWORD aBufSize)
{
	return 0;
}

bool TextMem::_Seek(__int64 aDistance, int aOrigin)
{
	return false;
}

__int64 TextMem::_Tell() const
{
	return -1; // negative values means it is not supported
}

__int64 TextMem::_Length() const
{
	return mLength;
}
