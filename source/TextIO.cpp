#include "stdafx.h"
#include "TextIO.h"
#include "script.h"
#include "script_object.h"

//
// TextStream
//
bool TextStream::Open(LPCTSTR aFileSpec, DWORD aFlags, UINT aCodePage)
{
	mLength = 0; // Set the default value here so _Open() can change it.
	if (!_Open(aFileSpec, aFlags))
		return false;
	SetCodePage(aCodePage == CP_ACP ? GetACP() : aCodePage);
	mFlags = aFlags;
	mEOF = false;
	mCacheInt = 0;

	int mode = aFlags & 3;
	if (mode != TextStream::WRITE) {
		// Detects the BOM of UTF-8 and UTF-16LE
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



WCHAR TextStream::ReadCharW()
{
	if (mCacheW[1]) { // surrogate pair or standalone CR
		WCHAR ret = mCacheW[1];
		mCacheW[1] = 0;
		return ret;
	}
	if (mEOF)
		return WEOF;

	if (mCodePage == CP_UTF16) {
		if (!ReadAtLeast(sizeof(WCHAR)))
			return WEOF;
		return *mPosW++;
	}
	else {
		if (!ReadAtLeast(sizeof(CHAR)))
			return WEOF;

		int iBytes;
		if (mCodePage == CP_UTF8)
		{
			if (*mPos < 0x80)
				// single byte UTF-8 character
				return (wchar_t) *mPosA++;
			// The size in bytes of UTF-8 characters.
			if ((*mPos & 0xE0) == 0xC0)
				iBytes = 2;
			else if ((*mPos & 0xF0) == 0xE0)
				iBytes = 3;
			else if ((*mPos & 0xF8) == 0xF0)
				iBytes = 4;
			else {
				// Invalid in current UTF-8 stardard.
				mPosA++;
				return '?';
			}
		}
		else if (mCodePageIsDBCS)
			iBytes = IsDBCSLeadByteEx(mCodePage, *mPos) ? 2 : 1;
		else
			iBytes = 1;

		if (!ReadAtLeast(iBytes))
			return WEOF;
		if (MultiByteToWideChar(mCodePage, MB_ERR_INVALID_CHARS, mPosA, iBytes, mCacheW, 2) > 0) {
			mPosA += iBytes;
			return mCacheW[0];
		}
		else
			mPosA++; // ignore invalid byte
	}
	return '?'; // invalid character
}



DWORD TextStream::Write(LPCWSTR aBuf, DWORD aBufLen)
{
	if (aBufLen == 0)
		aBufLen = wcslen(aBuf);
	if (mCodePage == CP_UTF16) {
		if (mFlags & EOL_CRLF) {
			DWORD dwWritten = 0;
			int i;
			for (i = 0;i < aBufLen;i++) {
				if (aBuf[i] == '\n' && mCacheW[0] != '\r')
					dwWritten += _Write(L"\r\n", 4);
				else
					dwWritten += _Write(&aBuf[i], 1);
				mCacheW[0] = aBuf[i];
			}
			return dwWritten;
		}
		return _Write(aBuf, aBufLen * sizeof(wchar_t));
	}
	else {
		CStringA sBuf;
		if (mCodePage == CP_UTF8)
			StringWCharToUTF8(aBuf, sBuf, aBufLen);
		else
			StringWCharToChar(aBuf, sBuf, aBufLen, '?', mCodePage);
		if (mFlags & EOL_CRLF) {
			DWORD dwWritten = 0;
			int i;
			for (i = 0;i < sBuf.GetLength();i++) {
				if (sBuf[i] == '\n' && mCacheA[0] != '\r')
					dwWritten += _Write("\r\n", 2);
				else
					dwWritten += _Write(&sBuf[i], 1);
				mCacheA[0] = sBuf[i];
			}
			return dwWritten;
		}
		return _Write(sBuf, sBuf.GetLength());
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
			dwShareMode = FILE_SHARE_READ;
			dwCreationDisposition = OPEN_EXISTING;
			break;
		case TextStream::WRITE:
			dwDesiredAccess = GENERIC_WRITE;
			dwShareMode = 0;
			dwCreationDisposition = CREATE_ALWAYS;
			break;
		case TextStream::APPEND:
		case TextStream::UPDATE:
			dwDesiredAccess = GENERIC_WRITE | GENERIC_READ;
			dwShareMode = 0;
			dwCreationDisposition = OPEN_ALWAYS;
			break;
	}
	// FILE_FLAG_SEQUENTIAL_SCAN is set, as sequential accesses are quite common for text files handling.
	mFile = CreateFile(aFileSpec, dwDesiredAccess, dwShareMode, NULL, dwCreationDisposition,
		(aFlags & (EOL_CRLF | EOL_ORPHAN_CR)) ? FILE_FLAG_SEQUENTIAL_SCAN : 0, NULL);
	if (mFile == INVALID_HANDLE_VALUE)
		return false;
	return true;
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

class FileObject : public Object
{
public:
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
								return r;
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
							return r;
						aResultToken.symbol = SYM_STRING;
						aResultToken.marker = (LPTSTR) aResultToken.circuit_token; // Store the address of the result for the caller.
						aResultToken.buf = (LPTSTR)(size_t) mFile.ReadLine(aResultToken.marker, READ_FILE_LINE_SIZE - 1); // MANDATORY FOR USERS OF CIRCUIT_TOKEN: "buf" is being overloaded to store the length for our caller.
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
					else if (!_tcsicmp(field + 3, _T("Write"))) // ReadWrite
						iReadWrite = 2;
					if (iReadWrite)
					// Reference: BIF_NumGet
					{
						size_t right_side_bound, target; // Don't make target a pointer-type because the integer offset might not be a multiple of 4 (i.e. the below increments "target" directly by "offset" and we don't want that to use pointer math).
						ExprTokenType &target_token = *aParam[1];
						if (target_token.symbol == SYM_VAR) // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
						{
							target = (size_t)target_token.var->Contents(); // Although Contents(TRUE) will force an update of mContents if necessary, it very unlikely to be necessary here because we're about to fetch a binary number from inside mContents, not a normal/text number.
							right_side_bound = target + target_token.var->ByteCapacity(); // This is first illegal address to the right of target.
						}
						else
							target = (size_t)TokenToInt64(target_token);

						size_t size = TokenToInt64(*aParam[2]);

						if (target < 1024 // Basic sanity check to catch incoming raw addresses that are zero or blank.
							|| target_token.symbol == SYM_VAR && target+size > right_side_bound) // i.e. it's ok if target+size==right_side_bound because the last byte to be read is actually at target+size-1. In other words, the position of the last possible terminator within the variable's capacity is considered an allowable address.
						{
							aResultToken.value_int64 = 0;
							return OK;
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
					aResultToken.value_int64 = mFile.Seek(TokenToInt64(*aParam[1]), TokenToInt64(*aParam[2])) ? 1 : 0;
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
		}

		return r;
	}

	TextFile mFile;
};

void BIF_FileOpen(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	LPTSTR aFileName = TokenToString(*aParam[0], aResultToken.buf);
	DWORD aFlags = (DWORD) TokenToInt64(*aParam[1]);
	UINT aCodePage = aParamCount > 2 ? (UINT) TokenToInt64(*aParam[2]) : CP_ACP;

	FileObject *fileObj = new FileObject();
	if (fileObj->mFile.Open(aFileName, aFlags, aCodePage))
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
