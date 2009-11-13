#include "stdafx.h"
#include "TextIO.h"

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
	memset(mCache, 0, sizeof(mCache));

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
	if (mode == TextStream::WRITE || mode == TextStream::APPEND && _Length() == 0) {
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
			dwDesiredAccess = GENERIC_WRITE | GENERIC_READ;
			dwShareMode = 0;
			dwCreationDisposition = OPEN_ALWAYS;
			break;
	}
	// FILE_FLAG_SEQUENTIAL_SCAN is set, sequential accesses are quite common for text files handling.
	mFile = CreateFile(aFileSpec, dwDesiredAccess, dwShareMode, NULL, dwCreationDisposition, FILE_FLAG_SEQUENTIAL_SCAN, NULL);
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

bool TextFile::_Seek(long aDistance, int aOrigin)
{
	return SetFilePointer(mFile, aDistance, NULL, aOrigin) != INVALID_SET_FILE_POINTER;
}

__int64 TextFile::_Length()
{
	LARGE_INTEGER size;
	GetFileSizeEx(mFile, &size);
	return size.QuadPart;
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

bool TextMem::_Seek(long aDistance, int aOrigin)
{
	return false;
}

__int64 TextMem::_Length()
{
	return mLength;
}
