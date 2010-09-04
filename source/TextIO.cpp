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

	SetCodePage(aCodePage);
	mFlags = aFlags;
	mEOF = false;
	mCacheInt = 0;
	mWriteCharW = 0;

	int mode = aFlags & ACCESS_MODE_MASK;
	if (mode == USEHANDLE)
		return true;
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
		else if (mLocale && _ismbblead_l(*mPos, mLocale))
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
		aBufLen = (DWORD)_tcslen(aBuf);
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
		DWORD len;
		CStringA buf_a;
#ifndef UNICODE
		// Since it's probably the most common case, optimize for the active codepage:
		if (g_ACP == mCodePage)
		{
			str = aBuf;
			len = aBufLen;
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
			len = (DWORD)buf_a.GetLength();
		}
		
		if (mFlags & EOL_CRLF)
		{
			DWORD dwWritten = 0;
			DWORD i;
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
	switch (aFlags & ACCESS_MODE_MASK) {
		case READ:
			dwDesiredAccess = GENERIC_READ;
			dwCreationDisposition = OPEN_EXISTING;
			break;
		case WRITE:
			dwDesiredAccess = GENERIC_WRITE;
			dwCreationDisposition = CREATE_ALWAYS;
			break;
		case APPEND:
		case UPDATE:
			dwDesiredAccess = GENERIC_WRITE | GENERIC_READ;
			dwCreationDisposition = OPEN_ALWAYS;
			break;
		case USEHANDLE:
			if (!GetFileType((HANDLE)aFileSpec))
				return false;
			mFile = (HANDLE)aFileSpec;
			return true;
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
		if ((mFlags & ACCESS_MODE_MASK) != USEHANDLE)
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
class FileObject : public ObjectBase // fincs: No longer allowing the script to manipulate File objects
{
	FileObject() {}
	~FileObject() {}

	enum MemberID {
		INVALID = 0,
		// methods
		Read,
		Write,
		ReadLine,
		WriteLine,
		NumReadWrite,
		RawReadWrite,
		LastMethodPlusOne,
		// properties
		Position,
		Length,
		AtEOF,
		Handle,
		Close
	};

	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
	// Reference: MetaObject::Invoke
	{
		if (!aParamCount) // file[]
			return INVOKE_NOT_HANDLED;
		
		--aParamCount; // Exclude name from param count.
		LPTSTR name = TokenToString(*aParam[0]); // Name of method or property.
		MemberID member = INVALID;

		// Read' and Write' must be handled differently to support ReadUInt(), WriteShort(), etc.
		if (!_tcsnicmp(name, _T("Read"), 4))
		{
			if (!name[4])
				member = Read;
			else if (!_tcsicmp(name + 4, _T("Line")))
				member = ReadLine;
			else
				member = NumReadWrite;
		}
		else if (!_tcsnicmp(name, _T("Write"), 5))
		{
			if (!name[5])
				member = Write;
			else if (!_tcsicmp(name + 5, _T("Line")))
				member = WriteLine;
			else
				member = NumReadWrite;
		}
	#define if_member(s,e)	else if (!_tcsicmp(name, _T(s))) member = e;
		if_member("RawRead", RawReadWrite)
		if_member("RawWrite", RawReadWrite)
		if_member("Pos", Position)
		if_member("Length", Length)
		if_member("AtEOF", AtEOF)
		if_member("__Handle", Handle) // Prefix with underscores because it is designed for advanced users.
		if_member("Close", Close)
		// Supported for enhanced clarity:
		if_member("Position", Position)
		// Legacy names:
		if_member("Seek", Position)
		if_member("Tell", Position)
	#undef if_member
		if (member == INVALID)
			return INVOKE_NOT_HANDLED;

		// Syntax validation:
		if (!IS_INVOKE_CALL)
		{
			if (member < LastMethodPlusOne)
				// Member requires parentheses().
				return OK;
			if (aParamCount != (IS_INVOKE_SET ? 1 : 0))
				// Get: disallow File.Length[newLength] and File.Seek[dist,origin].
				// Set: disallow File[]:=PropertyName and File["Pos",dist]:=origin.
				return OK;
		}

		aResultToken.symbol = SYM_INTEGER; // Set default return type -- the most common cases return integer.

		switch (member)
		{
		case NumReadWrite:
			{
				bool reading = (*name == 'R' || *name == 'r');
				LPCTSTR type = name + (reading ? 4 : 5);

				// Based on BIF_NumGet:

				BOOL is_signed, is_float = FALSE;
				DWORD size = 0;

				if (ctoupper(*type) == 'U') // Unsigned.
				{
					++type; // Remove the first character from further consideration.
					is_signed = FALSE;
				}
				else
					is_signed = TRUE;

				switch(ctoupper(*type)) // Override "size" and aResultToken.symbol if type warrants it. Note that the above has omitted the leading "U", if present, leaving type as "Int" vs. "Uint", etc.
				{
				case 'I':
					if (_tcschr(type, '6')) // Int64. It's checked this way for performance, and to avoid access violation if string is bogus and too short such as "i64".
						size = 8;
					else
						size = 4;
					break;
				case 'S': size = 2; break; // Short.
				case 'C': size = 1; break; // Char.
				case 'P': size = sizeof(INT_PTR); break;

				case 'D': size = 8; is_float = true; break; // Double.
				case 'F': size = 4; is_float = true; break; // Float.
				}
				if (!size)
					break; // and return ""

				union {
						__int64 i8;
						int i4;
						short i2;
						char i1;
						double d;
						float f;
					} buf;

				if (reading)
				{
					buf.i8 = 0;
					if ( !mFile.Read(&buf, size) )
						break; // and return ""

					if (is_float)
					{
						aResultToken.value_double = (size == 4) ? buf.f : buf.d;
						aResultToken.symbol = SYM_FLOAT;
					}
					else
					{
						if (is_signed)
						{
							// sign-extend to 64-bit
							switch (size)
							{
							case 4: buf.i8 = buf.i4; break;
							case 2: buf.i8 = buf.i2; break;
							case 1: buf.i8 = buf.i1; break;
							//case 8: not needed.
							}
						}
						//else it's unsigned. No need to zero-extend thanks to init done earlier.
						aResultToken.value_int64 = buf.i8;
						//aResultToken.symbol = SYM_INTEGER; // This is the default.
					}
				}
				else
				{
					if (aParamCount != 1)
						break; // and return ""

					ExprTokenType &token_to_write = *aParam[1];
					
					if (is_float)
					{
						buf.d = TokenToDouble(token_to_write);
						if (size == 4)
							buf.f = (float)buf.d;
					}
					else
					{
						if (size == 8 && !is_signed && !IS_NUMERIC(token_to_write.symbol))
							buf.i8 = (__int64)ATOU64(TokenToString(token_to_write)); // For comments, search for ATOU64 in BIF_DllCall().
						else
							buf.i8 = TokenToInt64(token_to_write);
					}
					
					aResultToken.value_int64 = mFile.Write(&buf, size);
				}
				return OK;
			}
			break;

		case Read:
			if (aParamCount <= 1)
			{
				aResultToken.symbol = SYM_STRING; // Set for both paths below.
				DWORD length;
				if (aParamCount)
					length = (DWORD)TokenToInt64(*aParam[1]);
				else
					length = (DWORD)(mFile.Length() - mFile.Tell());
				if (length == -1 || !TokenSetResult(aResultToken, NULL, length)) // Relies on short-circuit order. TokenSetResult requires non-NULL aResult if aResultLength == -1.
				{
					// Our caller set marker to a default result of "", which should still be in place.
					return OK; // FAIL vs OK currently has no real effect here, but in Line::ExecUntil it is used to exit the current thread when a critical error occurs.  Since that behaviour might be implemented for objects someday and in this particular case a bad parameter is more likely than critically low memory, FAIL seems inappropriate.
				}
				length = mFile.Read(aResultToken.marker, length);
				aResultToken.marker[length] = '\0';
				aResultToken.buf = (LPTSTR)(size_t) length; // Update buf to the actual number of characters read. Only strictly necessary in some cases; see TokenSetResult.
				return OK;
			}
			break;
		
		case ReadLine:
			if (aParamCount == 0)
			{	// See above for comments.
				aResultToken.symbol = SYM_STRING;
				if (!TokenSetResult(aResultToken, NULL, READ_FILE_LINE_SIZE))
					return OK; 
				DWORD length = mFile.ReadLine(aResultToken.marker, READ_FILE_LINE_SIZE - 1);
				aResultToken.marker[length] = '\0';
				aResultToken.buf = (LPTSTR)(size_t) length;
				return OK;
			}
			break;

		case Write:
		case WriteLine:
			if (aParamCount <= 1)
			{
				DWORD written = 0;
				if (aParamCount)
				{
					written = mFile.Write(TokenToString(*aParam[1], aResultToken.buf));
				}
				if (member == WriteLine)
				{
					written += mFile.Write(_T("\n"), 1);
				}
				aResultToken.value_int64 = written;
				return OK;
			}
			break;

		case RawReadWrite:
			if (aParamCount == 2)
			{
				bool reading = (name[3] == 'R' || name[3] == 'r');

				LPVOID target;
				ExprTokenType &target_token = *aParam[1];
				DWORD size = (DWORD)TokenToInt64(*aParam[2]);

				if (target_token.symbol == SYM_VAR) // SYM_VAR's Type() is always VAR_NORMAL (except lvalues in expressions).
				{
					// Check if the user requested a size larger than the variable.
					if ( size > target_token.var->ByteCapacity()
						// Too small: expand the target variable if reading; abort otherwise.
						&& (!reading || !target_token.var->SetCapacity(size, false, false)) ) // Relies on short-circuit order.
					{
						aResultToken.value_int64 = 0;
						return OK;
					}
					target = target_token.var->Contents();
				}
				else
					target = (LPVOID)TokenToInt64(target_token);

				DWORD result;
				if (target < (LPVOID)1024) // Basic sanity check to catch incoming raw addresses that are zero or blank.
					result = 0;
				else if (reading)
					result = mFile.Read(target, size);
				else
					result = mFile.Write(target, size);
				aResultToken.value_int64 = result;
				return OK;
			}
			break;

		case Position:
			if (aParamCount == 0)
			{
				aResultToken.value_int64 = mFile.Tell();
				return OK;
			}
			else if (aParamCount <= 2)
			{
				__int64 distance = TokenToInt64(*aParam[1]);
				int origin;
				if (aParamCount == 2)
					origin = (int)TokenToInt64(*aParam[2]);
				else // Defaulting to SEEK_END when distance is negative seems more useful than allowing it to be interpreted as an unsigned value (> 9.e18 bytes).
					origin = (distance < 0) ? SEEK_END : SEEK_SET;

				aResultToken.value_int64 = mFile.Seek(distance, origin);
				return OK;
			}
			break;

		case Length:
			if (aParamCount == 0)
			{
				aResultToken.value_int64 = mFile.Length();
				return OK;
			}
			else if (aParamCount == 1)
			{
				if (-1 != (aResultToken.value_int64 = mFile.Length(TokenToInt64(*aParam[1]))))
					return OK;
				else // Empty string seems like a more suitable failure indicator than -1.
					aResultToken.marker = _T("");
					// Let below set symbol back to SYM_STRING.
			}
			break;

		case AtEOF:
			if (aParamCount == 0)
				aResultToken.value_int64 = mFile.AtEOF();
			return OK;
		
		case Handle:
			if (aParamCount == 0)
				aResultToken.value_int64 = (UINT_PTR) mFile.Handle();
			return OK;

		case Close:
			if (aParamCount == 0)
				mFile.Close();
			return OK;
		}
		
		// Since above didn't return, an error must've occurred.
		aResultToken.symbol = SYM_STRING;
		// marker should already be set to "".
		return OK;
	}

	TextFile mFile;
	
public:
	static inline FileObject *Open(LPCTSTR aFileSpec, DWORD aFlags, UINT aCodePage)
	{
		FileObject *fileObj = new FileObject();
		if (fileObj && fileObj->mFile.Open(aFileSpec, aFlags, aCodePage))
			return fileObj;
		fileObj->Release();
		return NULL;
	}
};

void BIF_FileOpen(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	DWORD aFlags;
	UINT aEncoding;

	if (TokenIsPureNumeric(*aParam[1]))
	{
		aFlags = (DWORD) TokenToInt64(*aParam[1]);
	}
	else
	{
		LPCTSTR sflag = TokenToString(*aParam[1], aResultToken.buf);

		sflag = omit_leading_whitespace(sflag); // For consistency with the loop below.

		// Access mode must come first:
		switch (tolower(*sflag))
		{
		case 'r':
			if (tolower(sflag[1]) == 'w')
			{
				aFlags = TextStream::UPDATE;
				++sflag;
			}
			else
				aFlags = TextStream::READ;
			break;
		case 'w': aFlags = TextStream::WRITE; break;
		case 'a': aFlags = TextStream::APPEND; break;
		case 'h': aFlags = TextStream::USEHANDLE; break;
		default:
			// Invalid flag.
			aResultToken.value_int64 = 0;
			g->LastError = ERROR_INVALID_PARAMETER; // For consistency.
			return;
		}
		
		// Default to not locking file, for consistency with fopen/standard AutoHotkey and because it seems best for flexibility.
		aFlags |= TextStream::SHARE_ALL;

		for (++sflag; *sflag; ++sflag)
		{
			switch (ctolower(*sflag))
			{
			case '\n': aFlags |= TextStream::EOL_CRLF; break;
			case '\r': aFlags |= TextStream::EOL_ORPHAN_CR; break;
			case ' ':
			case '\t':
				// Allow spaces and tabs for readability.
				break;
			case '-':
				for (++sflag; ; ++sflag)
				{
					switch (ctolower(*sflag))
					{
					case 'r': aFlags &= ~TextStream::SHARE_READ; continue;
					case 'w': aFlags &= ~TextStream::SHARE_WRITE; continue;
					case 'd': aFlags &= ~TextStream::SHARE_DELETE; continue;
					// Whitespace not allowed here.  Outer loop allows "-r -w" but not "-r w".
					}
					if (sflag[-1] == '-')
						// Let "-" on its own be equivalent to "-rwd".
						aFlags &= ~TextStream::SHARE_ALL;
					break;
				}
				--sflag; // Point sflag at the last char of this option.  Outer loop will do ++sflag.
				break;
			default:
				// Invalid flag.
				aResultToken.value_int64 = 0;
				g->LastError = ERROR_INVALID_PARAMETER; // For consistency.
				return;
			}
		}
	}

	if (aParamCount > 2)
	{
		if (!TokenIsPureNumeric(*aParam[2]))
		{
			aEncoding = Line::ConvertFileEncoding(TokenToString(*aParam[2]));
			if (aEncoding == -1)
			{	// Invalid param.
				aResultToken.value_int64 = 0;
				g->LastError = ERROR_INVALID_PARAMETER; // For consistency.
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

	LPTSTR aFileName;
	if ((aFlags & TextStream::ACCESS_MODE_MASK) == TextStream::USEHANDLE)
		aFileName = (LPTSTR)(HANDLE)TokenToInt64(*aParam[0]);
	else
		aFileName = TokenToString(*aParam[0], aResultToken.buf);

	if (aResultToken.object = FileObject::Open(aFileName, aFlags, aEncoding & CP_AHKCP))
		aResultToken.symbol = SYM_OBJECT;
	else
		aResultToken.value_int64 = 0; // and symbol is already SYM_INTEGER.

	g->LastError = GetLastError();
}
#endif

//
// TextMem
//
bool TextMem::_Open(LPCTSTR aFileSpec, DWORD aFlags)
{
	ASSERT( (aFlags & ACCESS_MODE_MASK) == TextStream::READ ); // Only read mode is supported.

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
