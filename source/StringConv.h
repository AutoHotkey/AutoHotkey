#pragma once

#ifdef _WIN32

#define IsValidUTF8(str, cch) MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, (str), (cch), NULL, 0)
#define IsValidACP(str, cch) MultiByteToWideChar(CP_ACP, MB_ERR_INVALID_CHARS, (str), (cch), NULL, 0)

LPCWSTR StringUTF8ToWChar(LPCSTR sUTF8, CStringW &sWChar, int iChars = -1);
LPCWSTR StringCharToWChar(LPCSTR sChar, CStringW &sWChar, int iChars = -1, UINT codepage = CP_ACP);
LPCSTR StringWCharToUTF8(LPCWSTR sWChar, CStringA &sUTF8, int iChars = -1);
LPCSTR StringCharToUTF8(LPCSTR sChar, CStringA &sUTF8, int iChars = -1, UINT codepage = CP_ACP);
LPCSTR StringWCharToChar(LPCWSTR sWChar, CStringA &sChar, int iChars = -1, char chDef = '?', UINT codepage = CP_ACP);
LPCSTR StringUTF8ToChar(LPCSTR sUTF8, CStringA &sChar, int iChars = -1, char chDef = '?', UINT codepage = CP_ACP);

LPCWSTR _StringDummyConvW(LPCWSTR sSrc, CStringW &sDest, int iChars = -1);
LPCSTR _StringDummyConvA(LPCSTR sSrc, CStringA &sDest, int iChars = -1);

#ifdef _UNICODE

#define StringTCharToWChar	_StringDummyConvW
#define StringTCharToUTF8	StringWCharToUTF8
#define StringTCharToChar	StringWCharToChar
inline LPCWSTR StringWCharToTChar(LPCWSTR sWChar, CStringW &sTChar, int iChars = -1, char chDef = '?')
{ return _StringDummyConvW(sWChar, sTChar, iChars); }
inline LPCWSTR StringUTF8ToTChar(LPCSTR sUTF8, CStringW &sTChar, int iChars = -1, char chDef = '?')
{ return StringUTF8ToWChar(sUTF8, sTChar, iChars); }
#define StringCharToTChar	StringCharToWChar

#else
// ! _UNICODE
#define StringTCharToWChar	StringCharToWChar
#define StringTCharToUTF8	StringCharToUTF8
#define StringTCharToChar	_StringDummyConvA
#define StringWCharToTChar	StringWCharToChar
#define StringUTF8ToTChar	StringUTF8ToChar
#define StringCharToTChar	_StringDummyConvA

#endif // _UNICODE

class CStringWCharFromUTF8 : public CStringW
{
public:
	CStringWCharFromUTF8(LPCSTR sUTF8, int iChars = -1)
	{ StringUTF8ToWChar(sUTF8, *this, iChars); }
};

class CStringWCharFromChar : public CStringW
{
public:
	CStringWCharFromChar(LPCSTR sChar, int iChars = -1, UINT codepage = CP_ACP)
	{ StringCharToWChar(sChar, *this, iChars, codepage); }
};

class CStringUTF8FromWChar : public CStringA
{
public:
	CStringUTF8FromWChar(LPCWSTR sWChar, int iChars = -1)
	{ StringWCharToUTF8(sWChar, *this, iChars); }
};

class CStringUTF8FromChar : public CStringA
{
public:
	CStringUTF8FromChar(LPCSTR sChar, int iChars = -1, UINT codepage = CP_ACP)
	{ StringCharToUTF8(sChar, *this, iChars, codepage); }
};

class CStringCharFromWChar : public CStringA
{
public:
	CStringCharFromWChar(LPCWSTR sWChar, int iChars = -1, char chDef = '?', UINT codepage = CP_ACP)
	{ StringWCharToChar(sWChar, *this, iChars, chDef, codepage); }
};

class CStringCharFromUTF8 : public CStringA
{
public:
	CStringCharFromUTF8(LPCSTR sUTF8, int iChars = -1, char chDef = '?', UINT codepage = CP_ACP)
	{ StringUTF8ToChar(sUTF8, *this, iChars, chDef, codepage); }
};

#ifdef _UNICODE
#define CStringTCharFromWCharIfNeeded(s, ...) (s)
#define CStringWCharFromTCharIfNeeded(s, ...) (s)
#define CStringTCharFromCharIfNeeded(s, ...) CStringTCharFromChar((s), __VA_ARGS__)
#define CStringCharFromTCharIfNeeded(s, ...) CStringCharFromTChar((s), __VA_ARGS__)

class CStringTCharFromWChar : public CStringW
{
public:
	CStringTCharFromWChar(LPCWSTR sWChar, int iChars = -1, char chDef = '?')
	{ _StringDummyConvW(sWChar, *this, iChars); }
};
class CStringTCharFromUTF8 : public CStringW
{
public:
	CStringTCharFromUTF8(LPCSTR sUTF8, int iChars = -1, char chDef = '?')
	{ StringUTF8ToWChar(sUTF8, *this, iChars); }
};
typedef	CStringWCharFromChar	CStringTCharFromChar;
typedef CStringW				CStringWCharFromTChar;
typedef CStringUTF8FromWChar	CStringUTF8FromTChar;
typedef CStringCharFromWChar	CStringCharFromTChar;

#else
// ! _UNICODE
#define CStringTCharFromWCharIfNeeded(s, ...) CStringTCharFromWChar((s), __VA_ARGS__)
#define CStringWCharFromTCharIfNeeded(s, ...) CStringWCharFromTChar((s), __VA_ARGS__)
#define CStringTCharFromCharIfNeeded(s, ...) (s)
#define CStringCharFromTCharIfNeeded(s, ...) (s)

typedef CStringCharFromWChar	CStringTCharFromWChar;
typedef CStringCharFromUTF8		CStringTCharFromUTF8;
typedef	CStringA				CStringTCharFromChar;
typedef CStringWCharFromChar	CStringWCharFromTChar;
typedef CStringUTF8FromChar		CStringUTF8FromTChar;
typedef CStringA				CStringCharFromTChar;
#endif // _UNICODE

typedef CStringTCharFromWChar	CStringFromWChar;
typedef CStringTCharFromUTF8	CStringFromUTF8;
typedef	CStringTCharFromChar	CStringFromChar;

#endif // _WIN32

template <typename CHAR_T>
void CharConvEndian(CHAR_T *pChar)
{
	BYTE *pCh = (BYTE *) pChar, chTemp;
	int i, j;
	for (i = 0, j = sizeof(CHAR_T) - 1;i < (sizeof(CHAR_T) >> 1);i++, j--) {
		chTemp = pCh[i];
		pCh[i] = pCh[j];
		pCh[j] = chTemp;
	}
}

// uChars = 0 ==> treat 'sSrc' as NULL-terminal
template <typename CHAR_T>
unsigned long StringConvEndian(CHAR_T *sSrc, unsigned long uChars = 0)
{
	unsigned long i;
	for (i = 0;sSrc[i] && (uChars == 0 || i < uChars);i++)
		CharConvEndian(sSrc + i);
	return i;
}
