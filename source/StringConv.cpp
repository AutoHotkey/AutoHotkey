#include "stdafx.h"
#include "KuString.h"
#include "StringConv.h"

#ifdef _WIN32
LPCWSTR StringUTF8ToWChar(LPCSTR sUTF8, CStringW &sWChar, int iChars/* = -1*/)
{
	if (!sUTF8)
		return NULL;

	sWChar.Empty();
	int iLen = MultiByteToWideChar(CP_UTF8, 0, sUTF8, iChars, NULL, 0);
	if (iLen > 0) {
		LPWSTR sBuf = sWChar.GetBufferSetLength(iLen);
		iLen = MultiByteToWideChar(CP_UTF8, 0, sUTF8, iChars, sBuf, iLen);
		sWChar.ReleaseBufferSetLength(sBuf[iLen - 1] ? iLen : iLen - 1);
		return (iLen > 0) ? sWChar.GetString() : NULL;
	}

	return *sUTF8 != 0 ? sWChar.GetString() : NULL;
}

LPCWSTR StringCharToWChar(LPCSTR sChar, CStringW &sWChar, int iChars/* = -1*/, UINT codepage/* = CP_ACP*/)
{
	if (!sChar)
		return NULL;

	sWChar.Empty();
	int iLen = MultiByteToWideChar(codepage, 0, sChar, iChars, NULL, 0);
	if (iLen > 0) {
		LPWSTR sBuf = sWChar.GetBufferSetLength(iLen);
		MultiByteToWideChar(codepage, 0, sChar, iChars, sBuf, iLen);
		sWChar.ReleaseBufferSetLength(sBuf[iLen - 1] ? iLen : iLen - 1);
		return (iLen > 0) ? sWChar.GetString() : NULL;
	}

	return (*sChar != 0) ? sWChar.GetString() : NULL;
}

LPCSTR StringWCharToUTF8(LPCWSTR sWChar, CStringA &sUTF8, int iChars/* = -1*/)
{
	if (!sWChar)
		return NULL;

	sUTF8.Empty();
	int iLen = WideCharToMultiByte(CP_UTF8, 0, sWChar, iChars, NULL, 0, NULL, NULL);
	if (iLen > 0) {
		LPSTR sBuf = sUTF8.GetBufferSetLength(iLen);
		WideCharToMultiByte(CP_UTF8, 0, sWChar, iChars, sBuf, iLen, NULL, NULL);
		sUTF8.ReleaseBufferSetLength(sBuf[iLen - 1] ? iLen : iLen - 1);
		return (iLen > 0) ? sUTF8.GetString() : NULL;
	}

	return (*sWChar != 0) ? sUTF8.GetString() : NULL;
}

LPCSTR StringCharToUTF8(LPCSTR sChar, CStringA &sUTF8, int iChars/* = -1*/, UINT codepage/* = CP_ACP*/)
{
	return StringWCharToUTF8(CStringWCharFromChar(sChar, iChars, codepage), sUTF8);
}

LPCSTR StringWCharToChar(LPCWSTR sWChar, CStringA &sChar, int iChars/* = -1*/, char chDef/* = '?'*/, UINT codepage/* = CP_ACP*/)
{
	if (!sWChar)
		return NULL;

	sChar.Empty();
	int iLen = WideCharToMultiByte(codepage, WC_NO_BEST_FIT_CHARS, sWChar, iChars, NULL, 0, &chDef, NULL);
	if (iLen > 0) {
		LPSTR sBuf = sChar.GetBufferSetLength(iLen);
		WideCharToMultiByte(codepage, WC_NO_BEST_FIT_CHARS, sWChar, iChars, sBuf, iLen, &chDef, NULL);
		sChar.ReleaseBufferSetLength(sBuf[iLen - 1] ? iLen : iLen - 1);
		return (iLen > 0) ? sChar.GetString() : NULL;
	}

	return (*sWChar != 0) ? sChar.GetString() : NULL;
}

LPCSTR StringUTF8ToChar(LPCSTR sUTF8, CStringA &sChar, int iChars/* = -1*/, char chDef/* = '?'*/, UINT codepage/* = CP_ACP*/)
{
	return StringWCharToChar(CStringWCharFromUTF8(sUTF8, iChars), sChar, iChars, chDef, codepage);
}
#endif

template <typename SRC_T, typename DEST_T>
SRC_T _StringDummyConv(SRC_T sSrc, DEST_T sDest, int iChars = -1)
{
	if (!sSrc)
		return NULL;
	if (iChars >= 0)
		sDest.SetString(sSrc, iChars);
	else
		sDest = sSrc;
	return sDest;
}

LPCWSTR _StringDummyConvW(LPCWSTR sSrc, CStringW &sDest, int iChars/* = -1*/)
{
	return _StringDummyConv(sSrc, sDest, iChars);
}

LPCSTR _StringDummyConvA(LPCSTR sSrc, CStringA &sDest, int iChars/* = -1*/)
{
	return _StringDummyConv(sSrc, sDest, iChars);
}
