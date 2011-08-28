/* This file is part of KuShellExtension
 * Copyright (C) 2008-2009 Kai-Chieh Ku (kjackie@gmail.com)
 *
 * This program is free software; you can redistribute it and/or
 * modify it under the terms of the GNU General Public License
 * as published by the Free Software Foundation; either
 * version 3 of the License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.	See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program; if not, write to the Free Software
 * Foundation, Inc., 675 Mass Ave, Cambridge, MA 02139, USA.
 */

#pragma once

/*
	CKuStringT: A copy on write CString-like string class.

	The "copy on write" means,
		CKuString s1, s2, s3, s4;
		s1 = s2 = s3 = s4 = "this is a string";
	will share the same buffer, which saves spaces and does faster when copying strings.

	WARNING: It is not thread safe, to copy a string from a thread to another thread, use following syntax
	s2 = s1.GetString(); // s2 will have its own buffer, which is not shared by other instances.
*/

#ifdef _WIN32
#define WIN32_LEAN_AND_MEAN
#define _CRT_SECURE_NO_WARNINGS
#define _CRT_NON_CONFORMING_SWPRINTFS
#include <windows.h>
#define strcasecmp stricmp
#define wcscasecmp wcsicmp
#else
#define _vsnwprintf vswprintf
#endif

#include <stdio.h>
#include <stdlib.h>
#include <stdarg.h>
#include <string.h>
#include <ctype.h>
#ifndef NO_WCHAR // Most systems has wchar_t now, including *nix and Windows.
#include <wchar.h>
#include <wctype.h>
#endif

#ifndef ASSERT
#ifdef _DEBUG
#include <assert.h>
#define ASSERT(expr) assert(expr)
#else
#define ASSERT(expr)
#endif
#endif

#ifndef VERIFY
#ifdef _DEBUG
#define VERIFY(expr) ASSERT(expr)
#else
#define VERIFY(expr) expr
#endif
#endif

#ifndef va_copy
	#ifdef _WIN32
		#define va_copy(dst, src) { (dst) = (src); }
	#endif
#endif

template<typename T, typename U> class CKuStringT;
class CKuStringUtilA;
class CKuStringUtilW;

typedef CKuStringT<char, CKuStringUtilA> CKuStringA;
#ifndef NO_WCHAR
typedef CKuStringT<wchar_t, CKuStringUtilW> CKuStringW;
#endif

#if defined(_WIN32) && defined(_UNICODE) && !defined(NO_WCHAR)
typedef CKuStringW CKuString;
typedef CKuStringUtilW CKuStringUtil;
#else
typedef CKuStringA CKuString;
typedef CKuStringUtilA CKuStringUtil;
#endif

#ifndef __CSTRINGT_H__
typedef CKuStringA CStringA;
#ifndef NO_WCHAR
typedef CKuStringW CStringW;
#endif
typedef CKuString CString;
#endif

#ifdef _WIN32
extern LPCWSTR StringCharToWChar(LPCSTR sChar, CStringW &sWChar, int iChars/* = -1*/, UINT codepage/* = CP_ACP*/);
extern LPCSTR StringWCharToChar(LPCWSTR sWChar, CStringA &sChar, int iChars/* = -1*/, char chDef/* = '?'*/, UINT codepage/* = CP_ACP*/);
#endif

#ifdef _MSC_VER
#pragma warning(push)
#pragma warning(disable: 4996)
#endif

class CKuStringUtilA
{
public:
	inline static int FormatV(char *dest, size_t cch, const char *fmt, va_list ap)
	{ return vsnprintf(dest, cch, fmt, ap); }
	inline static int Compare(const char *a, const char *b)
	{ return strcmp(a, b); }
	inline static int CompareNoCase(const char *a, const char *b)
	{ return strcasecmp(a, b); }
	inline static size_t SpanIncluding(const char *str, const char *strCharSet)
	{ return strspn(str, strCharSet); }
	inline static size_t SpanExcluding(const char *str, const char *strCharSet)
	{ return strcspn(str, strCharSet); }
	inline static char * FindString(char *haystack, const char *needle)
	{ return strstr(haystack, needle); }
	inline static bool IsSpace(char c)
	{ return !!isspace((unsigned char) c); }
	inline static bool IsLower(char c)
	{ return !!islower((unsigned char) c); }
	inline static bool IsUpper(char c)
	{ return !!isupper((unsigned char) c); }
	inline static char ToLower(char c)
	{ return IsUpper(c) ? (char) tolower((unsigned char) c) : c; }
	inline static char ToUpper(char c)
	{ return IsLower(c) ? (char) toupper((unsigned char) c) : c; }
	inline static DWORD GetEnvironmentVariable(const char *lpName, char *lpBuffer, DWORD nSize)
#ifdef _WIN32
	{ return ::GetEnvironmentVariableA(lpName, lpBuffer, nSize); }
#else
	{
		if (lpBuffer) {
			int ret = 0;
			const char *env = getenv(lpName);
			while (ret < nSize && (*lpBuffer++ = *env++))
				ret++;
			return ret;
		}
		return strlen(getenv(lpName));
	}
#endif
#ifdef _WIN32
	inline static bool LoadString(CKuStringA &str, HINSTANCE hInstance, UINT nID, WORD wLanguage);
#endif
};

#ifndef NO_WCHAR
class CKuStringUtilW
{
public:
	inline static int FormatV(wchar_t *dest, size_t cch, const wchar_t *fmt, va_list ap)
	{ return _vsnwprintf(dest, cch, fmt, ap); }
	inline static int Compare(const wchar_t *a, const wchar_t *b)
	{ return wcscmp(a, b); }
	inline static int CompareNoCase(const wchar_t *a, const wchar_t *b)
	{ return wcscasecmp(a, b); }
	inline static size_t SpanIncluding(const wchar_t *str, const wchar_t *strCharSet)
	{ return wcsspn(str, strCharSet); }
	inline static size_t SpanExcluding(const wchar_t *str, const wchar_t *strCharSet)
	{ return wcscspn(str, strCharSet); }
	inline static wchar_t * FindString(wchar_t *haystack, const wchar_t *needle)
	{ return wcsstr(haystack, needle); }
	inline static bool IsSpace(wchar_t c)
	{ return !!iswspace((wint_t) c); }
	inline static bool IsLower(wchar_t c)
	{ return !!iswlower((wint_t) c); }
	inline static bool IsUpper(wchar_t c)
	{ return !!iswupper((wint_t) c); }
	inline static wchar_t ToLower(wchar_t c)
	{ return IsUpper(c) ? (wchar_t) towlower((wint_t) c) : c; }
	inline static wchar_t ToUpper(wchar_t c)
	{ return IsLower(c) ? (wchar_t) towupper((wint_t) c) : c; }
	inline static DWORD GetEnvironmentVariable(const wchar_t *lpName, wchar_t *lpBuffer, DWORD nSize)
#ifdef _WIN32
	{ return ::GetEnvironmentVariableW(lpName, lpBuffer, nSize); }
#else
	{
		int ret = wcstombs(NULL, lpName, 0);
		if (ret < 0)
			return ret;
		char *sName = (char *) malloc(ret + 1);
		wcstombs(sName, lpName, ret);
		ret = mbstowcs(lpBuffer, getenv(sName), nSize);
		free(sName);
		return ret;
	}
#endif
#ifdef _WIN32
	inline static bool LoadString(CKuStringW &str, HINSTANCE hInstance, UINT nID, WORD wLanguage);
#endif
};
#endif

#ifdef _MSC_VER
#pragma warning(pop) // C4996
#endif

template<typename T, typename U>
class CKuStringT
{
	typedef T * STRT;
	typedef const T * CSTRT;
	typedef INT_PTR SIZET;
public:
	inline static SIZET StringLength(CSTRT str)
	{
		ASSERT(str);
		SIZET len = 0;
		while (*str++)
			len++;
		return len;
	}
	inline static void CopyChars(STRT dst, CSTRT src, SIZET len = -1)
	{
		ASSERT(src);
		ASSERT(dst);
		if (len < 0)
			while (*dst = *src++)
				dst++;
		else {
			if (len > 0)
				memcpy(dst, src, len * sizeof(T));
			dst[len] = 0;
		}
	}
	inline static void CopyCharsOverlapped(STRT dst, CSTRT src, SIZET len = -1)
	{
		ASSERT(src);
		ASSERT(dst);
		if (len < 0)
			len = StringLength(src);
		if (len > 0)
			memmove(dst, src, len * sizeof(T));
		dst[len] = 0;
	}
	inline static SIZET StringLengthSafe(CSTRT str)
	{ return str ? StringLength(str) : 0; }
	inline static STRT Find(STRT str, T ch)
	{
		ASSERT(str);
		for (;*str;*str++)
			if (*str == ch)
				return str;
		return NULL;
	}
	inline static CSTRT Find(CSTRT str, T ch)
	{ return Find((STRT) str, ch); }

	inline static CSTRT GetStringSafe(CSTRT str)
	{ return str ? str : &m_sNULL; }

	typedef bool (*IsFunc)(T ch, CSTRT str);
	inline static bool IsSpace(T ch, CSTRT sSet)
	{ return !!U::IsSpace(ch); }
	inline static bool IsInSet(T ch, CSTRT sSet)
	{ return !!Find(sSet, ch); }
	inline static bool IsTheChar(T ch, CSTRT sSet)
	{ return ch == *sSet; }
	inline static bool IsLower(T ch, CSTRT sSet)
	{ return U::IsLower(ch); }
	inline static bool IsUpper(T ch, CSTRT sSet)
	{ return U::IsUpper(ch); }

	inline static CSTRT Find(CSTRT str, IsFunc isfunc, CSTRT sSet)
	{
		for (;*str;str++)
			if (isfunc(*str, sSet))
				return str;
		return NULL;
	}
private:
	class CKuStringDataT
	{
		friend class CKuStringT;

		CKuStringDataT()
		{ Init(); }
		CKuStringDataT(const CKuStringDataT &d)
		{ Init(); Copy(d); }
		CKuStringDataT(CSTRT src, SIZET len = -1)
		{ Init(); Copy(src, len); }
		CKuStringDataT(T ch)
		{ Init(); Alloc(1);	m_sString[0] = ch; m_iLength = 1; }
		~CKuStringDataT()
		{
			if (m_sData)
				free(m_sData);
		}

		unsigned int AddRef() { return ++m_iRefCount; }
		unsigned int Release()
		{
			unsigned int uRefCount = --m_iRefCount;
			if (uRefCount == 0)
				delete this;
			return uRefCount;
		}

		CKuStringDataT &Copy(CSTRT src, SIZET len = -1)
		{
			if (len == -1)
				len = StringLength(src);
			Alloc(len);
			CopyChars(m_sString, src, len);
			m_iLength = len;
			return *this;
		}
		CKuStringDataT &Copy(const CKuStringDataT &d)
		{ return Copy(d.m_sString, d.m_iLength); }
		bool Alloc(SIZET iCount)
		{
			ASSERT(iCount >= 0);
			ASSERT(m_sString >= m_sData);
			if (m_sData == m_sString) {
				if (!m_sData)
					m_sData = m_sString = (STRT) malloc((iCount + 1) * sizeof(T));
				else if (iCount > m_iSize)
					m_sData = m_sString = (STRT) realloc(m_sData, (iCount + 1) * sizeof(T));
				if (!m_sData)
					return false;
				m_sData[iCount] = 0;
				m_iSize = iCount;
			}
			else if (((SIZET)(m_sString - m_sData)) + iCount > m_iSize) {
				STRT sNew = (STRT) malloc((iCount + 1) * sizeof(T));
				if (!sNew)
					return false;
				CopyChars(sNew, m_sString, m_iLength);
				free(m_sData);
				m_sData = m_sString = sNew;
				m_iSize = iCount;
			}
			return true;
		}
		void Normalize() // let offset be zero
		{
			if (m_sString != m_sData) {
				ASSERT(m_sString);
				ASSERT(m_sData);
				if (m_iLength > 0)
					CopyCharsOverlapped(m_sData, m_sString, m_iLength);
				m_sString = m_sData;
			}
		}
		void FreeExtra(bool bForced = false)
		{
			if (!m_sData)
				return;
			if (m_iLength == 0) {
				free(m_sData);
				m_sData = m_sString = NULL;
				m_iSize = 0;
				Alloc(0);
			}
			else if (m_iSize > (m_iLength << 1) || bForced && m_iSize > m_iLength) { // free extra space only if unused > used or forced flag is specified.
				Normalize();
				VERIFY( m_sData = m_sString = (STRT) realloc(m_sData, (m_iLength + 1) * sizeof(T)) );
				m_iSize = m_iLength;
			}
		}
		void Offset(SIZET iOffset = 1)
		{
			m_sString += iOffset;
			m_iLength -= iOffset;
		}
		void Count()
		{
			m_sData[m_iSize] = 0;
			m_iLength = StringLength(m_sString);
		}
		void Init()
		{
			m_sData = m_sString = NULL;
			m_iSize = m_iLength = 0;
			m_iRefCount = 1;
		}

		STRT m_sString;		// pointer to the string
		STRT m_sData;		// allocated buffer
		SIZET m_iLength;	// current string length
		SIZET m_iSize;		// allocated size in chars w/o '\0'
		unsigned int volatile m_iRefCount; // reference counter
	};

	void Init()
	{
		m_pData = NULL;
		m_bDirty = false;
	}
protected:
	// Make sure data is ready for write. i.e. refcount = 1
	void New(bool bCopy = true)
	{
		ASSERT(!m_bDirty);
		if (!m_pData)
			m_pData = new CKuStringDataT;
		else if (m_pData->m_iRefCount > 1) {
			CKuStringDataT *pData = m_pData;
			if (bCopy)
				m_pData = new CKuStringDataT(*pData);
			else
				m_pData = new CKuStringDataT;
			pData->Release();
		}
		else if (!bCopy)
			m_pData->m_iLength = 0;
	}
private:
	CKuStringT& TrimLeft(IsFunc isfunc, CSTRT sSet)
	{
		ASSERT(isfunc);
		if (!IsEmpty() && isfunc(m_pData->m_sString[0], sSet)) {
			New();
			m_pData->Offset();
			while (isfunc(m_pData->m_sString[0], sSet))
				m_pData->Offset();
		}
		return *this;
	}
	CKuStringT& TrimRight(IsFunc isfunc, CSTRT sSet)
	{
		ASSERT(isfunc);
		if (!IsEmpty() && isfunc(m_pData->m_sString[m_pData->m_iLength - 1] , sSet)) {
			New();
			CSTRT str = m_pData->m_sString;
			STRT ptr = m_pData->m_sString + m_pData->m_iLength - 2;
			for (;ptr >= str && isfunc(*ptr, sSet);)
				ptr--;
			*++ptr = 0;
			m_pData->m_iLength = (SIZET)(ptr - str);
		}
		return *this;
	}
	CKuStringT& Trim(IsFunc isfunc, CSTRT sSet)
	{ TrimRight(isfunc, sSet); return TrimLeft(isfunc, sSet); }
public:
	CKuStringT()
	{ Init(); }
	virtual ~CKuStringT()
	{
		if (m_pData)
			m_pData->Release();
	}

	SIZET GetLength() const
	{ return m_pData ? m_pData->m_iLength : 0; }

	CKuStringT& Empty()
	{
		if (m_pData) {
			m_pData->Release();
			Init();
		}
		return *this;
	}
	bool IsEmpty() const
	{ return GetLength() == 0; }

	CKuStringT& SetString(const CKuStringT& str)
	{
		ASSERT(!str.m_bDirty);
		if (str.IsEmpty())
			Empty();
		else if (m_pData != str.m_pData) {
			if (m_pData)
				m_pData->Release();
			m_pData = str.m_pData;
			m_pData->AddRef();
		}
		return *this;
	}
	CKuStringT(const CKuStringT& str)
	{ Init(); SetString(str); }
	CKuStringT&  operator = (const CKuStringT& str)
	{ return SetString(str); }

	CKuStringT& SetString(CSTRT str, SIZET len = -1)
	{
		if (!str || !*str || !len)
			Empty();
		else {
			if (m_pData)
				m_pData->Release();
			m_pData = new CKuStringDataT(str, len);
			m_bDirty = false;
		}
		return *this;
	}
	CKuStringT(CSTRT str, SIZET len = -1)
	{ Init(); SetString(str, len); }
	CKuStringT&  operator = (CSTRT str)
	{ return SetString(str); }

	CKuStringT& SetString(T ch)
	{
		if (m_pData)
			m_pData->Release();
		m_pData = new CKuStringDataT(ch);
		m_bDirty = false;
		return *this;
	}
	explicit CKuStringT(T ch)
	{ Init(); SetString(ch); }
	CKuStringT&  operator = (T ch)
	{ return SetString(ch); }

	CSTRT GetString() const
	{ return m_pData ? GetStringSafe(m_pData->m_sString) : &m_sNULL; }
	operator CSTRT () const
	{ return GetString(); }

	CKuStringT& SetAt(SIZET i, T ch)
	{ ASSERT(i < GetLength()); m_pData->m_sString[i] = ch; return *this; }
	T GetAt(SIZET i) const
	{ ASSERT(i < GetLength()); return m_pData->m_sString[i]; }
	T &GetAt(SIZET i)
	{ ASSERT(i < GetLength()); return m_pData->m_sString[i]; }
	T operator [] (SIZET i) const
	{ return GetAt(i); }
	T &operator [] (SIZET i)
	{ return GetAt(i); }

	int GetAllocLength() const
	{ return m_pData ? m_pData->m_iSize : 0; }
	CKuStringT& Preallocate(int nLength)
	{
		if (!m_pData)
			m_pData = new CKuStringDataT;
		m_pData->Alloc(nLength);
		return *this;
	}
	CKuStringT& FreeExtra(bool bForced = false)
	{ ASSERT(!m_bDirty); if (m_pData) m_pData->FreeExtra(bForced); return *this; }


	STRT GetBuffer()
	{
		New();
		m_bDirty = true;
		return m_pData->m_sString;
	}
	STRT GetBufferSetLength(SIZET len)
	{
		GetBuffer();
		m_pData->Alloc(len);
		return m_pData->m_sString;
	}
	STRT GetBuffer(SIZET len)
	{ return GetBufferSetLength(len); }
	CKuStringT& ReleaseBufferSetLength(SIZET len)
	{
		if (m_bDirty) {
			m_pData->m_iLength = len;
			m_pData->m_sString[m_pData->m_iLength] = 0;
			m_bDirty = false;
			FreeExtra();
		}
		return *this;
	}
	CKuStringT& ReleaseBuffer(SIZET len = -1)
	{
		if (len >= 0)
			ReleaseBufferSetLength(len);
		else if (m_bDirty) {
			m_pData->Count();
			m_bDirty = false;
			FreeExtra();
		}
		return *this;
	}

	CKuStringT& AttachBuffer(STRT sBuffer, SIZET iCapacity = -1, SIZET iLength = -1)
	{
		ASSERT(sBuffer);
		ASSERT(iCapacity != 0);
		ASSERT(iCapacity >= iLength);
		if (m_pData)
			m_pData->Release();
		m_pData = new CKuStringDataT;
		if (iLength < 0)
			iLength = StringLength(sBuffer);
		if (iCapacity < 0)
			iCapacity = iLength;
		m_pData->m_sData = m_pData->m_sString = sBuffer;
		m_pData->m_iSize = iCapacity;
		m_pData->m_iLength = iLength;
		return *this;
	}
	STRT DetachBuffer()
	{
		if (!m_pData)
			return NULL;
		New();
		m_pData->Normalize();
		STRT sRet = m_pData->m_sData;
		m_pData->m_sData = NULL;
		Empty();
		return sRet;
	}

	CKuStringT& Append(CSTRT str, SIZET len = -1)
	{
		ASSERT(!m_bDirty);
		ASSERT(str);
		if (*str) {
			New();
			if (len < 0)
				len = StringLength(str);
			m_pData->Alloc(m_pData->m_iLength + len);
			CopyChars(m_pData->m_sString + m_pData->m_iLength, str, len);
			m_pData->m_iLength += len;
		}
		return *this;
	}
	CKuStringT& operator += (CSTRT str)
	{ return Append(str); }

	CKuStringT& Append(T ch)
	{
		ASSERT(!m_bDirty);
		New();
		m_pData->Alloc(m_pData->m_iLength + 1);
		m_pData->m_sString[m_pData->m_iLength++] = ch;
		m_pData->m_sString[m_pData->m_iLength] = 0;
		return *this;
	}
	CKuStringT& operator += (T ch)
	{ return Append(ch); }

	friend CKuStringT operator + (const CKuStringT& a, const CKuStringT& b)
	{ CKuStringT ret(a); return ret.Append(b); }
	friend CKuStringT operator + (const CKuStringT& a, CSTRT b)
	{ CKuStringT ret(a); return ret.Append(b); }
	friend CKuStringT operator + (CSTRT a, const CKuStringT& b)
	{ CKuStringT ret(a); return ret.Append(b); }
	friend CKuStringT operator + (const CKuStringT& a, T b)
	{ CKuStringT ret(a); return ret.Append(b); }
	friend CKuStringT operator + (T a, const CKuStringT& b)
	{ CKuStringT ret(a); return ret.Append(b); }

	CKuStringT& Insert(CSTRT str, SIZET idx = 0, SIZET len = -1)
	{
		ASSERT(str);
		ASSERT(idx >= 0);
		if (IsEmpty())
			return SetString(str, len);
		if (idx >= GetLength())
			return Append(str, len);
		if (len)
			len = StringLength(str);
		if (len > 0) {
			New();
			m_pData->Alloc(m_pData->m_iLength + len);
			CopyCharsOverlapped(m_pData->m_sString + idx + len, m_pData->m_sString + idx, m_pData->m_iLength - idx); // including '\0'
			memcpy(m_pData->m_sString + idx, str, len * sizeof(T)); // excluding '\0'
			m_pData->m_iLength += len;
		}
		return *this;
	}
	CKuStringT& Insert(T ch, SIZET idx = 0)
	{
		ASSERT(idx >= 0);
		ASSERT(ch);
		if (IsEmpty())
			return SetString(ch);
		if (idx >= GetLength())
			return Append(ch);
		New();
		m_pData->Alloc(m_pData->m_iLength + 1);
		CopyCharsOverlapped(m_pData->m_sString + idx + 1, m_pData->m_sString + idx, m_pData->m_iLength - idx); // including '\0'
		m_pData->m_sString[idx] = ch;
		m_pData->m_iLength++;
		return *this;
	}

	CKuStringT& Delete(SIZET idx, SIZET len = 1)
	{
		ASSERT(idx >= 0 && idx < GetLength());
		ASSERT(len >= 0);
		if (!IsEmpty() && len > 0) {
			New();
			if (idx + len >= m_pData->m_iLength)
				Truncate(idx);
			else if (idx == 0)
				m_pData->Offset(len);
			else {
				CopyCharsOverlapped(m_pData->m_sString + idx, m_pData->m_sString + idx + len, m_pData->m_iLength - idx - len); // including '\0'
				m_pData->m_iLength -= len;
				FreeExtra();
			}
		}
		return *this;
	}

	CKuStringT& FormatV(CSTRT fmt, va_list ap)
	{
		ASSERT(fmt);
		New(false);
		va_list ap2;
		va_copy(ap2, ap);
		int len = U::FormatV(NULL, 0, fmt, ap);
		ASSERT(len >= 0);
		// NOTE: The length argument of format function has different implementations.
		// Use safe one, though it may waste 1 character space.
		len++;
		m_pData->Alloc(len);
		m_pData->m_iLength = U::FormatV(m_pData->m_sString, len, fmt, ap2);
		return *this;
	}
	CKuStringT& Format(CSTRT fmt, ...)
	{
		va_list ap;
		va_start(ap, fmt);
		return FormatV(fmt, ap);
	}

	CKuStringT& AppendFormatV(CSTRT fmt, va_list ap)
	{
		ASSERT(fmt);
		New();
		va_list ap2;
		va_copy(ap2, ap);
		int len = U::FormatV(NULL, 0, fmt, ap);
		ASSERT(len >= 0);
		// NOTE: The length argument of format function has different implementations.
		// Use safe one, though it may waste 1 character space.
		len++;
		m_pData->Alloc(m_pData->m_iLength + len);
		m_pData->m_iLength += U::FormatV(m_pData->m_sString + m_pData->m_iLength, len, fmt, ap2);
		return *this;
	}
	CKuStringT& AppendFormat(CSTRT fmt, ...)
	{
		va_list ap;
		va_start(ap, fmt);
		return AppendFormatV(fmt, ap);
	}

	CKuStringT& GetEnvironmentVariable(CSTRT sName)
	{
		ASSERT(sName);
		DWORD len = U::GetEnvironmentVariable(sName, NULL, 0);
		if (len == 0)
			Empty();
		else {
			New(false);
			m_pData->Alloc(len);
			m_pData->m_iLength = U::GetEnvironmentVariable(sName, m_pData->m_sString, len);
		}
		return *this;
	}

	CKuStringT Mid(SIZET iFirst, SIZET iLen) const
	{ ASSERT(iFirst + iLen <= GetLength()); return CKuStringT(GetString() + iFirst, iLen); }
	CKuStringT Mid(SIZET iFirst) const
	{ ASSERT(iFirst <= GetLength()); return CKuStringT(GetString() + iFirst); }
	CKuStringT Left(SIZET iLen) const
	{ ASSERT(iLen <= GetLength()); return CKuStringT(GetString(), iLen); }
	CKuStringT Right(SIZET iLen) const
	{ ASSERT(iLen <= GetLength()); return CKuStringT(GetString() + GetLength() - iLen); }

	CKuStringT& Truncate(SIZET nNewLength)
	{
		ASSERT(!m_bDirty);
		ASSERT(nNewLength >= 0 && nNewLength <= GetLength());
		if (nNewLength < GetLength()) {
			New();
			m_pData->m_sString[nNewLength] = 0;
			m_pData->m_iLength = nNewLength;
			FreeExtra();
		}
		return *this;
	}

	CKuStringT& CutMid(SIZET iFirst)
	{
		ASSERT(!m_bDirty);
		ASSERT(iFirst >= 0 && iFirst <= GetLength());
		if (iFirst > 0) {
			New();
			m_pData->Offset(iFirst);
		}
		return *this;
	}
	CKuStringT& CutMid(SIZET iFirst, SIZET iLen)
	{
		if (iLen == 0)
			Empty();
		else {
			CutMid(iFirst);
			Truncate(iLen);
		}
		return *this;
	}
	CKuStringT& CutLeft(SIZET iLen)
	{ Truncate(iLen); return *this; }
	CKuStringT& CutRight(SIZET iLen)
	{ return CutMid(GetLength() - iLen); }

	CKuStringT& TrimLeft()
	{ return TrimLeft(IsSpace, NULL); }
	CKuStringT& TrimLeft(T ch)
	{ return TrimLeft(IsTheChar, &ch); }
	CKuStringT& TrimLeft(CSTRT sSet)
	{ ASSERT(sSet && *sSet); return TrimLeft(IsInSet, sSet); }
	CKuStringT& TrimRight()
	{ return TrimRight(IsSpace, NULL); }
	CKuStringT& TrimRight(T ch)
	{ return TrimRight(IsTheChar, &ch); }
	CKuStringT& TrimRight(CSTRT sSet)
	{ ASSERT(sSet && *sSet); return TrimRight(IsInSet, sSet); }
	CKuStringT& Trim()
	{ return Trim(IsSpace, NULL); }
	CKuStringT& Trim(T ch)
	{ return Trim(IsTheChar, &ch); }
	CKuStringT& Trim(CSTRT sSet)
	{ ASSERT(sSet && *sSet); return Trim(IsInSet, sSet); }

	CKuStringT SpanIncluding(CSTRT pszCharSet) const
	{
		ASSERT(pszCharSet);
		CKuStringT sRet;
		if (!IsEmpty()) {
			size_t len = U::SpanIncluding(m_pData->m_sString, pszCharSet);
			if (len > 0) {
				STRT ptr = sRet.GetBufferSetLength(len);
				CopyChars(ptr, m_pData->m_sString, len);
				sRet.ReleaseBuffer();
			}
		}
		return sRet;
	}

	CKuStringT SpanExcluding(CSTRT pszCharSet) const
	{
		ASSERT(pszCharSet);
		CKuStringT sRet;
		if (!IsEmpty()) {
			size_t len = U::SpanExcluding(m_pData->m_sString, pszCharSet);
			if (len > 0) {
				STRT ptr = sRet.GetBufferSetLength(len);
				CopyChars(ptr, m_pData->m_sString, len);
				sRet.ReleaseBuffer();
			}
		}
		return sRet;
	}

	CKuStringT Tokenize(CSTRT pszTokens, SIZET& iStart) const
	{
		ASSERT(!m_bDirty);
		ASSERT(pszTokens);
		ASSERT(iStart >= 0);
		CKuStringT sRet;
		if (iStart < GetLength()) {
			iStart += U::SpanIncluding(m_pData->m_sString + iStart, pszTokens);
			if (iStart < GetLength()) {
				size_t len = U::SpanExcluding(m_pData->m_sString + iStart, pszTokens);
				if (len > 0) {
					STRT ptr = sRet.GetBufferSetLength(len);
					CopyChars(ptr, m_pData->m_sString + iStart, len);
					sRet.ReleaseBuffer();
				}
				iStart += len;
			}
		}
		return sRet;
	}

	SIZET Find(CSTRT sFind, SIZET iStart = 0) const
	{
		ASSERT(sFind);
		ASSERT(iStart >= 0 && iStart <= GetLength());
		if (IsEmpty() || iStart == GetLength())
			return -1;
		CSTRT s = U::FindString(m_pData->m_sString + iStart, sFind);
		return s ? (SIZET)(s - m_pData->m_sString) : -1;
	}
	SIZET Find(T ch, SIZET iStart = 0) const
	{
		ASSERT(ch);
		ASSERT(iStart >= 0 && iStart <= GetLength());
		if (IsEmpty() || iStart == GetLength())
			return -1;
		CSTRT s = Find(m_pData->m_sString + iStart, ch);
		return s ? (SIZET)(s - m_pData->m_sString) : -1;
	}

	SIZET ReverseFind(T ch) const
	{
		ASSERT(ch);
		if (IsEmpty())
			return -1;
		SIZET i = m_pData->m_iLength - 1;
		for (;i >= 0;i--)
			if (m_pData->m_sString[i] == ch)
				return i;
		return -1;
	}

	CKuStringT& Replace(T chOld, T chNew)
	{
		ASSERT(chOld);
		ASSERT(chNew);
		SIZET iPos = Find(chOld);
		if (!IsEmpty() && iPos >= 0) {
			New();
			for (STRT ptr = m_pData->m_sString + iPos;*ptr;ptr++)
				if (*ptr == chOld) {
					*ptr = chNew;
				}
		}
		return *this;
	}
	CKuStringT& Replace(CSTRT sOld, CSTRT sNew)
	{
		ASSERT(sOld && *sOld);
		SIZET iPos = Find(sOld);

		if (!IsEmpty() && iPos >= 0) {
			SIZET iOld = StringLength(sOld);
			SIZET iNew = StringLengthSafe(sNew);
			SIZET iDiff = iNew - iOld;
			SIZET iStart = iPos, iNext;

			if (iNew > iOld) { // will need grow size
				SIZET iNewLen = m_pData->m_iLength + iDiff;
				
				iPos += iOld;
				while ((iPos = Find(sOld, iPos)) > 0) {
					iNewLen += iDiff;
					iPos += iOld;
				}

				iPos = iStart;
				CKuStringDataT *pNew = new CKuStringDataT;
				pNew->Alloc(iNewLen);
				pNew->m_iLength = iNewLen; // we have known the new string length

				STRT str = pNew->m_sString;
				if (iPos > 0)
					memcpy(str, m_pData->m_sString, iPos * sizeof(T)); // copy beginning mismatched string

				do {
					memcpy(str + iPos, sNew, iNew * sizeof(T));
					iStart += iOld; // next search position
					iPos += iNew; // next copy position
					iNext = Find(sOld, iStart);
					if (iNext > iStart) {
						memcpy(str + iPos, m_pData->m_sString + iStart, (iNext - iStart) * sizeof(T));
						iPos += (iNext - iStart);
						iStart = iNext;
					}
					else if (iNext < 0 && m_pData->m_iLength > iStart)
						memcpy(str + iPos, m_pData->m_sString + iStart, (m_pData->m_iLength - iStart) * sizeof(T)); // last loop, can't find more 'sOld'
				} while (iNext >= 0);
				str[iNewLen] = 0;

				m_pData->Release();
				m_pData = pNew;
			}
			else {
				New();
				SIZET iNewLen = m_pData->m_iLength;
				STRT str = m_pData->m_sString;
				do {
					if (iNew > 0)
						memcpy(str + iPos, sNew, iNew * sizeof(T));
					iNewLen += iDiff;
					iStart += iOld; // next search position
					iPos += iNew; // next copy position
					iNext = Find(sOld, iStart);
					if (iNext > iStart) {
						if (iStart > iPos)
							memmove(str + iPos, str + iStart, (iNext - iStart) * sizeof(T));
						iPos += (iNext - iStart);
						iStart = iNext;
					}
					else if (iNext < 0 && m_pData->m_iLength > iStart && iStart > iPos)
						memmove(str + iPos, str + iStart, (m_pData->m_iLength - iStart) * sizeof(T)); // last loop, can't find more 'sOld'
				} while (iNext >= 0);
				str[iNewLen] = 0;
				m_pData->m_iLength = iNewLen;
				FreeExtra();
			}
		}
		return *this;
	}

	CKuStringT& MakeLower()
	{
		if (!IsEmpty()) {
			STRT str  = (STRT) Find(m_pData->m_sString, IsUpper, NULL);
			if (str) {
				New();
				for (STRT ptr = str;*ptr;ptr++)
					*ptr = U::ToLower(*ptr);
			}
		}
		return *this;
	}
	CKuStringT& MakeUpper()
	{
		if (!IsEmpty()) {
			STRT str  = (STRT) Find(m_pData->m_sString, IsLower, NULL);
			if (str) {
				New();
				for (STRT ptr = str;*ptr;ptr++)
					*ptr = U::ToUpper(*ptr);
			}
		}
		return *this;
	}

	int Compare(CSTRT str) const
	{ ASSERT(str); return U::Compare(GetString(), str); }
	int CompareNoCase(CSTRT str) const
	{ ASSERT(str); return U::CompareNoCase(GetString(), str); }

	friend bool operator == (const CKuStringT& a, const CKuStringT& b)
	{ return a.Compare(b) == 0; }
	friend bool operator != (const CKuStringT& a, const CKuStringT& b)
	{ return a.Compare(b) != 0; }
	friend bool operator > (const CKuStringT& a, const CKuStringT& b)
	{ return a.Compare(b) > 0; }
	friend bool operator < (const CKuStringT& a, const CKuStringT& b)
	{ return a.Compare(b) < 0; }
	friend bool operator >= (const CKuStringT& a, const CKuStringT& b)
	{ return a.Compare(b) >= 0; }
	friend bool operator <= (const CKuStringT& a, const CKuStringT& b)
	{ return a.Compare(b) <= 0; }

	friend bool operator == (const CKuStringT& a, CSTRT b)
	{ return a.Compare(b) == 0; }
	friend bool operator != (const CKuStringT& a, CSTRT b)
	{ return a.Compare(b) != 0; }
	friend bool operator > (const CKuStringT& a, CSTRT b)
	{ return a.Compare(b) > 0; }
	friend bool operator < (const CKuStringT& a, CSTRT b)
	{ return a.Compare(b) < 0; }
	friend bool operator >= (const CKuStringT& a, CSTRT b)
	{ return a.Compare(b) >= 0; }
	friend bool operator <= (const CKuStringT& a, CSTRT b)
	{ return a.Compare(b) <= 0; }

	friend bool operator == (CSTRT a, const CKuStringT& b)
	{ return b.Compare(a) == 0; }
	friend bool operator != (CSTRT a, const CKuStringT& b)
	{ return b.Compare(a) != 0; }
	friend bool operator > (CSTRT a, const CKuStringT& b)
	{ return b.Compare(a) < 0; }
	friend bool operator < (CSTRT a, const CKuStringT& b)
	{ return b.Compare(a) > 0; }
	friend bool operator >= (CSTRT a, const CKuStringT& b)
	{ return b.Compare(a) <= 0; }
	friend bool operator <= (CSTRT a, const CKuStringT& b)
	{ return b.Compare(a) >= 0; }

	int Compare(T ch) const
	{ T str[2] = {ch, 0}; return Compare(str); }
	int CompareNoCase(T ch) const
	{ T str[2] = {ch, 0}; return CompareNoCase(str); }

	friend bool operator == (const CKuStringT& a, T b)
	{ return a.Compare(b) == 0; }
	friend bool operator != (const CKuStringT& a, T b)
	{ return a.Compare(b) != 0; }
	friend bool operator > (const CKuStringT& a, T b)
	{ return a.Compare(b) > 0; }
	friend bool operator < (const CKuStringT& a, T b)
	{ return a.Compare(b) < 0; }
	friend bool operator >= (const CKuStringT& a, T b)
	{ return a.Compare(b) >= 0; }
	friend bool operator <= (const CKuStringT& a, T b)
	{ return a.Compare(b) <= 0; }

	friend bool operator == (T a, const CKuStringT& b)
	{ return b.Compare(a) == 0; }
	friend bool operator != (T a, const CKuStringT& b)
	{ return b.Compare(a) != 0; }
	friend bool operator > (T a, const CKuStringT& b)
	{ return b.Compare(a) < 0; }
	friend bool operator < (T a, const CKuStringT& b)
	{ return b.Compare(a) > 0; }
	friend bool operator >= (T a, const CKuStringT& b)
	{ return b.Compare(a) <= 0; }
	friend bool operator <= (T a, const CKuStringT& b)
	{ return b.Compare(a) >= 0; }

#ifdef _WIN32
	bool LoadString(HINSTANCE hInstance, UINT nID, WORD wLanguage)
	{ return U::LoadString(*this, hInstance, nID, wLanguage); }
	bool LoadString(HINSTANCE hInstance, UINT nID)
	{ return LoadString(hInstance, nID, MAKELANGID(LANG_NEUTRAL, SUBLANG_NEUTRAL)); }
	bool LoadString(UINT nID)
	{ return LoadString(GetModuleHandle(NULL), nID); }

	explicit CKuStringT(HINSTANCE hInstance, UINT nID, WORD wLanguage)
	{ Init(); LoadString(hInstance, nID, wLanguage); }
	explicit CKuStringT(HINSTANCE hInstance, UINT nID)
	{ Init(); LoadString(hInstance, nID); }
	explicit CKuStringT(UINT nID)
	{ Init(); LoadString(nID); }
#endif
private:
	static const T m_sNULL;
	CKuStringDataT *m_pData;
	bool m_bDirty;
};
template<typename T, typename U> const T CKuStringT<T, U>::m_sNULL = 0;

#ifdef _WIN32
inline bool CKuStringUtilW::LoadString(CKuStringW &str, HINSTANCE hInstance, UINT nID, WORD wLanguage)
{
	HRSRC hResource = ::FindResourceEx(hInstance, RT_STRING, MAKEINTRESOURCE(((nID >> 4) + 1)), wLanguage);
	if( hResource == NULL )
		return false;
	HGLOBAL hGlobal = ::LoadResource(hInstance, hResource);
	if (hGlobal) {
		const WORD *pImage;
		const WORD *pImageEnd;
		pImage = (const WORD *) ::LockResource(hGlobal);
		if (pImage) {
			pImageEnd = (const WORD *)(((UINT_PTR) pImage) + ::SizeofResource(hInstance, hResource));
			UINT iIndex = nID & 0xF;

			while(iIndex > 0 && pImage < pImageEnd)
			{
				pImage = (const WORD *)(((UINT_PTR) pImage) + (sizeof(WORD) + (*pImage * sizeof(wchar_t))));
				iIndex--;
			}
			if (pImage < pImageEnd && *pImage > 0) {
				memcpy(str.GetBufferSetLength(*pImage), pImage + 1, *pImage * sizeof(wchar_t));
				str.ReleaseBufferSetLength(*pImage);
				FreeResource(hGlobal);
				return true;
			}
		}
		FreeResource(hGlobal);
	}

	return false;
}

inline bool CKuStringUtilA::LoadString(CKuStringA &str, HINSTANCE hInstance, UINT nID, WORD wLanguage)
{
	CKuStringW s;
	if (!s.LoadString(hInstance, nID, wLanguage))
		return false;
	return !!StringWCharToChar(s, str, -1, '?', CP_ACP);
}
#endif
