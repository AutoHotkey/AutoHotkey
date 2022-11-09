﻿///////////////////////////////////////////////////////////////////////////////
//
// AutoIt
//
// Copyright (C)1999-2007:
//		- Jonathan Bennett <jon@hiddensoft.com>
//		- Others listed at http://www.autoitscript.com/autoit3/docs/credits.htm
//      - Chris Mallett (support@autohotkey.com): adaptation of this file's
//        functions to interface with AutoHotkey.
//
// This program is free software; you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation; either version 2 of the License, or
// (at your option) any later version.
//
///////////////////////////////////////////////////////////////////////////////
//
// script_registry.cpp
//
// Contains registry handling routines.  Part of script.cpp
//
///////////////////////////////////////////////////////////////////////////////


// Includes
#include "stdafx.h" // pre-compiled headers
#include "script.h"
#include "util.h" // for strlcpy()
#include "globaldata.h"
#include "script_func_impl.h"
#include "abi.h"


bif_impl FResult IniRead(StrArg aFilespec, optl<StrArg> aSection, optl<StrArg> aKey, optl<StrArg> aDefault, StrRet &aRetVal)
{
	TCHAR	szFileTemp[T_MAX_PATH];
	TCHAR	*szFilePart, *cp;
	TCHAR	szBuffer[65535];					// Max ini file size is 65535 under 95
	*szBuffer = '\0';
	TCHAR	szEmpty[] = _T("");

	// Get the fullpathname (ini functions need a full path):
	GetFullPathName(aFilespec, _countof(szFileTemp), szFileTemp, &szFilePart);
	if (aKey.has_value()) // But let it be "", which would correspond to an entry such as "=value".
	{
		if (!aSection.has_value()) // But an explicit "" acceptable (see below).
			return FR_E_ARG(1);
		// An access violation can occur if the following conditions are met:
		//	1) aFilespec specifies a Unicode file.
		//	2) aSection is a read-only string, either empty or containing only spaces.
		//
		// Debugging at the assembly level indicates that in this specific situation,
		// it tries to write a zero at the end of aSection (which is already null-
		// terminated).
		//
		// The second condition can ordinarily only be met if Section is omitted,
		// since in all other cases aSection is writable.  Although Section was a
		// required parameter prior to revision 57, empty or blank section names
		// are actually valid.  Simply passing an empty writable buffer appears
		// to work around the problem effectively:
		GetPrivateProfileString(*aSection.value() ? aSection.value() : szEmpty
			, aKey.value(), aDefault.value_or_null(), szBuffer, _countof(szBuffer), szFileTemp);
	}
	else if (aSection.has_value()
		? GetPrivateProfileSection(aSection.value(), szBuffer, _countof(szBuffer), szFileTemp)
		: GetPrivateProfileSectionNames(szBuffer, _countof(szBuffer), szFileTemp))
	{
		// Convert null-terminated array of null-terminated strings to newline-delimited list.
		for (cp = szBuffer; ; ++cp)
			if (!*cp)
			{
				if (!cp[1])
					break;
				*cp = '\n';
			}
	}
	g->LastError = GetLastError();
	if (g->LastError)
	{
		if (!aDefault.has_value())
		{
			if (g->LastError == 2)
				return FError(_T("The requested key, section or file was not found."));
			return FR_E_WIN32(g->LastError);
		}
		if (!aKey.has_value())
		{
			aRetVal.SetTemp(aDefault.value());
			return OK;
		}
	}
	// The above function is supposed to set szBuffer to be aDefault if it can't find the
	// file, section, or key.  In other words, it always changes the contents of szBuffer.
	if (!aRetVal.Copy(szBuffer)) // Avoid using the length the API reported because it might be inaccurate if the data contains any binary zeroes, or if the data is double-terminated, etc.
		return FR_E_OUTOFMEM;
	return OK;
}

#ifdef UNICODE
static BOOL IniEncodingFix(LPCWSTR aFilespec, LPCWSTR aSection)
{
	BOOL result = TRUE;
	if (GetFileAttributes(aFilespec) == INVALID_FILE_ATTRIBUTES) // File does not appear to exist.
	{
		HANDLE hFile;
		DWORD dwWritten;

		// Create a Unicode file. (UTF-16LE BOM)
		hFile = CreateFile(aFilespec, GENERIC_WRITE, 0, NULL, CREATE_NEW, 0, NULL);
		if (hFile != INVALID_HANDLE_VALUE)
		{
			DWORD cc = (DWORD)wcslen(aSection);
			LPTSTR section = talloca(cc + 4);
			sntprintf(section, cc + 4, L"\xFEFF[%s]", aSection);
			DWORD cb = _TSIZE(cc + 3);
			
			// Write a UTF-16LE BOM to identify this as a Unicode file.
			// Write [%aSection%] to prevent WritePrivateProfileString from inserting an empty line (for consistency and style).
			if (!WriteFile(hFile, section, cb, &dwWritten, NULL) || dwWritten != cb)
				result = FALSE;

			if (!CloseHandle(hFile))
				result = FALSE;
		}
	}
	return result;
}
#endif

bif_impl FResult IniWrite(StrArg aValue, StrArg aFilespec, StrArg aSection, optl<StrArg> aKey)
{
	TCHAR	szFileTemp[T_MAX_PATH];
	TCHAR	*szFilePart;
	BOOL	result;
	// Get the fullpathname (INI functions need a full path) 
	GetFullPathName(aFilespec, _countof(szFileTemp), szFileTemp, &szFilePart);
#ifdef UNICODE
	// WritePrivateProfileStringW() always creates INIs using the system codepage.
	// IniEncodingFix() checks if the file exists and if it doesn't then it creates
	// an empty file with a UTF-16LE BOM.
	result = IniEncodingFix(szFileTemp, aSection);
	if(result){
#endif
		if (aKey.has_value()) // But let it be "", which would create an entry such as "=value".
		{
			result = WritePrivateProfileString(aSection, aKey.value(), aValue, szFileTemp);  // Returns zero on failure.
		}
		else
		{
			size_t value_len = _tcslen(aValue);
			TCHAR c, *cp, *szBuffer = talloca(value_len + 2); // +2 for double null-terminator.
			// Convert newline-delimited list to null-terminated array of null-terminated strings.
			for (cp = szBuffer; *aValue; ++cp, ++aValue)
			{
				c = *aValue;
				*cp = (c == '\n') ? '\0' : c;
			}
			*cp = '\0', cp[1] = '\0'; // Double null-terminator.
			result = WritePrivateProfileSection(aSection, szBuffer, szFileTemp);
		}
		if (result)
			WritePrivateProfileString(NULL, NULL, NULL, szFileTemp);	// Flush
#ifdef UNICODE
	}
#endif
	return result ? OK : FR_E_WIN32;
}



bif_impl FResult IniDelete(StrArg aFilespec, StrArg aSection, optl<StrArg> aKey)
// Note that aKey can be NULL, in which case the entire section will be deleted.
{
	TCHAR	szFileTemp[T_MAX_PATH];
	TCHAR	*szFilePart;
	// Get the fullpathname (ini functions need a full path) 
	GetFullPathName(aFilespec, _countof(szFileTemp), szFileTemp, &szFilePart);
	BOOL result = WritePrivateProfileString(aSection, aKey.value_or_null(), NULL, szFileTemp);  // Returns zero on failure.
	g->LastError = GetLastError();
	WritePrivateProfileString(NULL, NULL, NULL, szFileTemp);	// Flush
	return result ? OK : FR_E_WIN32(g->LastError);
}



void RegRead(ResultToken &aResultToken, HKEY aRootKey, LPTSTR aRegSubkey, LPTSTR aValueName, ExprTokenType *aDefault)
{
	HKEY	hRegKey;
	DWORD	dwRes, dwBuf, dwType;
	LONG    result;

	_f_set_retval_p(_T(""), 0); // Set default in case of early return.

	// Open the registry key
	result = RegOpenKeyEx(aRootKey, aRegSubkey, 0, KEY_READ | g->RegView, &hRegKey);
	if (result != ERROR_SUCCESS)
		goto finish_skip_close;

	// Read the value and determine the type.  If aValueName is the empty string, the key's default value is used.
	result = RegQueryValueEx(hRegKey, aValueName, NULL, &dwType, NULL, NULL);
	if (result != ERROR_SUCCESS)
		goto finish;

	LPTSTR contents;

	// The way we read is different depending on the type of the key
	switch (dwType)
	{
		case REG_DWORD:
			dwRes = sizeof(dwBuf);
			result = RegQueryValueEx(hRegKey, aValueName, NULL, NULL, (LPBYTE)&dwBuf, &dwRes);
			if (result == ERROR_SUCCESS)
			{
				// Set the return value but don't return yet.
				aResultToken.Return(dwBuf);
			}
			break;

		// Note: The contents of any of these types can be >64K on NT/2k/XP+ (though that is probably rare):
		case REG_SZ:
		case REG_EXPAND_SZ:
		case REG_MULTI_SZ:
		{
			dwRes = 0; // Retained for backward compatibility because values >64K may cause it to fail on Win95 (unverified, and MSDN implies its value should be ignored for the following call).
			// MSDN: If lpData is NULL, and lpcbData is non-NULL, the function returns ERROR_SUCCESS and stores
			// the size of the data, in bytes, in the variable pointed to by lpcbData.
			result = RegQueryValueEx(hRegKey, aValueName, NULL, NULL, NULL, &dwRes); // Find how large the value is.
			if (result != ERROR_SUCCESS || !dwRes) // Can't find size (realistically might never happen), or size is zero.
				break;
			DWORD dwCharLen = dwRes / sizeof(TCHAR);
			// Set up the result buffer to receive the contents.
			// Since dwRes includes the space for the zero terminator (if the MSDN docs
			// are accurate), this will size it to be 1 byte larger than we need,
			// which leaves room for the final newline character to be inserted after
			// the last item.  But add 2 to the requested capacity in case the data isn't
			// terminated in the registry, which allows double-NULL to be put in for REG_MULTI_SZ later.
			if (!TokenSetResult(aResultToken, NULL, dwCharLen + 2))
				break; // aResultToken.Error() was already called, but we still need to call RegCloseKey().
			//aResultToken.symbol = SYM_STRING; // Already set by _f_set_reval_p().
			contents = aResultToken.marker; // This target buf should now be large enough for the result.

			result = RegQueryValueEx(hRegKey, aValueName, NULL, NULL, (LPBYTE)contents, &dwRes);

			if (result != ERROR_SUCCESS || !dwRes) // Relies on short-circuit boolean order.
			{
				*contents = '\0'; // MSDN says the contents of the buffer is undefined after the call in some cases, so reset it.
			}
			else
			{
				dwCharLen = dwRes / sizeof(TCHAR);
				// See ReadRegString() for more comments about the following:
				// The MSDN docs state that we should ensure that the buffer is NULL-terminated ourselves:
				// "If the data has the REG_SZ, REG_MULTI_SZ or REG_EXPAND_SZ type, then lpcbData will also
				// include the size of the terminating null character or characters ... If the data has the
				// REG_SZ, REG_MULTI_SZ or REG_EXPAND_SZ type, the string may not have been stored with the
				// proper null-terminating characters. Applications should ensure that the string is properly
				// terminated before using it, otherwise, the application may fail by overwriting a buffer."
				//
				// The MSDN docs cited above are a little unclear.  The most likely interpretation is that
				// dwRes contains the true size retrieved.  For example, if dwRes is 1, the first char
				// in the buffer is either a NULL or an actual non-NULL character that was originally
				// stored in the registry incorrectly (i.e. without a terminator).
				//
				// Include all data up to dwCharLen, excluding one null-terminator if present,
				// but ensure the final value is properly null-terminated outside of dwCharLen.
				if (!contents[dwCharLen - 1])
					--dwCharLen;
				else
					contents[dwCharLen] = '\0';
				if (dwType == REG_MULTI_SZ) // Convert NULL-delimiters into newline delimiters.
				{
					// Translate each null-terminated string within dwCharLen to a newline-terminated string.
					// There used to be a comment here indicating that terminating the final item with \n
					// makes parsing easier, but I don't think that's true when using Loop Parse/StrSplit.
					// In any case, RegRead's documentation indicates that all items will end with `n.
					// Cases covered here:
					//  A. item1\0item2\0\0  ; Correctly terminated (one \0 was excluded from dwCharLen above)
					//  B. item1\0item2\0    ; Incorrectly terminated
					//  C. item1\0item2      ; Incorrectly unterminated (but it was terminated above)
					// RegRead distinguishes between A & B/C by only showing a trailing blank line for A,
					// but doing that here would be contrary to the documentation and probably wouldn't be
					// useful since B and C are technically invalid.
					if (!dwCharLen // Value was just \0 on its own (dwCharLen is 0 due to --dwCharLen above).
						|| contents[dwCharLen - 1]) // Case B or C above.
					{
						++dwCharLen; // Let the null-terminator (guaranteed above) be translated to \n.
						contents[dwCharLen] = '\0'; // Ensure there's still an actual null-terminator.
					}
					// Empty items (signified by \0\0 within the value) are specifically permitted by this loop,
					// and are known to commonly occur in the PendingRenameOperations value written by MoveFileEx.
					for (DWORD i = 0; i < dwCharLen; ++i)
						if (!contents[i])
							contents[i] = '\n';
				}
			}
			aResultToken.marker_length = _tcslen(contents); // Due to conservative buffer sizes above, length is probably too large by 3. So update to reflect the true length.
			break;
		}
		case REG_BINARY:
		{
			// See case REG_SZ for comments.
			dwRes = 0;
			result = RegQueryValueEx(hRegKey, aValueName, NULL, NULL, NULL, &dwRes); // Find how large the value is.
			if (result != ERROR_SUCCESS || !dwRes) // Can't find size (realistically might never happen), or size is zero.
				break;

			// Set up the result buffer.
			// AutoIt3: Each byte will turned into 2 digits, plus a final null:
			if (!TokenSetResult(aResultToken, NULL, dwRes * 2))
				break; // aResultToken.Error() was already called, but we still need to call RegCloseKey().
			//aResultToken.symbol = SYM_STRING; // Already set by _f_set_reval_p().
			contents = aResultToken.marker;
			*contents = '\0';
			
			// Read the binary data into the variable, placed so that the last byte of
			// binary data will be overwritten as the hexadecimal conversion completes.
			LPBYTE pRegBuffer = (LPBYTE)(contents + dwRes * 2) - dwRes;
			result = RegQueryValueEx(hRegKey, aValueName, NULL, NULL, pRegBuffer, &dwRes);
			if (result != ERROR_SUCCESS)
			{
				aResultToken.marker_length = 0; // Override the length set by TokenSetResult().
				break;
			}

			int j = 0;
			DWORD i, n; // i and n must be unsigned to work
			static const TCHAR szHexData[] = _T("0123456789ABCDEF");
			for (i = 0; i < dwRes; ++i)
			{
				n = pRegBuffer[i];				// Get the value and convert to 2 digit hex
				contents[j + 1] = szHexData[n % 16];
				n /= 16;
				contents[j] = szHexData[n % 16];
				j += 2;
			}
			contents[j] = '\0'; // Terminate
			// aResultToken.marker_length was already set by TokenSetResult().
			break;
		}
		default:
			result = ERROR_UNSUPPORTED_TYPE; // Indicate the error.
			break;
	}

finish:
	// The function is designed to reach this point only if the key was opened,
	// rather than initializing hRegKey to NULL and checking here, since it's
	// not clear whether NULL is actually an invalid registry handle value:
	//if (hRegKey)
	RegCloseKey(hRegKey);
finish_skip_close:
	g->LastError = result;
	if (result == ERROR_FILE_NOT_FOUND && aDefault)
	{
		aResultToken.CopyValueFrom(*aDefault);
		if (aResultToken.symbol == SYM_OBJECT)
			aResultToken.object->AddRef(); // Must be done for safety, and supporting objects (returned as is) might have some utility.
	}
	else if (result != ERROR_SUCCESS)
		_f_throw_win32(g->LastError);
} // RegRead()



void RegWrite(ResultToken &aResultToken, ExprTokenType *aValue, DWORD aValueType, HKEY aRootKey, LPTSTR aRegSubkey, LPTSTR aValueName)
// If aValueName is the empty string, the key's default value is used.
// if aValue is NULL, aValueType must be REG_NONE; a key will be created but no value will be set.
{
	HKEY	hRegKey;
	DWORD	dwBuf;

	TCHAR nbuf[MAX_NUMBER_SIZE];
	LPTSTR value;
	size_t length;
	
	LONG result;

	if (aValue) // RegWrite, not RegCreateKey
	{
		if (aValueType == REG_NONE) // Omitted
			return (void)aResultToken.ParamError(1, nullptr);

		if (aValueType != REG_DWORD)
			value = TokenToString(*aValue, nbuf, &length);
		else if (TokenIsNumeric(*aValue))
			dwBuf = (DWORD)TokenToInt64(*aValue);
		else
			return (void)aResultToken.ParamError(0, aValue, _T("Number"));
	}

	// Open/Create the registry key
	// The following works even on root keys (i.e. blank subkey), although values can't be created/written to
	// HKCU's root level, perhaps because it's an alias for a subkey inside HKEY_USERS.  Even when RegOpenKeyEx()
	// is used on HKCU (which is probably redundant since it's a pre-opened key?), the API can't create values
	// there even though RegEdit can.
	result = RegCreateKeyEx(aRootKey, aRegSubkey, 0, _T(""), REG_OPTION_NON_VOLATILE, KEY_WRITE | g->RegView, NULL, &hRegKey, NULL);
	if (result != ERROR_SUCCESS)
		goto finish;

	// Write the registry differently depending on type of variable we are writing
	switch (aValueType)
	{
	case REG_SZ:
		result = RegSetValueEx(hRegKey, aValueName, 0, REG_SZ, (CONST BYTE *)value, (DWORD)(length+1) * sizeof(TCHAR));
		break;

	case REG_EXPAND_SZ:
		result = RegSetValueEx(hRegKey, aValueName, 0, REG_EXPAND_SZ, (CONST BYTE *)value, (DWORD)(length+1) * sizeof(TCHAR));
		break;
	
	case REG_MULTI_SZ:
	{
		// Allocate some temporary memory because aValue might not be a writable string,
		// and we would need to write to it to temporarily change the newline delimiters into
		// zero-delimiters.  Even if we were to require callers to give us a modifiable string,
		// its capacity may be 1 byte too small to handle the double termination that's needed
		// (i.e. if the last item in the list happens to not end in a newline):
		LPTSTR buf = tmalloc(length + 2);
		if (!buf)
		{
			result = ERROR_OUTOFMEMORY;
			break;
		}
		tcslcpy(buf, value, length + 1);
		// Double-terminate:
		buf[length + 1] = '\0';

		// Remove any final newline the user may have provided since we don't want the length
		// to include it when calling RegSetValueEx() -- it would be too large by 1:
		if (length > 0 && buf[length - 1] == '\n')
			buf[--length] = '\0';

		// Replace the script's delimiter char with the zero-delimiter needed by RegSetValueEx():
		for (LPTSTR cp = buf; *cp; ++cp)
			if (*cp == '\n')
				*cp = '\0';

		result = RegSetValueEx(hRegKey, aValueName, 0, REG_MULTI_SZ, (CONST BYTE *)buf
								, (DWORD)(length ? length + 2 : 0) * sizeof(TCHAR));
		free(buf);
		break;
	}

	case REG_DWORD:
		result = RegSetValueEx(hRegKey, aValueName, 0, REG_DWORD, (CONST BYTE *)&dwBuf, sizeof(dwBuf));
		break;

	case REG_BINARY:
	{
		int nLen = (int)length;

		// String length must be a multiple of 2 
		if (nLen % 2)
		{
			result = ERROR_INVALID_PARAMETER;
			break;
		}

		int nBytes = nLen / 2;
		LPBYTE pRegBuffer = (LPBYTE) malloc(nBytes);
		if (!pRegBuffer)
		{
			result = ERROR_OUTOFMEMORY;
			break;
		}

		// Really crappy hex conversion
		int j = 0, i = 0, nVal, nMult;
		while (i < nLen && j < nBytes)
		{
			nVal = 0;
			for (nMult = 16; nMult >= 0; nMult = nMult - 15)
			{
				if (value[i] >= '0' && value[i] <= '9')
					nVal += (value[i] - '0') * nMult;
				else if (value[i] >= 'A' && value[i] <= 'F')
					nVal += (((value[i] - 'A'))+10) * nMult;
				else if (value[i] >= 'a' && value[i] <= 'f')
					nVal += (((value[i] - 'a'))+10) * nMult;
				else
				{
					free(pRegBuffer);
					RegCloseKey(hRegKey);
					result = ERROR_INVALID_PARAMETER;
					goto finish;
				}
				++i;
			}
			pRegBuffer[j++] = (char)nVal;
		}

		result = RegSetValueEx(hRegKey, aValueName, 0, REG_BINARY, pRegBuffer, (DWORD)j);
		free(pRegBuffer);
		break;
	}

	case REG_NONE: // RegCreateKey
		break;

	default: // Could be any unsupported value stored in reg_item.type by RegEnumValue().
		result = ERROR_INVALID_PARAMETER; // Anything other than ERROR_SUCCESS.
		break;
	} // switch()

	RegCloseKey(hRegKey);
	// Additionally, fall through to below:

finish:
	g->LastError = result;
	if (result != ERROR_SUCCESS)
		_f_throw_win32(result);
	_f_return_empty;
} // RegWrite()



LONG RegRemoveSubkeys(HKEY hRegKey)
{
	// Removes all subkeys to the given key.  Will not touch the given key.
	TCHAR Name[256];
	DWORD dwNameSize;
	FILETIME ftLastWrite;
	HKEY hSubKey;
	LONG result;

	for (;;) 
	{ // infinite loop 
		dwNameSize = _countof(Name)-1;
		if (RegEnumKeyEx(hRegKey, 0, Name, &dwNameSize, NULL, NULL, NULL, &ftLastWrite) == ERROR_NO_MORE_ITEMS)
			return ERROR_SUCCESS;
		result = RegOpenKeyEx(hRegKey, Name, 0, KEY_READ | g->RegView, &hSubKey);
		if (result != ERROR_SUCCESS)
			break;
		result = RegRemoveSubkeys(hSubKey);
		RegCloseKey(hSubKey);
		if (result != ERROR_SUCCESS)
			break;
		result = RegDeleteKey(hRegKey, Name);
		if (result != ERROR_SUCCESS)
			break;
	}
	return result;
}



void RegDelete(ResultToken &aResultToken, HKEY aRootKey, LPTSTR aRegSubkey, LPTSTR aValueName)
{
	LONG result;

	if (!aRootKey)
		_f_throw_value(ERR_PARAM1_INVALID);
	// Fix for v1.0.48: Don't remove the entire key if it's a root key!  According to MSDN,
	// the root key would be opened by RegOpenKeyEx() further below whenever aRegSubkey is NULL
	// or an empty string. aValueName is also checked to preserve the ability to delete a value
	// that exists directly under a root key.
	if (  (!aRegSubkey || !*aRegSubkey) && !aValueName  )
		_f_throw_value(_T("Cannot delete root key"));

	// Open the key we want
	HKEY hRegKey;
	result = RegOpenKeyEx(aRootKey, aRegSubkey, 0, KEY_READ | KEY_WRITE | g->RegView, &hRegKey);
	if (result != ERROR_SUCCESS)
		goto finish;

	if (!aValueName) // Caller's signal to delete the entire subkey.
	{
		// Remove the entire Key
		result = RegRemoveSubkeys(hRegKey); // Delete any subitems within the key.
		RegCloseKey(hRegKey); // Close parent key.  Not sure if this needs to be done only after the above.
		if (result == ERROR_SUCCESS)
		{
			if (g->RegView)
				result = RegDeleteKeyEx(aRootKey, aRegSubkey, g->RegView, 0);
			else
				result = RegDeleteKey(aRootKey, aRegSubkey);
		}
	}
	else
	{
		// Remove the value.
		result = RegDeleteValue(hRegKey, aValueName);
		RegCloseKey(hRegKey);
	}

finish:
	g->LastError = result;
	if (result != ERROR_SUCCESS)
		_f_throw_win32(result);
	_f_return_empty;
} // RegDelete()



struct RegRootKeyType
{
	LPTSTR short_name;
	LPTSTR long_name;
	HKEY key;
};

static RegRootKeyType sRegRootKeyTypes[] =
{
	{_T("HKLM"), _T("HKEY_LOCAL_MACHINE"), HKEY_LOCAL_MACHINE},
	{_T("HKCR"), _T("HKEY_CLASSES_ROOT"), HKEY_CLASSES_ROOT},
	{_T("HKCC"), _T("HKEY_CURRENT_CONFIG"), HKEY_CURRENT_CONFIG},
	{_T("HKCU"), _T("HKEY_CURRENT_USER"), HKEY_CURRENT_USER},
	{_T("HKU"), _T("HKEY_USERS"), HKEY_USERS}
};

HKEY Line::RegConvertRootKeyType(LPTSTR aName)
{
	for (int i = 0; i < _countof(sRegRootKeyTypes); ++i)
		if (!_tcsicmp(aName, sRegRootKeyTypes[i].short_name)
			|| !_tcsicmp(aName, sRegRootKeyTypes[i].long_name))
			return sRegRootKeyTypes[i].key;
	return NULL;
}

LPTSTR Line::RegConvertRootKeyType(HKEY aKey)
{
	for (int i = 0; i < _countof(sRegRootKeyTypes); ++i)
		if (aKey == sRegRootKeyTypes[i].key)
			return sRegRootKeyTypes[i].long_name;
	// These are either unused or so rarely used that they aren't supported:
	// HKEY_PERFORMANCE_DATA, HKEY_PERFORMANCE_TEXT, HKEY_PERFORMANCE_NLSTEXT
	return _T("");
}



BIF_DECL(BIF_Reg)
{
	TCHAR key_buf[MAX_REG_ITEM_SIZE];
	BuiltInFunctionID action = _f_callee_id;
	HKEY root_key;
	LPTSTR sub_key;
	LPTSTR value_name = action == FID_RegDeleteKey ? NULL : _T(""); // Set default.
	DWORD value_type = REG_NONE; // RegWrite
	ExprTokenType *value = nullptr; // RegWrite
	bool close_root;
	if (action == FID_RegWrite)
	{
		value = aParam[0];
		if (aParamCount > 1)
		{
			if (aParam[1]->symbol != SYM_MISSING)
			{
				value_type = Line::RegConvertValueType(ParamIndexToString(1));
				if (value_type == REG_NONE) // In this case REG_NONE means unknown/invalid vs. omitted.
					_f_throw_param(1);
			}
			aParamCount -= 2;
		}
		else
			aParamCount = 0; // It was 1.
		aParam += 2; // There might not have been 2, but it won't be dereferenced in that case anyway.
	}
	if (ParamIndexIsOmitted(0) && g->mLoopRegItem) // Default to the registry loop's current item.
	{
		RegItemStruct &reg_item = *g->mLoopRegItem;
		root_key = reg_item.root_key;
		if (reg_item.type == REG_SUBKEY)
		{
			if (*reg_item.subkey)
			{
				sntprintf(key_buf, _countof(key_buf), _T("%s\\%s"), reg_item.subkey, reg_item.name);
				sub_key = key_buf;
			}
			else
				sub_key = reg_item.name;
		}
		else
		{
			sub_key = reg_item.subkey;
			if (action != FID_RegDeleteKey)
			{
				value_name = reg_item.name; // Set default.
				if (value_type == REG_NONE)
					value_type = reg_item.type;
			}
		}
		// Do not use RegCloseKey() on this, even if it's a remote key, since our caller handles that:
		close_root = false;
	}
	else
	{
		LPTSTR key_name = ParamIndexToOptionalString(0); // No buf needed since numbers aren't valid root keys.
		root_key = Line::RegConvertKey(key_name, &sub_key, &close_root);
		if (!root_key)
			return (void)aResultToken.ParamError(action == FID_RegWrite ? 2 : 0, aParamCount ? aParam[0] : nullptr);
	}
	if (!ParamIndexIsOmitted(1)) // Implies this isn't RegDeleteKey.
	{
		if (ParamIndexToObject(1))
			return (void)aResultToken.ParamError(action == FID_RegWrite ? 3 : 1, aParam[1], _T("String"));
		value_name = ParamIndexToString(1, _f_number_buf);
	}

	switch (action)
	{
	case FID_RegRead:  RegRead(aResultToken, root_key, sub_key, value_name, ParamIndexIsOmitted(2) ? nullptr : aParam[2]); break;
	case FID_RegCreateKey:
	case FID_RegWrite: RegWrite(aResultToken, value, value_type, root_key, sub_key, value_name); break;
	default:           RegDelete(aResultToken, root_key, sub_key, value_name); break;
	}

	if (close_root)
		RegCloseKey(root_key);
}
