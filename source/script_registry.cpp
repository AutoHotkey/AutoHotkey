///////////////////////////////////////////////////////////////////////////////
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


ResultType Line::IniRead(LPTSTR aFilespec, LPTSTR aSection, LPTSTR aKey, LPTSTR aDefault)
{
	if (!aDefault || !*aDefault)
		aDefault = _T("ERROR");  // This mirrors what AutoIt2 does for its default value.
	TCHAR	szFileTemp[_MAX_PATH+1];
	TCHAR	*szFilePart, *cp;
	TCHAR	szBuffer[65535] = _T("");					// Max ini file size is 65535 under 95
	// Get the fullpathname (ini functions need a full path):
	GetFullPathName(aFilespec, _MAX_PATH, szFileTemp, &szFilePart);
	if (*aKey)
	{
		GetPrivateProfileString(aSection, aKey, aDefault, szBuffer, _countof(szBuffer), szFileTemp);
	}
	else if (*aSection
		? GetPrivateProfileSection(aSection, szBuffer, _countof(szBuffer), szFileTemp)
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
	// The above function is supposed to set szBuffer to be aDefault if it can't find the
	// file, section, or key.  In other words, it always changes the contents of szBuffer.
	return OUTPUT_VAR->Assign(szBuffer); // Avoid using the length the API reported because it might be inaccurate if the data contains any binary zeroes, or if the data is double-terminated, etc.
	// Note: ErrorLevel is not changed by this command since the aDefault value is returned
	// whenever there's an error.
}

#ifdef UNICODE
static BOOL IniEncodingFix(LPWSTR aFilespec, LPWSTR aSection)
{
	BOOL result = TRUE;
	if (!DoesFilePatternExist(aFilespec))
	{
		HANDLE hFile;
		DWORD dwWritten;

		// Create a Unicode file. (UTF-16LE BOM)
		hFile = CreateFile(aFilespec, GENERIC_WRITE, 0, NULL, CREATE_NEW, 0, NULL);
		if (hFile != INVALID_HANDLE_VALUE)
		{
			DWORD cc = (DWORD)wcslen(aSection);
			DWORD cb = (cc + 1) * sizeof(WCHAR);
			
			aSection[cc] = ']'; // Temporarily replace the null-terminator.

			// Write a UTF-16LE BOM to identify this as a Unicode file.
			// Write [%aSection%] to prevent WritePrivateProfileString from inserting an empty line (for consistency and style).
			if (   !WriteFile(hFile, L"\xFEFF[", 4, &dwWritten, NULL) || dwWritten != 4
				|| !WriteFile(hFile, aSection, cb, &dwWritten, NULL) || dwWritten != cb   )
				result = FALSE;

			if (!CloseHandle(hFile))
				result = FALSE;

			aSection[cc] = '\0'; // Re-terminate.
		}
	}
	return result;
}
#endif

ResultType Line::IniWrite(LPTSTR aValue, LPTSTR aFilespec, LPTSTR aSection, LPTSTR aKey)
{
	TCHAR	szFileTemp[_MAX_PATH+1];
	TCHAR	*szFilePart;
	BOOL	result;
	// Get the fullpathname (INI functions need a full path) 
	GetFullPathName(aFilespec, _MAX_PATH, szFileTemp, &szFilePart);
#ifdef UNICODE
	// WritePrivateProfileStringW() always creates INIs using the system codepage.
	// IniEncodingFix() checks if the file exists and if it doesn't then it creates
	// an empty file with a UTF-16LE BOM.
	result = IniEncodingFix(szFileTemp, aSection);
	if(result){
#endif
		if (*aKey)
		{
			result = WritePrivateProfileString(aSection, aKey, aValue, szFileTemp);  // Returns zero on failure.
		}
		else
		{
			size_t value_len = ArgLength(1);
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
		WritePrivateProfileString(NULL, NULL, NULL, szFileTemp);	// Flush
#ifdef UNICODE
	}
#endif
	return g_script.mIsAutoIt2 ? OK : g_ErrorLevel->Assign(result ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
}



ResultType Line::IniDelete(LPTSTR aFilespec, LPTSTR aSection, LPTSTR aKey)
// Note that aKey can be NULL, in which case the entire section will be deleted.
{
	TCHAR	szFileTemp[_MAX_PATH+1];
	TCHAR	*szFilePart;
	// Get the fullpathname (ini functions need a full path) 
	GetFullPathName(aFilespec, _MAX_PATH, szFileTemp, &szFilePart);
	BOOL result = WritePrivateProfileString(aSection, aKey, NULL, szFileTemp);  // Returns zero on failure.
	WritePrivateProfileString(NULL, NULL, NULL, szFileTemp);	// Flush
	return g_script.mIsAutoIt2 ? OK : g_ErrorLevel->Assign(result ? ERRORLEVEL_NONE : ERRORLEVEL_ERROR);
}



ResultType Line::RegRead(HKEY aRootKey, LPTSTR aRegSubkey, LPTSTR aValueName)
{
	Var &output_var = *OUTPUT_VAR;
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.
	output_var.Assign(); // Init.  Tell it not to free the memory by not calling with "".

	HKEY	hRegKey;
	DWORD	dwRes, dwBuf, dwType;
	LONG    result;
	// My: Seems safest to keep the limit just below 64K in case Win95 has problems with larger values.
	TCHAR	szRegBuffer[65535]; // Only allow reading of 64Kb from a key

	if (!aRootKey)
		return OK;  // Let ErrorLevel tell the story.

	// Open the registry key
	if (RegOpenKeyEx(aRootKey, aRegSubkey, 0, KEY_READ, &hRegKey) != ERROR_SUCCESS)
		return OK;  // Let ErrorLevel tell the story.

	// Read the value and determine the type.  If aValueName is the empty string, the key's default value is used.
	if (RegQueryValueEx(hRegKey, aValueName, NULL, &dwType, NULL, NULL) != ERROR_SUCCESS)
	{
		RegCloseKey(hRegKey);
		return OK;  // Let ErrorLevel tell the story.
	}

	LPTSTR contents, cp;

	// The way we read is different depending on the type of the key
	switch (dwType)
	{
		case REG_DWORD:
			dwRes = sizeof(dwBuf);
			RegQueryValueEx(hRegKey, aValueName, NULL, NULL, (LPBYTE)&dwBuf, &dwRes);
			RegCloseKey(hRegKey);
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
			return output_var.Assign((DWORD)dwBuf);

		// Note: The contents of any of these types can be >64K on NT/2k/XP+ (though that is probably rare):
		case REG_SZ:
		case REG_EXPAND_SZ:
		case REG_MULTI_SZ:
		{
			dwRes = 0; // Retained for backward compatibility because values >64K may cause it to fail on Win95 (unverified, and MSDN implies its value should be ignored for the following call).
			// MSDN: If lpData is NULL, and lpcbData is non-NULL, the function returns ERROR_SUCCESS and stores
			// the size of the data, in bytes, in the variable pointed to by lpcbData.
			if (RegQueryValueEx(hRegKey, aValueName, NULL, NULL, NULL, &dwRes) != ERROR_SUCCESS // Find how large the value is.
				|| !dwRes) // Can't find size (realistically might never happen), or size is zero.
			{
				RegCloseKey(hRegKey);
				return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // For backward compatibility, indicate success (these conditions should be very rare anyway).
			}
			DWORD dwCharLen = dwRes / sizeof(TCHAR);
			// Set up the variable to receive the contents, enlarging it if necessary:
			// Since dwRes includes the space for the zero terminator (if the MSDN docs
			// are accurate), this will enlarge it to be 1 byte larger than we need,
			// which leaves room for the final newline character to be inserted after
			// the last item.  But add 2 to the requested capacity in case the data isn't
			// terminated in the registry, which allows double-NULL to be put in for REG_MULTI_SZ later.
			if (output_var.AssignString(NULL, (VarSizeType)(dwCharLen + 2)) != OK)
			{
				RegCloseKey(hRegKey);
				return FAIL; // FAIL is only returned when the error is a critical one such as this one.
			}

			contents = output_var.Contents(); // This target buf should now be large enough for the result.

			result = RegQueryValueEx(hRegKey, aValueName, NULL, NULL, (LPBYTE)contents, &dwRes);
			RegCloseKey(hRegKey);

			if (result != ERROR_SUCCESS || !dwRes) // Relies on short-circuit boolean order.
				*contents = '\0'; // MSDN says the contents of the buffer is undefined after the call in some cases, so reset it.
				// Above realistically probably never happens; for backward compatibility (and simplicity),
				// consider it a success.
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
				// Double-terminate so that the loop can find out the true end of the buffer.
				// The MSDN docs cited above are a little unclear.  The most likely interpretation is that
				// dwRes contains the true size retrieved.  For example, if dwRes is 1, the first char
				// in the buffer is either a NULL or an actual non-NULL character that was originally
				// stored in the registry incorrectly (i.e. without a terminator).  In either case, do
				// not change the first character, just leave it as is and add a NULL at the 2nd and
				// 3rd character positions to ensure that it is double terminated in every case:
				contents[dwCharLen] = contents[dwCharLen + 1] = '\0';

				if (dwType == REG_MULTI_SZ) // Convert NULL-delimiters into newline delimiters.
				{
					for (cp = contents;; ++cp)
					{
						if (!*cp)
						{
							// Unlike AutoIt3, it seems best to have a newline character after the
							// last item in the list also.  It usually makes parsing easier:
							*cp = '\n';	// Convert to \n for later storage in the user's variable.
							if (!*(cp + 1)) // Buffer is double terminated, so this is safe.
								// Double null terminator marks the end of the used portion of the buffer.
								break;
						}
					}
					// else the buffer is empty (see above notes for explanation).  So don't put any newlines
					// into it at all, since each newline should correspond to an item in the buffer.
				}
			}
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
			output_var.SetCharLength((VarSizeType)_tcslen(contents)); // Due to conservative buffer sizes above, length is probably too large by 3. So update to reflect the true length.
			return output_var.Close(); // Must be called after Assign(NULL, ...) or when Contents() has been altered because it updates the variable's attributes and properly handles VAR_CLIPBOARD.
		}
		case REG_BINARY:
		{
			LPBYTE pRegBuffer = (LPBYTE) szRegBuffer;
			dwRes = sizeof(szRegBuffer);
			result = RegQueryValueEx(hRegKey, aValueName, NULL, NULL, pRegBuffer, &dwRes);
			RegCloseKey(hRegKey);

			if (result == ERROR_MORE_DATA)
				// The API docs state that the buffer's contents are undefined in this case,
				// so for no we don't support values larger than our buffer size:
				return OK; // Let ErrorLevel tell the story.  The output variable has already been made blank.

			// Set up the variable to receive the contents, enlarging it if necessary.
			// AutoIt3: Each byte will turned into 2 digits, plus a final null:
			if (output_var.AssignString(NULL, (VarSizeType)(dwRes * 2)) != OK)
				return FAIL;
			contents = output_var.Contents();
			*contents = '\0';

			int j = 0;
			DWORD i, n; // i and n must be unsigned to work
			TCHAR szHexData[] = _T("0123456789ABCDEF");  // Access to local vars might be faster than static ones.
			for (i = 0; i < dwRes; ++i)
			{
				n = pRegBuffer[i];				// Get the value and convert to 2 digit hex
				contents[j + 1] = szHexData[n % 16];
				n /= 16;
				contents[j] = szHexData[n % 16];
				j += 2;
			}
			contents[j] = '\0'; // Terminate
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
			return output_var.Close(); // Length() was already set by the earlier call to Assign().
		}
	}

	// Since above didn't return, this is an unsupported value type.
	return OK;  // Let ErrorLevel tell the story.
} // RegRead()



ResultType Line::RegWrite(DWORD aValueType, HKEY aRootKey, LPTSTR aRegSubkey, LPTSTR aValueName, LPTSTR aValue)
// If aValueName is the empty string, the key's default value is used.
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.

	HKEY	hRegKey;
	DWORD	dwRes, dwBuf;

	// My: Seems safest to keep the limit just below 64K in case Win95 has problems with larger values.
	TCHAR szRegBuffer[65535], *buf; // Only allow writing of 64Kb to a key for Win9x, which is all it supports.
	#define SET_REG_BUF \
		if (g_os.IsWin9x())\
		{\
			tcslcpy(szRegBuffer, aValue, sizeof(szRegBuffer));\
			buf = szRegBuffer;\
		}\
		else\
			buf = aValue;

	if (!aRootKey || aValueType == REG_NONE || aValueType == REG_SUBKEY) // Can't write to these.
		return OK;  // Let ErrorLevel tell the story.

	// Open/Create the registry key
	// The following works even on root keys (i.e. blank subkey), although values can't be created/written to
	// HKCU's root level, perhaps because it's an alias for a subkey inside HKEY_USERS.  Even when RegOpenKeyEx()
	// is used on HKCU (which is probably redundant since it's a pre-opened key?), the API can't create values
	// there even though RegEdit can.
	if (RegCreateKeyEx(aRootKey, aRegSubkey, 0, _T(""), REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hRegKey, &dwRes)
		!= ERROR_SUCCESS)
		return OK;  // Let ErrorLevel tell the story.

	// Write the registry differently depending on type of variable we are writing
	switch (aValueType)
	{
	case REG_SZ:
		SET_REG_BUF
		if (RegSetValueEx(hRegKey, aValueName, 0, REG_SZ, (CONST BYTE *)buf, (DWORD)(_tcslen(buf)+1) * sizeof(TCHAR)) == ERROR_SUCCESS)
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		RegCloseKey(hRegKey);
		return OK;

	case REG_EXPAND_SZ:
		SET_REG_BUF
		if (RegSetValueEx(hRegKey, aValueName, 0, REG_EXPAND_SZ, (CONST BYTE *)buf, (DWORD)(_tcslen(buf)+1) * sizeof(TCHAR)) == ERROR_SUCCESS)
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		RegCloseKey(hRegKey);
		return OK;
	
	case REG_MULTI_SZ:
	{
		// Don't allow values over 64K for this type because aValue might not be a writable
		// string, and we would need to write to it to temporarily change the newline delimiters
		// into zero-delimiters.  Even if we were to require callers to give us a modifiable string,
		// its capacity be 1 byte too small to handle the double termination that's needed
		// (i.e. if the last item in the list happens to not end in a newline):
		tcslcpy(szRegBuffer, aValue, _countof(szRegBuffer) - 1);  // -1 to leave space for a 2nd terminator.
		// Double-terminate:
		size_t length = _tcslen(szRegBuffer);
		szRegBuffer[length + 1] = '\0';

		// Remove any final newline the user may have provided since we don't want the length
		// to include it when calling RegSetValueEx() -- it would be too large by 1:
		if (length > 0 && szRegBuffer[length - 1] == '\n')
			szRegBuffer[--length] = '\0';

		// Replace the script's delimiter char with the zero-delimiter needed by RegSetValueEx():
		for (LPTSTR cp = szRegBuffer; *cp; ++cp)
			if (*cp == '\n')
				*cp = '\0';

		if (RegSetValueEx(hRegKey, aValueName, 0, REG_MULTI_SZ, (CONST BYTE *)szRegBuffer
			, (DWORD)(length ? length + 2 : 0) * sizeof(TCHAR)) == ERROR_SUCCESS)
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		RegCloseKey(hRegKey);
		return OK;
	}

	case REG_DWORD:
		if (*aValue)
			dwBuf = ATOU(aValue);  // Changed to ATOU() for v1.0.24 so that hex values are supported.
		else // Default to 0 when blank.
			dwBuf = 0;
		if (RegSetValueEx(hRegKey, aValueName, 0, REG_DWORD, (CONST BYTE *)&dwBuf, sizeof(dwBuf) ) == ERROR_SUCCESS)
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		RegCloseKey(hRegKey);
		return OK;

	case REG_BINARY:
	{
		int nLen = (int)_tcslen(aValue);

		// Stringlength must be a multiple of 2 
		if (nLen % 2)
		{
			RegCloseKey(hRegKey);
			return OK;  // Let ErrorLevel tell the story.
		}

		// Really crappy hex conversion
		int j = 0, i = 0, nVal, nMult;
		LPBYTE pRegBuffer = (LPBYTE) szRegBuffer;
		while (i < nLen && j < sizeof(szRegBuffer))
		{
			nVal = 0;
			for (nMult = 16; nMult >= 0; nMult = nMult - 15)
			{
				if (aValue[i] >= '0' && aValue[i] <= '9')
					nVal += (aValue[i] - '0') * nMult;
				else if (aValue[i] >= 'A' && aValue[i] <= 'F')
					nVal += (((aValue[i] - 'A'))+10) * nMult;
				else if (aValue[i] >= 'a' && aValue[i] <= 'f')
					nVal += (((aValue[i] - 'a'))+10) * nMult;
				else
				{
					RegCloseKey(hRegKey);
					return OK;  // Let ErrorLevel tell the story.
				}
				++i;
			}
			pRegBuffer[j++] = (char)nVal;
		}

		if (RegSetValueEx(hRegKey, aValueName, 0, REG_BINARY, pRegBuffer, (DWORD)j) == ERROR_SUCCESS)
			g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
		// else keep the default failure value for ErrorLevel
	
		RegCloseKey(hRegKey);
		return OK;
	}
	} // switch()

	// If we reached here then the requested type was unknown/unsupported.
	// Let ErrorLevel tell the story.
	RegCloseKey(hRegKey);
	return OK;
} // RegWrite()



bool Line::RegRemoveSubkeys(HKEY hRegKey)
{
	// Removes all subkeys to the given key.  Will not touch the given key.
	TCHAR Name[256];
	DWORD dwNameSize;
	FILETIME ftLastWrite;
	HKEY hSubKey;
	bool Success;

	for (;;) 
	{ // infinite loop 
		dwNameSize=255;
		if (RegEnumKeyEx(hRegKey, 0, Name, &dwNameSize, NULL, NULL, NULL, &ftLastWrite) == ERROR_NO_MORE_ITEMS)
			break;
		if (RegOpenKeyEx(hRegKey, Name, 0, KEY_READ, &hSubKey) != ERROR_SUCCESS)
			return false;
		
		Success=RegRemoveSubkeys(hSubKey);
		RegCloseKey(hSubKey);
		if (!Success)
			return false;
		else if (RegDeleteKey(hRegKey, Name) != ERROR_SUCCESS)
			return false;
	}
	return true;
}



ResultType Line::RegDelete(HKEY aRootKey, LPTSTR aRegSubkey, LPTSTR aValueName)
{
	g_ErrorLevel->Assign(ERRORLEVEL_ERROR); // Set default ErrorLevel.

	// Fix for v1.0.48: Don't remove the entire key if it's a root key!  According to MSDN,
	// the root key would be opened by RegOpenKeyEx() further below whenever aRegSubkey is NULL
	// or an empty string. aValueName is also checked to preserve the ability to delete a value
	// that exists directly under a root key.
	if (   !aRootKey
		|| (!aRegSubkey || !*aRegSubkey) && (!aValueName || !*aValueName)   ) // See comment above.
		return OK;  // Let ErrorLevel tell the story.

	// Open the key we want
	HKEY hRegKey;
	if (RegOpenKeyEx(aRootKey, aRegSubkey, 0, KEY_READ | KEY_WRITE, &hRegKey) != ERROR_SUCCESS)
		return OK;  // Let ErrorLevel tell the story.

	if (!aValueName || !*aValueName)
	{
		// Remove the entire Key
		bool success = RegRemoveSubkeys(hRegKey); // Delete any subitems within the key.
		RegCloseKey(hRegKey); // Close parent key.  Not sure if this needs to be done only after the above.
		if (!success)
			return OK;  // Let ErrorLevel tell the story.
		if (RegDeleteKey(aRootKey, aRegSubkey) != ERROR_SUCCESS) 
			return OK;  // Let ErrorLevel tell the story.
	}
	else
	{
		// Remove Value.  The special phrase "ahk_default" indicates that the key's default
		// value (displayed as "(Default)" by RegEdit) should be deleted.  This is done to
		// distinguish a blank (which deletes the entire subkey) from the default item.
		LONG lRes = RegDeleteValue(hRegKey, _tcsicmp(aValueName, _T("ahk_default")) ? aValueName : _T(""));
		RegCloseKey(hRegKey);
		if (lRes != ERROR_SUCCESS)
			return OK;  // Let ErrorLevel tell the story.
	}

	return g_ErrorLevel->Assign(ERRORLEVEL_NONE); // Indicate success.
} // RegDelete()
