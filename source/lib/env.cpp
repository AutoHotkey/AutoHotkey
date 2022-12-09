/*
AutoHotkey

Copyright 2003-2009 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#include "stdafx.h"
#include "script.h"
#include "window.h" // for MonitorInfoPackage
#include "script_func_impl.h"
#include "abi.h"



#pragma region Environment Variables


bif_impl FResult EnvGet(StrArg aEnvVarName, StrRet &aRetVal)
{
	// According to MSDN, 32767 is exactly large enough to handle the largest variable plus its zero terminator.
	// Update: In practice, at least on Windows 7, the limit only applies to the ANSI functions.
	TCHAR stack_buf[32767];

	// GetEnvironmentVariable() could be called twice, the first time to get the actual size.  But calling it once
	// and then copying the string performs better, except in the rare cases where the initial buffer is too small.
	DWORD length = GetEnvironmentVariable(aEnvVarName, stack_buf, _countof(stack_buf));
	//_tcscpy(stack_buf, _T("Hello, world"));
	//DWORD length = (DWORD)_tcslen(_T("Hello, world"));
	if (!length) // Generally empty or undefined.
	{
		//aRetVal->SetEmpty();
		return OK;
	}
	LPTSTR buf = aRetVal.Alloc(length);
	if (!buf)
		return FR_E_OUTOFMEM;
	if (length >= _countof(stack_buf))
	{
		// In this case, length indicates the required buffer size, and the contents of the buffer are undefined.
		// Since our buffer is 32767 characters, the var apparently exceeds the documented limit, as can happen
		// if the var was set with the Unicode API.
		auto size = length;
		length = GetEnvironmentVariable(aEnvVarName, buf, size);
		if (!length || length >= size) // Check against size in case the value could have changed since the first call.
		{
			*buf = '\0'; // Ensure var is null-terminated.
			length = 0;
		}
	}
	else
		tmemcpy(buf, stack_buf, (size_t)length + 1);
		//tmemcpy(buf, _T("Hello, world"), (size_t)length + 1);
	aRetVal.SetLength(length);
	return OK;
}


bif_impl FResult EnvSet(StrArg aName, optl<StrArg> aValue)
{
	// MSDN: "If [the 2nd] parameter is NULL, the variable is deleted from the current process's environment."
	// No checking is currently done to ensure that aValue isn't longer than 32K, since testing shows that
	// larger values are supported by the Unicode functions, at least in some OSes.  If not supported,
	// SetEnvironmentVariable() should return 0 (fail) anyway.
	// Note: It seems that env variable names can contain spaces and other symbols, so it's best not to
	// validate aEnvVarName the same way we validate script variables (i.e. just let the return value
	// determine whether there's an error).
	return SetEnvironmentVariable(aName, aValue.value_or_null()) ? OK : FR_E_WIN32;
}


#pragma endregion



bif_impl int SysGet(int aIndex)
{
	return GetSystemMetrics(aIndex);
}



#pragma region MonitorGet and related


static BOOL CALLBACK EnumMonitorProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM lParam)
{
	MonitorInfoPackage &mip = *(MonitorInfoPackage *)lParam;  // For performance and convenience.
	if (mip.monitor_number_to_find == COUNT_ALL_MONITORS)
	{
		++mip.count;
		return TRUE;  // Enumerate all monitors so that they can be counted.
	}
	if (!GetMonitorInfo(hMonitor, &mip.monitor_info_ex)) // Failed.  Probably very rare.
		return FALSE; // Due to the complexity of needing to stop at the correct monitor number, do not continue.
		// In the unlikely event that the above fails when the caller wanted us to find the primary
		// monitor, the caller will think the primary is the previously found monitor (if any).
		// This is just documented here as a known limitation since this combination of circumstances
		// is probably impossible.
	++mip.count; // So that caller can detect failure, increment only now that failure conditions have been checked.
	if (mip.monitor_number_to_find) // Caller gave a specific monitor number, so don't search for the primary monitor.
	{
		if (mip.count == mip.monitor_number_to_find) // Since the desired monitor has been found, must not continue.
			return FALSE;
	}
	else // Caller wants the primary monitor found.
		// MSDN docs are unclear that MONITORINFOF_PRIMARY is a bitwise value, but the name "dwFlags" implies so:
		if (mip.monitor_info_ex.dwFlags & MONITORINFOF_PRIMARY)
			return FALSE;  // Primary has been found and "count" contains its number. Must not continue the enumeration.
			// Above assumes that it is impossible to not have a primary monitor in a system that has at least
			// one monitor.  MSDN certainly implies this through multiple references to the primary monitor.
	// Otherwise, continue the enumeration:
	return TRUE;
}


static void EnumForMonitorGet(MonitorInfoPackage &mip, int aIndex)
{
	mip.monitor_info_ex.cbSize = sizeof(MONITORINFOEX);
	mip.monitor_number_to_find = aIndex; // If 0, it will default to the primary monitor.
	EnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
}


// For the next few BIFs, I'm not sure if it is possible to have zero monitors.  Obviously it's possible
// to not have a monitor turned on or not connected at all.  But it seems likely that these various API
// functions will provide a "default monitor" in the absence of a physical monitor connected to the
// system.  To be safe, all of the below will assume that zero is possible, at least on some OSes or
// under some conditions.


bif_impl int MonitorGetCount()
{
	// Don't use GetSystemMetrics(SM_CMONITORS) because of this:
	// MSDN: "GetSystemMetrics(SM_CMONITORS) counts only display monitors. This is different from
	// EnumDisplayMonitors, which enumerates display monitors and also non-display pseudo-monitors."
	MonitorInfoPackage mip = {0};
	EnumForMonitorGet(mip, COUNT_ALL_MONITORS);
	return mip.count;
}


bif_impl int MonitorGetPrimary()
{
	// Even if the first monitor to be retrieved by the EnumProc is always the primary (which is doubtful
	// since there's no mention of this in the MS docs) it seems best to have this function in case that
	// policy ever changes.
	MonitorInfoPackage mip = {0};
	EnumForMonitorGet(mip, 0);
	return mip.count;
}


static FResult MonitorGet(optl<int> aIndex, int *aLeft, int *aTop, int *aRight, int *aBottom, int &aRetVal, bool aWorkArea)
{
	MonitorInfoPackage mip = {0};
	EnumForMonitorGet(mip, aIndex.value_or(0));
	if (!mip.count) // Might be virtually impossible.
		return FR_E_WIN32;
	if (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count)
		return FR_E_ARG(0);
	// Otherwise:
	auto &monitor_rect = aWorkArea ? mip.monitor_info_ex.rcWork : mip.monitor_info_ex.rcMonitor;
	if (aLeft) *aLeft = monitor_rect.left;
	if (aTop) *aTop = monitor_rect.top;
	if (aRight) *aRight = monitor_rect.right;
	if (aBottom) *aBottom = monitor_rect.bottom;
	aRetVal = mip.count;
	return OK;
}


bif_impl FResult MonitorGet(optl<int> aIndex, int *aLeft, int *aTop, int *aRight, int *aBottom, int &aRetVal)
{
	return MonitorGet(aIndex, aLeft, aTop, aRight, aBottom, aRetVal, false);
}


bif_impl FResult MonitorGetWorkArea(optl<int> aIndex, int *aLeft, int *aTop, int *aRight, int *aBottom, int &aRetVal)
{
	return MonitorGet(aIndex, aLeft, aTop, aRight, aBottom, aRetVal, true);
}


bif_impl FResult MonitorGetName(optl<int> aIndex, StrRet &aRetVal)
{
	MonitorInfoPackage mip = {0};
	EnumForMonitorGet(mip, aIndex.value_or(0));
	if (!mip.count) // Might be virtually impossible.
		return FR_E_WIN32;
	if (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count)
		return FR_E_ARG(0);
	// Using _countof() vs. _tcslen() may help both code size and performance,
	// due to inlining of Alloc() and its use of mCallerBuf for small strings:
	LPTSTR buf = aRetVal.Alloc(_countof(mip.monitor_info_ex.szDevice));
	if (!buf)
		return FR_E_OUTOFMEM;
	_tcscpy(buf, mip.monitor_info_ex.szDevice);
	return OK;
}


#pragma endregion
