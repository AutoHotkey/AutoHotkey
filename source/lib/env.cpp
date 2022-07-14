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



///////////////////////////
// Environment Functions //
///////////////////////////


BIF_DECL(BIF_Env)
// Value := EnvGet(EnvVarName)
// EnvSet(EnvVarName, Value)
{
	LPTSTR aEnvVarName = ParamIndexToString(0, _f_number_buf);
	// Don't use a size greater than 32767 because that will cause it to fail on Win95 (tested by Robert Yalkin).
	// According to MSDN, 32767 is exactly large enough to handle the largest variable plus its zero terminator.
	// Update: In practice, at least on Windows 7, the limit only applies to the ANSI functions.
	TCHAR buf[32767];

	if (_f_callee_id == FID_EnvSet)
	{
		// MSDN: "If [the 2nd] parameter is NULL, the variable is deleted from the current process's environment."
		// No checking is currently done to ensure that aValue isn't longer than 32K, since testing shows that
		// larger values are supported by the Unicode functions, at least in some OSes.  If not supported,
		// SetEnvironmentVariable() should return 0 (fail) anyway.
		// Note: It seems that env variable names can contain spaces and other symbols, so it's best not to
		// validate aEnvVarName the same way we validate script variables (i.e. just let the return value
		// determine whether there's an error).
		LPTSTR aValue = ParamIndexIsOmitted(1) ? NULL : ParamIndexToString(1, buf);
		if (!SetEnvironmentVariable(aEnvVarName, aValue))
			_f_throw_win32();
		_f_return_empty;
	}

	// GetEnvironmentVariable() could be called twice, the first time to get the actual size.  But that would
	// probably perform worse since GetEnvironmentVariable() is a very slow function, so it seems best to fetch
	// it into a large buffer then just copy it to dest-var.
	DWORD length = GetEnvironmentVariable(aEnvVarName, buf, _countof(buf));
	if (length >= _countof(buf))
	{
		// In this case, length indicates the required buffer size, and the contents of the buffer are undefined.
		// Since our buffer is 32767 characters, the var apparently exceeds the documented limit, as can happen
		// if the var was set with the Unicode API.
		if (!TokenSetResult(aResultToken, NULL, length - 1))
			return;
		length = GetEnvironmentVariable(aEnvVarName, aResultToken.marker, length);
		if (!length)
			*aResultToken.marker = '\0'; // Ensure var is null-terminated.
		aResultToken.symbol = SYM_STRING;
		aResultToken.marker_length = length; // Update length in case it changed.
		return;
	}
	_f_return(length ? buf : _T(""), length);
}



BIF_DECL(BIF_SysGet)
{
	_f_return(GetSystemMetrics(ParamIndexToInt(0)));
}



BIF_DECL(BIF_MonitorGet)
{
	MonitorInfoPackage mip = {0};  // Improves maintainability to initialize unconditionally, here.
	mip.monitor_info_ex.cbSize = sizeof(MONITORINFOEX); // Also improves maintainability.

	BuiltInFunctionID cmd = _f_callee_id;
	switch (cmd)
	{
	// For the next few cases, I'm not sure if it is possible to have zero monitors.  Obviously it's possible
	// to not have a monitor turned on or not connected at all.  But it seems likely that these various API
	// functions will provide a "default monitor" in the absence of a physical monitor connected to the
	// system.  To be safe, all of the below will assume that zero is possible, at least on some OSes or
	// under some conditions.  However, on Win95/NT, "1" is assumed since there is probably no way to tell
	// for sure if there are zero monitors except via GetSystemMetrics(SM_CMONITORS), which is a different
	// animal as described below.
	case FID_MonitorGetCount:
		// Don't use GetSystemMetrics(SM_CMONITORS) because of this:
		// MSDN: "GetSystemMetrics(SM_CMONITORS) counts only display monitors. This is different from
		// EnumDisplayMonitors, which enumerates display monitors and also non-display pseudo-monitors."
		mip.monitor_number_to_find = COUNT_ALL_MONITORS;
		EnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		_f_return_i(mip.count); // Will return zero if the API ever returns a legitimate zero.

	// Even if the first monitor to be retrieved by the EnumProc is always the primary (which is doubtful
	// since there's no mention of this in the MSDN docs) it seems best to have this sub-cmd in case that
	// policy ever changes:
	case FID_MonitorGetPrimary:
		// The mip struct's values have already initialized correctly for the below:
		EnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		_f_return_i(mip.count); // Will return zero if the API ever returns a legitimate zero.

	case FID_MonitorGet:
	case FID_MonitorGetWorkArea:
	// Params: N, Left, Top, Right, Bottom
	{
		mip.monitor_number_to_find = ParamIndexToOptionalInt(0, 0);  // If this returns 0, it will default to the primary monitor.
		EnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		if (!mip.count || (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count))
			break;
		// Otherwise:
		LONG *monitor_rect = (LONG *)((cmd == FID_MonitorGetWorkArea) ? &mip.monitor_info_ex.rcWork : &mip.monitor_info_ex.rcMonitor);
		for (int i = 1; i <= 4; ++i) // Params: N (0), Left (1), Top, Right, Bottom.
			if (Var *var = ParamIndexToOutputVar(i))
				var->Assign(monitor_rect[i-1]);
		_f_return_i(mip.count); // Return the monitor number.
	}

	case FID_MonitorGetName: // Param: N
		mip.monitor_number_to_find = ParamIndexToOptionalInt(0, 0);  // If this returns 0, it will default to the primary monitor.
		EnumDisplayMonitors(NULL, NULL, EnumMonitorProc, (LPARAM)&mip);
		if (!mip.count || (mip.monitor_number_to_find && mip.monitor_number_to_find != mip.count))
			break;
		_f_return(mip.monitor_info_ex.szDevice);
	} // switch()
	
	// Since above didn't return, an error was detected.
	if (!mip.count) // Might be virtually impossible.
		_f_throw_win32();
	_f_throw_param(0);
}



BOOL CALLBACK EnumMonitorProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM lParam)
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
