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
#include "defines.h"
#include "globaldata.h"
#include "application.h"
#include <Psapi.h>
#include "script_func_impl.h"



bif_impl UINT ProcessExist(optl<StrArg> aProcess)
{
	// Return the discovered PID or zero if none.
	return aProcess.has_value() ? ProcessExist(aProcess.value()) : GetCurrentProcessId();
}



bif_impl UINT ProcessGetParent(optl<StrArg> aProcess)
{
	TCHAR buf[MAX_INTEGER_SIZE];
	return ProcessExist(aProcess.has_value() ? aProcess.value() : ITOA(GetCurrentProcessId(), buf), true);
}



bif_impl FResult ProcessClose(StrArg aProcess, UINT &aRetVal)
{
	aRetVal = 0; // Set default in case of failure.
	if (auto pid = ProcessExist(aProcess))  // Assign
	{
		if (auto hProcess = OpenProcess(PROCESS_TERMINATE, FALSE, pid))
		{
			if (TerminateProcess(hProcess, 0))
				aRetVal = pid; // Indicate success.
			CloseHandle(hProcess);
		}
	}
	return OK;
}



static FResult ProcessGetPathName(optl<StrArg> aProcess, StrRet &aRetVal, bool aGetNameOnly)
{
	auto pid = aProcess.has_value() ? ProcessExist(aProcess.value()) : GetCurrentProcessId();
	if (!pid)
		return FError(ERR_NO_PROCESS, nullptr, ErrorPrototype::Target);
	TCHAR process_name[MAX_PATH];
	if (!GetProcessName(pid, process_name, _countof(process_name), aGetNameOnly))
		return FR_E_WIN32;
	return aRetVal.Copy(process_name) ? OK : FR_E_OUTOFMEM;
}

bif_impl FResult ProcessGetName(optl<StrArg> aProcess, StrRet &aRetVal)
{
	return ProcessGetPathName(aProcess, aRetVal, true);
}

bif_impl FResult ProcessGetPath(optl<StrArg> aProcess, StrRet &aRetVal)
{
	return ProcessGetPathName(aProcess, aRetVal, false);
}



void MsgSleepWithListLines(int aSleepDuration, Line *waiting_line, DWORD start_time);

static FResult ProcessWait(StrArg aProcess, optl<double> aTimeout, UINT &aRetVal, bool aWaitClose)
{
	DWORD pid;
	// This section is similar to that used for WINWAIT and RUNWAIT:
	bool wait_indefinitely;
	int sleep_duration;
	DWORD start_time = GetTickCount();
	Line *waiting_line = g_script.mCurrLine;
	if (aTimeout.has_value()) // The param containing the timeout value was specified.
	{
		wait_indefinitely = false;
		sleep_duration = (int)(aTimeout.value_or(0) * 1000); // Can be zero.
	}
	else
	{
		wait_indefinitely = true;
		sleep_duration = 0; // Just to catch any bugs.
	}
	for (;;)
	{ // Always do the first iteration so that at least one check is done.
		pid = ProcessExist(aProcess);
		if ((!aWaitClose) == (pid != 0)) // i.e. condition of this cmd is satisfied.
		{
			// For WaitClose: Since PID cannot always be determined (i.e. if process never existed,
			// there was no need to wait for it to close), for consistency, return 0 on success.
			aRetVal = pid;
			return OK;
		}
		// Must cast to int or any negative result will be lost due to DWORD type:
		if (wait_indefinitely || (int)(sleep_duration - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
		{
			MsgSleepWithListLines(100, waiting_line, start_time);  // For performance reasons, don't check as often as the WinWait family does.
		}
		else // Done waiting.
		{
			// Return 0 if ProcessWait times out; or the PID of the process that still exists
			// if ProcessWaitClose times out.
			aRetVal = pid;
			return OK;
		}
	} // for()
}

bif_impl FResult ProcessWait(StrArg aProcess, optl<double> aTimeout, UINT &aRetVal)
{
	return ProcessWait(aProcess, aTimeout, aRetVal, false);
}

bif_impl FResult ProcessWaitClose(StrArg aProcess, optl<double> aTimeout, UINT &aRetVal)
{
	return ProcessWait(aProcess, aTimeout, aRetVal, true);
}



bif_impl FResult ProcessSetPriority(StrArg aPriority, optl<StrArg> aProcess, UINT &aRetVal)
{
	DWORD priority;
	switch (_totupper(*aPriority))
	{
	case 'L': priority = IDLE_PRIORITY_CLASS; break;
	case 'B': priority = BELOW_NORMAL_PRIORITY_CLASS; break;
	case 'N': priority = NORMAL_PRIORITY_CLASS; break;
	case 'A': priority = ABOVE_NORMAL_PRIORITY_CLASS; break;
	case 'H': priority = HIGH_PRIORITY_CLASS; break;
	case 'R': priority = REALTIME_PRIORITY_CLASS; break;
	default:
		// Since above didn't break, aPriority was invalid.
		return FR_E_ARG(0);
	}

	DWORD pid = aProcess.has_nonempty_value() ? ProcessExist(aProcess.value()) : GetCurrentProcessId();
	if (!pid)
		return FError(ERR_NO_PROCESS, nullptr, ErrorPrototype::Target);

	HANDLE hProcess = OpenProcess(PROCESS_SET_INFORMATION, FALSE, pid);
	if (!hProcess)
		return FR_E_WIN32;

	DWORD error = SetPriorityClass(hProcess, priority) ? NOERROR : GetLastError();
	CloseHandle(hProcess);
	if (error)
		return FR_E_WIN32(error);
	aRetVal = pid;
	return OK;
}



bif_impl FResult Run(StrArg aTarget, optl<StrArg> aWorkingDir, optl<StrArg> aOptions, ResultToken *aOutPID)
{
	HANDLE hprocess;
	auto result = g_script.ActionExec(aTarget, nullptr, aWorkingDir.value_or_null(), true
		, aOptions.value_or_null(), &hprocess, true, true);
	if (aOutPID && hprocess)
		aOutPID->SetValue((UINT)GetProcessId(hprocess));
	return result ? OK : FR_FAIL;
}



DWORD GetProcessName(DWORD aProcessID, LPTSTR aBuf, DWORD aBufSize, bool aGetNameOnly)
{
	*aBuf = '\0'; // Set default.
	HANDLE hproc;
	// Windows XP/2003 would require PROCESS_QUERY_INFORMATION, but those OSes are not supported.
	if (  !(hproc = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, FALSE, aProcessID))  )
		return 0;

	// Benchmarks showed that attempting GetModuleBaseName/GetModuleFileNameEx
	// first did not help performance.  Also, QueryFullProcessImageName appeared
	// to be slower than the following.
	DWORD buf_length = GetProcessImageFileName(hproc, aBuf, aBufSize);
	if (buf_length)
	{
		LPTSTR cp;
		if (aGetNameOnly)
		{
			// Convert full path to just name.
			cp = _tcsrchr(aBuf, '\\');
			if (cp)
				tmemmove(aBuf, cp + 1, _tcslen(cp)); // Includes the null terminator.
		}
		else
		{
			// Convert device path to logical path.
			TCHAR device_path[MAX_PATH];
			TCHAR letter[3];
			letter[1] = ':';
			letter[2] = '\0';
			// For simplicity and because GetLogicalDriveStrings does not exist on Win2k, it is not used.
			for (*letter = 'A'; *letter <= 'Z'; ++(*letter))
			{
				DWORD device_path_length = QueryDosDevice(letter, device_path, _countof(device_path));
				if (device_path_length > 2) // Includes two null terminators.
				{
					device_path_length -= 2;
					if (!_tcsncmp(device_path, aBuf, device_path_length)
						&& aBuf[device_path_length] == '\\') // Relies on short-circuit evaluation.
					{
						// Copy drive letter:
						aBuf[0] = letter[0];
						aBuf[1] = letter[1];
						// Contract path to remove remainder of device name.
						tmemmove(aBuf + 2, aBuf + device_path_length, buf_length - device_path_length + 1);
						buf_length -= device_path_length - 2;
						break;
					}
				}
			}
		}
	}

	CloseHandle(hproc);
	return buf_length;
}
