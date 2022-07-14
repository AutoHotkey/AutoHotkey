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

#include "stdafx.h" // pre-compiled headers
#include <winioctl.h> // For PREVENT_MEDIA_REMOVAL and CD lock/unlock.
#include "qmath.h" // Used by Transform() [math.h incurs 2k larger code size just for ceil() & floor()]
#include "script.h"
#include "window.h" // for IF_USE_FOREGROUND_WINDOW
#include "application.h" // for MsgSleep()
#include "resources/resource.h"  // For InputBox.
#include "TextIO.h"
#include <Psapi.h> // for GetModuleBaseName.
#include "shlwapi.h" // StrCmpLogicalW

#include <mmdeviceapi.h> // for SoundSet/SoundGet.
#include <endpointvolume.h> // for SoundSet/SoundGet.
#include <functiondiscoverykeys.h>

#define PCRE_STATIC             // For RegEx. PCRE_STATIC tells PCRE to declare its functions for normal, static
#include "lib_pcre/pcre/pcre.h" // linkage rather than as functions inside an external DLL.

#include "script_func_impl.h"



///////////////////////
// Process Functions //
///////////////////////


BIF_DECL(BIF_Process)
{
	_f_param_string_opt(aProcess, 0);

	HANDLE hProcess;
	DWORD pid;
	
	BuiltInFunctionID process_cmd = _f_callee_id;
	switch (process_cmd)
	{
	case FID_ProcessExist:
		// Return the discovered PID or zero if none.
		_f_return_i(*aProcess ? ProcessExist(aProcess) : GetCurrentProcessId());

	case FID_ProcessClose:
		_f_set_retval_i(0); // Set default.
		if (pid = ProcessExist(aProcess))  // Assign
		{
			if (hProcess = OpenProcess(PROCESS_TERMINATE, FALSE, pid))
			{
				if (TerminateProcess(hProcess, 0))
					_f_set_retval_i(pid); // Indicate success.
				CloseHandle(hProcess);
			}
		}
		// Since above didn't return, yield a PID of 0 to indicate failure.
		_f_return_retval;

	case FID_ProcessWait:
	case FID_ProcessWaitClose:
	{
		// This section is similar to that used for WINWAIT and RUNWAIT:
		bool wait_indefinitely;
		int sleep_duration;
		DWORD start_time;
		if (aParamCount > 1) // The param containing the timeout value was specified.
		{
			if (!ParamIndexIsNumeric(1))
				_f_throw_param(1, _T("Number"));
			wait_indefinitely = false;
			sleep_duration = (int)(TokenToDouble(*aParam[1]) * 1000); // Can be zero.
			start_time = GetTickCount();
		}
		else
		{
			wait_indefinitely = true;
			sleep_duration = 0; // Just to catch any bugs.
		}
		for (;;)
		{ // Always do the first iteration so that at least one check is done.
			pid = ProcessExist(aProcess);
			if ((process_cmd == FID_ProcessWait) == (pid != 0)) // i.e. condition of this cmd is satisfied.
			{
				// For WaitClose: Since PID cannot always be determined (i.e. if process never existed,
				// there was no need to wait for it to close), for consistency, return 0 on success.
				_f_return_i(pid);
			}
			// Must cast to int or any negative result will be lost due to DWORD type:
			if (wait_indefinitely || (int)(sleep_duration - (GetTickCount() - start_time)) > SLEEP_INTERVAL_HALF)
			{
				MsgSleep(100);  // For performance reasons, don't check as often as the WinWait family does.
			}
			else // Done waiting.
			{
				// Return 0 if "Process Wait" times out; or the PID of the process that still exists
				// if "Process WaitClose" times out.
				_f_return_i(pid);
			}
		} // for()

	case FID_ProcessGetName:
	case FID_ProcessGetPath:
		pid = *aProcess ? ProcessExist(aProcess) : GetCurrentProcessId();
		if (!pid)
			_f_throw(ERR_NO_PROCESS, ErrorPrototype::Target);
		TCHAR process_name[MAX_PATH];
		if (!GetProcessName(pid, process_name, _countof(process_name), process_cmd == FID_ProcessGetName))
			_f_throw_win32();
		_f_return(process_name);
	} // case
	} // switch()
}



BIF_DECL(BIF_ProcessSetPriority)
{
	_f_param_string_opt(aPriority, 0);
	_f_param_string_opt(aProcess, 1);

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
		_f_throw_param(0);
	}

	DWORD pid = *aProcess ? ProcessExist(aProcess) : GetCurrentProcessId();
	if (!pid)
		_f_throw(ERR_NO_PROCESS, ErrorPrototype::Target);

	HANDLE hProcess = OpenProcess(PROCESS_SET_INFORMATION, FALSE, pid);
	if (!hProcess)
		_f_throw_win32();

	DWORD error = SetPriorityClass(hProcess, priority) ? NOERROR : GetLastError();
	CloseHandle(hProcess);
	if (error)
		_f_throw_win32(error);
	_f_return_i(pid);
}



BIF_DECL(BIF_Run)
{
	_f_param_string(target, 0);
	_f_param_string_opt(working_dir, 1);
	_f_param_string_opt(options, 2);
	Var *output_var_pid = ParamIndexToOutputVar(3);
	if (!g_script.ActionExec(target, nullptr, working_dir, true, options, nullptr, true, true, output_var_pid))
		_f_return_FAIL;
	_f_return_empty;
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
