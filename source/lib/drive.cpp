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
#include <winioctl.h> // For PREVENT_MEDIA_REMOVAL and CD lock/unlock.
#include "script.h"
#include "script_func_impl.h"



void DriveSpace(ResultToken &aResultToken, LPTSTR aPath, bool aGetFreeSpace)
// Because of NTFS's ability to mount volumes into a directory, a path might not necessarily
// have the same amount of free space as its root drive.  However, I'm not sure if this
// method here actually takes that into account.
{
	ASSERT(aPath && *aPath);

	TCHAR buf[MAX_PATH]; // MAX_PATH vs T_MAX_PATH because testing shows it doesn't support long paths even with \\?\.
	tcslcpy(buf, aPath, _countof(buf));
	size_t length = _tcslen(buf);
	if (buf[length - 1] != '\\' // Trailing backslash is absent,
		&& length + 1 < _countof(buf)) // and there's room to fix it.
	{
		buf[length++] = '\\';
		buf[length] = '\0';
	}
	//else it should still work unless this is a UNC path.

	// MSDN: "The GetDiskFreeSpaceEx function returns correct values for all volumes, including those
	// that are greater than 2 gigabytes."
	__int64 free_space;
	ULARGE_INTEGER total, free, used;
	if (!GetDiskFreeSpaceEx(buf, &free, &total, &used))
		_f_throw_win32();
	// Casting this way allows sizes of up to 2,097,152 gigabytes:
	free_space = (__int64)((unsigned __int64)(aGetFreeSpace ? free.QuadPart : total.QuadPart)
		/ (1024*1024));

	_f_return(free_space);
}



ResultType DriveLock(TCHAR aDriveLetter, bool aLockIt); // Forward declaration for BIF_Drive.

BIF_DECL(BIF_Drive)
{
	BuiltInFunctionID drive_cmd = _f_callee_id;

	LPTSTR aValue = ParamIndexToOptionalString(0, _f_number_buf);

	bool successful = false;
	TCHAR path[MAX_PATH]; // MAX_PATH vs. T_MAX_PATH because SetVolumeLabel() can't seem to make use of long paths.
	size_t path_length;

	// Notes about the below macro:
	// - It adds a backslash to the contents of the path variable because certain API calls or OS versions
	//   might require it.
	// - It is used by both Drive() and DriveGet().
	// - Leave space for the backslash in case its needed.
	#define DRIVE_SET_PATH \
		tcslcpy(path, aValue, _countof(path) - 1);\
		path_length = _tcslen(path);\
		if (path_length && path[path_length - 1] != '\\')\
			path[path_length++] = '\\';

	switch(drive_cmd)
	{
	case FID_DriveLock:
	case FID_DriveUnlock:
		successful = DriveLock(*aValue, drive_cmd == FID_DriveLock);
		break;

	case FID_DriveEject:
	case FID_DriveRetract:
	{
		// Don't do DRIVE_SET_PATH in this case since trailing backslash is not wanted.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash, which might prevent DriveGetType() from working on some OSes.
		// 2) Eject (and more rarely, retract) works on some other drive types.
		// 3) CreateFile or DeviceIoControl will simply fail or have no effect if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM)
		//	_f_throw(ERR_FAILED);
		BOOL retract = (drive_cmd == FID_DriveRetract);
		TCHAR path[] { '\\', '\\', '.', '\\', 0, ':', '\0', '\0' };
		if (*aValue)
		{
			// Testing showed that a Volume GUID of the form \\?\Volume{...} will work even when
			// the drive isn't mapped to a drive letter.
			if (cisalpha(aValue[0]) && (!aValue[1] || aValue[1] == ':' && (!aValue[2] || aValue[2] == '\\' && !aValue[3])))
			{
				path[4] = aValue[0];
				aValue = path;
			}
		}
		else // When drive is omitted, operate upon the first CD/DVD drive.
		{
			path[6] = '\\'; // GetDriveType technically requires a slash, although it may work without.
			// Testing with mciSendString() while changing or removing drive letters showed that
			// its "default" drive is really just the first drive found in alphabetical order,
			// which is also the most obvious/intuitive choice.
			for (TCHAR c = 'A'; ; ++c)
			{
				path[4] = c;
				if (GetDriveType(path) == DRIVE_CDROM)
					break;
				if (c == 'Z')
					_f_throw(ERR_FAILED); // No CD/DVD drive found with a drive letter.  
			}
			path[6] = '\0'; // Remove the trailing slash for CreateFile to open the volume.
			aValue = path;
		}
		// Testing indicates neither this method nor the MCI method work with mapped drives or UNC paths.
		// That makes sense when one considers that the following opens the *volume*, whereas a network
		// share would correspond to a directory; i.e. this needs "\\.\D:" and not "\\.\D:\".
		HANDLE hVol = CreateFile(aValue, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
		if (hVol == INVALID_HANDLE_VALUE)
			break;
		DWORD unused;
		successful = DeviceIoControl(hVol, retract ? IOCTL_STORAGE_LOAD_MEDIA : IOCTL_STORAGE_EJECT_MEDIA, NULL, 0, NULL, 0, &unused, NULL);
		CloseHandle(hVol);
		break;
	}

	case FID_DriveSetLabel: // Note that it is possible and allowed for the new label to be blank.
		DRIVE_SET_PATH
		successful = SetVolumeLabel(path, ParamIndexToOptionalString(1, _f_number_buf)); // _f_number_buf can potentially be used by both parameters, since aValue is copied into path above.
		break;

	} // switch()

	if (!successful)
		_f_throw_win32();
	_f_return_empty;
}



ResultType DriveLock(TCHAR aDriveLetter, bool aLockIt)
{
	HANDLE hdevice;
	DWORD unused;
	BOOL result;
	TCHAR filename[64];
	_stprintf(filename, _T("\\\\.\\%c:"), aDriveLetter);
	// FILE_READ_ATTRIBUTES is not enough; it yields "Access Denied" error.  So apparently all or part
	// of the sub-attributes in GENERIC_READ are needed.  An MSDN example implies that GENERIC_WRITE is
	// only needed for GetDriveType() == DRIVE_REMOVABLE drives, and maybe not even those when all we
	// want to do is lock/unlock the drive (that example did quite a bit more).  In any case, research
	// indicates that all CD/DVD drives are ever considered DRIVE_CDROM, not DRIVE_REMOVABLE.
	// Due to this and the unlikelihood that GENERIC_WRITE is ever needed anyway, GetDriveType() is
	// not called for the purpose of conditionally adding the GENERIC_WRITE attribute.
	hdevice = CreateFile(filename, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
	if (hdevice == INVALID_HANDLE_VALUE)
		return FAIL;
	PREVENT_MEDIA_REMOVAL pmr;
	pmr.PreventMediaRemoval = aLockIt;
	result = DeviceIoControl(hdevice, IOCTL_STORAGE_MEDIA_REMOVAL, &pmr, sizeof(PREVENT_MEDIA_REMOVAL)
		, NULL, 0, &unused, NULL);
	CloseHandle(hdevice);
	return result ? OK : FAIL;
}



BIF_DECL(BIF_DriveGet)
{
	BuiltInFunctionID drive_get_cmd = _f_callee_id;

	// A parameter is mandatory for some, but optional for others:
	LPTSTR aValue = ParamIndexToOptionalString(0, _f_number_buf);
	if (!*aValue && drive_get_cmd != FID_DriveGetList && drive_get_cmd != FID_DriveGetStatusCD)
		_f_throw_value(ERR_PARAM1_MUST_NOT_BE_BLANK);
	
	if (drive_get_cmd == FID_DriveGetCapacity || drive_get_cmd == FID_DriveGetSpaceFree)
	{
		DriveSpace(aResultToken, aValue, drive_get_cmd == FID_DriveGetSpaceFree);
		return;
	}
	
	TCHAR path[T_MAX_PATH]; // T_MAX_PATH vs. MAX_PATH, though testing indicates only GetDriveType() supports long paths.
	size_t path_length;

	switch(drive_get_cmd)
	{

	case FID_DriveGetList:
	{
		UINT drive_type;
		#define ALL_DRIVE_TYPES 256
		if (!*aValue) drive_type = ALL_DRIVE_TYPES;
		else if (!_tcsicmp(aValue, _T("CDROM"))) drive_type = DRIVE_CDROM;
		else if (!_tcsicmp(aValue, _T("Removable"))) drive_type = DRIVE_REMOVABLE;
		else if (!_tcsicmp(aValue, _T("Fixed"))) drive_type = DRIVE_FIXED;
		else if (!_tcsicmp(aValue, _T("Network"))) drive_type = DRIVE_REMOTE;
		else if (!_tcsicmp(aValue, _T("RAMDisk"))) drive_type = DRIVE_RAMDISK;
		else if (!_tcsicmp(aValue, _T("Unknown"))) drive_type = DRIVE_UNKNOWN;
		else
			goto invalid_parameter;

		TCHAR found_drives[32];  // Need room for all 26 possible drive letters.
		int found_drives_count;
		UCHAR letter;
		TCHAR buf[128], *buf_ptr;

		for (found_drives_count = 0, letter = 'A'; letter <= 'Z'; ++letter)
		{
			buf_ptr = buf;
			*buf_ptr++ = letter;
			*buf_ptr++ = ':';
			*buf_ptr++ = '\\';
			*buf_ptr = '\0';
			UINT this_type = GetDriveType(buf);
			if (this_type == drive_type || (drive_type == ALL_DRIVE_TYPES && this_type != DRIVE_NO_ROOT_DIR))
				found_drives[found_drives_count++] = letter;  // Store just the drive letters.
		}
		found_drives[found_drives_count] = '\0';  // Terminate the string of found drive letters.
		// An empty list should not be flagged as failure, even for FIXED drive_type.
		// For example, when booting Windows PE from a REMOVABLE drive, it mounts a RAMDISK
		// drive but there may be no FIXED drives present.
		//if (!*found_drives)
		//	goto error;
		_f_return(found_drives);
	}

	case FID_DriveGetFilesystem:
	case FID_DriveGetLabel:
	case FID_DriveGetSerial:
	{
		TCHAR volume_name[256];
		TCHAR file_system[256];
		DRIVE_SET_PATH
		DWORD serial_number, max_component_length, file_system_flags;
		if (!GetVolumeInformation(path, volume_name, _countof(volume_name) - 1, &serial_number, &max_component_length
			, &file_system_flags, file_system, _countof(file_system) - 1))
			goto error;
		switch(drive_get_cmd)
		{
		case FID_DriveGetFilesystem: _f_return(file_system);
		case FID_DriveGetLabel:      _f_return(volume_name);
		case FID_DriveGetSerial:     _f_return(serial_number);
		}
		break;
	}

	case FID_DriveGetType:
	{
		DRIVE_SET_PATH
		switch (GetDriveType(path))
		{
		case DRIVE_UNKNOWN:   _f_return_p(_T("Unknown"));
		case DRIVE_REMOVABLE: _f_return_p(_T("Removable"));
		case DRIVE_FIXED:     _f_return_p(_T("Fixed"));
		case DRIVE_REMOTE:    _f_return_p(_T("Network"));
		case DRIVE_CDROM:     _f_return_p(_T("CDROM"));
		case DRIVE_RAMDISK:   _f_return_p(_T("RAMDisk"));
		default: // DRIVE_NO_ROOT_DIR
			_f_return_empty;
		}
		break;
	}

	case FID_DriveGetStatus:
	{
		DRIVE_SET_PATH
		DWORD sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters;
		switch (GetDiskFreeSpace(path, &sectors_per_cluster, &bytes_per_sector, &free_clusters, &total_clusters)
			? ERROR_SUCCESS : GetLastError())
		{
		case ERROR_SUCCESS:        _f_return_p(_T("Ready"));
		case ERROR_PATH_NOT_FOUND: _f_return_p(_T("Invalid"));
		case ERROR_NOT_READY:      _f_return_p(_T("NotReady"));
		case ERROR_WRITE_PROTECT:  _f_return_p(_T("ReadOnly"));
		default:                   _f_return_p(_T("Unknown"));
		}
		break;
	}

	case FID_DriveGetStatusCD:
		// Don't do DRIVE_SET_PATH in this case since trailing backslash might prevent it from
		// working on some OSes.
		// It seems best not to do the below check since:
		// 1) aValue usually lacks a trailing backslash so that it will work correctly with "open c: type cdaudio".
		//    That lack might prevent DriveGetType() from working on some OSes.
		// 2) It's conceivable that tray eject/retract might work on certain types of drives even though
		//    they aren't of type DRIVE_CDROM.
		// 3) One or both of the calls to mciSendString() will simply fail if the drive isn't of the right type.
		//if (GetDriveType(aValue) != DRIVE_CDROM) // Testing reveals that the below method does not work on Network CD/DVD drives.
		//	_f_throw(ERR_FAILED);
		TCHAR mci_string[256], status[128];
		// Note that there is apparently no way to determine via mciSendString() whether the tray is ejected
		// or not, since "open" is returned even when the tray is closed but there is no media.
		if (!*aValue) // When drive is omitted, operate upon default CD/DVD drive.
		{
			if (mciSendString(_T("status cdaudio mode"), status, _countof(status), NULL)) // Error.
				goto failed;
		}
		else // Operate upon a specific drive letter.
		{
			sntprintf(mci_string, _countof(mci_string), _T("open %s type cdaudio alias cd wait shareable"), aValue);
			if (mciSendString(mci_string, NULL, 0, NULL)) // Error.
				goto failed;
			MCIERROR error = mciSendString(_T("status cd mode"), status, _countof(status), NULL);
			mciSendString(_T("close cd wait"), NULL, 0, NULL);
			if (error)
				goto failed;
		}
		// Otherwise, success:
		_f_return(status);

	} // switch()

failed: // Failure where GetLastError() isn't relevant.
	_f_throw(ERR_FAILED);

error: // Win32 error
	_f_throw_win32();

invalid_parameter:
	_f_throw_param(0);
}
