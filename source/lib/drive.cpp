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
#include "abi.h"



static bool IsValidDriveLetter(LPCTSTR aDrive)
{
	return cisalpha(*aDrive)
		&& (!aDrive[1] || (aDrive[1] == ':'
			&& !aDrive[2] || (aDrive[2] == '\\'
				&& !aDrive[3])));
}



template<size_t n>
static void DriveFixPath(LPCTSTR &aDrive, TCHAR (&aBuf)[n])
{
	// Append a backslash because certain API calls or OS versions might require it;
	// but leave aDrive as-is if it's too long for the buffer, so that if long path
	// support within the OS is improved in future, callers can utilize it.
	auto path_length = _tcslen(aDrive);
	if (path_length && path_length < _countof(aBuf) && aDrive[path_length - 1] != '\\')
	{
		tmemcpy(aBuf, aDrive, path_length);
		aBuf[path_length++] = '\\';
		aBuf[path_length] = '\0';
		aDrive = aBuf;
	}
}



static FResult DriveSpace(LPCTSTR aPath, __int64 *aRetVal, bool aGetFreeSpace)
// Because of NTFS's ability to mount volumes into a directory, a path might not necessarily
// have the same amount of free space as its root drive.  However, I'm not sure if this
// method here actually takes that into account.
{
	if (!*aPath)
		return FR_E_ARG(0);

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
	ULARGE_INTEGER total, free, used;
	if (!GetDiskFreeSpaceEx(buf, &free, &total, &used))
		return FR_E_WIN32;
	// Casting this way allows sizes of up to 2,097,152 gigabytes:
	*aRetVal = (__int64)((unsigned __int64)(aGetFreeSpace ? free.QuadPart : total.QuadPart) / 1048576);
	return OK;
}

bif_impl FResult DriveGetCapacity(LPCTSTR aPath, __int64 *aRetVal)
{
	return DriveSpace(aPath, aRetVal, false);
}

bif_impl FResult DriveGetSpaceFree(LPCTSTR aPath, __int64 *aRetVal)
{
	return DriveSpace(aPath, aRetVal, true);
}



static FResult DriveEject(LPCTSTR aDrive, bool aEject)
{
	// Don't do DRIVE_SET_PATH in this case since trailing backslash is not wanted.
	// It seems best not to do the below check since:
	// 1) aValue usually lacks a trailing backslash, which might prevent DriveGetType() from working on some OSes.
	// 2) Eject (and more rarely, retract) works on some other drive types.
	// 3) CreateFile or DeviceIoControl will simply fail or have no effect if the drive isn't of the right type.
	//if (GetDriveType(aDrive) != DRIVE_CDROM)
	//	return FR_E_FAILED;
	TCHAR path[] { '\\', '\\', '.', '\\', 0, ':', '\0', '\0' };
	if (aDrive)
	{
		// Testing showed that a Volume GUID of the form \\?\Volume{...} will work even when
		// the drive isn't mapped to a drive letter.
		if (IsValidDriveLetter(aDrive))
		{
			path[4] = aDrive[0];
			aDrive = path;
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
				return FR_E_FAILED; // No CD/DVD drive found with a drive letter.  
		}
		path[6] = '\0'; // Remove the trailing slash for CreateFile to open the volume.
		aDrive = path;
	}
	// Testing indicates neither this method nor the MCI method work with mapped drives or UNC paths.
	// That makes sense when one considers that the following opens the *volume*, whereas a network
	// share would correspond to a directory; i.e. this needs "\\.\D:" and not "\\.\D:\".
	HANDLE hVol = CreateFile(aDrive, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
	if (hVol == INVALID_HANDLE_VALUE)
		return FR_E_WIN32;
	DWORD unused;
	auto successful = DeviceIoControl(hVol, aEject ? IOCTL_STORAGE_EJECT_MEDIA : IOCTL_STORAGE_LOAD_MEDIA, NULL, 0, NULL, 0, &unused, NULL);
	CloseHandle(hVol);
	return successful ? OK : FR_E_WIN32;
}

bif_impl FResult DriveEject(LPCTSTR aDrive)
{
	return DriveEject(aDrive, true);
}

bif_impl FResult DriveRetract(LPCTSTR aDrive)
{
	return DriveEject(aDrive, false);
}



bif_impl FResult DriveSetLabel(LPCTSTR aDrive, LPCTSTR aLabel)
{
	TCHAR buf[MAX_PATH]; // MAX_PATH vs. T_MAX_PATH because SetVolumeLabel() can't seem to make use of long paths.
	DriveFixPath(aDrive, buf);
	return SetVolumeLabel(aDrive, aLabel) ? OK : FR_E_WIN32;
}



static FResult DriveLock(LPCTSTR aDrive, bool aLockIt)
{
	if (!IsValidDriveLetter(aDrive))
		return FR_E_ARG(0);

	HANDLE hdevice;
	DWORD unused;
	BOOL result;
	TCHAR filename[64];
	_stprintf(filename, _T("\\\\.\\%c:"), *aDrive);
	// FILE_READ_ATTRIBUTES is not enough; it yields "Access Denied" error.  So apparently all or part
	// of the sub-attributes in GENERIC_READ are needed.  An MSDN example implies that GENERIC_WRITE is
	// only needed for GetDriveType() == DRIVE_REMOVABLE drives, and maybe not even those when all we
	// want to do is lock/unlock the drive (that example did quite a bit more).  In any case, research
	// indicates that all CD/DVD drives are ever considered DRIVE_CDROM, not DRIVE_REMOVABLE.
	// Due to this and the unlikelihood that GENERIC_WRITE is ever needed anyway, GetDriveType() is
	// not called for the purpose of conditionally adding the GENERIC_WRITE attribute.
	hdevice = CreateFile(filename, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING, 0, NULL);
	if (hdevice == INVALID_HANDLE_VALUE)
		return FR_E_WIN32;
	PREVENT_MEDIA_REMOVAL pmr;
	pmr.PreventMediaRemoval = aLockIt;
	result = DeviceIoControl(hdevice, IOCTL_STORAGE_MEDIA_REMOVAL, &pmr, sizeof(PREVENT_MEDIA_REMOVAL)
		, NULL, 0, &unused, NULL);
	CloseHandle(hdevice);
	return result ? OK : FR_E_WIN32;
}

bif_impl FResult DriveLock(LPCTSTR aDrive)
{
	return DriveLock(aDrive, true);
}

bif_impl FResult DriveUnlock(LPCTSTR aDrive)
{
	return DriveLock(aDrive, false);
}



bif_impl FResult DriveGetList(LPCTSTR aType, StrRet &aRetVal)
{
	UINT drive_type;
	#define ALL_DRIVE_TYPES 256
	if (!aType) drive_type = ALL_DRIVE_TYPES;
	else if (!_tcsicmp(aType, _T("CDROM"))) drive_type = DRIVE_CDROM;
	else if (!_tcsicmp(aType, _T("Removable"))) drive_type = DRIVE_REMOVABLE;
	else if (!_tcsicmp(aType, _T("Fixed"))) drive_type = DRIVE_FIXED;
	else if (!_tcsicmp(aType, _T("Network"))) drive_type = DRIVE_REMOTE;
	else if (!_tcsicmp(aType, _T("RAMDisk"))) drive_type = DRIVE_RAMDISK;
	else if (!_tcsicmp(aType, _T("Unknown"))) drive_type = DRIVE_UNKNOWN;
	else
		return FR_E_ARG(0);

	LPTSTR found_drives = aRetVal.Alloc(26); // Need room for all 26 possible drive letters.
	if (!found_drives)
		return FR_E_OUTOFMEM;
	int found_drives_count;
	TCHAR letter[] = { 0, ':', '\\', '\0' };
	for (found_drives_count = 0, *letter = 'A'; *letter <= 'Z'; ++*letter)
	{
		UINT this_type = GetDriveType(letter);
		if (this_type == drive_type || (drive_type == ALL_DRIVE_TYPES && this_type != DRIVE_NO_ROOT_DIR))
			found_drives[found_drives_count++] = *letter;  // Store just the drive letters.
	}
	found_drives[found_drives_count] = '\0';  // Terminate the string of found drive letters.
	// An empty list should not be flagged as failure, even for FIXED drive_type.
	// For example, when booting Windows PE from a REMOVABLE drive, it mounts a RAMDISK
	// drive but there may be no FIXED drives present.
	//if (!*found_drives)
	//	return FR_E_FAILED;
	return OK;
}



bif_impl FResult DriveGetFilesystem(LPCTSTR aDrive, StrRet &aRetVal)
{
	TCHAR buf[MAX_PATH]; // MAX_PATH vs. T_MAX_PATH because testing in 2019 indicated GetVolumeInformation() did not support long paths.
	DriveFixPath(aDrive, buf);
	LPTSTR file_system = aRetVal.CallerBuf();
	if (!GetVolumeInformation(aDrive, NULL, 0, NULL, NULL, NULL, file_system, StrRet::CallerBufSize))
		return FR_E_WIN32;
	aRetVal.SetTemp(file_system);
	return OK;
}



bif_impl FResult DriveGetLabel(LPCTSTR aDrive, StrRet &aRetVal)
{
	TCHAR buf[MAX_PATH]; // MAX_PATH vs. T_MAX_PATH because testing in 2019 indicated GetVolumeInformation() did not support long paths.
	DriveFixPath(aDrive, buf);
	LPTSTR volume_name = aRetVal.CallerBuf();
	if (!GetVolumeInformation(aDrive, volume_name, StrRet::CallerBufSize, NULL, NULL, NULL, NULL, 0))
		return FR_E_WIN32;
	aRetVal.SetTemp(volume_name);
	return OK;
}



bif_impl FResult DriveGetSerial(LPCTSTR aDrive, __int64 *aRetVal)
{
	TCHAR buf[MAX_PATH]; // MAX_PATH vs. T_MAX_PATH because testing in 2019 indicated GetVolumeInformation() did not support long paths.
	DriveFixPath(aDrive, buf);
	DWORD serial_number;
	if (!GetVolumeInformation(aDrive, NULL, 0, &serial_number, NULL, NULL, NULL, 0))
		return FR_E_WIN32;
	*aRetVal = serial_number;
	return OK;
}



bif_impl FResult DriveGetType(LPCTSTR aDrive, StrRet &aRetVal)
{
	TCHAR buf[T_MAX_PATH]; // T_MAX_PATH vs. MAX_PATH because GetDriveType() can support long paths.
	DriveFixPath(aDrive, buf);
	switch (GetDriveType(aDrive))
	{
	case DRIVE_UNKNOWN:   aRetVal.SetStatic(_T("Unknown")); break;
	case DRIVE_REMOVABLE: aRetVal.SetStatic(_T("Removable")); break;
	case DRIVE_FIXED:     aRetVal.SetStatic(_T("Fixed")); break;
	case DRIVE_REMOTE:    aRetVal.SetStatic(_T("Network")); break;
	case DRIVE_CDROM:     aRetVal.SetStatic(_T("CDROM")); break;
	case DRIVE_RAMDISK:   aRetVal.SetStatic(_T("RAMDisk")); break;
	default: // DRIVE_NO_ROOT_DIR
		aRetVal.SetEmpty();
	}
	return OK;
}



bif_impl FResult DriveGetStatus(LPCTSTR aDrive, StrRet &aRetVal)
{
	TCHAR buf[MAX_PATH]; // MAX_PATH vs. T_MAX_PATH because testing in 2019 indicated GetDiskFreeSpace() did not support long paths.
	DriveFixPath(aDrive, buf);
	DWORD sectors_per_cluster, bytes_per_sector, free_clusters, total_clusters;
	switch (GetDiskFreeSpace(aDrive, &sectors_per_cluster, &bytes_per_sector, &free_clusters, &total_clusters)
		? ERROR_SUCCESS : GetLastError())
	{
	case ERROR_SUCCESS:			aRetVal.SetStatic(_T("Ready")); break;
	case ERROR_FILE_NOT_FOUND:
	case ERROR_PATH_NOT_FOUND:	aRetVal.SetStatic(_T("Invalid")); break;
	case ERROR_NOT_READY:		aRetVal.SetStatic(_T("NotReady")); break;
	case ERROR_WRITE_PROTECT:	aRetVal.SetStatic(_T("ReadOnly")); break;
	default:					aRetVal.SetStatic(_T("Unknown")); break;
	}
	return OK;
}



bif_impl FResult DriveGetStatusCD(LPCTSTR aDrive, StrRet &aRetVal)
{
	// Explicitly validate aDrive since it'll be inserted into the MCI command string.
	if (!IsValidDriveLetter(aDrive))
		return FR_E_ARG(0);

	// Don't do DRIVE_SET_PATH in this case since trailing backslash might prevent it from
	// working on some OSes.
	// It seems best not to do the below check since:
	// 1) aValue usually lacks a trailing backslash so that it will work correctly with "open c: type cdaudio".
	//    That lack might prevent DriveGetType() from working on some OSes.
	// 2) It's conceivable that tray eject/retract might work on certain types of drives even though
	//    they aren't of type DRIVE_CDROM.
	// 3) One or both of the calls to mciSendString() will simply fail if the drive isn't of the right type.
	//if (GetDriveType(aValue) != DRIVE_CDROM) // Testing reveals that the below method does not work on Network CD/DVD drives.
	//	return FR_E_FAILED;
	TCHAR mci_string[256];
	const auto status_size = 16;
	LPTSTR status = aRetVal.Alloc(status_size);
	// Note that there is apparently no way to determine via mciSendString() whether the tray is ejected
	// or not, since "open" is returned even when the tray is closed but there is no media.
	if (!aDrive) // When drive is omitted, operate upon default CD/DVD drive.
	{
		if (mciSendString(_T("status cdaudio mode"), status, status_size, NULL)) // Error.
			return FR_E_FAILED;
	}
	else // Operate upon a specific drive letter.
	{
		sntprintf(mci_string, _countof(mci_string), _T("open %s type cdaudio alias cd wait shareable"), aDrive);
		if (mciSendString(mci_string, NULL, 0, NULL)) // Error.
			return FR_E_FAILED;
		MCIERROR error = mciSendString(_T("status cd mode"), status, status_size, NULL);
		mciSendString(_T("close cd wait"), NULL, 0, NULL);
		if (error)
			return FR_E_FAILED;
	}
	// Otherwise, success:
	return OK;
}
