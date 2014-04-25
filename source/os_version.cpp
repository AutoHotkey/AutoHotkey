
///////////////////////////////////////////////////////////////////////////////
//
// AutoIt
//
// Copyright (C)1999-2003:
//		- Jonathan Bennett <jon@hiddensoft.com>
//		- Others listed at http://www.autoitscript.com/autoit3/docs/credits.htm
//
// This program is free software; you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation; either version 2 of the License, or
// (at your option) any later version.
//
///////////////////////////////////////////////////////////////////////////////
//
// os_version.cpp
//
// A standalone class for easy checking of the OS version.
//
///////////////////////////////////////////////////////////////////////////////


// Includes
#include "stdafx.h" // pre-compiled headers

#ifndef _MSC_VER								// Includes for non-MS compilers
	#include <windows.h>
#endif

#include "os_version.h"


/*
OSVERSIONINFO structure details
===============================

dwOSVersionInfoSize
Specifies the size, in bytes, of this data structure. Set this member to sizeof(OSVERSIONINFO)
before calling the GetVersionEx function.

dwMajorVersion
Identifies the major version number of the operating system as follows. Operating System Value
Windows 95				4
Windows 98				4
Windows Me				4
Windows NT 3.51			3
Windows NT 4.0			4
Windows 2000/XP/2003	5
Windows Vista/7			6

dwMinorVersion
Identifies the minor version number of the operating system as follows. Operating System Value
Windows 95			0
Windows 98			10
Windows Me			90
Windows NT 3.51		51
Windows NT 4.0		0
Windows 2000		0
Windows XP			1
Windows 2003		2
Windows Vista		0 (probably 0 for all Vista variants)
Windows 7			1

dwBuildNumber
Windows NT/2000: Identifies the build number of the operating system.
Windows 95/98: Identifies the build number of the operating system in the low-order word.
The high-order word contains the major and minor version numbers.

dwPlatformId
Identifies the operating system platform. This member can be one of the following values. Value Platform
VER_PLATFORM_WIN32s			Win32s on Windows 3.1.
VER_PLATFORM_WIN32_WINDOWS	Windows 95, Windows 98, or Windows Me.
VER_PLATFORM_WIN32_NT		Windows NT 3.51, Windows NT 4.0, Windows 2000, Windows XP, Windows Vista or Windows 7.

szCSDVersion
Windows NT/2000, Whistler: Contains a null-terminated string, such as "Service Pack 3",
that indicates the latest Service Pack installed on the system. If no Service Pack has
been installed, the string is empty.
Windows 95/98/Me: Contains a null-terminated string that indicates additional version
information. For example, " C" indicates Windows 95 OSR2 and " A" indicates Windows 98 SE.

*/


///////////////////////////////////////////////////////////////////////////////
// Init()
///////////////////////////////////////////////////////////////////////////////

void OS_Version::Init(void)
{
	int		i;
	int		nTemp;

	// Get details of the OS we are running on
	m_OSvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
	GetVersionEx(&m_OSvi);

	// Populate Major and Minor version numbers
	m_dwMajorVersion	= m_OSvi.dwMajorVersion;
	m_dwMinorVersion	= m_OSvi.dwMinorVersion;
	m_dwBuildNumber		= m_OSvi.dwBuildNumber;

	// Get CSD information
	nTemp = (int)_tcslen(m_OSvi.szCSDVersion);

	if (nTemp > 0)
	{
		//	strip trailing
		for (i=nTemp-1; i>0; i--)
		{
			if (m_OSvi.szCSDVersion[i] != ' ') 
				break;
			m_OSvi.szCSDVersion[i] = '\0';
		}

		//	strip leading
		nTemp = i;
		for (i=0; i<nTemp; i++)
		{
			if (m_OSvi.szCSDVersion[i] != ' ') 
				break;
		}
		_tcscpy(m_szCSDVersion, &m_OSvi.szCSDVersion[i]);
	}
	else
		m_szCSDVersion[0] = '\0';				// No CSD info, make it blank to avoid errors


	// Set all options to false by default
#ifdef CONFIG_WIN9X
	m_bWinNT	= false;
	m_bWin9x	= false;

	m_bWin95	= false;	m_bWin95orLater		= false;
	m_bWin98	= false;	m_bWin98orLater		= false;
	m_bWinMe	= false;	m_bWinMeorLater		= false;
#endif
#ifdef CONFIG_WINNT4
	m_bWinNT4	= false;	m_bWinNT4orLater	= false;
#endif
	m_bWin2000	= false;	m_bWin2000orLater	= false;
	m_bWinXP	= false;	m_bWinXPorLater		= false;
	m_bWin2003  = false;
	m_bWinVista = false;	m_bWinVistaOrLater	= false;
	m_bWin7		= false;	m_bWin7OrLater		= false;
	m_bWin8		= false;

#ifdef CONFIG_WIN9X
	// Work out if NT or 9x
	if (m_OSvi.dwPlatformId == VER_PLATFORM_WIN32_NT)
	{
		// Windows NT
		m_bWinNT = true;
#endif

		switch (m_dwMajorVersion)
		{
#ifdef CONFIG_WINNT4
			case 4:								// NT 4
				m_bWinNT4 = true;
				m_bWinNT4orLater = true;
				break;
#endif

			case 5:								// Win2000 / XP
#ifdef CONFIG_WINNT4
				m_bWinNT4orLater = true;
#endif
				m_bWin2000orLater = true;
				if ( m_dwMinorVersion == 0 )	// Win2000
					m_bWin2000 = true;
				else // minor is 1 (XP), 2 (2003), or beyond.
				{
					m_bWinXPorLater = true;
					if ( m_dwMinorVersion == 1 )
						m_bWinXP = true;
					else if ( m_dwMinorVersion == 2 )
						m_bWin2003 = true;
					//else it's something later than XP/2003, so there is nothing more to be done.
				}
				break;
			case 6:
				if (m_dwMinorVersion == 0)
					m_bWinVista = true;
				else {
					m_bWin7OrLater = true;
					if (m_dwMinorVersion == 1)
						m_bWin7 = true;
					else if (m_dwMinorVersion == 2)
						m_bWin8 = true;
					else if (m_dwMinorVersion == 3)
						m_bWin8_1 = true;
				}
				m_bWinVistaOrLater = true;
				m_bWinXPorLater = true;
				m_bWin2000orLater = true;
#ifdef CONFIG_WINNT4
				m_bWinNT4orLater = true;
#endif
				break;
			default:
				if (m_dwMajorVersion > 6)
				{
					m_bWin7OrLater = true;
					m_bWinVistaOrLater = true;
					m_bWinXPorLater = true;
					m_bWin2000orLater = true;
#ifdef CONFIG_WINNT4
					m_bWinNT4orLater = true;
#endif
				}
  				break;

		} // End Switch
#ifdef CONFIG_WIN9X
	}
	else
	{
		// Windows 9x -- all major versions = 4
		m_bWin9x = true;
		m_bWin95orLater = true;
		m_dwBuildNumber	= (WORD) m_OSvi.dwBuildNumber;	// Build number in lower word on 9x

		switch ( m_dwMinorVersion )
		{
			case 0:								// 95
				m_bWin95 = true;
				break;

			case 10:							// 98
				m_bWin98 = m_bWin98orLater = true;
				break;

			case 90:							// ME
				m_bWinMe = 	m_bWinMeorLater = m_bWin98orLater = true;
				break;
		} // End Switch
	} // End If
#endif

} // Init()

