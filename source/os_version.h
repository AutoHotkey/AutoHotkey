#ifndef __OS_VERSION_H
#define __OS_VERSION_H

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
// os_version.h
//
// A standalone class for easy checking of the OS version.
//
///////////////////////////////////////////////////////////////////////////////


// Includes
#include <windows.h>

class OS_Version
{
public:
	// Functions
	OS_Version() { Init(); }									// Constructor
	void	Init(void);											// Call first before use

#ifdef UNICODE
//	bool	IsWinNT(void) {return true;}
//	bool	IsWin9x(void) {return false;}
#else
	bool	IsWinNT(void) {return m_bWinNT;}					// Returns true if NT/2k/XP and family.
	bool	IsWin9x(void) {return m_bWin9x;}					// Returns true if 9x
#endif

	bool	IsWinNT4(void) {return m_bWinNT4;}					// Returns true if WinNT 4
	bool	IsWin2000(void) {return m_bWin2000;}				// Returns true if Win2000
	bool	IsWinXP(void) {return m_bWinXP;}					// Returns true if WinXP
	bool	IsWin2003(void) {return m_bWin2003;}				// Returns true if Win2003
	bool	IsWinVista(void) {return m_bWinVista;}				// Returns true if WinVista (v1.0.44.13)
	bool	IsWin7(void) {return m_bWin7; }						// Returns true if Win7
	bool	IsWinNT4orLater(void) {return m_bWinNT4orLater;}	// Returns true if WinNT 4+
	bool	IsWin2000orLater(void) {return m_bWin2000orLater;}	// Returns true if Win2000+
	bool	IsWinXPorLater(void) {return m_bWinXPorLater;}		// Returns true if WinXP+
	bool	IsWinVistaOrLater(void) {return m_bWinVistaOrLater;}// Returns true if WinVista or later (v1.0.48.01)
	bool	IsWin7OrLater(void) {return m_bWin7OrLater; }		// Returns true if Win7+

#ifdef UNICODE
	//bool	IsWin95(void) {return false;}
	//bool	IsWin98(void) {return false;}
	//bool	IsWinMe(void) {return false;}
	//bool	IsWin95orLater(void) {return false;}
	//bool	IsWin98orLater(void) {return false;}
	//bool	IsWinMeorLater(void) {return false;}
#else
	bool	IsWin95(void) {return m_bWin95;}					// Returns true if 95
	bool	IsWin98(void) {return m_bWin98;}					// Returns true if 98
	bool	IsWinMe(void) {return m_bWinMe;}					// Returns true if Me
	bool	IsWin95orLater(void) {return m_bWin95orLater;}		// Returns true if 95
	bool	IsWin98orLater(void) {return m_bWin98orLater;}		// Returns true if 98
	bool	IsWinMeorLater(void) {return m_bWinMeorLater;}		// Returns true if Me
#endif

	DWORD	BuildNumber(void) {return m_dwBuildNumber;}
	LPCTSTR CSD(void) {return m_szCSDVersion;}

private:
	// Variables
	OSVERSIONINFO	m_OSvi;						// OS Version data

	DWORD			m_dwMajorVersion;			// Major OS version
	DWORD			m_dwMinorVersion;			// Minor OS version
	DWORD			m_dwBuildNumber;			// Build number
	TCHAR			m_szCSDVersion [256];

#ifndef UNICODE
	bool			m_bWinNT;
	bool			m_bWin9x;
#endif

	bool			m_bWinNT4;
	bool			m_bWinNT4orLater;
	bool			m_bWinXP;
	bool			m_bWin2003;
	bool			m_bWinVista;
	bool			m_bWin7;
	bool			m_bWinXPorLater;
	bool			m_bWin2000;
	bool			m_bWin2000orLater;
	bool			m_bWinVistaOrLater;
	bool			m_bWin7OrLater;
#ifndef UNICODE
	bool			m_bWin98;
	bool			m_bWin98orLater;
	bool			m_bWin95;
	bool			m_bWin95orLater;
	bool			m_bWinMe;
	bool			m_bWinMeorLater;
#endif
};

///////////////////////////////////////////////////////////////////////////////

#endif

