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

	bool	IsWinVista(void) {return m_bWinVista;}				// Returns true if WinVista (v1.0.44.13)
	bool	IsWin7(void) {return m_bWin7; }						// Returns true if Win7
	bool	IsWin8(void) {return m_bWin8; }						// Returns true if Win8
	bool	IsWin8_1(void) {return m_bWin8_1; }					// Returns true if Win8.1
	bool	IsWin7OrLater(void) {return m_bWin7OrLater; }		// Returns true if Win7+
	bool	IsWin10OrLater(void) {return m_dwMajorVersion >= 10;} // Excludes early pre-release builds.

	DWORD	BuildNumber(void) {return m_dwBuildNumber;}
	//LPCTSTR CSD(void) {return m_szCSDVersion;}
	LPCTSTR Version() {return m_szVersion;}

private:
	// Variables
	OSVERSIONINFOW	m_OSvi;						// OS Version data

	DWORD			m_dwMajorVersion;			// Major OS version
	DWORD			m_dwMinorVersion;			// Minor OS version
	DWORD			m_dwBuildNumber;			// Build number
	//TCHAR			m_szCSDVersion[128];
	TCHAR			m_szVersion[32];			// "Major.Minor.Build" -- longest known number is 9 chars (plus terminator), but 32 should be future-proof.

#ifdef CONFIG_WIN9X
	bool			m_bWinNT;
	bool			m_bWin9x;

	bool			m_bWin95;
	bool			m_bWin95orLater;
	bool			m_bWin98;
	bool			m_bWin98orLater;
	bool			m_bWinMe;
	bool			m_bWinMeorLater;
#endif
#ifdef CONFIG_WINNT4
	bool			m_bWinNT4;
	bool			m_bWinNT4orLater;
#endif
#ifdef CONFIG_WIN2K
	bool			m_bWin2000;
	bool			m_bWin2000orLater; // For simplicity, this is always left in even though it is not used when !(CONFIG_WIN9X || CONFIG_WINNT4).
	bool			m_bWinXPorLater;
#endif
	bool			m_bWinXP;
	bool			m_bWin2003;
	bool			m_bWinVista;
	bool			m_bWinVistaOrLater;
	bool			m_bWin7;
	bool			m_bWin7OrLater;
	bool			m_bWin8;
	bool			m_bWin8_1;
};

///////////////////////////////////////////////////////////////////////////////

#endif

