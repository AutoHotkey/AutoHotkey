#pragma once
// This file defines some macros for compile-time configrations.
// (Like many projects on *nix that using autotools.)

#if defined(WIN32) && !defined(_WIN64)
#define WIN32_PLATFORM
#endif

#ifdef _MSC_VER
	#if defined(WIN32_PLATFORM) || defined(_WIN64)
	#define ENABLE_DLLCALL
	#define ENABLE_REGISTERCALLBACK
	#endif
#endif

#if defined(AUTOHOTKEYSC) && !defined(_WIN64)
#define ENABLE_EXEARC
#endif

#if !defined(_MBCS) && !defined(_UNICODE) && !defined(UNICODE) // If not set in project settings...

// L: Comment out the next line to enable UNICODE:
//#define _MBCS

#ifndef _MBCS
#define _UNICODE
#define UNICODE
#endif
#endif

#ifndef AUTOHOTKEYSC
// DBGp
#define CONFIG_DEBUGGER
// Send vars as UTF-16
#ifdef UNICODE
//#define CONFIG_DBG_UTF16_SPEED_HACKS
#endif
#endif

// Generates warnings to help we check whether the codes are ready to handle Unicode or not.
//#define CONFIG_UNICODE_CHECK

// Includes experimental features
#define CONFIG_EXPERIMENTAL

#if !defined(UNICODE) && (!defined(_MSC_VER) || _MSC_VER < 1500)
// These should be defined if the compiler supports these platforms, otherwise run-time OS checks may be inaccurate.
#define CONFIG_WIN9X
#define CONFIG_WINNT4
#endif