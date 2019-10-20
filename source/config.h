#pragma once
// This file defines some macros for compile-time configurations.
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
#endif

// Generates warnings to help we check whether the codes are ready to handle Unicode or not.
//#define CONFIG_UNICODE_CHECK

// This is now defined via Config.vcxproj if supported by the current platform toolset.
//#ifndef _WIN64
//#define CONFIG_WIN2K
//#endif