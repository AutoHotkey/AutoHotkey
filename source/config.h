#pragma once
// This file defines some macros for compile-time configrations.
// (Like many projects on *nix that using autotools.)

#ifndef AUTOHOTKEYSC
// DBGp
#define CONFIG_DEBUGGER
// Send vars as UTF-16
//#define CONFIG_DBG_UTF16_SPEED_HACKS
#endif

// Generates warnings to help we check whether the codes are ready to handle Unicode or not.
//#define CONFIG_UNICODE_CHECK

// A *lite* version of AutoHotkeyU perhaps.
//#define CONFIG_AUTOHOTKEY_LITE

// Includes experimental features
#define CONFIG_EXPERIMENTAL

#ifndef UNICODE
#define CONFIG_WIN9X
#endif

#if !defined(UNICODE) && (!defined(_MSC_VER) || _MSC_VER < 1500)
#define CONFIG_WINNT4
#endif

// If you do not have ATL (Express version of VC++), undef this.
#define HAVE_ATL
