#pragma once
// This file defines some macros for compile-time configrations.
// (Like many projects on *nix that using autotools.)

#ifndef AUTOHOTKEYSC
#define CONFIG_DEBUGGER
#endif

// Generates warnings to help we check whether the codes are ready to handle Unicode or not.
//#define CONFIG_UNICODE_CHECK

// A *lite* version of AutoHotkeyU perhaps.
//#define CONFIG_AUTOHOTKEY_LITE

// Includes experimental features
#define CONFIG_EXPERIMENTAL

// If you do not have ATL (Express version of VC++), undef this.
#define HAVE_ATL
