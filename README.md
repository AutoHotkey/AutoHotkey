# AutoHotkey_L #

AutoHotkey is a free, open source macro-creation and automation software utility that allows users to automate repetitive tasks. It is driven by a custom scripting language that is aimed specifically at providing keyboard shortcuts, otherwise known as hotkeys.

AutoHotkey_L started as a fork of AutoHotkey but has been the main branch for some time.

https://autohotkey.com/


## How to Compile ##

AutoHotkey is developed with [Microsoft Visual Studio Community 2015 Express](https://www.visualstudio.com/products/visual-studio-community-vs), which is a free download from Microsoft.

  - Get the source code.
  - Open AutoHotkeyx.sln in Visual Studio.
  - Select the appropriate Build and Platform.
  - Build.

The project is configured to build with the Visual C++ 2010 toolset if available, primarily to facilitate Windows 2000 support but also because it appears to produce smaller 32-bit binaries than later versions. If the 2010 toolset is not available for a given platform, the project should automatically fall back to v140 (2015), v120 (2013) or v110 (2012).

Note that the fallback toolsets do not support targetting Windows XP. For that, install VS 2010 or change the platform toolset to v110_xp, v120_xp or v140_xp (if installed).

The project should also build in Visual C++ 2010, 2012 or 2013.


## Build Configurations ##

AutoHotkeyx.vcxproj contains several combinations of build configurations.  The main configurations are:

  - **Debug**: AutoHotkey.exe in debug mode.
  - **Release**: AutoHotkey.exe for general use.
  - **Self-contained**: AutoHotkeySC.bin, used for compiled scripts.

Secondary configurations are:

  - **(mbcs)**: ANSI (multi-byte character set). Configurations without this suffix are Unicode.
  - **(minimal)**: Alternative project settings for producing a smaller binary, possibly with lower performance and added dependencies.


## Platforms ##

AutoHotkeyx.vcxproj includes the following Platforms:

  - **Win32**: for Windows 32-bit.
  - **x64**: for Windows x64.

Visual C++ 2010 officially supports XP SP2 and later.  AutoHotkey supports Windows XP pre-SP2 and Windows 2000 via an asm patch (win2kcompat.asm).  Older versions are not supported.

## AutoHotkey v2 Alpha  ##

https://autohotkey.com/v2/

[v2 Branch](https://github.com/Lexikos/AutoHotkey_L/tree/alpha)


