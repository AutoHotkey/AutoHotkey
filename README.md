# AutoHotkey_L # - 

pending LCCB approval in Shanghai for local use in Consular Operations

AutoHotkey is a free, open source macro-creation and automation software utility that allows users to automate repetitive tasks. It is driven by a custom scripting language that is aimed specifically at providing keyboard shortcuts, otherwise known as hotkeys.

AutoHotkey_L started as a fork of AutoHotkey but has been the main branch for some time.

http://ahkscript.org/


## How to Compile ##

AutoHotkey_L is developed with [Microsoft Visual C++ 2010 Express](http://www.microsoft.com/visualstudio/en-us/products/2010-editions/visual-cpp-express), which is a free download from Microsoft.

  - Get the source code.
  - Open AutoHotkeyx.sln in VC++ 2010 Express.
  - Select the appropriate Build and Platform.
  - Build.

Windows 7.1 SDK may be required for x64 builds.  Installing the SDK compilers and Visual Studio 2010 SP1 together can prove difficult; see [KB2519277](http://support.microsoft.com/kb/2519277) for the solution.

Visual Studio 2010 or MSBuild in the Windows 7.1 SDK (with Visual C++ 2010 compilers installed) should also work.

Visual Studio 2013 will work if the Platform Toolset is changed in the project settings, but does not support Windows XP SP2 or earlier.


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

Visual C++ 2010 officially supports XP SP2 and later.  AutoHotkey_L supports Windows XP pre-SP2 and Windows 2000 via an asm patch (win2kcompat.asm).  Older versions are not supported.
