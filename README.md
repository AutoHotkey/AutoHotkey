# AutoHotkey_L #

AutoHotkey is a free, open source macro-creation and automation software utility that allows users to automate repetitive tasks. It is driven by a custom scripting language that is aimed specifically at providing keyboard shortcuts, otherwise known as hotkeys.

AutoHotkey_L started as a fork of AutoHotkey but has been the main branch for some time.

https://autohotkey.com/


## How to Compile ##

AutoHotkey v2 is developed with [Microsoft Visual Studio Community 2019](https://www.visualstudio.com/products/visual-studio-community-vs), which is a free download from Microsoft.

  - Get the source code.
  - Open AutoHotkeyx.sln in Visual Studio.
  - Select the appropriate Build and Platform.
  - Build.

The project is configured in a way that allows building with Visual Studio 2012 or later. However, for the v2 branch, some newer C++ language features are used and therefore a later version of the compiler might be required.

The project is configured to use a platform toolset with "_xp" suffix, if available.


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

AutoHotkey supports Windows XP with or without service packs and Windows 2000 via an asm patch (win2kcompat.asm).  Support may be removed if maintaining it becomes non-trivial.  Older versions are not supported.

## AutoHotkey v2 Alpha  ##

https://autohotkey.com/v2/

[v2 Branch](https://github.com/Lexikos/AutoHotkey_L/tree/alpha)


