# AutoHotkey #

AutoHotkey is a free, open source macro-creation and automation software utility that allows users to automate repetitive tasks. It is driven by a custom scripting language that has special provision for defining keyboard shortcuts, otherwise known as hotkeys.

https://www.autohotkey.com/


## Support ##

The [AutoHotkey Community forum](https://www.autohotkey.com/boards/) is the primary source of support for AutoHotkey.

AutoHotkey v1 is not being maintained, but support is provided by community members.

* **Scripts not working**: For assistance getting code AutoHotkey scripts to work the way you want, start a topic in the [Ask for Help (v2)](https://www.autohotkey.com/boards/viewforum.php?f=82) or [Ask for Help (v1)](https://www.autohotkey.com/boards/viewforum.php?f=76) subforum, depending on your AutoHotkey version.
* **Bug reporting**: If in doubt about the nature of your issue, please post in [Ask for Help (v2)](https://www.autohotkey.com/boards/viewforum.php?f=82) for confirmation. Otherwise, bugs should be reported in the [Bug Reports](https://www.autohotkey.com/boards/viewforum.php?f=14) subforum.
* **False positives**: If you notice any AutoHotkey files (downloaded from official sources) are being flagged as suspicious or a virus, they are likely false positives. Please refer to our page on how to resolve or report these [here](https://www.autohotkey.com/download/safe.htm).
* **Other development topics**: For any other development related enquiries, please utilize the [AutoHotkey Development](https://www.autohotkey.com/boards/viewforum.php?f=37) subforum.


## How to Compile ##

AutoHotkey is developed with [Microsoft Visual Studio Community 2022](https://www.visualstudio.com/products/visual-studio-community-vs), which is a free download from Microsoft.

  - Get the source code.
  - Open AutoHotkeyx.sln in Visual Studio.
  - Select the appropriate Build and Platform.
  - Build.

The project is configured in a way that allows building with Visual Studio 2012 or later, but only the 2022 toolset is regularly tested. Some newer C++ language features are used and therefore a later version of the compiler might be required.


## Developing in VS Code ##

AutoHotkey v2 can also be built and debugged in VS Code.

Requirements:
  - [C/C++ for Visual Studio Code](https://marketplace.visualstudio.com/items?itemName=ms-vscode.cpptools). VS Code might prompt you to install this if you open a .cpp file.
  - [Build Tools for Visual Studio 2022](https://aka.ms/vs/17/release/vs_BuildTools.exe) with the "Desktop development with C++" workload, or similar (some older or newer versions and different products should work).



## Build Configurations ##

AutoHotkeyx.vcxproj contains several combinations of build configurations.  The main configurations are:

  - **Debug**: AutoHotkey.exe in debug mode.
  - **Release**: AutoHotkey.exe for general use.
  - **Self-contained**: AutoHotkeySC.bin, used for compiled scripts.

Secondary configurations are:

  - **(mbcs)**: ANSI (multi-byte character set). Configurations without this suffix are Unicode.
  - **.dll**: Builds an experimental dll for use hosting the interpreter, such as to enable the use of v1 libraries in a v2 script. See [README-LIB.md](README-LIB.md).


## Platforms ##

AutoHotkeyx.vcxproj includes the following Platforms:

  - **Win32**: for Windows 32-bit.
  - **x64**: for Windows x64.

AutoHotkey v2 should run on Windows 7 and later, but each release is tested only on Windows 11.

Support for Windows XP was dropped by commit b1ac557a. It may be possible to build binaries for Windows XP by adjusting linker settings and using an appropriate compiler toolset, but some features may not work. The removed support included version checks affecting use of fonts packaged with later OS versions, owner-drawn menu icons, Sound APIs, and some other miscellaneous version checks for pre-Vista compatibility.
