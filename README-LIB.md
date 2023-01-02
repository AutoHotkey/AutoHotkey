# AutoHotkeyLib

This project includes the capability to compile as a dll, for hosting the interpreter in another application. This was added in 2022, and is not to be confused with the AutoHotkey.dll first created by tinku99 in 2009 and later developed by HotKeyIt ([HotKeyIt/ahkdll](https://github.com/HotKeyIt/ahkdll/)) and thqby ([thqby/AutoHotkey_H](https://github.com/thqby/AutoHotkey_H/)). In contrast with those projects, this dll aims to provide a more convenient interface for hosting the interpreter, behaving as close to base AutoHotkey as is practical.

The dll has two primary use cases:
  - Host and execute a script; in particular, allow the use of v1 code within a v2 process.
  - Examine a script without executing it; provide data for auto-complete, calltips, navigating to function definitions, etc.

The dll is Unicode-only; ANSI is currently unsupported.

## Experimental

The syntax and behaviour described here may change between releases, without notice. Version numbers are purely for the base version of AutoHotkey and do not reflect the status of the Lib API.

## Building

Select the `Release.dll` build configuration and appropriate platform to build the dll.

Use the `Debug.dll` configuration for debugging, but in the project settings, set the debug command line to the path of an exe which will load the dll. If both the dll and the exe are debug builds, it should be possible to debug both at once (even if they are not in the same project/directory).

## Entry Points

The dll currently has two entry points.

### Host

Creates and returns an instance of AutoHotkeyLib.

```C++
HRESULT Host(IDispatch **ppLib);
```

This can be used from AutoHotkey v2 as follows, where *dll* contains the path or filename:
```AutoHotkey
if !hmod := DllCall("LoadLibrary", "str", dll, "ptr")
    throw OSError()
DllCall(dll "\Host", "ptr*", Lib := ComValue(9, 0), "hresult")
```

`Lib` then contains a COM object as described below.

The dll does not yet include the set of functions needed to implement a COM class factory, which would allow activation by means such as CoCreateInstance, ComObjCreate, ComObject or CreateObject (depending on the language/version). Such use would require registering the dll, whereas `Host` does not.

### Main

Executes the script in the same way that AutoHotkey.exe would. 

```C++
int Main(int argc, LPTSTR argv[]);
```
This was mainly for testing during early development, and is internally used by `Lib.Main(CmdLine)`.

## Lib API

The methods and properties exposed by the Lib object are defined in [ahklib.idl](source/ahklib.idl), in the  `IAutoHotkeyLib` interface. These are subject to change.

**General:**

  - `ExitCode := Lib.Main(CmdLine)` executes a command line more or less the same as AutoHotkey.exe would. If the script is persistent, it does not return. *CmdLine* must include "parameter 0" (which is normally the path of the exe), although its value is ignored. *ExitCode* is either 0 or 2 (CRITICAL_ERROR); ExitApp causes the whole program to terminate.
  - `Lib.LoadFile(FileName)` loads a script file but does not execute it. It can only be called once, unless the dll is unloaded and loaded again. The script is partially initialized for execution, but hotkeys are not manifested.
  - `Lib.OnProblem(Callback)` registers a callback to be called for each warning or error, including load-time errors (if registered before the script is loaded) and exceptions that aren't caught or handled by OnError. If a callback is set, the default error message is not shown. *Callback* receives a single parameter containing the thrown value or exception. If an exception object is created for warnings and errors that otherwise wouldn't have one, `What` is set to "Warn" or "Error".

**Execution:**

  - `ExitCode := Lib.Execute()` manifests hotkeys and excecutes the auto-execute section, then returns. Unlike *Main*, it does not initialize the command line arg variables or check for a previous instance of the script. *Execute* can be called multiple times. If this function is not called, it is possible to extract information about the script without executing it, or to execute specific functions without (or before) executing the auto-execute section.
  - `Lib.Script` returns an object which can be used to retrieve or set global variables (as properties) or call functions (as methods).

**Informational:**

  - `Lib.Funcs`, `Lib.Vars` and `Lib.Labels` return collections of objects describing functions, variables and labels. See `IDescribeFunc`, `IDescribeVar`, `IDescribeLabel` and `IDescribeParam` in [ahklib.idl](source/ahklib.idl) for usage. Collections use the `IDispCollection` interface, with Funcs, Vars and Labels accepting a name for *Index* and Params accepting a one-based index.

## Not Implemented

`Lib` does not provide any means to enumerate classes. It can be done by enumerating global variable names with `Lib.Vars`, retrieving the classes from `Lib.Script` and inspecting them, but this carries some risk because invoking the class object may cause unknown script to execute.

There are many other ways that the API could be extended, either to extract information or to control how the script executes or integrates with its host. 

## Known Issues

Events that execute in a new pseudo-thread may cause undefined behaviour if `Lib.Execute()` was skipped.

`A_AhkPath` refers to the path of the current process, so attempts to use it to execute scripts may not behave as intended.

Functions and menu items which execute a command line derived from the path of the current executable may not behave as intended. This includes `Reload`.

## Unloading/Reloading

Although the dll only supports loading one main script file, it is possible to execute or examine multiple scripts, one at a time, by unloading the dll and loading it again.
```AutoHotkey
DllCall("FreeLibrary", "ptr", hmod)  ; hmod was previously returned by LoadLibrary
```

Before unloading the dll, release all external references to objects created by the dll and ensure the keyboard/mouse hook is uninstalled (such as by suspending hotkeys). If the keyboard/mouse hook is still installed, calling FreeLibrary generally will not unload the dll. `Lib` currently does not provide any way to remove the hooks explicitly, so in some cases (e.g. if `#InstallKeybdHook` was used) it may not be possible.
