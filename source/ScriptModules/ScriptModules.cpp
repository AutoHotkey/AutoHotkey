/*
AutoHotkey
Copyright 2003-2009 Chris Mallett (support@autohotkey.com)
This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.
This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#include "..\stdafx.h" // pre-compiled headers
#include "ScriptModules.h"

const LPTSTR ScriptModule::sUnamedModuleName = SMODULES_UNNAMED_NAME;

ScriptModule::~ScriptModule()
{
	if (mNested)
		delete mNested;
	if (mOptionalModules)
		delete mOptionalModules;
#ifndef AUTOHOTKEYSC
	FreeSourceFileIndexList();
#endif
}

// ScriptModule methods,
#include "ScriptModulesMisc.h"				// Misc. methods.
#include "ScriptModulesOptional.h"			// Handling of optional modules, "#Import *i ...".
#include "ScriptModulesAddNames.h"			// Handling of #UseXXX directives.
#include "ScriptModulesReleaseObjects.h"	// For releasing of objects when the program terminates.
#include "ScriptModulesNested.h"			// Handling of nested modules.

#ifndef AUTOHOTKEYSC
#include "ScriptModulesSourceFileIndex.h"	// For tracking #include:ed files,
											// to allow a file to be included into more than one module
											// without using #includeAgain
					
#endif 
// End ScriptModule methods.

#include "ModuleListMethods.h"