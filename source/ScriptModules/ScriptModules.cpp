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

bool ScriptModule::ValidateName(LPTSTR aName)
{
	// return true if this name can be used as a variable name in this module, i.e., it is not the same as one of its nested module names or any of the reserved names.
	// else returns false.
	// Relies on being called from Var::ValidateName.
	if SMODULES_NAMES_MATCH(aName, SMODULES_OUTER_MODULE_NAME) // Need to check since SMODULES_DEFAULT_MODULE_NAME's mOuter is NULL.
		return false;
	return !GetNestedModule(aName, true);
}

bool ScriptModule::AddOptionalModule(LPTSTR aName)
{
	// adds aName to a list of module names which will not cause load time error if not found.
	// to allow the "*i" option for SMODULES_INCLUDE_DIRECTIVE_NAME.
	// returns false on out of memory, else true.

	if (!mOptionalModules)
		mOptionalModules = new ModuleNameList;
	LPTSTR name = _tcsdup(aName);
	if (!name)
		return false;
	if (!mOptionalModules->AddItem(name))
	{
		free(name);
		return false;
	}
	return true;
}

bool ScriptModule::IsOptionalModule(LPTSTR aName)
{
	// returns true if aName is in the list of optional modules.
	if (!mOptionalModules)
		return false;
	return mOptionalModules->HasItem(aName);
}

ScriptModule* ScriptModule::FindModuleFromDotDelimitedString(LPTSTR aString)
{
	// This is very similar to FindClassFromDotDelimitedString, maintain together.
	// aString, a dot (.) delitmited string. Relies on that each module name is verified elsewhere.

	// Call BIF_StrSplit:
	// Array: = StrSplit(String, Delimiters, OmitChars, MaxParts : = -1)
	CALL_BIF(StrSplit, strsplit_result,
		aString,							// String
		_T("."),							// Delimiters
		_T(""))								// OmitChars
	if (CALL_TO_BIF_FAILED(strsplit_result))
		return NULL;	// This is probably out of memory, strsplit already displayed the error.

	Array* arr = (Array*)TokenToObject(strsplit_result);	// Array

	int count = arr->Length();			// Numbers elements in the array.
	
	ScriptModule* found_module = this;	// A resolved module in which to conduct the search
	ScriptModule* tmp_module = NULL;	// Used to not overwrite found_module and detect errors after the loop

	// For each item in array:
	ARRAY_FOR_EACH(arr, i, item) 
	{
		LPTSTR name = item.marker; // item i
		if (tmp_module = found_module->GetNestedModule(name, true))
		{
			// It is a module, the next item in the array will be searched for in this module.
			found_module = tmp_module;
			continue;
		}
		arr->Release();
		return NULL; // Not found
	}
	arr->Release();
	return found_module;
}

#include "ScriptModulesAddNames.h"
#include "ScriptModulesReleaseObjects.h"
#include "ScriptModulesNested.h"

#ifndef AUTOHOTKEYSC
#include "ScriptModulesSourceFileIndex.h"
#endif 

ScriptModule* ScriptModule::GetReservedModule(LPTSTR aName, ScriptModule* aSource /* = NULL */) // static public
{
	// aName, the name of the module to return.
	// aSource, if aName doesn't match any of the "standard/default" module, find a module relative to aSource, for example aSource outer module. Can be NULL (omitted).
	// Get one of the module whose name are reserved. See ScriptModulesDefines.h
	if (SMODULES_NAMES_MATCH(aName, SMODULES_STANDARD_MODULE_NAME))
		return g_StandardModule;
	if (SMODULES_NAMES_MATCH(aName, SMODULES_DEFAULT_MODULE_NAME))
		return g_script.mModuleSimpleList.GetItem(0);

	// Handle relative names:
	if (!aSource)
		return NULL;
	if (SMODULES_NAMES_MATCH(aName, SMODULES_OUTER_MODULE_NAME))
		return aSource->mOuter;
	return NULL;
}

Object* ScriptModule::FindClassFromDotDelimitedString(LPTSTR aString)
{
	// aString, a dot (.) delitmited string. Relies on that each class name is verified elsewhere.
	
	// Call BIF_StrSplit:
	// Array: = StrSplit(String, Delimiters, OmitChars, MaxParts : = -1)
	CALL_BIF(StrSplit, strsplit_result,
		aString,							// String
		_T("."),							// Delimiters
		_T(""))								// OmitChars
	if (CALL_TO_BIF_FAILED(strsplit_result))
		return NULL;	// This is probably out of memory, strsplit already displayed the error, 
						// there will be an "Unknown class" error too but this is too rare to worry about.

	Array* arr = (Array*)TokenToObject(strsplit_result);	// Array
	
	int count = arr->Length();			// Numbers elements in the array.
	if (count == 1)
	{
		arr->Release();
		return g_script.FindClass(aString, 0, this);
	}

	ScriptModule* found_module = this;	// A resolved module in which to conduct the search
	ScriptModule* tmp_module = NULL;	// Used to not overwrite found_module and detect errors after the loop
	Object* class_object = NULL;		// The object to return

	// For each item in array:
	ARRAY_FOR_EACH(arr, i, item)
	{
		LPTSTR name = item.marker; // Item i
		size_t name_length = item.marker_length;
		
		if (tmp_module = found_module->GetNestedModule(name, true))
		{
			// It is a module, the next item in the array will be searched for in this module.
			found_module = tmp_module;
			continue;
		}
		// Here, the module is completely resolved. Find the class object within it.
		// Example, aString = "myMod.myOtherMod.MyClass.MyNestedClass", extract
		// class_name = "MyClass.MyNestedClass" and pass it to FindClass.
		LPTSTR class_name;
		
		// Call BIF_InStr to find the i:th dot. At this stage i must be at least 1.
		// FoundPos := InStr(Haystack, Needle , CaseSensitive := false, StartingPos := 1, Occurrence := 1)
		CALL_BIF(InStr, instr_result,
			aString,		// Haystack
			_T("."),		// Needle
			true,			// CaseSensitive
			1,				// StartingPos
			(__int64)i);	// Occurrence
		if (CALL_TO_BIF_FAILED(instr_result)) // Check this for maintainability.
		{
			arr->Release();
			return NULL;
		}
		__int64 FoundPos = instr_result.value_int64;
		class_name = aString + FoundPos; // FoundPos is 1-based so no need to +1 to move past the dot (.)
		class_object = g_script.FindClass(class_name, 0, found_module);
		break; // Must not continue the loop!
	}
	if (tmp_module == found_module) // Only to catch bugs
	{
		// The last iteration was resolved to a module, eg, something like "class x extends MyModule.MyNestedModule"
		ASSERT(!class_object);
	}
	arr->Release();
	return class_object;
}
#include "ModuleListMethods.h"