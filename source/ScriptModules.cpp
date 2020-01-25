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

#include "stdafx.h" // pre-compiled headers
#include "ScriptModules.h"

const LPTSTR ScriptModule::sUnamedModuleName = SMODULES_UNNAMED_NAME;

void ScriptModule::RemoveLastModule()
{
	// caller has ensured mNested must exist.
	mNested->RemoveLastModule();
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

void ScriptModule::FreeOptionalModuleList()
{
	// frees the list of optional modules for this module and all of its nested ones.
	if (mNested)
		mNested->FreeOptionalModuleList();
	if (mOptionalModules)
		delete mOptionalModules;
	mOptionalModules = NULL;
}

#ifndef AUTOHOTKEYSC
bool ScriptModule::AddSourceFileIndex(int aIndex)
{
	// Adds an index to the list of files which was included in this module.
	// Each index in the list corresponds to a source file in Line::sSourceFile
	// returns false on out of memory, else true.

	if (!mSourceFileIndexList)
		mSourceFileIndexList = new SimpleList<int>;
	return mSourceFileIndexList->AddItem(aIndex);

}

void ScriptModule::FreeSourceFileIndexList()
{
	// frees the list of source file indices for this module and all of its nested ones.
	if (mNested)
		mNested->FreeSourceFileIndexList();
	if (mSourceFileIndexList)
		delete mSourceFileIndexList;
	mSourceFileIndexList = NULL;
}

bool ScriptModule::HasIncludedSourceFile(TCHAR aPath[])
{
	// Finds if aPath was #include:ed in this module.
	if (!mSourceFileIndexList)
		return false;		// no list, not included.
	int index = 0;			// iteration index for GetItem (which does the bound checks)
	int source_index;		// result from get item
	bool was_found;			// break loop criteria.
	// Note:
	// mSourceFileIndexList only contains indicies which are valid for Line::sSourceFile, so the below is safe.
	while (1)
	{
		source_index = mSourceFileIndexList->GetItem(index++, &was_found);
		if (!was_found) // means no more items in the list to check.
			return false;
		if (!lstrcmpi(Line::sSourceFile[source_index], aPath)) // Case insensitive like the file system (testing shows that "Ä" == "ä" in the NTFS, which is hopefully how lstrcmpi works regardless of locale).
			return true;
	}
}
#endif // #ifndef AUTOHOTKEYSC


ScriptModule* ScriptModule::InsertNestedModule(LPTSTR aName, int aFuncsInitSize, ScriptModule* aOuter) // public
{
	// Creates a new ScriptModule and inserts it in the this ScriptModule's mNested list.
	// Duplicate module names are detected and causes NULL to be returned. Otherwise the new ScriptModule is returned.
	// NULL might also be returned in case of some failed memory allocation

	// aName, the name of the new module to create and insert.
	// aFuncsInitSize, the number of functions to allocate space for in the module's FuncList mFuncs.
	// aOuter, the enclosing module. This will mostlikey always be g_CurrentModule unless g_CurrentModule is one of the top level modules.

	// Since using SimpleHeap::Malloc for allocation of modules, callers of this function should avoid
	// making "many" calls where InsertNestedModule below fails due to duplicate name. The only current caller, DefineModule
	// causes the program to end on failure.
	ScriptModule* new_module = new ScriptModule(aName, aFuncsInitSize, aOuter);
	if (!new_module) // Check since `new` is overloaded to use SimpleHeap::Malloc
		return NULL;
	if (InsertNestedModule(new_module))
		return new_module;
	delete new_module;
	return NULL;
}

bool ScriptModule::InsertNestedModule(ScriptModule *aModule) // public
{
	// aModule, the module to insert in this module's mNested list.
	// Returns false on out of memory or if this module was already added. Otherwise true.
	if (!mNested)	// If mNested doesn't exist, it is created.
		if (!(mNested = new ModuleList())) // Check since `new` is overloaded to use SimpleHeap::Malloc
			return false;
	return mNested->Add(aModule);
}

ScriptModule* ScriptModule::GetNestedModule(LPTSTR aModuleName, bool aAllowReserved /*= false*/) // public
{
	ScriptModule* found;
	if (aAllowReserved)
		if (found = GetReservedModule(aModuleName, this))
			return found;
	// not a reserved module name or not looking for one, search nested modules, if any.
	if (!mNested)
		return NULL; // no nested
	mNested->find(aModuleName, &found);
	return found;
}

ScriptModule* ScriptModule::GetReservedModule(LPTSTR aName, ScriptModule* aSource /* = NULL */) // static public
{
	// aName, the name of the module to return.
	// aSource, if aName doesn't match any of the "standard/default" module, find a module relative to aSource, for example aSource outer module. Can be NULL (omitted).
	// Get one of the module whose name are reserved. See ScriptModulesDefines.h
	if (SMODULES_NAMES_MATCH(aName, SMODULES_STANDARD_MODULE_NAME))
		return NULL;
	if (SMODULES_NAMES_MATCH(aName, SMODULES_DEFAULT_MODULE_NAME))
		return NULL;

	// Handle relative names:
	if (!aSource)
		return NULL;
	if (SMODULES_NAMES_MATCH(aName, SMODULES_OUTER_MODULE_NAME))
		return aSource->mOuter;
	return NULL;
}

//
//	ModuleList methods.
//
#define MODULELIST_INITIAL_SIZE (5)
#define MODULELIST_SIZE_GROW (5)
bool ModuleList::Add(ScriptModule* aModule)
{
	// Adds aModule to mList
	// return true on succsess, else false
	if (!aModule)
		return false;
	if (IsInList(aModule))	// module names must be unique
		return false;
	if (mCount == mListSize) // Expand if needed
	{
		size_t new_size = mListSize ? MODULELIST_INITIAL_SIZE : mListSize + MODULELIST_SIZE_GROW;
		ScriptModule** new_list = (ScriptModule**)realloc(mList, sizeof(ScriptModule*) * new_size);
		if (!new_list)
			return false; // out of memory
		mListSize = new_size;
		mList = new_list;
	}
	// Add module and increment count
	mList[mCount] = aModule;
	mCount++;
	return true;
}


bool ModuleList::IsInList(ScriptModule* aModule)
{
	return find(aModule->GetName());
}

void ModuleList::RemoveLastModule()
{
	// removes the last module.
	ASSERT(mList);
	delete mList[--mCount];
}

bool ModuleList::find(LPTSTR aName, ScriptModule **aFound)
{
	// Anonymous modules never match.
	for (size_t i = 0; i < mCount; ++i)
	{
		if (SMODULES_NAMES_MATCH(aName, mList[i]->GetName()))
		{
			if (aFound)
				*aFound = mList[i];
			return true;
		}
	}
	return false;
}

void ModuleList::FreeOptionalModuleList()
{
	// see the corresponding ScriptModule:: method of comments
	for (size_t i = 0; i < mCount; ++i)
		mList[i]->FreeOptionalModuleList();
}
#ifndef AUTOHOTKEYSC
void ModuleList::FreeSourceFileIndexList()
{
	// see the corresponding ScriptModule:: method of comments
	for (size_t i = 0; i < mCount; ++i)
		mList[i]->FreeSourceFileIndexList();
}
#endif

#undef MODULELIST_INITIAL_SIZE
#undef MODULELIST_SIZE_GROW