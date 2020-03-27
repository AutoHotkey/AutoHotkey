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

bool ScriptModule::ValidateName(LPTSTR aName)
{
	// return true if this name can be used as a variable name in this module, i.e., it is not the same as one of its nested module names or any of the reserved names.
	// else returns false.
	// Relies on being called from Var::ValidateName.
	if SMODULES_NAMES_MATCH(aName, SMODULES_OUTER_MODULE_NAME) // Need to check since SMODULES_DEFAULT_MODULE_NAME's mOuter is NULL.
		return false;
	return !GetNestedModule(aName, true);
}

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

//
// Methods for "using" names - start
//

ResultType ScriptModule::AddObjectError(LPTSTR aErrorMsg, LPTSTR aExtraInfo, UseParams* aUp)
{
	// Used by the below methods for error reporting.
	g_script.mCurrLine = NULL;
	g_script.ScriptError(aErrorMsg, aExtraInfo);
	return FAIL;
}

bool ScriptModule::ResolveUseParams() 
{
	// Some names might not have been able to be imported when the relevant import directive was reached,
	// due to the script not being fully parsed. Try to reslove any such issues here, if not possible, the script will terminate.
	if (!mUseParams)
		return true; // No unresolved UseParams
	LPTSTR name = NULL; // For error reporting
	
	UseParams* up;
	int i = 0;
	while (up = mUseParams->GetItem(i)) // GetItem does the bound checks
	{
		switch (AddObject(up))
		{
		case CONDITION_TRUE:
			// Not added, no error message shown, show one.
			
			name = up->scope_symbol == SYM_MODULE
				? up->param1		// An item in the name list couldn't be resolved. (This case might not be possible now, but added for maintainability)
				: up->param2.str;	// The source could not be resolved.
			AddObjectError(ERR_SMODULES_UNRESOLVED_NAME, name);
			// fall through:
		case FAIL:
			// The error was either shown by AddObject or the above case if fell through.
			return false;
			//case OK:
		}
		// Added, continue with the next item:
		++i;
	}
	// Only clean up on success, the program should terminate anyways on failure.
	delete mUseParams;
	mUseParams = NULL;
	return true; // All succeeded.
}

// This "series" of methods should return CONDITION_TRUE if the UseParams are saved (and not added).
// OK if the params are fully handled. FAIL on any failure. Any saved UseParams will be resolved later, see ResolveUseParams.
ResultType ScriptModule::AddObject(LPTSTR aObjList, LPTSTR aModuleName, SymbolType aTypeSymbol)
{
	// Atempts to add each object in aObjList (comma delimited) to aModuleName.
	// The type of the objects are indicated by aTypeSymbol.
	UseParams *up = new UseParams(SYM_STRING, aTypeSymbol, aObjList, NULL, aModuleName);
	// Handle this later in ResolveUseParams, otherwise it might fail depending on the posistion of the directive. 
	// Save this UseParams so that it can be resolved later:
	if (!mUseParams)
		mUseParams = new UseParamsList;
	if (!mUseParams->AddItem(up))
		return FAIL; // The program should fail so no point in doing "delete up;" here. Unlikely so avoid error message, ERR_OUTOFMEM
	return CONDITION_TRUE;
}
ResultType ScriptModule::AddObject(UseParams* aObjs)
{
	ScriptModule* mod = NULL;
	switch (aObjs->scope_symbol) // resolve the source module
	{
	case SYM_STRING: 
		mod = FindModuleFromDotDelimitedString(aObjs->param2.str); 
		if (mod)
		{	// Save it to avoid searching for it again if cannot add the objects below.
			free(aObjs->param2.str);
			aObjs->param2.mod = mod;
			aObjs->scope_symbol = SYM_MODULE;
		}
		else
			return AddObjectError(ERR_SMODULES_NOT_FOUND, aObjs->param2.str, aObjs);
		break;
	case SYM_MODULE: mod = aObjs->param2.mod; break;
	}
	if (!mod)
		return FAIL;
	
	// Call BIF_StrSplit:
	// Array: = StrSplit(String, Delimiters, OmitChars, MaxParts : = -1)
	CALL_BIF(StrSplit, strsplit_result,
		aObjs->param1,							// String
		_T(","),								// Delimiters: comma
		_T(" \t"))								// OmitChars: spaces and tabs
	if (CALL_TO_BIF_FAILED(strsplit_result))
		return FAIL;	// This is probably out of memory, strsplit already displayed the error.
	Array* arr = (Array*)TokenToObject(strsplit_result);	// Array
	
	switch (aObjs->type_symbol)
	{
	case SYM_FUNC: return AddFuncFromList(arr, mod);
	case SYM_VAR: return AddVarFromList(arr, mod);
	
	}
	// This shouldn't be reached:
	return FAIL;
}
ResultType ScriptModule::AddFuncFromList(Array* aFuncList, ScriptModule* aModule)
{
	// Adds all functions in aFuncList if found in aModule, to this module
	if (aModule == GetReservedModule(SMODULES_STANDARD_MODULE_NAME)) // The standard module is not supported.
		return AddObjectError(ERR_SMODULES_NOT_SUPPORTED, SMODULES_STANDARD_MODULE_NAME);
	ARRAY_FOR_EACH(aFuncList, i, result)
	{
		if (!FindAndAddFunc(result.marker, aModule))
			return FAIL; // Message already shown.
	}
	return OK;
}

ResultType ScriptModule::FindAndAddFunc(LPTSTR aFuncName, ScriptModule* aModule)
{
	if (!_tcsicmp(aFuncName, SMODULES_IMPORT_NAME_ALL_MARKER))
		return AddAllFuncs(aModule);
	// Add only one func:
	Func* func = aModule->mFuncs.Find(aFuncName, NULL); // Find the func in the source module
	if (!func) 
		// This func doesn't exist
		return AddObjectError(ERR_SMODULES_FUNC_NOT_FOUND, aFuncName);
	
	int insert_pos;
	Func* this_func = mFuncs.Find(aFuncName, &insert_pos);	// search for a function with the same name in this module
	if (this_func)
	{
		if (this_func == func) // This exact func already exists
			return OK;
		// A different function with the same name already exists in this module
		return AddObjectError(ERR_DUPLICATE_FUNCTION, aFuncName);
	}
	if (!mFuncs.Insert(func, insert_pos)) // Now add it to this module
		return AddObjectError(ERR_OUTOFMEM);
	return OK;
}

ResultType ScriptModule::AddAllFuncs(ScriptModule* aModule)
{
	auto &funcs = aModule->mFuncs.mItem;
	size_t func_count = aModule->mFuncs.mCount;
	LPCTSTR name;
	for (size_t i = 0; i < func_count; ++i) // Visit each func in the source module's list of functions.
	{
		name = funcs[i]->mName;
		if (_tcschr(name, '.')) // Avoid adding class methods such as "cls.prototype.method".
			continue;
		if (!FindAndAddFunc((LPTSTR)name, aModule))
			return FAIL; // The error message has already been displayed.
	}
	return OK;
}

ResultType ScriptModule::AddVarFromList(Array *aVarList, ScriptModule *aModule)
{
	ARRAY_FOR_EACH(aVarList, i, result)
	{
		if (!FindAndAddVar(result.marker, (int)result.marker_length, aModule))
			return FAIL;
	}
	return OK;

}
ResultType ScriptModule::FindAndAddVar(LPTSTR aVarName, int aNameLength, ScriptModule *aModule)
{
	if (!_tcsicmp(aVarName, SMODULES_IMPORT_NAME_ALL_MARKER))
		return AddAllVars(aModule);
	// Add aVarName to this modules var list, first find it in aModule:
	Var* var = g_script.FindVar(aVarName, aNameLength, NULL, VAR_GLOBAL, false, aModule);
	if (!var)
		return AddObjectError(ERR_SMODULES_VAR_NOT_FOUND, aVarName);
		
	
	return AddVar(var); // Now add it to this module
}
ResultType ScriptModule::AddAllVars(ScriptModule *aModule) 
{
	// Adds all super-global vars in aModule to this module.
	Var	**vars, *var;
	int var_count;
	for (int i = 0; i < 2; ++i) // Two iterations, one for each var list.
	{
		vars = i ? aModule->mVar : aModule->mLazyVar;
		var_count = i ? aModule->mVarCount : aModule->mLazyVarCount; // Perhaps unlikely but could occur if a module contains some massive list of constants.
		for (int j = 0; j < var_count; ++j)	// one iteration for each var in the var list.
		{
			var = vars[j];
			if (var->IsSuperGlobal())
			{
				if (!AddVar(var))
					return FAIL;
			}
		}
	}
	return OK; // The above need not have added any vars, but that is not an error.
}

ResultType ScriptModule::AddVar(Var* aVar)
{	
	// aVar must not be NULL.
	int pos;
	Var* var = g_script.FindVar(aVar->mName, 0, &pos, FINDVAR_GLOBAL, false, this); // Search for a variable with this name first
	if (var && var != aVar)
	{
		// Another variable with this name already exists,
		// This can happen if this module have declared this variable as super global. That should be an error.
		// But it can also happen if the variable is declared global in a function.
		if (var->IsSuperGlobal())
			return AddObjectError(ERR_SMODULES_BAD_DECLARATION);
		// Replace var with aVar.
		ReplaceGlobalVar(var, aVar);
		return OK;
	}
	if (var) // This (exact) var already exists in this module. Note "var != aVar" above.
		return OK;
	// Insert the var in this modules list:
	return g_script.AddVar(NULL, 0, pos, VAR_GLOBAL, this, aVar) ? OK : FAIL;
}

ResultType ScriptModule::ReplaceGlobalVar(Var* aVar1, Var* aVar2)
{
	// replaces aVar1 with aVar2, marks aVar1 as replaced. 
	if (ReplaceVar(mVar, mVarCount, aVar1, aVar2)
		|| ReplaceVar(mLazyVar, mLazyVarCount, aVar1, aVar2))
		return OK;
	return FAIL;
}

ResultType ScriptModule::ReplaceVar(Var** aVars, int aVarCount, Var* aVar1, Var* aVar2)
{
	// Finds aVar1 and puts aVar2 in its place. Marks aVar1 as replaced (it shouldn't be used after this.)
	// Code taken from Script::FindVar.
	// Init for binary search loop:
	int left, right, mid, result;  // left/right must be ints to allow them to go negative and detect underflow.
	Var** var;  // An array of pointers-to-var.

	LPTSTR var_name = aVar1->mName;
	var = aVars;
	right = aVarCount - 1;
	

	// Binary search:
	for (left = 0; left <= right;) // "right" was already initialized above.
	{
		mid = (left + right) / 2;
		result = _tcsicmp(var_name, var[mid]->mName); // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales; 3) performance.
		if (result > 0)
			left = mid + 1;
		else if (result < 0)
			right = mid - 1;
		else
		{
			if (var[mid] != aVar1) // This shouldn't happen
				return FAIL;
			var[mid] = aVar2;
			aVar1->MarkReplaced();
			return OK;
		}
	}
	return FAIL;
}

//
// Methods for "using" names - end
//



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
	// frees the list of source file indices for this module.
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
		if (found = ScriptModule::GetReservedModule(aModuleName, this))
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
		return g_script.mModuleSimpleList.GetItem(1);
	if (SMODULES_NAMES_MATCH(aName, SMODULES_DEFAULT_MODULE_NAME))
		return g_script.mModuleSimpleList.GetItem(0);

	// Handle relative names:
	if (!aSource)
		return NULL;
	if (SMODULES_NAMES_MATCH(aName, SMODULES_OUTER_MODULE_NAME))
		return aSource->mOuter;
	return NULL;
}

// Methods for releasing objects in vars:
void ScriptModule::ReleaseVarObjects(Var** aVar, int aVarCount) // private
{
	for (int i = 0; i < aVarCount; ++i)
		if (aVar[i]->IsObject())
			aVar[i]->ReleaseObject(); // ReleaseObject() vs Free() for performance (though probably not important at this point).
		// Otherwise, maybe best not to free it in case an object's __Delete meta-function uses it?
	return;
}

void ScriptModule::ReleaseStaticVarObjects(Var** aVar, int aVarCount) // private
{
	for (int i = 0; i < aVarCount; ++i)
		if (aVar[i]->IsStatic() && aVar[i]->IsObject()) // For consistency, only free static vars (see below).
			aVar[i]->ReleaseObject();
}

void ScriptModule::ReleaseStaticVarObjects(FuncList& aFuncs) // private
{
	for (int i = 0; i < aFuncs.mCount; ++i)
	{
		if (aFuncs.mItem[i]->IsBuiltIn())
			continue;
		auto &f = *(UserFunc *)aFuncs.mItem[i];
		// Since it doesn't seem feasible to release all var backups created by recursive function
		// calls and all tokens in the 'stack' of each currently executing expression, currently
		// only static and global variables are released.  It seems best for consistency to also
		// avoid releasing top-level non-static local variables (i.e. which aren't in var backups).
		ReleaseStaticVarObjects(f.mVar, f.mVarCount);
		ReleaseStaticVarObjects(f.mLazyVar, f.mLazyVarCount);
		if (f.mFuncs.mCount)
			ReleaseStaticVarObjects(f.mFuncs);
	}
}

void ScriptModule::ReleaseVarObjects()	// public
{
	ReleaseVarObjects(mVar, mVarCount);
	ReleaseVarObjects(mLazyVar, mLazyVarCount);
	ReleaseStaticVarObjects(mFuncs);
	if (mNested)
		mNested->ReleaseVarObjects(); // release nested last, it seems more likely that the outer module refers to the nested ones, than the other way around.
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

#undef MODULELIST_INITIAL_SIZE
#undef MODULELIST_SIZE_GROW

void ModuleList::ReleaseVarObjects() // public
{
	for (size_t i = 0; i < mCount; ++i)
		mList[i]->ReleaseVarObjects(); // this releases the vars in the nested modules too.
}