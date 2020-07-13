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

// This file contains methods for handling importing of names between script modules.
// See for example SMODULES_IMPORT_VARS_DIRECTIVE_NAME in ScriptModulesDefines.h

// These methods should return:
// - CONDITION_TRUE, if the UseParams (the parameters of #UseXXX) were not added, but no error occured.
// - OK, if the params are fully handled. 
// - FAIL, on any failure.

#pragma once

ResultType ScriptModule::AddObjectError(LPTSTR aErrorMsg, LPTSTR aExtraInfo, UseParams* aUp)
{
	// Used by the below methods for error reporting.
	g_script.mCurrLine = NULL;
	g_script.ScriptError(aErrorMsg, aExtraInfo);
	return FAIL;
}

ResultType ScriptModule::AddObject(LPTSTR aObjList, LPTSTR aModuleName, SymbolType aTypeSymbol)
{
	// The actual additiaon of these objects is handled later by ResolveUseParams
	// otherwise it might fail depending on the posistion of the directive. 
	// For now just save the UseParams.
	UseParams* up = new UseParams(SYM_STRING, aTypeSymbol, aObjList, NULL, aModuleName);
	
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
			aObjs->scope_symbol = SYM_OBJECT;
		}
		else
			return AddObjectError(ERR_SMODULES_NOT_FOUND, aObjs->param2.str, aObjs);
		break;
	case SYM_OBJECT: mod = aObjs->param2.mod; break;
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
	if (aModule == g_StandardModule) // The standard module is not supported.
		return AddObjectError(ERR_SMODULES_NOT_SUPPORTED, SMODULES_STANDARD_MODULE_NAME);
	ARRAY_FOR_EACH(aFuncList, i, result)
	{
		if (!FindAndAddFunc(result.marker, aModule))
			return FAIL; // Message already shown.
	}
	return OK;
}

ResultType ScriptModule::FindAndAddFunc(LPTSTR aFuncName, ScriptModule* aModule, bool aAllowConflict)
{
	
	// aAllowConflict should be true if an explicit function definition takes precedence
	// over a declaration. Currently, aAllowConflict is only true when importing "*all" funcs.
	// It is difficult to imagine another case where passing true would make sense.

	if (!_tcsicmp(aFuncName, SMODULES_IMPORT_NAME_ALL_MARKER))
		return AddAllFuncs(aModule);
	// From this point, handle only one function.
	
	auto current_func = mCurrentUseParam->current_func; // might be nullptr.
	
	// First, check for a conflicting function definition 
	// in the list of expicitly defined functions.
	{
		auto& funcs = current_func ? current_func->mFuncs : mFuncs;
		int left;
		auto defined_func = funcs.Find(aFuncName, &left);
		if (defined_func)
		{	// There already exist an explicitly defined function with this name in this scope.
			if (!aAllowConflict)
				return AddObjectError(ERR_DUPLICATE_DECLARATION, aFuncName);
			else
				// If it ever becomes relevant, note that aFuncName is not guaranteed,
				// by the logic of this method, to exist in aModule at this return site. 
				// However, when adding this code, this isn't relevant, 
				// but acutally guaranteed by the current use of this method.
				return OK;
		}
	}
	
	auto func = aModule->mFuncs.Find(aFuncName, NULL); // Find the func in the source module
	if (!func)
		// This func doesn't exist
		return AddObjectError(ERR_SMODULES_FUNC_NOT_FOUND, aFuncName);

	int insert_pos;
	auto &used_funcs = current_func ? current_func->mUsedFuncs : mUsedFuncs;
	auto used_func = used_funcs.Find(aFuncName, &insert_pos);	// search for a function with the same name in the "used scope"
	
	if (used_func)
	{
		if ( used_func == func	 // This exact func already exists in this scope.
								 // Seems best to allow this in case, for example, multiple includes in the same scope
								 // uses the same library and imports funcs from it.
			 || aAllowConflict ) // Or conflicts are allowed. This can happen if a scope imports all names
								 // names from two modules which each defines a function named aFuncName.
								 // This makes the declaration positional as the first import will take precedence
								 // but it also makes scripts more maintainable so seems worth it.
								 // If conflicts are known by script the author,
								 // it can make explicit declarations before importing all.
			return OK;
		// A different function with the same name already exists in this scope
		return AddObjectError(ERR_DUPLICATE_DECLARATION, aFuncName);
	}
	if (!used_funcs.Insert(func, insert_pos)) // Now add it to this scope
		return AddObjectError(ERR_OUTOFMEM);
	return OK;
}

ResultType ScriptModule::AddAllFuncs(ScriptModule* aModule)
{
	auto& funcs = aModule->mFuncs.mItem;
	size_t func_count = aModule->mFuncs.mCount;
	LPCTSTR name;
	for (size_t i = 0; i < func_count; ++i) // Visit each func in the source module's list of functions.
	{
		name = funcs[i]->mName;
		if (_tcschr(name, '.')) // Avoid adding class methods such as "cls.prototype.method".
			continue;
		if (!FindAndAddFunc((LPTSTR)name, aModule, /* aAllowConflict = */ true))
			return FAIL; // The error message has already been displayed.
	}
	return OK;
}

ResultType ScriptModule::AddVarFromList(Array* aVarList, ScriptModule* aModule)
{
	ARRAY_FOR_EACH(aVarList, i, result)
	{
		if (!FindAndAddVar(result.marker, (int)result.marker_length, aModule))
			return FAIL;
	}
	return OK;

}
ResultType ScriptModule::FindAndAddVar(LPTSTR aVarName, int aNameLength, ScriptModule* aModule)
{
	if (!_tcsicmp(aVarName, SMODULES_IMPORT_NAME_ALL_MARKER))
		return AddAllVars(aModule);
	// Add aVarName to this modules var list, first find it in aModule:
	Var* var = g_script.FindVar(aVarName, aNameLength, NULL, VAR_GLOBAL, false, aModule);
	if (!var)
		return AddObjectError(ERR_SMODULES_VAR_NOT_FOUND, aVarName);

	return AddVar(var); // Now add it to this module
}
ResultType ScriptModule::AddAllVars(ScriptModule* aModule)
{
	// Adds all super-global vars in aModule to this module.
	Var** vars, * var;
	int var_count;
	for (int i = 0; i < 2; ++i) // Two iterations, one for each var list.
	{
		// The lazy list is unlikely to contain anything, but could be non-empty if a module contains some massive list of constants.
		vars = i ? aModule->mVar : aModule->mLazyVar;
		var_count = i ? aModule->mVarCount : aModule->mLazyVarCount;	
		for (int j = 0; j < var_count; ++j)
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

bool ScriptModule::ResolveUseParams()
{
	// This method should be called only when the script has been fully parsed.
	// Names are not added when the relevant #UseXXX directive was reached, due to the script not being fully parsed.
	// Try to reslove these now, if not possible, the script will terminate.
	if (!mUseParams)
		return true; // No unresolved UseParams
	LPTSTR name = NULL; // For error reporting

	UseParams* up;
	int i = 0;
	while (up = mUseParams->GetItem(i)) // GetItem does the bound checks
	{
		mCurrentUseParam = up;
		switch (AddObject(up))
		{
		case CONDITION_TRUE:
			// Not added, no error message shown, show one.

			name = up->scope_symbol == SYM_OBJECT
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
