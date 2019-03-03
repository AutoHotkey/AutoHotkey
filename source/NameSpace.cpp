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
#include "NameSpace.h"

const LPTSTR NameSpace::sAnonymousNameSpaceName = NAMESPACE_ANONYMOUS_NAME;

ResultType NameSpace::AutoExecSection()
// Returns FAIL if can't run due to critical error.  Otherwise returns OK.
{
	ResultType nested_result;
	if (mNestedNameSpaces)
	{
		if ((nested_result = mNestedNameSpaces->AutoExecSection()) != OK)
			return nested_result;
	}
	// At this point all nested namespaces of this namespace, and their nested namespaces has executed their respective auto execute section, and completed their static var initializers.
	 
	// Handle this namespace:
	SetCurrentNameSpace();	
	DEBUGGER_STACK_PUSH(_T("Auto-execute"))	// Only after the current namespace has been set.
	// v2: Ensure the Hotkey command defaults to no criterion rather than the last #IfWin.  
	mSettings.HotCriterion = NULL;
	
	ResultType ExecUntil_result;
	if (!mFirstLine) // In case it's ever possible to be empty.
		ExecUntil_result = OK;
	// And continue on to do normal exit routine so that the right ExitCode is returned by the program.
	else
	{
		// The below also sets g->ThreadStartTime and g->UninterruptibleDuration.  Notes about this:
		// In case the AutoExecute section takes a long time (or never completes), allow interruptions
		// such as hotkeys and timed subroutines after a short time. Use g->AllowThreadToBeInterrupted
		// vs. g_AllowInterruption in case commands in the AutoExecute section need exclusive use of
		// g_AllowInterruption (i.e. they might change its value to false and then back to true,
		// which would interfere with our use of that var).

		// In comment below g is now t. References to autoexec timer are obsolete.
		// UPDATE v1.0.48: g->ThreadStartTime and g->UninterruptibleDuration were added so that IsInterruptible()
		// won't make the AutoExec section interruptible prematurely.  In prior versions, KILL_AUTOEXEC_TIMER() did this,
		// but with the new IsInterruptible() function, doing it in KILL_AUTOEXEC_TIMER() wouldn't be reliable because
		// it might already have been done by IsInterruptible() [or vice versa], which might provide a window of
		// opportunity in which any use of Critical by the AutoExec section would be undone by the second timeout.
		// More info: Since AutoExecSection() never calls InitNewThread(), it never used to set the uninterruptible
		// timer.  Instead, it had its own timer.  But now that IsInterruptible() checks for the timeout of
		// "Thread Interrupt", AutoExec might become interruptible prematurely unless it uses the new method below.
		
		// Set these so that all auto execute sections have an (as) equal (as possible) chance to do their thing.
		t->AllowThreadToBeInterrupted = false;
		t->ThreadStartTime = GetTickCount();
		t->UninterruptibleDuration = AUTO_EXEC_UNINTERRUPTIBLE_DURATION;
		t->UninterruptibleDuration = 0; // So that ExecUntil below calls SetThreadCriticalStatus. Better to do it this way for maintainability.

		// v1.0.25: This is now done here, closer to the actual execution of the first line in the (edit: namespace rather than script),
		// to avoid an unnecessary Sleep(10) that would otherwise occur in ExecUntil:
		g_script.mLastPeekTime = GetTickCount();
		ExecUntil_result = mFirstLine->ExecUntil(UNTIL_RETURN); // Might never return (e.g. infinite loop or ExitApp).
		
	}	

	// REMEMBER: The ExecUntil() call above will never return if the AutoExec section never finishes
	// (e.g. infinite loop) or it uses Exit/ExitApp.

	// It seems best to set ErrorLevel to NONE after the auto-execute part of the script is done.
	// However, it isn't set to NONE right before launching each new thread (e.g. hotkey subroutine)
	// because it's more flexible that way (i.e. the user may want one hotkey subroutine to use the value
	// of ErrorLevel set by another).  This reset was also done by LoadFromFile(), but it is done again
	// here in case the auto-execute section changed it:
	g_ErrorLevel->Assign(ERRORLEVEL_NONE);
	DEBUGGER_STACK_POP()
	return ExecUntil_result;
}

Var *NameSpace::FindVar(LPTSTR aVarName, size_t aVarNameLength, int *apInsertPos, int aScope, bool *apIsLocal)
{
	NameSpace *prev_namespace = g_CurrentNameSpace;	// save current namespace
	SetCurrentNameSpace();	// Could assign to g_CurrentNameSpace but done this way for maintainability
	Var *found_var = g_script.FindVar(aVarName, aVarNameLength, apInsertPos, aScope, apIsLocal);
	prev_namespace->SetCurrentNameSpace(); // restore namespace
	return found_var;
}

Var *NameSpace::FindOrAddVar(LPTSTR aVarName, size_t aVarNameLength /* = 0 */, int aScope /*= FINDVAR_DEFAULT*/) // public
{
	NameSpace *prev_namespace = g_CurrentNameSpace;	// save current namespace
	SetCurrentNameSpace();	// Could assign to g_CurrentNameSpace but done this way for maintainability
	Var *found_var = g_script.FindOrAddVar(aVarName, aVarNameLength, aScope);
	prev_namespace->SetCurrentNameSpace(); // restore namespace
	return found_var;
}

Func *NameSpace::FindFunc(LPCTSTR aFuncName, size_t aFuncNameLength /*= -1*/, int *apInsertPos /*= NULL*/, bool aAllowNested /*= true*/)  // public
{
	NameSpace *prev_namespace = g_CurrentNameSpace;	// save current namespace
	SetCurrentNameSpace();	// cannot just assign g_CurrentNameSpace.
	Func *found_func = g_script.FindFunc(aFuncName, aFuncNameLength, apInsertPos, aAllowNested);
	prev_namespace->SetCurrentNameSpace(); // restore namespace
	return found_func;
}

// End Script methods

Var * NameSpace::FindVarFromScopeSymbolDelimitedString(LPTSTR aString, bool aAllowAddVar, NameSpace **aFoundNameSpace)
{
	// aString, a OPERATOR_SCOPE_SYMBOL delitmited string. Caller is responsible for aString containing at least one delimiter. 
	//			White space is trimmed but any other invalid symbols are not detected, see next param for exception.
	// aAllowAddVar, set to true to indicate that if the variable isn't found, it can be added. The var name is validated only if aAllowAddVar is true.
	// aFoundNameSpace, the namespace in which var was searched for. If *aFoundNameSpace is NULL when this function returns the namespace(s) doesn't exist.
	
	// Call BIF_StrSplit to split the string and trim any spaces or tabs. Then use Object::ArrayToStrings.
	// Done this way for convenience, it isn't the most efficient way imaginable.
	// Array: = StrSplit(String, Delimiters, OmitChars, MaxParts : = -1)
	CALL_BIF(StrSplit, result_token,
		aString,							// String
		OPERATOR_SCOPE_SYMBOL,				// Delimiters
		_T(" \t"))							// OmitChars
	if (CALL_TO_BIF_FAILED(result_token))
		return NULL; // This is probably out of memory, currently this is not handled but the program should soon fail somewhere else if such a small memory commitment fails here.
	
	Object *arr = (Object *)TokenToObject(result_token);	// Array
	
	Var* var = NULL;
	int count = arr->GetNumericItemCount();
	LPTSTR *strings = (LPTSTR *)alloca(count * sizeof LPTSTR);
	NameSpace *found_namespace = this;
	if (arr->ArrayToStrings(strings, count, count)) // should never fail but better safe than sorry.
	{
		for (int i = 0; i < count - 1; ++i) // the last string is the variable name, hence the "-1"
			if (!(found_namespace = found_namespace->GetNestedNameSpace(strings[i], true)))
				break;
		LPTSTR var_name = strings[count - 1]; // AddVar will validate the name but not FindVar.
		if (found_namespace)
		{
			var = aAllowAddVar ? found_namespace->FindOrAddVar(var_name, 0, FINDVAR_GLOBAL)			// can add
				: found_namespace->FindVar(var_name, 0, NULL, FINDVAR_GLOBAL, NULL);				// cannot add
			if (aFoundNameSpace)
				*aFoundNameSpace = found_namespace; // caller wants the found namespace
		}
	}
	arr->Release();
	return var; // Can be NULL
}

NameSpace::~NameSpace()
{
	if (mNestedNameSpaces)
		delete mNestedNameSpaces;
	if (mOptionalNameSpaces)
		delete mOptionalNameSpaces;
#ifndef AUTOHOTKEYSC
	if (mSourceFileIndexList)
		delete mSourceFileIndexList;
#endif
}

bool NameSpace::AddOptionalNameSpace(LPTSTR aName)
{
	// adds aName to a list of namespace names which will not cause load time error if not found.
	// to allow "#import *i". Search word: NAMESPACE_INCLUDE_DIRECTIVE_NAME
	// returns false on out of memory, else true.

	if (!mOptionalNameSpaces)
		mOptionalNameSpaces = new NameSpaceNameList;
	LPTSTR name = _tcsdup(aName);
	if (!name)
		return false;
	if (!mOptionalNameSpaces->AddItem(name))
	{
		free(name);
		return false;
	}
	return true;
}

bool NameSpace::IsOptionalNameSpace(LPTSTR aName)
{
	// returns true if aName is in the list of optional namespaces.
	if (!mOptionalNameSpaces)
		return false;
	return mOptionalNameSpaces->HasItem(aName);
}

void NameSpace::FreeOptionalNameSpaceList() 
{
	// frees the list of optional namespaces for this namespace and all of its nested ones.
	if (mNestedNameSpaces)
		mNestedNameSpaces->FreeOptionalNameSpaceList();
	if (mOptionalNameSpaces)
		delete mOptionalNameSpaces;
	mOptionalNameSpaces = NULL;
}
#ifndef AUTOHOTKEYSC
bool NameSpace::AddSourceFileIndex(int aIndex)
{
	// Adds an index to the list of files which was included in this namespace.
	// Each index in the list corresponds to a source file in Line::sSourceFile
	// returns false on out of memory, else true.

	if (!mSourceFileIndexList)
		mSourceFileIndexList = new SimpleList<int>;
	return mSourceFileIndexList->AddItem(aIndex);
	
}

void NameSpace::FreeSourceFileIndexList()
{
	// frees the list of source file indices for this namespace and all of its nested ones.
	if (mNestedNameSpaces)
		mNestedNameSpaces->FreeSourceFileIndexList();
	if (mSourceFileIndexList)
		delete mSourceFileIndexList;
	mSourceFileIndexList = NULL;
}

bool NameSpace::HasIncludedSourceFile(TCHAR aPath[])
{
	// Finds if aPath was #include:ed in this namespace.
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

NameSpace *NameSpace::GetNestedNameSpace(LPTSTR aNameSpaceName, bool aAllowReserved /*= false*/) // public
{
	NameSpace *found;
	if (aAllowReserved)
		if (found = GetReservedNameSpace(aNameSpaceName, this))
			return found;
	// not a reserved namespace or not looking for one, search nested namespaces, if any.
	if (!mNestedNameSpaces)
		return NULL; // no nested
	mNestedNameSpaces->find(aNameSpaceName, &found); 
	return found;
}

NameSpace *NameSpace::GetReservedNameSpace(LPTSTR aName, NameSpace* aSource /* = NULL */) // static public
{
	// aName, the name of the namespace to return.
	// aSource, if aName doesn't match any of the "standard/default" namespaces, find a namespace relative to aSource, for example aSource outer namespace. Can be NULL (omitted).
	// Get one of the namespaces whose name are reserved. See NameSpaceDefines.h
	if (NAMESPACE_NAMES_MATCH(aName, NAMESPACE_STANDARD_NAMESPACE_NAME))
		return g_StandardNameSpace;
	if (NAMESPACE_NAMES_MATCH(aName, NAMESPACE_DEFAULT_NAMESPACE_NAME))
		return g_DefaultNameSpace;
	
	// Handle relative names:
	if (!aSource)
		return NULL;
	if (NAMESPACE_NAMES_MATCH(aName, NAMESPACE_OUTER_NAMESPACE_NAME))
		return aSource->mOuter;
	if (NAMESPACE_NAMES_MATCH(aName, NAMESPACE_TOP_NAMESPACE_NAME))
		return aSource->GetTopNameSpace();
	return NULL;
}

bool NameSpace::SetCurrentNameSpace()	// public
{
	// The caller is responsible for restoring the previous namespace if needed.
	if (this != g_CurrentNameSpace)
	{
		g_CurrentNameSpace = this;					// set this to the current namespace.
		g = &mSettings;								// use the settings of this namespace.
		return true;	// to indicate namespace was changed.
	}
	return false;	// to indicate no restoring needed.
}

global_struct &NameSpace::GetSettings()
{
	// This is needed by HOT_IF_WIN
	return mSettings;
}

NameSpace * NameSpace::GetTopNameSpace()
{
	// Finds the top namespace for this namespace.
	NameSpace *top = this; // this function can never return "this" namespace.
	while ((top = top->GetOuterNameSpace()) && !top->mIsTopNameSpace);
	return top;
}

NameSpace * NameSpace::GetOuterNameSpace()
{
	// Returns the outer namespace, this can be NULL. 
	// g_DefaultNameSpace and g_StandardNameSpace does not have any outer namespace, this is relies upon at least in Debugger::context_get.
	return mOuter;
}

// Func methods, public:

Func* NameSpace::FuncFind(LPCTSTR aName, int *apInsertPos)
{
	return mFuncs.Find(aName, apInsertPos);
}
Func* NameSpace::FuncFind(LPCTSTR aName, size_t aNameLength, int *apInsertPos, bool aAllowNested)
{
	Func *found_func;
	NameSpace *prev_namespace = g_CurrentNameSpace;
	SetCurrentNameSpace();									// Must be set for Script::FindFunc
	found_func = g_script.FindFunc(aName, aNameLength, apInsertPos, aAllowNested);
	prev_namespace->SetCurrentNameSpace();
	return found_func;
}

ResultType NameSpace::FuncInsert(Func *aFunc, int aInsertPos)
{
	return mFuncs.Insert(aFunc, aInsertPos);
}
ResultType NameSpace::FuncAlloc(int aAllocCount)
{
	return mFuncs.Alloc(aAllocCount);
}

FuncList &NameSpace::GetFuncs()
{
	return mFuncs;
}

// Var methods, public:
Var** &NameSpace::GetVar() 
{
	return mVar;
}
Var** &NameSpace::GetLazyVar()
{
	return mLazyVar;
}
int &NameSpace::GetVarCount()
{
	return mVarCount;
}
int &NameSpace::GetLazyVarCount()
{
	return mLazyVarCount;
}
int &NameSpace::GetVarCountMax()
{
	return mVarCountMax;
}

// Methods for releasing objects in vars:
void NameSpace::ReleaseVarObjects(Var** aVar, int aVarCount) // private
{
	for (int i = 0; i < aVarCount; ++i)
		if (aVar[i]->IsObject())
			aVar[i]->ReleaseObject(); // ReleaseObject() vs Free() for performance (though probably not important at this point).
		// Otherwise, maybe best not to free it in case an object's __Delete meta-function uses it?
	return;
}
void NameSpace::ReleaseStaticVarObjects(Var **aVar, int aVarCount) // private
{
	for (int i = 0; i < aVarCount; ++i)
		if (aVar[i]->IsStatic() && aVar[i]->IsObject()) // For consistency, only free static vars (see below).
			aVar[i]->ReleaseObject();
}
void NameSpace::ReleaseStaticVarObjects(FuncList &aFuncs) // private
{
	for (int i = 0; i < aFuncs.mCount; ++i)
	{
		Func &f = *aFuncs.mItem[i];
		if (f.mIsBuiltIn)
			continue;
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
void NameSpace::ReleaseVarObjects()	// public
{
	ReleaseVarObjects(mVar, mVarCount);
	ReleaseVarObjects(mLazyVar, mLazyVarCount);
	ReleaseStaticVarObjects(mFuncs);
	if (mNestedNameSpaces)
		mNestedNameSpaces->ReleaseVarObjects(); // release nested last, it seems more likely that the outer namespace refers to the nested ones, than the other way around.
}

Func * NameSpace::TokenToFunc(ExprTokenType & aToken, bool aAllowNested)
{
	// equivalent to calling ::TokenToFunc in the context of this namespace and restoring the previous namespace.
	Func *func;
	NameSpace *prev_namespace = g_CurrentNameSpace;
	SetCurrentNameSpace();
	// the following is similar to ::TokenToFunc.
	if (!(func = dynamic_cast<Func *>(TokenToObject(aToken))))
	{
		LPTSTR func_name = TokenToString(aToken);
		if (*func_name)
			func = FindFunc(func_name, -1, NULL, aAllowNested);
	}
	prev_namespace->SetCurrentNameSpace();
	return func;
}

NameSpace *NameSpace::InsertNestedNameSpace(LPTSTR aName, int aFuncsInitSize, NameSpace *aOuter, bool aIsTopNameSpace) // public
{
	// Creates a new namespace and inserts it in the this namespace' mNestedNameSpaces.
	// Duplicate namespace names are detected and causes NULL to be returned. Otherwise the new namespace is returned.
	// NULL might also be returned in case of some failed memory allocation

	// aName, the name of the new namespace to create and insert
	// aFuncsInitSize, the number of functions to allocate space for in the namespace' FuncList mFuncs.
	// aOuter, the enclosing namespace. This will mostlikey always be g_CurrentNameSpace unless g_CurrentNameSpace is one of the top level namespaces.
	
	// Since using SimpleHeap::Malloc for allocation of namespaces, callers of this function should avoid
	// making "many" calls where InsertNestedNameSpace below fails due to duplicate name. The only current caller, DefineNameSpace
	// causes the program to end on failure.
	NameSpace *new_namespace = new NameSpace(aName, aFuncsInitSize, aOuter, aIsTopNameSpace);
	if (InsertNestedNameSpace(new_namespace))
		return new_namespace;
	delete new_namespace;
	return NULL;
}	
bool NameSpace::InsertNestedNameSpace(NameSpace *aNameSpace) // public
{
	// aNameSpace, the namespace to insert in this namespace' mNestedNameSpaces.
	// Returns false on out of memory or if this namespace was already added. Otherwise true.
	if (!mNestedNameSpaces)	// If mNestedNameSpaces doesn't exist, it is created.
		if (!(mNestedNameSpaces = new NameSpaceList())) // Check since `new` is overloaded to use SimpleHeap::Malloc
			return false;
	return mNestedNameSpaces->Add(aNameSpace);
}

void NameSpace::RemoveLastNameSpace()
{
	// caller has ensured mNestedNameSpaces must exist.
	mNestedNameSpaces->RemoveLastNameSpace();
}

//
//	NameSpaceList methods.
//
#define NAMESPACELIST_INITIAL_SIZE (5)
#define NAMESPACELIST_SIZE_GROW (5)
bool NameSpaceList::Add(NameSpace *aNameSpace)
{
	// Adds aNameSpace to mList
	// return true on succsess, else false
	if (!aNameSpace)
		return false;
	if (IsInList(aNameSpace))	// Namespace names must be unique
		return false;
	if (mCount == mListSize) // Expand if needed
	{
		size_t new_size = mListSize ? NAMESPACELIST_INITIAL_SIZE : mListSize + NAMESPACELIST_SIZE_GROW;
		NameSpace **new_list = (NameSpace**)realloc(mList, sizeof(NameSpace*) * new_size);
		if (!new_list)
			return false; // out of memory
		mListSize = new_size;
		mList = new_list;
	}
	// Add namespace and increment count
	mList[mCount] = aNameSpace;
	mCount++;
	return true;
}

void NameSpaceList::RemoveLastNameSpace()
{
	// removes the last namespace.
	ASSERT(mList);
	delete mList[--mCount];
}

bool NameSpaceList::IsInList(NameSpace* aNameSpace)
{
	return find(aNameSpace->GetName());
}

bool NameSpaceList::find(LPTSTR aName, NameSpace **aFound)
{
	// Anonymous namespaces never match.
	for (size_t i = 0; i < mCount; ++i)
	{
		if (NAMESPACE_NAMES_MATCH(aName, mList[i]->GetName()))
		{
			if (aFound)
				*aFound = mList[i];
			return true;
		}
	}
	return false;
}

ResultType NameSpaceList::AutoExecSection()
{
	// Run each namespace' auto exec section.
	ResultType result;
	for (size_t i = 0; i < mCount; ++i)
		if ((result = mList[i]->AutoExecSection()) != OK)
			return result;
	return OK;
}

void NameSpaceList::ReleaseVarObjects() // public
{
	for (size_t i = 0; i < mCount; ++i)
		mList[i]->ReleaseVarObjects(); // this releases the vars in the nested namespaces too.
}

void NameSpaceList::FreeOptionalNameSpaceList()
{
	// see the corresponding NameSpace:: method of comments
	for (size_t i = 0; i < mCount; ++i)
		mList[i]->FreeOptionalNameSpaceList();
}
#ifndef AUTOHOTKEYSC
void NameSpaceList::FreeSourceFileIndexList()
{
	// see the corresponding NameSpace:: method of comments
	for (size_t i = 0; i < mCount; ++i)
		mList[i]->FreeSourceFileIndexList();
}
#endif

#undef NAMESPACELIST_INITIAL_SIZE
#undef NAMESPACELIST_SIZE_GROW