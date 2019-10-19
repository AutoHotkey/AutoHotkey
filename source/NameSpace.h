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

#include "defines.h"
#include "var.h"
#include "globaldata.h"
#include "script.h"
#include "NameSpaceDefines.h"

// Both class NameSpace and class NameSpaceList uses SimpleHeap::Malloc with the `new` operator.

// Forward declarations.
class NameSpaceList;
struct FuncList;
class Line;
class NameSpace
{
private:
	LPTSTR mName;										// Name of namespace.
	NameSpaceList *mNestedNameSpaces;					// List of nested namespaces, this is NULL until a request to add a nested namespace, see InsertNestedNameSpace
	NameSpace *mOuter;									// The namesspace in which this name space resides. 
														// Having a reference to the enclosing namespace facilitates load time restoration of the outer namespace when the inner namespace definition ends.
														// Also used by scope operator to access the enclosing scope, referenced to as NAMESPACE_OUTER_NAMESPACE_NAME.
												
	FuncList mFuncs;									// List of functions
	Var **mVar, **mLazyVar;								// Array of pointers-to-variable, allocated upon first use and later expanded as needed.
	int mVarCount, mVarCountMax, mLazyVarCount;			// Count of items in the above array as well as the maximum capacity.
	
	// Only for load time -- start
	class NameSpaceNameList : public SimpleList<LPTSTR>
	{
	public:
		NameSpaceNameList() : SimpleList(true) {}
		bool AreEqual(LPTSTR aName1, LPTSTR aName2) { return NAMESPACE_NAMES_MATCH(aName1, aName2); } // virtual
		void FreeItem(LPTSTR aName) { free(aName); }	// virtual
	} *mOptionalNameSpaces;								// A list of optional namespaces to allow "#import *i". Search word: NAMESPACE_INCLUDE_DIRECTIVE_NAME

#ifndef AUTOHOTKEYSC
	SimpleList<int> *mSourceFileIndexList;				// A list of numbers corresponding to indices in Line::sSourceFile, to allow #include duplicates across namespaces.
#endif
	
	// Only for load time -- end

	global_struct mSettings;							// The settings belonging to the namespace.


	bool mLaunchesCriticalThreads;						// If true, each new thread which has its first line in this namespace, will be critical.
	DWORD mPeekFrequency;								// The peek frequency for critical threads.

	
	// Methods for releasing objects in vars.
	void ReleaseVarObjects(Var** aVar, int aCount);
	void ReleaseStaticVarObjects(Var **aVar, int aVarCount);
	void ReleaseStaticVarObjects(FuncList &aFuncs);

public:
	static const LPTSTR sAnonymousNameSpaceName;							// All anonymous namespaces will share this name
																			// to facilitate implementation and debugging.

	// To control the order of the autoexecute sections and static var initializers,
	// the following "Script::" vars and methods are introduced.
	Line *mFirstLine, *mLastLine;											// For starting the autoexec section for each namespace.
	Line *mFirstStaticLine, *mLastStaticLine;								// The first and last static var initializer.

	ResultType AutoExecSection();											// Each namespace has its own autoexecute section.

	// only for load time:
	Object *mUnresolvedClasses;		// A list of unresolved classes. To enable extending a class which definition has not yet been encountered.
	bool mNoHotkeyLabels;			// To indicate that the first hotkey in this namespace should be "prefixed" with a return line.	
	bool mHasPreparsedExpressions;	// If true, indicates that when auto including a func due to the scope operator, the new lines must be preparsed separately.

	Var *FindVar(LPTSTR aVarName, size_t aVarNameLength = 0, int *apInsertPos = NULL, int aScope = FINDVAR_DEFAULT, bool *apIsLocal = NULL);
	Var *FindOrAddVar(LPTSTR aVarName, size_t aVarNameLength = 0, int aScope = FINDVAR_DEFAULT);
	Func *FindFunc(LPCTSTR aFuncName, size_t aFuncNameLength = -1, int *apInsertPos = NULL, bool aAllowNested = true);
	// end "Script::" vars and methods.

	Var *FindVarFromScopeSymbolDelimitedString(LPTSTR aString, bool aAllowAddVar, NameSpace **aFoundNameSpace = NULL);

	NameSpace(LPTSTR aName, int aFuncsInitSize = 0, NameSpace *aOuter = NULL) :
		mNestedNameSpaces(NULL),
		mOptionalNameSpaces(NULL),
#ifndef AUTOHOTKEYSC
		mSourceFileIndexList(NULL),
#endif
		mFirstLine(NULL), mLastLine(NULL),
		mFirstStaticLine(NULL), mLastStaticLine(NULL),
		mUnresolvedClasses(NULL),
		mNoHotkeyLabels(true),
		mHasPreparsedExpressions(false),
		mVar(NULL), mLazyVar(NULL),
		mVarCount(0), mVarCountMax(0), mLazyVarCount(0),
		mOuter(aOuter),
		mLaunchesCriticalThreads(false),
		mPeekFrequency(DEFAULT_PEEK_FREQUENCY)
	{
		if (aName != NAMESPACE_ANONYMOUS)
			mName = SimpleHeap::Malloc(aName);	// copy the name for simplicity.
		else
			mName = NAMESPACE_ANONYMOUS;
		if (!mName)
			return; // out of memory
		if (aFuncsInitSize && !mFuncs.Alloc(aFuncsInitSize))
			return; // out of memory

		global_init(mSettings);			// sets defaults
	}
	~NameSpace();
	LPTSTR GetName() { return mName; }
	NameSpaceList *GetNestedNameSpaces() { return mNestedNameSpaces; }		// Returns the list on nested namespaces.
	bool IsNestedNamespace() { return mOuter != NULL; }						// Returns true if this namesspace is nested inside another nested namespace.
	bool LeaveCurrentNameSpace() { return mOuter->SetCurrentNameSpace(); }	// Sets the enclosing namespace to be the current namespace.
																			// Caller must ensure mOuter is not NULL. This should only happen for the default and standard namespaces.
	bool AddOptionalNameSpace(LPTSTR aName);								// Manages a list of optional namespaces. To avoid load time errors when using "#import *i". Search word: NAMESPACE_INCLUDE_DIRECTIVE_NAME
	bool IsOptionalNameSpace(LPTSTR aName);									// Finds a name in the list of optional namespaces.
	void FreeOptionalNameSpaceList();
#ifndef AUTOHOTKEYSC
	bool AddSourceFileIndex(int aIndex);
	void FreeSourceFileIndexList();
	bool HasIncludedSourceFile(TCHAR aPath[]);
#endif

	// For handling critical:
	void SetThreadCriticalStatus(ScriptThread &t) { t.ThreadIsCritical = mLaunchesCriticalThreads; t.PeekFrequency = mPeekFrequency; t.AllowThreadToBeInterrupted = (bool)mLaunchesCriticalThreads; }
	DWORD GetPeekFrequency() { return mPeekFrequency; }
	bool GetCritical() { return mLaunchesCriticalThreads; }
	void MakeNameSpaceCritical(DWORD aNewPeekFrequency) { mLaunchesCriticalThreads = true;  mPeekFrequency = aNewPeekFrequency; }
	void MakeNameSpaceNonCritical() { mLaunchesCriticalThreads = false; mPeekFrequency = DEFAULT_PEEK_FREQUENCY; }

	NameSpace *GetNestedNameSpace(LPTSTR aNameSpaceName, bool aAllowReserved = false); // returns the namespace if this namespace has a nested namespace with name aNameSpaceName, else NULL.
	

	static NameSpace *GetReservedNameSpace(LPTSTR aName, NameSpace* aSource = NULL); // Returns one of the reserved namespace if aName matches.

	
	bool SetCurrentNameSpace();												// Sets this namespace to be the current one, should always be used instead of assigning g_CurrentNameSpace directly.

	global_struct &GetSettings();											// This is needed by HOT_IF_ACTIVE et al.

	NameSpace *GetOuterNameSpace();
	// Methods for inserting nested namespaces:
	NameSpace *InsertNestedNameSpace(LPTSTR aName, int aFuncsInitSize = 0, NameSpace *aOuter = NULL);
	bool InsertNestedNameSpace(NameSpace *aNameSpace);
	void RemoveLastNameSpace();

	// Misc methods:
	void ReleaseVarObjects();			// Called during application termination.

	Func *TokenToFunc(ExprTokenType &aToken, bool aAllowNested = false);

	// FuncList method wrappers:
	Func *FuncFind(LPCTSTR aName, int *apInsertPos);
	Func* FuncFind(LPCTSTR aName, size_t aNameLength, int *apInsertPos = NULL, bool aAllowNested = false);
	ResultType FuncInsert(Func *aFunc, int aInsertPos);
	ResultType FuncAlloc(int aAllocCount);
	FuncList &GetFuncs();

	// Var methods
	Var** &GetVar();
	Var** &GetLazyVar();
	int &GetVarCount();
	int &GetLazyVarCount();
	int &GetVarCountMax();

	// Operators
	void *operator new(size_t aBytes) { return SimpleHeap::Malloc(aBytes); }
	void *operator new[](size_t aBytes) { return SimpleHeap::Malloc(aBytes); }
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}

};

class NameSpaceList
{
	NameSpace **mList;				// an array of namespaces
	size_t mCount;					// the number of namespaces in mList
	size_t mListSize;				// the size of mList
public:
	NameSpaceList() : mList(NULL), mCount(0), mListSize(0) {}	// constructor
	~NameSpaceList()
	{
		if (mList)
			free(mList);
	}
	// list management:
	bool Add(NameSpace * aNameSpace);
	void RemoveLastNameSpace();
	bool IsInList(NameSpace *aNameSpace);
	bool find(LPTSTR aName, NameSpace **aFound = NULL);

	// Script methods:
	ResultType AutoExecSection();
	void ReleaseVarObjects();
	
	// misc methods:
	void FreeOptionalNameSpaceList();	// calls the corresponding NameSpace method for all namespaces in the list.
#ifndef AUTOHOTKEYSC
	void FreeSourceFileIndexList();		// ---
#endif
	
	// Operators
	void *operator new(size_t aBytes) { return SimpleHeap::Malloc(aBytes); }
	void *operator new[](size_t aBytes) { return SimpleHeap::Malloc(aBytes); }
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};