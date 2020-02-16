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
#include "ScriptModulesDefines.h"

// Forward declaration
class ModuleList; 

class ScriptModule
{
private:
	LPTSTR mName;										// Name of module.
	ScriptModule* mOuter;								// The module in which this module resides. 
														// Having a reference to the enclosing module facilitates load time restoration of the outer module when the inner module definition ends.
														// Can also be used to refer to the outer module via scope resolution.
	ModuleList* mNested;								// List of nested modules.
	
	
	void ReleaseVarObjects(Var** aVar, int aVarCount);
	void ReleaseStaticVarObjects(Var** aVar, int aVarCount);
	void ReleaseStaticVarObjects(FuncList& aFuncs);


public:

	FuncList mFuncs;									// List of functions
	Var** mVar, ** mLazyVar;							// Array of pointers-to-variable, allocated upon first use and later expanded as needed.
	int mVarCount, mVarCountMax, mLazyVarCount;			// Count of items in the above array as well as the maximum capacity.

	static const LPTSTR sUnamedModuleName;				// All unnamed modules will share this name
														// to facilitate implementation and debugging.

	LPTSTR GetName() { return mName; }					// Get the name of the module.

	bool SetCurrentModule() { g_CurrentModule = this; return true; }	// Use this rather than direct assignment for maintainability.
	bool LeaveCurrentModule() { return mOuter->SetCurrentModule(); }	// Sets the enclosing module to be the current module.

	bool ValidateName(LPTSTR aName);
	
	
	// Only for load time
	void RemoveLastModule();
	bool AddOptionalModule(LPTSTR aName);
	bool IsOptionalModule(LPTSTR aName);
	class ModuleNameList : public SimpleList<LPTSTR>
	{
	public:
		ModuleNameList() : SimpleList(true) {}
		bool AreEqual(LPTSTR aName1, LPTSTR aName2) { return SMODULES_NAMES_MATCH(aName1, aName2); } // virtual
		void FreeItem(LPTSTR aName) { free(aName); }	// virtual
	} *mOptionalModules;								// A list of optional modules to allow the "*i" option for SMODULES_INCLUDE_DIRECTIVE_NAME.


#ifndef AUTOHOTKEYSC
	SimpleList<int>* mSourceFileIndexList;				// A list of numbers corresponding to indices in Line::sSourceFile, to allow #include duplicates across modules.
	// For "including" modules.
	bool AddSourceFileIndex(int aIndex);
	void FreeSourceFileIndexList();
	bool HasIncludedSourceFile(TCHAR aPath[]);
#endif // #ifndef AUTOHOTKEYSC
	// To manage nested modules:
	ScriptModule* InsertNestedModule(LPTSTR aName, int aFuncsInitSize, ScriptModule* aOuter);
	bool InsertNestedModule(ScriptModule* aModule);
	ScriptModule* GetNestedModule(LPTSTR aModuleName, bool aAllowReserved = false); // returns the module if this module has a nested module with name aModuleName, else NULL.
	
	static ScriptModule* GetReservedModule(LPTSTR aName, ScriptModule* aSource = NULL);

	
	void ReleaseVarObjects();

	void FreeOptionalModuleList();


	ScriptModule(LPTSTR aName, int aFuncsInitSize = 0, ScriptModule* aOuter = NULL) :
		mVar(NULL), mLazyVar(NULL),
		mVarCount(0), mVarCountMax(0), mLazyVarCount(0),
		mOuter(aOuter), mNested(NULL),
		mOptionalModules(NULL), mSourceFileIndexList(NULL)
	{
		if (aName != SMODULES_UNNAMED_NAME)
			mName = SimpleHeap::Malloc(aName);	// copy the name for simplicity.
		else
			mName = SMODULES_UNNAMED_NAME;
		if (!mName)
			return; // out of memory
		if (aFuncsInitSize && !mFuncs.Alloc(aFuncsInitSize))
			return; // out of memory
	}

	// Operators
	void* operator new(size_t aBytes) { return SimpleHeap::Malloc(aBytes); }
	void* operator new[](size_t aBytes) { return SimpleHeap::Malloc(aBytes); }
	void operator delete(void* aPtr) {}
	void operator delete[](void* aPtr) {}

};

class ModuleList
{
	ScriptModule** mList;			// an array of ScriptModules
	size_t mCount;					// the number of Modules in mList
	size_t mListSize;				// the size of mList
public:
	ModuleList() : mList(NULL), mCount(0), mListSize(0) {}	// constructor
	~ModuleList()
	{
		if (mList)
			free(mList);
	}
	// list management:
	bool Add(ScriptModule* aModule);
	bool IsInList(ScriptModule* aModule);
	void RemoveLastModule();
	bool find(LPTSTR aName, ScriptModule** aFound = NULL);

	void FreeOptionalModuleList();
#ifndef AUTOHOTKEYSC
	void FreeSourceFileIndexList();
	void ReleaseVarObjects();
#endif
	// Operators
	void* operator new(size_t aBytes) { return SimpleHeap::Malloc(aBytes); }
	void* operator new[](size_t aBytes) { return SimpleHeap::Malloc(aBytes); }
	void operator delete(void* aPtr) {}
	void operator delete[](void* aPtr) {}
};