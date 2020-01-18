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

class ScriptModule
{
private:
	LPTSTR mName;										// Name of module.
	ScriptModule* mOuter;								// The module in which this module resides. 
														// Having a reference to the enclosing module facilitates load time restoration of the outer module when the inner module definition ends.
														// Can also be used to refer to the outer module via scope resolution.

	FuncList mFuncs;									// List of functions
	Var** mVar, ** mLazyVar;							// Array of pointers-to-variable, allocated upon first use and later expanded as needed.
	int mVarCount, mVarCountMax, mLazyVarCount;			// Count of items in the above array as well as the maximum capacity.

public:

	static const LPTSTR sUnamedModuleName;							// All anonymous namespaces will share this name
																			// to facilitate implementation and debugging.


	ScriptModule(LPTSTR aName, int aFuncsInitSize = 0, ScriptModule* aOuter = NULL) :
		mVar(NULL), mLazyVar(NULL),
		mVarCount(0), mVarCountMax(0), mLazyVarCount(0),
		mOuter(aOuter)
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

};