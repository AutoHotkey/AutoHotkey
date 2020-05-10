#pragma once

void ScriptModule::RemoveLastModule()
{
	// caller has ensured mNested must exist.
	mNested->RemoveLastModule();
}

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

bool ScriptModule::InsertNestedModule(ScriptModule* aModule) // public
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
