#pragma once
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