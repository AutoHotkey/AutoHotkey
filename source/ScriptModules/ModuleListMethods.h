#pragma once

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

bool ModuleList::find(LPTSTR aName, ScriptModule** aFound)
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

#undef MODULELIST_INITIAL_SIZE
#undef MODULELIST_SIZE_GROW

void ModuleList::ReleaseVarObjects() // public
{
	for (size_t i = 0; i < mCount; ++i)
		mList[i]->ReleaseVarObjects(); // this releases the vars in the nested modules too.
}