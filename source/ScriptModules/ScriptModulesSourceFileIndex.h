#pragma once
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
