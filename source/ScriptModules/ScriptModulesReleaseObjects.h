#pragma once
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
		auto& f = *(UserFunc*)aFuncs.mItem[i];
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
