#pragma once


struct ScriptImport
{
	LPTSTR names = nullptr, mod_path = nullptr, mod_name = nullptr, var_name = nullptr;
	ScriptModule *mod = nullptr;
	ScriptImport *next = nullptr;
	LineNumberType line_number = 0;
	FileIndexType file_index = 0;
	bool wildcard = false;
	bool is_export = false;

	ScriptImport() {}
	ScriptImport(ScriptModule *aMod) : mod(aMod), names(_T("*")), wildcard(true) {}

	void *operator new(size_t aBytes) {return SimpleHeap::Alloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Alloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}
};


class ScriptModule : public ObjectBase
{
public:
	LPCTSTR mName = nullptr;
	Line *mFirstLine = nullptr, *mLastLine = nullptr;
	Label *mFirstLabel = nullptr, *mLastLabel = nullptr;
	ScriptImport *mImports = nullptr;
	ScriptModule *mPrev = nullptr;
	VarList mVars;
	Var *mSelf = nullptr;
	UnresolvedBaseClass *mUnresolvedBaseClass = nullptr;
	FileIndexType *mFiles = nullptr, mFilesCount = 0, mFilesCountMax = 0;
	LineNumberType mDirectiveLineNumber = 0;
	FileIndexType mSelfFileIndex = ABSOLUTE_MAX_SOURCE_FILES;
	FileIndexType mOuterFileIndex = ABSOLUTE_MAX_SOURCE_FILES;
	FileIndexType mDirectiveFileIndex = ABSOLUTE_MAX_SOURCE_FILES;
	bool mExecuted = false;
	bool mIsBuiltinModule = false;
	bool mHasWildcardExports = false;
	bool mBackCompatMode = true; // Compatibility mode for module startup code.  true = requires v2.0, false = requires v2.1
	bool mBackCompatModeWasSet = false;

	// #Warn settings
	WarnMode Warn_LocalSameAsGlobal = WARNMODE_OFF;
	WarnMode Warn_Unreachable = WARNMODE_ON;
	WarnMode Warn_VarUnset = WARNMODE_ON;

	bool IsFileModule() const { return mSelfFileIndex != ABSOLUTE_MAX_SOURCE_FILES; }

	bool HasFileIndex(FileIndexType aFile);
	ResultType AddFileIndex(FileIndexType aFile);

	IObject *FindGlobalObject(LPCTSTR aName);
	Var *FindImportedVar(LPCTSTR aName);
	Var *FindImportableVar(LPCTSTR aName, bool aAllowCreate = false);
	Var *AddNewImportVar(LPTSTR aVarName, Var *aAliasFor, IObject *aModule, bool aExport);

	ScriptModule() {}
	ScriptModule(LPCTSTR aName) : mName(aName) {}

	void *operator new(size_t aBytes) {return SimpleHeap::Alloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Alloc(aBytes);}
	void operator delete(void *aPtr) {}
	void operator delete[](void *aPtr) {}

	IObject_Type_Impl("Module");
	ResultType Invoke(IObject_Invoke_PARAMS_DECL) override;
	Object *Base() override { return sPrototype; }
	static Object *sPrototype;
};

typedef ScriptItemList<ScriptModule, 16> ScriptModuleList;
