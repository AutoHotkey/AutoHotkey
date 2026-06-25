#include "stdafx.h"
#include "script.h"
#include "globaldata.h"


Object *ScriptModule::sPrototype;


ResultType ScriptModule::Invoke(IObject_Invoke_PARAMS_DECL)
{
	Var *var = aName ? FindImportableVar(aName) : nullptr;
	if (!var)
		return ObjectBase::Invoke(IObject_Invoke_PARAMS);

	if (IS_INVOKE_SET && aParamCount == 1)
		return var->Assign(*aParam[0]);

	if (var->IsUninitialized())
	{
		if (var->CanSelfInitialize()) // Uninitialized class or import.
		{
			auto result = var->SelfInitialize();
			if (result == FAIL || result == EARLY_EXIT)
				return result;
		}
		if (var->IsUninitialized())
			return g_script.VarUnsetError(var);
	}

	var->Get(aResultToken);
	if (aResultToken.Exited() || aParamCount == 0 && IS_INVOKE_GET)
		return aResultToken.Result();

	return Object::ApplyParams(aResultToken, aFlags, aParam, aParamCount);
}


ScriptModule *Script::FindDirectiveModule(LPCTSTR aName, ScriptModule *aList)
{
	ScriptModule *mod;
	for (mod = aList; mod && !mod->IsFileModule(); mod = mod->mPrev)
		if (!_tcsicmp(aName, mod->mName))
			return mod;
	if (!_tcsicmp(aName, _T("AHK")))
		return &mBuiltinModule;
	if (*aName == '_')
	{
		if (!_tcsicmp(aName, _T("__Init")))
			return mod;
		if (!_tcsicmp(aName, _T("__Main")))
			return &mDefaultModule;
	}
	return nullptr;
}


ResultType Script::ParseModuleDirective(LPCTSTR aName)
{
	auto mod = FindDirectiveModule(aName, mLastModule);
	if (!mod)
	{
		if (!Var::ValidateName(aName, DISPLAY_MODULE_ERROR))
			return FAIL;

		mod = CreateModule(SimpleHeap::Alloc(aName));
		mod->mDirectiveFileIndex = mCurrFileIndex;
		mod->mDirectiveLineNumber = mCombinedLineNumber;
		auto outer = mLastModule;
		while (!outer->IsFileModule()) outer = outer->mPrev;
		mod->mOuterFileIndex = outer->mSelfFileIndex;
		mod->mBackCompatMode = outer->mBackCompatMode; // Mode for startup code.  Set default based on __Init module.
		mBackCompatMode = mod->mBackCompatMode; // Mode for function definitions to inherit.

		// Let any previous #Warn settings carry over from the previous module, by default.
		mod->Warn_LocalSameAsGlobal = mCurrentModule->Warn_LocalSameAsGlobal;
		mod->Warn_Unreachable = mCurrentModule->Warn_Unreachable;
		mod->Warn_VarUnset = mCurrentModule->Warn_VarUnset;
	}
	else if (mCombinedLineNumber == mod->mDirectiveLineNumber && mCurrFileIndex == mod->mDirectiveFileIndex)
	{
		return ScriptError(_T("#Module in duplicate #Include."));
	}
	if (mod != mCurrentModule)
	{
		ReopenModule(mod);
	}
	return CONDITION_TRUE;
}


ResultType Script::ParseImportDirective(LPTSTR aBuf)
{
	bool is_export = false;
	if (!_tcsnicmp(aBuf, _T("Export"), 6) && IS_SPACE_OR_TAB(aBuf[6]))
	{
		auto cp = omit_leading_whitespace(aBuf + 7);
		if ( !(*cp == '{' || !_tcsnicmp(cp, _T("as"), 2) && IS_SPACE_OR_TAB(cp[2])) )
		{
			is_export = true;
			aBuf = cp;
		}
	}

	LPTSTR cp = aBuf;
	LPTSTR mod_name = nullptr, mod_name_end = nullptr;
	LPTSTR var_name = nullptr, var_name_end = nullptr;
	LPTSTR names = nullptr, names_end = nullptr;
	bool import_file = false, import_wildcard = false;

	if (import_file = (*cp == '"' || *cp == '\''))
	{
		mod_name = cp + 1;
		cp += FindTextDelim(cp, *cp, 1);
		if (!*cp)
			return FAIL;
		mod_name_end = cp++;
	}
	else
	{
		mod_name = cp;
		mod_name_end = cp = find_identifier_end(cp);
		if (cp == mod_name)
			return FAIL;
		if (*cp)
			++cp;
	}
	if (*cp) // There's a character after the module name/path.
	{
		cp = omit_leading_whitespace(cp);
		if (!_tcsnicmp(cp, _T("as"), 2) && IS_SPACE_OR_TAB(cp[2]))
		{
			var_name = omit_leading_whitespace(cp + 3);
			var_name_end = find_identifier_end(var_name);
			if (var_name_end == var_name)
				return FAIL;
			cp = omit_leading_whitespace(var_name_end);
		}
		if (*cp == '{')
		{
			names = cp + 1;
			cp += FindExprDelim(cp, '}', 1);
			if (!*cp || cp[1])
				return FAIL;
			names_end = cp;
		}
		else if (*cp)
			return FAIL;
	}

	auto imp = new ScriptImport();
	imp->names = SimpleHeap::Alloc(names, names_end - names);
	imp->mod_path = SimpleHeap::Alloc(mod_name, mod_name_end - mod_name);
	if (import_file)
	{
		for (cp = mod_name_end - 1; cp > mod_name && IS_IDENTIFIER_CHAR(*cp); --cp);
		if (*cp == ':') // The last non-word character is ':'.
		{
			imp->mod_path[cp - mod_name] = '\0';
			imp->mod_name = imp->mod_path + (cp - mod_name) + 1;
		}
		ConvertEscapeSequences(imp->mod_path);
	}
	if (var_name)
		imp->var_name = SimpleHeap::Alloc(var_name, var_name_end - var_name);
	else if (!import_file)
		imp->var_name = imp->mod_path;
	imp->wildcard = import_wildcard;
	imp->is_export = is_export;
	imp->line_number = mCombinedLineNumber;
	imp->file_index = mCurrFileIndex;
	imp->next = mCurrentModule->mImports;
	mCurrentModule->mImports = imp;
	return OK;
}


Var *Script::FindImportedVar(LPCTSTR aVarName)
{
	return CurrentModule()->FindImportedVar(aVarName);
}


Var *ScriptModule::FindImportedVar(LPCTSTR aVarName)
{
	for (auto imp = mImports; imp; imp = imp->next)
	{
		if (imp->wildcard && imp->mod) // mod can be null during DerefInclude().
		{
			auto var = imp->mod->mVars.Find(aVarName);
			if (!var && imp->mod->mHasWildcardExports)
			{
				bool hwe = mHasWildcardExports;
				mHasWildcardExports = false; // Prevent infinite recursion.
				var = imp->mod->FindImportedVar(aVarName);
				mHasWildcardExports = hwe;
			}
			if (var && var->IsExported())
			{
				if (imp->mod->mExecuted)
					return var;
				// AddNewImportVar must be used to support executing the module on first reference.
				return AddNewImportVar(var->mName, var, imp->mod, imp->is_export);
			}
		}
	}
	return nullptr;
}


Var *ScriptModule::FindImportableVar(LPCTSTR aVarName, bool aAllowCreate)
{
	int at;
	auto var = mVars.Find(aVarName, &at);
	if (var)
		return var;
	if (!var && mIsBuiltinModule)
		var = g_script.FindOrAddBuiltInVar(aVarName, true, nullptr);
	if (!var)
		var = FindImportedVar(aVarName);
	if (!var && aAllowCreate)
	{
		// Since ResolveImports() does everything in one stage, imported names for some
		// modules don't exist yet.  Undeclared (but assigned) global variables also do
		// not exist at this stage.  The following allows both to be imported.
		var = new Var(const_cast<LPTSTR>(aVarName), VAR_GLOBAL);
		if (!var || !mVars.Insert(var, at))
			return nullptr;
	}
	return var;
}


// Add a new Var to the current module.
// Raises an error if a conflicting declaration exists.
// May use an existing Var if not previously marked as declared, such as if created by Export.
// Caller provides persistent memory for aVarName.
Var *ScriptModule::AddNewImportVar(LPTSTR aVarName, Var *aAliasFor, IObject *aModule, bool aExport)
{
	ASSERT(aVarName && aModule);
	int at;
	auto var = mVars.Find(aVarName, &at);
	if (var)
	{
		// mVars should contain only declared or exported variables at this point.
		if (var->IsDeclared() || aAliasFor && aAliasFor->ResolveAlias() == var)
		{
			if (aAliasFor ? var->IsAlias() && var->GetAliasFor() == aAliasFor : var->ToObject() == aModule)
				return var; // Already imported.
			g_script.ConflictingDeclarationError(_T("import"), var);
			return nullptr;
		}
		ASSERT(!var->IsAssignedSomewhere());
		var->Scope() |= VAR_DECLARED;
	}
	else
		var = new Var(aVarName, VAR_DECLARE_GLOBAL);
	if (!mVars.Insert(var, at))
	{
		delete var;
		MemoryError();
		return nullptr;
	}
	var->SetImport(aModule, aAliasFor);
	if (aExport)
		var->Scope() |= VAR_EXPORTED;
	return var;
}


ResultType Script::ResolveImports(ScriptModule *aTerminator)
{
	ScriptModule *directive_list = mLastModule;
	for (auto mod = mCurrentModule = mLastModule; mod != aTerminator; mod = mCurrentModule = mCurrentModule->mPrev)
	{
		if (!directive_list && !mod->IsFileModule())
			directive_list = mod;

		for (auto imp = mod->mImports; imp; imp = imp->next)
		{
			if (!imp->mod && !ResolveImports(*imp, directive_list))
				return FAIL;
		}

		if (mod->IsFileModule())
			directive_list = nullptr;
	}
	return OK;
}


void Script::ResolveIndirectImports()
{
	// Check any variables which have been added by other modules explicitly importing them
	// from this module, to see if they should be imported from some other module via wildcard.
	// This must be done after modules are resolved and assignments are parsed, since only an
	// unassigned var should implicitly import via wildcard.
	for (int i = 0; i < mCurrentModule->mVars.mCount; ++i)
	{
		auto var = mCurrentModule->mVars.mItem[i];
		if (!var->IsDeclared() && !var->IsAssignedSomewhere())
		{
			ASSERT(!var->IsAlias() && !var->IsDirectConstant());
			auto alias = mCurrentModule->FindImportedVar(var->mName);
			ASSERT(alias == var);
		}
	}
}


ResultType Script::ResolveImports(ScriptImport &imp, ScriptModule *aDirectiveList)
{
	if (!imp.mod_name)
	{
		// #Import mod_path
		imp.mod = FindDirectiveModule(imp.mod_path, aDirectiveList);
	}
	if (!imp.mod)
	{
		// #Import mod_path
		// #Import "mod_path"
		// #Import "mod_path:mod_name"
		// #Import ":mod_name"
		FileIndexType file_index;
		switch (FindModuleFileIndex(imp.mod_path, file_index, imp.file_index))
		{
		case FAIL:	return FAIL;
		case OK:	break;
		default:
			mCurrLine = nullptr;
			mCurrFileIndex = imp.file_index;
			mCombinedLineNumber = imp.line_number;
			return ScriptError(_T("Module not found"), imp.mod_path);
		}
		ScriptModule *new_last = nullptr;
		// Search by file index, which corresponds to the file's full path.
		for (auto m = mLastModule; m; m = m->mPrev)
		{
			if (m->mSelfFileIndex == file_index)
			{
				if (!new_last) // A file with no #modules
					new_last = m; // should still resolve :__Init.
				imp.mod = m;
				break;
			}
			if (!new_last && m->mOuterFileIndex == file_index)
				new_last = m;
		}
		if (!imp.mod)
		{
			auto path = Line::sSourceFile[file_index];
			auto cur_mod = mCurrentModule;
			imp.mod = mCurrentModule = CreateModule(imp.mod_name);
			imp.mod->mSelfFileIndex = file_index;
			if (!LoadIncludedFile(path, false, false))
				return FAIL;
			new_last = mLastModule;
			if (!CloseCurrentModule())
				return FAIL;
			if (!ResolveImports(imp.mod->mPrev)) // Resolve imports in all modules that were just included.
				return FAIL;
			mCurrentModule = cur_mod;
		}
		if (imp.mod && imp.mod_name && *imp.mod_name) // #Import "file:submodule"
		{
			// Now that "file" has been resolved and loaded, resolve "submodule".
			imp.mod = FindDirectiveModule(imp.mod_name, new_last);
			if (!imp.mod)
			{
				mCurrLine = nullptr;
				mCurrFileIndex = imp.file_index;
				mCombinedLineNumber = imp.line_number;
				return ScriptError(_T("Module not found"), imp.mod_name);
			}
		}
	}

	// Set in case of error.
	mCurrLine = nullptr;
	mCombinedLineNumber = imp.line_number;
	mCurrFileIndex = imp.file_index;

	if (imp.var_name)
	{
		// Do not reuse mSelf or a previous Var created by an import even if mod_name == var_name,
		// since the exported status of the Var (VAR_EXPORTED) shouldn't propagate between modules.
		auto var = mCurrentModule->AddNewImportVar(imp.var_name, imp.mod->mSelf, imp.mod, imp.is_export);
		if (!var)
			return FAIL;
	}

	if (imp.names)
	{
		for (LPTSTR cp = imp.names; *(cp = omit_leading_whitespace(cp)); ++cp)
		{
			TCHAR c;
			if (*cp == '*')
			{
				if (imp.mod != mCurrentModule)
				{
					imp.wildcard = true;
					if (imp.is_export)
						mCurrentModule->mHasWildcardExports = true;
					cp = omit_leading_whitespace(cp + 1);
				}
				c = *cp;
			}
			else
			{
				LPTSTR var_name = cp, mod_name = cp;
				c = *(cp = find_identifier_end(cp));
				*cp = '\0';
				while (IS_SPACE_OR_TAB(c)) c = *++cp; // Find next non-whitespace.
				if (!_tcsnicmp(cp, _T("as"), 2) && IS_SPACE_OR_TAB(cp[2]))
				{
					var_name = omit_leading_whitespace(cp + 3);
					cp = find_identifier_end(var_name);
					c = *cp;
					*cp = '\0';
					while (IS_SPACE_OR_TAB(c)) c = *++cp; // Find next non-whitespace.
				}
				auto exported = imp.mod->FindImportableVar(mod_name, true);
				if (!exported)
					return MemoryError();
				auto imported = mCurrentModule->AddNewImportVar(var_name, exported, imp.mod, imp.is_export);
				if (!imported)
					return FAIL;
			}
			if (!c)
				break;
			if (c != ',')
			{
				*cp = c;
				return ScriptError(_T("Invalid import"), cp);
			}
		}
	}

	return OK;
}


ResultType Script::CloseCurrentModule()
{
	// Reset these so they don't affect the next module.
	g->HotCriterion = nullptr;
	*mClassStructPack = 0;

	mCurrentModule->mLastLine = mLastLine;
	mLastLine = nullptr; // Start a new linked list of lines.
	mPendingRelatedLine = nullptr;
	return OK;
}


void Script::ReopenModule(ScriptModule *aMod)
{
	ASSERT(mCurrentModule != aMod);
	CloseCurrentModule();
	mCurrentModule = aMod;
	mLastLine = aMod->mLastLine; // Null unless we reopened an existing module.
	if (mLastLine)
		mPendingRelatedLine = mLastLine->mParentLine;
}


ScriptModule *Script::CreateModule(LPCTSTR aName)
{
	auto mod = new ScriptModule(aName);
	mod->mPrev = mLastModule;
	mLastModule = mod;
	return mod;
}


bool ScriptModule::HasFileIndex(FileIndexType aFile)
{
	for (int i = 0; i < mFilesCount; ++i)
		if (mFiles[i] == aFile)
			return true;
	return false;
}


ResultType ScriptModule::AddFileIndex(FileIndexType aFile)
{
	if (mFilesCount == mFilesCountMax)
	{
		auto new_size = mFilesCount ? mFilesCount * 2 : 8;
		auto new_files = (FileIndexType*)realloc(mFiles, new_size * sizeof(FileIndexType));
		if (!new_files)
			return MemoryError();
		mFiles = new_files;
		mFilesCountMax = new_size;
	}
	mFiles[mFilesCount++] = aFile;
	return OK;
}


IObject *ScriptModule::FindGlobalObject(LPCTSTR aName)
{
	if (auto var = mVars.Find(aName))
	{
		ASSERT(!var->IsVirtual());
		return var->ToObject();
	}
	return nullptr;
}


LPCWSTR Script::InitModuleSearchPath()
{
	SetCurrentDirectory(mFileDir);

	// The size of this array sets the total limit for the entire list of directories, resolved to full paths.
	// This is enough for extreme cases, up to the total limit for all environment variables (32767 chars).
	TCHAR buf[MAX_WIDE_PATH], var_buf[MAX_WIDE_PATH];
	LPTSTR d, cp = buf;
	LPCTSTR path_spec;
	DWORD len, space;

	if (GetEnvironmentVariable(L"AhkImportPath", var_buf, _countof(var_buf)))
		path_spec = var_buf;
	else
		path_spec = _T(".;%A_MyDocuments%\\AutoHotkey;%A_AhkPath%\\..");

	// Ensure a consistent result if A_LineFile is used, since this is only
	// done once for the lifetime of the program, not once for each Import.
	mCurrFileIndex = 0;
	mCombinedLineNumber = 0;

	LPTSTR deref_path;
	if (!DerefInclude(deref_path, path_spec))
		return _T("");

	for (auto p = deref_path; *p; p = d)
	{
		for (d = p; *d && *d != ';'; ++d);
		if (*d)
			*d++ = '\0'; // Terminate within deref_path.
		if (!*p)
			continue;
		if (cp - buf >= _countof(buf)) // Due to rarity, ignore any items that won't fit.
			break;
		space = (DWORD)(_countof(buf) - (cp - buf));
		len = GetFullPathName(p, space, cp, nullptr);
		if (!len || len >= space)
			continue; // Ignore this item.
		// Ignore nonexistent directories so each Import won't do unnecessary checks.
		DWORD attr = GetFileAttributes(cp);
		if (attr == INVALID_FILE_ATTRIBUTES || !(attr & FILE_ATTRIBUTE_DIRECTORY))
			continue; // Ignore this item.
		cp += len + 1; // Write the next item after the null terminator.
	}
	if (cp == buf) // No items; could happen in the case of a malformed env var.
		*cp++ = '\0'; // Ensure double-null termination.
	
	free(deref_path);
	return SimpleHeap::Alloc(buf, cp - buf);
}


ResultType Script::FindModuleFileIndex(LPCTSTR aName, FileIndexType &aFileIndex, FileIndexType aLocalFileIndex)
{
	if (!*aName)
	{
		aFileIndex = mCurrentModule->IsFileModule() ? mCurrentModule->mSelfFileIndex : mCurrentModule->mOuterFileIndex;
		return OK;
	}

	if (*aName == '*') // *Resource name or stdin.
		return SourceFileIndex(aName, aFileIndex);

	static auto search_path = InitModuleSearchPath();
	const auto suffix = EXT_AUTOHOTKEY;
	const auto suffix_length = _countof(EXT_AUTOHOTKEY) - 1;
	const auto dir_suffix = _T("\\__Init") EXT_AUTOHOTKEY;
	const auto dir_suffix_length = _countof(_T("\\__Init") EXT_AUTOHOTKEY) - 1;

	TCHAR buf[T_MAX_PATH];

	// Always search the directory of the current file first.
	LPCTSTR dir = search_path;
	auto line_file = Line::sSourceFile[aLocalFileIndex];
	if (*line_file != '*')
	{
		auto p = _tcsrchr(line_file, '\\');
		if (p && p - line_file < _countof(buf))
		{
			tmemcpy(buf, line_file, p - line_file + 1);
			buf[p - line_file + 1] = '\0';
			dir = buf;
		}
	}

	auto name_length = _tcslen(aName);
	if (name_length + dir_suffix_length >= _countof(buf))
		return CONDITION_FALSE;

	for (LPCTSTR next_dir; *dir; dir = next_dir)
	{
		if (dir == buf)
			next_dir = search_path;
		else
			next_dir = dir + _tcslen(dir) + 1;

		if (!SetCurrentDirectory(dir))
			continue; // Ignore this item.

		tmemcpy(buf, aName, name_length + 1);

		// Try exact aName first.
		DWORD attr = GetFileAttributes(buf);
		auto p = buf + name_length;
		if ((attr & FILE_ATTRIBUTE_DIRECTORY) && attr != INVALID_FILE_ATTRIBUTES)
		{
			// Try aName\__module.ahk.
			tmemcpy(p, dir_suffix, dir_suffix_length + 1);
			attr = GetFileAttributes(buf);
		}
		if (attr & FILE_ATTRIBUTE_DIRECTORY) // No file found yet.
		{
			// Try aName.ahk.
			tmemcpy(p, suffix, suffix_length + 1);
			attr = GetFileAttributes(buf);
		}
		if ((attr & FILE_ATTRIBUTE_DIRECTORY) == 0) // File exists and is not a directory.
			return SourceFileIndex(buf, aFileIndex);
	}
	return CONDITION_FALSE;
}
