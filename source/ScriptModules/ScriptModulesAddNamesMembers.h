#pragma once

// Used for importing names from other modules:
struct UseParams
{
	__int64 line_file_info;		// line/file info for error reporting.
	UserFunc* current_func;
	LPTSTR param1;	// The objects to use, eg a list of vars or funcs.
	union
	{	// Identifies the scope of the object(s) to use.
		LPTSTR str;				// SYM_STRING
		ScriptModule* mod;		// SYM_OBJECT
	} param2;
	SymbolType type_symbol;		// Indicates the type of the objects specified by param1.
	SymbolType scope_symbol;	// Indicates which member of param2 to use.
	
	

	UseParams(SymbolType aScopeSymbol, SymbolType aTypeSymbol, LPTSTR aObjList, ScriptModule* aModule, LPTSTR aModuleName) :
		param1(_tcsdup(aObjList)),
		type_symbol(aTypeSymbol),
		scope_symbol(aScopeSymbol)
	{
		if (scope_symbol == SYM_STRING)
			param2.str = _tcsdup(aModuleName);
		else
			param2.mod = aModule;
		current_func = g->CurrentFunc;
		line_file_info = g_script.GetFileLineInfo();
	}
	~UseParams()
	{
		if (param1) free(param1);
		if (scope_symbol == SYM_STRING) free(param2.str);
	}

};
class UseParamsList : public SimpleList<UseParams*>
{
public:
	UseParamsList() : SimpleList(true) {}
	void FreeItem(UseParams* aParams) { delete aParams; }	// virtual
} *mUseParams;								// A list of object to use.

UseParams* mCurrentUseParam;	// When resolving use params, this will refer to the current one.
								// Done this way to avoid passing around it and or using globals.