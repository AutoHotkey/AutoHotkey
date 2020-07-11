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

// These names are reserved (as module names) and cannot be defined by the script.
// They can be changed here to any valid "var name". 
#define SMODULES_DEFAULT_MODULE_NAME _T("Script")	// This is the name of the module which the entire script resides in. 
													// It allows easy access to the top layer in a general and recognizable way.

#define SMODULES_STANDARD_MODULE_NAME _T("std")		// This is the name of the standard module for easy access to built-in functions and possibly future addititions. This module is parallel to the script's default module.

#define SMODULES_OUTER_MODULE_NAME _T("Outer")

// For unnamed modules:
#define SMODULES_UNNAMED_NAME _T("<unnamed ") SMODULES_DECLARATION_KEYWORD_NAME _T(">")	// For listlines / errors etc. This name is invalid for scripts due to '<', ' ' and '>'. This name can be any string. At least IsDirective assumes this is not _T("").
#define SMODULES_UNNAMED_STR (ScriptModule::sUnamedModuleName)	// This is a static string for quickly identifying an unnamed module and avoid allocating a string for each unnamed module.

// For literal module definitions:
#define SMODULES_DECLARATION_KEYWORD_NAME_QUOTED "Module"
#define SMODULES_DECLARATION_KEYWORD_NAME _T(SMODULES_DECLARATION_KEYWORD_NAME_QUOTED)
#define SMODULES_DECLARATION_KEYWORD_NAME_LENGTH (_countof(SMODULES_DECLARATION_KEYWORD_NAME) - 1) // - 1 to exclude the '\0
#define SMODULES_DECLARATION_KEYWORD_NAME_LC _T("module") // Lower case version for error reporting etc.

// For including file to module:
#define SMODULES_INCLUDE_DIRECTIVE_NAME _T("#Import")
#define SMODULES_INCLUDE_DIRECTIVE_NAME_LENGTH (_countof(SMODULES_INCLUDE_DIRECTIVE_NAME) - 1) // - 1 to exclude the '\0'
#define SMODULES_INCLUDE_DIRECTIVE_OPTIONAL_MARKER _T("*i")
#define SMODULES_STR_EQUALS_INCLUDE_DIRECTIVE_OPTIONAL_MARKER(str) (str[0] == '*' && ctoupper(str[1]) == 'I') // Relies on short-circuit boolean order.
#define SMODULES_INCLUDE_DIRECTIVE_OPTIONAL_MARKER_LENGTH (_countof(SMODULES_INCLUDE_DIRECTIVE_OPTIONAL_MARKER) - 1) // - 1 to exclude the '\0'

#define SMODULES_INCLUDE_DIRECTIVE_FILE_MODULE_SEP _T("as") // The separator is required to be enclosed in whitespace
#define SMODULES_INCLUDE_DIRECTIVE_FILE_MODULE_SEP_LENGTH (_countof(SMODULES_INCLUDE_DIRECTIVE_FILE_MODULE_SEP) - 1) // - 1 to exclude the '\0'

// For importing names:
#define SMODULES_IMPORT_VARS_DECLARATION_KEYWORD _T("UseVar")
#define SMODULES_IMPORT_FUNCS_DECLARATION_KEYWORD _T("UseFunc")

#define SMODULES_IMPORT_VARS_DECLARATION_KEYWORD_LENGTH (_countof(SMODULES_IMPORT_VARS_DECLARATION_KEYWORD) - 1) // - 1 to exclude the '\0'
#define SMODULES_IMPORT_FUNCS_DECLARATION_KEYWORD_LENGTH (_countof(SMODULES_IMPORT_FUNCS_DECLARATION_KEYWORD) - 1) // - 1 to exclude the '\0'


#define SMODULES_IMPORT_NAME_SEP _T("in")
#define SMODULES_IMPORT_NAME_SEP_LENGTH (_countof(SMODULES_IMPORT_NAME_SEP) - 1) // - 1 to exclude the '\0'

#define SMODULES_IMPORT_NAME_ALL_MARKER _T("*All")

// Rule for module names match:
#define SMODULES_NAMES_MATCH(name1, name2) ((name1) != SMODULES_UNNAMED_STR && (name2) != SMODULES_UNNAMED_STR && !_tcsicmp( (name1), (name2)))

// Error messages:
#define ERR_SMODULES_NOT_FOUND _T("Module not found.")
#define ERR_SMODULES_DUPLICATE_NAME _T("Duplicate module definition.")
#define ERR_SMODULES_IN_FUNCTION _T("Functions cannot contain modules.")
#define ERR_SMODULES_IN_CLASS _T("Classes cannot contain modules.")
#define ERR_SMODULES_IN_BLOCK _T("This block cannot contain modules.")
#define ERR_SMODULES_DEFINITION_SYNTAX _T("Syntax error in module definition.") // Also used with SMODULES_INCLUDE_DIRECTIVE_NAME.
#define ERR_SMODULES_INVALID_SCOPE_RESOLUTION _T("Invalid scope resolution.")

#define ERR_SMODULES_NOT_SUPPORTED _T("Not supported.")

#define ERR_SMODULES_BAD_DECLARATION _T("Bad declaration.")
#define ERR_SMODULES_VAR_NOT_FOUND _T("Variable not found.")
#define ERR_SMODULES_FUNC_NOT_FOUND _T("Function not found.")
#define ERR_SMODULES_UNRESOLVED_NAME _T("Could not resolve name.")

// Misc

// Warning: This macro declares a variable outside the for block
#define FOR_EACH_MODULE(mod) int Macro_Index = 0; for ( ScriptModule *mod; mod = g_script.mModuleSimpleList.GetItem(Macro_Index); ++Macro_Index)