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

// These names are reserved (as namespace names) and cannot be defined by the script.
// They can be changed here to any valid "var name". 
#define NAMESPACE_DEFAULT_NAMESPACE_NAME _T("Script")	// This is the name of the namespace which the entire script resides in. 
														// It allows easy access to the top layer in a general and recognizable way.

#define NAMESPACE_STANDARD_NAMESPACE_NAME _T("std")		// This is the name of the standard namespace for easy access to built-in functions and possibly future addititions. This namespace is parallel to the scripts namespace.

#define NAMESPACE_TOP_NAMESPACE_NAME _T("Library")
#define NAMESPACE_OUTER_NAMESPACE_NAME _T("Outer")

// For anonymous namespaces:
#define NAMESPACE_ANONYMOUS_NAME _T("<anonymous namespace>")		// For listlines / errors etc. This name is invalid for scripts due to '<', ' ' and '>'. This name can be any string.
#define NAMESPACE_ANONYMOUS	(NameSpace::sAnonymousNameSpaceName)	// This is a static string for quickly identifying a anonymous namespace and avoid allocating a string for each anonymous namespace.

// the scope operator:
// Note that changing the symbols here doesn't automatically change the symbol
// for example, some switches will have case '-': rather than case OPERATOR_SCOPE_SYMBOL_PART1: 
// since '-' is used for other operators too.

// Maintain these together,
#define OPERATOR_SCOPE_SYMBOL_PART1 '-'
#define OPERATOR_SCOPE_SYMBOL_PART2 '>'
#define OPERATOR_SCOPE_SYMBOL _T("->")	
#define OPERATOR_SCOPE_SYMBOL_LENGTH (_countof(OPERATOR_SCOPE_SYMBOL) - 1)	 // - 1 to exclude the '\0'

// For literal namespace definitions:
#define NAMESPACE_DECLARATION_KEYWORD_NAME _T("Namespace")
#define NAMESPACE_DECLARATION_KEYWORD_NAME_LENGTH (_countof(NAMESPACE_DECLARATION_KEYWORD_NAME) - 1) // - 1 to exclude the '\0

#define NAMESPACE_DECLARATION_TOP_KEYWORD_NAME _T("Library") // This can be changed to a multiword keyword, eg, Namespace Library
#define NAMESPACE_DECLARATION_TOP_KEYWORD_NAME_LENGTH (_countof(NAMESPACE_DECLARATION_TOP_KEYWORD_NAME) - 1) // - 1 to exclude the '\0

// For including file to namespace:
#define NAMESPACE_INCLUDE_DIRECTIVE_NAME _T("#Import")
#define NAMESPACE_INCLUDE_DIRECTIVE_NAME_LENGTH (_countof(NAMESPACE_INCLUDE_DIRECTIVE_NAME) - 1) // - 1 to exclude the '\0'

#define NAMESPACE_INCLUDE_TOP_DIRECTIVE_NAME _T("#ImportLibrary")
#define NAMESPACE_INCLUDE_TOP_DIRECTIVE_NAME_LENGTH (_countof(NAMESPACE_INCLUDE_TOP_DIRECTIVE_NAME) - 1) // - 1 to exclude the '\0'


// Rule for namespace names match:
#define NAMESPACE_NAMES_MATCH(name1, name2) ((name1) != NAMESPACE_ANONYMOUS && (name2) != NAMESPACE_ANONYMOUS && !_tcsicmp( (name1), (name2)))

// Error messages:
#define ERR_NAMESPACE_NOT_FOUND _T("Namespace not found.")
#define ERR_NAMESPACE_DUPLICATE_NAME _T("Duplicate namespace definition.")
#define ERR_NAMESPACE_IN_FUNCTION _T("Functions cannot contain namespaces.")
#define ERR_NAMESPACE_IN_CLASS _T("Classes cannot contain namespaces.")
#define ERR_NAMESPACE_IN_BLOCK _T("This block cannot contain namespaces.")
#define ERR_NAMESPACE_DEFINITION_SYNTAX _T("Syntax error in namespace definition.") // Also used with #import.
#define ERR_SCOPE_OPERATOR_MISSING_VAR_OR_FUNC _T("Scope resolution operator missing its end variable or function.")

// Misc

// Warning: This macro declares a variable outside the for block
#define FOR_EACH_NAMESPACE(ns) int Macro_Index = 0; for ( NameSpace *ns; ns = g_script.mNameSpaceSimpleList.GetItem(Macro_Index); ++Macro_Index)