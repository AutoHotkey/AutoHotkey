#pragma once
#define SM_RETURN_ERROR(...) return aResultToken.Error(__VA_ARGS__)
ResultType ScriptModule::Invoke(IObject_Invoke_PARAMS_DECL)
{
	aResultToken.symbol = SYM_STRING; // Set default
	aResultToken.marker = _T("");

	int invoke_type = aFlags;

	if (!aName) // Invalid, eg, %'MyModule'%[x]
		SM_RETURN_ERROR(ERR_SMODULES_INVALID_SCOPE_RESOLUTION);

	size_t name_length = _tcslen(aName);

	if (invoke_type & IT_CALL) // It is a function call, find the func and call it.
	{
		// This must be done before searching for a module since functions and modules can have the same name.
		Func* func = g_script.FindFunc(aName, name_length, 0, this);
		if (!func)
		{
			// Invoke value base. Because MyModule.%'NonExistentFunc'%(p*) "should"
			// be consistent with %'NonExistentFunc'%(p*).
			aParam--;
			aParamCount++;
			aResultToken.func->mFID = (BuiltInFunctionID)(invoke_type | IF_DEFAULT);
			Op_ObjInvoke(aResultToken, aParam, aParamCount);
			return aResultToken.Result();
		}
		if (func->IsBuiltIn() && this != g_StandardModule)
			SM_RETURN_ERROR(ERR_SMODULES_FUNC_NOT_FOUND);
		func->Call(aResultToken, aParam, aParamCount, false); // "false" since variadic parameters has already been expanded.
		return aResultToken.Result();
	}
	// Trying to resolve either a module or a variable.
	ScriptModule* found = NULL;

	if ((found = GetNestedModule(aName, true))
		|| IsOptionalModule(aName))
	{
		// It is a module.
		if (invoke_type & IT_SET) // Invalid, eg, MyModule.%MyOtherModule% := val
			SM_RETURN_ERROR(ERR_SMODULES_INVALID_SCOPE_RESOLUTION);

		aResultToken.symbol = SYM_OBJECT;
		aResultToken.object = found ? (IObject*)found : (IObject*)g_OptSM;;
	}
	else
	{	// It is a variable.
		if (!*aName) // Eg, MyModule.%''%
			SM_RETURN_ERROR(ERR_DYNAMIC_BLANK);
		Var* var = g_script.FindVar(aName, name_length, NULL, FINDVAR_SUPER_GLOBAL, NULL, this);
		if (!var || !var->IsSuperGlobal())
			SM_RETURN_ERROR(ERR_SMODULES_VAR_NOT_FOUND, aName);

		if (var->Type() == VAR_VIRTUAL) // Doesn't support built-in vars.
			SM_RETURN_ERROR(ERR_SMODULES_INVALID_SCOPE_RESOLUTION);
		if (invoke_type & IT_SET
			&& aParamCount == 1)
		{
			// Assiging a variable. Eg, MyModule.%expr% := value
			if (VAR_IS_READONLY(*var)) // For maintainability
				return	g_script.VarIsReadOnlyError(var, Script::AssignmentErrorType::INVALID_ASSIGNMENT);
			// Assign the value to the variable
			if (!(var->Assign(**aParam)))
			{
				// The error message should have been displayed, but the exit result must be set.
				return aResultToken.SetExitResult(FAIL);;
			}
			// Proceed to set the variable as the result of the assignment.
		}
		else if (aParamCount)
		{	
			// Parameters were passed, eg MyModule.%expr%[p*] [ := value ]
			ExprTokenType this_token;
			this_token.symbol = SYM_VAR;
			this_token.var = var;

			// Get the obj which will be invoked:
			IObject* obj;
			if (var->HasObject())
				obj = var->Object();
			else
			{
				// MyModule.%expr%[p*] should be equivalent to %expr%[p*] also when %expr%
				// is resolved to a var which doesn't contain an object.
				obj = Object::ValueBase(this_token);
				invoke_type |= IF_DEFAULT;
			}
			// AddRef and Release unconditionally for brevity. 
			// Done to avoid the object being deleted in case the variable is reasigned during the operation.
			obj->AddRef();
			ResultType result = obj->Invoke(aResultToken, invoke_type, NULL, this_token, aParam, aParamCount);
			obj->Release();
			return result;
		}
		// The result is a variable, eg, MyModule.%'var'% [:= value].
		aResultToken.symbol = SYM_VAR;
		aResultToken.var = var;
	}
	return OK;
}
#undef SM_RETURN_ERROR