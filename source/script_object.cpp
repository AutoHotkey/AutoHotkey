#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"
#include "application.h"

#include "script_object.h"
#include "script_func_impl.h"

#include <errno.h> // For ERANGE.


//
// CallMethod - Invoke a method with no parameters, discarding the result.
//

ResultType CallMethod(IObject *aInvokee, IObject *aThis, LPTSTR aMethodName
	, ExprTokenType *aParamValue, int aParamCount, INT_PTR *aRetVal // For event handlers.
	, int aExtraFlags) // For Object.__Delete().
{
	ResultToken result_token;
	TCHAR result_buf[MAX_NUMBER_SIZE];
	result_token.InitResult(result_buf);

	ExprTokenType this_token(aThis), name_token(aMethodName);

	++aParamCount; // For the method name.
	ExprTokenType **param = (ExprTokenType **)_alloca(aParamCount * sizeof(ExprTokenType *));
	param[0] = &name_token;
	for (int i = 1; i < aParamCount; ++i)
		param[i] = aParamValue + (i-1);

	ResultType result = aInvokee->Invoke(result_token, this_token, IT_CALL | aExtraFlags, param, aParamCount);

	if (result != EARLY_EXIT && result != FAIL)
	{
		// Indicate to caller whether an integer value was returned (for MsgMonitor()).
		result = TokenIsEmptyString(result_token) ? OK : EARLY_RETURN;
	}
	
	if (aRetVal) // Always set this as some callers don't initialize it:
		*aRetVal = result == EARLY_RETURN ? (INT_PTR)TokenToInt64(result_token) : 0;

	result_token.Free();
	return result;
}


//
// Object::Create - Called by BIF_Object to create a new object, optionally passing key/value pairs to set.
//

Object *Object::Create(ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount & 1)
		return NULL; // Odd number of parameters - reserved for future use.

	Object *obj = new Object();
	if (obj && aParamCount)
	{
		if (aParamCount > 8)
			// Set initial capacity to avoid multiple expansions.
			// For simplicity, failure is handled by the loop below.
			obj->SetInternalCapacity(aParamCount >> 1);
		// Otherwise, there are 4 or less key-value pairs.  When the first
		// item is inserted, a default initial capacity of 4 will be set.

		TCHAR buf[MAX_NUMBER_SIZE];
		FieldType *field;
		SymbolType key_type;
		KeyType key;
		IndexType insert_pos;
		
		for (int i = 0; i + 1 < aParamCount; i += 2)
		{
			if (aParam[i]->symbol == SYM_MISSING || aParam[i+1]->symbol == SYM_MISSING)
				continue; // For simplicity.

			field = obj->FindField(*aParam[i], buf, key_type, key, insert_pos);
			if (!field // Probably NULL, but calling FindField() first avoids an extra call to TokenToString() below.
				&& key_type == SYM_STRING && !_tcsicmp(key.s, _T("base")))
			{
				// For consistency with assignments, the following is allowed to overwrite a previous
				// base object (although having "base" occur twice in the parameter list would be quite
				// useless) or set mBase to NULL if the value parameter is not an object.
				if (obj->mBase)
					obj->mBase->Release();
				if (obj->mBase = TokenToObject(*aParam[i + 1]))
					obj->mBase->AddRef();
				continue;
			}
			if (  !(field
				 || (field = obj->Insert(key_type, key, insert_pos)))
				|| !field->Assign(*aParam[i + 1])  )
			{	// Out of memory.
				obj->Release();
				return NULL;
			}
		}
	}
	return obj;
}

Object *Object::CreateArray(ExprTokenType *aValue[], int aValueCount)
{
	Object *obj = new Object();
	if (obj && aValueCount && !obj->InsertAt(0, 1, aValue, aValueCount))
	{
		obj->Release(); // InsertAt failed.
		obj = NULL;
	}
	return obj;
}


//
// Object::Clone - Used for variadic function-calls.
//

Object *Object::Clone(BOOL aExcludeIntegerKeys)
// Creates an object and copies to it the fields at and after the given offset.
{
	IndexType aStartOffset = aExcludeIntegerKeys ? mKeyOffsetObject : 0;

	Object *objptr = new Object();
	if (!objptr|| aStartOffset >= mFieldCount)
		return objptr;
	
	Object &obj = *objptr;

	// Allocate space in destination object.
	IndexType field_count = mFieldCount - aStartOffset;
	if (!obj.SetInternalCapacity(field_count))
	{
		obj.Release();
		return NULL;
	}

	FieldType *fields = obj.mFields; // Newly allocated by above.
	int failure_count = 0; // See comment below.
	IndexType i;

	obj.mFieldCount = field_count;
	obj.mKeyOffsetObject = mKeyOffsetObject - aStartOffset;
	obj.mKeyOffsetString = mKeyOffsetString - aStartOffset;
	if (obj.mKeyOffsetObject < 0) // Currently might always evaluate to false.
	{
		obj.mKeyOffsetObject = 0; // aStartOffset excluded all integer and some or all object keys.
		if (obj.mKeyOffsetString < 0)
			obj.mKeyOffsetString = 0; // aStartOffset also excluded some string keys.
	}
	//else no need to check mKeyOffsetString since it should always be >= mKeyOffsetObject.

	for (i = 0; i < field_count; ++i)
	{
		FieldType &dst = fields[i];
		FieldType &src = mFields[aStartOffset + i];

		// Copy key.
		if (i >= obj.mKeyOffsetString)
		{
			if ( !(dst.key.s = _tcsdup(src.key.s)) )
			{
				// Key allocation failed. At this point, all int and object keys
				// have been set and values for previous fields have been copied.
				// Rather than trying to set up the object so that what we have
				// so far is valid in order to break out of the loop, continue,
				// make all fields valid and then allow them to be freed. 
				++failure_count;
			}
		}
		else if (i >= obj.mKeyOffsetObject)
			(dst.key.p = src.key.p)->AddRef();
		else
			dst.key.i = src.key.i;

		// Copy value.
		switch (dst.symbol = src.symbol)
		{
		case SYM_STRING:
			dst.string.Init();
			if (!dst.Assign(src.string, src.string.Length(), true)) // Pass true to conserve memory (no space is allowed for future expansion).
				++failure_count; // See failure comment further above.
			break;
		case SYM_OBJECT:
			(dst.object = src.object)->AddRef();
			break;
		//case SYM_INTEGER:
		//case SYM_FLOAT:
		default:
			dst.n_int64 = src.n_int64; // Union copy.
		}
	}
	if (failure_count)
	{
		// One or more memory allocations failed.  It seems best to return a clear failure
		// indication rather than an incomplete copy.  Now that the loop above has finished,
		// the object's contents are at least valid and it is safe to free the object:
		obj.Release();
		return NULL;
	}
	return &obj;
}


//
// Object::ArrayToParams - Used for variadic function-calls.
//

void Object::ArrayToParams(ExprTokenType *token, ExprTokenType **param_list, int extra_params
	, ExprTokenType **aParam, int aParamCount)
// Expands this object's contents into the parameter list.  Due to the nature
// of the parameter list, only fields with integer keys are used (named params
// aren't supported).
// Return value is FAIL if a required parameter was omitted or malloc() failed.
{
	// Find the first and last field to be used.
	int start = (int)mKeyOffsetInt;
	int end = (int)mKeyOffsetObject; // For readability.
	while (start < end && mFields[start].key.i < 1)
		++start; // Skip any keys <= 0 (consistent with UDF-calling behaviour).
	
	int param_index;
	IndexType field_index;

	// For each extra param...
	for (field_index = start, param_index = 0; field_index < end; ++field_index, ++param_index)
	{
		for ( ; param_index + 1 < (int)mFields[field_index].key.i; ++param_index)
		{
			token[param_index].symbol = SYM_MISSING;
			token[param_index].marker = _T("");
			token[param_index].marker_length = 0;
		}
		mFields[field_index].ToToken(token[param_index]);
	}
	
	ExprTokenType **param_ptr = param_list;

	// Init the array of param token pointers.
	for (param_index = 0; param_index < aParamCount; ++param_index)
		*param_ptr++ = aParam[param_index]; // Caller-supplied param token.
	for (param_index = 0; param_index < extra_params; ++param_index)
		*param_ptr++ = &token[param_index]; // New param.
}


//
// Object::ArrayToStrings - Used by BIF_StrSplit.
//

ResultType Object::ArrayToStrings(LPTSTR *aStrings, int &aStringCount, int aStringsMax)
{
	int i, j;
	for (i = 0, j = 0; i < aStringsMax && j < mKeyOffsetObject; ++j)
		if (SYM_STRING == mFields[j].symbol)
			aStrings[i++] = mFields[j].string;
		else
			return FAIL;
	aStringCount = i;
	return OK;
}


//
// Object::Delete - Called immediately before the object is deleted.
//					Returns false if object should not be deleted yet.
//

bool Object::Delete()
{
	if (mBase)
	{
		KeyType key;
		IndexType insert_pos;
		key.s = _T("__Class");
		if (FindField(SYM_STRING, key, insert_pos))
			// This object appears to be a class definition, so it would probably be
			// undesirable to call the super-class' __Delete() meta-function for this.
			return ObjectBase::Delete();

		// L33: Privatize the last recursion layer's deref buffer in case it is in use by our caller.
		// It's done here rather than in Var::FreeAndRestoreFunctionVars (even though the below might
		// not actually call any script functions) because this function is probably executed much
		// less often in most cases.
		PRIVATIZE_S_DEREF_BUF;

		// If an exception has been thrown, temporarily clear it for execution of __Delete.
		ResultToken *exc = g->ThrownToken;
		g->ThrownToken = NULL;

		CallMethod(mBase, this, _T("__Delete"), NULL, 0, NULL, IF_METAOBJ); // base.__Delete()

		// Reset any exception cleared above.
		if (exc)
		{
			if (g->ThrownToken)
				// Let the original exception take precedence over this secondary exception.
				g_script.FreeExceptionToken(g->ThrownToken);
			g->ThrownToken = exc;
		}

		DEPRIVATIZE_S_DEREF_BUF; // L33: See above.

		// Above may pass the script a reference to this object to allow cleanup routines to free any
		// associated resources.  Deleting it is only safe if the script no longer holds any references
		// to it.  Since cleanup routines may (intentionally or unintentionally) copy this reference,
		// ensure this object really has no more references before proceeding with deletion:
		if (mRefCount > 1)
			return false;
	}
	return ObjectBase::Delete();
}


Object::~Object()
{
	if (mBase)
		mBase->Release();

	if (mFields)
	{
		if (mFieldCount)
		{
			IndexType i = mFieldCount - 1;
			// Free keys: first strings, then objects (objects have a lower index in the mFields array).
			for ( ; i >= mKeyOffsetString; --i)
				free(mFields[i].key.s);
			for ( ; i >= mKeyOffsetObject; --i)
				mFields[i].key.p->Release();
			// Free values.
			while (mFieldCount) 
				mFields[--mFieldCount].Free();
		}
		// Free fields array.
		free(mFields);
	}
}


//
// Object::Invoke
//

ResultType STDMETHODCALLTYPE Object::Invoke(
                                            ResultToken &aResultToken,
                                            ExprTokenType &aThisToken,
                                            int aFlags,
                                            ExprTokenType *aParam[],
                                            int aParamCount
                                            )
// L40: Revised base mechanism for flexibility and to simplify some aspects.
//		obj[] -> obj.base.__Get -> obj.base[] -> obj.base.__Get etc.
{
	SymbolType key_type;
	KeyType key;
    FieldType *field, *prop_field;
	IndexType insert_pos;
	Property *prop = NULL; // Set default.

	// If this is some object's base and is being invoked in that capacity, call
	//	__Get/__Set/__Call as defined in this base object before searching further.
	if (SHOULD_INVOKE_METAFUNC)
	{
		key.s = sMetaFuncName[INVOKE_TYPE];
		// Look for a meta-function definition directly in this base object.
		if (field = FindField(SYM_STRING, key, /*out*/ insert_pos))
		{
			// Seems more maintainable to copy params rather than assume aParam[-1] is always valid.
			ExprTokenType **meta_params = (ExprTokenType **)_alloca((aParamCount + 1) * sizeof(ExprTokenType *));
			// Shallow copy; points to the same tokens.  Leave a space for param[0], which must be the param which
			// identified the field (or in this case an empty space) to replace with aThisToken when appropriate.
			memcpy(meta_params + 1, aParam, aParamCount * sizeof(ExprTokenType*));

			Line *curr_line = g_script.mCurrLine;
			ResultType r = CallField(field, aResultToken, aThisToken, aFlags, meta_params, aParamCount + 1);
			g_script.mCurrLine = curr_line; // Allows exceptions thrown by later meta-functions to report a more appropriate line.
			//if (r == EARLY_RETURN)
				// Propagate EARLY_RETURN in case this was the __Call meta-function of a
				// "function object" which is used as a meta-function of some other object.
				//return EARLY_RETURN; // TODO: Detection of 'return' vs 'return empty_value'.
			if (r != OK) // Likely EARLY_RETURN, FAIL or EARLY_EXIT.
				return r;
		}
	}
	
	int param_count_excluding_rvalue = aParamCount;

	if (IS_INVOKE_SET)
	{
		// Due to the way expression parsing works, the result should never be negative
		// (and any direct callers of Invoke must always pass aParamCount >= 1):
		--param_count_excluding_rvalue;
	}
	
	if (param_count_excluding_rvalue && aParam[0]->symbol != SYM_MISSING)
	{
		field = FindField(*aParam[0], _f_number_buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos);

		static Property sProperty;

		// v1.1.16: Handle class property accessors:
		if (field && field->symbol == SYM_OBJECT && *(void **)field->object == *(void **)&sProperty)
		{
			// The "type check" above is used for speed.  Simple benchmarks of x[1] where x := [[]]
			// shows this check to not affect performance, whereas dynamic_cast<> hurt performance by
			// about 25% and typeid()== by about 20%.  We can safely assume that the vtable pointer is
			// stored at the beginning of the object even though it isn't guaranteed by the C++ standard,
			// since COM fundamentally requires it:  http://msdn.microsoft.com/en-us/library/dd757710
			prop = (Property *)field->object;
			prop_field = field;
			if (IS_INVOKE_SET ? prop->CanSet() : prop->CanGet())
			{
				if (aParamCount > 2 && IS_INVOKE_SET)
				{
					// Do some shuffling to put value before the other parameters.  This relies on above
					// having verified that we're handling this invocation; otherwise the parameters would
					// need to be swapped back later in case they're passed to a base's meta-function.
					ExprTokenType *value = aParam[aParamCount - 1];
					for (int i = aParamCount - 1; i > 1; --i)
						aParam[i] = aParam[i - 1];
					aParam[1] = value; // Corresponds to the setter's hidden "value" parameter.
				}
				ExprTokenType *name_token = aParam[0];
				aParam[0] = &aThisToken; // For the hidden "this" parameter in the getter/setter.
				// Pass IF_FUNCOBJ so that it'll pass all parameters to the getter/setter.
				// For a functor Object, we would need to pass a token representing "this" Property,
				// but since Property::Invoke doesn't use it, we pass our aThisToken for simplicity.
				ResultType result = prop->Invoke(aResultToken, aThisToken, aFlags | IF_FUNCOBJ, aParam, aParamCount);
				aParam[0] = name_token;
				return result == EARLY_RETURN ? OK : result;
			}
			// The property was missing get/set (whichever this invocation is), so continue as
			// if the property itself wasn't defined.
			key_type = SYM_INVALID;
			field = NULL;
		}
	}
	else
	{
		key_type = SYM_INVALID; // Allow key_type checks below without requiring that param_count_excluding_rvalue also be checked.
		field = NULL;
	}
	
	if (!field)
	{
		// This field doesn't exist, so let our base object define what happens:
		//		1) __Get, __Set or __Call.  If these don't return a value, processing continues.
		//		2) For GET and CALL only, check the base object's own fields.
		//		3) Repeat 1 through 3 for the base object's own base.
		if (mBase)
		{
			// aFlags: If caller specified IF_METAOBJ but not IF_METAFUNC, they want to recursively
			// find and execute a specific meta-function (__new or __delete) but don't want any base
			// object to invoke __call.  So if this is already a meta-invocation, don't change aFlags.
			ResultType r = mBase->Invoke(aResultToken, aThisToken, aFlags | (IS_INVOKE_META ? 0 : IF_META), aParam, aParamCount);
			if (r != INVOKE_NOT_HANDLED // Base handled it.
				|| key_type == SYM_INVALID) // Nothing left to do in this case.
				return r;

			// Since the above may have inserted or removed fields (including the specified one),
			// insert_pos may no longer be correct or safe.  Updating field also allows a meta-function
			// to initialize a field and allow processing to continue as if it already existed.
			field = FindField(key_type, key, /*out*/ insert_pos);
			if (prop)
			{
				// This field was a property.
				if (field && field->symbol == SYM_OBJECT && field->object == prop)
				{
					// This field is still a property (and the same one).
					prop_field = field; // Must update this pointer in case the field is to be overwritten.
					field = NULL; // Act like the field doesn't exist (until the time comes to insert a value).
				}
				else
					prop = NULL; // field was reassigned or removed, so ignore the property.
			}
		}

		// Since the base object didn't handle this op, check for built-in properties/methods.
		// This must apply only to the original target object (aThisToken), not one of its bases.
		if (!IS_INVOKE_META && key_type == SYM_STRING && !field) // v1.1.16: Check field again so if __Call sets a field, it gets called.
		{
			//
			// BUILT-IN METHODS
			//
			if (IS_INVOKE_CALL)
			{
				// Since above has not handled this call and no field exists, check for built-in methods.
				return CallBuiltin(GetBuiltinID(key.s), aResultToken, aParam + 1, aParamCount - 1); // +/- 1 to exclude the method identifier.
			}
			//
			// BUILT-IN "BASE" PROPERTY
			//
			else if (param_count_excluding_rvalue == 1)
			{
				if (!_tcsicmp(key.s, _T("base")))
				{
					if (IS_INVOKE_SET)
					{
						IObject *obj = TokenToObject(*aParam[1]);
						if (obj)
							obj->AddRef(); // for mBase
						if (mBase)
							mBase->Release();
						mBase = obj; // May be NULL.
					}
	
					if (mBase)
					{
						mBase->AddRef();
						_o_return(mBase);
					}
					//else return empty string.
					_o_return_empty;
				}
				else if (!_tcsicmp(key.s, _T("Length")) && IS_INVOKE_GET)
				{
					return _Length(aResultToken);
				}
			}
		} // if (!IS_INVOKE_META && key_type == SYM_STRING)
	} // if (!field)

	//
	// OPERATE ON A FIELD WITHIN THIS OBJECT
	//

	// CALL
	if (IS_INVOKE_CALL)
	{
		if (!field)
			return INVOKE_NOT_HANDLED;
		// v1.1.18: The following flag is set whenever a COM client invokes with METHOD|PROPERTYGET,
		// such as X.Y in VBScript or C#.  Some convenience is gained at the expense of purity by treating
		// it as METHOD if X.Y is a Func object or PROPERTYGET in any other case.
		// v1.1.19: Handling this flag here rather than in CallField() has the following benefits:
		//  - Reduces code duplication.
		//  - Fixes X.__Call being returned instead of being called, if X.__Call is a string.
		//  - Allows X.Y(Z) and similar to work like X.Y[Z], instead of ignoring the extra parameters.
		if ( !(aFlags & IF_CALL_FUNC_ONLY) || (field->symbol == SYM_OBJECT && dynamic_cast<Func *>(field->object)) )
			return CallField(field, aResultToken, aThisToken, aFlags, aParam, aParamCount);
		aFlags = (aFlags & ~(IT_BITMASK | IF_CALL_FUNC_ONLY)) | IT_GET;
	}

	// MULTIPARAM[x,y] -- may be SET[x,y]:=z or GET[x,y], but always treated like GET[x].
	if (param_count_excluding_rvalue > 1)
	{
		// This is something like this[x,y] or this[x,y]:=z.  Since it wasn't handled by a meta-mechanism above,
		// handle only the "x" part (automatically creating and storing an object if this[x] didn't already exist
		// and an assignment is being made) and recursively invoke.  This has at least two benefits:
		//	1) Objects natively function as multi-dimensional arrays.
		//	2) Automatic initialization of object-fields.
		//		For instance, this["base","__Get"]:="MyObjGet" does not require a prior this.base:=Object().
		IObject *obj = NULL;
		if (field)
		{
			if (field->symbol == SYM_OBJECT)
				// AddRef not used.  See below.
				obj = field->object;
			else if (!IS_INVOKE_META)
				_o_throw(ERR_ARRAY_NOT_MULTIDIM);
		}
		else if (!IS_INVOKE_META)
		{
			// This section applies only to the target object (aThisToken) and not any of its base objects.
			// Allow obj["base",x] to access a field of obj.base; L40: This also fixes obj.base[x] which was broken by L36.
			if (key_type == SYM_STRING && !_tcsicmp(key.s, _T("base")))
			{
				if (!mBase && IS_INVOKE_SET)
					mBase = new Object();
				obj = mBase; // If NULL, above failed and below will detect it.
			}
			// Automatically create a new object for the x part of obj[x,y]:=z.
			else if (IS_INVOKE_SET)
			{
				Object *new_obj = new Object();
				if (new_obj)
				{
					if ( field = prop ? prop_field : Insert(key_type, key, insert_pos) )
					{
						if (prop) // Otherwise, field is already empty.
							prop->Release();
						// Don't do field->Assign() since it would do AddRef() and we would need to counter with Release().
						field->symbol = SYM_OBJECT;
						field->object = obj = new_obj;
					}
					else
					{	// Create() succeeded but Insert() failed, so free the newly created obj.
						new_obj->Release();
						new_obj = NULL;
					}
				}
				if (!new_obj)
					_o_throw(ERR_OUTOFMEM);
			}
			else if (IS_INVOKE_GET)
			{
				// Treat x[y,z] like x[y] when x[y] is not set: just return "", don't throw an exception.
				// On the other hand, if x[y] is set to something which is not an object, the "if (field)"
				// section above raises an error.
				_o_return_empty;
			}
		}
		if (obj) // Object was successfully found or created.
		{
			// obj now contains a pointer to the object contained by this field, possibly newly created above.
			ExprTokenType obj_token(obj);
			// References in obj_token and obj weren't counted (AddRef wasn't called), so Release() does not
			// need to be called before returning, and accessing obj after calling Invoke() would not be safe
			// since it could Release() the object (by overwriting our field via script) as a side-effect.
			// Recursively invoke obj, passing remaining parameters; remove IF_META to correctly treat obj as target:
			return obj->Invoke(aResultToken, obj_token, aFlags & ~IF_META, aParam + 1, aParamCount - 1);
			// Above may return INVOKE_NOT_HANDLED in cases such as obj[a,b] where obj[a] exists but obj[a][b] does not.
		}
	} // MULTIPARAM[x,y]

	// SET
	else if (IS_INVOKE_SET)
	{
		if (!IS_INVOKE_META && param_count_excluding_rvalue)
		{
			ExprTokenType &value_param = *aParam[1];
			// L34: Assigning an empty string no longer removes the field.
			if ( (field || (field = prop ? prop_field : Insert(key_type, key, insert_pos)))
				&& field->Assign(value_param) )
			{
				// See the similar call below for comments.
				field->Get(aResultToken);
				return OK;
			}
			_o_throw(ERR_OUTOFMEM);
		}
	}

	// GET
	else
	{
		if (field)
		{
			// Caller takes care of copying the result into persistent memory when necessary, and must
			// ensure this is done before they Release() this object.  For ExpandExpression(), there are
			// two different danger scenarios:
			//   1) Command % {value:"string"}.value  ; Temporary object could be released prematurely.
			//   2) Fn( obj.value, obj := "" )        ; Object is freed by the assignment.
			// For #1, SYM_OBJECT refs are released after the result of the expression is copied into the
			// deref buf (as of commit d1ab199).  For #2, the value is copied immediately after we return,
			// because the result of any BIF is assumed to be volatile if expression eval isn't finished.
			field->Get(aResultToken);
			return OK;
		}
		// If 'this' is the target object (not its base), produce OK so that something like if(!foo.bar) is
		// considered valid even when foo.bar has not been set.
		if (!IS_INVOKE_META && aParamCount)
			_o_return_empty;
	}

	// Fell through from one of the sections above: invocation was not handled.
	return INVOKE_NOT_HANDLED;
}


int Object::GetBuiltinID(LPCTSTR aName)
{
	switch (toupper(*aName))
	{
	case 'I':
		if (!_tcsicmp(aName, _T("InsertAt")))
			return FID_ObjInsertAt;
		break;
	case 'R':
		if (!_tcsicmp(aName, _T("RemoveAt")))
			return FID_ObjRemoveAt;
		break;
	case 'D':
		if (!_tcsicmp(aName, _T("Delete")))
			return FID_ObjDelete;
		break;
	case 'P':
		if (!_tcsicmp(aName, _T("Push")))
			return FID_ObjPush;
		if (!_tcsicmp(aName, _T("Pop")))
			return FID_ObjPop;
		break;
	case 'H':
		if (!_tcsicmp(aName, _T("HasKey")))
			return FID_ObjHasKey;
		break;
	case 'L':
		if (!_tcsicmp(aName, _T("Length")))
			return FID_ObjLength;
		break;
	case '_':
		if (!_tcsicmp(aName, _T("_NewEnum")))
			return FID_ObjNewEnum;
		break;
	case 'G':
		if (!_tcsicmp(aName, _T("GetAddress")))
			return FID_ObjGetAddress;
		if (!_tcsicmp(aName, _T("GetCapacity")))
			return FID_ObjGetCapacity;
		break;
	case 'S':
		if (!_tcsicmp(aName, _T("SetCapacity")))
			return FID_ObjSetCapacity;
		break;
	case 'C':
		if (!_tcsicmp(aName, _T("Clone")))
			return FID_ObjClone;
		break;
	case 'M':
		if (!_tcsicmp(aName, _T("MaxIndex")))
			return FID_ObjMaxIndex;
		if (!_tcsicmp(aName, _T("MinIndex")))
			return FID_ObjMinIndex;
		break;
	}
	return -1;
}


ResultType Object::CallBuiltin(int aID, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case FID_ObjInsertAt:		return _InsertAt(aResultToken, aParam, aParamCount);
	case FID_ObjDelete:			return _Delete(aResultToken, aParam, aParamCount);
	case FID_ObjRemoveAt:		return _RemoveAt(aResultToken, aParam, aParamCount);
	case FID_ObjPush:			return _Push(aResultToken, aParam, aParamCount);
	case FID_ObjPop:			return _Pop(aResultToken, aParam, aParamCount);
	case FID_ObjLength:			return _Length(aResultToken);
	case FID_ObjHasKey:			return _HasKey(aResultToken, aParam, aParamCount);
	case FID_ObjGetCapacity:	return _GetCapacity(aResultToken, aParam, aParamCount);
	case FID_ObjSetCapacity:	return _SetCapacity(aResultToken, aParam, aParamCount);
	case FID_ObjGetAddress:		return _GetAddress(aResultToken, aParam, aParamCount);
	case FID_ObjClone:			return _Clone(aResultToken, aParam, aParamCount);
	case FID_ObjNewEnum:		return _NewEnum(aResultToken, aParam, aParamCount);
	case FID_ObjMaxIndex:		return _MaxIndex(aResultToken);
	case FID_ObjMinIndex:		return _MinIndex(aResultToken);
	}
	return INVOKE_NOT_HANDLED;
}


//
// Internal: Object::CallField - Used by Object::Invoke to call a function/method stored in this object.
//

ResultType Object::CallField(FieldType *aField, ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
// aParam[0] contains the identifier of this field or an empty space (for __Get etc.).
{
	if (aField->symbol == SYM_OBJECT)
	{
		// Allocate a new array of param pointers that we can modify.
		ExprTokenType **params = (ExprTokenType **)_alloca((aParamCount + 1) * sizeof(ExprTokenType *));
		// Shallow copy; points to the same tokens.  Skip aParam[0], which contains
		// the key used to find aField.  We want to invoke aField.Call(this, aParams*).
		memcpy(params + 2, aParam + 1, (aParamCount - 1) * sizeof(ExprTokenType*));
		// Where fn = this[key], call fn.call(this, params*).
		ExprTokenType field_token(aField->object); // fn
		ExprTokenType method_name(_T("call"), 4); // Works with JScript functions as well, unlike "Call".
		params[0] = &method_name; // call
		params[1] = &aThisToken; // this
		return aField->object->Invoke(aResultToken, field_token, IT_CALL, params, aParamCount + 1);
	}
	if (aField->symbol == SYM_STRING)
	{
		Func *func = g_script.FindFunc(aField->string, aField->string.Length());
		if (func)
		{
			ExprTokenType *tmp = aParam[0];
			// v2: Always pass "this" as the first parameter.  The old behaviour of passing it only when called
			// indirectly via mBase was confusing to many users, and isn't needed now that the script can do
			// %this.func%() instead of this.func() if they don't want to pass "this".
			// For this type of call, "this" object is included as the first parameter.  To do this, aParam[0] is
			// temporarily overwritten with a pointer to aThisToken.  Note that aThisToken contains the original
			// object specified in script, not the C++ "this" which is actually a meta-object/base of that object.
			aParam[0] = &aThisToken;
			func->Call(aResultToken, aParam, aParamCount);
			aParam[0] = tmp;
			return aResultToken.Result();
		}
	}
	// The field's value is neither a function reference nor the name of a known function.
	ExprTokenType tok;
	aField->ToToken(tok);
	_o_throw(ERR_NONEXISTENT_FUNCTION, TokenToString(tok, _f_number_buf));
}


//
// Helper function for WinMain()
//

Object *Object::CreateFromArgV(LPTSTR *aArgV, int aArgC)
{
	ExprTokenType *token = (ExprTokenType *)_alloca(aArgC * sizeof(ExprTokenType));
	ExprTokenType **param = (ExprTokenType **)_alloca(aArgC * sizeof(ExprTokenType*));
	for (int j = 0; j < aArgC; ++j)
	{
		token[j].SetValue(aArgV[j]);
		param[j] = &token[j];
	}
	return CreateArray(param, aArgC);
}



//
// Helper function for StrSplit/WinGetList/WinGetControls
//

bool Object::Append(ExprTokenType &aValue)
{
	if (mFieldCount == mFieldCountMax && !Expand()) // Attempt to expand if at capacity.
		return false;

	FieldType &field = mFields[mKeyOffsetObject];
	if (mKeyOffsetObject < mFieldCount)
		// For maintainability. This might never be done, because our caller
		// doesn't use string/object keys. Move existing fields to make room:
		memmove(&field + 1, &field, (mFieldCount - mKeyOffsetObject) * sizeof(FieldType));
	++mFieldCount; // Only after memmove above.
	++mKeyOffsetObject;
	++mKeyOffsetString;
	
	// The following relies on the fact that callers of this function ONLY use
	// this function, so the last integer key == the number of integer keys.
	field.key.i = mKeyOffsetObject;

	field.symbol = SYM_INVALID; // Init for Assign(): indicate that it does not contain a valid string or object.
	return field.Assign(aValue);
}


//
// Helper function used with class definitions.
//

void Object::EndClassDefinition()
{
	// Instance variables were previously created as keys in the class object to prevent duplicate or
	// conflicting declarations.  Since these variables will be added at run-time to the derived objects,
	// we don't want them in the class object.  So delete any key-value pairs with the special marker
	// value (currently any integer, since static initializers haven't been evaluated yet).
	for (IndexType i = mFieldCount - 1; i >= 0; --i)
		if (mFields[i].symbol == SYM_INTEGER)
		{
			if (i >= mKeyOffsetString) // Must be checked since key can be an integer, such as for "0 := (expr)".
				free(mFields[i].key.s);
			if (i < --mFieldCount)
				memmove(mFields + i, mFields + i + 1, (mFieldCount - i) * sizeof(FieldType));
		}
}


//
// Helper function for 'is' operator: is aBase a direct or indirect base object of this?
//

bool Object::IsDerivedFrom(IObject *aBase)
{
	IObject *ibase;
	Object *base;
	for (ibase = mBase; ; ibase = base->mBase)
	{
		if (ibase == aBase)
			return true;
		if (  !(base = dynamic_cast<Object *>(ibase))  )  // ibase may be NULL.
			return false;
	}
}
	

//
// Object:: Built-in Methods
//

bool Object::InsertAt(INT_PTR aOffset, INT_PTR aKey, ExprTokenType *aValue[], int aValueCount)
{
	IndexType actual_count = (IndexType)aValueCount;
	for (int i = 0; i < aValueCount; ++i)
		if (aValue[i]->symbol == SYM_MISSING)
			actual_count--;
	IndexType need_capacity = mFieldCount + actual_count;
	if (need_capacity > mFieldCountMax && !SetInternalCapacity(need_capacity))
		// Fail.
		return false;
	FieldType *field = mFields + aOffset;
	if (aOffset < mFieldCount)
		memmove(field + actual_count, field, (mFieldCount - aOffset) * sizeof(FieldType));
	mFieldCount += actual_count;
	mKeyOffsetObject += actual_count; // ints before objects
	mKeyOffsetString += actual_count; // and strings
	FieldType *field_end;
	// Set keys and copy value params into the fields.
	for (int i = 0; i < aValueCount; ++i, ++aKey)
	{
		ExprTokenType &value = *(aValue[i]);
		if (value.symbol != SYM_MISSING)
		{
			field->key.i = aKey;
			field->symbol = SYM_INVALID; // Init for Assign(): indicate that it does not contain a valid string or object.
			field->Assign(value);
			field++;
		}
	}
	// Adjust keys of fields which have been moved.
	for (field_end = mFields + mKeyOffsetObject; field < field_end; ++field)
	{
		field->key.i += aValueCount; // NOT =++key.i since keys might not be contiguous.
	}
	return true;
}

ResultType Object::_InsertAt(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
// InsertAt(index, value1, ...)
{
	if (aParamCount < 2)
		_o_throw(ERR_TOO_FEW_PARAMS);

	SymbolType key_type;
	KeyType key;
	IndexType insert_pos;
	FieldType *field = FindField(**aParam, _f_number_buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos);
	if (key_type != SYM_INTEGER)
		_o_throw(ERR_PARAM1_INVALID, key_type == SYM_STRING ? key.s : _T(""));
		
	if (field)
	{
		insert_pos = field - mFields; // insert_pos wasn't set in this case.
		field = NULL; // Insert, don't overwrite.
	}

	if (!InsertAt(insert_pos, key.i, aParam + 1, aParamCount - 1))
		_o_throw(ERR_OUTOFMEM);
	
	_o_return_empty; // No return value.
}

ResultType Object::_Push(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
// Push(value1, ...)
{
	IndexType insert_pos = mKeyOffsetObject; // int keys end here.;
	IntKeyType start_index = (insert_pos ? mFields[insert_pos - 1].key.i + 1 : 1);
	if (!InsertAt(insert_pos, start_index, aParam, aParamCount))
		_o_throw(ERR_OUTOFMEM);

	// Return the new "length" of the array.
	_o_return(start_index + aParamCount - 1);
}

ResultType Object::_Remove_impl(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, RemoveMode aMode)
// Remove(first_key [, last_key := first_key])
// RemoveAt(index [, virtual_count := 1])
// Pop()
{
	FieldType *min_field;
	IndexType min_pos, max_pos, pos;
	SymbolType min_key_type;
	KeyType min_key, max_key;
	IntKeyType logical_count_removed = 1;

	LPTSTR number_buf = _f_number_buf;

	// Find the position of "min".
	if (aMode == RM_Pop)
	{
		if (mKeyOffsetObject) // i.e. at least one int field; use Length()
		{
			min_field = &mFields[min_pos = mKeyOffsetObject - 1];
			min_key = min_field->key;
			min_key_type = SYM_INTEGER;
		}
		else // No appropriate field to remove, just return "".
			_o_return_empty;
	}
	else
	{
		if (!aParamCount)
			_o_throw(ERR_TOO_FEW_PARAMS);

		if (min_field = FindField(*aParam[0], number_buf, min_key_type, min_key, min_pos))
			min_pos = min_field - mFields; // else min_pos was already set by FindField.
		
		if (min_key_type != SYM_INTEGER && aMode != RM_RemoveKey)
			_o_throw(ERR_PARAM1_INVALID);
	}
	
	if (aParamCount > 1) // Removing a range of keys.
	{
		SymbolType max_key_type;
		FieldType *max_field;
		if (aMode == RM_RemoveAt)
		{
			logical_count_removed = (IntKeyType)TokenToInt64(*aParam[1]);
			// Find the next position >= [aParam[1] + Count].
			max_key_type = SYM_INTEGER;
			max_key.i = min_key.i + logical_count_removed;
			if (max_field = FindField(max_key_type, max_key, max_pos))
				max_pos = max_field - mFields;
		}
		else
		{
			// Find the next position > [aParam[1]].
			if (max_field = FindField(*aParam[1], number_buf, max_key_type, max_key, max_pos))
				max_pos = max_field - mFields + 1;
		}
		// Since the order of key-types in mFields is of no logical consequence, require that both keys be the same type.
		// Do not allow removing a range of object keys since there is probably no meaning to their order.
		if (max_key_type != min_key_type || max_key_type == SYM_OBJECT || max_pos < min_pos
			// min and max are different types, are objects, or max < min.
			|| (max_pos == min_pos && (max_key_type == SYM_INTEGER ? max_key.i < min_key.i : _tcsicmp(max_key.s, min_key.s) < 0)))
			// max < min, but no keys exist in that range so (max_pos < min_pos) check above didn't catch it.
		{
			_o_throw(ERR_PARAM2_INVALID);
		}
		//else if (max_pos == min_pos): specified range is valid, but doesn't match any keys.
		//	Continue on, adjust integer keys as necessary and return 0.
	}
	else // Removing a single item.
	{
		if (!min_field) // Nothing to remove.
		{
			if (aMode == RM_RemoveAt)
				for (pos = min_pos; pos < mKeyOffsetObject; ++pos)
					mFields[pos].key.i--;
			// Our return value when only one key is given is supposed to be the value
			// previously at this[key], which has just been removed.  Since this[key]
			// would return "", it makes sense to return the same in this case.
			_o_return_empty;
		}
		// Since only one field (at maximum) can be removed in this mode, it
		// seems more useful to return the field being removed than a count.
		switch (aResultToken.symbol = min_field->symbol)
		{
		case SYM_STRING:
			// For simplicity, just discard the memory of min_field->marker (can't return it as-is
			// since the string isn't at the start of its memory block).  Scripts can use the 2-param
			// mode to avoid any performance penalty this may incur.
			TokenSetResult(aResultToken, min_field->string, min_field->string.Length());
			break;
		case SYM_OBJECT:
			aResultToken.object = min_field->object;
			min_field->symbol = SYM_INVALID; // Prevent Free() from calling object->Release(), since caller is taking ownership of the ref.
			break;
		//case SYM_INTEGER:
		//case SYM_FLOAT:
		default:
			aResultToken.value_int64 = min_field->n_int64; // Effectively also value_double = n_double.
		}
		// If the key is an object, release it now because Free() doesn't.
		// Note that object keys can only be removed in the single-item mode.
		if (min_key_type == SYM_OBJECT)
			min_field->key.p->Release();
		// Set max_pos for the loops below.
		max_pos = min_pos + 1;
	}

	for (pos = min_pos; pos < max_pos; ++pos)
		// Free each field in the range being removed.
		mFields[pos].Free();

	if (min_key_type == SYM_STRING)
		// Free all string keys in the range being removed.
		for (pos = min_pos; pos < max_pos; ++pos)
			free(mFields[pos].key.s);

	IndexType remaining_fields = mFieldCount - max_pos;
	if (remaining_fields)
		// Move remaining fields left to fill the gap left by the removed range.
		memmove(mFields + min_pos, mFields + max_pos, remaining_fields * sizeof(FieldType));
	// Adjust count by the actual number of fields in the removed range.
	IndexType actual_count_removed = max_pos - min_pos;
	mFieldCount -= actual_count_removed;
	// Adjust key offsets and numeric keys as necessary.
	if (min_key_type != SYM_STRING) // i.e. SYM_OBJECT or SYM_INTEGER
	{
		mKeyOffsetString -= actual_count_removed;
		if (min_key_type == SYM_INTEGER)
		{
			mKeyOffsetObject -= actual_count_removed;
			if (aMode == RM_RemoveAt)
			{
				// Regardless of whether any fields were removed, min_pos contains the position of the field which
				// immediately followed the specified range.  Decrement each numeric key from this position onward.
				if (logical_count_removed > 0)
					for (pos = min_pos; pos < mKeyOffsetObject; ++pos)
						mFields[pos].key.i -= logical_count_removed;
			}
		}
	}
	if (aParamCount > 1)
	{
		// Return actual number of fields removed:
		_o_return(actual_count_removed);
	}
	//else result was set above.
	return OK;
}

ResultType Object::_Delete(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	return _Remove_impl(aResultToken, aParam, aParamCount, RM_RemoveKey);
}

ResultType Object::_RemoveAt(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	return _Remove_impl(aResultToken, aParam, aParamCount, RM_RemoveAt);
}

ResultType Object::_Pop(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	// Unwanted parameters are ignored, as is conventional for dynamic calls.
	// _Remove_impl relies on aParamCount == 0 for Pop().
	return _Remove_impl(aResultToken, NULL, 0, RM_Pop);
}


ResultType Object::_Length(ResultToken &aResultToken)
{
	IntKeyType max_index = mKeyOffsetObject ? mFields[mKeyOffsetObject - 1].key.i : 0;
	_o_return(max_index > 0 ? max_index : 0);
}

ResultType Object::_MaxIndex(ResultToken &aResultToken)
{
	if (mKeyOffsetObject)
		_o_return(mFields[mKeyOffsetObject - 1].key.i);
	else
		_o_return_empty;
}

ResultType Object::_MinIndex(ResultToken &aResultToken)
{
	if (mKeyOffsetObject)
		_o_return(mFields[0].key.i);
	else
		_o_return_empty;
}

ResultType Object::_GetCapacity(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount)
	{
		SymbolType key_type;
		KeyType key;
		IndexType insert_pos;
		FieldType *field;

		if ( (field = FindField(*aParam[0], _f_number_buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos))
			&& field->symbol == SYM_STRING )
		{
			size_t size = field->string.Capacity();
			_o_return(size ? _TSIZE(size - 1) : 0); // -1 to exclude null-terminator.
		}
		//else: wrong type of field, so return an empty string.
	}
	else // aParamCount == 0
	{
		_o_return(mFieldCountMax);
	}
	_o_return_empty;
}

ResultType Object::_SetCapacity(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
// SetCapacity([field_name,] new_capacity)
{
	if (!aParamCount || !TokenIsNumeric(*aParam[aParamCount - 1]))
		_o_throw(ERR_PARAM_INVALID);

	__int64 desired_capacity = TokenToInt64(*aParam[aParamCount - 1]);
	if (aParamCount >= 2) // Field name was specified.
	{
		if (desired_capacity < 0) // Check before sign is dropped.
			_o_throw(ERR_PARAM2_INVALID);
		size_t desired_size = (size_t)desired_capacity;

		SymbolType key_type;
		KeyType key;
		IndexType insert_pos;
		FieldType *field;

		if ( (field = FindField(*aParam[0], _f_number_buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos))
			|| (field = Insert(key_type, key, insert_pos)) )
		{	
			// Field was successfully found or inserted.
			if (field->symbol != SYM_STRING)
				// Wrong type of field.
				_o_throw(ERR_INVALID_VALUE);
			if (!desired_size)
			{
				// Caller specified zero - empty the field but do not remove it.
				field->Clear();
				_o_return(0);
			}
#ifdef UNICODE
			// Convert size in bytes to size in chars.
			desired_size = (desired_size >> 1) + (desired_size & 1);
#endif
			size_t length = field->string.Length();
			if (desired_size < length)
				desired_size = length; // Consistent with Object.SetCapacity(n): never truncate data.
			// Unlike VarSetCapacity: allow fields to shrink, preserves existing data.
			if (field->string.SetCapacity(desired_size + 1)) // Like VarSetCapacity, reserve one char for the null-terminator.
			{
				// Return new size, excluding null-terminator.
				_o_return(_TSIZE(desired_size));
			}
			// Since above didn't return, it failed.
		}
		// Out of memory.  Throw an error, otherwise scripts might assume success and end up corrupting the heap:
		_o_throw(ERR_OUTOFMEM);
	}
	// else aParamCount == 1: set the capacity of this object.
	IndexType desired_count = (IndexType)desired_capacity;
	if (desired_count < mFieldCount)
	{
		// It doesn't seem intuitive to allow SetCapacity to truncate the fields array, so just reallocate
		// as necessary to remove any unused space.  Allow negative values since SetCapacity(-1) seems more
		// intuitive than SetCapacity(0) when the contents aren't being discarded.
		desired_count = mFieldCount;
	}
	if (!desired_count)
	{	// Caller wants to shrink object to current contents but there aren't any, so free mFields.
		if (mFields)
		{
			free(mFields);
			mFields = NULL;
			mFieldCountMax = 0;
		}
		//else mFieldCountMax should already be 0.
		// Since mFieldCountMax and desired_size are both 0, below will return 0 and won't call SetInternalCapacity.
	}
	if (desired_count == mFieldCountMax || SetInternalCapacity(desired_count))
	{
		_o_return(mFieldCountMax);
	}
	// At this point, failure isn't critical since nothing is being stored yet.  However, it might be easier to
	// debug if an error is thrown here rather than possibly later, when the array attempts to resize itself to
	// fit new items.  This also avoids the need for scripts to check if the return value is less than expected:
	_o_throw(ERR_OUTOFMEM);
}

ResultType Object::_GetAddress(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
// GetAddress(key)
{
	if (!aParamCount)
		_o_throw(ERR_TOO_FEW_PARAMS);
	
	SymbolType key_type;
	KeyType key;
	IndexType insert_pos;
	FieldType *field;

	if ( (field = FindField(*aParam[0], _f_number_buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos))
		&& field->symbol == SYM_STRING && field->string.Capacity() != 0 )
	{
		_o_return((__int64)field->string.Value());
	}
	//else: field has no memory allocated, so return an empty string.
	_o_return_empty;
}

ResultType Object::_NewEnum(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (IObject *enm = new Enumerator(this))
		_o_return(enm);
	else
		_o_throw(ERR_OUTOFMEM);
}

ResultType Object::_HasKey(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (!aParamCount)
		_o_throw(ERR_TOO_FEW_PARAMS);
	
	SymbolType key_type;
	KeyType key;
	INT_PTR insert_pos;
	FieldType *field = FindField(*aParam[0], _f_number_buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos);
	_o_return(field != NULL);
}

ResultType Object::_Clone(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	Object *clone = Clone();
	if (!clone)
		_o_throw(ERR_OUTOFMEM);	
	if (mBase)
		(clone->mBase = mBase)->AddRef();
	_o_return(clone);
}


//
// Object::FieldType
//

void Object::FieldType::Clear()
{
	Free();
	symbol = SYM_STRING;
	string.Init();
}

bool Object::FieldType::Assign(LPTSTR str, size_t len, bool exact_size)
{
	if (len == -1)
		len = _tcslen(str);

	if (!len) // Check len, not *str, since it might be binary data or not null-terminated.
	{
		Clear();
		return true;
	}

	if (symbol != SYM_STRING || len >= string.Capacity())
	{
		Free(); // Free object or previous buffer (which was too small).
		symbol = SYM_STRING;
		string.Init();

		size_t new_size = len + 1;
		if (!exact_size)
		{
			// Use size calculations equivalent to Var:
			if (new_size < 16)
				new_size = 16; // 16 seems like a good size because it holds nearly any number.  It seems counterproductive to go too small because each malloc has overhead.
			else if (new_size < MAX_PATH)
				new_size = MAX_PATH;  // An amount that will fit all standard filenames seems good.
			else if (new_size < (160 * 1024)) // MAX_PATH to 160 KB or less -> 10% extra.
				new_size = (size_t)(new_size * 1.1);
			else if (new_size < (1600 * 1024))  // 160 to 1600 KB -> 16 KB extra
				new_size += (16 * 1024);
			else if (new_size < (6400 * 1024)) // 1600 to 6400 KB -> 1% extra
				new_size += (new_size / 100);
			else  // 6400 KB or more: Cap the extra margin at some reasonable compromise of speed vs. mem usage: 64 KB
				new_size += (64 * 1024);
		}
		if (!string.SetCapacity(new_size))
			return false; // And leave string empty (as set by Free() above).
	}
	// else we have a buffer with sufficient capacity already.

	LPTSTR buf = string.Value();
	tmemcpy(buf, str, len);
	buf[len] = '\0'; // Must be done separately since some callers pass a substring.
	string.Length() = len;
	return true; // Success.
}

bool Object::FieldType::Assign(ExprTokenType &aParam)
{
	ExprTokenType temp, *val; // Seems more maintainable to use a copy; avoid any possible side-effects.
	if (aParam.symbol == SYM_VAR)
	{
		// Primary reason for this check: If var has a cached binary integer, we want to use it and
		// not the stringified copy of it.  It seems unlikely that scripts will depend on the string
		// format of a literal number such as 0x123 or 00001, and even less likely for a number stored
		// in an object (even implicitly via a variadic function).  If the value is eventually passed
		// to a COM method call, it can be important that it is passed as VT_I4 and not VT_BSTR.
		aParam.var->ToTokenSkipAddRef(temp); // Skip AddRef() if applicable because it's called below.
		val = &temp;
	}
	else
		val = &aParam;

	switch (val->symbol)
	{
	case SYM_STRING:
		return Assign(val->marker, val->marker_length);
	case SYM_OBJECT:
		Free(); // Free string or object, if applicable.
		symbol = SYM_OBJECT; // Set symbol *after* calling Free().
		object = val->object;
		object->AddRef();
		break;
	//case SYM_INTEGER:
	//case SYM_FLOAT:
	default:
		Free(); // Free string or object, if applicable.
		symbol = val->symbol; // Either SYM_INTEGER or SYM_FLOAT.  Set symbol *after* calling Free().
		n_int64 = val->value_int64; // Also handles value_double via union.
		break;
	}
	return true;
}

void Object::FieldType::Get(ExprTokenType &result)
{
	switch (result.symbol = symbol) // Assign.
	{
	case SYM_STRING:
		result.marker = string;
		result.marker_length = string.Length();
		break;
	case SYM_OBJECT:
		object->AddRef();
		result.object = object;
		break;
	default:
		result.value_int64 = n_int64; // Union copy.
	}
}

void Object::FieldType::ToToken(ExprTokenType &aToken)
// Used when we want the value as is, in a token.  Does not AddRef() or copy strings.
{
	switch (aToken.symbol = symbol) // Assign.
	{
	case SYM_STRING:
		aToken.marker = string;
		aToken.marker_length = string.Length();
		break;
	default:
		aToken.value_int64 = n_int64; // Union copy.
	}
}

void Object::FieldType::Free()
// Only the value is freed, since keys only need to be freed when a field is removed
// entirely or the Object is being deleted.  See Object::Delete.
// CONTAINED VALUE WILL NOT BE VALID AFTER THIS FUNCTION RETURNS.
{
	if (symbol == SYM_STRING)
		string.Free();
	else if (symbol == SYM_OBJECT)
		object->Release();
}


//
// Enumerator
//

ResultType STDMETHODCALLTYPE EnumBase::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (IS_INVOKE_SET)
		return INVOKE_NOT_HANDLED;

	if (IS_INVOKE_CALL)
	{
		if (aParamCount && !_tcsicmp(ParamIndexToString(0), _T("Next")))
		{	// This is something like enum.Next(var); exclude "Next" so it is treated below as enum[var].
			++aParam; --aParamCount;
		}
		else
			return INVOKE_NOT_HANDLED;
	}
	Var *var0 = ParamIndexToOptionalVar(0);
	Var *var1 = ParamIndexToOptionalVar(1);
	_o_return(Next(var0, var1));
}

int Object::Enumerator::Next(Var *aKey, Var *aVal)
{
	if (++mOffset < mObject->mFieldCount)
	{
		FieldType &field = mObject->mFields[mOffset];
		if (aKey)
		{
			if (mOffset < mObject->mKeyOffsetObject) // mKeyOffsetInt < mKeyOffsetObject
				aKey->Assign(field.key.i);
			else if (mOffset < mObject->mKeyOffsetString) // mKeyOffsetObject < mKeyOffsetString
				aKey->Assign(field.key.p);
			else // mKeyOffsetString < mFieldCount
				aKey->Assign(field.key.s);
		}
		if (aVal)
		{
			switch (field.symbol)
			{
			case SYM_STRING:	aVal->AssignString(field.string, field.string.Length());	break;
			case SYM_INTEGER:	aVal->Assign(field.n_int64);		break;
			case SYM_FLOAT:		aVal->Assign(field.n_double);		break;
			case SYM_OBJECT:	aVal->Assign(field.object);			break;
			}
		}
		return true;
	}
	return false;
}

	

//
// Object:: Internal Methods
//

template<typename T>
Object::FieldType *Object::FindField(T val, INT_PTR left, INT_PTR right, INT_PTR &insert_pos)
// Template used below.  left and right must be set by caller to the appropriate bounds within mFields.
{
	INT_PTR mid, result;
	while (left <= right)
	{
		mid = (left + right) / 2;
		
		FieldType &field = mFields[mid];
		
		result = field.CompareKey(val);
		
		if (result < 0)
			right = mid - 1;
		else if (result > 0)
			left = mid + 1;
		else
			return &field;
	}
	insert_pos = left;
	return NULL;
}

Object::FieldType *Object::FindField(SymbolType key_type, KeyType key, IndexType &insert_pos)
// Searches for a field with the given key.  If found, a pointer to the field is returned.  Otherwise
// NULL is returned and insert_pos is set to the index a newly created field should be inserted at.
// key_type and key are output for creating a new field or removing an existing one correctly.
// left and right must indicate the appropriate section of mFields to search, based on key type.
{
	IndexType left, right;

	if (key_type == SYM_STRING)
	{
		left = mKeyOffsetString;
		right = mFieldCount - 1; // String keys are last in the mFields array.

		return FindField<LPTSTR>(key.s, left, right, insert_pos);
	}
	else // key_type == SYM_INTEGER || key_type == SYM_OBJECT
	{
		if (key_type == SYM_INTEGER)
		{
			left = mKeyOffsetInt;
			right = mKeyOffsetObject - 1; // Int keys end where Object keys begin.
		}
		else
		{
			left = mKeyOffsetObject;
			right = mKeyOffsetString - 1; // Object keys end where String keys begin.
		}
		// Both may be treated as integer since left/right exclude keys of an incorrect type:
		return FindField<IntKeyType>(key.i, left, right, insert_pos);
	}
}

void Object::ConvertKey(ExprTokenType &key_token, LPTSTR buf, SymbolType &key_type, KeyType &key)
// Converts key_token to the appropriate key_type and key.
// The exact type of the key is not preserved, since that often produces confusing behaviour;
// for example, guis[WinExist()] := x ... x := guis[A_Gui] would fail because A_Gui returns a
// string.  Strings are converted to integers only where conversion back to string produces
// the same string, so for instance, "01" and " 1 " and "+0x8000" are left as strings.
// Integers which are too large for IntKeyType are converted to strings, to avoid data loss.
{
	SymbolType inner_type = key_token.symbol;
	if (inner_type == SYM_VAR)
	{
		switch (key_token.var->IsPureNumericOrObject())
		{
		case VAR_ATTRIB_IS_INT64:	inner_type = SYM_INTEGER; break;
		case VAR_ATTRIB_IS_OBJECT:	inner_type = SYM_OBJECT; break;
		case VAR_ATTRIB_IS_DOUBLE:	inner_type = SYM_FLOAT; break;
		default:					inner_type = SYM_STRING; break;
		}
	}
	if (inner_type == SYM_OBJECT)
	{
		key_type = SYM_OBJECT;
		key.p = TokenToObject(key_token);
		return;
	}
	if (inner_type == SYM_INTEGER)
	{
		__int64 token_int = TokenToInt64(key_token);
		key.i = IntKeyType(token_int);
		if (__int64(key.i) == token_int) // Confirm round trip is possible.
		{
			key_type = SYM_INTEGER;
			return;
		}
		// Round trip isn't possible, so store it as a string.
	}
	key_type = SYM_STRING; // Set default for simplicity.
	key.s = TokenToString(key_token, buf);
	if (inner_type == SYM_STRING)
	{
		// Only SYM_STRING needs to be checked in this way, because SYM_INTEGER already
		// returned an integer if possible, and SYM_FLOAT never produces a decimal integer.
		LPTSTR cp = key.s;
		if (*cp == '-')
		{
			++cp;
			if (*cp == '0')
				return; // Things like "-0" or "-0x1" or "-007" won't round trip.
		}
		if (*cp == '0')
		{
			// "0" on its own will round trip, but anything else with a leading "0" will not.
			if (!cp[1])
			{
				key.i = 0;
				key_type = SYM_INTEGER;
			}
			return;
		}
		if (*cp <= '9' && *cp >= '1') // Rule out many non-numeric values and leading '+' or whitespace.
		{
			LPTSTR endptr;
			errno = 0; // Clear any previous error number.
			IntKeyType i = ObjParseIntKey(key.s, &endptr);
			if (errno != ERANGE // The number is not too large for IntKeyType.
				&& !*endptr) // There is no trailing whitespace or other non-numeric suffix.
			{
				key.i = i;
				key_type = SYM_INTEGER;
				return; // For maintainability.
			}
		}
	}
}

Object::FieldType *Object::FindField(ExprTokenType &key_token, LPTSTR aBuf, SymbolType &key_type, KeyType &key, IndexType &insert_pos)
// Searches for a field with the given key, where the key is a token passed from script.
{
	ConvertKey(key_token, aBuf, key_type, key);
	return FindField(key_type, key, insert_pos);
}
	
bool Object::SetInternalCapacity(IndexType new_capacity)
// Expands mFields to the specified number if fields.
// Caller *must* ensure new_capacity >= 1 && new_capacity >= mFieldCount.
{
	FieldType *new_fields = (FieldType *)realloc(mFields, new_capacity * sizeof(FieldType));
	if (!new_fields)
		return false;
	mFields = new_fields;
	mFieldCountMax = new_capacity;
	return true;
}
	
Object::FieldType *Object::Insert(SymbolType key_type, KeyType key, IndexType at)
// Inserts a single field with the given key at the given offset.
// Caller must ensure 'at' is the correct offset for this key.
{
	if (mFieldCount == mFieldCountMax && !Expand()  // Attempt to expand if at capacity.
		|| key_type == SYM_STRING && !(key.s = _tcsdup(key.s)))  // Attempt to duplicate key-string.
	{	// Out of memory.
		return NULL;
	}
	// There is now definitely room in mFields for a new field.

	FieldType &field = mFields[at];
	if (at < mFieldCount)
		// Move existing fields to make room.
		memmove(&field + 1, &field, (mFieldCount - at) * sizeof(FieldType));
	++mFieldCount; // Only after memmove above.
	
	// Update key-type offsets based on where and what was inserted; also update this key's ref count:
	if (key_type != SYM_STRING)
	{
		// Must be either SYM_INTEGER or SYM_OBJECT, which both precede SYM_STRING.
		++mKeyOffsetString;

		if (key_type != SYM_OBJECT)
			// Must be SYM_INTEGER, which precedes SYM_OBJECT.
			++mKeyOffsetObject;
		else
			key.p->AddRef();
	}
	
	field.key = key; // Above has already copied string or called key.p->AddRef() as appropriate.
	field.symbol = SYM_STRING;
	field.string.Init(); // Initialize to empty string.  Caller will likely reassign.

	return &field;
}


//
// Property: Invoked when a derived object gets/sets the corresponding key.
//

ResultType STDMETHODCALLTYPE Property::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	Func **member;

	if (aFlags & IF_FUNCOBJ)
	{
		// Use mGet even if IS_INVOKE_CALL, for symmetry:
		//   obj.prop()         ; Get
		//   obj.prop() := val  ; Set (translation is performed by the preparser).
		member = IS_INVOKE_SET ? &mSet : &mGet;
	}
	else
	{
		if (!aParamCount)
			return INVOKE_NOT_HANDLED;

		LPTSTR name = TokenToString(*aParam[0]);
		
		if (!_tcsicmp(name, _T("Get")))
		{
			member = &mGet;
		}
		else if (!_tcsicmp(name, _T("Set")))
		{
			member = &mSet;
		}
		else
			return INVOKE_NOT_HANDLED;
		// ALL CODE PATHS ABOVE MUST RETURN OR SET member.

		if (!IS_INVOKE_CALL)
		{
			if (IS_INVOKE_SET)
			{
				if (aParamCount != 2)
					return OK;
				// Allow changing the GET/SET function, since it's simple and seems harmless.
				*member = TokenToFunc(*aParam[1]); // Can be NULL.
				--aParamCount;
			}
			if (*member && aParamCount == 1)
			{
				aResultToken.symbol = SYM_OBJECT;
				aResultToken.object = *member;
			}
			return OK;
		}
		// Since above did not return, we're explicitly calling Get or Set.
		++aParam;      // Omit method name.
		--aParamCount; // 
	}
	// Since above did not return, member must have been set to either &mGet or &mSet.
	// However, their actual values might be NULL:
	if (!*member)
		return INVOKE_NOT_HANDLED;
	(*member)->Call(aResultToken, aParam, aParamCount);
	return aResultToken.Result();
}


//
// Func: Script interface, accessible via "function reference".
//

ResultType STDMETHODCALLTYPE Func::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	LPTSTR member;

	if (!aParamCount)
		aFlags |= IF_FUNCOBJ;
	else
		member = TokenToString(*aParam[0]);

	if (!IS_INVOKE_CALL && !(aFlags & IF_FUNCOBJ))
	{
		if (IS_INVOKE_SET || aParamCount > 1)
			return INVOKE_NOT_HANDLED;

		     if (!_tcsicmp(member, _T("Name")))			_o_return(mName);
		else if (!_tcsicmp(member, _T("MinParams")))	_o_return(mMinParams);
		else if (!_tcsicmp(member, _T("MaxParams")))	_o_return(mParamCount);
		else if (!_tcsicmp(member, _T("IsBuiltIn")))	_o_return(mIsBuiltIn);
		else if (!_tcsicmp(member, _T("IsVariadic")))	_o_return(mIsVariadic);
		else return INVOKE_NOT_HANDLED;
	}
	
	if (  !(aFlags & IF_FUNCOBJ)  )
	{
		if (!_tcsicmp(member, _T("IsOptional")) && aParamCount <= 2)
		{
			if (aParamCount == 2)
			{
				int param = (int)TokenToInt64(*aParam[1]); // One-based.
				if (param > 0 && (param <= mParamCount || mIsVariadic))
					_o_return(param > mMinParams);
				else
					_o_return_empty;
			}
			else
				_o_return(mMinParams != mParamCount || mIsVariadic); // True if any params are optional.
		}
		else if (!_tcsicmp(member, _T("IsByRef")) && aParamCount <= 2 && !mIsBuiltIn)
		{
			if (aParamCount == 2)
			{
				int param = (int)TokenToInt64(*aParam[1]); // One-based.
				if (param > 0 && (param <= mParamCount || mIsVariadic))
					_o_return(param <= mParamCount && mParam[param-1].is_byref);
				else
					_o_return_empty;
			}
			else
			{
				for (int param = 0; param < mParamCount; ++param)
					if (mParam[param].is_byref)
						_o_return(TRUE);
				_o_return(FALSE);
			}
		}
		else if (!_tcsicmp(member, _T("Bind")))
		{
			if (BoundFunc *bf = BoundFunc::Bind(this, aParam+1, aParamCount-1, IT_CALL | IF_FUNCOBJ))
			{
				aResultToken.symbol = SYM_OBJECT;
				aResultToken.object = bf;
				return OK;
			}
			_o_throw(ERR_OUTOFMEM);
		}
		if (_tcsicmp(member, _T("Call")))
			return INVOKE_NOT_HANDLED; // Reserved.
		// Called explicitly by script, such as by "%obj.funcref%()" or "x := obj.funcref, x.()"
		// rather than implicitly, like "obj.funcref()".
		++aParam;		// Discard the "method name" parameter.
		--aParamCount;	// 
	}
	Call(aResultToken, aParam, aParamCount);
	return aResultToken.Result();
}


ResultType STDMETHODCALLTYPE BoundFunc::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (  !(aFlags & IF_FUNCOBJ) && aParamCount  )
	{
		// No methods/properties implemented yet, except Call().
		if (!TokenIsEmptyString(*aParam[0]) && _tcsicmp(TokenToString(*aParam[0]), _T("Call")))
			return INVOKE_NOT_HANDLED; // Reserved.
		++aParam;
		--aParamCount;
	}

	// Combine the bound parameters with the supplied parameters.
	int bound_count = mParams->MaxIndex();
	if (bound_count > 0)
	{
		ExprTokenType *token = (ExprTokenType *)_alloca(bound_count * sizeof(ExprTokenType));
		ExprTokenType **param = (ExprTokenType **)_alloca((bound_count + aParamCount) * sizeof(ExprTokenType *));
		mParams->ArrayToParams(token, param, bound_count, NULL, 0);
		memcpy(param + bound_count, aParam, aParamCount * sizeof(ExprTokenType *));
		aParam = param;
		aParamCount += bound_count;
	}

	ExprTokenType this_token;
	this_token.symbol = SYM_OBJECT;
	this_token.object = mFunc;

	// Call the function or object.
	return mFunc->Invoke(aResultToken, this_token, mFlags, aParam, aParamCount);
	//return CallFunc(*mFunc, aResultToken, params, param_count);
}

BoundFunc *BoundFunc::Bind(IObject *aFunc, ExprTokenType **aParam, int aParamCount, int aFlags)
{
	if (Object *params = Object::CreateArray(aParam, aParamCount))
	{
		if (BoundFunc *bf = new BoundFunc(aFunc, params, aFlags))
		{
			aFunc->AddRef();
			// bf has taken over our reference to params.
			return bf;
		}
		// malloc failure; release params and return.
		params->Release();
	}
	return NULL;
}

BoundFunc::~BoundFunc()
{
	mFunc->Release();
	mParams->Release();
}


ResultType STDMETHODCALLTYPE Label::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	// Labels are never returned to script, so no need to check flags or parameters.
	return Execute();
}

ResultType LabelPtr::ExecuteInNewThread(TCHAR *aNewThreadDesc, ExprTokenType *aParamValue, int aParamCount, INT_PTR *aRetVal) const
{
	DEBUGGER_STACK_PUSH(aNewThreadDesc)
	ResultType result = CallMethod(mObject, mObject, _T("call"), aParamValue, aParamCount, aRetVal); // Lower-case "call" for compatibility with JScript.
	DEBUGGER_STACK_POP()
	return result;
}


LabelPtr::CallableType LabelPtr::getType(IObject *aObject)
{
	static const Func sFunc(NULL, false);
	// Comparing [[vfptr]] produces smaller code and is perhaps 10% faster than dynamic_cast<>.
	void *vfptr = *(void **)aObject;
	if (vfptr == *(void **)g_script.mPlaceholderLabel)
		return Callable_Label;
	if (vfptr == *(void **)&sFunc)
		return Callable_Func;
	return Callable_Object;
}

Line *LabelPtr::getJumpToLine(IObject *aObject)
{
	switch (getType(aObject))
	{
	case Callable_Label: return ((Label *)aObject)->mJumpToLine;
	case Callable_Func: return ((Func *)aObject)->mJumpToLine;
	default: return NULL;
	}
}

bool LabelPtr::IsExemptFromSuspend() const
{
	if (Line *line = getJumpToLine(mObject))
		return line->IsExemptFromSuspend();
	return false;
}

ActionTypeType LabelPtr::TypeOfFirstLine() const
{
	if (Line *line = getJumpToLine(mObject))
		return line->mActionType;
	return ACT_INVALID;
}

LPTSTR LabelPtr::Name() const
{
	switch (getType(mObject))
	{
	case Callable_Label: return ((Label *)mObject)->mName;
	case Callable_Func: return ((Func *)mObject)->mName;
	default: return _T("Object");
	}
}



ResultType MsgMonitorList::Call(ExprTokenType *aParamValue, int aParamCount, int aInitNewThreadIndex)
{
	ResultType result = OK;
	INT_PTR retval = 0;
	
	for (MsgMonitorInstance inst (*this); inst.index < inst.count; ++inst.index)
	{
		if (inst.index >= aInitNewThreadIndex) // Re-initialize the thread.
			InitNewThread(0, true, false, ACT_INVALID);
		
		IObject *func = mMonitor[inst.index].func;

		if (!CallMethod(func, func, _T("call"), aParamValue, aParamCount, &retval))
		{
			result = FAIL; // Callback encountered an error.
			break;
		}
		if (retval)
		{
			result = CONDITION_TRUE;
			break;
		}
	}
	return result;
}



//
// MetaObject - Defines behaviour of object syntax when used on a non-object value.
//

MetaObject g_MetaObject;

LPTSTR Object::sMetaFuncName[] = { _T("__Get"), _T("__Set"), _T("__Call") };

ResultType STDMETHODCALLTYPE MetaObject::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	// For something like base.Method() in a class-defined method:
	// It seems more useful and intuitive for this special behaviour to take precedence over
	// the default meta-functions, especially since "base" may become a reserved word in future.
	if (aThisToken.symbol == SYM_VAR && !_tcsicmp(aThisToken.var->mName, _T("base"))
		&& !aThisToken.var->HasContents() // Since scripts are able to assign to it, may as well let them use the assigned value.
		&& g->CurrentFunc && g->CurrentFunc->mClass) // We're in a function defined within a class (i.e. a method).
	{
		if (IObject *this_class_base = g->CurrentFunc->mClass->Base())
		{
			ExprTokenType this_token;
			this_token.symbol = SYM_VAR;
			this_token.var = g->CurrentFunc->mParam[0].var;
			ResultType result = this_class_base->Invoke(aResultToken, this_token, (aFlags & ~IF_METAFUNC) | IF_METAOBJ, aParam, aParamCount);
			// Avoid returning INVOKE_NOT_HANDLED in this case so that our caller never
			// shows an "uninitialized var" warning for base.Foo() in a class method.
			if (result != INVOKE_NOT_HANDLED)
				return result;
		}
		_o_return_empty;
	}

	// Allow script-defined meta-functions to override the default behaviour:
	return Object::Invoke(aResultToken, aThisToken, aFlags, aParam, aParamCount);
}


#ifdef CONFIG_DEBUGGER

void ObjectBase::DebugWriteProperty(IDebugProperties *aDebugger, int aPage, int aPageSize, int aMaxDepth)
{
	DebugCookie cookie;
	aDebugger->BeginProperty(NULL, "object", 0, cookie);
	//if (aPage == 0)
	//{
	//	// This is mostly a workaround for debugger clients which make it difficult to
	//	// tell when a property contains an object with no child properties of its own:
	//	aDebugger->WriteProperty("Note", _T("This object doesn't support debugging."));
	//}
	aDebugger->EndProperty(cookie);
}

void Func::DebugWriteProperty(IDebugProperties *aDebugger, int aPage, int aPageSize, int aMaxDepth)
{
	DebugCookie cookie;
	aDebugger->BeginProperty(NULL, "object", 1, cookie);
	if (aPage == 0)
		aDebugger->WriteProperty("Name", mName);
	aDebugger->EndProperty(cookie);
}

#endif
