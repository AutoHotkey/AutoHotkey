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
	, ExprTokenType *aParamValue, int aParamCount, __int64 *aRetVal // For event handlers.
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

	// Exceptions are thrown by Invoke for too few/many parameters, but not for non-existent method.
	// Check for that here, with the exception that objects are permitted to lack a __Delete method.
	if (result == INVOKE_NOT_HANDLED && !(aExtraFlags & IF_METAOBJ))
		result = g_script.ThrowRuntimeException(ERR_UNKNOWN_METHOD, NULL, aMethodName);

	if (result != EARLY_EXIT && result != FAIL)
	{
		// Indicate to caller whether an integer value was returned (for MsgMonitor()).
		result = TokenIsEmptyString(result_token) ? OK : EARLY_RETURN;
	}
	
	if (aRetVal) // Always set this as some callers don't initialize it:
		*aRetVal = result == EARLY_RETURN ? TokenToInt64(result_token) : 0;

	result_token.Free();
	return result;
}


//
// Helpers for Invoke
//

ResultType ObjectMember::Invoke(ObjectMember aMembers[], int aMemberCount, IObject *const aThis
	, ResultToken &aResultToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (IS_INVOKE_SET)
		--aParamCount; // Let aParamCount be just the ones in [].
	if (aParamCount < 1)
		_o_throw(ERR_INVALID_USAGE);

	LPCTSTR name = ParamIndexToString(0, _f_retval_buf);

	for (int i = 0; i < aMemberCount; ++i)
	{
		auto &member = aMembers[i];
		if ((INVOKE_TYPE == IT_CALL) == (member.invokeType == IT_CALL)
			&& !_tcsicmp(member.name, name))
		{
			if (member.invokeType == IT_GET && IS_INVOKE_SET)
				_o_throw(ERR_INVALID_USAGE);
			--aParamCount;
			++aParam;
			if (aParamCount < member.minParams)
				_o_throw(ERR_TOO_FEW_PARAMS);
			if (aParamCount > member.maxParams && member.maxParams != MAXP_VARIADIC)
				_o_throw(ERR_TOO_MANY_PARAMS);
			for (int i = 0; i < member.minParams; ++i)
				if (aParam[i]->symbol == SYM_MISSING)
					_o_throw(ERR_PARAM_REQUIRED);
			return (aThis->*member.method)(aResultToken, member.id, aFlags, aParam, aParamCount);
		}
	}
	return INVOKE_NOT_HANDLED;
}


//
// Object::Create - Create a new Object given an array of property name/value pairs.
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
		IndexType insert_pos;
		
		for (int i = 0; i + 1 < aParamCount; i += 2)
		{
			if (aParam[i]->symbol == SYM_MISSING || aParam[i+1]->symbol == SYM_MISSING)
				continue; // For simplicity.

			auto name = TokenToString(*aParam[i], buf);

			if (!_tcsicmp(name, _T("base")))
			{
				// For consistency with assignments, the following is allowed to overwrite a previous
				// base object (although having "base" occur twice in the parameter list would be quite
				// useless) or set mBase to NULL if the value parameter is not an object.
				obj->SetBase(TokenToObject(*aParam[i + 1]));
				continue;
			}

			auto field = obj->FindField(name, insert_pos);
			if (  !(field
				 || (field = obj->Insert(name, insert_pos)))
				|| !field->Assign(*aParam[i + 1])  )
			{	// Out of memory.
				obj->Release();
				return NULL;
			}
		}
	}
	return obj;
}


//
// Map::Create - Create a new Map given an array of key/value pairs.
//

Map *Map::Create(ExprTokenType *aParam[], int aParamCount)
{
	ASSERT(!(aParamCount & 1));

	Map *obj = new Map();
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


//
// Object::Clone - Used for variadic function-calls.
//

// Creates an object and copies the content of this object to it.
Object *Object::Clone()
{
	auto objptr = new Object();
	return CloneTo(*objptr);
}

// Helper function for temporary implementation of Clone by Object and subclasses.
// Should be eliminated once revision of the object model is complete.
Object *Object::CloneTo(Object &obj)
{
	// Allocate space in destination object.
	IndexType field_count = mFieldCount;
	if (!obj.SetInternalCapacity(field_count))
	{
		obj.Release();
		return NULL;
	}

	int failure_count = 0; // See comment below.
	IndexType i;

	obj.mFieldCount = field_count;

	for (i = 0; i < field_count; ++i)
	{
		FieldType &dst = obj.mFields[i];
		FieldType &src = mFields[i];

		// Copy name.
		dst.key_c = src.key_c;
		if ( !(dst.name = _tcsdup(src.name)) )
		{
			// Rather than trying to set up the object so that what we have
			// so far is valid in order to break out of the loop, continue,
			// make all fields valid and then allow them to be freed. 
			++failure_count;
		}

		// Copy value.
		if (!dst.InitCopy(src))
			++failure_count;
	}
	if (failure_count)
	{
		// One or more memory allocations failed.  It seems best to return a clear failure
		// indication rather than an incomplete copy.  Now that the loop above has finished,
		// the object's contents are at least valid and it is safe to free the object:
		obj.Release();
		return NULL;
	}
	if (mBase)
		(obj.mBase = mBase)->AddRef();
	return &obj;
}

Map *Map::CloneTo(Map &obj)
{
	Object::CloneTo(obj);

	if (!obj.SetInternalCapacity(mFieldCount))
	{
		obj.Release();
		return NULL;
	}

	int failure_count = 0; // See Object::CloneT() for comments.
	IndexType i;

	obj.mFieldCount = mFieldCount;
	obj.mKeyOffsetObject = mKeyOffsetObject;
	obj.mKeyOffsetString = mKeyOffsetString;
	if (obj.mKeyOffsetObject < 0) // Currently might always evaluate to false.
	{
		obj.mKeyOffsetObject = 0; // aStartOffset excluded all integer and some or all object keys.
		if (obj.mKeyOffsetString < 0)
			obj.mKeyOffsetString = 0; // aStartOffset also excluded some string keys.
	}
	//else no need to check mKeyOffsetString since it should always be >= mKeyOffsetObject.

	for (i = 0; i < mFieldCount; ++i)
	{
		FieldType &dst = obj.mFields[i];
		FieldType &src = mFields[i];

		// Copy key.
		if (i >= obj.mKeyOffsetString)
		{
			dst.key_c = src.key_c;
			if ( !(dst.key.s = _tcsdup(src.key.s)) )
			{
				// Key allocation failed. At this point, all int and object keys
				// have been set and values for previous fields have been copied.
				++failure_count;
			}
		}
		else 
		{
			// Copy whole key; search "(IntKeyType)(INT_PTR)" for comments.
			dst.key = src.key;
			if (i >= obj.mKeyOffsetObject)
				dst.key.p->AddRef();
		}

		// Copy value.
		if (!dst.InitCopy(src))
			++failure_count;
	}
	if (failure_count)
	{
		obj.Release();
		return NULL;
	}
	return &obj;
}


//
// Array::ToParams - Used for variadic function-calls.
//

// Copies this array's elements into the parameter list.
// Caller has ensured param_list can fit aParamCount + Length().
void Array::ToParams(ExprTokenType *token, ExprTokenType **param_list, ExprTokenType **aParam, int aParamCount)
{
	for (index_t i = 0; i < mLength; ++i)
		mItem[i].ToToken(token[i]);
	
	ExprTokenType **param_ptr = param_list;
	for (int i = 0; i < aParamCount; ++i)
		*param_ptr++ = aParam[i]; // Caller-supplied param token.
	for (index_t i = 0; i < mLength; ++i)
		*param_ptr++ = &token[i]; // New param.
}

// Calls an Enumerator repeatedly and returns an Array of all first-arg values.
// This is used in conjunction with Array::ToParams to support other objects.
Array *Array::FromEnumerable(IObject *aEnumerable)
{
	FuncResult result_token;
	ExprTokenType t_this(aEnumerable);
	ExprTokenType param[2], *params[2] = { param, param + 1 };

	param[0].SetValue(_T("_NewEnum"), 8);
	auto result = aEnumerable->Invoke(result_token, t_this, IT_CALL, params, 1);
	if (result == FAIL || result == EARLY_EXIT || result == INVOKE_NOT_HANDLED)
	{
		if (result == INVOKE_NOT_HANDLED)
			g_script.ScriptError(ERR_UNKNOWN_METHOD, _T("_NewEnum"));
		return nullptr;
	}
	IObject *enumerator = TokenToObject(result_token);
	if (!enumerator)
	{
		g_script.ScriptError(ERR_TYPE_MISMATCH, _T("_NewEnum"));
		result_token.Free();
		return nullptr;
	}
	enumerator->AddRef();

	Var var;
	t_this.SetValue(enumerator);
	param[0].SetValue(_T("Next"), 4);
	param[1].symbol = SYM_VAR;
	param[1].var = &var;
	Array *vargs = new Array();
	for (;;)
	{
		result_token.Free();
		result_token.InitResult(result_token.buf);
		auto result = enumerator->Invoke(result_token, t_this, IT_CALL, params, 2);
		if (result == FAIL || result == EARLY_EXIT || result == INVOKE_NOT_HANDLED)
		{
			if (result == INVOKE_NOT_HANDLED)
				g_script.ScriptError(ERR_UNKNOWN_METHOD, _T("Next"));
			vargs->Release();
			vargs = nullptr;
			break;
		}
		if (!TokenToBOOL(result_token))
			break;
		ExprTokenType value;
		var.ToTokenSkipAddRef(value);
		vargs->Append(value);
	}
	result_token.Free();
	enumerator->Release();
	return vargs;
}


//
// Array::ToStrings - Used by BIF_StrSplit.
//

ResultType Array::ToStrings(LPTSTR *aStrings, int &aStringCount, int aStringsMax)
{
	for (index_t i = 0; i < mLength; ++i)
		if (SYM_STRING == mItem[i].symbol)
			aStrings[i] = mItem[i].string;
		else
			return FAIL;
	aStringCount = mLength;
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
		if (FindField(_T("__Class")))
			// This object appears to be a class definition, so it would probably be
			// undesirable to call the super-class' __Delete() meta-function for this.
			return ObjectBase::Delete();

		// L33: Privatize the last recursion layer's deref buffer in case it is in use by our caller.
		// It's done here rather than in Var::FreeAndRestoreFunctionVars (even though the below might
		// not actually call any script functions) because this function is probably executed much
		// less often in most cases.
		PRIVATIZE_S_DEREF_BUF;

		Line *curr_line = g_script.mCurrLine;

		// If an exception has been thrown, temporarily clear it for execution of __Delete.
		ResultToken *exc = g->ThrownToken;
		g->ThrownToken = NULL;
		
		// This prevents an erroneous "The current thread will exit" message when an error occurs,
		// by causing LineError() to throw an exception:
		int outer_excptmode = g->ExcptMode;
		g->ExcptMode |= EXCPTMODE_DELETE;

		CallMethod(mBase, this, _T("__Delete"), NULL, 0, NULL, IF_METAOBJ); // base.__Delete()

		g->ExcptMode = outer_excptmode;

		// Exceptions thrown by __Delete are reported immediately because they would not be handled
		// consistently by the caller (they would typically be "thrown" by the next function call),
		// and because the caller must be allowed to make additional __Delete calls.
		if (g->ThrownToken)
			g_script.FreeExceptionToken(g->ThrownToken);

		// If an exception has been thrown by our caller, it's likely that it can and should be handled
		// reliably by our caller, so restore it.
		if (exc)
			g->ThrownToken = exc;

		g_script.mCurrLine = curr_line; // Prevent misleading error reports/Exception() stack trace.

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
			for (IndexType i = 0; i < mFieldCount; ++i)
			{
				free(mFields[i].name);
				mFields[i].Free();
			}
		}
		// Free fields array.
		free(mFields);
	}
}


Map::~Map()
{
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
// Invoke - dynamic dispatch
//

ObjectMember Object::sMembers[] =
{
	Object_Method1(Delete, 1, 2),
	Object_Method1(Count, 0, 0),
	Object_Method1(SetCapacity, 1, 1),
	Object_Method1(GetCapacity, 0, 0),
	Object_Method1(_NewEnum, 0, 0),
	Object_Method1(HasKey, 1, 1),
	Object_Method1(Clone, 0, 0)
};

ResultType STDMETHODCALLTYPE Object::Invoke(
                                            ResultToken &aResultToken,
                                            ExprTokenType &aThisToken,
                                            int aFlags,
                                            ExprTokenType *aParam[],
                                            int aParamCount
                                            )
{
	name_t name;
    FieldType *field, *prop_field;
	IndexType insert_pos;
	Property *prop = NULL; // Set default.

	// If this is some object's base and is being invoked in that capacity, call
	//	__Get/__Set/__Call as defined in this base object before searching further.
	// Meta-functions are skipped for .__Item for these reasons:
	//  1) They would be called for every access with [k] when __Item is not defined, 
	//     instead of just those where [k] is not defined.  For the short term, this
	//     breaks expectations (and scripts).  For the long term it seems redundant.
	//  2) __Item handling is currently designed to be backward-compatible to ease
	//     the transition.  If not present, it will cause recursion, which may then
	//     call the meta-functions in the pre-established manner.
	//  3) Even once the fallback behaviour is removed, it is probably more useful
	//     for __Item and __Get/__Set to be mutually exclusive, and more intuitive
	//     for __Item to be handled like a meta-function (don't call other meta-s
	//     in the process of calling __Item).
	if ((aFlags & IF_METAFUNC) && (aFlags & (IF_DEFAULT|IT_CALL)) != IF_DEFAULT)
	{
		// Look for a meta-function definition directly in this base object.
		if (field = FindField(sMetaFuncName[INVOKE_TYPE]))
		{
			Line *curr_line = g_script.mCurrLine;
			ResultType r = CallField(field, aResultToken, aThisToken, aFlags, aParam, aParamCount);
			g_script.mCurrLine = curr_line; // Allows exceptions thrown by later meta-functions to report a more appropriate line.
			//if (r == EARLY_RETURN)
				// Propagate EARLY_RETURN in case this was the __Call meta-function of a
				// "function object" which is used as a meta-function of some other object.
				//return EARLY_RETURN; // TODO: Detection of 'return' vs 'return empty_value'.
			if (r != OK) // Likely EARLY_RETURN, FAIL or EARLY_EXIT.
				return r;
		}
	}
	
	auto actual_param = aParam; // Actual first parameter between [] or ().
	int actual_param_count = aParamCount; // Actual number of parameters between [] or ().

	if (IS_INVOKE_SET)
	{
		// Due to the way expression parsing works, the result should never be negative
		// (and any direct callers of Invoke must always pass aParamCount >= 1):
		--actual_param_count;
	}
	
	ASSERT(actual_param_count || (aFlags & IF_DEFAULT));
	{
		if ((aFlags & IF_DEFAULT) || aParam[0]->symbol == SYM_MISSING)
			name = IS_INVOKE_CALL ? _T("Call") : _T("__Item");
		else
			name = ParamIndexToString(0, _f_number_buf);

		if (!(aFlags & IF_DEFAULT))
		{
			++actual_param;
			--actual_param_count;
		}
		
		field = FindField(name, insert_pos);

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
			if (IS_INVOKE_SET ? prop->CanSet() : IS_INVOKE_GET && prop->CanGet())
			{
				// Prepare the parameter list: this, value, actual_param*
				ExprTokenType **prop_param = (ExprTokenType **)_alloca((actual_param_count + 2) * sizeof(ExprTokenType *));
				prop_param[0] = &aThisToken; // For the hidden "this" parameter in the getter/setter.
				int prop_param_count = 1;
				if (IS_INVOKE_SET)
				{
					// Put the setter's hidden "value" parameter before the other parameters.
					prop_param[prop_param_count++] = actual_param[actual_param_count];
				}
				memcpy(prop_param + prop_param_count, actual_param, actual_param_count * sizeof(ExprTokenType *));
				prop_param_count += actual_param_count;
				
				// Pass IF_DEFAULT so that it'll pass all parameters to the getter/setter.
				// For a functor Object, we would need to pass a token representing "this" Property,
				// but since Property::Invoke doesn't use it, we pass our aThisToken for simplicity.
				ResultType result = prop->Invoke(aResultToken, aThisToken, aFlags | IF_DEFAULT, prop_param, prop_param_count);
				return result == EARLY_RETURN ? OK : result;
			}
			// The property was missing get/set (whichever this invocation is), so continue as
			// if the property itself wasn't defined.
			field = NULL;
		}
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
				|| prop) // Nothing left to do in this case.
				return r;

			// Since the above may have inserted or removed fields (including the specified one),
			// insert_pos may no longer be correct or safe.  Updating field also allows a meta-function
			// to initialize a field and allow processing to continue as if it already existed.
			field = FindField(name, /*out*/ insert_pos);
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
		if (!IS_INVOKE_META && !field) // v1.1.16: Check field again so if __Call sets a field, it gets called.
		{
			//
			// BUILT-IN METHODS
			//
			if (IS_INVOKE_CALL)
			{
				// Since above has not handled this call and no field exists, check for built-in methods.
				return ObjectMember::Invoke(sMembers, _countof(sMembers), this, aResultToken, aFlags, aParam, aParamCount);
			}
			//
			// BUILT-IN "BASE" PROPERTY
			//
			else if (actual_param_count == 0)
			{
				if (!_tcsicmp(name, _T("base")))
				{
					if (IS_INVOKE_SET)
					{
						IObject *obj = TokenToObject(**actual_param);
						if (obj)
							obj->AddRef(); // for mBase
						if (mBase)
							mBase->Release();
						mBase = obj; // May be NULL.
						return OK;
					}
	
					if (mBase)
					{
						mBase->AddRef();
						_o_return(mBase);
					}
					//else return empty string.
					_o_return_empty;
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
			return CallField(field, aResultToken, aThisToken, aFlags & ~IF_DEFAULT, actual_param, actual_param_count);
		aFlags = (aFlags & ~(IT_BITMASK | IF_CALL_FUNC_ONLY)) | IT_GET;
	}

	// This next section handles both this[x,y] (handled as this.__Item[x,y]) and this.x[y]
	// Here, we only retrieve or create the sub-object this[x]/this.x, never set the actual
	// property (since that's handled via recursion into the sub-object).
	if (actual_param_count > 0)
	{
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
			if (!_tcsicmp(name, _T("base")))
			{
				if (!mBase && IS_INVOKE_SET)
					mBase = new Object();
				obj = mBase; // If NULL, above failed and below will detect it.
			}
			else if (aFlags & IF_DEFAULT)
			{
				// There's no this.__Item property or key-value pair and it was not handled by __get/__set.
				// Do not create this.__Item; instead, fall back to the old behaviour (use the next
				// parameter as a key).  Once there are dedicated map/dictionary/array types, this
				// and the next section should be removed (__Item should be explicitly defined).
				obj = this;
				// IF_DEFAULT will be removed below to prevent infinite recursion.
			}
			// Automatically create a new object for the x part of obj.x[y]:=z.
			else if (IS_INVOKE_SET)
			{
				Object *new_obj = new Object();
				if (new_obj)
				{
					if ( field = prop ? prop_field : Insert(name, insert_pos) )
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
				// For now, x[y] (that is, x.__item[y]) is permitted since x.__item will be initialized
				// automatically if the script assigns to x[y].
				//if (aFlags & IF_DEFAULT)

				// Treat x.y[z] like x.y when x.y is not set: just return "", don't throw an exception.
				// On the other hand, if x.y is set to something which is not an object, the "if (field)"
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
			if (IS_INVOKE_SET)
				++actual_param_count;
			// Recursively invoke obj, passing remaining parameters; remove IF_META to correctly treat obj as target:
			// Toggle IF_DEFAULT so that this[1,2] expands to this.__Item.1.__Item.2, same as this[1][2].
			// This should be changed once a dedicated map/dictionary type is implemented.
			return obj->Invoke(aResultToken, obj_token, (aFlags & ~IF_META) ^ IF_DEFAULT, actual_param, actual_param_count);
			// Above may return INVOKE_NOT_HANDLED in cases such as obj[a,b] where obj[a] exists but obj[a][b] does not.
		}
	} // MULTIPARAM[x,y]

	// SET
	else if (IS_INVOKE_SET)
	{
		if (!IS_INVOKE_META)
		{
			ExprTokenType &value_param = **actual_param;
			if ( (field || (field = prop ? prop_field : Insert(name, insert_pos)))
				&& field->Assign(value_param) )
				return OK;
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
			//   1) Fn {value:"string"}.value   ; Temporary object could be released prematurely.
			//   2) Fn( obj.value, obj := "" )  ; Object is freed by the assignment.
			// For both cases, the value is copied immediately after we return, because the result of any
			// BIF is assumed to be volatile if expression eval isn't finished.  The function call in #1
			// is handled by ExpandExpression() since commit 2a276145.
			field->ReturnRef(aResultToken);
			return OK;
		}
		// If 'this' is the target object (not its base), produce OK so that something like if(!foo.bar) is
		// considered valid even when foo.bar has not been set.
		if (!IS_INVOKE_META)
			_o_return_empty;
	}

	// Fell through from one of the sections above: invocation was not handled.
	return INVOKE_NOT_HANDLED;
}


ResultType Object::CallBuiltin(int aID, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case FID_ObjDelete:			return Delete(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjCount:			return Count(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjHasKey:			return HasKey(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjGetCapacity:	return GetCapacity(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjSetCapacity:	return SetCapacity(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjClone:			return Clone(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjNewEnum:		return _NewEnum(aResultToken, 0, IT_CALL, aParam, aParamCount);
	}
	return INVOKE_NOT_HANDLED;
}


ObjectMember Map::sMembers[] =
{
	Object_Member(Capacity, Capacity, 0, IT_SET),
	Object_Member(Count, Count, 0, IT_GET),
	Object_Method1(Clone, 0, 0),
	Object_Method1(Delete, 1, 2),
	Object_Method1(Has, 1, 1),
	Object_Method1(_NewEnum, 0, 0)
};

ResultType STDMETHODCALLTYPE Map::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if ((aFlags & (IF_DEFAULT|IT_CALL)) == IF_DEFAULT)
	{
		if (aParamCount != (IS_INVOKE_SET ? 2 : 1))
			_o_throw(ERR_INVALID_USAGE);
		if (IS_INVOKE_GET)
		{
			if (GetItem(aResultToken, *aParam[0]))
			{
				if (aResultToken.symbol == SYM_OBJECT)
					aResultToken.object->AddRef();
				return OK;
			}
			_o_throw(_T("Key not found."), ParamIndexToString(0, _f_number_buf));
		}
		else
		{
			if (!SetItem(*aParam[0], *aParam[1]))
				_o_throw(ERR_OUTOFMEM);
		}
		return OK;
	}
	auto result = ObjectMember::Invoke(sMembers, _countof(sMembers), this, aResultToken, aFlags, aParam, aParamCount);
	// See Array::Invoke for comments regarding this part:
	if (result == INVOKE_NOT_HANDLED)
		result = Object::Invoke(aResultToken, aThisToken, aFlags, aParam, aParamCount);
	return result;
}


//
// Internal: Object::CallField - Used by Object::Invoke to call a function/method stored in this object.
//

ResultType Object::CallField(FieldType *aField, ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
// aParam[0] contains the identifier of this field or an empty space (for __Get etc.).
{
	// Allocate a new array of param pointers that we can modify.
	ExprTokenType **param = (ExprTokenType **)_alloca((aParamCount + 1) * sizeof(ExprTokenType *));
	// Where fn = this.%key%, call %fn%(this, params*).
	ExprTokenType field_token(aField->object); // fn
	param[0] = &aThisToken; // this
	memcpy(param + 1, aParam, aParamCount * sizeof(ExprTokenType*)); // params*
	++aParamCount;

	if (aField->symbol == SYM_OBJECT)
	{
		return aField->object->Invoke(aResultToken, field_token, IT_CALL|IF_DEFAULT, param, aParamCount);
	}
	if (aField->symbol == SYM_STRING)
	{
		Func *func = g_script.FindFunc(aField->string, aField->string.Length());
		if (func)
		{
			// v2: Always pass "this" as the first parameter.  The old behaviour of passing it only when called
			// indirectly via mBase was confusing to many users, and isn't needed now that the script can do
			// %this.func%() instead of this.func() if they don't want to pass "this".
			func->Call(aResultToken, param, aParamCount);
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

Array *Array::FromArgV(LPTSTR *aArgV, int aArgC)
{
	ExprTokenType *token = (ExprTokenType *)_alloca(aArgC * sizeof(ExprTokenType));
	ExprTokenType **param = (ExprTokenType **)_alloca(aArgC * sizeof(ExprTokenType*));
	for (int j = 0; j < aArgC; ++j)
	{
		token[j].SetValue(aArgV[j]);
		param[j] = &token[j];
	}
	return Create(param, aArgC);
}



//
// Helper function for StrSplit/WinGetList/WinGetControls
//

bool Array::Append(ExprTokenType &aValue)
{
	if (mLength == MaxIndex || !EnsureCapacity(mLength + 1))
		return false;
	auto &item = mItem[mLength++];
	item.Minit();
	return item.Assign(aValue);
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
			free(mFields[i].name);
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
// Object::Type() - Returns the object's type/class name.
//

static LPTSTR sObjectTypeName = _T("Object");

LPTSTR Object::Type()
{
	IObject *ibase;
	Object *base;
	ExprTokenType value;
	if (GetItem(value, _T("__Class")))
		return _T("Class"); // This object is a class.
	for (ibase = mBase; base = dynamic_cast<Object *>(ibase); ibase = base->mBase)
		if (base->GetItem(value, _T("__Class")))
			return TokenToString(value); // This object is an instance of base.
	return sObjectTypeName; // This is an Object of undetermined type, like Object() or {}.
}

LPTSTR Array::Type()
{
	// Override default type name, but let script redefine it as per Object rules.
	auto type = Object::Type();
	if (type == sObjectTypeName)
		return _T("Array");
	return type;
}

LPTSTR Map::Type()
{
	// Override default type name, but let script redefine it as per Object rules.
	auto type = Object::Type();
	if (type == sObjectTypeName)
		return _T("Map");
	return type;
}

	

//
// Object:: and Map:: Built-in Methods
//

ResultType Object::Delete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto field = FindField(ParamIndexToString(0, _f_number_buf));
	if (!field)
		_o_return_empty;
	
	// Return the removed value.
	field->ReturnMove(aResultToken);

	// Free the value (if not transferred to aResultToken) and name.
	field->Free();
	free(field->name);

	// Remove the field.
	memmove(field, field + 1, (mFieldCount - (field - mFields)) * sizeof(FieldType));
	mFieldCount--;
	return OK;
}

ResultType Map::Delete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Delete(first_key [, last_key := first_key])
{
	FieldType *min_field;
	IndexType min_pos, max_pos, pos;
	SymbolType min_key_type;
	KeyType min_key, max_key;
	IntKeyType logical_count_removed = 1;

	LPTSTR number_buf = _f_number_buf;

	// Find the position of "min".
	if (min_field = FindField(*aParam[0], number_buf, min_key_type, min_key, min_pos))
		min_pos = min_field - mFields; // else min_pos was already set by FindField.

	if (aParamCount > 1) // Removing a range of keys.
	{
		SymbolType max_key_type;
		FieldType *max_field;

		// Find the next position > [aParam[1]].
		if (max_field = FindField(*aParam[1], number_buf, max_key_type, max_key, max_pos))
			max_pos = max_field - mFields + 1;

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
			// Our return value when only one key is given is supposed to be the value
			// previously at this[key], which has just been removed.  Since this[key]
			// would return "", it makes sense to return the same in this case.
			_o_return_empty;
		}
		// Since only one field (at maximum) can be removed in this mode, it
		// seems more useful to return the field being removed than a count.
		min_field->ReturnMove(aResultToken);
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


ResultType Object::Count(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return((__int64)mFieldCount);
}

ResultType Map::Count(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return((__int64)mFieldCount);
}

ResultType Object::GetCapacity(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return(mFieldCountMax);
}

ResultType Object::SetCapacity(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (!ParamIndexIsNumeric(0))
		_o_throw(ERR_PARAM1_INVALID);

	IndexType desired_count = (IndexType)ParamIndexToInt64(0);
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

ResultType Map::Capacity(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (IS_INVOKE_GET)
	{
		_o_return(mFieldCountMax);
	}

	if (!ParamIndexIsNumeric(0))
		_o_throw(ERR_PARAM1_INVALID);

	IndexType desired_count = (IndexType)ParamIndexToInt64(0);
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

ResultType Object::_NewEnum(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (IObject *enm = new Enumerator(this))
		_o_return(enm);
	else
		_o_throw(ERR_OUTOFMEM);
}

ResultType Map::_NewEnum(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (IObject *enm = new Enumerator(this))
		_o_return(enm);
	else
		_o_throw(ERR_OUTOFMEM);
}

ResultType Object::HasKey(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return(FindField(ParamIndexToString(0, _f_number_buf)) != NULL);
}

ResultType Map::Has(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	SymbolType key_type;
	KeyType key;
	INT_PTR insert_pos;
	FieldType *field = FindField(*aParam[0], _f_number_buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos);
	_o_return(field != NULL);
}

ResultType Object::Clone(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	Object *clone = Clone();
	if (!clone)
		_o_throw(ERR_OUTOFMEM);	
	_o_return(clone);
}

ResultType Map::Clone(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto clone = new Map();
	if (!CloneTo(*clone))
		_o_throw(ERR_OUTOFMEM);
	_o_return(clone);
}


//
// Object::Variant
//

void Object::Variant::Minit()
{
	symbol = SYM_MISSING;
	string.Init();
}

void Object::Variant::AssignEmptyString()
{
	Free();
	symbol = SYM_STRING;
	string.Init();
}

void Object::Variant::AssignMissing()
{
	Free();
	Minit();
}

bool Object::Variant::Assign(LPTSTR str, size_t len, bool exact_size)
{
	if (len == -1)
		len = _tcslen(str);

	if (!len) // Check len, not *str, since it might be binary data or not null-terminated.
	{
		AssignEmptyString();
		return true;
	}

	if (symbol != SYM_STRING || len >= string.Capacity())
	{
		AssignEmptyString(); // Free object or previous buffer (which was too small).

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

bool Object::Variant::Assign(ExprTokenType &aParam)
{
	ExprTokenType temp, *val; // Seems more maintainable to use a copy; avoid any possible side-effects.
	if (aParam.symbol == SYM_VAR)
	{
		aParam.var->ToTokenSkipAddRef(temp); // Skip AddRef() if applicable because it's called below.
		val = &temp;
	}
	else
		val = &aParam;

	switch (val->symbol)
	{
	case SYM_STRING:
		return Assign(val->marker, val->marker_length);
	case SYM_MISSING:
		AssignMissing();
		return OK;
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

// Copy a value from a Variant into this uninitialized Variant.
bool Object::Variant::InitCopy(Variant &val)
{
	switch (symbol = val.symbol)
	{
	case SYM_STRING:
		string.Init();
		return Assign(val.string, val.string.Length(), true); // Pass true to conserve memory (no space is allowed for future expansion).
	case SYM_OBJECT:
		(object = val.object)->AddRef();
		break;
	//case SYM_INTEGER:
	//case SYM_FLOAT:
	default:
		n_int64 = val.n_int64; // Union copy.
	}
	return true;
}

// Return value, knowing Variant will be kept around.
// Copying of value can be skipped.
void Object::Variant::ReturnRef(ResultToken &result)
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
	case SYM_MISSING:
		result.SetValue(_T(""), 0);
		break;
	//case SYM_INTEGER:
	//case SYM_FLOAT:
	default:
		result.value_int64 = n_int64; // Union copy.
	}
}

// Return value, knowing Variant will shortly be deleted.
// Value may be moved from Variant into ResultToken.
void Object::Variant::ReturnMove(ResultToken &result)
{
	switch (result.symbol = symbol)
	{
	case SYM_STRING:
		// For simplicity, just discard the memory of item.string (can't return it as-is since
		// the string isn't at the start of its memory block).  Scripts can use the 2-param mode
		// to avoid any performance penalty this may incur.
		TokenSetResult(result, string, string.Length());
		break;
	case SYM_OBJECT:
		result.object = object;
		Minit(); // Let item forget the object ref since we are taking ownership.
		break;
	case SYM_MISSING:
		result.SetValue(_T(""), 0);
		break;
	//case SYM_INTEGER:
	//case SYM_FLOAT:
	default:
		result.value_int64 = n_int64; // Effectively also value_double = n_double.
	}
}

// Used when we want the value as is, in a token.  Does not AddRef() or copy strings.
void Object::Variant::ToToken(ExprTokenType &aToken)
{
	switch (aToken.symbol = symbol) // Assign.
	{
	case SYM_STRING:
	case SYM_MISSING:
		aToken.marker = string;
		aToken.marker_length = string.Length();
		break;
	default:
		aToken.value_int64 = n_int64; // Union copy.
	}
}

void Object::Variant::Free()
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
// Array
//

ResultType Array::SetCapacity(index_t aNewCapacity)
{
	if (mLength > aNewCapacity)
		RemoveAt(aNewCapacity, mLength - aNewCapacity);
	auto new_item = (Variant *)realloc(mItem, sizeof(Variant) * aNewCapacity);
	if (!new_item)
		return FAIL;
	mItem = new_item;
	mCapacity = aNewCapacity;
	return OK;
}

ResultType Array::EnsureCapacity(index_t aRequired)
{
	if (mCapacity >= aRequired)
		return OK;
	// Simple doubling of previous capacity, if that's enough, seems adequate.
	// Otherwise, allocate exactly the amount required with no room to spare.
	// v1 Object doubled in capacity when needed to add a new field, but started
	// at 4 and did not allocate any extra space when inserting with InsertAt or
	// Push.  By contrast, this approach:
	//  1) Wastes no space in the possibly common case where Array::InsertAt is
	//     called exactly once (such as when constructing the Array).
	//  2) Expands exponentially if Push is being used repeatedly, which should
	//     perform much better than expanding by 1 each time.
	if (aRequired < (mCapacity << 1))
		aRequired = (mCapacity << 1);
	return SetCapacity(aRequired);
}

template<typename TokenT>
ResultType Array::InsertAt(index_t aIndex, TokenT aValue[], index_t aCount)
{
	ASSERT(aIndex <= mLength);

	if (!EnsureCapacity(mLength + aCount))
		return FAIL;

	if (aIndex < mLength)
	{
		memmove(mItem + aIndex + aCount, mItem + aIndex, (mLength - aIndex) * sizeof(mItem[0]));
	}
	for (index_t i = 0; i < aCount; ++i)
	{
		mItem[aIndex + i].Minit();
		mItem[aIndex + i].Assign(aValue[i]);
	}
	mLength += aCount;
	return OK;
}

template ResultType Array::InsertAt(index_t, ExprTokenType *[], index_t);
template ResultType Array::InsertAt(index_t, ExprTokenType [], index_t);

void Array::RemoveAt(index_t aIndex, index_t aCount)
{
	ASSERT(aIndex + aCount <= mLength);

	for (index_t i = 0; i < aCount; ++i)
	{
		mItem[aIndex + i].Free();
	}
	if (aIndex < mLength)
	{
		memmove(mItem + aIndex, mItem + aIndex + aCount, (mLength - aIndex - aCount) * sizeof(mItem[0]));
	}
	mLength -= aCount;
}

ResultType Array::SetLength(index_t aNewLength)
{
	if (mLength > aNewLength)
	{
		RemoveAt(aNewLength, mLength - aNewLength);
		return OK;
	}
	if (aNewLength > mCapacity && !SetCapacity(aNewLength))
		return FAIL;
	for (index_t i = mLength; i < aNewLength; ++i)
	{
		mItem[i].Minit();
	}
	mLength = aNewLength;
	return OK;
}

Array::~Array()
{
	RemoveAt(0, mLength);
	free(mItem);
}

Array *Array::Create(ExprTokenType *aValue[], index_t aCount)
{
	auto arr = new Array();
	if (!aCount || arr->InsertAt(0, aValue, aCount))
		return arr;
	arr->Release();
	return nullptr;
}

Array *Array::Clone(BOOL aMembersOnly)
{
	auto arr = new Array();
	if (!CloneTo(*arr))
		return nullptr; // CloneTo() released arr.
	if (aMembersOnly) // Flag used in variadic calls (Func::Call).
		return arr;
	if (!arr->SetCapacity(mCapacity))
		return nullptr;
	for (index_t i = 0; i < mLength; ++i)
	{
		auto &new_item = arr->mItem[arr->mLength++];
		new_item.Minit();
		ExprTokenType value;
		mItem[i].ToToken(value);
		if (!new_item.Assign(value))
		{
			arr->Release();
			return nullptr;
		}
	}
	return arr;
}

bool Array::ItemToToken(index_t aIndex, ExprTokenType &aToken)
{
	if (aIndex >= mLength)
		return false;
	mItem[aIndex].ToToken(aToken);
	return true;
}

ObjectMember Array::sMembers[] =
{
	Object_Property_get_set(Length),
	Object_Property_get_set(Capacity),
	Object_Method(InsertAt, 2, MAXP_VARIADIC),
	Object_Method(Push, 1, MAXP_VARIADIC),
	Object_Method(RemoveAt, 1, 2),
	Object_Method(Pop, 0, 0),
	Object_Method(Has, 1, 1),
	Object_Method(Delete, 1, 1),
	Object_Method(Clone, 0, 0),
	Object_Method(_NewEnum, 0, 0)
};

ResultType STDMETHODCALLTYPE Array::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if ((aFlags & (IF_DEFAULT|IT_CALL)) == IF_DEFAULT)
	{
		if (aParamCount != (IS_INVOKE_SET ? 2 : 1))
			_o_throw(ERR_INVALID_USAGE);
		auto index = ParamToZeroIndex(*aParam[0]);
		if (index >= mLength)
			_o_throw(ERR_INVALID_INDEX, ParamIndexToString(0, _f_number_buf));
		auto &item = mItem[index];
		if (IS_INVOKE_GET)
			item.ReturnRef(aResultToken);
		else
			if (!item.Assign(*aParam[1]))
				_o_throw(ERR_OUTOFMEM);
		return OK;
	}
	auto result = ObjectMember::Invoke(sMembers, _countof(sMembers), this, aResultToken, aFlags, aParam, aParamCount);
	// There's currently no support for overriding built-in Array members.
	// That aside, this approach allows user-defined properties and methods
	// and built-in Object members (if they aren't shadowed by Array members),
	// which would be difficult with the current design.  This will be fixed
	// when Object is redesigned to allow defining all built-ins via mapping,
	// rather than virtual functions, recursion and special flags.
	if (result == INVOKE_NOT_HANDLED)
		result = Object::Invoke(aResultToken, aThisToken, aFlags, aParam, aParamCount);
	return result;
}

ResultType Array::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case P_Length:
	case P_Capacity:
		if (IS_INVOKE_SET)
		{
			auto arg64 = (UINT64)ParamIndexToInt64(0);
			if (arg64 < 0 || arg64 > MaxIndex || !ParamIndexIsNumeric(0))
				_o_throw(ERR_INVALID_VALUE);
			if (!(aID == P_Capacity ? SetCapacity((index_t)arg64) : SetLength((index_t)arg64)))
				_o_throw(ERR_OUTOFMEM);
			return OK;
		}
		_o_return(aID == P_Capacity ? Capacity() : Length());

	case M_InsertAt:
	case M_Push:
	{
		index_t index;
		if (aID == M_InsertAt)
		{
			index = ParamToZeroIndex(*aParam[0]);
			if (index > mLength || index + (index_t)aParamCount > MaxIndex) // The second condition is very unlikely.
				_o_throw(ERR_PARAM1_INVALID);
			aParam++;
			aParamCount--;
		}
		else
			index = mLength;
		if (!InsertAt(index, aParam, aParamCount))
			_o_throw(ERR_OUTOFMEM);
		return OK;
	}

	case M_RemoveAt:
	case M_Pop:
	{
		index_t index;
		if (aID == M_RemoveAt)
		{
			index = ParamToZeroIndex(*aParam[0]);
			if (index >= mLength)
				_o_throw(ERR_PARAM1_INVALID);
		}
		else
		{
			if (!mLength)
				_o_throw(_T("Array is empty."));
			index = mLength - 1;
		}
		
		index_t count = (index_t)ParamIndexToOptionalInt64(1, 1);
		if (index + count > mLength)
			_o_throw(ERR_PARAM2_INVALID);

		if (aParamCount < 2) // Remove-and-return mode.
		{
			mItem[index].ReturnMove(aResultToken);
			if (aResultToken.Exited())
				return aResultToken.Result();
		}
		
		RemoveAt(index, count);
		return OK;
	}
	
	case M_Has:
	{
		auto index = ParamToZeroIndex(*aParam[0]);
		_o_return(index >= 0 && index < mLength && mItem[index].symbol != SYM_MISSING
			? (index + 1) : 0);
	}

	case M_Delete:
	{
		auto index = ParamToZeroIndex(*aParam[0]);
		if (index < mLength)
		{
			mItem[index].ReturnMove(aResultToken);
			mItem[index].AssignMissing();
		}
		return OK;
	}

	case M_Clone:
		if (auto *arr = Clone())
			_o_return(arr);
		_o_throw(ERR_OUTOFMEM);

	case M__NewEnum:
		_o_return(new Enumerator(this));
	}
	return INVOKE_NOT_HANDLED;
}

Array::index_t Array::ParamToZeroIndex(ExprTokenType &aParam)
{
	if (!TokenIsNumeric(aParam))
		return -1;
	auto index = TokenToInt64(aParam);
	if (index <= 0) // Let -1 be the last item and 0 be the first unused index.
		index += mLength + 1;
	--index; // Convert to zero-based.
	return index >= 0 && index <= MaxIndex ? UINT(index) : BadIndex;
}

Implement_DebugWriteProperty_via_sMembers(Array)


int Array::Enumerator::Next(Var *aVal, Var *aReserved)
{
	if (mIndex < mArray->mLength)
	{
		auto &item = mArray->mItem[mIndex++];
		switch (item.symbol)
		{
		default:	aVal->AssignString(item.string, item.string.Length());	break;
		case SYM_INTEGER:	aVal->Assign(item.n_int64);			break;
		case SYM_FLOAT:		aVal->Assign(item.n_double);		break;
		case SYM_OBJECT:	aVal->Assign(item.object);			break;
		}
		if (aReserved)
			aReserved->Assign();
		return true;
	}
	return false;
}



//
// Enumerator
//

ObjectMember EnumBase::sMembers[] =
{
	Object_Method_(Next, 0, 2, Invoke, 0)
};

ResultType EnumBase::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	Var *var0 = ParamIndexToOptionalVar(0);
	Var *var1 = ParamIndexToOptionalVar(1);
	_o_return(Next(var0, var1));
}

ResultType STDMETHODCALLTYPE EnumBase::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	return ObjectMember::Invoke(sMembers, _countof(sMembers), this, aResultToken, aFlags, aParam, aParamCount);
}

int Object::Enumerator::Next(Var *aKey, Var *aVal)
{
	if (++mOffset < mObject->mFieldCount)
	{
		FieldType &field = mObject->mFields[mOffset];
		if (aKey)
		{
			aKey->Assign(field.name);
		}
		if (aVal)
		{
			ExprTokenType value;
			field.ToToken(value);
			aVal->Assign(value);
		}
		return true;
	}
	return false;
}

int Map::Enumerator::Next(Var *aKey, Var *aVal)
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
			ExprTokenType value;
			field.ToToken(value);
			aVal->Assign(value);
		}
		return true;
	}
	return false;
}

	

//
// Object:: and Map:: Internal Methods
//

Map::FieldType *Map::FindField(IntKeyType val, IndexType left, IndexType right, IndexType &insert_pos)
// left and right must be set by caller to the appropriate bounds within mFields.
{
	IndexType mid;
	// Optimize for common arrays such as [a,b,c] where keys are consecutive numbers starting at 1.
	// In such cases, the needed key can be found immediately.  Benchmarks show that starting the
	// search this way can also benefit sparse arrays, and has little effect on associative arrays
	// (keyed with a precalculated set of 100 or 2000 random integers between 0x10000 and 0x2000000).
	if ((mid = left + (IndexType)val - 1) > right)
	{
		// I couldn't come up with a data set or pattern where the standard starting calculation
		// actually performed better, so start the search by comparing the last element's key.
		// This improves performance when appending to an array via assignment.  Benchmarks show
		// marginal improvements for other cases, probably due to slightly smaller code size.
		//if (--mid != right) // Optimize for appending via incrementing index: this[n++].
		//	mid = (left + right) / 2; // Fall back to standard binary search.
		mid = right;
	}
	else if (mid < left) // val was negative.
		mid = left; // As above.
	for ( ; left <= right; mid = (left + right) / 2)
	{
		FieldType &field = mFields[mid];

		auto result = val - field.key.i;

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

Object::FieldType *Object::FindField(name_t name, IndexType &insert_pos)
// left and right must be set by caller to the appropriate bounds within mFields.
{
	IndexType left = 0, mid, right = mFieldCount - 1;
	int first_char = *name;
	if (first_char <= 'Z' && first_char >= 'A')
		first_char += 32;
	while (left <= right)
	{
		mid = (left + right) / 2;
		
		FieldType &field = mFields[mid];
		
		// key_c contains the lower-case version of field.name[0].  Checking key_c first
		// allows the _tcsicmp() call to be skipped whenever the first character differs.
		// This also means that .name isn't dereferenced, which means one less potential
		// CPU cache miss (where we wait for the data to be pulled from RAM into cache).
		// field.key_c might cause a cache miss, but it's very likely that key.s will be
		// read into cache at the same time (but only the pointer value, not the chars).
		int result = first_char - field.key_c;
		if (!result)
			result = _tcsicmp(name, field.name);
		
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

Map::FieldType *Map::FindField(LPTSTR val, IndexType left, IndexType right, IndexType &insert_pos)
// left and right must be set by caller to the appropriate bounds within mFields.
{
	IndexType mid;
	int first_char = *val;
	while (left <= right)
	{
		mid = (left + right) / 2;

		FieldType &field = mFields[mid];

		// key_c contains key.s[0], cached there for performance.
		int result = first_char - field.key_c;
		if (!result)
			result = _tcscmp(val, field.key.s);

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

Map::FieldType *Map::FindField(SymbolType key_type, KeyType key, IndexType &insert_pos)
// Searches for a field with the given key.  If found, a pointer to the field is returned.  Otherwise
// NULL is returned and insert_pos is set to the index a newly created field should be inserted at.
// key_type and key are output for creating a new field or removing an existing one correctly.
// left and right must indicate the appropriate section of mFields to search, based on key type.
{
	IndexType left, right;

	switch (key_type)
	{
	case SYM_STRING:
		left = mKeyOffsetString;
		right = mFieldCount - 1; // String keys are last in the mFields array.
		return FindField(key.s, left, right, insert_pos);
	case SYM_OBJECT:
		left = mKeyOffsetObject;
		right = mKeyOffsetString - 1; // Object keys end where String keys begin.
		// left and right restrict the search to just the portion with object keys.
		// Reuse the integer search function to reduce code size.  On 32-bit builds,
		// this requires that the upper 32 bits of each key have been initialized.
		return FindField((IntKeyType)(INT_PTR)key.p, left, right, insert_pos);
		//case SYM_INTEGER:
	default:
		left = mKeyOffsetInt;
		right = mKeyOffsetObject - 1; // Int keys end where Object keys begin.
		return FindField(key.i, left, right, insert_pos);
	}
}

void Map::ConvertKey(ExprTokenType &key_token, LPTSTR buf, SymbolType &key_type, KeyType &key)
// Converts key_token to the appropriate key_type and key.
// The exact type of the key is not preserved, since that often produces confusing behaviour;
// for example, guis[WinExist()] := x ... x := guis[A_Gui] would fail because A_Gui returns a
// string.  Strings are converted to integers only where conversion back to string produces
// the same string, so for instance, "01" and " 1 " and "+0x8000" are left as strings.
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
		// Set i to support the way FindField() is used.  Otherwise on 32-bit builds the
		// upper 32 bits would be potentially uninitialized, and searches could fail.
		key.i = (IntKeyType)(INT_PTR)TokenToObject(key_token);
		return;
	}
	if (inner_type == SYM_INTEGER)
	{
		key.i = TokenToInt64(key_token);
		key_type = SYM_INTEGER;
		return;
	}
	key_type = SYM_STRING;
	key.s = TokenToString(key_token, buf);
}

Map::FieldType *Map::FindField(ExprTokenType &key_token, LPTSTR aBuf, SymbolType &key_type, KeyType &key, IndexType &insert_pos)
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

bool Map::SetInternalCapacity(IndexType new_capacity)
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
	
Object::FieldType *Object::Insert(name_t name, IndexType at)
// Inserts a single field with the given key at the given offset.
// Caller must ensure 'at' is the correct offset for this key.
{
	if (mFieldCount == mFieldCountMax && !Expand()  // Attempt to expand if at capacity.
		|| !(name = _tcsdup(name)))  // Attempt to duplicate key-string.
	{	// Out of memory.
		return NULL;
	}
	// There is now definitely room in mFields for a new field.

	FieldType &field = mFields[at];
	if (at < mFieldCount)
		// Move existing fields to make room.
		memmove(&field + 1, &field, (mFieldCount - at) * sizeof(FieldType));
	++mFieldCount; // Only after memmove above.
	
	field.key_c = ctolower(*name);
	
	field.name = name; // Above has already copied string or called key.p->AddRef() as appropriate.
	field.symbol = SYM_STRING;
	field.string.Init(); // Initialize to empty string.  Caller will likely reassign.

	return &field;
}

Map::FieldType *Map::Insert(SymbolType key_type, KeyType key, IndexType at)
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
	if (key_type == SYM_STRING)
		field.key_c = *key.s;
	else
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

	if (aFlags & IF_DEFAULT)
	{
		if (IS_INVOKE_CALL) // May be impossible due to how IF_DEFAULT is used.
			return INVOKE_NOT_HANDLED;
		
		member = IS_INVOKE_SET ? &mSet : &mGet;
	}
	else
	{
		if (!aParamCount) // May be impossible as this[] would use flag IF_DEFAULT.
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
				// Allow changing the GET/SET function, since it's simple and seems harmless.
				if (aParamCount == 2)
					*member = TokenToFunc(*aParam[1]); // Can be NULL.
				return OK;
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

ObjectMember Func::sMembers[] =
{
	Object_Method(Call, 0, MAXP_VARIADIC),
	Object_Method(Bind, 0, MAXP_VARIADIC),
	Object_Method(IsOptional, 0, MAX_FUNCTION_PARAMS),
	Object_Method(IsByRef, 0, MAX_FUNCTION_PARAMS),

	Object_Property_get(Name),
	Object_Property_get(MinParams),
	Object_Property_get(MaxParams),
	Object_Property_get(IsBuiltIn),
	Object_Property_get(IsVariadic)
};

ResultType STDMETHODCALLTYPE Func::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (aFlags == (IT_CALL | IF_DEFAULT))
	{
		Call(aResultToken, aParam, aParamCount);
		return aResultToken.Result();
	}
	return ObjectMember::Invoke(sMembers, _countof(sMembers), this, aResultToken, aFlags, aParam, aParamCount);
}

ResultType Func::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (MemberID(aID))
	{
	case M_Call:
		Call(aResultToken, aParam, aParamCount);
		return aResultToken.Result();

	case M_Bind:
		if (BoundFunc *bf = BoundFunc::Bind(this, aParam, aParamCount, IT_CALL | IF_DEFAULT))
			_o_return(bf);
		_o_throw(ERR_OUTOFMEM);

	case M_IsOptional:
		if (aParamCount)
		{
			int param = ParamIndexToInt(0);
			if (param > 0 && (param <= mParamCount || mIsVariadic))
				_o_return(param > mMinParams);
			else
				_o_throw(ERR_PARAM1_INVALID);
		}
		else
			_o_return(mMinParams != mParamCount || mIsVariadic); // True if any params are optional.
	
	case M_IsByRef:
		if (aParamCount)
		{
			int param = ParamIndexToInt(0);
			if (param <= 0 || param > mParamCount && !mIsVariadic)
				_o_throw(ERR_PARAM1_INVALID);
			if (mIsBuiltIn)
				_o_return(ArgIsOutputVar(param-1));
			_o_return(param <= mParamCount && mParam[param-1].is_byref);
		}
		else
		{
			for (int param = 0; param < mParamCount; ++param)
				if (mParam[param].is_byref)
					_o_return(TRUE);
			_o_return(FALSE);
		}

	case P_Name: _o_return(mName);
	case P_MinParams: _o_return(mMinParams);
	case P_MaxParams: _o_return(mParamCount);
	case P_IsBuiltIn: _o_return(mIsBuiltIn);
	case P_IsVariadic: _o_return(mIsVariadic);
	}
	return INVOKE_NOT_HANDLED;
}


ResultType STDMETHODCALLTYPE BoundFunc::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (aFlags != (IT_CALL | IF_DEFAULT))
	{
		// No methods/properties implemented yet, except Call().
		if (!IS_INVOKE_CALL || _tcsicmp(TokenToString(*aParam[0]), _T("Call")))
			return INVOKE_NOT_HANDLED; // Reserved.
		++aParam;
		--aParamCount;
	}

	// Combine the bound parameters with the supplied parameters.
	int bound_count = mParams->Length();
	if (bound_count > 0)
	{
		ExprTokenType *token = (ExprTokenType *)_alloca(bound_count * sizeof(ExprTokenType));
		ExprTokenType **param = (ExprTokenType **)_alloca((bound_count + aParamCount) * sizeof(ExprTokenType *));
		mParams->ToParams(token, param, NULL, 0);
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
	if (auto params = Array::Create(aParam, aParamCount))
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


ResultType STDMETHODCALLTYPE Closure::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (aFlags != (IT_CALL | IF_DEFAULT))
	{
		if (!IS_INVOKE_CALL || _tcsicmp(ParamIndexToString(0), _T("Call"))) // i.e. not Call.
			return mFunc->Invoke(aResultToken, aThisToken, aFlags, aParam, aParamCount);
		++aParam;
		--aParamCount;
	}
	mFunc->Call(aResultToken, aParam, aParamCount, false, mVars);
	return aResultToken.Result();
}

Closure::~Closure()
{
	mVars->Release();
}


ResultType STDMETHODCALLTYPE Label::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	// Labels are never returned to script, so no need to check flags or parameters.
	return Execute();
}

ResultType LabelPtr::ExecuteInNewThread(TCHAR *aNewThreadDesc, ExprTokenType *aParamValue, int aParamCount, __int64 *aRetVal) const
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
	__int64 retval = 0;
	
	for (MsgMonitorInstance inst (*this); inst.index < inst.count; ++inst.index)
	{
		if (inst.index >= aInitNewThreadIndex) // Re-initialize the thread.
			InitNewThread(0, true, false);
		
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



ResultType MsgMonitorList::Call(ExprTokenType *aParamValue, int aParamCount, UINT aMsg, UCHAR aMsgType, GuiType *aGui, INT_PTR *aRetVal)
{
	ResultType result = OK;
	__int64 retval = 0;
	BOOL thread_used = FALSE;
	
	for (MsgMonitorInstance inst (*this); inst.index < inst.count; ++inst.index)
	{
		MsgMonitorStruct &mon = mMonitor[inst.index];
		if (mon.msg != aMsg || mon.msg_type != aMsgType)
			continue;

		IObject *func = mon.is_method ? aGui->mEventSink : mon.func; // is_method == true implies the GUI has an event sink object.
		LPTSTR method_name = mon.is_method ? mon.method_name : _T("call");

		if (thread_used) // Re-initialize the thread.
			InitNewThread(0, true, false);
		
		// Set last found window (as documented).
		g->hWndLastUsed = aGui->mHwnd;
		
		result = CallMethod(func, func, method_name, aParamValue, aParamCount, &retval);
		if (result == FAIL) // Callback encountered an error.
			break;
		if (result == EARLY_RETURN) // Callback returned a non-empty value.
			break;
		thread_used = TRUE;
	}
	if (aRetVal)
		*aRetVal = (INT_PTR)retval;
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
	ResultType result = Object::Invoke(aResultToken, aThisToken, aFlags, aParam, aParamCount);
	if (result == INVOKE_NOT_HANDLED && (aFlags & (IF_DEFAULT|IF_METAOBJ|IT_CALL)) == (IF_DEFAULT|IF_METAOBJ))
		// No __Item defined; provide fallback behaviour for temporary backward-compatibility.
		// This should be removed once the paradigm shift is complete.
		result = Object::Invoke(aResultToken, aThisToken, aFlags & ~IF_DEFAULT, aParam, aParamCount);
	return result;
}



//
// Buffer
//

ObjectMember BufferObject::sMembers[] =
{
	Object_Property_get(Ptr),
	Object_Property_get_set(Size),
	Object_Property_get(Data)
};

ResultType BufferObject::Invoke(ResultToken &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	return ObjectMember::Invoke(sMembers, _countof(sMembers), this, aResultToken, aFlags, aParam, aParamCount);
}

ResultType BufferObject::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case P_Ptr:
		_o_return((size_t)mData);
	case P_Size:
		if (IS_INVOKE_SET)
		{
			if (!ParamIndexIsNumeric(0))
				_o_throw(ERR_TYPE_MISMATCH);
			auto new_size = ParamIndexToInt64(0);
			if (new_size < 0 || new_size > SIZE_MAX)
				_o_throw(ERR_INVALID_VALUE);
			if (!Resize((size_t)new_size))
				_o_throw(ERR_OUTOFMEM);
			return OK;
		}
		_o_return(mSize);
	case P_Data:
		// Return the data as a binary string, which can be passed to FileAppend().
		_o_return_p((LPTSTR)mData, mSize / sizeof(TCHAR));
	}
	return INVOKE_NOT_HANDLED;
}

ResultType BufferObject::Resize(size_t aNewSize)
{
	auto new_data = realloc(mData, aNewSize);
	if (!new_data)
		return FAIL;
	if (aNewSize > mSize)
		memset((BYTE*)new_data + mSize, 0, aNewSize - mSize);
	mData = new_data;
	mSize = aNewSize;
	return OK;
}


BIF_DECL(BIF_BufferAlloc)
{
	if (!ParamIndexIsNumeric(0))
		_f_throw(ERR_TYPE_MISMATCH);
	auto size = ParamIndexToInt64(0);
	if (size < 0 || size > SIZE_MAX)
		_f_throw(ERR_PARAM1_INVALID);
	auto data = malloc((size_t)size);
	if (!data)
		_f_throw(ERR_OUTOFMEM);
	_f_return(new BufferObject(data, (size_t)size));
}


BIF_DECL(BIF_ClipboardAll)
{
	void *data;
	size_t size;
	if (!aParamCount)
	{
		// Retrieve clipboard contents.
		if (!Var::GetClipboardAll(&data, &size))
			_f_return_FAIL;
	}
	else
	{
		// Use caller-supplied data.
		void *caller_data;
		if (TokenIsPureNumeric(*aParam[0]))
		{
			// Caller supplied an address.
			caller_data = (void *)ParamIndexToIntPtr(0);
			if ((size_t)caller_data < 65536) // Basic check to catch incoming raw addresses that are zero or blank.  On Win32, the first 64KB of address space is always invalid.
				_f_throw(ERR_PARAM1_INVALID);
			size = -1;
		}
		else
		{
			// Caller supplied a binary string or variable, such as from File.RawRead(var, n).
			caller_data = ParamIndexToString(0, NULL, &size);
			size *= sizeof(TCHAR);
		}
		if (!ParamIndexIsOmitted(1))
			size = (size_t)ParamIndexToIntPtr(1);
		else if (size == -1) // i.e. it can be omitted when size != -1 (a string was passed).
			_f_throw(ERR_PARAM2_MUST_NOT_BE_BLANK);
		size_t extra = sizeof(TCHAR); // For an additional null-terminator in case the caller passed invalid data.
		#ifdef UNICODE
		if (size & 1) // Odd; not a multiple of sizeof(WCHAR).
			++extra; // Align the null-terminator.
		#endif
		if (  !(data = malloc(size + extra))  ) // More likely to be due to invalid parameter than out of memory.
			_f_throw(ERR_OUTOFMEM);
		memcpy(data, caller_data, size);
		// Although data returned by GetClipboardAll() should already be terminated with
		// a null UINT, the caller may have passed invalid data.  So align the data to a
		// multiple of sizeof(TCHAR) and terminate with a proper null character in case
		// `this.Data` is used with something expecting a null-terminated string.
		#ifdef UNICODE
		if (size & 1)
			((LPBYTE)data)[size++] = 0; // Size is rounded up so that `this.Data` will not truncate the last byte.
		#endif
		*LPTSTR(LPBYTE(data)+size) = '\0'; // Cast to LPBYTE first because size is in bytes, not TCHARs.
	}
	_f_return(new ClipboardAll(data, size));
}



#ifdef CONFIG_DEBUGGER

void IObject::DebugWriteProperty(IDebugProperties *aDebugger, int aPage, int aPageSize, int aMaxDepth)
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

Implement_DebugWriteProperty_via_sMembers(Func)
Implement_DebugWriteProperty_via_sMembers(BufferObject)

void ObjectMember::DebugWriteProperty(ObjectMember aMembers[], int aMemberCount, IObject *const aThis
	, IDebugProperties *aDebugger, int aPage, int aPageSize, int aMaxDepth)
{
	int num_children = 0;
	for (int imem = 0, iprop = -1; imem < aMemberCount; ++imem)
		if (aMembers[imem].invokeType != IT_CALL)
			++num_children;
	
	DebugCookie cookie;
	aDebugger->BeginProperty(NULL, "object", num_children, cookie);

	if (aMaxDepth)
	{
		int page_begin = aPageSize * aPage, page_end = page_begin + aPageSize;
		for (int imem = 0, iprop = -1; imem < aMemberCount; ++imem)
		{
			auto &member = aMembers[imem];
			if (member.invokeType == IT_CALL || member.minParams > 0)
				continue;
			++iprop;
			if (iprop < page_begin)
				continue;
			if (iprop >= page_end)
				break;
			FuncResult result_token;
			auto result = (aThis->*member.method)(result_token, member.id, IT_GET, NULL, 0);
			aDebugger->WriteProperty(ExprTokenType(member.name), result_token);
			result_token.Free();
		}
	}

	aDebugger->EndProperty(cookie);
}

#endif
