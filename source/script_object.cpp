#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"
#include "application.h"

#include "script_object.h"
#include "script_func_impl.h"
#include "input_object.h"

#include <errno.h> // For ERANGE.
#include <initializer_list>


//
// CallMethod - Invoke a method with no parameters, discarding the result.
//

ResultType CallMethod(IObject *aInvokee, IObject *aThis, LPTSTR aMethodName
	, ExprTokenType *aParamValue, int aParamCount, __int64 *aRetVal // For event handlers.
	, int aExtraFlags, bool aReturnBoolean)
{
	ResultToken result_token;
	TCHAR result_buf[MAX_NUMBER_SIZE];
	result_token.InitResult(result_buf);

	ExprTokenType this_token(aThis);

	ExprTokenType **param = (ExprTokenType **)_alloca(aParamCount * sizeof(ExprTokenType *));
	for (int i = 0; i < aParamCount; ++i)
		param[i] = aParamValue + i;

	ResultType result = aInvokee->Invoke(result_token, IT_CALL | aExtraFlags, aMethodName, this_token, param, aParamCount);

	// Exceptions are thrown by Invoke for too few/many parameters, but not for non-existent method.
	// Check for that here, with the exception that objects are permitted to lack a __Delete method.
	if (result == INVOKE_NOT_HANDLED && !(aExtraFlags & IF_BYPASS_METAFUNC))
		result = ResultToken().UnknownMemberError(this_token, IT_CALL, aMethodName);

	if (result != EARLY_EXIT && result != FAIL)
	{
		if (aReturnBoolean)
			result = TokenToBOOL(result_token) ? CONDITION_TRUE : CONDITION_FALSE;
		else
			// Indicate to caller whether an integer value was returned (for MsgMonitor()).
			result = TokenIsEmptyString(result_token) ? OK : EARLY_RETURN;
	}
	
	if (aRetVal) // Always set this as some callers don't initialize it:
		*aRetVal = result == EARLY_RETURN ? TokenToInt64(result_token) : 0;

	result_token.Free();
	return result;
}


//
// Object::Create - Create a new Object given an array of property name/value pairs.
//

Object *Object::Create()
{
	Object *obj = new Object();
	obj->SetBase(Object::sPrototype);
	return obj;
}

Object *Object::Create(ExprTokenType *aParam[], int aParamCount, ResultToken *apResultToken)
{
	if (aParamCount & 1)
	{
		apResultToken->Error(ERR_PARAM_COUNT_INVALID);
		return NULL; // Odd number of parameters - reserved for future use.
	}

	Object *obj = Object::Create();
	if (aParamCount)
	{
		if (aParamCount > 8)
			// Set initial capacity to avoid multiple expansions.
			// For simplicity, failure is handled by the loop below.
			obj->SetInternalCapacity(aParamCount >> 1);
		// Otherwise, there are 4 or less key-value pairs.  When the first
		// item is inserted, a default initial capacity of 4 will be set.

		TCHAR buf[MAX_NUMBER_SIZE];
		
		for (int i = 0; i + 1 < aParamCount; i += 2)
		{
			if (aParam[i]->symbol == SYM_MISSING || aParam[i+1]->symbol == SYM_MISSING)
				continue; // For simplicity and flexibility.

			auto name = TokenToString(*aParam[i], buf);

			if (!_tcsicmp(name, _T("Base")) && apResultToken)
			{
				auto base = dynamic_cast<Object *>(TokenToObject(*aParam[i + 1]));
				if (!obj->SetBase(base, *apResultToken))
				{
					obj->Release();
					return nullptr;
				}
				continue;
			}

			if (!obj->SetOwnProp(name, *aParam[i + 1]))
			{
				if (apResultToken)
					apResultToken->MemoryError();
				obj->Release();
				return NULL;
			}
		}
	}
	return obj;
}


//
// Map::Create - Create a new Map given an array of key/value pairs.
// Map::SetItems - Add or set items given an array of key/value pairs.
//

Map *Map::Create(ExprTokenType *aParam[], int aParamCount)
{
	ASSERT(!(aParamCount & 1));

	Map *map = new Map();
	map->SetBase(Map::sPrototype);
	if (aParamCount && !map->SetItems(aParam, aParamCount))
	{
		// Out of memory.
		map->Release();
		return NULL;
	}
	return map;
}

ResultType Map::SetItems(ExprTokenType *aParam[], int aParamCount)
{
	ASSERT(!(aParamCount & 1)); // Caller should verify and throw.

	if (!aParamCount)
		return OK;

	// Calculate the maximum number of items that will exist after all items are added.
	// There may be an excess if some items already exist, so instead of allocating this
	// exact amount up front, postpone it until the last possible moment.
	index_t max_capacity_required = mCapacity + (aParamCount >> 1);

	for (int i = 0; i + 1 < aParamCount; i += 2)
	{
		if (aParam[i]->symbol == SYM_MISSING || aParam[i+1]->symbol == SYM_MISSING)
			continue; // For simplicity and flexibility.

		// See comments above.  HasItem() is checked so that the capacity won't be expanded
		// unnecessarily if all of the remaining items already exist.  This produces smaller
		// code than inlining FindItem()/Assign() here and benchmarks faster than allowing
		// unnecessary expansion for a case like Map('c',x,'b',x,'a',x).set('a',y,'b',y).
		if (mCapacity == mCount && !HasItem(*aParam[i]))
			SetInternalCapacity(max_capacity_required);

		if (!SetItem(*aParam[i], *aParam[i + 1]))
			return FAIL; // Out of memory.
	}

	return OK;
}


//
// Cloning
//

// Helper function for temporary implementation of Clone by Object and subclasses.
// Should be eliminated once revision of the object model is complete.
Object *Object::CloneTo(Object &obj)
{
	// Allocate space in destination object.
	auto field_count = mFields.Length();
	if (!obj.SetInternalCapacity(field_count))
	{
		obj.Release();
		return NULL;
	}

	int failure_count = 0; // See comment below.
	index_t i;

	obj.mFields.Length() = field_count;

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
		obj.SetBase(mBase);
	return &obj;
}

Map *Map::CloneTo(Map &obj)
{
	Object::CloneTo(obj);

	if (!obj.SetInternalCapacity(mCount))
	{
		obj.Release();
		return NULL;
	}

	int failure_count = 0; // See Object::CloneT() for comments.
	index_t i;

	obj.mFlags = mFlags;
	obj.mCount = mCount;
	obj.mKeyOffsetObject = mKeyOffsetObject;
	obj.mKeyOffsetString = mKeyOffsetString;
	if (obj.mKeyOffsetObject < 0) // Currently might always evaluate to false.
	{
		obj.mKeyOffsetObject = 0; // aStartOffset excluded all integer and some or all object keys.
		if (obj.mKeyOffsetString < 0)
			obj.mKeyOffsetString = 0; // aStartOffset also excluded some string keys.
	}
	//else no need to check mKeyOffsetString since it should always be >= mKeyOffsetObject.

	for (i = 0; i < mCount; ++i)
	{
		Pair &dst = obj.mItem[i];
		Pair &src = mItem[i];

		// Copy key.
		if (i >= obj.mKeyOffsetString)
		{
			dst.key_c = src.key_c;
			if ( !(dst.key.s = _tcsdup(src.key.s)) )
			{
				// Key allocation failed. At this point, all int and object keys
				// have been set and values for previous items have been copied.
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

ResultType GetEnumerator(IObject *&aEnumerator, ExprTokenType &aEnumerable, int aVarCount, bool aDisplayError)
{
	FuncResult result_token;
	ExprTokenType t_count(aVarCount), *param[] = { &t_count };
	IObject *invokee = TokenToObject(aEnumerable);
	if (!invokee)
		invokee = Object::ValueBase(aEnumerable);
	if (invokee)
	{
		// enum := object.__Enum(number of vars)
		// IF_NEWENUM causes ComObjects to invoke a _NewEnum method or property.
		// IF_BYPASS_METAFUNC causes Objects to skip the __Call meta-function if __Enum is not found.
		auto result = invokee->Invoke(result_token, IT_CALL | IF_NEWENUM | IF_BYPASS_METAFUNC, _T("__Enum"), aEnumerable, param, 1);
		if (result == FAIL || result == EARLY_EXIT)
			return result;
		if (result == INVOKE_NOT_HANDLED)
		{
			aEnumerator = invokee;
			aEnumerator->AddRef();
			return OK;
		}
		aEnumerator = TokenToObject(result_token);
		if (aEnumerator)
			return OK;
	}
	result_token.Free();
	if (aDisplayError)
		g_script.RuntimeError(ERR_TYPE_MISMATCH, _T("__Enum"), FAIL, nullptr, ErrorPrototype::Type);
	return FAIL;
}

ResultType CallEnumerator(IObject *aEnumerator, ExprTokenType *aParam[], int aParamCount, bool aDisplayError)
{
	FuncResult result_token;
	ExprTokenType t_this(aEnumerator);
	for (int i = 0; i < aParamCount; ++i)
		if (aParam[i]->symbol == SYM_OBJECT)
		{
			ASSERT(dynamic_cast<VarRef *>(aParam[i]->object));
			((VarRef *)aParam[i]->object)->Uninitialize(VAR_NEVER_FREE);
		}
	auto result = aEnumerator->Invoke(result_token, IT_CALL, nullptr, t_this, aParam, aParamCount);
	if (result == FAIL || result == EARLY_EXIT || result == INVOKE_NOT_HANDLED)
	{
		if (result == INVOKE_NOT_HANDLED && aDisplayError)
			return g_script.RuntimeError(ERR_NOT_ENUMERABLE, nullptr, FAIL, nullptr, ErrorPrototype::Type); // Object not callable -> wrong type of object.
		return result;
	}
	result = TokenToBOOL(result_token) ? CONDITION_TRUE : CONDITION_FALSE;
	result_token.Free();
	return result;
}

// Calls an Enumerator repeatedly and returns an Array of all first-arg values.
// This is used in conjunction with Array::ToParams to support other objects.
Array *Array::FromEnumerable(ExprTokenType &aEnumerable)
{
	IObject *enumerator;
	auto result = GetEnumerator(enumerator, aEnumerable, 1, true);
	if (result == FAIL || result == EARLY_EXIT)
		return nullptr;
	
	auto varref = new VarRef();
	ExprTokenType tvar { varref }, *param = &tvar;
	Array *vargs = Array::Create();
	for (;;)
	{
		auto result = CallEnumerator(enumerator, &param, 1, true);
		if (result == FAIL)
		{
			vargs->Release();
			vargs = nullptr;
			break;
		}
		if (result != CONDITION_TRUE)
			break;
		ExprTokenType value;
		varref->ToTokenSkipAddRef(value);
		vargs->Append(value);
	}
	varref->Release();
	enumerator->Release();
	return vargs;
}


//
// Array::ToStrings - Used by StrSplit.
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

		// If an exception has been thrown, temporarily clear it for execution of __Delete.
		ResultToken *exc = g->ThrownToken;
		g->ThrownToken = NULL;
		
		// This prevents an erroneous "The current thread will exit" message when an error occurs,
		// by causing LineError() to throw an exception:
		int outer_excptmode = g->ExcptMode;
		g->ExcptMode |= EXCPTMODE_DELETE;

		{
			FuncResult rt;
			CallMeta(_T("__Delete"), rt, ExprTokenType(this), nullptr, 0);
			rt.Free();
		}

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
}


void Map::Clear()
{
	while (mCount)
	{
		--mCount;
		// Copy key before Free() since it might cause re-entry via __delete.
		auto key = mItem[mCount].key;
		mItem[mCount].Free();
		if (mCount >= mKeyOffsetString)
			free(key.s);
		else 
		{
			--mKeyOffsetString;
			if (mCount >= mKeyOffsetObject)
				key.p->Release(); // Might also cause re-entry.
			else
				--mKeyOffsetObject;
		}
	}
}


//
// Invoke - dynamic dispatch
//

ObjectMember Object::sMembers[] =
{
	Object_Method1(Clone, 0, 0),
	Object_Method1(DefineProp, 2, 2),
	Object_Method1(DeleteProp, 1, 1),
	Object_Method1(GetOwnPropDesc, 1, 1),
	Object_Method1(HasOwnProp, 1, 1),
	Object_Method1(OwnProps, 0, 0)
};

LPTSTR Object::sMetaFuncName[] = { _T("__Get"), _T("__Set"), _T("__Call") };

ResultType Object::Invoke(IObject_Invoke_PARAMS_DECL)
{
	// In debug mode, verify aResultToken has been initialized correctly.
	ASSERT(aResultToken.symbol == SYM_STRING && aResultToken.marker && !*aResultToken.marker);
	ASSERT(aResultToken.Result() == OK);

	name_t name;
	if (!aName)
	{
		name = IS_INVOKE_CALL ? _T("Call") : _T("__Item");
		aFlags |= IF_BYPASS_METAFUNC;
	}
	else
		name = aName;

	auto actual_param = aParam; // Actual first parameter between [] or ().
	int actual_param_count = aParamCount; // Actual number of parameters between [] or ().

	bool hasprop = false; // Whether any kind of property was found.
	bool setting = IS_INVOKE_SET;
	bool calling = IS_INVOKE_CALL;
	bool handle_params_recursively = calling;
	ResultToken token_for_recursion;
	IObject *etter = nullptr, *method = nullptr;
	Variant *field = nullptr;
	index_t insert_pos, other_pos;
	Object *that;

	if (setting)
	{
		// Due to the way expression parsing works, the result should never be negative
		// (and any direct callers of Invoke must always pass aParamCount >= 1):
		ASSERT(actual_param_count > 0);
		--actual_param_count;
	}

	for (that = this; that; that = that->mBase)
	{
		// Search each object from this to its most distance base, but set insert_pos only when
		// searching this object, since it needs to be the position we can insert a new field at.
		field = that->FindField(name, that == this ? insert_pos : other_pos);
		if (!field) // 'that' has no own property.
			continue;
		if (field->symbol != SYM_DYNAMIC) // 'that' has a value property.
		{
			if (hasprop && setting)
				// This value property has been overridden with a getter, but no setter.
				// Treat it as read-only rather than allowing the getter to implicitly be overridden.
				field = nullptr;
			hasprop = true;
			// This value property takes precedence over any getter, setter or method defined in a base.
			break;
		}
		hasprop = true;
		// Since above did not break or continue, 'that' has a dynamic property.
		if (calling)
		{
			if (method = field->prop->Method())
			{
				etter = nullptr; // Method takes precedence.
				break;
			}
			// Record the first (most derived) getter, if any, in case there is no method:
			if (!etter)
				etter = field->prop->Getter();
			field = nullptr;
			continue;
		}
		if (actual_param_count > 0 && field->prop->MaxParams == 0) // Prop cannot accept parameters.
		{
			setting = false; // GET this property's value.
			handle_params_recursively = true; // Apply parameters by passing them to value->Invoke().
		}
		// Can this Property actually handle this operation?
		if (setting)
			etter = field->prop->Setter();
		else if (  !(etter = field->prop->Getter()) && !method  )
			method = field->prop->Method(); // Fall back to returning this if no getter is found.
		// Reset field to simplify detection of dynamic property vs. value.
		// Note that field would be reset by the next iteration, if there is one.
		field = nullptr;
		if (etter)
			break;
		// This part of the property isn't implemented here, so keep searching.
		continue;
	} // for (that = each base)

	if (!hasprop && aName)
	{
		// Invoke a meta-function in place of this non-existent property.
		auto result = CallMetaVarg(aFlags, aName, aResultToken, aThisToken, actual_param, actual_param_count);
		if (result != INVOKE_NOT_HANDLED)
			return result;
	}

	if (etter) // Property with getter/setter.
	{
		// Prepare the parameter list: this, [value,] actual_param*
		ExprTokenType this_etter(etter);
		ExprTokenType **prop_param = (ExprTokenType **)_malloca((actual_param_count + 2) * sizeof(ExprTokenType *));
		if (!prop_param)
			return aResultToken.MemoryError();
		prop_param[0] = &aThisToken; // For the hidden "this" parameter in the getter/setter.
		int prop_param_count = 1;
		if (setting)
		{
			// Put the setter's hidden "value" parameter before the other parameters.
			prop_param[prop_param_count++] = actual_param[actual_param_count];
		}
		if (!handle_params_recursively)
		{
			memcpy(prop_param + prop_param_count, actual_param, actual_param_count * sizeof(ExprTokenType *));
			prop_param_count += actual_param_count;
		}
		// Call getter/setter.
		auto result = etter->Invoke(aResultToken, IT_CALL, nullptr, this_etter, prop_param, prop_param_count);
		_freea(prop_param);
		if (result == INVOKE_NOT_HANDLED)
			return aResultToken.UnknownMemberError(this_etter, IT_CALL, nullptr);
		if ((!handle_params_recursively && !calling) || result == FAIL || result == EARLY_EXIT)
			return result;
		// Otherwise, handle_params_recursively || calling.
		token_for_recursion.CopyValueFrom(aResultToken);
		token_for_recursion.mem_to_free = aResultToken.mem_to_free;
		aResultToken.mem_to_free = nullptr;
		aResultToken.SetValue(_T(""));
	}

	if (calling)
	{
		ExprTokenType func_token;

		if (etter)
			func_token.CopyValueFrom(token_for_recursion);
		else if (!field)
			return INVOKE_NOT_HANDLED;
		else if (field->symbol == SYM_DYNAMIC)
			func_token.SetValue(field->prop->Method());
		else
			field->ToToken(func_token);
		auto result = CallAsMethod(func_token, aResultToken, aThisToken, actual_param, actual_param_count);
		if (etter)
			token_for_recursion.Free();
		return result;
	}

	if (actual_param_count > 0)
	{
		// This section handles parameters being passed to a property, such as this.x[y],
		// when that property doesn't accept parameters (i.e. none were declared, or the
		// property is undefined or just a value).
		if (!etter)
		{
			if (field)
				field->ToToken(token_for_recursion);
			else if (method)
				token_for_recursion.SetValue(method);
			else
				return INVOKE_NOT_HANDLED;
		}
		
		if (IS_INVOKE_SET)
			++actual_param_count; // Fix the parameter count.

		IObject *obj_for_recursion = TokenToObject(token_for_recursion);
		if (!obj_for_recursion)
		{
			obj_for_recursion = ValueBase(token_for_recursion);
			aFlags |= IF_SUBSTITUTE_THIS;
		}
		
		// Recursively invoke obj_for_recursion, passing remaining parameters:
		auto result = obj_for_recursion->Invoke(aResultToken, (aFlags & IT_BITMASK)
			, nullptr, token_for_recursion, actual_param, actual_param_count);
		
		if (aResultToken.symbol == SYM_STRING && !aResultToken.mem_to_free && aResultToken.marker != aResultToken.buf)
		{
			// Before releasing obj_for_recursion, make a copy of the string in case it points
			// to memory contained by obj_for_recursion, which might be deleted via Release().
			if (!TokenSetResult(aResultToken, aResultToken.marker, aResultToken.marker_length))
				result = FAIL;
		}
		if (result == INVOKE_NOT_HANDLED)
		{
			// Something like obj.x[y] where obj.x exists but obj.x[y] does not.  Throw here
			// to override the default error message, which would indicate that "x" is unknown.
			result = aResultToken.UnknownMemberError(token_for_recursion, aFlags, nullptr);
		}
		if (etter)
			token_for_recursion.Free();
		return result;
	}

	// SET
	else if (setting)
	{
		if (!field && hasprop || (aFlags & (IF_SUBSTITUTE_THIS | IF_SUPER)))
		{
			if ((aFlags & IF_SUPER) && (field || !hasprop))
			{
				// This is `super.x := y` where x is either a value property or undefined.
				// If aThisToken is an Object, use the base Object implementation of set: create a value property.
				if (auto real_this = dynamic_cast<Object *>(TokenToObject(aThisToken)))
				{
					if (!real_this->SetOwnProp(aName, **actual_param))
						return aResultToken.MemoryError();
					return OK;
				}
			}
			// This property has a getter but no setter; or either IF_SUBSTITUTE_THIS or IF_SUPER was set and above
			// did not return, in which case the property should be considered read-only, since it can't be stored
			// in the actual target object (which is aThisToken, not C++ `this`).
			return hasprop ? aResultToken.Error(ERR_PROPERTY_READONLY, name) : INVOKE_NOT_HANDLED;
		}
		
		if (!field || this != that) // No such property in this object yet.
		{
			if (actual_param[0]->symbol == SYM_MISSING)
				return OK; // No action needed for x.y := unset.
			if (  !(field = Insert(name, insert_pos))  )
				return aResultToken.MemoryError();
		}
		else if (actual_param[0]->symbol == SYM_MISSING) // x.y := unset
		{
			// Completely delete the property, since other sections currently aren't designed to handle properties
			// with no value (unlike Array and Map items).
			mFields.Remove((index_t)(field - mFields), 1);
			return OK;
		}
		if (field->Assign(**actual_param))
			return OK;
		return aResultToken.MemoryError();
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
		else if (method)
		{
			method->AddRef();
			return aResultToken.Return(method);
		}
	}

	// Fell through from one of the sections above: invocation was not handled.
	return INVOKE_NOT_HANDLED;
}


ResultType ObjectBase::Invoke(IObject_Invoke_PARAMS_DECL)
{
	if (auto base = Base())
	{
		aFlags |= IF_SUBSTITUTE_THIS;
		return base->Invoke(IObject_Invoke_PARAMS);
	}
	return INVOKE_NOT_HANDLED;
}


void Object::CallBuiltin(int aID, ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case FID_ObjOwnPropCount:	return PropCount(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjHasOwnProp:		return HasOwnProp(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjGetCapacity:	return GetCapacity(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjSetCapacity:	return SetCapacity(aResultToken, 0, IT_CALL, aParam, aParamCount);
	case FID_ObjOwnProps:		return OwnProps(aResultToken, 0, IT_CALL, aParam, aParamCount);
	}
}


ObjectMember Map::sMembers[] =
{
	Object_Member(__Item, __Item, 0, IT_SET, 1, 1),
	Object_Member(Capacity, Capacity, 0, IT_SET),
	Object_Member(CaseSense, CaseSense, 0, IT_SET),
	Object_Member(Count, Count, 0, IT_GET),
	Object_Method1(__Enum, 0, 1),
	Object_Member(__New, Set, 0, IT_CALL, 0, MAXP_VARIADIC),
	Object_Method1(Clear, 0, 0),
	Object_Method1(Clone, 0, 0),
	Object_Method1(Delete, 1, 1),
	Object_Member(Get, __Item, 0, IT_CALL, 1, 2),
	Object_Method1(Has, 1, 1),
	Object_Method1(Set, 0, MAXP_VARIADIC)  // Allow 0 for flexibility with variadic calls.
};


void Map::__Item(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (!IS_INVOKE_SET) // Get or call.
	{
		if (!GetItem(aResultToken, *aParam[0]))
		{
			if (ParamIndexIsOmitted(1))
			{
				auto result = Invoke(aResultToken, IT_GET, _T("Default"), ExprTokenType { this }, nullptr, 0);
				if (result == INVOKE_NOT_HANDLED)
					_o_throw(ERR_ITEM_UNSET, *aParam[0], ErrorPrototype::UnsetItem);
				return;
			}
			// Otherwise, caller provided a default value.
			aResultToken.CopyValueFrom(*aParam[1]);
		}
		if (aResultToken.symbol == SYM_OBJECT)
			aResultToken.object->AddRef();
		return;
	}
	else
	{
		if (!SetItem(*aParam[1], *aParam[0]))
			_o_throw_oom;
	}
}


void Map::Set(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (aParamCount & 1)
		_o_throw(ERR_PARAM_COUNT_INVALID);
	if (!SetItems(aParam, aParamCount))
		_o_throw_oom;
	AddRef();
	_o_return(this);
}



//
// Internal
//

ResultType Object::CallAsMethod(ExprTokenType &aFunc, ResultToken &aResultToken, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount)
{
	auto func = TokenToObject(aFunc);
	if (!func)
		func = ValueBase(aFunc);
	ExprTokenType **param = (ExprTokenType **)_malloca((aParamCount + 1) * sizeof(ExprTokenType *));
	if (!param)
		return aResultToken.MemoryError();
	param[0] = &aThisToken;
	memcpy(param + 1, aParam, aParamCount * sizeof(ExprTokenType *));
	// return %func%(this, aParam*)
	auto invoke_result = func->Invoke(aResultToken, IT_CALL, nullptr, aFunc, param, aParamCount + 1);
	_freea(param);
	return invoke_result;
}

ResultType Object::CallMeta(LPTSTR aName, ResultToken &aResultToken, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount)
{
	IObject *method;
	if (method = GetMethod(aName))
	{
		return CallAsMethod(ExprTokenType(method), aResultToken, aThisToken, aParam, aParamCount);
	}
	return INVOKE_NOT_HANDLED;
}

ResultType Object::CallMetaVarg(int aFlags, LPTSTR aName, ResultToken &aResultToken, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount)
{
	auto func = GetMethod(sMetaFuncName[INVOKE_TYPE]);
	if (!func)
		return INVOKE_NOT_HANDLED;
	auto vargs = Array::Create(aParam, aParamCount);
	if (!vargs)
		return aResultToken.MemoryError();
	ExprTokenType name_token(aName), args_token(vargs), *param[4];
	param[0] = &aThisToken; // this
	param[1] = &name_token; // name
	param[2] = &args_token; // args
	int param_count = 3;
	if (IS_INVOKE_SET)
		param[param_count++] = aParam[aParamCount]; // value
	// return %func%(this, name, args [, value])
	ResultType aResult = func->Invoke(aResultToken, IT_CALL, nullptr, ExprTokenType(func), param, param_count);
	vargs->Release();
	return aResult;
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
	auto &obj = *(Object *)GetOwnPropObj(_T("Prototype"));
	// Each variable declaration created a 'missing' property in the class or prototype object to prevent
	// duplicate or conflicting declarations.  Remove them now so that the declaration acts like a normal
	// assignment (i.e. invokes property setters and __Set), for flexibility and consistency.
	RemoveMissingProperties();
	obj.RemoveMissingProperties();
}

void Object::RemoveMissingProperties()
{
	for (index_t i = mFields.Length(); i > 0; )
	{
		i--;
		if (mFields[i].symbol == SYM_MISSING)
			mFields.Remove(i, 1);
	}
}



bool ObjectBase::IsOfType(Object *aPrototype)
{
	auto base = Base();
	return base == aPrototype || base->IsDerivedFrom(aPrototype);
}

bool Object::IsOfType(Object *aPrototype)
{
	return aPrototype == Object::sPrototype || (!IsClassPrototype() && IsDerivedFrom(aPrototype));
}


bool Object::IsDerivedFrom(IObject *aBase)
{
	Object *base;
	for (base = mBase; base; base = base->mBase)
		if (base == aBase)
			return true;
	return false;
}


Object *Object::GetNativeBase()
{
	Object *base;
	for (base = mBase; base; base = base->mBase)
		if (base->IsNativeClassPrototype())
			return base;
	return nullptr;
}


bool Object::CanSetBase(Object *aBase)
{
	auto new_native_base = (!aBase || aBase->IsNativeClassPrototype())
		? aBase : aBase->GetNativeBase();
	return new_native_base == GetNativeBase() // Cannot change native type.
		&& !aBase->IsDerivedFrom(this); // Cannot create loops.
}


ResultType Object::SetBase(Object *aNewBase, ResultToken &aResultToken)
{
	if (!CanSetBase(aNewBase))
		return aResultToken.ValueError(ERR_INVALID_BASE);
	SetBase(aNewBase);
	return OK;
}


//
// Object::Type() - Returns the object's type/class name.
//

LPTSTR Object::Type()
{
	Object *base;
	ExprTokenType value;
	if (GetOwnProp(value, _T("__Class")))
		return _T("Prototype"); // This object is a prototype.
	for (base = mBase; base; base = base->mBase)
		if (base->GetOwnProp(value, _T("__Class")))
			return TokenToString(value); // This object is an instance of that class.
	return _T("Object"); // Provide a default in case __Class has been removed from all of the base objects.
}


Object *Object::CreateClass(Object *aPrototype)
{
	auto cls = new Object();
	cls->SetBase(sClassPrototype);
	cls->SetOwnProp(_T("Prototype"), aPrototype);
	return cls;
}


Object *Object::CreatePrototype(LPTSTR aClassName, Object *aBase)
{
	auto obj = new Object();
	obj->mFlags |= ClassPrototype;
	obj->SetOwnProp(_T("__Class"), ExprTokenType(aClassName));
	obj->SetBase(aBase);
	return obj;
}


Object *Object::CreatePrototype(LPTSTR aClassName, Object *aBase, ObjectMember aMember[], int aMemberCount)
{
	auto obj = CreatePrototype(aClassName, aBase);
	return DefineMembers(obj, aClassName, aMember, aMemberCount);
}

Object *Object::CreatePrototype(LPTSTR aClassName, Object *aBase, ObjectMemberMd aMember[], int aMemberCount)
{
	auto obj = CreatePrototype(aClassName, aBase);
	return DefineMetadataMembers(obj, aClassName, aMember, aMemberCount);
}

Object *Object::CreatePrototype(LPTSTR aClassName, Object *aBase, ObjectMemberListType aMember, int aMemberCount)
{
	if (aMember.duck)
		return CreatePrototype(aClassName, aBase, aMember.duck, aMemberCount);
	else
		return CreatePrototype(aClassName, aBase, aMember.meta, aMemberCount);
}


Object *Object::DefineMembers(Object *obj, LPTSTR aClassName, ObjectMember aMember[], int aMemberCount)
{
	if (aMemberCount)
		obj->mFlags |= NativeClassPrototype;

	TCHAR full_name[MAX_VAR_NAME_LENGTH + 1];
	TCHAR *name = full_name + _stprintf(full_name, _T("%s.Prototype."), aClassName);

	for (int i = 0; i < aMemberCount; ++i)
	{
		const auto &member = aMember[i];
		_tcscpy(name, member.name);
		if (member.invokeType == IT_CALL)
		{
			auto func = new BuiltInMethod(SimpleHeap::Alloc(full_name));
			func->mBIM = member.method;
			func->mMID = member.id;
			func->mMIT = IT_CALL;
			func->mMinParams = member.minParams + 1; // Includes `this`.
			func->mIsVariadic = member.maxParams == MAXP_VARIADIC;
			func->mParamCount = func->mIsVariadic ? func->mMinParams : member.maxParams + 1;
			func->mClass = obj; // AddRef not needed since neither mClass nor our caller's reference to obj is ever Released.
			obj->DefineMethod(member.name, func);
			func->Release();
		}
		else
		{
			auto prop = obj->DefineProperty(name);
			prop->MinParams = member.minParams;
			prop->MaxParams = member.maxParams;
			
			auto op_name = _tcschr(name, '\0');

			_tcscpy(op_name, _T(".Get"));
			auto func = new BuiltInMethod(SimpleHeap::Alloc(full_name));
			func->mBIM = member.method;
			func->mMID = member.id;
			func->mMIT = IT_GET;
			func->mMinParams = member.minParams + 1; // Includes `this`.
			func->mParamCount = member.maxParams + 1;
			func->mIsVariadic = member.maxParams == MAXP_VARIADIC;
			func->mClass = obj;
			prop->SetGetter(func);
			func->Release();
			
			if (member.invokeType == IT_SET)
			{
				_tcscpy(op_name, _T(".Set"));
				func = new BuiltInMethod(SimpleHeap::Alloc(full_name));
				func->mBIM = member.method;
				func->mMID = member.id;
				func->mMIT = IT_SET;
				func->mMinParams = member.minParams + 2; // Includes `this` and `value`.
				func->mParamCount = member.maxParams + 2;
				func->mIsVariadic = member.maxParams == MAXP_VARIADIC;
				func->mClass = obj;
				prop->SetSetter(func);
				func->Release();
			}
		}
	}

	return obj;
}

Object *Object::CreateClass(LPTSTR aClassName, Object *aBase, Object *aPrototype, ClassFactoryDef aFactory)
{
	auto class_obj = CreateClass(aPrototype);

	class_obj->SetBase(aBase);

	if (aFactory.call)
	{
		TCHAR full_name[MAX_VAR_NAME_LENGTH + 1];
		_stprintf(full_name, _T("%s.Call"), aClassName);
		auto ctor = new BuiltInFunc(SimpleHeap::Alloc(full_name));
		ctor->mBIF = aFactory.call;
		ctor->mFID = FID_Object_New;
		ctor->mMinParams = aFactory.min_params; // Usually 1, the class object.
		ctor->mParamCount = aFactory.max_params;
		ctor->mIsVariadic = aFactory.is_variadic; // Usually variadic since __new(...) may be redefined/overridden.
		class_obj->DefineMethod(_T("Call"), ctor);
		ctor->Release();
	}

	auto var = g_script.FindOrAddVar(aClassName, 0, VAR_DECLARE_GLOBAL);
	var->AssignSkipAddRef(class_obj);
	var->MakeReadOnly();

	return class_obj;
}


//
// Object:: and Map:: Built-ins
//

void Object::DeleteProp(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto field = FindField(ParamIndexToString(0, _f_number_buf));
	if (!field)
		_o_return_empty;
	field->ReturnMove(aResultToken); // Return the removed value.
	mFields.Remove((index_t)(field - mFields), 1);
	_o_return_empty;
}

void Map::Delete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	Pair *item;
	index_t pos;
	SymbolType key_type;
	Key key;

	if (item = FindItem(*aParam[0], _f_number_buf, key_type, key, pos))
		pos = index_t(item - mItem); // else min_pos was already set by FindItem.

	if (!item) // Nothing to remove.
	{
		// Our return value when only one arg is given is supposed to be the value
		// removed from this[arg].  Since this[arg] would throw an exception...
		_o_throw(ERR_ITEM_UNSET, *aParam[0], ErrorPrototype::UnsetItem);
	}
	// Set return value to the removed item.
	item->ReturnMove(aResultToken);
	// Copy item to temporary memory so that Free() and Release() can be postponed,
	// in case they cause re-entry via __delete.  ReturnMove() may have transferred
	// an object value, but not the key or Property getter/setter.
	auto copy = (Pair *)_alloca(sizeof(*item));
	memcpy(copy, item, sizeof(*item));
	// Remove item.
	memmove(item, item + 1, (mCount - (pos + 1)) * sizeof(Pair));
	mCount--;
	// Free item and keys.
	copy->Free();
	if (key_type == SYM_STRING)
		free(copy->key.s);
	else // i.e. SYM_OBJECT or SYM_INTEGER
	{
		mKeyOffsetString--;
		if (key_type == SYM_INTEGER)
			mKeyOffsetObject--;
		else
			copy->key.p->Release();
	}
	_o_return_retval;
}


void Map::Clear(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	Clear();
	_o_return_empty;
}


void Object::PropCount(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return((__int64)mFields.Length());
}

void Map::Count(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return((__int64)mCount);
}

void Map::CaseSense(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (IS_INVOKE_GET)
	{
		if (mFlags & MapUseLocale)
			_o_return_p(_T("Locale"));
		if (mFlags & MapCaseless)
			_o_return_p(_T("Off"));
		_o_return_p(_T("On"));
	}

	// Do not permit a change of flags if the Map contains string keys, as the expected order
	// may not match the actual order.  To simplify the error message and documentation,
	// the Map must be empty of other types of keys as well.
	if (mCount)
		_o_throw(_T("Map must be empty"));

	switch (TokenToStringCase(*aParam[0]))
	{
	case SCS_SENSITIVE:
		mFlags &= ~(MapCaseless | MapUseLocale);
		break;
	case SCS_INSENSITIVE_LOCALE:
		mFlags |= (MapCaseless | MapUseLocale);
		break;
	case SCS_INSENSITIVE:
		mFlags = (mFlags | MapCaseless) & ~MapUseLocale;
		break;
	default:
		_o_throw(ERR_INVALID_VALUE, *aParam[0], ErrorPrototype::Value);
	}
}

void Object::GetCapacity(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return(mFields.Capacity());
}

void Object::SetCapacity(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (!ParamIndexIsNumeric(0))
	{
		aResultToken.ParamError(1, aParam[0], _T("Number")); // Param index differs because this is actually the global function ObjSetCapacity().
		return;
	}

	index_t desired_count = (index_t)ParamIndexToInt64(0);
	if (desired_count < mFields.Length())
	{
		// It doesn't seem intuitive to allow SetCapacity to truncate the fields array, so just reallocate
		// as necessary to remove any unused space.  Allow negative values since SetCapacity(-1) seems more
		// intuitive than SetCapacity(0) when the contents aren't being discarded.
		desired_count = mFields.Length();
	}
	if (desired_count == 0)
	{
		mFields.Free();
		ASSERT(desired_count == mFields.Capacity());
	}
	if (desired_count == mFields.Capacity() || SetInternalCapacity(desired_count))
	{
		_o_return(mFields.Capacity());
	}
	// At this point, failure isn't critical since nothing is being stored yet.  However, it might be easier to
	// debug if an error is thrown here rather than possibly later, when the array attempts to resize itself to
	// fit new items.  This also avoids the need for scripts to check if the return value is less than expected:
	_o_throw_oom;
}

void Map::Capacity(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (IS_INVOKE_GET)
	{
		_o_return(mCapacity);
	}

	if (!ParamIndexIsNumeric(0))
		_o_throw_type(_T("Number"), *aParam[0]);

	index_t desired_count = (index_t)ParamIndexToInt64(0);
	if (desired_count < mCount)
	{
		// It doesn't seem intuitive to allow SetCapacity to truncate the item array, so just reallocate
		// as necessary to remove any unused space.  Allow negative values since SetCapacity(-1) seems more
		// intuitive than SetCapacity(0) when the contents aren't being discarded.
		desired_count = mCount;
	}
	if (!desired_count)
	{
		if (mItem)
		{
			free(mItem);
			mItem = nullptr;
			mCapacity = 0;
		}
		//else mCapacity should already be 0.
		// Since mCapacity and desired_size are both 0, below will return 0 and won't call SetInternalCapacity.
	}
	if (desired_count == mCapacity || SetInternalCapacity(desired_count))
	{
		_o_return(mCapacity);
	}
	// At this point, failure isn't critical since nothing is being stored yet.  However, it might be easier to
	// debug if an error is thrown here rather than possibly later, when the array attempts to resize itself to
	// fit new items.  This also avoids the need for scripts to check if the return value is less than expected:
	_o_throw_oom;
}

void Object::OwnProps(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return(new IndexEnumerator(this, ParamIndexToOptionalInt(0, 1)
		, static_cast<IndexEnumerator::Callback>(&Object::GetEnumProp)));
}

void Map::__Enum(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return(new IndexEnumerator(this, ParamIndexToOptionalInt(0, 1)
		, static_cast<IndexEnumerator::Callback>(&Map::GetEnumItem)));
}

void Object::HasOwnProp(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return(FindField(ParamIndexToString(0, _f_number_buf)) != nullptr);
}

void Map::Has(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	SymbolType key_type;
	Key key;
	index_t insert_pos;
	auto item = FindItem(*aParam[0], _f_number_buf, /*out*/ key_type, /*out*/ key, /*out*/ insert_pos);
	_o_return(item != nullptr);
}

void Object::Clone(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (GetNativeBase() != Object::sPrototype)
		_o_throw(ERR_TYPE_MISMATCH, ErrorPrototype::Type); // Cannot construct an instance of this class using Object::Clone().
	auto clone = new Object();
	if (!CloneTo(*clone))
		_o_throw_oom;	
	_o_return(clone);
}

void Map::Clone(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto clone = new Map();
	if (!CloneTo(*clone))
		_o_throw_oom;
	_o_return(clone);
}

bool Object::DefineMethod(name_t aName, IObject *aFunc)
{
	if (auto prop = DefineProperty(aName))
	{
		prop->SetMethod(aFunc);
		return true;
	}
	return false;
}

Property *Object::DefineProperty(name_t aName)
{
	index_t insert_pos;
	auto field = FindField(aName, insert_pos);
	if (!field && !(field = Insert(aName, insert_pos)))
		return nullptr;
	if (field->symbol != SYM_DYNAMIC)
	{
		field->Free();
		field->symbol = SYM_DYNAMIC;
		field->prop = new Property();
	}
	return field->prop;
}

ResultType GetObjMaxParams(IObject *aObj, int &aMaxParams, ResultToken &aResultToken)
{
	__int64 propval = 0;
	auto result = GetObjectIntProperty(aObj, _T("MaxParams"), propval, aResultToken, true);
	switch (result)
	{
	case FAIL:
	case EARLY_EXIT:
		return result;
	case OK:
		aMaxParams = (int)propval;
		propval = 0;
		result = GetObjectIntProperty(aObj, _T("IsVariadic"), propval, aResultToken, true);
		switch (result)
		{
		case FAIL:
		case EARLY_EXIT:
			return result;
		case INVOKE_NOT_HANDLED:
			return OK;
		case OK:
			if (propval)
				aMaxParams = INT_MAX;
		}
	}
	return result;
}

void Object::DefineProp(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto name = ParamIndexToString(0, _f_number_buf);
	if (!*name)
		_o_throw_param(0);
	ExprTokenType getter, setter, method, value;
	getter.symbol = SYM_INVALID;
	setter.symbol = SYM_INVALID;
	method.symbol = SYM_INVALID;
	value.symbol = SYM_INVALID;
	auto desc = dynamic_cast<Object *>(ParamIndexToObject(1));
	if (!desc // Must be an Object.
		|| desc->GetOwnProp(getter, _T("Get")) && getter.symbol != SYM_OBJECT  // If defined, must be an object.
		|| desc->GetOwnProp(setter, _T("Set")) && setter.symbol != SYM_OBJECT
		|| desc->GetOwnProp(method, _T("Call")) && method.symbol != SYM_OBJECT
		|| desc->GetOwnProp(value, _T("Value")) && (getter.symbol != SYM_INVALID || setter.symbol != SYM_INVALID || method.symbol != SYM_INVALID)
		// To help prevent errors, throw if none of the above properties were present.  This also serves to
		// reserve some cases for possible future use, such as passing a function object to imply {get:...}.
		|| getter.symbol == SYM_INVALID && setter.symbol == SYM_INVALID && method.symbol == SYM_INVALID && value.symbol == SYM_INVALID)
		_o_throw_param(1);
	if (value.symbol != SYM_INVALID) // Above already verified that neither Get nor Set was present.
	{
		if (!SetOwnProp(name, value))
			_o_throw_oom;
		AddRef();
		_o_return(this);
	}
	auto prop = DefineProperty(name);
	if (!prop)
		_o_throw_oom;
	if (getter.symbol == SYM_OBJECT) prop->SetGetter(getter.object);
	if (setter.symbol == SYM_OBJECT) prop->SetSetter(setter.object);
	if (method.symbol == SYM_OBJECT) prop->SetMethod(method.object);
	prop->MaxParams = -1;
	if (auto obj = prop->Getter())
	{
		int max_params;
		switch (GetObjMaxParams(obj, max_params, aResultToken))
		{
		case FAIL:
		case EARLY_EXIT:
			return;
		case OK:
			prop->MaxParams = max_params - 1;
		}
	}
	if (auto obj = prop->Setter())
	{
		int max_params;
		switch (GetObjMaxParams(obj, max_params, aResultToken))
		{
		case FAIL:
		case EARLY_EXIT:
			return;
		case OK:
			if (prop->MaxParams < max_params - 2)
				prop->MaxParams = max_params - 2;
		}
	}
	AddRef();
	_o_return(this);
}

void Object::GetOwnPropDesc(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto name = ParamIndexToString(0, _f_number_buf);
	if (!*name)
		_o_throw_param(0);
	auto field = FindField(name);
	if (!field)
		_o__ret(aResultToken.UnknownMemberError(ExprTokenType(this), IT_GET, name));
	auto desc = Object::Create();
	desc->SetInternalCapacity(field->symbol == SYM_DYNAMIC ? 3 : 1);
	if (field->symbol == SYM_DYNAMIC)
	{
		if (auto getter = field->prop->Getter()) desc->SetOwnProp(_T("Get"), getter);
		if (auto setter = field->prop->Setter()) desc->SetOwnProp(_T("Set"), setter);
		if (auto method = field->prop->Method()) desc->SetOwnProp(_T("Call"), method);
	}
	else
	{
		ExprTokenType value;
		field->ToToken(value);
		desc->SetOwnProp(_T("Value"), value);
	}
	_o_return(desc);
}


//
// Class objects
//

ResultType Object::New(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	Object *base = dynamic_cast<Object *>(ParamIndexToObject(0));
	Object *proto = base ? dynamic_cast<Object *>(base->GetOwnPropObj(_T("Prototype"))) : nullptr;
	if (!proto)
	{
		Release();
		return aResultToken.ParamError(0, aParam[0]);
	}
	if (!SetBase(proto, aResultToken))
	{
		Release();
		return FAIL;
	}
	return Construct(aResultToken, aParam + 1, aParamCount - 1);
}

ResultType Object::Construct(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	ExprTokenType this_token(this);
	ResultType result;

	// __Init was added so that instance variables can be initialized in the correct order
	// (beginning at the root class and ending at class_object) before __New is called.
	// It shouldn't be explicitly defined by the user, but auto-generated in DefineClassVars().
	result = CallMeta(_T("__Init"), aResultToken, this_token, nullptr, 0);
	if (result != INVOKE_NOT_HANDLED)
	{
		// It's possible that __Init is user-defined (despite recommendations in the
		// documentation) or built-in, so make sure the return value, if any, is freed:
		aResultToken.Free();
		// Reset to defaults for __New, invoked below.
		aResultToken.InitResult(aResultToken.buf);
		if (result == FAIL || result == EARLY_EXIT) // Checked only after Free() and InitResult() as caller might expect mem_to_free == NULL.
		{
			Release();
			return aResultToken.SetExitResult(result); // SetExitResult is necessary because result was reset by InitResult.
		}
	}

	// __New may be defined by the script for custom initialization code.
	result = CallMeta(_T("__New"), aResultToken, this_token, aParam, aParamCount);
	aResultToken.Free();
	if (result == INVOKE_NOT_HANDLED && aParamCount)
	{
		// Maybe the caller expects the parameters to be used in some way, but they won't
		// since there's no __New.  Treat it the same as having __New without parameters.
		result = aResultToken.Error(ERR_TOO_MANY_PARAMS);
	}
	if (result == FAIL || result == EARLY_EXIT)
	{
		// An error was raised within __New() or while trying to call it, or Exit was called.
		Release();
		return result;
	}

	aResultToken.SetValue(this); // No AddRef() since Object::New() would need to Release().
	return aResultToken.SetResult(OK);
}

BIF_DECL(Any___Init)
{
	_f_return_empty;
}


//
// Object::Variant
//

void Object::Variant::Minit()
{
	symbol = SYM_MISSING;
	new (&string) String();
}

void Object::Variant::AssignEmptyString()
{
	Free();
	symbol = SYM_STRING;
	new (&string) String();
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
		new (&string) String();
		return Assign(val.string, val.string.Length(), true); // Pass true to conserve memory (no space is allowed for future expansion).
	case SYM_OBJECT:
		(object = val.object)->AddRef();
		break;
	case SYM_DYNAMIC:
		prop = new Property();
		prop->SetGetter(val.prop->Getter());
		prop->SetSetter(val.prop->Setter());
		prop->SetMethod(val.prop->Method());
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
	case SYM_DYNAMIC:
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
		aToken.marker = string;
		aToken.marker_length = string.Length();
		break;
	case SYM_DYNAMIC:
		aToken.SetValue(_T(""), 0);
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
	switch (symbol)
	{
	case SYM_STRING: string.~String(); break;
	case SYM_OBJECT: object->Release(); break;
	case SYM_DYNAMIC: delete prop; break;
	}
}



//
// Array
//

ResultType Array::SetCapacity(index_t aNewCapacity)
{
	if (mLength > aNewCapacity)
		RemoveAt(aNewCapacity, mLength - aNewCapacity);
	auto new_item = (Variant *)realloc(mItem, sizeof(Variant) * aNewCapacity);
	if (!new_item && aNewCapacity)
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
	arr->SetBase(Array::sPrototype);
	if (!aCount || arr->InsertAt(0, aValue, aCount))
		return arr;
	arr->Release();
	return nullptr;
}

Array *Array::Clone()
{
	auto arr = new Array();
	if (!CloneTo(*arr))
		return nullptr; // CloneTo() released arr.
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
	Object_Property_get_set(__Item, 1, 1),
	Object_Property_get_set(Capacity),
	Object_Property_get_set(Length),
	Object_Member(__New, Invoke, M_Push, IT_CALL, 0, MAXP_VARIADIC),
	Object_Method(__Enum, 0, 1),
	Object_Method(Clone, 0, 0),
	Object_Method(Delete, 1, 1),
	Object_Method(Get, 1, 2),
	Object_Method(Has, 1, 1),
	Object_Method(InsertAt, 1, MAXP_VARIADIC),
	Object_Method(Pop, 0, 0),
	Object_Method(Push, 0, MAXP_VARIADIC),
	Object_Method(RemoveAt, 1, 2)
};

void Array::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case P___Item:
	case M_Get:
	{
		auto index = ParamToZeroIndex(*aParam[IS_INVOKE_SET ? 1 : 0]);
		if (index >= mLength)
			_o_throw(ERR_INVALID_INDEX, *aParam[IS_INVOKE_SET ? 1 : 0], ErrorPrototype::Index);
		auto &item = mItem[index];
		if (IS_INVOKE_SET)
		{
			if (!item.Assign(*aParam[0]))
				_o_throw_oom;
			return;
		}
		if (item.symbol == SYM_MISSING)
		{
			if (!ParamIndexIsOmitted(1)) // Get(index, default)
			{
				aResultToken.CopyValueFrom(*aParam[1]);
				return;
			}
			auto result = Object::Invoke(aResultToken, IT_GET, _T("Default"), ExprTokenType{this}, nullptr, 0);
			if (result != INVOKE_NOT_HANDLED)
				_o_return_retval;
			_o_throw(ERR_ITEM_UNSET, *aParam[0], ErrorPrototype::UnsetItem);
		}
		item.ReturnRef(aResultToken);
		_o_return_retval;
	}

	case P_Length:
	case P_Capacity:
		if (IS_INVOKE_SET)
		{
			if (!ParamIndexIsNumeric(0))
				_o_throw_type(_T("Number"), *aParam[0]);
			auto arg64 = (UINT64)ParamIndexToInt64(0);
			if (arg64 < 0 || arg64 > MaxIndex)
				_o_throw_value(ERR_INVALID_VALUE);
			if (!(aID == P_Capacity ? SetCapacity((index_t)arg64) : SetLength((index_t)arg64)))
				_o_throw_oom;
			return;
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
				_o_throw_param(0);
			aParam++;
			aParamCount--;
		}
		else
			index = mLength;
		if (!InsertAt(index, aParam, aParamCount))
			_o_throw_oom;
		_o_return_empty;
	}

	case M_RemoveAt:
	case M_Pop:
	{
		index_t index;
		if (aID == M_RemoveAt)
		{
			index = ParamToZeroIndex(*aParam[0]);
			if (index >= mLength)
				_o_throw_param(0);
		}
		else
		{
			if (!mLength)
				_o_throw(_T("Array is empty."));
			index = mLength - 1;
		}
		
		index_t count = 1;
		if (!ParamIndexIsOmitted(1))
		{
			Throw_if_Param_NaN(1);
			count = (index_t)ParamIndexToInt64(1);
		}
		if (index + count > mLength)
			_o_throw_param(1);

		if (aParamCount < 2) // Remove-and-return mode.
		{
			mItem[index].ReturnMove(aResultToken);
			if (aResultToken.Exited())
				return;
		}
		
		RemoveAt(index, count);
		return;
	}
	
	case M_Has:
	{
		auto index = ParamToZeroIndex(*aParam[0]);
		_o_return(index >= 0 && index < mLength && mItem[index].symbol != SYM_MISSING);
	}

	case M_Delete:
	{
		auto index = ParamToZeroIndex(*aParam[0]);
		if (index >= mLength)
			_o_throw_param(0);
		mItem[index].ReturnMove(aResultToken);
		mItem[index].AssignMissing();
		_o_return_retval;
	}

	case M_Clone:
		if (auto *arr = Clone())
			_o_return(arr);
		_o_throw_oom;

	case M___Enum:
		_o_return(new IndexEnumerator(this, ParamIndexToOptionalInt(0, 1)
			, static_cast<IndexEnumerator::Callback>(&Array::GetEnumItem)));
	}
}

Array::index_t Array::ParamToZeroIndex(ExprTokenType &aParam)
{
	if (!TokenIsNumeric(aParam))
		return BadIndex;
	auto index = TokenToInt64(aParam);
	if (index <= 0) // Let -1 be the last item and 0 be the first unused index.
		index += mLength + 1;
	--index; // Convert to zero-based.
	return index >= 0 && index <= MaxIndex ? UINT(index) : BadIndex;
}


ResultType Array::GetEnumItem(UINT &aIndex, Var *aVal, Var *aReserved, int aVarCount)
{
	if (aIndex < mLength)
	{
		if (aVarCount > 1)
		{
			// Put the index first, only when there are two parameters.
			if (aVal)
				aVal->Assign((__int64)aIndex + 1);
			aVal = aReserved;
		}
		if (aVal)
		{
			auto &item = mItem[aIndex];
			switch (item.symbol)
			{
			default:
				if (item.symbol == SYM_MISSING)
					aVal->Uninitialize();
				else
					aVal->AssignString(item.string, item.string.Length());
				
				break;
			case SYM_INTEGER:	aVal->Assign(item.n_int64);			break;
			case SYM_FLOAT:		aVal->Assign(item.n_double);		break;
			case SYM_OBJECT:	aVal->Assign(item.object);			break;
			}
		}
		return CONDITION_TRUE;
	}
	return CONDITION_FALSE;
}



//
// Enumerator
//

bool EnumBase::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	Var *var0 = ParamIndexToOutputVar(0);
	Var *var1 = ParamIndexToOutputVar(1);
	auto result = Next(var0, var1);
	switch (result)
	{
	case CONDITION_TRUE:
	case CONDITION_FALSE:
		aResultToken.SetValue(result == CONDITION_TRUE);
		return true;
	default: // Probably FAIL or EARLY_EXIT.
		aResultToken.SetExitResult(result);
		return false;
	}
}


ResultType IndexEnumerator::Next(Var *var0, Var *var1)
{
	return (mObject->*mGetItem)(++mIndex, var0, var1, mParamCount);
}


ResultType Object::GetEnumProp(UINT &aIndex, Var *aName, Var *aVal, int aVarCount)
{
	for  ( ; aIndex < mFields.Length(); ++aIndex)
	{
		FieldType &field = mFields[aIndex];
		if (aVal)
		{
			if (field.symbol == SYM_DYNAMIC)
			{
				// Skip it if it can't be called without parameters, or if there's no getter in this object
				// (consistent with inherited properties that have neither getter nor setter defined here).
				// Also skip if this is a class prototype, since that isn't an instance of the class and
				// therefore isn't a valid target for a method/property call.
				if (field.prop->MaxParams > 0 || !field.prop->Getter() || IsClassPrototype())
					continue;

				FuncResult result_token;
				ExprTokenType getter(field.prop->Getter());
				ExprTokenType object(this);
				auto *param = &object;
				auto result = getter.object->Invoke(result_token, IT_CALL, nullptr, getter, &param, 1);
				if (result == FAIL || result == EARLY_EXIT)
					return result;
				if (result_token.mem_to_free)
				{
					ASSERT(result_token.symbol == SYM_STRING && result_token.mem_to_free == result_token.marker);
					aVal->AcceptNewMem(result_token.mem_to_free, result_token.marker_length);
				}
				else
				{
					aVal->Assign(result_token);
					result_token.Free();
				}
			}
			else
			{
				ExprTokenType value;
				field.ToToken(value);
				aVal->Assign(value);
			}
		}
		if (aName)
		{
			aName->Assign(field.name);
		}
		return CONDITION_TRUE;
	}
	return CONDITION_FALSE;
}


ResultType Map::GetEnumItem(UINT &aIndex, Var *aKey, Var *aVal, int aVarCount)
{
	if (aIndex < mCount)
	{
		auto &item = mItem[aIndex];
		if (aKey)
		{
			if (aIndex < mKeyOffsetObject) // mKeyOffsetInt < mKeyOffsetObject
				aKey->Assign(item.key.i);
			else if (aIndex < mKeyOffsetString) // mKeyOffsetObject < mKeyOffsetString
				aKey->Assign(item.key.p);
			else // mKeyOffsetString < mCount
				aKey->Assign(item.key.s);
		}
		if (aVal)
		{
			ExprTokenType value;
			item.ToToken(value);
			aVal->Assign(value);
		}
		return CONDITION_TRUE;
	}
	return CONDITION_FALSE;
}


ResultType RegExMatchObject::GetEnumItem(UINT &aIndex, Var *aKey, Var *aVal, int aVarCount)
{
	if (aIndex >= (UINT)mPatternCount)
		return CONDITION_FALSE;
	// In single-var mode, return the subpattern values.
	// Otherwise, return the subpattern names first and values second.
	if (aVarCount < 2)
	{
		aVal = aKey;
		aKey = nullptr;
	}
	if (aKey)
	{
		if (mPatternName && mPatternName[aIndex])
			aKey->Assign(mPatternName[aIndex]);
		else
			aKey->Assign((__int64)aIndex);
	}
	if (aVal)
	{
		aVal->Assign(mHaystack - mHaystackStart + mOffset[aIndex*2], mOffset[aIndex*2+1]);
	}
	return CONDITION_TRUE;
}



//
// Object:: and Map:: Internal Methods
//

Map::Pair *Map::FindItem(IntKeyType val, index_t left, index_t right, index_t &insert_pos)
// left and right must be set by caller to the appropriate bounds within mItem.
{
	while (left < right)
	{
		index_t mid = left + ((right - left) >> 1);
		auto &item = mItem[mid];

		auto result = val - item.key.i;

		if (result < 0)
			right = mid;
		else if (result > 0)
			left = mid + 1;
		else
			return &item;
	}
	insert_pos = left;
	return nullptr;
}

Object::FieldType *Object::FindField(name_t name, index_t &insert_pos)
{
	index_t left = 0, mid, right = mFields.Length();
	int first_char = *name;
	if (first_char <= 'Z' && first_char >= 'A')
		first_char += 32;
	while (left < right)
	{
		mid = left + ((right - left) >> 1);
		
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
			right = mid;
		else if (result > 0)
			left = mid + 1;
		else
			return &field;
	}
	insert_pos = left;
	return nullptr;
}

bool Object::HasProp(name_t name)
{
	return FindField(name) || mBase && mBase->HasProp(name);
}

IObject *Object::GetMethod(name_t name)
{
	// Return the function(?) object which would be called if the named property is called,
	// or nullptr if that would require invoking a getter.  Does not verify that the object
	// is callable, and does not support primitive values (even in the unusual case that Call
	// has been implemented via the value's base/prototype).
	bool dynamic_only = false;
	for (Object *that = this; that; that = that->mBase)
	{
		if (auto field = that->FindField(name))
		{
			if (field->symbol != SYM_DYNAMIC)
				return (dynamic_only || field->symbol != SYM_OBJECT) ? nullptr : field->object;
			if (auto func = field->prop->Method())
				return func; // Method takes precedence over any inherited value or getter.
			if (field->prop->Getter())
				dynamic_only = true; // Getter takes precedence over any inherited value.
		}
	}
	return nullptr;
}

bool Object::HasMethod(name_t aName)
{
	return GetMethod(aName) != nullptr;
}

Map::Pair *Map::FindItem(LPTSTR val, index_t left, index_t right, index_t &insert_pos)
// left and right must be set by caller to the appropriate bounds within mItem.
{
	bool caseless = mFlags & MapCaseless;
	bool use_locale = mFlags & MapUseLocale;
	index_t mid;
	int first_char = caseless ? 0 : *val;
	while (left < right)
	{
		mid = left + ((right - left) >> 1);

		auto &item = mItem[mid];

		// key_c contains key.s[0], cached there for performance if !caseless.
		// If caseless, key_c is 0 since this simple formula is insufficient to
		// replicate the sort order of _tcsicmp and lstrcmpi.
		int result = first_char - item.key_c;
		if (!result)
			result = !caseless ? _tcscmp(val, item.key.s)
				: use_locale ? lstrcmpi(val, item.key.s) : _tcsicmp(val, item.key.s);

		if (result < 0)
			right = mid;
		else if (result > 0)
			left = mid + 1;
		else
			return &item;
	}
	insert_pos = left;
	return nullptr;
}

Map::Pair *Map::FindItem(SymbolType key_type, Key key, index_t &insert_pos)
// Searches for an item with the given key.  If found, a pointer to the item is returned.  Otherwise
// NULL is returned and insert_pos is set to the index a newly created item should be inserted at.
// key_type and key are output for creating a new item or removing an existing one correctly.
// left and right must indicate the appropriate section of mItem to search, based on key type.
{
	index_t left, right;

	switch (key_type)
	{
	case SYM_STRING:
		left = mKeyOffsetString;
		right = mCount; // String keys are last in the mItem array.
		return FindItem(key.s, left, right, insert_pos);
	case SYM_OBJECT:
		left = mKeyOffsetObject;
		right = mKeyOffsetString; // Object keys end where String keys begin.
		// left and right restrict the search to just the portion with object keys.
		// Reuse the integer search function to reduce code size.  On 32-bit builds,
		// this requires that the upper 32 bits of each key have been initialized.
		return FindItem((IntKeyType)(INT_PTR)key.p, left, right, insert_pos);
	//case SYM_INTEGER:
	default:
		left = mKeyOffsetInt;
		right = mKeyOffsetObject; // Int keys end where Object keys begin.
		return FindItem(key.i, left, right, insert_pos);
	}
}

void Map::ConvertKey(ExprTokenType &key_token, LPTSTR buf, SymbolType &key_type, Key &key)
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
		// Set i to support the way FindItem() is used.  Otherwise on 32-bit builds the
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

Map::Pair *Map::FindItem(ExprTokenType &key_token, LPTSTR aBuf, SymbolType &key_type, Key &key, index_t &insert_pos)
// Searches for an item with the given key, where the key is a token passed from script.
{
	ConvertKey(key_token, aBuf, key_type, key);
	return FindItem(key_type, key, insert_pos);
}
	
bool Object::SetInternalCapacity(index_t new_capacity)
// Expands mFields to the specified number if fields.
// Caller *must* ensure new_capacity >= 1 && new_capacity >= mFields.Length().
{
	return mFields.SetCapacity(new_capacity);
}

bool Map::SetInternalCapacity(index_t new_capacity)
// Caller *must* ensure new_capacity >= 1 && new_capacity >= mCount.
{
	Pair *new_fields = (Pair *)realloc(mItem, new_capacity * sizeof(Pair));
	if (!new_fields)
		return false;
	mItem = new_fields;
	mCapacity = new_capacity;
	return true;
}
	
Object::FieldType *Object::Insert(name_t name, index_t at)
// Inserts a single field with the given key at the given offset.
// Caller must ensure 'at' is the correct offset for this key.
{
	if (mFields.Length() == mFields.Capacity() && !Expand()  // Attempt to expand if at capacity.
		|| !(name = _tcsdup(name)))  // Attempt to duplicate key-string.
	{	// Out of memory.
		return nullptr;
	}
	// There is now definitely room in mFields for a new field.
	FieldType &field = *mFields.InsertUninitialized(at, 1);
	field.key_c = ctolower(*name);
	field.name = name; // Above has already copied string or called key.p->AddRef() as appropriate.
	field.Minit(); // Initialize to default value.  Caller will likely reassign.
	return &field;
}

Map::Pair *Map::Insert(SymbolType key_type, Key key, index_t at)
// Inserts a single item with the given key at the given offset.
// Caller must ensure 'at' is the correct offset for this key.
{
	if (mCount == mCapacity && !Expand()  // Attempt to expand if at capacity.
		|| key_type == SYM_STRING && !(key.s = _tcsdup(key.s)))  // Attempt to duplicate key-string.
	{	// Out of memory.
		return NULL;
	}
	// There is now definitely room in mItem for a new item.

	auto &item = mItem[at];
	if (at < mCount)
		// Move existing items to make room.
		memmove(&item + 1, &item, (mCount - at) * sizeof(Pair));
	++mCount; // Only after memmove above.

	// Update key-type offsets based on where and what was inserted; also update this key's ref count:
	if (key_type == SYM_STRING)
	{
		item.key_c = (mFlags & MapCaseless) ? 0 : *key.s;
	}
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

	item.key = key; // Above has already copied string or called key.p->AddRef() as appropriate.
	item.Minit(); // Initialize to default value.  Caller will likely reassign.

	return &item;
}



//
// Func: A function, either built-in or created by a function definition.
//

ResultType Func::Invoke(IObject_Invoke_PARAMS_DECL)
{
	if (!aName && IS_INVOKE_CALL && !HasOwnProps()) // Very rough check that covers the most common cases.
	{
		// Take a shortcut for performance.
		Call(aResultToken, aParam, aParamCount);
		return aResultToken.Result();
	}
	return Object::Invoke(IObject_Invoke_PARAMS);
}

void Func::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (MemberID(aID))
	{
	case M_Call:
		Call(aResultToken, aParam, aParamCount);
		return;

	case M_Bind:
		if (BoundFunc *bf = BoundFunc::Bind(this, IT_CALL, nullptr, aParam, aParamCount))
			_o_return(bf);
		_o_throw_oom;

	case M_IsOptional:
		if (aParamCount)
		{
			int param = ParamIndexToInt(0);
			if (param > 0 && (param <= mParamCount || mIsVariadic))
				_o_return(ArgIsOptional(param-1));
			else
				_o_throw_param(0);
		}
		else
			_o_return(mMinParams != mParamCount || mIsVariadic); // True if any params are optional.
	
	case M_IsByRef:
		if (aParamCount)
		{
			int param = ParamIndexToInt(0);
			if (param <= 0 || param > mParamCount && !mIsVariadic)
				_o_throw_param(0);
			_o_return(ArgIsOutputVar(param-1));
		}
		else
		{
			for (int param = 0; param < mParamCount; ++param)
				if (ArgIsOutputVar(param))
					_o_return(TRUE);
			_o_return(FALSE);
		}

	case P_Name: _o_return(const_cast<LPTSTR>(mName));
	case P_MinParams: _o_return(mMinParams);
	case P_MaxParams: _o_return(mParamCount);
	case P_IsBuiltIn: _o_return(IsBuiltIn());
	case P_IsVariadic: _o_return(mIsVariadic);
	}
}


bool BoundFunc::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	// Combine the bound parameters with the supplied parameters.
	int bound_count = mParams->Length();
	if (bound_count > 0)
	{
		ExprTokenType *token = (ExprTokenType *)_alloca(bound_count * sizeof(ExprTokenType));
		ExprTokenType **param = (ExprTokenType **)_alloca((bound_count + aParamCount) * sizeof(ExprTokenType *));
		mParams->ToParams(token, param, NULL, 0);
		// Fill in any missing parameters with those that were supplied.
		// Provides greater utility than binding to the parameter's default value.
		for (int i = 0; i < bound_count && aParamCount; ++i)
		{
			if (param[i]->symbol == SYM_MISSING)
			{
				param[i] = *(aParam++);
				--aParamCount;
			}
		}
		memcpy(param + bound_count, aParam, aParamCount * sizeof(ExprTokenType *));
		aParam = param;
		aParamCount += bound_count;
	}

	ExprTokenType this_token;
	this_token.symbol = SYM_OBJECT;
	this_token.object = mFunc;

	// Call the function or object.
	switch (mFunc->Invoke(aResultToken, mFlags, mMember, this_token, aParam, aParamCount))
	{
	case FAIL:
		return FAIL;
	default:
		return OK;
	case INVOKE_NOT_HANDLED:
		return aResultToken.UnknownMemberError(this_token, IT_CALL, mMember);
	}
}

BoundFunc *BoundFunc::Bind(IObject *aFunc, int aFlags, LPCTSTR aMember, ExprTokenType **aParam, int aParamCount)
{
	LPTSTR member;
	if (!aMember)
		member = nullptr;
	else if (!(member = _tcsdup(aMember)))
		return nullptr;
	if (auto params = Array::Create(aParam, aParamCount))
	{
		aFunc->AddRef();
		// BoundFunc takes our reference to params.
		return new BoundFunc(aFunc, member, params, aFlags);
	}
	free(member);
	return nullptr;
}

BoundFunc::~BoundFunc()
{
	mFunc->Release();
	mParams->Release();
	free(mMember);
}


bool Closure::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	AddRef();	// Avoid it being deleted during the call.
	auto result = mFunc->Call(aResultToken, aParam, aParamCount, mVars);
	Release();
	return result;
}

Closure::~Closure()
{
	if (!(mFlags & ClosureGroupedFlag))
		mVars->Release();
}

bool Closure::Delete()
{
	if ((mFlags & ClosureGroupedFlag) && !mVars->FullyReleased(mRefCount))
		return false;
	return Func::Delete();
}

bool FreeVars::FullyReleased(ULONG aRefPendingRelease)
{
	// This function is part of a workaround for circular references that occur because all closures
	// have a reference to this FreeVars, while any closure referenced by itself or another closure
	// has a reference in mVar[].
	if (mRefCount)
		return mRefCount < 0;
	int circular_closures = 0;
	for (int i = 0; i < mVarCount; ++i)
		if (mVar[i].Type() == VAR_CONSTANT)
		{
			ASSERT(mVar[i].HasObject() && dynamic_cast<ObjectBase*>(mVar[i].Object())); // Any object in VAR_CONSTANT must derive from ObjectBase.
			auto obj = (ObjectBase *)mVar[i].Object();
			if (obj->RefCount() && aRefPendingRelease == 0)
				return false;
			--aRefPendingRelease;
			++circular_closures;
		}
	--mRefCount; // Now that delete is certain, make this non-zero to prevent reentry.
	if (circular_closures)
	{
		// All closures in downvars have mRefCount == 0, meaning their only reference is the
		// uncounted one in mVar[].  In order to free the object properly, mRefCount needs to
		// be restored to 1 prior to Release(), which will be called by Var::Free().
		for (int i = 0; i < mVarCount; ++i)
			if (mVar[i].Type() == VAR_CONSTANT)
			{
				auto obj = (ObjectBase *)mVar[i].Object();
				obj->AddRef();
			}
	}
	delete this;
	return true;
}


ResultType IObjectPtr::ExecuteInNewThread(TCHAR *aNewThreadDesc, ExprTokenType *aParamValue, int aParamCount, bool aReturnBoolean) const
{
	DEBUGGER_STACK_PUSH(aNewThreadDesc)
	ResultType result = CallMethod(mObject, mObject, nullptr, aParamValue, aParamCount, nullptr, 0, aReturnBoolean);
	DEBUGGER_STACK_POP()
	return result;
}


Func *IObjectPtr::ToFunc() const
{
	return dynamic_cast<Func *>(mObject);
}

LPCTSTR IObjectPtr::Name() const
{
	if (auto func = ToFunc()) return func->mName;
	return mObject->Type();
}



ResultType MsgMonitorList::Call(ExprTokenType *aParamValue, int aParamCount, int aInitNewThreadIndex, __int64 *aRetVal)
{
	ResultType result = OK;
	__int64 retval = 0;
	
	for (MsgMonitorInstance inst (*this); inst.index < inst.count; ++inst.index)
	{
		if (inst.index >= aInitNewThreadIndex) // Re-initialize the thread.
			InitNewThread(0, true, false);
		
		IObject *func = mMonitor[inst.index].func;

		if (!CallMethod(func, func, nullptr, aParamValue, aParamCount, &retval))
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
	if (aRetVal)
		*aRetVal = retval;
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
		LPTSTR method_name = mon.is_method ? mon.method_name : nullptr;

		if (thread_used) // Re-initialize the thread.
			InitNewThread(0, true, false);
		
		// Set last found window (as documented).
		g->hWndLastUsed = aGui->mHwnd;
		
		// If we're about to call a method of the Gui itself, don't pass the Gui as the first parameter
		// since it will be in `this`.  Doing this here rather than when the parameters are built ensures
		// that message monitor functions (not methods) still receive the expected Gui parameter.
		int skip_arg = func == aGui && mon.is_method && aParamValue->symbol == SYM_OBJECT && aParamValue->object == aGui;

		result = CallMethod(func, func, method_name, aParamValue + skip_arg, aParamCount - skip_arg, &retval);
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
// Buffer
//

BufferObject *BufferObject::Create(void *aData, size_t aSize)
{
	auto obj = new BufferObject(aData, aSize);
	obj->SetBase(BufferObject::sPrototype);
	return obj;
}

ObjectMember BufferObject::sMembers[] =
{
	Object_Method(__New, 0, 2),
	Object_Property_get(Ptr),
	Object_Property_get_set(Size)
};

void BufferObject::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case P_Ptr:
		_o_return((size_t)mData);
	case P_Size: // Size or __New
		if (!IS_INVOKE_GET)
		{
			if (!ParamIndexIsOmitted(0))
			{
				if (!ParamIndexIsNumeric(0))
					if (IS_INVOKE_SET)
						_o_throw_type(_T("Number"), *aParam[0]);
					else
						_o_throw_param(0, _T("Number"));
				auto new_size = ParamIndexToInt64(0);
				if (new_size < 0 || new_size > SIZE_MAX)
					_o_throw_value(ERR_INVALID_VALUE);
				if (!Resize((size_t)new_size))
					_o_throw_oom;
			}
			if (!ParamIndexIsOmitted(1))
			{
				if (!ParamIndexIsNumeric(1))
					_o_throw_param(1, _T("Number"));
				memset(mData, (char)ParamIndexToInt64(1), mSize);
			}
			return;
		}
		_o_return(mSize);
	}
}

ResultType BufferObject::Resize(size_t aNewSize)
{
	// It seems worthwhile to guarantee that no reallocation is performed if size is the same.
	// Testing (in 2022) showed realloc() to return new blocks even if the same size is passed
	// in multiple times.
	if (aNewSize == mSize)
		return OK;
	auto new_data = realloc(mData, aNewSize);
	if (!new_data && aNewSize)
		return FAIL;
	mData = new_data;
	mSize = aNewSize;
	return OK;
}


void ClipboardAll::__New(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	void *data;
	size_t size;
	if (!aParamCount)
	{
		// Retrieve clipboard contents.
		if (!Var::GetClipboardAll(&data, &size))
			_o_return_FAIL;
	}
	else
	{
		// Use caller-supplied data.
		size_t caller_data;
		if (auto obj = ParamIndexToObject(0))
		{
			GetBufferObjectPtr(aResultToken, obj, caller_data, size);
			if (aResultToken.Exited())
				return;
		}
		else
		{
			// Caller supplied an address.
			Throw_if_Param_NaN(0);
			caller_data = (size_t)ParamIndexToIntPtr(0);
			if (caller_data < 65536) // Basic check to catch incoming raw addresses that are zero or blank.  On Win32, the first 64KB of address space is always invalid.
				_o_throw_param(0);
			size = -1;
		}
		if (!ParamIndexIsOmitted(1))
		{
			Throw_if_Param_NaN(1);
			size = (size_t)ParamIndexToIntPtr(1);
		}
		else if (size == -1) // i.e. it can be omitted when size != -1 (a string was passed).
			return (void)aResultToken.ParamError(1, nullptr);
		if (  !(data = malloc(size))  ) // More likely to be due to invalid parameter than out of memory.
			_o_throw_oom;
		memcpy(data, (void *)caller_data, size);
	}
	if (mData != data)
		free(mData); // In case of explicit call to __New.
	mData = data;
	mSize = size;
}


Object *ClipboardAll::Create()
{
	auto obj = new ClipboardAll();
	obj->SetBase(ClipboardAll::sPrototype);
	return obj;
}


ObjectMember ClipboardAll::sMembers[]
{
	Object_Method1(__New, 0, 2)
};



ObjectMember Func::sMembers[] =
{
	Object_Method(Bind, 0, MAXP_VARIADIC),
	Object_Method(Call, 0, MAXP_VARIADIC),
	Object_Method(IsByRef, 0, MAX_FUNCTION_PARAMS),
	Object_Method(IsOptional, 0, MAX_FUNCTION_PARAMS),

	Object_Property_get(IsBuiltIn),
	Object_Property_get(IsVariadic),
	Object_Property_get(MaxParams),
	Object_Property_get(MinParams),
	Object_Property_get(Name)
};



ObjectMember RegExMatchObject::sMembers[] =
{
	Object_Method(__Enum, 0, 1),
	Object_Method(__Get, 2, 2),
	Object_Member(__Item, Invoke, M_Value, IT_GET, 0, 1),
	Object_Member(Count, Invoke, M_Count, IT_GET, 0, 0),
	Object_Method(Len, 0, 1),
	Object_Member(Len, Invoke, M_Len, IT_GET, 0, 1),
	Object_Member(Mark, Invoke, M_Mark, IT_GET, 0, 0),
	Object_Method(Name, 1, 1),
	Object_Member(Name, Invoke, M_Name, IT_GET, 1, 1),
	Object_Method(Pos, 0, 1),
	Object_Member(Pos, Invoke, M_Pos, IT_GET, 0, 1),
};



ObjectMember Object::sErrorMembers[]
{
	Object_Member(__New, Error__New, M_Error__New, IT_CALL, 0, 3)
};

ObjectMember Object::sOSErrorMembers[]
{
	Object_Member(__New, Error__New, M_OSError__New, IT_CALL, 0, 3)
};



struct ClassDef
{
	LPCTSTR name;
	Object **proto_var;
	ClassFactoryDef factory;
	ObjectMemberListType members;
	int member_count;
	std::initializer_list<ClassDef> subclasses;
};

void DefineClasses(Object *aBaseClass, Object *aBaseProto, std::initializer_list<ClassDef> aClasses)
{
	for (auto &c : aClasses)
	{
		auto proto = (c.proto_var && *c.proto_var) ? *c.proto_var
			: Object::CreatePrototype(const_cast<LPTSTR>(c.name), aBaseProto, c.members, c.member_count);
		if (c.proto_var)
			*c.proto_var = proto;
		auto cobj = Object::CreateClass(const_cast<LPTSTR>(c.name), aBaseClass, proto, c.factory);
		if (c.subclasses.size())
			DefineClasses(cobj, proto, c.subclasses);
	}
}


void Object::CreateRootPrototypes()
{
	// Create the root prototypes before defining any members, since
	// each member relies on Func::sPrototype having been initialized.
	sAnyPrototype = CreatePrototype(_T("Any"), nullptr);
	sPrototype = CreatePrototype(_T("Object"), sAnyPrototype);
	Func::sPrototype = CreatePrototype(_T("Func"), Object::sPrototype);

	// These methods correspond to global functions, as BuiltInMethod
	// only handles Objects, and these must handle primitive values.
	static const LPTSTR sFuncs[] = { _T("GetMethod"), _T("HasBase"), _T("HasMethod"), _T("HasProp") };
	for (int i = 0; i < _countof(sFuncs); ++i)
		sAnyPrototype->DefineMethod(sFuncs[i], g_script.FindGlobalFunc(sFuncs[i]));
	auto prop = sAnyPrototype->DefineProperty(_T("Base"));
	prop->MinParams = 0;
	prop->MaxParams = 0;
	prop->SetGetter(g_script.FindGlobalFunc(_T("ObjGetBase")));
	prop->SetSetter(g_script.FindGlobalFunc(_T("ObjSetBase")));
	
	// Define __Init so that Script::DefineClassInit can add an unconditional super.__Init().
	static auto __Init = new BuiltInFunc { _T(""), Any___Init, 1, 1 };
	sAnyPrototype->DefineMethod(_T("__Init"), __Init);

	DefineMembers(sPrototype, _T("Object"), sMembers, _countof(sMembers));
	DefineMembers(Func::sPrototype, _T("Func"), Func::sMembers, _countof(Func::sMembers));

	// Create classes.
	//

	sClassPrototype = Object::CreatePrototype(_T("Class"), Object::sPrototype);
	auto anyClass = CreateClass(_T("Any"), sClassPrototype, sAnyPrototype, nullptr);
	Object::sClass = CreateClass(_T("Object"), anyClass, Object::sPrototype, NewObject<Object>);

	ObjectCtor no_ctor = nullptr;
	ObjectMember *no_members = nullptr;

	DefineClasses(Object::sClass, Object::sPrototype, {
		{_T("Array"), &Array::sPrototype, NewObject<Array>
			, Array::sMembers, _countof(Array::sMembers)},
		{_T("Buffer"), &BufferObject::sPrototype, NewObject<BufferObject>, BufferObject::sMembers, _countof(BufferObject::sMembers), {
			{_T("ClipboardAll"), &ClipboardAll::sPrototype, NewObject<ClipboardAll>
				, ClipboardAll::sMembers, _countof(ClipboardAll::sMembers)}
		}},
		{_T("Class"), &Object::sClassPrototype},
		{_T("Error"), &ErrorPrototype::Error, no_ctor, sErrorMembers, _countof(sErrorMembers), {
			{_T("MemoryError"), &ErrorPrototype::Memory},
			{_T("OSError"), &ErrorPrototype::OS, no_ctor, sOSErrorMembers, _countof(sOSErrorMembers)},
			{_T("TargetError"), &ErrorPrototype::Target},
			{_T("TimeoutError"), &ErrorPrototype::Timeout},
			{_T("TypeError"), &ErrorPrototype::Type},
			{_T("UnsetError"), &ErrorPrototype::Unset, no_ctor, no_members, 0, {
				{_T("MemberError"), &ErrorPrototype::Member, no_ctor, no_members, 0, {
					{_T("PropertyError"), &ErrorPrototype::Property},
					{_T("MethodError"), &ErrorPrototype::Method}
				}},
				{_T("UnsetItemError"), &ErrorPrototype::UnsetItem}
			}},
			{_T("ValueError"), &ErrorPrototype::Value, no_ctor, no_members, 0, {
				{_T("IndexError"), &ErrorPrototype::Index}
			}},
			{_T("ZeroDivisionError"), &ErrorPrototype::ZeroDivision}
		}},
		{_T("Func"), &Func::sPrototype, no_ctor, Func::sMembers, _countof(Func::sMembers), {
			{_T("BoundFunc"), &BoundFunc::sPrototype},
			{_T("Closure"), &Closure::sPrototype},
			{_T("Enumerator"), &EnumBase::sPrototype}
		}},
		{_T("Gui"), &GuiType::sPrototype, NewObject<GuiType>
			, GuiType::sMembers, GuiType::sMemberCount},
		{_T("InputHook"), &InputObject::sPrototype, NewObject<InputObject>
			, InputObject::sMembers, InputObject::sMemberCount},
		{_T("Map"), &Map::sPrototype, NewObject<Map>
			, Map::sMembers, _countof(Map::sMembers)},
		{_T("Menu"), &UserMenu::sPrototype, NewObject<UserMenu>
			, UserMenu::sMembers, UserMenu::sMemberCount, {
			{_T("MenuBar"), &UserMenu::sBarPrototype, NewObject<UserMenu::Bar>}
		}},
		{_T("RegExMatchInfo"), &RegExMatchObject::sPrototype, no_ctor
			, RegExMatchObject::sMembers, _countof(RegExMatchObject::sMembers)}
	});

	// Parameter counts are specified for static Call in the following classes
	// but not those using NewObject<> because the latter passes parameters on
	// to __New, which can be redefined by a subclass.  Specifying counts here
	// sets MinParams/MaxParams/IsVariadic appropriately and avoids the need to
	// validate aParamCount in each function, which reduces code size.
	// Note that the `this` parameter (the class itself) is counted.
	DefineClasses(anyClass, sAnyPrototype, {
		{_T("ComValue"), &sComValuePrototype, {ComValue_Call, 3, 4}, no_members, 0, {
			{_T("ComObjArray"), &sComArrayPrototype, {ComObjArray_Call, 3, 10}},
			{_T("ComObject"), &sComObjectPrototype, {ComObject_Call, 2, 3}},
			{_T("ComValueRef"), &sComRefPrototype}
		}},
		{_T("Primitive"), &Object::sPrimitivePrototype, no_ctor, no_members, 0, {
			{_T("Number"), &Object::sNumberPrototype, {BIF_Number, 2, 2}, no_members, 0, {
				{_T("Float"), &Object::sFloatPrototype, {BIF_Float, 2, 2}},
				{_T("Integer"), &Object::sIntegerPrototype, {BIF_Integer, 2, 2}}
			}},
			{_T("String"), &Object::sStringPrototype, {BIF_String, 2, 2}}
		}},
		{_T("VarRef"), &sVarRefPrototype}
	});

	GuiControlType::DefineControlClasses();
	DefineComPrototypeMembers();
	DefineFileClass();

	// Permit Object.Call to construct Error objects.
	ErrorPrototype::Error->mFlags &= ~NativeClassPrototype;
	ErrorPrototype::OS->mFlags &= ~NativeClassPrototype;
}

Object *Object::sAnyPrototype;
Object *Func::sPrototype;
Object *Object::sPrototype;

Object *Object::sClassPrototype;
Object *Array::sPrototype;
Object *Map::sPrototype;

Object *Object::sClass;

Object *Closure::sPrototype;
Object *BoundFunc::sPrototype;
Object *EnumBase::sPrototype;

Object *BufferObject::sPrototype;
Object *ClipboardAll::sPrototype;

Object *RegExMatchObject::sPrototype;

Object *GuiType::sPrototype;
Object *UserMenu::sPrototype;
Object *UserMenu::sBarPrototype;

namespace ErrorPrototype
{
	Object *Error, *Memory, *Type, *Value, *OS, *ZeroDivision;
	Object *Target, *Unset, *Member, *Property, *Method, *Index, *UnsetItem;
	Object *Timeout;
}

Object *Object::sVarRefPrototype;
Object *Object::sComObjectPrototype, *Object::sComValuePrototype, *Object::sComArrayPrototype, *Object::sComRefPrototype;


//
// Primitive values as objects
//

Object *Object::sPrimitivePrototype;
Object *Object::sStringPrototype;
Object *Object::sNumberPrototype;
Object *Object::sIntegerPrototype;
Object *Object::sFloatPrototype;

Object *Object::ValueBase(ExprTokenType &aValue)
{
	switch (TypeOfToken(aValue))
	{
	case SYM_STRING: return Object::sStringPrototype;
	case SYM_INTEGER: return Object::sIntegerPrototype;
	case SYM_FLOAT: return Object::sFloatPrototype;
	}
	return nullptr;
}



void Object::DefineClass(name_t aName, Object *aClass)
{
	auto prop = DefineProperty(aName);

	ExprTokenType values[] { aClass, aName }, *param[] { values, values + 1 };

	auto info = SimpleHeap::Alloc<NestedClassInfo>();
	info->class_object = aClass;
	info->constructed = false;
	aClass->AddRef();

	auto get = new BuiltInFunc { _T(""), Class_GetNestedClass, 1, 1, false, info };
	prop->MinParams = 0;
	prop->MaxParams = 0;
	prop->SetGetter(get);

	auto call = new BuiltInFunc { _T(""), Class_CallNestedClass, 1, 1, true, info };
	prop->SetMethod(call);
}


BIF_DECL(Class_GetNestedClass)
{
	auto info = (NestedClassInfo *)aResultToken.func->mData;
	auto cls = info->class_object;
	cls->AddRef();
	if (info->constructed)
		_f_return(cls);
	info->constructed = true;
	cls->Construct(aResultToken, nullptr, 0);
}


BIF_DECL(Class_CallNestedClass)
{
	auto info = (NestedClassInfo *)aResultToken.func->mData;
	auto cls = info->class_object;
	if (!info->constructed)
	{
		info->constructed = true;
		cls->AddRef(); // Necessary because Construct() calls Release() on failure/exit.
		if (cls->Construct(aResultToken, nullptr, 0) != OK) // FAIL or EXIT
			return;
		cls->Release();
		aResultToken.InitResult(aResultToken.buf);
	}
	else
		aResultToken.symbol = SYM_STRING; // Set the default expected by Invoke.
	cls->Invoke(aResultToken, IT_CALL, nullptr, ExprTokenType { cls }, aParam + 1, aParamCount - 1);
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

#endif
