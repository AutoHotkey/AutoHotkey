#include "stdafx.h" // pre-compiled headers
#include "defines.h"
#include "globaldata.h"
#include "script.h"
#include "application.h"

#include "script_object.h"
#include "script_func_impl.h"
#include "script_gui.h"
#include "input_object.h"

#include <errno.h> // For ERANGE.
#include <initializer_list>


#define STRUCT_PTR_CLASS_NAME _T("Ptr")
#define STRUCT_PTR_CLASS_SUFFIX _T(".") STRUCT_PTR_CLASS_NAME

// Used on functions called only during program startup where inlining just wastes space.
#define STARTUP_FUNCTION __declspec(noinline)


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
			result = TokenIsBlank(result_token) ? OK : EARLY_RETURN;
	}
	
	if (aRetVal) // Always set this as some callers don't initialize it:
		*aRetVal = result == EARLY_RETURN ? TokenToInt64(result_token) : 0;

	result_token.Free();
	return result;
}


//
// Object Construction
//

void *Object::operator new(size_t aObjectSize)
{
	return operator new(aObjectSize, 0);
}

void *Object::operator new(size_t aObjectSize, size_t aAdditional)
{
	if (auto p = malloc(aObjectSize + aAdditional))
		return p;
	g_script.CriticalError(ERR_OUTOFMEM);
	return nullptr; // Never executed (process terminated).
}

void Object::operator delete(void *p)
{
	free(p);
}

void Object::operator delete(void *p, size_t)
{
	free(p);
}

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

Object *Object::CreateStruct(Object *aBase, UINT_PTR aPtr, UINT aFlags)
{
	auto &si = *aBase->GetStructInfo(true);
	Object *obj = new (si.nested_object_size + si.size) Object(aFlags | CannotOwnProps | DataIsSuffix);
	void *data = obj + 1;
	if (si.nested_object_size)
	{
		ZeroMemory(data, si.nested_object_size);
		data = (char*)data + si.nested_object_size;
	}
	if (si.size)
	{
		if (aPtr)
			memcpy(data, (void*)aPtr, si.size);
		else
			ZeroMemory(data, si.size);
	}
	obj->SetBase(aBase);
	return obj;
}

Object *Object::CreateStructPtr(Object *aBase, UINT_PTR aPtr, UINT aFlags)
{
	auto &si = *aBase->GetStructInfo(true);
	Object *obj = new (si.nested_object_size + sizeof(void*)) Object(aFlags | CannotOwnProps | DataIsSuffixPtr);
	void *data = obj + 1;
	if (si.nested_object_size)
	{
		ZeroMemory(data, si.nested_object_size);
		data = (char*)data + si.nested_object_size;
	}
	*(UINT_PTR*)data = aPtr;
	obj->SetBase(aBase);
	return obj;
}

ResultType Object::CreateStruct(ResultToken &aResultToken, Object *aBase, ExprTokenType *aParam[], int aParamCount)
{
	auto obj = CreateStruct(aBase);
	return obj->Initialize(aResultToken, aParam, aParamCount);
}

void Object::NewInstance(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	auto nproto = (Object*)aResultToken.callee_id;
	auto nsi = (StructInfo*)(nproto + 1);
	IObject *cls = ParamIndexToObject(0);
	// For backward-compatibility, this must permit any Object (not just a Class)
	// with a Prototype own property which is any Object (not just a Prototype).
	Object *proto = cls && cls->IsOfType(Object::sPrototype) ? ((Object*)cls)->ClassGetPrototypeBackwardCompatible() : nullptr;
	auto si = proto ? proto->GetStructInfo(true) : nullptr;
	if (!si || si->create != nsi->create)
		_f_throw_value(ERR_INVALID_BASE);

	auto suffix = si->nested_object_size + si->size;
	auto obj = si->create(suffix);
	if (suffix)
	{
		auto ptr = (UINT_PTR)obj + si->object_size;
		ZeroMemory((void*)ptr, suffix);
		obj->mFlags |= DataIsSuffix;
	}
	obj->SetBase(proto);
	obj->Initialize(aResultToken, aParam + 1, aParamCount - 1);
}

template<class T>
T *NewObject(size_t aSuffixSize)
{
	return new (aSuffixSize) T();
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
	if (field_count && !obj.SetInternalCapacity(field_count))
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
			((VarRef *)aParam[i]->object)->UninitializeNonVirtual(VAR_NEVER_FREE);
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
	if (mOuter)
	{
		// The current object (inner) and mOuter refer to each other.  The circular dependency
		// is handled by counting inner's reference to outer only while there are external refs
		// to inner (mRefCount>0).  Delete is called when mRefCount==1 about to become 0, meaning
		// the last external reference is released, and inner must Release outer.
		// Outer's __delete may rely on the inner objects, and yet inner's __delete can't execute
		// safely if its DataPtr() points to deleted data.  So outer is always destructed first,
		// and it becomes responsible for recursively destructing inner.
		if (mRefCount)
		{
			mRefCount--; // To reflect that this object doesn't have a counted ref to outer during outer's __delete.
			if (mOuter->Release() == 0)
				return true; // this was deleted, so don't do mRefCount++.
			mRefCount++; // Caller will --mRefCount.
			return false;
		}
		mRefCount++; // Must be non-zero during __Delete.
	}

	// __Delete shouldn't be called for Prototype objects.  Although it would be more efficient to
	// exclusively use the flag, it has been documented that __Delete isn't called if __Class exists.
	if (!(mFlags & NoCallDelete) && !FindField(_T("__Class")))
	{
		// L33: Privatize the last recursion layer's deref buffer in case it is in use by our caller.
		// It's done here rather than in Var::FreeAndRestoreFunctionVars (even though the below might
		// not actually call any script functions) because this function is probably executed much
		// less often in most cases.
		PRIVATIZE_S_DEREF_BUF;

		// If an exception has been thrown, temporarily clear it for execution of __Delete.
		ResultToken *exc = g->ThrownToken;
		g->ThrownToken = NULL;
		
		// EXCPTMODE_DELETE is used to replace "The current thread will exit" in error messages
		// with something more accurate (since an error here can't cause the thread to exit).
		// EXCPTMODE_CATCH is temporarily removed to ensure the error is actually reported
		// (otherwise it would be ignored if an object is deleted within try-catch).
		int outer_excptmode = g->ExcptMode;
		g->ExcptMode = (g->ExcptMode & ~EXCPTMODE_CATCH) | EXCPTMODE_DELETE;

		CallMetaDelete();

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


void Object::CallMetaDelete()
{
	// Caller has prepared the thread for __Delete to be called directly.
	ASSERT(mRefCount == 1 && mBase && (g->ExcptMode & EXCPTMODE_DELETE));

	FuncResult rt;
	CallMeta(_T("__Delete"), rt, ExprTokenType(this), nullptr, 0);
	rt.Free();

	// Call all nested destructors before anything is deleted.
	auto si = mBase->GetStructInfo();
	if (si->item_count)
	{
		if (!si->pointed_class) // Primitive values.
			return;
		// Element type is inferred by overall nest size and count.
		size_t nested_size = si->nested_object_size / si->item_count;
		if (nested_size == sizeof(Object*))
			return; // Pointers are released by ~Object().
		ASSERT(nested_size >= sizeof(Object));
		char *nest = (char*)this + si->object_size + si->nested_object_size;
		for (size_t i = 0; i < si->item_count; ++i)
		{
			auto nested = (Object*)(nest -= nested_size); // Destruct right to left.
			ASSERT(*(UINT_PTR*)nested);
			++nested->mRefCount;
			nested->CallMetaDelete();
			--nested->mRefCount;
		}
	}
	else
	{
		for (auto tp = si->last_field; tp; tp = tp->prev_field) // prev_field list includes inherited fields.
			if (tp->object_offset && !tp->pointed_proto)
			{
				auto nested = (Object*)((char*)this + tp->object_offset);
				ASSERT(*(UINT_PTR*)nested);
				++nested->mRefCount;
				nested->CallMetaDelete();
				--nested->mRefCount;
			}
	}
}


Object::~Object()
{
	if (mFlags & ClassPrototype)
	{
		auto &si = *(StructInfo*)(this + 1);
		// Iterate first_field & next_field (this prototype's own definitions),
		// not last_field & prev_field (own and inherited definitions).
		for (TypedProperty *next, *tp = si.first_field; tp; tp = next)
		{
			next = tp->next_field;
			delete tp;
		}
		// The pointer class holds a counted reference to the pointed class only while
		// external references to the pointer class exist.  At this stage all external
		// references to both have been released and pointed_class has been deleted.
		if (si.pointer_class)
			si.pointer_class->Delete();
		//if (si.pointed_class)
		//	si.pointed_class->Release();
		if (si.array_class_map)
			si.array_class_map->Release();
	}
	else if (mFlags & (DataIsSuffix | DataIsSuffixPtr | ObjectIsClass))
	{
		// Call native destructor for each nested object and release any pointer held for a Ptr field.
		auto si = mBase->GetStructInfo();
		for (auto tp = si->last_field; tp; tp = tp->prev_field) // prev_field list includes inherited fields.
			if (tp->object_offset)
			{
				auto nest = (char*)this + tp->object_offset;
				if (tp->pointed_proto)
				{
					auto p = (Object**)nest;
					if (*p)
						(*p)->Release();
				}
				else
				{
					auto p = (Object*)nest;
					if (*(UINT_PTR*)p) // vftbl initialized
						p->~Object();
				}
			}
		if (si->pointed_class) // Struct.Array or Struct.Ptr class, or a Class.
		{
			// Element type is inferred by overall nest size and count.
			size_t count = max(si->item_count, 1);
			size_t nested_size = si->nested_object_size / count;
			char *nest = (char*)this + si->object_size + si->nested_object_size;
			if (nested_size == sizeof(Object*))
			{
				auto p = (Object**)nest;
				for (size_t i = 0; i < count; ++i)
					if (*--p)
						(*p)->Release();
			}
			else
			{
				ASSERT(nested_size >= sizeof(Object));
				for (size_t i = 0; i < count; ++i)
				{
					auto nested = (Object*)(nest -= nested_size); // Destruct right to left.
					nested->~Object();
				}
			}
		}
	}
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
	Object_Method1(__Ref, 1, 1),
	Object_Method1(Clone, 0, 0),
	Object_Method1(DefineProp, 2, 2),
	Object_Method1(DeleteProp, 1, 1),
	Object_Method1(GetOwnPropDesc, 1, 1),
	Object_Method1(HasOwnProp, 1, 1),
	Object_Method1(OwnProps, 0, 0),
};

LPTSTR Object::sMetaFuncName[] = { _T("__Get"), _T("__Set"), _T("__Call") };

ResultType Object::Invoke(IObject_Invoke_PARAMS_DECL)
{
	// In debug mode, verify aResultToken has been initialized correctly.
	ASSERT(aResultToken.symbol == SYM_MISSING);
	ASSERT(aResultToken.Result() == OK);

	name_t name;
	if (!aName)
	{
		name = IS_INVOKE_CALL ? _T("Call") : _T("__Item");
		aFlags |= IF_BYPASS_METAFUNC;
	}
	else
		name = aName;
	
	ResultType result;

	switch (INVOKE_TYPE)
	{
	case IT_GET: result = GetProperty(aResultToken, aFlags, name, aThisToken, aParam, aParamCount); break;
	case IT_SET: result = SetProperty(aResultToken, aFlags, name, aThisToken, aParam, aParamCount); break;
	default: result = CallProperty(aResultToken, aFlags, name, aThisToken, aParam, aParamCount); break;
	}

	if (result == INVOKE_NOT_HANDLED && !(aFlags & IF_BYPASS_METAFUNC))
		result = CallMetaVarg(aFlags, aName, aResultToken, aThisToken, aParam, aParamCount);

	return result;
}


ResultType Object::GetProperty(ResultToken &aResultToken, int aFlags, name_t aName, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount)
{
	IObject *method = nullptr;

	for (auto that = this; that; that = that->mBase)
	{
		auto field = that->FindField(aName);
		if (!field)
			continue;

		if (field->symbol != SYM_DYNAMIC || field->prop->NoParamGet)
		{
			int flag = aParamCount ? IF_BYPASS___VALUE : 0;
			auto result = GetFieldValue(aResultToken, aFlags | flag, *field, aThisToken);
			if (!aParamCount || result != OK)
				return result;
			return ApplyParams(aResultToken, aFlags, aParam, aParamCount);
		}
		else if (auto getter = field->prop->Getter())
		{
			return CallEtter(aResultToken, aFlags, getter, aThisToken, aParam, aParamCount);
		}
		else if (!method)
		{
			method = field->prop->Method();
		}
	}

	if (method)
	{
		method->AddRef();
		aResultToken.SetValue(method);
		return OK;
	}

	return INVOKE_NOT_HANDLED;
}


ResultType Object::GetFieldValue(ResultToken &aResultToken, int aFlags, FieldType &aField, ExprTokenType &aThisToken)
{
	if (aField.symbol == SYM_TYPED_FIELD)
	{
		auto that = GetThisForTypedValue(aResultToken, aFlags, aField.name, aThisToken);
		return that ? that->GetTypedValue(aResultToken, aFlags, *aField.tprop) : FAIL;
	}
	else if (aField.symbol == SYM_DYNAMIC)
	{
		if (aField.prop->Getter())
			return CallEtter(aResultToken, aFlags, aField.prop->Getter(), aThisToken, nullptr, 0);
		auto method = aField.prop->Method();
		method->AddRef();
		aResultToken.SetValue(method);
		return OK;
	}
	aField.ReturnRef(aResultToken);
	return OK;
}


ResultType Object::CallProperty(ResultToken &aResultToken, int aFlags, name_t aName, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount)
{
	ResultToken method_token;
	method_token.InitResult(aResultToken.buf);
	auto result = GetMethodValue(method_token, aFlags, aName, aThisToken);
	if (result != INVOKE_NOT_HANDLED)
	{
		if (result == OK)
			result = CallAsMethod(method_token, aResultToken, aThisToken, aParam, aParamCount);
		method_token.Free();
	}
	return result;
}


ResultType Object::GetMethodValue(ResultToken &aResultToken, int aFlags, name_t aName, ExprTokenType &aThisToken)
{
	FieldType *getter = nullptr;
	for (auto that = this; that; that = that->mBase)
	{
		auto field = that->FindField(aName);
		if (!field)
			continue;
		if (field->symbol != SYM_DYNAMIC)
		{
			getter = field;
			break;
		}
		if (auto method = field->prop->Method())
		{
			method->AddRef();
			aResultToken.SetValue(method);
			return OK;
		}
		if (!getter && field->prop->Getter())
			getter = field;
	}
	if (getter)
		return GetFieldValue(aResultToken, (aFlags & ~IT_BITMASK) | IF_BYPASS___VALUE, *getter, aThisToken);
	return INVOKE_NOT_HANDLED;
}


ResultType Object::SetProperty(ResultToken &aResultToken, int aFlags, name_t aName, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount)
{
	Object *that;
	index_t insert_pos, other_pos;
	FieldType *field = nullptr;

	for (that = this; that; that = that->mBase)
	{
		auto candidate = that->FindField(aName, that == this ? insert_pos : other_pos);
		if (!candidate)
			continue;
		if (candidate->symbol != SYM_DYNAMIC)
		{
			// This value property takes precedence over anything inherited from that->mBase,
			// but any previously found getter or method implies that this property is read-only.
			if (!field)
				field = candidate;
			break;
		}
		auto setter = candidate->prop->Setter();
		if (setter && !(aParamCount > 1 && candidate->prop->NoParamSet))
		{
			// Setter hasn't been shadowed by a value/typed property, and either takes parameters
			// or none were passed.  If it takes parameters, the search stops here even if there
			// are insufficient parameters to successfully call the setter.
			return CallEtter(aResultToken, aFlags, setter, aThisToken, aParam, aParamCount);
		}
		// Save the first getter in case no setter is found, or the first method if no getters.
		if (  !(field && field->prop->Getter()) && candidate->prop->Getter()
			|| !field && candidate->prop->Method()  )
			field = candidate;
	}

	if (!field && !(aFlags & IF_BYPASS_METAFUNC))
	{
		// Call __Set before creating a field.
		auto result = CallMetaVarg(aFlags, aName, aResultToken, aThisToken, aParam, aParamCount);
		if (result != INVOKE_NOT_HANDLED)
			return result;
 	}

	if (aParamCount > 1)
	{
		if (!field)
			return INVOKE_NOT_HANDLED;

		// Apply parameters to the property's value, since there is no setter which accepts parameters.
		auto result = GetFieldValue(aResultToken, (aFlags & ~IT_BITMASK) | IF_BYPASS___VALUE, *field, aThisToken);
		if (result != OK)
			return result;
		return ApplyParams(aResultToken, aFlags, aParam, aParamCount);
	}

	if (field && field->symbol == SYM_TYPED_FIELD)
	{
		auto that = GetThisForTypedValue(aResultToken, aFlags, aName, aThisToken);
		return that ? that->SetTypedValue(aResultToken, aFlags, aName, *field->tprop, **aParam) : FAIL;
	}
	
	if (field && field->symbol == SYM_DYNAMIC || (aFlags & (IF_SUBSTITUTE_THIS | IF_SUPER)))
	{
		if ((aFlags & IF_SUPER) && !(field && field->symbol == SYM_DYNAMIC))
		{
			// This is `super.x := y` where x is either a value property or undefined.
			// If aThisToken is an Object, use the base Object implementation of set: create a value property.
			if (auto real_this = dynamic_cast<Object *>(TokenToObject(aThisToken)))
			{
				if (!real_this->SetOwnProp(aName, **aParam))
					return aResultToken.MemoryError();
				return OK;
			}
		}
		// This property has a getter but no setter; or either IF_SUBSTITUTE_THIS or IF_SUPER was set and above
		// did not return, in which case the property should be considered read-only, since it can't be stored
		// in the actual target object (which is aThisToken, not C++ `this`).
		return field ? aResultToken.Error(ERR_PROPERTY_READONLY, aName) : INVOKE_NOT_HANDLED;
	}

	if (this != that)
	{
		if ((aFlags & IF_NO_NEW_PROPS) || (mFlags & CannotOwnProps))
			return INVOKE_NOT_HANDLED;
		if (aParam[0]->symbol == SYM_MISSING)
			return OK; // No action needed for x.y := unset.
		if (  !(field = Insert(aName, insert_pos))  )
			return aResultToken.MemoryError();
	}
	else if (aParam[0]->symbol == SYM_MISSING) // x.y := unset
	{
		// Completely delete the property, since other sections currently aren't designed to handle properties
		// with no value (unlike Array and Map items).
		mFields.Remove((index_t)(field - mFields), 1);
		return OK;
	}

	return field->Assign(**aParam) ? OK : aResultToken.MemoryError();
}


ResultType Object::CallEtter(ResultToken &aResultToken, int aFlags, IObject *aEtter, ExprTokenType &aThisToken, ExprTokenType *aParam[], int aParamCount)
{
	// Prepare the parameter list: this, [value,] actual_param*
	ExprTokenType this_etter(aEtter);
	auto prop_param = (ExprTokenType **)_malloca((aParamCount + 1) * sizeof(ExprTokenType *));
	if (!prop_param)
		return aResultToken.MemoryError();
	prop_param[0] = &aThisToken; // For the hidden "this" parameter in the getter/setter.
	int prop_param_count = 1;
	if (IS_INVOKE_SET)
		// Put the setter's hidden "value" parameter before the other parameters.
		prop_param[prop_param_count++] = aParam[--aParamCount];
	memcpy(prop_param + prop_param_count, aParam, aParamCount * sizeof(ExprTokenType *));
	prop_param_count += aParamCount;
	// Call getter/setter.
	auto result = aEtter->Invoke(aResultToken, IT_CALL, nullptr, this_etter, prop_param, prop_param_count);
	_freea(prop_param);
	if (result == INVOKE_NOT_HANDLED)
		return aResultToken.UnknownMemberError(this_etter, IT_CALL, nullptr);
	return result;
}


Object *Object::GetThisForTypedValue(ResultToken &aResultToken, int aFlags, name_t aName, ExprTokenType &aThisToken)
{
	auto realthis = this;
	if (aFlags & (IF_SUBSTITUTE_THIS | IF_SUPER))
		realthis = dynamic_cast<Object*>(TokenToObject(aThisToken));
	if (realthis && realthis->HasData())
		return realthis;
	aResultToken.Error(_T("Property invalid for object with null data."), aName);
	return nullptr;
}


ResultType Object::GetTypedValue(ResultToken &aResultToken, int aFlags, TypedProperty &aProp)
{
	auto ptr = DataPtr() + aProp.data_offset;
	if (aProp.class_object) // Struct type.
	{
		if (aProp.pointed_proto) // Pointer type.
		{
			return GetBoxedPointer(aResultToken, *(UINT_PTR*)ptr, aProp.pointed_proto, aProp.object_offset);
		}
		auto nested = (Object*)((char*)this + aProp.object_offset);
		if (*(UINT_PTR*)nested == 0) // Since it wasn't constructed, this must be a pointer, not a real struct.
		{
			auto result = NestedSparseInit(aResultToken, aProp, ptr);
			if (result != OK)
				return result;
		}
		if (++nested->mRefCount == 1) // First external reference.
			++mRefCount; // Keep this alive while nested is referenced externally.
		if (!(aFlags & IF_BYPASS___VALUE))
		{
			auto result = nested->Invoke(aResultToken, IT_GET | IF_BYPASS_METAFUNC, _T("__value"), ExprTokenType(nested), nullptr, 0);
			if (result != INVOKE_NOT_HANDLED)
			{
				if (--nested->mRefCount == 0)
					--mRefCount;
				return result;
			}
		}
		aResultToken.SetValue(nested);
	}
	else
	{
		TypedPtrToToken(aProp.type, (void*)ptr, aResultToken);
		ASSERT(aResultToken.symbol != SYM_OBJECT); // Shouldn't happen since we don't support typed Object-pointer properties, but if it happened we may need to AddRef().
	}
	return OK;
}


ResultType Object::GetBoxedPointer(ResultToken &aResultToken, UINT_PTR aPtr, Object *aPrototype, size_t aNestOffset)
{
	auto nest = (Object**)((char*)this + aNestOffset);
	auto sp = *nest;
	if (sp && sp->DataPtr() != aPtr)
	{
		// Cached struct object pointer no longer matches.
		*nest = nullptr;
		sp->Release();
		sp = nullptr;
	}
	if (!sp)
	{
		if (!aPtr)
		{
			aResultToken.Unset();
			return OK;
		}
		*nest = sp = CreateStructPtr(aPrototype, aPtr);
		if (!sp)
			return FAIL;
	}
	sp->AddRef();
	aResultToken.SetValue(sp);
	return OK;
}


ResultType Object::SetTypedValue(ResultToken &aResultToken, int aFlags, name_t aName, TypedProperty &aProp, ExprTokenType &aValue)
{
	auto ptr = DataPtr() + aProp.data_offset;
	if (aProp.class_object)
	{
		if (aProp.pointed_proto) // Pointer type.
		{
			return SetBoxedPointer(aResultToken, aValue, *(UINT_PTR*)ptr, aProp.pointed_proto, aProp.object_offset, aProp.class_object);
		}
		auto nested = (Object*)((char*)this + aProp.object_offset);
		if (*(UINT_PTR*)nested == 0) // Since it wasn't constructed, this must be a pointer, not a real struct.
		{
			auto result = NestedSparseInit(aResultToken, aProp, ptr);
			if (result != OK)
				return result;
		}
		mRefCount++; // Must be done at least when nested->mRefCount == 0 (and then reversed when nested->mRefCount reaches 0 again).
		nested->mRefCount++; // Avoid calling Delete() when the __value setter returns.
		auto param = &aValue;
		auto result = nested->Invoke(aResultToken, IT_SET | IF_BYPASS_METAFUNC | IF_NO_NEW_PROPS, _T("__Value"), ExprTokenType(nested), &param, 1);
		nested->mRefCount--;
		mRefCount--;
		if (result != INVOKE_NOT_HANDLED)
			return result;
		return aResultToken.Error(_T("Assignment to struct is not supported."));
	}
	return SetValueOfTypeAtPtr(aProp.type, (void*)ptr, aValue, aResultToken);
}


ResultType Object::SetBoxedPointer(ResultToken &aResultToken, ExprTokenType &aValue, UINT_PTR &aPtr, Object *aPrototype, size_t aNestOffset, Object *aPointerClass)
{
	auto v = TokenToObject(aValue);
	Object *p;
	UINT_PTR np;
	if (v && v->IsOfType(aPrototype))
	{
		p = (Object*)v;
		np = p->DataPtr();
		if (!np)
			p = nullptr;
	}
	else if (aValue.symbol == SYM_MISSING)
	{
		p = nullptr;
		np = 0;
	}
	else if (v && v->IsOfType(aPointerClass ? aPointerClass->ClassGetPrototype() : Base())
		&& (np = ((Object*)v)->DataPtr())) // A pointer struct with no typed data is invalid.
	{
		p = (Object*)v; // Struct.Ptr
		np = *(UINT_PTR*)np; // p->DataPtr() was the address of the pointer variable, so dereference it.
		auto pnest = (Object**)((char*)p + sizeof(Object));
		p = *pnest;
		if (p && p->DataPtr() != np)
			p = nullptr;
	}
	else
		return aResultToken.TypeError(aPrototype->GetOwnPropString(_T("__Class")), aValue);
	
	aPtr = np;

	if (p)
		p->AddRef();
	auto nest = (Object**)((char*)this + aNestOffset);
	if (*nest)
		(*nest)->Release();
	*nest = p;
	return OK;
}


void Object::StructGet(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case M_Struct_Ptr: _o_return(DataPtr());
	case M_Struct_Size: _o_return(mBase->StructSize());
	case M_CArray_Length:
		auto si = mBase->GetStructInfo();
		_o_return(si->item_count);
	}
}


void Object::CArrayItem(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto si = mBase->GetStructInfo();
	
	auto index = ParamIndexToInt64(aParamCount - 1);
	index += index < 0 ? si->item_count : -1;
	if (index < 0 || (size_t)index >= si->item_count)
		_o_throw(ERR_INVALID_INDEX, *aParam[aParamCount - 1], ErrorPrototype::Index);

	Object *item_class = si->native_type == MdType::Void ? si->pointed_class : nullptr;
	ASSERT(si->native_type != MdType::Void || item_class);

	Object *pointed_proto = nullptr;
	StructInfo *item_si = nullptr;
	if (item_class)
		if (auto proto = item_class->ClassGetPrototype())
		{
			item_si = proto->GetStructInfo(true);
			if (item_si->pointed_class && !item_si->item_count)
				pointed_proto = item_si->pointed_class->ClassGetPrototype();
		}

	size_t nested_size = item_si ? item_si->SizeWhenNested() : 0; // Private object size.
	size_t item_size = si->size / si->item_count;  // Public struct size.
	ASSERT(si->size == item_size * si->item_count);
	// TODO: cache some of the above information in si->first_field ?

	TypedProperty tp{ si->native_type, item_class, pointed_proto, (size_t)index * item_size, si->object_size + (size_t)index * nested_size };
	if (IS_INVOKE_GET)
		GetTypedValue(aResultToken, 0, tp);
	else
		SetTypedValue(aResultToken, 0, _T("__Item"), tp, *aParam[0]);
	tp.class_object = nullptr;
}


BIF_DECL(NewStruct)
{
	IObject *cls = ParamIndexToObject(0);
	Object *proto = cls && cls->IsOfType(Object::sPrototype) ? ((Object*)cls)->ClassGetPrototype() : nullptr;
	if (!proto)
		_f_throw_value(_T("Invalid class"));
	Object::CreateStruct(aResultToken, proto, aParam + 1, aParamCount - 1);
}


void Object::StructPtrInvoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto si = mBase->GetStructInfo();
	auto proto = si->pointed_class->ClassGetPrototype();
	auto &ptr = *(UINT_PTR*)DataPtr();
	if (IS_INVOKE_GET)
		GetBoxedPointer(aResultToken, ptr, proto, si->object_size);
	else
		SetBoxedPointer(aResultToken, *aParam[0], ptr, proto, si->object_size, nullptr);
}


ResultType Object::ApplyParams(ResultToken &aThisResultToken, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	// On input, aThisResultToken contains the value to invoke, having been either retrieved
	// from a field or returned from a property getter.  Callers rely on us to free any string
	// value if appropriate, regardless of whether the recursive invoke succeeds (it might
	// succeed if the script has defined String.Prototype.__Item, for instance).
	ResultToken this_token;
	this_token.CopyValueFrom(aThisResultToken);
	this_token.mem_to_free = aThisResultToken.mem_to_free;
	aThisResultToken.InitResult(aThisResultToken.buf);
	auto &aResultToken = aThisResultToken;
	
	IObject *this_obj = TokenToObject(this_token);
	if (!this_obj)
	{
		this_obj = ValueBase(this_token);
		aFlags |= IF_SUBSTITUTE_THIS;
	}

	auto result = this_obj->Invoke(aResultToken, aFlags, nullptr, this_token, aParam, aParamCount);

	if (aResultToken.symbol == SYM_STRING && !aResultToken.mem_to_free && aResultToken.marker != aResultToken.buf)
	{
		// Returned strings are sometimes in memory owned by the object, so make a copy
		// before potentially releasing this_obj via this_token.Free().
		if (!TokenSetResult(aResultToken, aResultToken.marker, aResultToken.marker_length))
			result = FAIL;
	}

	if (result == INVOKE_NOT_HANDLED)
	{
		// Something like obj.x[y] where obj.x exists but obj.x[y] does not.  Throw here
		// to override the default error message, which would indicate that "x" is unknown.
		result = aResultToken.UnknownMemberError(this_token, aFlags, nullptr);
	}

	this_token.Free();
	return result;
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
	Object_Member(__Item, __Item, 0, IT_SET | BIMF_UNSET_ARG_1, 1, 1),
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
				{
					if (g_script.BackCompatMode())
						_o_throw(ERR_ITEM_UNSET, *aParam[0], ErrorPrototype::UnsetItem);
					_o_return_unset;
				}
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
		if (aParam[0]->symbol == SYM_MISSING)
			return Delete(aResultToken, aID, aFlags, aParam + 1, 1);
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
	if (IS_INVOKE_SET)
		--aParamCount;
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
	auto &obj = *ClassGetPrototype();
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


Object *Object::ClassGetPrototypeBackwardCompatible()
{
	if (auto p = GetOwnPropObj(_T("Prototype")))
		return p->IsOfType(Object::sPrototype) ? (Object*)p : nullptr;
	return ClassGetPrototype();
}


bool Object::CanSetBase()
{
	switch (mFlags & (StructInfoInitialized | StructInfoLocked | DataIsSuffix | DataIsSuffixPtr))
	{
	case 0:
		return true;
	case StructInfoInitialized | StructInfoLocked:
		// Prototypes with no typed properties allow Base assignment for backward-compatibility.
		// Since it's locked, size == 0 is only possible if an instance was created or this is a
		// built-in Prototype which has been locked for whatever reason.
		return ((StructInfo*)(this + 1))->size == 0;
	case StructInfoInitialized: // StructInfo was already initialized from the current Base (and since it's not locked, size != 0 is implied).
	default: // this is an instance with typed properties.
		return false;
	}
}


bool Object::CanSetBase(Object *aBase)
{
	if (aBase && aBase->GetStructInfo()->size)
		return false;
	auto new_native_base = (!aBase || aBase->IsNativeClassPrototype())
		? aBase : aBase->GetNativeBase();
	return new_native_base == GetNativeBase() // Cannot change native type.
		&& !aBase->IsDerivedFrom(this) && aBase != this; // Cannot create loops.
}


ResultType Object::SetBase(Object *aNewBase, ResultToken &aResultToken)
{
	if (!CanSetBase())
		return aResultToken.ValueError(ERR_PROPERTY_READONLY);
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
	if (HasOwnProp(_T("__Class")))
		return _T("Prototype"); // This object is a prototype.
	for (base = mBase; base; base = base->mBase)
		if (auto classname = base->GetOwnPropString(_T("__Class")))
			return classname; // This object is an instance of that class.
	return _T("Object"); // Provide a default in case __Class has been removed from all of the base objects.
}


Object *Object::CreateClass(Object *aPrototype, Object *aBase)
{
	auto cls = new (sizeof(Object*)) Object(ObjectIsClass);
	cls->SetBase(aBase);
	if (!sStructPrototype || !aPrototype->IsDerivedFrom(sStructPrototype))
		cls->SetOwnProp(_T("Prototype"), aPrototype);
	*(Object**)(cls + 1) = aPrototype;
	aPrototype->AddRef();
	return cls;
}


Object *Object::CreatePrototype(LPTSTR aClassName, Object *aBase)
{
	auto obj = new (sizeof(StructInfo)) Object();
	obj->mFlags |= ClassPrototype;
	obj->SetOwnProp(_T("__Class"), ExprTokenType(aClassName), false);
	obj->SetBase(aBase);
	ZeroMemory(obj + 1, sizeof(StructInfo));
	return obj;
}


STARTUP_FUNCTION
Object *Object::CreatePrototype(LPTSTR aClassName, Object *aBase, ObjectMember aMember[], int aMemberCount)
{
	auto obj = CreatePrototype(aClassName, aBase);
	return DefineMembers(obj, aClassName, aMember, aMemberCount);
}

STARTUP_FUNCTION
Object *Object::CreatePrototype(LPTSTR aClassName, Object *aBase, ObjectMemberMd aMember[], int aMemberCount)
{
	auto obj = CreatePrototype(aClassName, aBase);
	return DefineMetadataMembers(obj, aClassName, aMember, aMemberCount);
}

Object *Object::CreatePrototype(LPTSTR aClassName, Object *aBase, ObjectMemberListType aMember)
{
	if (aMember.duck)
		return CreatePrototype(aClassName, aBase, aMember.duck, aMember.count);
	else
		return CreatePrototype(aClassName, aBase, aMember.meta, aMember.count);
}


STARTUP_FUNCTION
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
			prop->NoParamGet = prop->NoParamSet = member.maxParams == 0;
			prop->NoEnumGet = member.minParams > 0;
			
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
			
			if ((member.invokeType & IT_BITMASK) == IT_SET) // & allows for additional flags.
			{
				_tcscpy(op_name, _T(".Set"));
				func = new BuiltInMethod(SimpleHeap::Alloc(full_name));
				func->mBIM = member.method;
				func->mMID = member.id;
				func->mMIT = member.invokeType;
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

STARTUP_FUNCTION
Object *Object::CreateClass(LPTSTR aClassName, Object *aBase, Object *aPrototype, ClassFactoryDef aFactory)
{
	auto class_obj = CreateClass(aPrototype, aBase);

	if (aFactory.call)
	{
		TCHAR full_name[MAX_VAR_NAME_LENGTH + 1];
		_stprintf(full_name, _T("%s.Call"), aClassName);
		auto ctor = new BuiltInFunc(SimpleHeap::Alloc(full_name));
		if (aFactory.is_bif)
		{
			ctor->mBIF = (BuiltInFunctionType)aFactory.call;
			ctor->mFID = FID_Object_New;
		}
		else
		{
			auto &si = *(StructInfo*)(aPrototype + 1);
			si.create = (NewObjectProc)aFactory.call;
			si.object_size = aFactory.object_size;
			ctor->mData = aPrototype;
			ctor->mBIF = NewInstance;
			aPrototype->mFlags |= StructInfoInitialized | StructInfoLocked;
		}
		ctor->mMinParams = aFactory.min_params; // Usually 1, the class object.
		ctor->mParamCount = aFactory.max_params;
		ctor->mIsVariadic = aFactory.is_variadic; // Usually variadic since __new(...) may be redefined/overridden.
		class_obj->DefineMethod(_T("Call"), ctor);
		ctor->Release();
	}

	auto var = g_script.FindOrAddVar(aClassName, 0, VAR_DECLARE_GLOBAL | VAR_EXPORTED);
	var->AssignSkipAddRef(class_obj);
	var->MakeReadOnly();

	return class_obj;
}

void Object::CreatePtrClass(ResultToken &aResultToken, ExprTokenType &aToClass)
{
	auto sc_ = TokenToObject(aToClass);
	auto sc = sc_->IsOfType(Object::sPrototype) ? (Object*)sc_ : nullptr;
	Object *sp = sc ? sc->ClassGetPrototype() : nullptr;
	auto spsi = sp ? (StructInfo*)(sp + 1) : nullptr;
	if (spsi && spsi->pointer_class)
	{
		// Return the previously created class.
		if (++spsi->pointer_class->mRefCount == 1)
			++sc->mRefCount; // Must AddRef() the pointed class whenever the pointer class becomes ref-counted. 
		_f_return(spsi->pointer_class);
	}
	if (!spsi || !sp->IsDerivedFrom(Object::sStructPrototype))
		return (void)aResultToken.TypeError(_T("Struct class"), aToClass);

	auto ptr_cls = CreatePtrClass(sc, sp, spsi);
	//ptr_cls->AddRef(); // mRefCount remains at 1 because we want the corresponding Release() to call Delete().
	_f_return(ptr_cls);
}

Object *Object::CreatePtrClass(Object *sc, Object *sp, StructInfo *spsi)
{
	ASSERT(sc && sp && spsi && !spsi->pointer_class);

	Object *bsp = sp->Base();
	auto bsi = (StructInfo*)(bsp + 1);
	auto bpc = bsi->pointer_class;
	if (!bpc)
		bpc = CreatePtrClass(sc->Base(), bsp, bsi);
	else if (bpc->mRefCount == 0) // About to become 1.
		bpc->mOuter->mRefCount++; // Must AddRef() the pointed class whenever the pointer class becomes ref-counted. 
	ASSERT(bpc);

	LPTSTR aClassName = sp->GetOwnPropString(_T("__Class"));
	auto len = _tcslen(aClassName);
	auto buf = len ? (LPTSTR)_malloca((len + _countof(STRUCT_PTR_CLASS_SUFFIX)) * sizeof(TCHAR)) : nullptr;
	LPTSTR class_name;
	if (buf)
	{
		tmemcpy(buf, aClassName, len);
		tmemcpy(buf + len, STRUCT_PTR_CLASS_SUFFIX, _countof(STRUCT_PTR_CLASS_SUFFIX));
		class_name = buf;
	}
	else
		class_name = STRUCT_PTR_CLASS_NAME;

	auto ptr_pro = CreatePrototype(class_name, bpc->ClassGetPrototype());
	auto ptr_cls = CreateClass(ptr_pro, bpc);
	ptr_cls->mOuter = sc;
	spsi->pointer_class = ptr_cls;
	ptr_pro->mFlags |= StructInfoInitialized | StructInfoLocked;
	auto si = (StructInfo*)(ptr_pro + 1);
	si->align = si->size = sizeof(void*);
	si->nested_object_size = sizeof(Object*);
	si->pointed_class = sc;
	if (sc)
		sc->AddRef();
	si->object_size = sizeof(Object);
	if (spsi->dllcall_type && !spsi->pointed_class)
	{
		si->dllcall_type = spsi->dllcall_type;
		si->is_unsigned = spsi->is_unsigned;
	}
	ptr_pro->Release();
	_freea(buf);

	return ptr_cls;
}

BIF_DECL(StructClass_Ptr)
{
	Object::CreatePtrClass(aResultToken, *aParam[0]);
	if (_f_callee_id && aResultToken.symbol == SYM_OBJECT)
	{
		auto cls = (Object*)aResultToken.object;
		aResultToken.InitInvokeRetVal();
		cls->Invoke(aResultToken, IT_CALL, nullptr, ExprTokenType{ cls }, aParam + 1, aParamCount - 1);
		cls->Release();
	}
}

void Object::CreateCArrayClass(ResultToken &aResultToken, ExprTokenType &aOfClass, size_t aCount)
{
	auto sc_ = TokenToObject(aOfClass);
	auto sc = sc_->IsOfType(Object::sPrototype) ? (Object*)sc_ : nullptr;
	auto sp = sc ? sc->ClassGetPrototype() : nullptr;
	auto spsi = sp ? sp->GetStructInfo(true) : nullptr;
	Map *map = spsi ? spsi->array_class_map : nullptr;
	ExprTokenType key = (__int64)aCount;
	if (!map)
	{
		if (!spsi || !sp->IsDerivedFrom(Object::sStructPrototype))
			return (void)aResultToken.TypeError(_T("Struct class"), aOfClass);
		spsi->array_class_map = map = Map::Create();
	}
	else if (map->GetItem(aResultToken, key))
	{
		ASSERT(aResultToken.symbol == SYM_OBJECT);
		auto ac = (Object*)aResultToken.object;
		if (++ac->mRefCount == 1)
			++sc->mRefCount; // Must AddRef() the element class whenever the array class becomes ref-counted. 
		return;
	}

	TCHAR class_name[MAX_CLASS_NAME_LENGTH + 1];
	sntprintf(class_name, _countof(class_name), _T("%s[%zi]"), sp->GetOwnPropString(_T("__Class")), aCount);

	// No cached class, so create one.
	auto ap = CreatePrototype(class_name, Object::sCArrayPrototype);
	auto ac = CreateClass(ap, Object::sCArrayClass);
	auto si = (StructInfo*)(ap + 1);
	ap->mFlags |= StructInfoInitialized | StructInfoLocked;
	ap->Release();

	// Cache it.
	if (!map->SetItem(key, ExprTokenType(ac)))
	{
		ac->Release();
		return (void)aResultToken.MemoryError();
	}
	ac->mOuter = sc;
	ac->mRefCount--;
	ASSERT(ac->mRefCount == 1); // Only the reference to be returned below is counted.
	sc->AddRef();

	si->object_size = sizeof(Object);
	if (!spsi->item_count)
		si->native_type = spsi->native_type;
	if (si->native_type == MdType::Void)
	{
		si->pointed_class = sc;
		si->nested_object_size = aCount * spsi->SizeWhenNested();
	}
	si->size = aCount * spsi->size;
	si->align = spsi->align;
	si->item_count = aCount;

	aResultToken.SetValue(ac);
}

BIF_DECL(StructClass_Item)
{
	if (!ParamIndexIsNumeric(1))
		return (void)aResultToken.ParamError(0, aParam[1], _T("Integer"));
	auto count = ParamIndexToInt64(1);
	if (count < 1)
		return (void)aResultToken.ParamError(0, aParam[1]);
	Object::CreateCArrayClass(aResultToken, *aParam[0], (size_t)count);
}


//
// Object:: and Map:: Built-ins
//

void Object::DeleteProp(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto field = FindField(ParamIndexToString(0, _f_number_buf));
	if (!field)
		_o_return_unset_blank;
	field->ReturnMove(aResultToken); // Return the removed value.
	mFields.Remove((index_t)(field - mFields), 1);
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
		// removed from this[arg], but there wasn't one.
		if (g_script.BackCompatMode())
			_o_throw(ERR_ITEM_UNSET, *aParam[0], ErrorPrototype::UnsetItem);
		_o_return_unset;
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
	_o_return(new IndexEnumerator(this, ParamIndexToOptionalInt(0, 0)
		, static_cast<IndexEnumerator::Callback>(&Object::GetEnumProp)));
}

void Map::__Enum(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	_o_return(new IndexEnumerator(this, ParamIndexToOptionalInt(0, 0)
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

Property *Object::DefineProperty(name_t aName, bool aEnumerable)
{
	index_t insert_pos;
	auto field = FindField(aName, insert_pos);
	if (!field && !(field = Insert(aName, insert_pos)))
		return nullptr;
	field->enumerable = aEnumerable;
	if (field->symbol != SYM_DYNAMIC)
	{
		field->Free();
		field->symbol = SYM_DYNAMIC;
		field->prop = new Property();
	}
	return field->prop;
}

TypedProperty *Object::DefineTypedProperty(name_t aName)
{
	ASSERT((mFlags & (ClassPrototype | StructInfoInitialized)) == (ClassPrototype | StructInfoInitialized));
	index_t insert_pos;
	auto field = FindField(aName, insert_pos);
	if (field)
		field->Free();
	else if (!(field = Insert(aName, insert_pos)))
		return nullptr;
	auto tprop = new TypedProperty();
	// Add it to the new field.
	field->symbol = SYM_TYPED_FIELD;
	field->tprop = tprop;
	// Add it to the Prototype's linked list of struct fields.
	auto &si = *(StructInfo*)(this + 1);
	if (si.first_field) // Check first_field and not last_field, which may belong to a superclass.
		si.last_field->next_field = tprop;
	else
		si.first_field = tprop;
	tprop->next_field = nullptr;
	tprop->prev_field = si.last_field;
	si.last_field = tprop;
	return tprop;
}

FResult Object::DefineTypedProperty(name_t aName, Object *aClass, size_t aPack, size_t aOffset)
{
	size_t psize = 0, palign = 0;
	MdType native_type = MdType::Void;
	StructInfo *psi = nullptr;
	if (aClass)
	{
		auto proto = aClass->ClassGetPrototype();
		if (proto && proto->IsDerivedFrom(Object::sStructPrototype))
		{
			psi = proto->GetStructInfo(true);
			if (psi->native_type != MdType::Void && !psi->item_count)
			{
				aClass = nullptr;
				native_type = psi->native_type;
			}
			psize = psi->size;
			palign = psi->align;
		}
	}
	if (!psize)
		return FR_E_ARGS;
	auto si = (mFlags & (StructInfoLocked | ClassPrototype)) != ClassPrototype ? nullptr
		: GetStructInfo(false);
	if (!si)
		return FR_E_FAILED;
	auto tprop = DefineTypedProperty(aName);
	if (!tprop)
		return FR_E_OUTOFMEM;
	tprop->type = native_type;
	tprop->pointed_proto = nullptr;
	if (tprop->class_object = aClass)
	{
		aClass->AddRef();
		tprop->object_offset = si->object_size + si->nested_object_size;
		si->nested_object_size += psi->SizeWhenNested();
		if (psi->IsPointerType())
		{
			tprop->pointed_proto = psi->pointed_class->ClassGetPrototype();
			tprop->pointed_proto->AddRef();
		}
	}
	if (aPack && palign > aPack)
		palign = aPack;
	if (palign > si->align)
		si->align = palign;
	ASSERT(palign && ((palign & (palign - 1)) == 0)); // Must be a power of 2.
	if (aOffset == -1)
		aOffset = (si->size + palign - 1) & ~(palign - 1);
	tprop->data_offset = aOffset;
	aOffset += psize;
	if (si->size < aOffset)
		si->size = aOffset; // size may be unaligned until the struct definition is closed (if palign < si-align).
	return OK;
}

Object::StructInfo *Object::GetStructInfo()
{
	// Callers use this simple version when it's known that the StructInfo
	// was already initialized and locked; e.g. because the caller is an
	// instance of the struct.
	if (!(mFlags & StructInfoInitialized))
		return mBase->GetStructInfo();
	return (StructInfo*)(this + 1);
}

Object::StructInfo *Object::GetStructInfo(bool aLock)
{
	if (!(mFlags & StructInfoInitialized))
	{
		// Lock base definition, and either return it for caller or initialize ours.
		auto &bsi = *mBase->GetStructInfo(true);

		if (!(mFlags & ClassPrototype)) // Only Prototypes can have StructInfo.
			return &bsi;

		// Even if this ends up being locked below as an exact copy of base,
		// initialize it to reduce the need for recursive calls at runtime.
		auto &si = *(StructInfo*)(this + 1);
		si.create = bsi.create;
		si.object_size = bsi.object_size;
		si.size = bsi.size;
		si.align = bsi.align;
		si.nested_object_size = bsi.nested_object_size;
		si.item_count = bsi.item_count;
		si.last_field = bsi.last_field; // With prev_field, this forms a reversed list of all fields, including inherited fields.
		si.pointed_class = bsi.pointed_class;
		// The following were already zero-initialized:
		//si.first_field = nullptr; // This lists fields defined *directly* within this Prototype.
		//si.pointer_class = nullptr; // Each subclass should get its own (dynamically).
		//si.array_class_map = nullptr; // Each subclass must create its own map.
		//si.native_type = MdType::Void; // Revert to a normal struct if extending a numeric type.
		//si.dllcall_type = DLL_ARG_INVALID; // As above.
		mFlags |= StructInfoInitialized;
		if (aLock)
			mFlags |= StructInfoLocked;
	}
	else if (aLock && !(mFlags & StructInfoLocked))
	{
		// Lock to prevent more fields from being added, such as when another struct's
		// layout depends on the size of this one (could be a derived class or one which
		// uses this class as a property or element type).
		mFlags |= StructInfoLocked;
		// Apply the struct's final alignment requirement to its size.
		auto &si = *(StructInfo*)(this + 1);
		si.size = (si.size + si.align - 1) & ~(si.align - 1);
	}
	return (StructInfo*)(this + 1);
}

UINT_PTR Object::StructSize()
{
	if (!(mFlags & StructInfoInitialized))
		return mBase->StructSize();
	return ((StructInfo*)(this + 1))->size;
}

MdType Object::GetStructMdType()
{
	if (!(mFlags & StructInfoInitialized))
		return MdType::Void;
	auto &si = *(StructInfo*)(this + 1);
	return si.item_count == 0 ? si.native_type : MdType::Void;
}

ResultType FillPropertyFlags(IObject *aObj, bool aSetter, Property &aProp, ResultToken &aResultToken)
{
	bool &no_param = aSetter ? aProp.NoParamSet : aProp.NoParamGet;
	no_param = false; // Reset to default, in case of error or undefined MaxParams.
	__int64 propval;
	ResultType result;
	if (!aSetter)
	{
		aProp.NoEnumGet = false; // Reset to default, in case of error or undefined MaxParams.
		propval = 0;
		result = GetObjectIntProperty(aObj, _T("MinParams"), propval, aResultToken, true);
		switch (result)
		{
		case FAIL:
		case EARLY_EXIT:
			return result;
		case OK:
			aProp.NoEnumGet = propval > 1;
		}
	}
	propval = 0;
	result = GetObjectIntProperty(aObj, _T("MaxParams"), propval, aResultToken, true);
	switch (result)
	{
	case FAIL:
	case EARLY_EXIT:
		return result;
	case OK:
		no_param = propval == (aSetter ? 2 : 1);
		if (!no_param)
			break; // No need to query IsVariadic.
		propval = 0;
		result = GetObjectIntProperty(aObj, _T("IsVariadic"), propval, aResultToken, true);
		switch (result)
		{
		case FAIL:
		case EARLY_EXIT:
			return result;
		case OK:
			if (propval)
				no_param = false; // Reset to false; property accepts parameters.
		}
	}
	return OK;
}

void Object::DefineProp(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (mFlags & CannotOwnProps)
		_o_throw_type(_T("Object"), ExprTokenType(this));
	auto name = ParamIndexToString(0, _f_number_buf);
	if (!*name)
		_o_throw_param(0 + aID);
	ExprTokenType getter, setter, method, value;
	getter.symbol = SYM_INVALID;
	setter.symbol = SYM_INVALID;
	method.symbol = SYM_INVALID;
	value.symbol = SYM_INVALID;
	auto desc = dynamic_cast<Object *>(ParamIndexToObject(1));
	if (desc && desc->GetOwnProp(value, _T("Type"))) // TODO: make this properly mutually exclusive with the others
	{
		Object *pclass = dynamic_cast<Object*>(TokenToObject(value));
		size_t pack = desc->GetOwnProp(value, _T("Pack")) ? (size_t)TokenToInt64(value) : 0;
		size_t offset = -1;
		if (desc->GetOwnProp(value, _T("Offset")))
		{
			if (value.symbol == SYM_STRING)
			{
				auto f = FindField(value.marker);
				if (f && f->symbol == SYM_TYPED_FIELD)
					offset = f->tprop->data_offset;
			}
			else if (value.symbol == SYM_INTEGER && value.value_int64 >= 0)
			{
				offset = (size_t)value.value_int64;
			}
			if (offset == -1)
				_o_throw_value(aID ? ERR_PARAM3_INVALID : ERR_PARAM2_INVALID);
		}
		switch (DefineTypedProperty(name, pclass, pack, offset))
		{
		case OK:
			AddRef();
			_o_return(this);
		case FR_E_ARGS:
			_o_throw_value(aID ? ERR_PARAM3_INVALID : ERR_PARAM2_INVALID);
		case FR_E_OUTOFMEM:
			_o_throw_oom;
		default:
			_o_throw(_T("Cannot add typed property."));
		}
	}
	if (!desc // Must be an Object.
		|| desc->GetOwnProp(getter, _T("Get")) && getter.symbol != SYM_OBJECT  // If defined, must be an object.
		|| desc->GetOwnProp(setter, _T("Set")) && setter.symbol != SYM_OBJECT
		|| desc->GetOwnProp(method, _T("Call")) && method.symbol != SYM_OBJECT
		|| desc->GetOwnProp(value, _T("Value")) && (getter.symbol != SYM_INVALID || setter.symbol != SYM_INVALID || method.symbol != SYM_INVALID)
		// To help prevent errors, throw if none of the above properties were present.  This also serves to
		// reserve some cases for possible future use, such as passing a function object to imply {get:...}.
		|| getter.symbol == SYM_INVALID && setter.symbol == SYM_INVALID && method.symbol == SYM_INVALID && value.symbol == SYM_INVALID)
		_o_throw_param(1 + aID);
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
	if (getter.symbol == SYM_OBJECT)
	{
		prop->SetGetter(getter.object);
		FillPropertyFlags(getter.object, false, *prop, aResultToken);
	}
	if (setter.symbol == SYM_OBJECT)
	{
		prop->SetSetter(setter.object);
		FillPropertyFlags(setter.object, true, *prop, aResultToken);
	}
	if (method.symbol == SYM_OBJECT)
	{
		prop->SetMethod(method.object);
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
	{
		if (g_script.BackCompatMode())
			_o__ret(aResultToken.UnknownMemberError(ExprTokenType(this), IT_GET, name));
		_o_return_unset;
	}
	auto desc = Object::Create();
	desc->SetInternalCapacity(field->symbol == SYM_DYNAMIC ? 3 : 1);
	if (field->symbol == SYM_DYNAMIC)
	{
		if (auto getter = field->prop->Getter()) desc->SetOwnProp(_T("Get"), getter);
		if (auto setter = field->prop->Setter()) desc->SetOwnProp(_T("Set"), setter);
		if (auto method = field->prop->Method()) desc->SetOwnProp(_T("Call"), method);
	}
	else if (field->symbol == SYM_TYPED_FIELD)
	{
		if (field->tprop->class_object)
			desc->SetOwnProp(_T("Type"), field->tprop->class_object);
		else
			desc->SetOwnProp(_T("Type"), sPrimitiveClass[(int)field->tprop->type-1]);
		desc->SetOwnProp(_T("Offset"), field->tprop->data_offset);
	}
	else
	{
		ExprTokenType value;
		field->ToToken(value);
		desc->SetOwnProp(_T("Value"), value);
	}
	_o_return(desc);
}

void NewPropRef(ResultToken &aResultToken, IObject *aObj, LPCTSTR aName)
{
	auto new_name = _tcsdup(aName);
	if (!new_name)
		_f_throw_oom;
	aObj->AddRef();
	_f_return(new PropRef(aObj, new_name));
}

void Object::__Ref(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	auto name = ParamIndexToString(0, _f_retval_buf);
	if (!*name)
		_f_throw_param(0);

	for (Object *that = this; that; that = that->mBase)
	{
		if (auto field = that->FindField(name))
		{
			if (field->symbol != SYM_TYPED_FIELD || !field->tprop->class_object || field->tprop->pointed_proto)
				break;
			auto nested = (Object*)((char*)this + field->tprop->object_offset);
			if (++nested->mRefCount == 1) // Nested objects have this unique requirement.
				++mRefCount;
			_o_return(nested);
		}
	}

	NewPropRef(aResultToken, this, name);
}

BIF_DECL(PropRef_Call)
{
	++aParam, --aParamCount; // Exclude "PropRef" itself.
	auto that = ParamIndexToObject(0);
	if (!that)
		_f_throw_param(0, _T("object"));
	auto name = ParamIndexToString(1, _f_retval_buf);
	if (!*name)
		_f_throw_param(1);
	NewPropRef(aResultToken, that, name);
}

void PropRef::__Value(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	if (mThat->Invoke(aResultToken, aFlags, mMember, ExprTokenType(mThat), aParam, aParamCount) == INVOKE_NOT_HANDLED)
		_o_return_unset;
}


//
// Class objects
//

ResultType Object::Initialize(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	if (auto si = mBase->GetStructInfo(true))
	{
		if (si->nested_object_size >= sizeof(Object)) // May have constructible properties.
		{
			auto result = si->item_count ? CArrayNew(aResultToken, si)
				: NestedNew(aResultToken, DataPtr(), mBase);
			if (result != OK)
				return result;
		}
	}
	return CallInitNew(aResultToken, aParam, aParamCount);
}

ResultType Object::NestedNew(ResultToken &aResultToken, UINT_PTR aPtr, Object *aBase)
{
	ASSERT(aBase->IsClassPrototype());
	auto si = (StructInfo*)(aBase + 1);
	if (si->nested_object_size < sizeof(Object)) // Definitely no constructible properties defined by aBase.
		return OK;

	ResultType result = OK;
	if (aBase->mBase) // Construct inherited nested objects first.
	{
		result = NestedNew(aResultToken, aPtr, aBase->mBase);
		if (result != OK)
			return result;
	}
	
	for (auto tprop = si->first_field; tprop; tprop = tprop->next_field)
	{
		if (!tprop->class_object || tprop->pointed_proto) // Primitive or Ptr
			continue;
		auto proto = tprop->class_object->ClassGetPrototype();
		
		// Construct the nested object in the space reserved for it.
		void *nest = (char*)this + tprop->object_offset;
		ASSERT(!*(UINT_PTR*)nest);
		auto nested = ::new (nest) Object(CannotOwnProps | DataIsSuffixPtr);
		++mRefCount;
		nested->mOuter = this;
		nested->SetBase(proto);
		nested->SetDataPtr(aPtr + tprop->data_offset);
		result = nested->Initialize(aResultToken, nullptr, 0);
		if (result != OK)
		{
			// On failure, Initialize already called nested->Release(), which resets its mRefCount
			// to 0 and counteracts our ++mRefCount.
			Release(); // this object won't be returned, since construction failed.
			break;
		}
		// During construction, 'nested' has a non-zero mRefCount and a counted reference to 'this'.
		// Now it needs to have mRefCount == 0 to reflect that there aren't any external references.
		if (--nested->mRefCount == 0)
			--mRefCount;
		aResultToken.symbol = SYM_INTEGER; // CallNew has set this to nested.  Reset to default without calling Release().
		ASSERT(nested->mRefCount == 0 && mRefCount);
	}
	return result;
}

ResultType Object::CArrayNew(ResultToken &aResultToken, StructInfo *si)
{
	ASSERT(si->nested_object_size && si->pointed_class);

	auto item_base = si->pointed_class->ClassGetPrototype();
	if (!item_base || !item_base->IsDerivedFrom(sStructPrototype)) // FIXME: make this check unnecessary, either by making StructClass.Prototype read-only or storing the prototype elsewhere
		return aResultToken.Error(_T("Bad Prototype"), nullptr, ErrorPrototype::Type);
	
	auto item_si = item_base->GetStructInfo();
	if (item_si->IsPointerType())
		return OK; // Nothing needed beyond the zero-initialization already performed.
	size_t nested_size = item_si->SizeWhenNested(); // Private object size.
	size_t item_size = item_si->size; // Public struct size.

	auto data_ptr = DataPtr();
	char *nest = (char*)this + si->object_size;

	ResultType result = OK;
	for (size_t i = 0; i < si->item_count; ++i, data_ptr += item_size, nest += nested_size)
	{
		// Construct the nested object in the space reserved for it.
		ASSERT(!*(UINT_PTR*)nest);
		auto nested = ::new (nest) Object(CannotOwnProps | DataIsSuffixPtr);
		++mRefCount;
		nested->mOuter = this;
		nested->SetBase(item_base);
		nested->SetDataPtr(data_ptr);
		result = nested->Initialize(aResultToken, nullptr, 0);
		if (result != OK)
		{
			Release(); // this object won't be returned, since construction failed.
			break;
		}
		// During construction, 'nested' has a non-zero mRefCount and a counted reference to 'this'.
		// Now it needs to have mRefCount == 0 to reflect that there aren't any external references.
		if (--nested->mRefCount == 0)
			--mRefCount;
		aResultToken.symbol = SYM_INTEGER; // CallNew has set this to nested.  Reset to default without calling Release().
		ASSERT(nested->mRefCount == 0 && mRefCount);
	}
	return result;
}

ResultType Object::NestedSparseInit(ResultToken& aResultToken, TypedProperty& aProp, UINT_PTR aPtr)
{
	ASSERT(!aProp.pointed_proto);
	auto proto = aProp.class_object->ClassGetPrototype();
	if (!proto)
		return INVOKE_NOT_HANDLED;
	auto nest = (char*)this + aProp.object_offset;
	ASSERT(*(UINT_PTR*)nest == 0);
	auto nested = ::new (nest) Object(CannotOwnProps | DataIsSuffixPtr | NoCallDelete);
	nested->mOuter = this;
	nested->SetBase(proto);
	nested->SetDataPtr(aPtr);
	nested->mRefCount--; // Zero refcount to signify there are no external references yet.
	return OK;
}

ResultType Object::CallInitNew(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
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

	return CallNew(aResultToken, aParam, aParamCount, this_token);
}

ResultType Object::CallNew(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, ExprTokenType &aThisToken)
{
	// __New may be defined by the script for custom initialization code.
	auto result = CallMeta(_T("__New"), aResultToken, aThisToken, aParam, aParamCount);
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
		prop = new Property(*val.prop);
		if (auto obj = prop->Getter()) obj->AddRef();
		if (auto obj = prop->Setter()) obj->AddRef();
		if (auto obj = prop->Method()) obj->AddRef();
		break;
	case SYM_TYPED_FIELD:
		// FIXME: This and other parts of the cloning process do not support typed properties.
		AssignMissing();
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
	//case SYM_MISSING: // Callers don't need special handling for this.
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
		// This implements "blank if none" documented for some methods in v2.0.
		result.Unset(UnsetKind::Blank); // Behaves as unset only in v2.1 mode.
		break;
	case SYM_DYNAMIC:
	case SYM_TYPED_FIELD: // This is a field definition; it can't have a value.
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
	case SYM_TYPED_FIELD:
		// This can be reached via Object::GetOwnProp.
		aToken.symbol = SYM_INVALID; // Allow caller to detect this as an error.
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

TypedProperty::~TypedProperty()
{
	if (class_object)
		class_object->Release();
	if (pointed_proto)
		pointed_proto->Release();
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
	Object_Member(__Item, Invoke, P___Item, IT_SET | BIMF_UNSET_ARG_1, 1, 1),
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
			if (g_script.BackCompatMode())
				_o_throw(ERR_ITEM_UNSET, *aParam[0], ErrorPrototype::UnsetItem);
			_o_return_unset;
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
		_o_return_unset_blank;
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
		bool return_it = ParamIndexIsOmitted(1);
		if (!return_it)
		{
			Throw_if_Param_NaN(1);
			count = (index_t)ParamIndexToInt64(1);
		}
		if (index + count > mLength)
			_o_throw_param(1);

		if (return_it) // Remove-and-return mode.
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
		_o_return(new IndexEnumerator(this, ParamIndexToOptionalInt(0, 0)
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
		ResultType result = OK;
		if (aVarCount > 1)
		{
			// Put the index first, only when there are two parameters.
			if (aVal)
				result = aVal->Assign((__int64)aIndex + 1);
			aVal = aReserved;
		}
		if (aVal && result)
		{
			auto &item = mItem[aIndex];
			switch (item.symbol)
			{
			default:
				if (item.symbol == SYM_MISSING)
					result = aVal->AssignUnset();
				else
					result = aVal->AssignString(item.string, item.string.Length());
				break;
			case SYM_INTEGER:	result = aVal->Assign(item.n_int64);	break;
			case SYM_FLOAT:		result = aVal->Assign(item.n_double);	break;
			case SYM_OBJECT:	result = aVal->Assign(item.object);		break;
			}
		}
		return result ? CONDITION_TRUE : FAIL;
	}
	return CONDITION_FALSE;
}



//
// Enumerator
//

bool EnumBase::Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount)
{
	Var *var[] { nullptr, nullptr };
	for (int i = 0; i < _countof(var); ++i)
		if (i < aParamCount)
			if (IObject *obj = ParamIndexToObject(i))
			{
				var[i] = new (_alloca(sizeof(Var))) Var(obj); // mType = VAR_VIRTUAL_OBJ
			}
			else if (aParam[i]->symbol != SYM_MISSING)
			{
				aResultToken.ParamError(i, aParam[i], _T("variable reference"));
				return false;
			}

	auto result = Next(var[0], var[1]);
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
	return (mObject->*mGetItem)(++mIndex, var0, var1, mParamCount ? mParamCount : var1 ? 2 : 1);
}


ResultType Object::GetEnumProp(UINT &aIndex, Var *aName, Var *aVal, int aVarCount)
{
	for  ( ; aIndex < mFields.Length(); ++aIndex)
	{
		FieldType &field = mFields[aIndex];
		// Assign name first to ensure stability in case the field is deleted by the property getter.
		if (aName)
			aName->Assign(field.name);
		if (aVal)
		{
			if (field.symbol == SYM_DYNAMIC)
			{
				// Skip it if it can't be called without parameters, or if there's no getter in this object
				// (consistent with inherited properties that have neither getter nor setter defined here).
				// Also skip if this is a class prototype, since that isn't an instance of the class and
				// therefore isn't a valid target for a method/property call.
				if (field.prop->NoEnumGet || !field.prop->Getter() || IsClassPrototype())
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
			else if (field.symbol == SYM_TYPED_FIELD)
			{
				// Typed properties are owned by the prototype, but have values only in the instances.
				continue;
			}
			else
			{
				ExprTokenType value;
				field.ToToken(value);
				aVal->Assign(value);
			}
		}
		return CONDITION_TRUE;
	}
	return CONDITION_FALSE;
}


Object::PropEnum::PropEnum(Object *aObject)
{
	for (Object *p = aObject; p; p = p->mBase)
		++mIndexCount;
	mIndex = new index_t[mIndexCount];
	memset(mIndex, 0, mIndexCount * sizeof(index_t));
	mObject = aObject;
	mObject->AddRef();
	mThisToken.SetValue(mObject);
}


Object::PropEnum::PropEnum(Object *aObject, ExprTokenType &aThisToken)
	: PropEnum(aObject)
{
	mThisToken.CopyValueFrom(aThisToken);
}


Object::PropEnum::~PropEnum()
{
	mObject->Release();
	delete[] mIndex;
}


ResultType Object::PropEnum::Next(Var *aName, Var *aVal)
{
	int nextidx, testidx = 0;
	Object *nextobj = nullptr;

	// Property getters should not be called for Prototype objects, since they are not instances.
	// Checking via mThisToken rather than mObject supports the substitution performed by the debugger.
	bool is_proto = mThisToken.symbol == SYM_OBJECT
		&& mThisToken.object->IsOfType(Object::sPrototype)
		&& static_cast<Object*>(mThisToken.object)->IsClassPrototype();

	for (Object *testobj = mObject;; )
	{
		if (mIndex[testidx] < testobj->mFields.Length())
		{
			auto &testfld = testobj->mFields[mIndex[testidx]];
			if (!testfld.enumerable
				|| testfld.symbol == SYM_DYNAMIC && (testfld.prop->NoEnumGet || !testfld.prop->Getter() || is_proto))
			{
				++mIndex[testidx]; // Skip this property.
				continue;
			}
			int r = nextobj ? _tcsicmp(testfld.name, nextobj->mFields[mIndex[nextidx]].name) : -1;
			if (r < 0)
			{
				nextidx = testidx;
				nextobj = testobj;
			}
			else if (r == 0)
			{
				++mIndex[testidx]; // Skip this shadowed property.
				// No need to consider the name at the new index, since r > 0 can be inferred.
			}
		}
		++testidx, testobj = testobj->mBase;
		if (!testobj || testidx >= mIndexCount)
			break; // No more bases.
	}
	if (!nextobj)
		return CONDITION_FALSE;

	UINT tempidx = mIndex[nextidx];

	auto &field = nextobj->mFields[mIndex[nextidx]++];

	ResultType result = OK;
	if (aName)
		result = aName->Assign(field.name);

	if (aVal && result)
	{
		FuncResult result_token;
		auto result = mObject->GetFieldValue(result_token, IT_GET | IF_BYPASS___VALUE, field, mThisToken);
		if (result == FAIL || result == EARLY_EXIT)
			return result;
		if (result_token.mem_to_free)
		{
			ASSERT(result_token.symbol == SYM_STRING && result_token.mem_to_free == result_token.marker);
			aVal->AcceptNewMem(result_token.mem_to_free, result_token.marker_length);
		}
		else
		{
			result = aVal->Assign(result_token);
			result_token.Free();
		}
	}
	
	return result ? CONDITION_TRUE : FAIL;
}


ResultType Map::GetEnumItem(UINT &aIndex, Var *aKey, Var *aVal, int aVarCount)
{
	if (aIndex < mCount)
	{
		auto &item = mItem[aIndex];
		ResultType result = OK;
		if (aKey)
		{
			if (aIndex < mKeyOffsetObject) // mKeyOffsetInt < mKeyOffsetObject
				result = aKey->Assign(item.key.i);
			else if (aIndex < mKeyOffsetString) // mKeyOffsetObject < mKeyOffsetString
				result = aKey->Assign(item.key.p);
			else // mKeyOffsetString < mCount
				result = aKey->Assign(item.key.s);
		}
		if (aVal && result)
		{
			ExprTokenType value;
			item.ToToken(value);
			result = aVal->Assign(value);
		}
		return result ? CONDITION_TRUE : FAIL;
	}
	return CONDITION_FALSE;
}


ResultType RegExMatchObject::GetEnumItem(UINT &aIndex, Var *aKey, Var *aVal, int aVarCount)
{
	if (aIndex >= (UINT)mPatternCount)
		return CONDITION_FALSE;
	// In single-var mode, return the subpattern values.
	// Otherwise, return the subpattern names first and values second.
	if (aVarCount == 1)
	{
		aVal = aKey;
		aKey = nullptr;
	}
	ResultType result = OK;
	if (aKey)
	{
		if (mPatternName && mPatternName[aIndex])
			result = aKey->Assign(mPatternName[aIndex]);
		else
			result = aKey->Assign((__int64)aIndex);
	}
	if (aVal && result)
	{
		result = aVal->Assign(mHaystack - mHaystackStart + mOffset[aIndex*2], mOffset[aIndex*2+1]);
	}
	return result ? CONDITION_TRUE : FAIL;
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
	field.enumerable = true;
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
	// aRefPendingRelease is 0 when this is being released the normal way (due to either the outer
	// function returning or a non-grouped Closure being deleted) and 1 when the script releases its
	// last direct reference to a grouped Closure.
	if (mRefCount)
		return mRefCount < 0; // Negative values mean a previous call (still on the stack) is deleting this.
	int circular_closures = 0;
	ULONG extra_references = 0;
	for (int i = 0; i < mVarCount; ++i)
		if (mVar[i].Type() == VAR_CONSTANT)
		{
			ASSERT(mVar[i].HasObject() && dynamic_cast<ObjectBase*>(mVar[i].Object())); // Any object in VAR_CONSTANT must derive from ObjectBase.
			auto obj = (ObjectBase *)mVar[i].Object();
			extra_references += obj->RefCount();
			++circular_closures;
		}
	// aRefPendingRelease == 0 && extra_references > 0: keep alive.
	// aRefPendingRelease == 1 && extra_references > 1: keep alive.
	// aRefPendingRelease == 1 && extra_references == 1: delete, because that one Closure is being deleted.
	if (extra_references > aRefPendingRelease)
		return false;
	--mRefCount; // Now that delete is certain, make this non-zero to prevent reentry.
	if (circular_closures)
	{
		// Any closure which is in a downvar and is not also an upvar (i.e. it is defined in this function,
		// not an outer one) has mRefCount == 0 at this point, meaning its only reference is the uncounted
		// one in mVar[].  In order to free the object properly, mRefCount needs to be restored to 1 prior
		// to Release(), which will be called by Var::Free() via ~FreeVars().
		for (int i = 0; i < mVarCount; ++i)
			if (mVar[i].IsDirectConstant())
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
	DEBUGGER_STACK_PUSH(_T("Gui"))
	ResultType result = OK;
	__int64 retval = 0;
	BOOL thread_used = FALSE;
	UINT_PTR event_info = g->EventInfo;
	
	for (MsgMonitorInstance inst (*this); inst.index < inst.count; ++inst.index)
	{
		MsgMonitorStruct &mon = mMonitor[inst.index];
		if (mon.msg != aMsg || mon.msg_type != aMsgType)
			continue;

		IObject *func = mon.is_method ? aGui->mEventSink : mon.func; // is_method == true implies the GUI has an event sink object.
		LPTSTR method_name = mon.is_method ? mon.method_name : nullptr;

		if (thread_used) // Re-initialize the thread.
		{
			InitNewThread(0, true, false);
			g->EventInfo = event_info;
		}
		
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
	DEBUGGER_STACK_POP()
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
	Object_Member(__New, Error__New, M_Error__New, IT_CALL, 0, 3),
	Object_Member(Show, Error_Show, 0, IT_CALL, 0, 2),
};

ObjectMember Object::sOSErrorMembers[]
{
	Object_Member(__New, Error__New, M_OSError__New, IT_CALL, 0, 3)
};



ObjectMember VarRef::sMembers[]
{
	Object_Member(__Value, __Value, 0, IT_SET | BIMF_UNSET_ARG_1)
};

ObjectMember PropRef::sMembers[]
{
	Object_Member(__Value, __Value, 0, IT_SET | BIMF_UNSET_ARG_1)
};


ObjectMember Object::sStructMembers[]
{
	Object_Method1(__Ref, 1, 1),
	Object_Member(Ptr, StructGet, M_Struct_Ptr, IT_GET),
	Object_Member(Size, StructGet, M_Struct_Size, IT_GET)
};

ObjectMember Object::sPtrMembers[]
{
	Object_Member(__Value, StructPtrInvoke, 0, IT_SET | BIMF_UNSET_ARG_1)
};

ObjectMember Object::sCArrayMembers[]
{
	Object_Member(__Item, CArrayItem, 0, IT_SET, 1, 1),
	Object_Member(Length, StructGet, M_CArray_Length, IT_GET)
};



struct ClassDef
{
	LPCTSTR name;
	Object **proto_var;
	ClassFactoryDef factory;
	ObjectMemberListType members;
	std::initializer_list<ClassDef> subclasses;
};

void DefineClasses(Object *aBaseClass, Object *aBaseProto, std::initializer_list<ClassDef> aClasses)
{
	for (auto &c : aClasses)
	{
		auto proto = (c.proto_var && *c.proto_var) ? *c.proto_var
			: Object::CreatePrototype(const_cast<LPTSTR>(c.name), aBaseProto, c.members);
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

	// This ensures GetStructInfo() has something to return for every Prototype or Object:
	sAnyPrototype->mFlags |= StructInfoInitialized | StructInfoLocked | NativeClassPrototype;

	// These methods correspond to global functions, as BuiltInMethod
	// only handles Objects, and these must handle primitive values.
	static const LPTSTR sFuncs[] = { _T("GetMethod"), _T("HasBase"), _T("HasMethod"), _T("HasProp"), _T("Props") };
	for (int i = 0; i < _countof(sFuncs); ++i)
		sAnyPrototype->DefineMethod(sFuncs[i], g_script.GetBuiltinObject(sFuncs[i]));
	auto prop = sAnyPrototype->DefineProperty(_T("Base"), false);
	prop->NoParamGet = prop->NoParamSet = true;
	prop->SetGetter(g_script.GetBuiltinObject(_T("ObjGetBase")));
	prop->SetSetter(g_script.GetBuiltinObject(_T("ObjSetBase")));
	
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
	Object::sObjectCall = Object::sClass->GetOwnPropMethod(_T("Call"));
	{
		// Each Class is suffixed with a pointer to the Prototype. This instructs ~Object() to release it.
		sClassPrototype->mFlags |= StructInfoInitialized | StructInfoLocked;
		auto &si = *(StructInfo*)(sClassPrototype + 1);
		si.nested_object_size = sizeof(Object*); // Pointer to Prototype.
		si.pointed_class = sClass; // Must be non-zero for ~Object().
		si.object_size = sizeof(Object); // For Class() without parameters.
		si.create = NewObject<Object>;
	}

	ObjectCtor no_ctor = nullptr;
	ObjectMemberListType no_members;

	DefineClasses(Object::sClass, Object::sPrototype, {
		{_T("Array"), &Array::sPrototype, NewObject<Array>, Array::sMembers},
		{_T("Buffer"), &BufferObject::sPrototype, NewObject<BufferObject>, BufferObject::sMembers, {
			{_T("ClipboardAll"), &ClipboardAll::sPrototype, NewObject<ClipboardAll>, ClipboardAll::sMembers}
		}},
		{_T("Class"), &Object::sClassPrototype, {Class_New, 0, 2, true}},
		{_T("Error"), &ErrorPrototype::Error, no_ctor, sErrorMembers, {
			{_T("MemoryError"), &ErrorPrototype::Memory},
			{_T("OSError"), &ErrorPrototype::OS, no_ctor, sOSErrorMembers},
			{_T("TargetError"), &ErrorPrototype::Target},
			{_T("TimeoutError"), &ErrorPrototype::Timeout},
			{_T("TypeError"), &ErrorPrototype::Type},
			{_T("UnsetError"), &ErrorPrototype::Unset, no_ctor, no_members, {
				{_T("MemberError"), &ErrorPrototype::Member, no_ctor, no_members, {
					{_T("PropertyError"), &ErrorPrototype::Property},
					{_T("MethodError"), &ErrorPrototype::Method}
				}},
				{_T("UnsetItemError"), &ErrorPrototype::UnsetItem}
			}},
			{_T("ValueError"), &ErrorPrototype::Value, no_ctor, no_members, {
				{_T("IndexError"), &ErrorPrototype::Index}
			}},
			{_T("ZeroDivisionError"), &ErrorPrototype::ZeroDivision}
		}},
		{_T("Func"), &Func::sPrototype, no_ctor, Func::sMembers, {
			{_T("BoundFunc"), &BoundFunc::sPrototype},
			{_T("Closure"), &Closure::sPrototype},
			{_T("Enumerator"), &EnumBase::sPrototype}
		}},
		{_T("Gui"), &GuiType::sPrototype, NewObject<GuiType>, {GuiType::sMembers, GuiType::sMemberCount}},
		{_T("InputHook"), &InputObject::sPrototype, NewObject<InputObject>, {InputObject::sMembers, InputObject::sMemberCount}},
		{_T("Map"), &Map::sPrototype, NewObject<Map>, Map::sMembers},
		{_T("Menu"), &UserMenu::sPrototype, NewObject<UserMenu>, {UserMenu::sMembers, UserMenu::sMemberCount}, {
			{_T("MenuBar"), &UserMenu::sBarPrototype, UserMenu::NewMenuBar}
		}},
		{_T("RegExMatchInfo"), &RegExMatchObject::sPrototype, no_ctor, RegExMatchObject::sMembers}
	});

	// Parameter counts are specified for static Call in the following classes
	// but not those using NewObject<> because the latter passes parameters on
	// to __New, which can be redefined by a subclass.  Specifying counts here
	// sets MinParams/MaxParams/IsVariadic appropriately and avoids the need to
	// validate aParamCount in each function, which reduces code size.
	// Note that the `this` parameter (the class itself) is counted.
	DefineClasses(anyClass, sAnyPrototype, {
		{_T("ComValue"), &sComValuePrototype, {ComValue_Call, 3, 4}, no_members, {
			{_T("ComObjArray"), &sComArrayPrototype, {ComObjArray_Call, 3, 10}},
			{_T("ComObject"), &sComObjectPrototype, {ComObject_Call, 2, 3}},
			{_T("ComValueRef"), &sComRefPrototype}
		}},
		{_T("Primitive"), &Object::sPrimitivePrototype, no_ctor, no_members, {
			{_T("Number"), &Object::sNumberPrototype, {BIF_Number, 2, 2}, no_members, {
				{_T("Float"), &Object::sFloatPrototype, {BIF_Float, 2, 2}},
				{_T("Integer"), &Object::sIntegerPrototype, {BIF_Integer, 2, 2}}
			}},
			{_T("String"), &Object::sStringPrototype, {BIF_String, 2, 2}}
		}},
		{_T("Module"), &ScriptModule::sPrototype},
		{_T("PropRef"), &PropRef::sPrototype, {PropRef_Call, 3, 3}, PropRef::sMembers},
		{_T("Struct"), &sStructPrototype, NewStruct, sStructMembers},
		{_T("VarRef"), &sVarRefPrototype, no_ctor, VarRef::sMembers}
	});

	sStructClass = (Object*)g_script.FindGlobalVar(_T("Struct"), 6)->Object();
	sStructClass->DefinePrototypeGetter();
	sStructClass->DefineMethod(_T("At"), new BuiltInFunc {_T("Struct.At"), StructClass_At, 2, 2});
	prop = sStructClass->DefineProperty(_T("__Item"));
	prop->SetGetter(new BuiltInFunc{ _T("Struct.__Item"), StructClass_Item, 2, 2 });
	prop->NoEnumGet = true;
	prop = sStructClass->DefineProperty(_T("Ptr"));
	prop->SetMethod(new BuiltInFunc{ _T("Struct.Ptr"), StructClass_Ptr, 1, 1, true, (void*)1 });
	prop->SetGetter(new BuiltInFunc{ _T("Struct.Ptr"), StructClass_Ptr, 1, 1, false, (void*)0 });
	prop->NoEnumGet = true;
	prop->NoParamGet = true;

	sPtrPrototype = CreatePrototype(_T("Struct") STRUCT_PTR_CLASS_SUFFIX, sStructPrototype, sPtrMembers, _countof(sPtrMembers));
	sPtrClass = CreateClass(sPtrPrototype, sStructClass);
	{
		sPtrPrototype->mFlags &= ~NativeClassPrototype; // Allow Struct.Call to construct Ptr.
		sPtrPrototype->mFlags |= StructInfoInitialized | StructInfoLocked;
		auto &tp = *sPtrPrototype->DefineTypedProperty(_T("Value"));
		tp.type = MdType::IntPtr;
		tp.class_object = nullptr;
		tp.pointed_proto = nullptr;
		tp.data_offset = 0;
		auto &psi = *(StructInfo*)(sPtrPrototype + 1);
		psi.align = psi.size = sizeof(void*);
		psi.pointed_class = sStructClass; // This is not a counted reference.
		psi.nested_object_size = sizeof(Object*);
		psi.object_size = sizeof(Object);
		auto &ssi = *(StructInfo*)(sStructPrototype + 1);
		ssi.align = 1;
		ssi.pointer_class = sPtrClass;
		ssi.object_size = sizeof(Object);
		++sPtrClass->mRefCount; // For correctness, though it should never be released.
		sStructPrototype->mFlags |= StructInfoInitialized | StructInfoLocked;
	}

	sCArrayPrototype = CreatePrototype(_T("Struct.Array"), sStructPrototype, sCArrayMembers, _countof(sCArrayMembers));
	sCArrayPrototype->mFlags &= ~NativeClassPrototype;
	sCArrayPrototype->mFlags |= StructInfoLocked;
	sCArrayClass = CreateClass(sCArrayPrototype, sStructClass);
	sStructClass->DefineClass(_T("Array"), sCArrayClass, true);

	LPTSTR const type_names[]{ _T("Float32"), _T("Float64"), _T("Int16"), _T("Int32"), _T("Int64"), _T("Int8"), _T("IntPtr"), _T("UInt16"), _T("UInt32"), _T("UInt8") };
	MdType const type_codes[]{ MdType::Float32, MdType::Float64, MdType::Int16, MdType::Int32, MdType::Int64, MdType::Int8, MdType::IntPtr, MdType::UInt16, MdType::UInt32, MdType::UInt8 };
	UCHAR const type_dllcall[]{ DLL_ARG_FLOAT, DLL_ARG_DOUBLE, DLL_ARG_SHORT, DLL_ARG_INT, DLL_ARG_INT64, DLL_ARG_CHAR, Exp32or64(DLL_ARG_INT,DLL_ARG_INT64), DLL_ARG_SHORT, DLL_ARG_INT, DLL_ARG_CHAR};
	for (int i = 0; i < _countof(type_names); ++i)
	{
		auto p = CreatePrototype(type_names[i], sStructPrototype);
		auto si = (StructInfo*)(p + 1);
		p->mFlags |= StructInfoInitialized | StructInfoLocked;
		si->object_size = sizeof(Object);
		si->native_type = type_codes[i];
		si->dllcall_type = type_dllcall[i];
		si->is_unsigned = type_names[i][0] == 'U';
		si->align = si->size = TypeSize(type_codes[i]);
		auto tp = p->DefineTypedProperty(_T("__Value"));
		tp->type = type_codes[i];
		tp->class_object = nullptr;
		tp->pointed_proto = nullptr;
		tp->data_offset = 0;
		auto c = CreateClass(type_names[i], sStructClass, p, nullptr);
		//CreatePtrClass(c, p, si);
		sPrimitiveClass[(int)type_codes[i] - 1] = c;
	}

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
Object *Object::sStructPrototype, *Object::sPtrPrototype, *Object::sCArrayPrototype;
Object *Array::sPrototype;
Object *Map::sPrototype;

Object *Object::sClass;
Object *Object::sStructClass, *Object::sPtrClass, *Object::sCArrayClass;
Object *Object::sPrimitiveClass[(int)MdType::LastSupportedPropertyType];

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
Object *PropRef::sPrototype;
Object *Object::sComObjectPrototype, *Object::sComValuePrototype, *Object::sComArrayPrototype, *Object::sComRefPrototype;

IObject *Object::sObjectCall;


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



void Object::DefineClass(name_t aName, Object *aClass, bool aIsStructPtrClass)
{
	auto prop = DefineProperty(aName);

	ExprTokenType values[] { aClass, aName }, *param[] { values, values + 1 };

	auto info = SimpleHeap::Alloc<NestedClassInfo>();
	info->class_object = aClass;
	info->constructed = aIsStructPtrClass;
	aClass->AddRef();

	auto get = new BuiltInFunc { _T(""), Class_GetNestedClass, 1, 1, false, info };
	prop->NoParamGet = prop->NoParamSet = true;
	prop->SetGetter(get);

	auto call = new BuiltInFunc { _T(""), Class_CallNestedClass, 1, 1, true, info };
	prop->SetMethod(call);
}


void Object::DefinePrototypeGetter()
{
	static BuiltInFunc sClassPrototypeGet{ _T("Class.Prototype.Get"), Class_Prototype, 1, 1 };

	auto prop = DefineProperty(_T("Prototype"));
	prop->SetGetter(&sClassPrototypeGet);
	prop->NoParamGet = true;
}


BIF_DECL(Class_Prototype)
{
	auto obj0 = ParamIndexToObject(0);
	auto p = obj0 && obj0->IsOfType(Object::sPrototype) ? ((Object*)obj0)->ClassGetPrototype() : nullptr;
	if (!p)
		_o_throw_type(_T("Class"), *aParam[0]);
	p->AddRef();
	_o_return(p);
}


BIF_DECL(Class_GetNestedClass)
{
	auto info = (NestedClassInfo *)aResultToken.callee_id;
	auto cls = info->class_object;
	cls->AddRef();
	if (info->constructed)
		_f_return(cls);
	info->constructed = true;
	cls->CallInitNew(aResultToken, nullptr, 0);
}


BIF_DECL(Class_CallNestedClass)
{
	auto info = (NestedClassInfo *)aResultToken.callee_id;
	auto cls = info->class_object;
	if (!info->constructed)
	{
		info->constructed = true;
		cls->AddRef(); // Necessary because CallInitNew() calls Release() on failure/exit.
		if (cls->CallInitNew(aResultToken, nullptr, 0) != OK) // FAIL or EXIT
			return;
		cls->Release();
		aResultToken.InitResult(aResultToken.buf);
	}
	else
		aResultToken.InitInvokeRetVal();
	cls->Invoke(aResultToken, IT_CALL, nullptr, ExprTokenType { cls }, aParam + 1, aParamCount - 1);
}


BIF_DECL(Class_New)
{
	// For backward-compatibility, Class() is the same as (Object.Call)(Class).
	// Class(unset) would have thrown ERR_TOO_MANY_PARAMS, so is exempted from this.
	if (aParamCount == 1)
	{
		aResultToken.callee_id = Object::sPrototype;
		return Object::NewInstance(aResultToken, aParam, aParamCount);
	}

	// aParam[0] is implicit and mandatory, as this is a method.  Usually it should be Class itself,
	// but might be something else if the script explicitly calls (Class.Call)(this) or extends Class
	// itself (which is not the same as an instance of Class, as that derives from Class.Prototype).
	++aParam, --aParamCount; // Exclude "Class" itself.

	// Get the optional name and class parameters.
	LPTSTR name = _T("");
	IObject *obj0 = ParamIndexToObject(0);
	if (!obj0)
	{
		name = ParamIndexToString(0, _f_retval_buf);
		++aParam, --aParamCount;
	}
	Object *base_class = obj0 ? dynamic_cast<Object *>(obj0) : ParamIndexIsOmitted(0) ? Object::sClass : dynamic_cast<Object *>(ParamIndexToObject(0));
	Object *base_proto = base_class ? base_class->ClassGetPrototype() : nullptr;
	if (!base_proto)
		return (void)aResultToken.ParamError(obj0 ? 0 : 1, aParam[0], _T("Class"));
	if (aParamCount)
		++aParam, --aParamCount;

	auto proto = Object::CreatePrototype(name, base_proto);
	auto class_obj = Object::CreateClass(proto, base_class);
	proto->Release();

	// Don't call any inherited __Init, since that would reinitialize static variables and duplicate
	// any typed properties defined by that one class.  This either releases or returns class_obj:
	class_obj->CallNew(aResultToken, aParam, aParamCount, ExprTokenType(class_obj));
}
