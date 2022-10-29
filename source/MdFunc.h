#pragma once

#include "script.h"
#include "MdType.h"


void TypedPtrToToken(MdType aType, void *aPtr, ExprTokenType &aToken);
ResultType FResultToError(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount, FResult aResult);


class MdFunc : public NativeFunc
{
	void *mMcFunc; // Pointer to native function.
	Object *mPrototype; // Prototype object used for type checking; a non-null value implies mMcFunc is a member function.
	MdType *mArgType; // Sequence of native arg types and modifiers.
	MdType mRetType; // Type of native return value (not necessarily the script return value).
	UINT8 mMaxResultTokens; // Number of ResultTokens that might be allocated for conversions.
	UINT8 mArgSlots; // Number of DWORD_PTRs needed for the parameter list.
	bool mThisCall;

public:
	MdFunc(LPCTSTR aName, void *aMcFunc, MdType aRetType, MdType *aArg, UINT aArgSize, Object *aPrototype = nullptr);

	bool Call(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount) override;
	bool ArgIsOutputVar(int aIndex) override;
	bool ArgIsOptional(int aIndex) override;
};