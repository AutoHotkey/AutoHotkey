#pragma once

class InputObject : public ObjectBase
{
	input_type input;

public:
	IObject *onEnd;

	InputObject() : onEnd(NULL)
	{
		input.ScriptObject = this;
	}

	~InputObject()
	{
		if (onEnd)
			onEnd->Release();
	}

	ResultType Setup(LPTSTR aOptions, LPTSTR aEndKeys, LPTSTR aMatchList, size_t aMatchList_length);
	ResultType KeyOpt(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
	IObject_Type_Impl("InputHook")
};
