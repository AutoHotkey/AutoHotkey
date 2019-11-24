#pragma once

class InputObject : public ObjectBase
{
	input_type input;

public:
	IObject *onEnd, *onKeyDown, *onChar, *onKeyUp;

	InputObject() : onEnd(NULL), onKeyDown(NULL), onChar(NULL), onKeyUp(NULL)
	{
		input.ScriptObject = this;
	}

	~InputObject()
	{
		if (onEnd)
			onEnd->Release();
		if (onKeyDown)
			onKeyDown->Release();
		if (onChar)
			onChar->Release();
		if (onKeyUp)
			onKeyUp->Release();
	}

	ResultType Setup(LPTSTR aOptions, LPTSTR aEndKeys, LPTSTR aMatchList, size_t aMatchList_length);
	ResultType KeyOpt(ExprTokenType &aResultToken, ExprTokenType *aParam[], int aParamCount);

	ResultType STDMETHODCALLTYPE Invoke(ExprTokenType &aResultToken, ExprTokenType &aThisToken, int aFlags, ExprTokenType *aParam[], int aParamCount);
	IObject_Type_Impl("InputHook")
};
