#pragma once

class InputObject : public ObjectBase
{
	input_type input;

public:
	IObject *onEnd, *onKeyDown, *onChar;

	InputObject() : onEnd(NULL), onKeyDown(NULL), onChar(NULL)
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
	}

	ResultType Setup(LPTSTR aOptions, LPTSTR aEndKeys, LPTSTR aMatchList, size_t aMatchList_length);
	ResultType KeyOpt(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount);

	ResultType Invoke(IObject_Invoke_PARAMS_DECL);
	IObject_Type_Impl("InputHook")
};
