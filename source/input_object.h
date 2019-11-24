#pragma once

class InputObject : public Object
{
	input_type input;

	enum Members
	{
		M_Start,
		M_Wait,
		M_Stop,
		//M_KeyOpt,
		P_Input,
		P_InProgress,
		P_EndReason,
		P_EndKey,
		P_EndMods,
		P_Match,
		P_OnChar,
		P_OnKeyDown,
		P_OnKeyUp,
		P_OnEnd,
		P_MinSendLevel,
		P_Timeout,
		P_BackspaceIsUndo,
		P_CaseSensitive,
		P_FindAnywhere,
		P_VisibleText,
		P_VisibleNonText,
		P_NotifyNonText
	};
	
public:
	IObject *onEnd = nullptr, *onKeyDown = nullptr, *onChar = nullptr, *onKeyUp = nullptr;

	static Object *sPrototype;
	static ObjectMember sMembers[];

	InputObject();
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
	ResultType KeyOpt(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType OnX(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType BoolOpt(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
	ResultType Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount);
};
