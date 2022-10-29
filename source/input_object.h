#pragma once

class InputObject : public Object
{
	input_type input;
	
	FResult get_On(IObject *&aRetVal, IObject *&aOn);
	FResult set_On(ExprTokenType &aValue, IObject *&aOn, int aValidParamCount);
	
public:
	IObject *onEnd = nullptr, *onKeyDown = nullptr, *onChar = nullptr, *onKeyUp = nullptr;

	static Object *sPrototype;
	static ObjectMemberMd sMembers[];
	static int sMemberCount;

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

	static Object *Create();
	
	FResult __New(optl<StrArg> aOptions, optl<StrArg> aEndKeys, optl<StrArg> aMatchList);
	
	FResult Start();
	FResult Wait(optl<double> aMaxTime, StrRet &aRetVal);
	FResult Stop();
	FResult KeyOpt(StrArg aKeys, StrArg aKeyOptions);
	
	FResult get_InProgress(BOOL &aRetVal)  { aRetVal = input.InProgress(); return OK; }
	
	FResult get_EndReason(StrRet &aRetVal);
	FResult get_EndKey(StrRet &aRetVal);
	FResult get_EndMods(StrRet &aRetVal);
	FResult get_Input(StrRet &aRetVal);
	FResult get_Match(StrRet &aRetVal);
	
	#define ONX_OPTION(X, N) \
		FResult get_On##X(IObject *&aRetVal) { \
			return get_On(aRetVal, on##X); \
		} \
		FResult set_On##X(ExprTokenType &aValue) { \
			return set_On(aValue, on##X, N); \
		}
	ONX_OPTION(End, 1);
	ONX_OPTION(Char, 2);
	ONX_OPTION(KeyDown, 3);
	ONX_OPTION(KeyUp, 3);
	#undef ONX_OPTION
	
	#define BOOL_OPTION(OPT) \
		FResult get_##OPT(BOOL &aRetVal) { \
			aRetVal = input.OPT; \
			return OK; \
		} \
		FResult set_##OPT(BOOL aValue) { \
			input.OPT = aValue; \
			return OK; \
		}
	BOOL_OPTION(BackspaceIsUndo);
	BOOL_OPTION(CaseSensitive);
	BOOL_OPTION(FindAnywhere);
	BOOL_OPTION(NotifyNonText);
	BOOL_OPTION(VisibleNonText);
	BOOL_OPTION(VisibleText);
	#undef BOOL_OPTION
	
	FResult get_MinSendLevel(int &aRetVal) { aRetVal = input.MinSendLevel; return OK; }
	FResult set_MinSendLevel(int aValue);
	
	FResult get_Timeout(double &aRetVal) { aRetVal = input.Timeout / 1000.0; return OK; }
	FResult set_Timeout(double aValue);
};
