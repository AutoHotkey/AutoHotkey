/*
AutoHotkey

Copyright 2003-2009 Chris Mallett (support@autohotkey.com)

This program is free software; you can redistribute it and/or
modify it under the terms of the GNU General Public License
as published by the Free Software Foundation; either version 2
of the License, or (at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.
*/

#include "stdafx.h" // pre-compiled headers
#include <Shlwapi.h>
#include <uxtheme.h>
#include "script.h"
#include "globaldata.h" // for a lot of things
#include "application.h" // for MsgSleep()
#include "window.h" // for SetForegroundWindowEx()
#include "qmath.h" // for qmathLog()
#include "script_func_impl.h"


// Gui methods use this macro either if they require the Gui to have a window,
// or if they aren't safe to use after Dispose() is called.
#define GUI_MUST_HAVE_HWND if (!mHwnd) return GuiNoWindowError()
// Using a noinline function reduces code by a few hundred bytes across all methods.
// Returning something like FR_E_WIN32(ERROR_INVALID_WINDOW_HANDLE) instead is about
// the same, so it seems better to use the clearer error message.
__declspec(noinline) FResult GuiNoWindowError() { return FError(ERR_GUI_NO_WINDOW); }

#define CTRL_THROW_IF_DESTROYED if (!hwnd) return ControlDestroyedError()
__declspec(noinline) FResult ControlDestroyedError() { return FError(_T("The control is destroyed.")); }


// Window class atom used by Guis.
static ATOM sGuiWinClass;


LPTSTR GuiControlType::sTypeNames[GUI_CONTROL_TYPE_COUNT] = { GUI_CONTROL_TYPE_NAMES };

GuiControls GuiControlType::ConvertTypeName(LPCTSTR aTypeName)
{
	for (int i = 1; i < _countof(sTypeNames); ++i) // Start at 1 because 0 is GUI_CONTROL_INVALID.
		if (!_tcsicmp(aTypeName, sTypeNames[i]))
			return (GuiControls)i;
	if (!_tcsicmp(aTypeName, _T("DropDownList"))) return GUI_CONTROL_DROPDOWNLIST;
	if (!_tcsicmp(aTypeName, _T("Picture"))) return GUI_CONTROL_PIC;
	return GUI_CONTROL_INVALID;
}

LPTSTR GuiControlType::GetTypeName()
{
	if (type == GUI_CONTROL_TAB)
	{
		// Tab2 and Tab3 are changed into Tab at an early stage.
		if (attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR)
			return sTypeNames[GUI_CONTROL_TAB2];
		if (GetProp(hwnd, _T("ahk_dlg")))
			return sTypeNames[GUI_CONTROL_TAB3];
	}
	return sTypeNames[type];
}

GuiControlType::TypeAttribs GuiControlType::TypeHasAttrib(GuiControls aType, TypeAttribs aAttrib)
{
	static TypeAttribs sAttrib[] = { 0,
		/*Text*/       TYPE_STATICBACK | TYPE_SUPPORTS_BGTRANS | TYPE_NO_SUBMIT,
		/*Pic*/        TYPE_STATICBACK | TYPE_SUPPORTS_BGTRANS | TYPE_NO_SUBMIT | TYPE_HAS_NO_TEXT | TYPE_RESERVE_UNION,
		/*GroupBox*/   TYPE_STATICBACK | TYPE_SUPPORTS_BGTRANS | TYPE_NO_SUBMIT, // Background affects only the label.
		/*Button*/     TYPE_MSGBKCOLOR | TYPE_SUPPORTS_BGTRANS | TYPE_NO_SUBMIT, // Background affects only the button's border, on Windows 10.
		/*CheckBox*/   TYPE_STATICBACK,
		/*Radio*/      TYPE_STATICBACK | TYPE_NO_SUBMIT, // No-submit: Radio controls are handled separately, by group.
		/*DDL*/        TYPE_MSGBKCOLOR | TYPE_HAS_ITEMS, // Background affects only the main control, and requires -Theme.
		/*ComboBox*/   TYPE_MSGBKCOLOR | TYPE_HAS_ITEMS, // Background affects only the main control.
		/*ListBox*/    TYPE_MSGBKCOLOR | TYPE_HAS_ITEMS,
		/*ListView*/   TYPE_SETBKCOLOR | TYPE_NO_SUBMIT | TYPE_RESERVE_UNION | TYPE_HAS_ITEMS,
		/*TreeView*/   TYPE_SETBKCOLOR | TYPE_NO_SUBMIT,
		/*Edit*/       TYPE_MSGBKCOLOR,
		/*DateTime*/   0,
		/*MonthCal*/   0,
		/*Hotkey*/     0,
		/*UpDown*/     TYPE_HAS_NO_TEXT,
		/*Slider*/     TYPE_STATICBACK | TYPE_HAS_NO_TEXT,
		/*Progress*/   TYPE_SETBKCOLOR | TYPE_HAS_NO_TEXT | TYPE_NO_SUBMIT,
		/*Tab*/        TYPE_STATICBACK | TYPE_HAS_ITEMS,
		/*Tab2*/       TYPE_STATICBACK | TYPE_HAS_ITEMS, // Only used prior to control construction; at least TYPE_HAS_ITEMS is needed here.
		/*Tab3*/       TYPE_STATICBACK | TYPE_HAS_ITEMS, // As above.
		/*ActiveX*/    TYPE_HAS_NO_TEXT | TYPE_NO_SUBMIT | TYPE_RESERVE_UNION,
		/*Link*/       TYPE_STATICBACK | TYPE_NO_SUBMIT,
		/*Custom*/     TYPE_MSGBKCOLOR | TYPE_NO_SUBMIT, // Custom controls *may* use WM_CTLCOLOR.
		/*StatusBar*/  TYPE_SETBKCOLOR | TYPE_NO_SUBMIT,
	};
	return sAttrib[aType] & aAttrib;
}

UCHAR **ConstructEventSupportArray()
{
	// This is mostly static data, but I haven't figured out how to construct it with just
	// an array initialiser (other than by using a fixed size for the sub-array, which
	// wastes space due to ListView and TreeView having many more events than other types).
	static UCHAR *raises[GUI_CONTROL_TYPE_COUNT];
	// Not necessary since the array is static and therefore initialised to 0 (NULL):
	//#define RAISES_NONE(ctrl) \
	//	raises[ctrl] = NULL;
	#define RAISES(ctrl, ...) { \
		static UCHAR events[] = { __VA_ARGS__, 0 }; \
		raises[ctrl] = events; \
	}

	RAISES(GUI_CONTROL_TEXT,         GUI_EVENT_CLICK, GUI_EVENT_DBLCLK)
	RAISES(GUI_CONTROL_PIC,          GUI_EVENT_CLICK, GUI_EVENT_DBLCLK)
	//RAISES_NONE(GUI_CONTROL_GROUPBOX)
	RAISES(GUI_CONTROL_BUTTON,       GUI_EVENT_CLICK, GUI_EVENT_DBLCLK, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_CHECKBOX,     GUI_EVENT_CLICK, GUI_EVENT_DBLCLK, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_RADIO,        GUI_EVENT_CLICK, GUI_EVENT_DBLCLK, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_DROPDOWNLIST, GUI_EVENT_CHANGE, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_COMBOBOX,     GUI_EVENT_CHANGE, GUI_EVENT_DBLCLK, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_LISTBOX,      GUI_EVENT_CHANGE, GUI_EVENT_DBLCLK, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_LISTVIEW,     GUI_EVENT_DBLCLK, GUI_EVENT_COLCLK, GUI_EVENT_CLICK, GUI_EVENT_ITEMFOCUS, GUI_EVENT_ITEMSELECT, GUI_EVENT_ITEMCHECK, GUI_EVENT_ITEMEDIT, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_TREEVIEW,     GUI_EVENT_ITEMSELECT, GUI_EVENT_DBLCLK, GUI_EVENT_CLICK, GUI_EVENT_ITEMEXPAND, GUI_EVENT_ITEMCHECK, GUI_EVENT_ITEMEDIT, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_EDIT,         GUI_EVENT_CHANGE, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_DATETIME,     GUI_EVENT_CHANGE, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS)
	RAISES(GUI_CONTROL_MONTHCAL,     GUI_EVENT_CHANGE)
	RAISES(GUI_CONTROL_HOTKEY,       GUI_EVENT_CHANGE)
	RAISES(GUI_CONTROL_UPDOWN,       GUI_EVENT_CHANGE)
	RAISES(GUI_CONTROL_SLIDER,       GUI_EVENT_CHANGE)
	//RAISES_NONE(GUI_CONTROL_PROGRESS)
	RAISES(GUI_CONTROL_TAB,          GUI_EVENT_CHANGE)
	//RAISES(GUI_CONTROL_TAB2,         ) // Not used since it is changed to TAB at an early stage.
	//RAISES(GUI_CONTROL_TAB3,         )
	//RAISES_NONE(GUI_CONTROL_ACTIVEX)
	RAISES(GUI_CONTROL_LINK,         GUI_EVENT_CLICK)
	//RAISES_NONE(GUI_CONTROL_CUSTOM)
	RAISES(GUI_CONTROL_STATUSBAR,    GUI_EVENT_CLICK, GUI_EVENT_DBLCLK)

	//#undef RAISES_NONE
	#undef RAISES

	return raises;
}

UCHAR **GuiControlType::sRaisesEvents = ConstructEventSupportArray();

bool GuiControlType::SupportsEvent(GuiEventType aEvent)
{
	if (UCHAR *this_raises = sRaisesEvents[type])
	{
		for (;; ++this_raises)
		{
			if (*this_raises == aEvent)
				return true;
			if (!*this_raises) // Checked after aEvent, as it will often be the first item in the array.
				break;
		}
	}
	if (aEvent == GUI_EVENT_CONTEXTMENU)
		return true; // All controls implicitly support this.
	return false;
}


// Helper function used to convert a token to an Array object.
static Array* TokenToArray(ExprTokenType &token)
{
	return dynamic_cast<Array *>(TokenToObject(token));
}


// Enumerator for GuiType objects.
ResultType GuiType::GetEnumItem(UINT &aIndex, Var *aOutputVar1, Var *aOutputVar2, int aVarCount)
{
	if (aIndex >= mControlCount) // Use >= vs. == in case the Gui was destroyed.
		return CONDITION_FALSE;
	GuiControlType* ctrl = mControl[aIndex];
	if (aVarCount == 1) // `for x in myGui` or `[myGui*]`
	{
		aOutputVar2 = aOutputVar1; // Return the more useful value in single-var mode: the control object.
		aOutputVar1 = nullptr;
	}
	if (aOutputVar1)
		aOutputVar1->AssignHWND(ctrl->hwnd);
	if (aOutputVar2)
		aOutputVar2->Assign(ctrl);
	return CONDITION_TRUE;
}


FResult GuiType::__Enum(optl<int> aVarCount, IObject *&aRetVal)
{
	GUI_MUST_HAVE_HWND;
	aRetVal = new IndexEnumerator(this, aVarCount.value_or(0)
		, static_cast<IndexEnumerator::Callback>(&GuiType::GetEnumItem));
	return OK;
}


ObjectMemberMd GuiType::sMembers[] =
{
	md_member(GuiType, __Enum, CALL, (In_Opt, Int32, VarCount), (Ret, Object, RetVal)),
	md_member(GuiType, __New, CALL, (In_Opt, String, Options), (In_Opt, String, Title), (In_Opt, Object, EventObj)),
	
	md_member(GuiType, Add, CALL, (In, String, ControlType), (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddActiveX, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddButton, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddCheckBox, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddComboBox, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddCustom, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddDateTime, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddDDL, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddDropDownList, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddEdit, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddGroupBox, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddHotkey, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddLink, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddListBox, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddListView, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddMonthCal, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddPic, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddPicture, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddProgress, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddRadio, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddSlider, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddStatusBar, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddTab, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddTab2, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddTab3, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddText, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddTreeView, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	md_member(GuiType, AddUpDown, CALL, (In_Opt, String, Options), (In_Opt, Variant, Content), (Ret, Object, RetVal)),
	
	md_member(GuiType, Destroy, CALL, md_arg_none),
	md_member(GuiType, Flash, CALL, (In_Opt, Bool32, Blink)),
	md_member(GuiType, GetClientPos, CALL, (Out_Opt, Int32, X), (Out_Opt, Int32, Y), (Out_Opt, Int32, Width), (Out_Opt, Int32, Height)),
	md_member(GuiType, GetPos, CALL, (Out_Opt, Int32, X), (Out_Opt, Int32, Y), (Out_Opt, Int32, Width), (Out_Opt, Int32, Height)),
	md_member(GuiType, Hide, CALL, md_arg_none),
	md_member(GuiType, Maximize, CALL, md_arg_none),
	md_member(GuiType, Minimize, CALL, md_arg_none),
	md_member(GuiType, Move, CALL, (In_Opt, Int32, X), (In_Opt, Int32, Y), (In_Opt, Int32, Width), (In_Opt, Int32, Height)),
	md_member(GuiType, OnEvent, CALL, (In, String, EventName), (In, Variant, Callback), (In_Opt, Int32, AddRemove)),
	md_member(GuiType, OnMessage, CALL, (In, UInt32, Number), (In, Variant, Callback), (In_Opt, Int32, AddRemove)),
	md_member(GuiType, Opt, CALL, (In, String, Options)),
	md_member(GuiType, Restore, CALL, md_arg_none),
	md_member(GuiType, SetFont, CALL, (In_Opt, String, Options), (In_Opt, String, FontName)),
	md_member(GuiType, Show, CALL, (In_Opt, String, Options)),
	md_member(GuiType, Submit, CALL, (In_Opt, Bool32, Hide), (Ret, Object, RetVal)),
	
	md_member		(GuiType, __Item, GET, (In, Variant, Index), (Ret, Object, RetVal)),
	md_property_get	(GuiType, Hwnd, UInt32),
	md_property		(GuiType, Title, String),
	md_property		(GuiType, Name, String),
	md_property_get	(GuiType, FocusedCtrl, Object),
	md_property		(GuiType, BackColor, Variant),
	md_property		(GuiType, MarginX, Int32),
	md_property		(GuiType, MarginY, Int32),
	md_property		(GuiType, MenuBar, Variant)
};

int GuiType::sMemberCount = _countof(sMembers);


FResult GuiType::Add(StrArg aCtrlType, optl<StrArg> aOptions, ExprTokenType *aContent, IObject *&aRetVal)
{
	auto ctrl_type = GuiControlType::ConvertTypeName(aCtrlType);
	if (ctrl_type == GUI_CONTROL_INVALID)
		return FValueError(_T("Invalid control type."), aCtrlType);
	auto fr = AddControl(aOptions, aContent, aRetVal, ctrl_type);
	return fr;
}


FResult GuiType::AddControl(optl<StrArg> aOptions, ExprTokenType *aContent, IObject *&aRetVal, GuiControls ctrl_type)
{
	auto options = aOptions.value_or_empty();
	TCHAR text_buf[MAX_NUMBER_SIZE], *text = _T("");
	Array *text_obj = nullptr;
	if (aContent)
	{
		auto temp_obj = TokenToObject(*aContent);
		if (GuiControlType::TypeHasAttrib(ctrl_type, GuiControlType::TYPE_HAS_ITEMS)) // Need an Array.
		{
			if (  !(text_obj = dynamic_cast<Array*>(temp_obj))  )
				return FTypeError(_T("Array"), *aContent);
		}
		else // Need a String.
		{
			if (temp_obj)
				return FTypeError(_T("String"), *aContent);
			text = TokenToString(*aContent, text_buf);
		}
	}
	GUI_MUST_HAVE_HWND;
	GuiControlType* pcontrol;
	ResultType result = AddControl(ctrl_type, options, text, pcontrol, text_obj);
	if (pcontrol)
	{
		pcontrol->AddRef();
		aRetVal = pcontrol;
	}
	return result ? OK : FR_FAIL;
}


FResult GuiType::__New(optl<StrArg> aOptions, optl<StrArg> aTitle, optl<IObject*> aEventObj)
{
	if (mHwnd || mDisposed)
		return FError(ERR_INVALID_USAGE);

	bool set_last_found_window = false;
	ToggleValueType own_dialogs = TOGGLE_INVALID;
	if (!ParseOptions(aOptions.value_or_empty(), set_last_found_window, own_dialogs))
		return FR_FAIL;

	mControl = (GuiControlType **)malloc(GUI_CONTROL_BLOCK_SIZE * sizeof(GuiControlType*));
	if (!mControl)
		return FR_E_OUTOFMEM;
	mControlCapacity = GUI_CONTROL_BLOCK_SIZE;

	if (aEventObj.has_value())
	{
		// The caller specified an object to use as event sink.
		mEventSink = aEventObj.value();
		if (mEventSink != this) // Primarily for custom GUI classes: prevent a circular reference.
			mEventSink->AddRef();
	}
	
	LPCTSTR title = aTitle.has_value() ? aTitle.value() : g_script.DefaultDialogTitle();

	// Create the Gui, now that we're past all other failure points.
	auto fr = Create(title);
	if (fr != OK)
		return fr;

	if (set_last_found_window)
		g->hWndLastUsed = mHwnd;
	SetOwnDialogs(own_dialogs);

	// Successful creation - add the Gui to the global list of Guis and return it
	AddGuiToList(this);
	return OK;
}


FResult GuiType::get_Hwnd(UINT &aRetVal)
{
	// Throwing if Destroy() has been called provides a clearer error message than
	// returning 0 when the Gui is being passed to a WinTitle parameter, and ensures
	// that it won't cause silent failures, such as with DllCall.
	GUI_MUST_HAVE_HWND;
	aRetVal = (UINT)(UINT_PTR)mHwnd;
	return OK;
}


FResult GuiType::get_Title(StrRet &aRetVal)
{
	GUI_MUST_HAVE_HWND;
	int length = GetWindowTextLength(mHwnd);
	auto buf = aRetVal.Alloc(length);
	if (!buf)
		return FR_E_OUTOFMEM;
	GetWindowText(mHwnd, buf, length+1);
	return OK;
}


FResult GuiType::set_Title(StrArg aValue)
{
	GUI_MUST_HAVE_HWND;
	SetWindowText(mHwnd, aValue);
	return OK;
}


FResult GuiType::get___Item(ExprTokenType &aIndex, IObject *&aRetVal)
{
	GUI_MUST_HAVE_HWND; // Seems clearer than relying on the error below.
	GuiControlType* ctrl = NULL;
	if (TokenIsPureNumeric(aIndex) == SYM_INTEGER)
	{
		HWND hWnd = (HWND)(UINT_PTR)TokenToInt64(aIndex);
		ctrl = FindControl(hWnd);
	}
	else
	{
		TCHAR buf[MAX_NUMBER_SIZE];
		GuiIndexType u = FindControl(TokenToString(aIndex, buf));
		if (u < mControlCount)
			ctrl = mControl[u];
	}
	if (!ctrl)
		return FError(_T("The specified control does not exist."), nullptr, ErrorPrototype::UnsetItem);

	ctrl->AddRef();
	aRetVal = ctrl;
	return OK;
}


FResult GuiType::get_FocusedCtrl(IObject *&aRetVal)
{
	GUI_MUST_HAVE_HWND;
	HWND hwnd = GetFocus();
	aRetVal = hwnd ? FindControl(hwnd) : nullptr;
	if (aRetVal)
		aRetVal->AddRef();
	return OK;
}


FResult GuiType::get_Margin(int &aRetVal, int &aMargin)
{
	if (aMargin == COORD_UNSPECIFIED)
	{
		// Avoid returning an arbitrary value such as COORD_UNSPECIFIED or -1 (Unscale(COORD_UNSPECIFIED)).
		// Instead, set the default margins if MarginX or MarginY is queried without having first been set,
		// prior to adding any controls.
		SetDefaultMargins();
	}
	aRetVal = Unscale(aMargin);
	return OK;
}


FResult GuiType::get_BackColor(ResultToken &aResultToken)
{
	GUI_MUST_HAVE_HWND;
	if (mBackgroundColorWin == CLR_DEFAULT)
		return OK; // "" is the default return value.
	_sntprintf(_f_retval_buf, _f_retval_buf_size, _T("%06X"), bgr_to_rgb(mBackgroundColorWin));
	aResultToken.SetValue(_f_retval_buf);
	return OK;
}


FResult GuiType::set_BackColor(ExprTokenType &aValue)
{
	GUI_MUST_HAVE_HWND;
	
	COLORREF new_color;
	if (!ColorToBGR(aValue, new_color))
		return FValueError(ERR_INVALID_VALUE);
	
	// AssignColor() takes care of deleting old brush, etc.
	AssignColor(new_color, mBackgroundColorWin, mBackgroundBrushWin);
	
	if (IsWindowVisible(mHwnd))
		// Force the window to repaint so that colors take effect immediately.
		// UpdateWindow() isn't enough sometimes/always, so do something more aggressive:
		InvalidateRect(mHwnd, NULL, TRUE);
	
	return OK;
}


FResult GuiType::get_Name(StrRet &aRetVal)
{
	aRetVal.SetTemp(mName); // Can safely be nullptr.
	return OK;
}


FResult GuiType::set_Name(StrArg aName)
{
	LPTSTR new_name = nullptr;
	if (*aName)
	{
		new_name = _tcsdup(aName);
		if (!new_name)
			return FR_E_OUTOFMEM;
	}
	free(mName);
	mName = new_name;
	return OK;
}


void GuiType::MethodGetPos(int *aX, int *aY, int *aWidth, int *aHeight, RECT &aPos, HWND aOrigin)
{
	MapWindowPoints(aOrigin, mOwner, (LPPOINT)&aPos, 2);
	
	if (aX)			*aX = Unscale(aPos.left);
	if (aY)			*aY = Unscale(aPos.top);
	if (aWidth)		*aWidth = Unscale(aPos.right - aPos.left);
	if (aHeight)	*aHeight = Unscale(aPos.bottom - aPos.top);
}


FResult GuiType::GetPos(int *aX, int *aY, int *aWidth, int *aHeight)
{
	GUI_MUST_HAVE_HWND;
	RECT rect;
	if (GetWindowRect(mHwnd, &rect))
		MethodGetPos(aX, aY, aWidth, aHeight, rect, NULL);
	return OK;
}


FResult GuiType::GetClientPos(int *aX, int *aY, int *aWidth, int *aHeight)
{
	GUI_MUST_HAVE_HWND;
	RECT rect;
	if (GetClientRect(mHwnd, &rect))
		MethodGetPos(aX, aY, aWidth, aHeight, rect, mHwnd);
	return OK;
}


FResult GuiType::Move(optl<int> aX, optl<int> aY, optl<int> aWidth, optl<int> aHeight)
{
	GUI_MUST_HAVE_HWND;
	// Like WinMove, this doesn't check if the window is minimized or do any autosizing.
	// Unlike WinMove, this does DPI scaling.
	RECT rect;
	GetWindowRect(mHwnd, &rect);
	rect.right -= rect.left; // Convert to width.
	rect.bottom -= rect.top; // Convert to height.
	if (HWND parent = GetParent(mHwnd)) // Allow for +Parent.
		ScreenToClient(parent, (LPPOINT)&rect); // Must do this before the loop, as coord[] already contains client coords.
	if (aX.has_value())			rect.left = Scale(aX.value());
	if (aY.has_value())			rect.top = Scale(aY.value());
	if (aWidth.has_value())		rect.right = Scale(aWidth.value());
	if (aHeight.has_value())	rect.bottom = Scale(aHeight.value());
	MoveWindow(mHwnd, rect.left, rect.top, rect.right, rect.bottom, TRUE);
	return OK;
}


bif_impl void GuiFromHwnd(UINT aHwnd, optl<BOOL> aRecurse, IObject *&aGui)
{
	if (aGui = aRecurse.value_or(FALSE) ? GuiType::FindGuiParent((HWND)(UINT_PTR)aHwnd) : GuiType::FindGui((HWND)(UINT_PTR)aHwnd))
		aGui->AddRef();
}


bif_impl void GuiCtrlFromHwnd(UINT aHwnd, IObject *&aGuiCtrl)
{
	if (GuiType* gui = GuiType::FindGuiParent((HWND)(UINT_PTR)aHwnd))
		if (aGuiCtrl = gui->FindControl((HWND)(UINT_PTR)aHwnd))
			aGuiCtrl->AddRef();
}


ObjectMemberMd GuiControlType::sMembers[] =
{
	md_member(GuiControlType, Focus, CALL, md_arg_none),
	md_member(GuiControlType, GetPos, CALL, (Out_Opt, Int32, X), (Out_Opt, Int32, Y), (Out_Opt, Int32, Width), (Out_Opt, Int32, Height)),
	md_member(GuiControlType, Move, CALL, (In_Opt, Int32, X), (In_Opt, Int32, Y), (In_Opt, Int32, Width), (In_Opt, Int32, Height)),
	md_member(GuiControlType, OnCommand, CALL, (In, Int32, NotifyCode), (In, Variant, Callback), (In_Opt, Int32, AddRemove)),
	md_member(GuiControlType, OnEvent, CALL, (In, String, EventName), (In, Variant, Callback), (In_Opt, Int32, AddRemove)),
	md_member(GuiControlType, OnNotify, CALL, (In, Int32, NotifyCode), (In, Variant, Callback), (In_Opt, Int32, AddRemove)),
	md_member(GuiControlType, Opt, CALL, (In, String, Options)),
	md_member(GuiControlType, Redraw, CALL, md_arg_none),
	md_member(GuiControlType, SetFont, CALL, (In_Opt, String, Options), (In_Opt, String, FontName)),
	
	md_property_get	(GuiControlType, ClassNN, String),
	md_property		(GuiControlType, Enabled, Bool32),
	md_property_get	(GuiControlType, Focused, Bool32),
	md_property_get	(GuiControlType, Gui, Object),
	md_property_get	(GuiControlType, Hwnd, UInt32),
	md_property		(GuiControlType, Name, String),
	md_property		(GuiControlType, Text, Variant),
	md_property_get	(GuiControlType, Type, String),
	md_property		(GuiControlType, Value, Variant),
	md_property		(GuiControlType, Visible, Bool32)
};

ObjectMemberMd GuiControlType::sMembersList[] =
{
	md_member_x(GuiControlType, Add, List_Add, CALL, (In, Variant, Value)),
	md_member_x(GuiControlType, Choose, List_Choose, CALL, (In, Variant, Value)),
	md_member_x(GuiControlType, Delete, List_Delete, CALL, (In_Opt, Int32, Index))
};

ObjectMemberMd GuiControlType::sMembersTab[] =
{
	md_member_x(GuiControlType, UseTab, Tab_UseTab, CALL, (In_Opt, Variant, Tab), (In_Opt, Bool32, ExactMatch))
};

ObjectMemberMd GuiControlType::sMembersDate[] =
{
	md_member_x(GuiControlType, SetFormat, DT_SetFormat, CALL, (In_Opt, String, Format))
};

#define FUN1(name, minp, maxp, bif) Object_Member(name, bif, 0, IT_CALL, minp, maxp)
#define FUNn(name, minp, maxp, bif, cat) Object_Member(name, bif, FID_##cat##_##name, IT_CALL, minp, maxp)

ObjectMemberMd GuiControlType::sMembersLV[] =
{
	md_member_x(GuiControlType, GetNext, LV_GetNext, CALL, (In_Opt, Int32, StartIndex), (In_Opt, String, RowType), (Ret, Int32, RetVal)),
	md_member_x(GuiControlType, GetCount, LV_GetCount, CALL, (In_Opt, String, Mode), (Ret, Int32, RetVal)),
	md_member_x(GuiControlType, GetText, LV_GetText, CALL, (In, Int32, Row), (In_Opt, Int32, Column), (Ret, String, RetVal)),
	md_member_x(GuiControlType, Add, LV_Add, CALL, (In_Opt, String, Options), (In, Params, Columns), (Ret, Int32, RetVal)),
	md_member_x(GuiControlType, Insert, LV_Insert, CALL, (In, Int32, Row), (In_Opt, String, Options), (In, Params, Columns), (Ret, Int32, RetVal)),
	md_member_x(GuiControlType, Modify, LV_Modify, CALL, (In, Int32, Row), (In_Opt, String, Options), (In, Params, Columns)),
	md_member_x(GuiControlType, Delete, LV_Delete, CALL, (In_Opt, Int32, Row)),
	md_member_x(GuiControlType, InsertCol, LV_InsertCol, CALL, (In_Opt, Int32, Column), (In_Opt, String, Options), (In_Opt, String, Title), (Ret, Int32, RetVal)),
	md_member_x(GuiControlType, ModifyCol, LV_ModifyCol, CALL, (In_Opt, Int32, Column), (In_Opt, String, Options), (In_Opt, String, Title)),
	md_member_x(GuiControlType, DeleteCol, LV_DeleteCol, CALL, (In, Int32, Column)),
	md_member_x(GuiControlType, SetImageList, LV_SetImageList, CALL, (In, UIntPtr, ImageListID), (In_Opt, Int32, IconType), (Ret, UIntPtr, RetVal))
};

ObjectMemberMd GuiControlType::sMembersTV[] =
{
	md_member_x(GuiControlType, Add, TV_Add, CALL, (In, String, Name), (In_Opt, UIntPtr, ParentItemID), (In_Opt, String, Options), (Ret, UIntPtr, RetVal)),
	md_member_x(GuiControlType, Delete, TV_Delete, CALL, (In_Opt, UIntPtr, ItemID)),
	md_member_x(GuiControlType, Get, TV_Get, CALL, (In, UIntPtr, ItemID), (In, String, Attribute), (Ret, UIntPtr, RetVal)),
	md_member_x(GuiControlType, GetChild, TV_GetChild, CALL, (In, UIntPtr, ItemID), (Ret, UIntPtr, RetVal)),
	md_member_x(GuiControlType, GetCount, TV_GetCount, CALL, (Ret, UInt32, RetVal)),
	md_member_x(GuiControlType, GetNext, TV_GetNext, CALL, (In_Opt, UIntPtr, ItemID), (In_Opt, String, ItemType), (Ret, UIntPtr, RetVal)),
	md_member_x(GuiControlType, GetParent, TV_GetParent, CALL, (In, UIntPtr, ItemID), (Ret, UIntPtr, RetVal)),
	md_member_x(GuiControlType, GetPrev, TV_GetPrev, CALL, (In, UIntPtr, ItemID), (Ret, UIntPtr, RetVal)),
	md_member_x(GuiControlType, GetSelection, TV_GetSelection, CALL, (Ret, UIntPtr, RetVal)),
	md_member_x(GuiControlType, GetText, TV_GetText, CALL, (In, UIntPtr, ItemID), (Ret, String, RetVal)),
	md_member_x(GuiControlType, Modify, TV_Modify, CALL, (In, UIntPtr, ItemID), (In_Opt, String, Options), (In_Opt, String, NewName), (Ret, UIntPtr, RetVal)),
	md_member_x(GuiControlType, SetImageList, TV_SetImageList, CALL, (In, UIntPtr, ImageListID), (In_Opt, Int32, IconType), (Ret, UIntPtr, RetVal))
};

ObjectMemberMd GuiControlType::sMembersSB[] =
{
	md_member_x(GuiControlType, SetIcon, SB_SetIcon, CALL, (In, String, Filename), (In_Opt, Int32, IconNumber), (In_Opt, UInt32, PartNumber), (Ret, UIntPtr, RetVal)),
	md_member_x(GuiControlType, SetParts, SB_SetParts, CALL, (In, Params, Filename), (Ret, UInt32, RetVal)),
	md_member_x(GuiControlType, SetText, SB_SetText, CALL, (In, String, NewText), (In_Opt, UInt32, PartNumber), (In_Opt, UInt32, Style))
};

#undef FUN1
#undef FUNn

Object *GuiControlType::sPrototype;
Object *GuiControlType::sPrototypeList;
Object *GuiControlType::sPrototypes[GUI_CONTROL_TYPE_COUNT];

void GuiControlType::DefineControlClasses()
{
	auto gui_class = (Object *)g_script.FindGlobalVar(_T("Gui"), 3)->Object();

	sPrototype = CreatePrototype(_T("Gui.Control"), Object::sPrototype, sMembers, _countof(sMembers));
	sPrototypeList = CreatePrototype(_T("Gui.List"), sPrototype, sMembersList, _countof(sMembersList));
	auto ctrl_class = CreateClass(sPrototype, Object::sClass);
	auto list_class = CreateClass(sPrototypeList, ctrl_class);
	gui_class->DefineClass(_T("Control"), ctrl_class);
	gui_class->DefineClass(_T("List"), list_class);

	for (int i = GUI_CONTROL_INVALID + 1; i < GUI_CONTROL_TYPE_COUNT; ++i)
	{
		if (i == GUI_CONTROL_TAB2 || i == GUI_CONTROL_TAB3)
			continue;

		// Determine base prototype and control-type-specific members.
		Object *base_proto = sPrototype, *base_class = ctrl_class;
		ObjectMemberListType more_items;
		int how_many = 0;
		switch (i)
		{
		case GUI_CONTROL_TAB: more_items = sMembersTab; how_many = _countof(sMembersTab); // Fall through:
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
		case GUI_CONTROL_LISTBOX: base_proto = sPrototypeList; base_class = list_class; break;
		case GUI_CONTROL_DATETIME: more_items = sMembersDate; how_many = _countof(sMembersDate); break;
		case GUI_CONTROL_LISTVIEW: more_items = sMembersLV; how_many = _countof(sMembersLV); break;
		case GUI_CONTROL_TREEVIEW: more_items = sMembersTV; how_many = _countof(sMembersTV); break;
		case GUI_CONTROL_STATUSBAR: more_items = sMembersSB; how_many = _countof(sMembersSB); break;
		}
		TCHAR buf[32];
		_sntprintf(buf, 32, _T("Gui.%s"), sTypeNames[i]);
		sPrototypes[i] = CreatePrototype(buf, base_proto, more_items, how_many);
		auto cls = CreateClass(sPrototypes[i]);
		cls->SetBase(base_class);
		gui_class->DefineClass(sTypeNames[i], cls);
	}
}

Object *GuiControlType::GetPrototype(GuiControls aType)
{
	ASSERT(aType <= _countof(sPrototypes));
	if (aType == GUI_CONTROL_TAB2 || aType == GUI_CONTROL_TAB3)
		aType = GUI_CONTROL_TAB; // Just make them all Gui.Tab.
	return sPrototypes[aType];
}


FResult GuiControlType::Focus()
{
	CTRL_THROW_IF_DESTROYED;
	// MSDN: "WM_NEXTDLGCTL updates the default pushbutton border, sets the default control identifier,
	// and automatically selects the text of an edit control (if the target window is an edit control)."
	// (It also sets the focus.)
	SendMessage(gui->mHwnd, WM_NEXTDLGCTL, (WPARAM)hwnd, TRUE);
	return OK;
}


FResult GuiControlType::GetPos(int *aX, int *aY, int *aWidth, int *aHeight)
{
	CTRL_THROW_IF_DESTROYED;
	
	RECT rect;
	GetWindowRect(hwnd, &rect);
	MapWindowPoints(NULL, gui->mHwnd, (LPPOINT)&rect, 2);
	
	if (aX)			*aX = gui->Unscale(rect.left);
	if (aY)			*aY = gui->Unscale(rect.top);
	if (aWidth)		*aWidth = gui->Unscale(rect.right - rect.left);
	if (aHeight)	*aHeight = gui->Unscale(rect.bottom - rect.top);
	
	return OK;
}


FResult GuiControlType::Move(optl<int> aX, optl<int> aY, optl<int> aWidth, optl<int> aHeight)
{
	CTRL_THROW_IF_DESTROYED;
	return gui->ControlMove(*this, aX.value_or(COORD_UNSPECIFIED), aY.value_or(COORD_UNSPECIFIED), aWidth.value_or(COORD_UNSPECIFIED), aHeight.value_or(COORD_UNSPECIFIED));
}


FResult GuiControlType::OnCommand(int aNotifyCode, ExprTokenType &aCallback, optl<int> aAddRemove)
{
	CTRL_THROW_IF_DESTROYED;
	return gui->OnEvent(this, aNotifyCode, GUI_EVENTKIND_COMMAND, aCallback, aAddRemove);
}


FResult GuiControlType::OnEvent(StrArg aEventName, ExprTokenType &aCallback, optl<int> aAddRemove)
{
	CTRL_THROW_IF_DESTROYED;
	UINT event_code = GuiType::ConvertEvent(aEventName);
	if (!event_code || !SupportsEvent(event_code))
		return FR_E_ARG(0);
	return gui->OnEvent(this, event_code, GUI_EVENTKIND_EVENT, aCallback, aAddRemove);
}


FResult GuiControlType::OnNotify(int aNotifyCode, ExprTokenType &aCallback, optl<int> aAddRemove)
{
	CTRL_THROW_IF_DESTROYED;
	return gui->OnEvent(this, aNotifyCode, GUI_EVENTKIND_NOTIFY, aCallback, aAddRemove);
}


FResult GuiControlType::Opt(StrArg aOptions)
{
	CTRL_THROW_IF_DESTROYED;
	GuiControlOptionsType go; // Its contents are not currently used here, but they might be in the future.
	gui->ControlInitOptions(go, *this);
	if (!gui->ControlParseOptions(aOptions, go, *this, GUI_HWND_TO_INDEX(hwnd)))
		return FR_FAIL;
	return OK;
}


FResult GuiControlType::Redraw()
{
	CTRL_THROW_IF_DESTROYED;
	RECT rect;
	GetWindowRect(hwnd, &rect); // Limit it to only that part of the client area that the control occupies.
	MapWindowPoints(NULL, gui->mHwnd, (LPPOINT)&rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
	InvalidateRect(gui->mHwnd, &rect, TRUE); // Seems safer to use TRUE, not knowing all possible overlaps, etc.
	return OK;
}


FResult GuiControlType::SetFont(optl<StrArg> aOptions, optl<StrArg> aFontName)
{
	CTRL_THROW_IF_DESTROYED;
	return gui->ControlSetFont(*this, aOptions.value_or_empty(), aFontName.value_or_empty()) ? OK : FR_FAIL;
}


FResult GuiControlType::get_ClassNN(StrRet &aRetVal)
{
	CTRL_THROW_IF_DESTROYED;
	TCHAR class_nn[WINDOW_CLASS_NN_SIZE];
	auto fr = ControlGetClassNN(gui->mHwnd, hwnd, class_nn, _countof(class_nn));
	return fr != OK ? fr : aRetVal.Copy(class_nn) ? OK : FR_E_OUTOFMEM;
}


FResult GuiControlType::get_Enabled(BOOL &aRetVal)
{
	CTRL_THROW_IF_DESTROYED;
	aRetVal = IsWindowEnabled(hwnd) ? 1 : 0;
	return OK;
}


FResult GuiControlType::set_Enabled(BOOL aValue)
{
	CTRL_THROW_IF_DESTROYED;
	gui->ControlSetEnabled(*this, aValue);
	return OK;
}


FResult GuiControlType::get_Focused(BOOL &aRetVal)
{
	CTRL_THROW_IF_DESTROYED;
	HWND focus = GetFocus();
	// For controls like ComboBox and ActiveX that may have a focused child window,
	// also co nsider the control focused if any child of the control is focused.
	aRetVal = (focus == hwnd || IsChild(hwnd, focus));
	return OK;
}


FResult GuiControlType::get_Gui(IObject *&aRetVal)
{
	CTRL_THROW_IF_DESTROYED;
	gui->AddRef();
	aRetVal = gui;
	return OK;
}


FResult GuiControlType::get_Hwnd(UINT &aRetVal)
{
	CTRL_THROW_IF_DESTROYED;
	aRetVal = (UINT)(UINT_PTR)hwnd;
	return OK;
}


FResult GuiControlType::get_Name(StrRet &aRetVal)
{
	aRetVal.SetTemp(name); // Can safely be nullptr.
	return OK;
}


FResult GuiControlType::set_Name(StrArg aValue)
{
	CTRL_THROW_IF_DESTROYED;
	return gui->ControlSetName(*this, aValue);
}


FResult GuiControlType::get_Text(ResultToken &aResultToken)
{
	CTRL_THROW_IF_DESTROYED;
	gui->ControlGetContents(aResultToken, *this, GuiType::Text_Mode);
	return aResultToken.Exited() ? FR_FAIL : OK;
}


FResult GuiControlType::set_Text(ExprTokenType &aValue)
{
	CTRL_THROW_IF_DESTROYED;
	return gui->ControlSetContents(*this, aValue, true);
}


FResult GuiControlType::get_Type(StrRet &aRetVal)
{
	aRetVal.SetStatic(GetTypeName());
	return OK;
}


FResult GuiControlType::get_Value(ResultToken &aResultToken)
{
	CTRL_THROW_IF_DESTROYED;
	gui->ControlGetContents(aResultToken, *this, GuiType::Value_Mode);
	return aResultToken.Exited() ? FR_FAIL : OK;
}


FResult GuiControlType::set_Value(ExprTokenType &aValue)
{
	CTRL_THROW_IF_DESTROYED;
	return gui->ControlSetContents(*this, aValue, false);
}


FResult GuiControlType::get_Visible(BOOL &aRetVal)
{
	CTRL_THROW_IF_DESTROYED;
	// Return the requested visibility state, not the actual state, which would
	// be affected by the parent window's visiblity and currently selected tab.
	aRetVal = (attrib & GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN) ? 0 : 1;
	return OK;
}


FResult GuiControlType::set_Visible(BOOL aValue)
{
	CTRL_THROW_IF_DESTROYED;
	gui->ControlSetVisible(*this, aValue);
	return OK;
}


FResult GuiControlType::Tab_UseTab(ExprTokenType *aTab, optl<BOOL> aExact)
{
	CTRL_THROW_IF_DESTROYED;
	BOOL whole_match = aExact.value_or(FALSE);
	int index; // int vs TabControlIndexType for detection of negative/out-of-range values.
	if (!aTab)
		index = -1;
	else if (TokenIsPureNumeric(*aTab) == SYM_INTEGER)
		index = (int)TokenToInt64(*aTab) - 1; // Convert to zero-based and let 0 indicate "no tab" (-1).
	else
	{
		TCHAR float_buf[MAX_NUMBER_SIZE];
		auto tab_name = TokenToString(*aTab, float_buf);
		index = gui->FindTabIndexByName(*this, tab_name, whole_match);
		if (index == -1 && *tab_name) // Invalid tab name (but "" is a valid way to specify "no tab").
			index = -2; // Flag it as invalid.
	}
	// Validate parameters before making any changes to the Gui's state:
	if (index < -1 || index >= MAX_TABS_PER_CONTROL)
		return FR_E_ARG(0);

	TabIndexType prev_tab_index = gui->mCurrentTabIndex;
	TabControlIndexType prev_tab_control_index = gui->mCurrentTabControlIndex;

	if (index == -1)
		gui->mCurrentTabControlIndex = MAX_TAB_CONTROLS; // i.e. "no tab"
	else
	{
		gui->mCurrentTabControlIndex = tab_index;
		if (gui->mCurrentTabIndex != index)
		{
			gui->mCurrentTabIndex = index;
			// Changing to a different tab should start a new radio group:
			gui->mInRadioGroup = false;
		}
	}
	if (gui->mCurrentTabControlIndex != prev_tab_control_index)
	{
		if (GuiControlType *tab_control = gui->FindTabControl(prev_tab_control_index))
		{
			// Autosize the previous tab control.  Has no effect if it is not a Tab3 control or has
			// already been autosized.  Doing it at this point allows the script to set different
			// margins for inside and outside the tab control, and is simpler than the alternative:
			// waiting until the next external control is added.  The main drawback is that the
			// script is unable to alternate between tab controls and still utilize autosizing.
			// On the flip side, scripts can use this to their advantage -- to add controls which
			// should not affect the tab control's size.
			gui->AutoSizeTabControl(*tab_control);
		}
		// Changing to a different tab control (or none at all when there
		// was one before, or vice versa) should start a new radio group:
		gui->mInRadioGroup = false;
	}
	return OK;
}


FResult GuiControlType::List_Add(ExprTokenType &aItems)
{
	CTRL_THROW_IF_DESTROYED;
	auto obj = TokenToArray(aItems);
	if (!obj)
		return FParamError(0, &aItems, _T("Array"));
	gui->ControlAddItems(*this, obj);
	if (type == GUI_CONTROL_TAB)
	{
		// Adding tabs may change the space usable by the tab dialog.
		gui->UpdateTabDialog(hwnd);
		// Appears to be necessary to resolve a redraw issue, at least for Tab3 on Windows 10.
		InvalidateRect(gui->mHwnd, NULL, TRUE);
	}
	return OK;
}


FResult GuiControlType::List_Delete(optl<int> aIndex)
{
	CTRL_THROW_IF_DESTROYED;
	UINT msg_one;
	UINT msg_all;
	switch (type)
	{
	case GUI_CONTROL_TAB:     msg_one = TCM_DELETEITEM; msg_all = TCM_DELETEALLITEMS; break;
	case GUI_CONTROL_LISTBOX: msg_one = LB_DELETESTRING; msg_all = LB_RESETCONTENT; break;
	default:                  msg_one = CB_DELETESTRING; msg_all = CB_RESETCONTENT; break; // ComboBox/DDL.
	}
	if (aIndex.has_value()) // Delete item with given index.
	{
		int index = aIndex.value() - 1;
		if (index < 0)
			return FR_E_ARG(0);
		SendMessage(hwnd, msg_one, (WPARAM)index, 0);
	}
	else // Delete all items.
	{
		SendMessage(hwnd, msg_all, 0, 0);
	}
	if (type == GUI_CONTROL_TAB)
	{
		// Removing tabs may change the space usable by the tab dialog.
		gui->UpdateTabDialog(hwnd);
	}
	return OK;
}


GuiType *GuiType::FindGui(HWND aHwnd)
{
	// Check that this window is an AutoHotkey Gui.
	ATOM atom = (ATOM)GetClassLong(aHwnd, GCW_ATOM);
	if (atom != sGuiWinClass)
		return NULL;

	// Retrieve the GuiType object associated to it.
	return (GuiType*)GetWindowLongPtr(aHwnd, GWLP_USERDATA);
}


GuiType *GuiType::FindGuiParent(HWND aHwnd)
// Returns the GuiType of aHwnd or its closest ancestor which is a Gui.
{
	for ( ; aHwnd; aHwnd = GetParent(aHwnd))
	{
		if (GuiType *gui = FindGui(aHwnd))
			return gui;
		if (!(GetWindowLong(aHwnd, GWL_STYLE) & WS_CHILD))
			break;
	}
	return NULL;
}


FResult GuiType::get_MenuBar(ResultToken &aResultToken)
{
	GUI_MUST_HAVE_HWND;
	if (mMenu)
	{
		mMenu->AddRef();
		aResultToken.Return(mMenu);
	}
	return OK;
}


FResult GuiType::set_MenuBar(ExprTokenType &aParam)
{
	GUI_MUST_HAVE_HWND;
	UserMenu *menu = NULL;
	if (!TokenIsEmptyString(aParam))
	{
		menu = dynamic_cast<UserMenu *>(TokenToObject(aParam));
		if (!menu || menu->mMenuType != MENU_TYPE_BAR)
			return FTypeError(_T("MenuBar"), aParam);
		menu->CreateHandle(); // Ensure the menu bar physically exists.
		menu->AddRef();
	}
	if (mMenu)
		mMenu->Release();
	mMenu = menu;
	::SetMenu(mHwnd, menu ? menu->mMenu : NULL);  // Add or remove the menu.
	if (menu) // v1.1.04: Keyboard accelerators.
		UpdateAccelerators(*menu);
	else
		RemoveAccelerators();
	return OK;
}


FResult GuiType::ControlSetName(GuiControlType &aControl, LPCTSTR aName)
{
	LPTSTR new_name = nullptr;
	if (aName && *aName)
	{
		// Make sure the name isn't already taken.
		for (GuiIndexType u = 0; u < mControlCount; ++u)
			if (mControl[u]->name && !_tcsicmp(mControl[u]->name, aName))
			{
				if (mControl[u] == &aControl) // aControl already has this name, but upper- vs lower-case may differ.
					break;
				return FError(_T("A control with this name already exists."), aName);
			}

		new_name = _tcsdup(aName);
		if (!new_name)
			return FR_E_OUTOFMEM;
	}
	if (aControl.name)
		free(aControl.name);
	aControl.name = new_name;
	return OK;
}


void GuiType::ControlSetEnabled(GuiControlType &aControl, bool aEnabled)
{
	GuiControlType *tab_control;
	int selection_index;

	// GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED is maintained for use with tab controls.  It allows controls
	// on inactive tabs to be marked for later enabling.  It also allows explicitly disabled controls to
	// stay disabled even when their tab/page becomes active. It is updated unconditionally for simplicity
	// and maintainability.  
	if (aEnabled)
		aControl.attrib &= ~GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED;
	else
		aControl.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED;
	if (tab_control = FindTabControl(aControl.tab_control_index)) // It belongs to a tab control that already exists.
	{
		if (GetWindowLong(tab_control->hwnd, GWL_STYLE) & WS_DISABLED) // But its tab control is disabled...
			return;
		selection_index = TabCtrl_GetCurSel(tab_control->hwnd);
		if (selection_index != aControl.tab_index && selection_index != -1)
			// There is no current tab/page or the one selected is not this control's:
			// Do not disable or re-enable the control in this case.
			// v1.0.48.04: Above now also checks for -1, which is a tab control containing zero tabs/pages.
			// The controls on such a tab control might be wrongly/inadvertently visible because
			// ControlUpdateCurrentTab() isn't capable of handling that situation.  Since fixing
			// ControlUpdateCurrentTab() would reduce backward compatibility -- and in case anyone is
			// using tabless tab controls for anything -- it seems best to allow these "wrongly visible"
			// controls to be explicitly manipulated by GuiControl Enable/Disable and Hide/Show.
			// UPDATE: ControlUpdateCurrentTab() has been fixed, so the controls won't be visible
			// unless the script specifically made them so.  The check is kept in case it is of
			// some use.  Prior to the fix, the controls would have only been visible if the tab
			// control previously had a tab but it was deleted, since controls are created in a
			// hidden state if the tab they are on is not selected.
			return;
	}
		
	// L23: Restrict focus workaround to when the control is/was actually focused. Fixes a bug introduced by L13: enabling or disabling a control caused the active Edit control to reselect its text.
	bool gui_control_was_focused = GetForegroundWindow() == mHwnd && GetFocus() == aControl.hwnd;

	// Since above didn't return, act upon the enabled/disable:
	EnableWindow(aControl.hwnd, aEnabled);
		
	// L23: Only if EnableWindow removed the keyboard focus entirely, reset the focus.
	if (gui_control_was_focused && !GetFocus())
		SetFocus(mHwnd);
		
	if (aControl.type == GUI_CONTROL_TAB) // This control is a tab control.
		// Update the control so that its current tab's controls will all be enabled or disabled (now
		// that the tab control itself has just been enabled or disabled):
		ControlUpdateCurrentTab(aControl, false);
}



void GuiType::ControlSetVisible(GuiControlType &aControl, bool aVisible)
{
	GuiControlType *tab_control;
	int selection_index;

	// GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN is maintained for use with tab controls.  It allows controls
	// on inactive tabs to be marked for later showing.  It also allows explicitly hidden controls to
	// stay hidden even when their tab/page becomes active. It is updated unconditionally for simplicity
	// and maintainability.
	if (aVisible)
		aControl.attrib &= ~GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN;
	else
		aControl.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN;
	if (tab_control = FindTabControl(aControl.tab_control_index)) // It belongs to a tab control that already exists.
	{
		if (!(GetWindowLong(tab_control->hwnd, GWL_STYLE) & WS_VISIBLE)) // But its tab control is hidden...
			return;
		selection_index = TabCtrl_GetCurSel(tab_control->hwnd);
		if (selection_index != aControl.tab_index && selection_index != -1)
			return; // v1.0.48.04: Concerning the line above, see comments in ControlSetEnabled.
	}
	// Since above didn't return, act upon the show/hide:
	ShowWindow(aControl.hwnd, aVisible ? SW_SHOWNOACTIVATE : SW_HIDE);
	if (aControl.type == GUI_CONTROL_TAB) // This control is a tab control.
		// Update the control so that its current tab's controls will all be shown or hidden (now
		// that the tab control itself has just been shown or hidden):
		ControlUpdateCurrentTab(aControl, false);
}


FResult GuiType::ControlMove(GuiControlType &aControl, int xpos, int ypos, int width, int height)
{
	RECT rect;
	GetWindowRect(aControl.hwnd, &rect); // Failure seems too rare to check for.
	POINT dest_pt = {rect.left, rect.top};
	ScreenToClient(mHwnd, &dest_pt); // Set default x/y target position, to be possibly overridden below.
	if (xpos != COORD_UNSPECIFIED)
		dest_pt.x = Scale(xpos);
	if (ypos != COORD_UNSPECIFIED)
		dest_pt.y = Scale(ypos);
	
	// Map to the GUI window's client area (inverse of GetPos) in case the
	// GUI window isn't the control's parent, such as if it is on a Tab3.
	MapWindowPoints(mHwnd, GetParent(aControl.hwnd), &dest_pt, 1);

	if (!MoveWindow(aControl.hwnd, dest_pt.x, dest_pt.y
		, width == COORD_UNSPECIFIED ? rect.right - rect.left : Scale(width)
		, height == COORD_UNSPECIFIED ? rect.bottom - rect.top : Scale(height)
		, TRUE))  // Do repaint.
		return FR_E_WIN32;

	// Note that GUI_CONTROL_UPDOWN has no special handling here.  This is because unlike slider buddies,
	// whose only purpose is to label the control, an up-down's is also content-linked to it, so the
	// inability to move the up-down to separate it from its buddy would be a loss of flexibility.  For
	// this reason and also to reduce code size, the control is not re-buddied to snap them together.
	if (aControl.type == GUI_CONTROL_SLIDER) // It seems buddies don't move automatically, so trigger the move.
	{
		HWND buddy1 = (HWND)SendMessage(aControl.hwnd, TBM_GETBUDDY, TRUE, 0);
		HWND buddy2 = (HWND)SendMessage(aControl.hwnd, TBM_GETBUDDY, FALSE, 0);
		if (buddy1)
		{
			SendMessage(aControl.hwnd, TBM_SETBUDDY, TRUE, (LPARAM)buddy1);
			// It doesn't always redraw the buddies correctly, at least on XP, so do it manually:
			InvalidateRect(buddy1, NULL, TRUE);
		}
		if (buddy2)
		{
			SendMessage(aControl.hwnd, TBM_SETBUDDY, FALSE, (LPARAM)buddy2);
			InvalidateRect(buddy2, NULL, TRUE);
		}
	}
	else if (aControl.type == GUI_CONTROL_TAB)
	{
		// Check for autosizing flags on this Tab control.
		int mask = (width == COORD_UNSPECIFIED ? 0 : TAB3_AUTOWIDTH)
			| (height == COORD_UNSPECIFIED ? 0 : TAB3_AUTOHEIGHT);
		int autosize = (int)(INT_PTR)GetProp(aControl.hwnd, _T("ahk_autosize"));
		if (autosize & mask) // There are autosize flags to unset (implies this is a Tab3 control).
		{
			autosize &= ~mask; // Unset the flags corresponding to dimensions we've just set.
			if (autosize)
				SetProp(aControl.hwnd, _T("ahk_autosize"), (HANDLE)(INT_PTR)autosize);
			else
				RemoveProp(aControl.hwnd, _T("ahk_autosize"));
		}
	}
	return OK;
}


ResultType GuiType::ControlSetFont(GuiControlType &aControl, LPCTSTR aOptions, LPCTSTR aFontName)
{
	if (!aControl.UsesFontAndTextColor()) // Control has no use for a font.
		return g_script.RuntimeError(ERR_GUI_NOT_FOR_THIS_TYPE);
	COLORREF color = CLR_INVALID;
	int font_index;
	if (*aOptions || *aFontName) // Use specified font.
	{
		FontType font;
		// Get control's current font, to inherit attributes.
		font.hfont = (HFONT)SendMessage(aControl.hwnd, WM_GETFONT, 0, 0);
		FontGetAttributes(font);

		// Find or create font with the changed attributes.
		font_index = FindOrCreateFont(aOptions, aFontName, &font, &color);
		if (font_index == -1)
			return FAIL;
	}
	else // Use GUI's current font.
	{
		font_index = mCurrentFontIndex;
		// It seems more correct to leave the control's current color, since there's
		// no way to know whether it should take precedence over the GUI's text color.
		// Require the caller to be explicit and specify the color if it should be changed.
		//color = mCurrentColor;
	}
	SendMessage(aControl.hwnd, WM_SETFONT, (WPARAM)sFont[font_index].hfont, 0);
	if (color != CLR_INVALID)
		ControlSetTextColor(aControl, color);
	ControlRedraw(aControl, false); // Required for refresh, at least for edit controls, probably some others.
	return OK;
}


void GuiType::ControlSetTextColor(GuiControlType &aControl, COLORREF aColor)
{
	switch (aControl.type)
	{
	case GUI_CONTROL_LISTVIEW:
		// Although MSDN doesn't say, ListView_SetTextColor() appears to support CLR_DEFAULT.
		ListView_SetTextColor(aControl.hwnd, aColor);
		break;
	case GUI_CONTROL_TREEVIEW:
		// Must pass -1 rather than CLR_DEFAULT for TreeView controls.
		TreeView_SetTextColor(aControl.hwnd, aColor == CLR_DEFAULT ? -1 : aColor);
		break;
	case GUI_CONTROL_DATETIME:
		ControlSetMonthCalColor(aControl, aColor, DTM_SETMCCOLOR);
		break;
	case GUI_CONTROL_MONTHCAL:
		ControlSetMonthCalColor(aControl, aColor, MCM_SETCOLOR);
		break;
	default:
		if (aControl.UsesUnionColor()) // For maintainability.  Overlaps with previous checks.
			aControl.union_color = aColor; // Used by WM_CTLCOLORSTATIC et. al. for some types of controls.
		break;
	}
}


void GuiType::ControlSetMonthCalColor(GuiControlType &aControl, COLORREF aColor, UINT aMsg)
{
	// Testing shows that changing the COLOR_WINDOWTEXT system color affects MonthCal
	// controls, but passing CLR_DEFAULT always sets the color to black regardless.
	// MSDN doesn't seem to specify any way to revert to defaults.
	if (aColor == CLR_DEFAULT)
		aColor = GetSysColor(COLOR_WINDOWTEXT);
	SendMessage(aControl.hwnd, aMsg, MCSC_TEXT, aColor);
}


ResultType GuiType::ControlChoose(GuiControlType &aControl, ExprTokenType &aParam, BOOL aOneExact)
{
	int selection_index = -1;
	bool is_choose_string = true;
	switch (TypeOfToken(aParam))
	{
	case SYM_INTEGER: is_choose_string = false; break;
	case SYM_OBJECT: goto error;
	}
	TCHAR buf[MAX_NUMBER_SIZE];
	
	UINT msg_set_index = 0, msg_select_string = 0, msg_find_string = 0;
	switch(aControl.type)
	{
	case GUI_CONTROL_TAB:
		msg_set_index = TCM_SETCURSEL;
		break;
	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
		msg_set_index = CB_SETCURSEL;
		if (aOneExact)
			msg_find_string = CB_FINDSTRINGEXACT;
		else
			msg_select_string = CB_SELECTSTRING;
		break;
	case GUI_CONTROL_LISTBOX:
		if (GetWindowLong(aControl.hwnd, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
		{
			// MSDN: Do not use [LB_SELECTSTRING] with a list box that has the LBS_MULTIPLESEL or the
			// LBS_EXTENDEDSEL styles:
			msg_set_index = LB_SETSEL;
			msg_find_string = aOneExact ? LB_FINDSTRINGEXACT : LB_FINDSTRING;
		}
		else // single-select listbox
		{
			msg_set_index = LB_SETCURSEL;
			if (aOneExact)
				msg_find_string = LB_FINDSTRINGEXACT;
			else
				msg_select_string = LB_SELECTSTRING;
		}
		break;
	default:  // Not a supported control type.
		// This shouldn't be possible due to checking of type elsewhere.
		//return g_script.RuntimeError(ERR_GUI_NOT_FOR_THIS_TYPE); // Disabled to reduce code size.
		ASSERT(!"invalid control type");
		goto error; // Seems better than returning OK, for the off chance this can happen.
	} // switch(control.type)

	LPTSTR item_string;
	if (is_choose_string)
	{
		item_string = TokenToString(aParam, buf);
		// Testing confirms that the CB and LB messages do not interpret an empty string
		// as the first item.  It seems best for "" to deselect the current item/tab.
		if (!*item_string)
		{
			selection_index = -1;
		}
		else if (msg_select_string)
		{
			if (SendMessage(aControl.hwnd, msg_select_string, -1, (LPARAM)item_string) == CB_ERR) // CB_ERR == LB_ERR
				goto error;
			return OK;
		}
		else if (msg_find_string)
		{
			LRESULT found_item = SendMessage(aControl.hwnd, msg_find_string, -1, (LPARAM)item_string);
			if (found_item == LB_ERR) // CB_ERR == LB_ERR
				goto error;
			selection_index = (int)found_item;
		}
		else
		{
			ASSERT(aControl.type == GUI_CONTROL_TAB);
			selection_index = FindTabIndexByName(aControl, item_string, aOneExact); // Returns -1 on failure.
			if (selection_index == -1)
				goto error;
		}
	}
	else // Choose by position vs. string.
	{
		selection_index = (int)TokenToInt64(aParam) - 1;
		if (selection_index < -1)
			goto error;
	}
	if (msg_set_index == LB_SETSEL) // Multi-select, so use the cumulative method.
	{
		if (aOneExact && selection_index >= 0) // But when selection_index == -1, it would be redundant.
			SendMessage(aControl.hwnd, msg_set_index, FALSE, -1); // Deselect all others.
		SendMessage(aControl.hwnd, msg_set_index, selection_index >= 0, selection_index);
		// The return value isn't checked since MSDN doesn't specify what it is in any
		// case other than failure.
		if (is_choose_string)
		{
			// Since this ListBox is multi-select, select ALL matching items:
			for (;;)
			{
				int found_item = (int)SendMessage(aControl.hwnd, msg_find_string, selection_index, (LPARAM)item_string);
				// MSDN: "When the search reaches the bottom of the list box, it continues
				// searching from the top of the list box back to the item specified by the
				// wParam parameter."  So if there are no more results, it will return the
				// original index found above, not -1.
				if (found_item <= selection_index) // The search has looped.
					break;
				selection_index = found_item;
				SendMessage(aControl.hwnd, LB_SETSEL, TRUE, selection_index);
			}
		}
	}
	else
	{
		int result = (int)SendMessage(aControl.hwnd, msg_set_index, selection_index, 0);
		if (msg_set_index == TCM_SETCURSEL)
		{
			if (result == -1)
			{
				if (SendMessage(aControl.hwnd, TCM_GETCURSEL, 0, 0) != selection_index)
					goto error;
				// Otherwise, it's on the right tab, so maybe result == -1 because there
				// was previously no tab selected.
			}
			if (result != selection_index)
				ControlUpdateCurrentTab(aControl, false);
		}
		else if (result == -1 && selection_index != -1)
			goto error;
	}
	return OK;

error:
	return ValueError(aOneExact ? ERR_INVALID_VALUE : ERR_PARAM1_INVALID, nullptr, FAIL_OR_OK); // Invalid parameter #1 is almost definitely the cause.
}

FResult GuiControlType::List_Choose(ExprTokenType &aValue)
{
	CTRL_THROW_IF_DESTROYED;
	return gui->ControlChoose(*this, aValue) ? OK : FR_FAIL;
}


FResult GuiType::ControlSetContents(GuiControlType &aControl, ExprTokenType &aValue, bool aIsText)
// aContents: The content as a string.
// aResultToken: Used only to set an exit result in the event of an error.
{
	if (TokenToObject(aValue))
		return FR_E_ARG(0); // Generic error because unsure which type is expected.
	TCHAR buf[MAX_NUMBER_SIZE];
	LPTSTR aContents = TokenToString(aValue, buf);
	// Control types for which Text and Value both have special handling:
	switch (aControl.type)
	{
	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
	case GUI_CONTROL_LISTBOX:
	case GUI_CONTROL_TAB:
		return ControlSetChoice(aControl, aContents, aIsText);

	case GUI_CONTROL_EDIT:
		return ControlSetEdit(aControl, aContents, aIsText);

	case GUI_CONTROL_DATETIME:
		if (aIsText)
			return FError(ERR_GUI_NOT_FOR_THIS_TYPE); // Let the user know the control doesn't support this.
		else
			return ControlSetDateTime(aControl, aContents);

	case GUI_CONTROL_TEXT:
		// It seems sensible to treat the caption text of a Text control as its "value",
		// since displaying text is its only purpose; it's just for output only, whereas
		// an Edit control can be for input or output.
		aIsText = true;
		break;
	}
	// The remaining control types have no need for special handling of the Text property,
	// either because they have normal caption text, or because they have no "text version"
	// of the value, or no "value" at all.
	if (aIsText)
	{
		SetWindowText(aControl.hwnd, aContents);
		if (aControl.background_color == CLR_TRANSPARENT) // +BackgroundTrans
			ControlRedraw(aControl);
		return OK;
	}
	// Otherwise, this is the Value property.
	switch (aControl.type)
	{
	case GUI_CONTROL_PIC: return ControlSetPic(aControl, aContents);
	case GUI_CONTROL_MONTHCAL: return ControlSetMonthCal(aControl, aContents);
	case GUI_CONTROL_HOTKEY: return ControlSetHotkey(aControl, aContents);

	case GUI_CONTROL_RADIO: // Same as below.
	case GUI_CONTROL_CHECKBOX:
	case GUI_CONTROL_UPDOWN:
	case GUI_CONTROL_SLIDER:
	case GUI_CONTROL_PROGRESS:
	{
		if (!TokenIsNumeric(aValue))
			return FTypeError(_T("Number"), aValue);
		int value = (int)TokenToInt64(aValue);
		switch (aControl.type)
		{
		case GUI_CONTROL_UPDOWN: return ControlSetUpDown(aControl, value);
		case GUI_CONTROL_SLIDER: return ControlSetSlider(aControl, value);
		case GUI_CONTROL_PROGRESS: return ControlSetProgress(aControl, value);
		case GUI_CONTROL_RADIO: // Same as below.
		case GUI_CONTROL_CHECKBOX: return ControlSetCheck(aControl, value);
		}
	}

	// Already handled in the switch() above:
	//case GUI_CONTROL_DATETIME:
	//case GUI_CONTROL_DROPDOWNLIST:
	//case GUI_CONTROL_COMBOBOX:
	//case GUI_CONTROL_LISTBOX:
	//case GUI_CONTROL_TAB:
	//case GUI_CONTROL_TEXT:

	case GUI_CONTROL_LINK:
	case GUI_CONTROL_GROUPBOX:
	case GUI_CONTROL_BUTTON:
	case GUI_CONTROL_STATUSBAR:
		// Above: The control has no "value".  GuiControl.Text should be used to set its text.
	case GUI_CONTROL_LISTVIEW: // More flexible methods are available.
	case GUI_CONTROL_TREEVIEW: // More flexible methods are available.
	case GUI_CONTROL_CUSTOM: // Don't know what this control's "value" is or how to set it.
	case GUI_CONTROL_ACTIVEX: // Assigning a value doesn't make sense.  SetWindowText() does nothing.
	default: // Should help code size by allowing the compiler to omit checks for the above cases.
		return FError(ERR_INVALID_USAGE);
	}
}


FResult GuiType::ControlSetPic(GuiControlType &aControl, LPTSTR aContents)
{
	// Update: The below doesn't work, so it will be documented that a picture control
	// should be always be referred to by its original filename even if the picture changes.
	// Set the text unconditionally even if the picture can't be loaded.  This text must
	// be set to allow GuiControl(Get) to be able to operate upon the picture without
	// needing to identify it via something like "Static14".
	//SetWindowText(control.hwnd, aContents);
	//SendMessage(control.hwnd, WM_SETTEXT, 0, (LPARAM)aContents);

	// Set default options, to be possibly overridden by any options actually present:
	// Fixed for v1.0.23: Below should use GetClientRect() vs. GetWindowRect(), otherwise
	// a size too large will be returned if the control has a border:
	RECT rect;
	GetClientRect(aControl.hwnd, &rect);
	int width = rect.right - rect.left;
	int height = rect.bottom - rect.top;
	int icon_number = 0; // Zero means "load icon or bitmap (doesn't matter)".
	// Note that setting the control's picture handle to NULL sometimes or always shrinks
	// the control down to zero dimensions, so must be done only after the above.

	// Parse any options that are present in front of the filename:
	LPTSTR next_option = omit_leading_whitespace(aContents);
	if (*next_option == '*') // Options are present.  Must check this here and in the for-loop to avoid omitting legitimate whitespace in a filename that starts with spaces.
	{
		LPTSTR option_end;
		TCHAR orig_char;
		for (; *next_option == '*'; next_option = omit_leading_whitespace(option_end))
		{
			// Find the end of this option item:
			if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
				option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
			// Permanently terminate in between options to help eliminate ambiguity for words contained
			// inside other words, and increase confidence in decimal and hexadecimal conversion.
			orig_char = *option_end;
			*option_end = '\0';
			++next_option; // Skip over the asterisk.  It might point to a zero terminator now.
			if (!_tcsnicmp(next_option, _T("Icon"), 4))
				icon_number = ATOI(next_option + 4); // LoadPicture() correctly handles any negative value.
			else
			{
				switch (ctoupper(*next_option))
				{
				case 'W':
					width = ATOI(next_option + 1);
					break;
				case 'H':
					height = ATOI(next_option + 1);
					break;
				// If not one of the above, such as zero terminator or a number, just ignore it.
				}
			}

			*option_end = orig_char; // Undo the temporary termination so that loop's omit_leading() will work.
		} // for() each item in option list

		// The below assigns option_end + 1 vs. next_option in case the filename is contained in a
		// variable ref and/ that filename contains leading spaces.  Example:
		// GuiControl,, MyPic, *w100 *h-1 %FilenameWithLeadingSpaces%
		// Update: Windows XP and perhaps other OSes will load filenames-containing-leading-spaces
		// even if those spaces are omitted.  However, I'm not sure whether all API calls that
		// use filenames do this, so it seems best to include those spaces whenever possible.
		aContents = *option_end ? option_end + 1 : option_end; // Set aContents to the start of the image's filespec.
	} 
	//else options are not present, so do not set aContents to be next_option because that would
	// omit legitimate spaces and tabs that might exist at the beginning of a real filename (file
	// names can start with spaces).

	// See comments in ControlLoadPicture():
	if (!ControlLoadPicture(aControl, aContents, width, height, icon_number))
		return FValueError(ERR_INVALID_VALUE, aContents); // A bit vague but probably correct.

	// Redraw is usually only needed for pictures in a tab control, but is sometimes
	// needed outside of that, for .gif and perhaps others.  So redraw unconditionally:
	ControlRedraw(aControl);
	return OK;
}


FResult GuiType::ControlSetCheck(GuiControlType &aControl, int checked)
{
	if (  !(checked == 0 || checked == 1 || (aControl.type == GUI_CONTROL_CHECKBOX && checked == -1))  )
		return FValueError(ERR_INVALID_VALUE);
	if (checked == -1)
		checked = BST_INDETERMINATE;
	//else the "checked" var is already set correctly.
	if (aControl.type == GUI_CONTROL_RADIO)
	{
		ControlCheckRadioButton(aControl, GUI_HWND_TO_INDEX(aControl.hwnd), checked);
		return OK;
	}
	// Otherwise, we're operating upon a checkbox.
	SendMessage(aControl.hwnd, BM_SETCHECK, checked, 0);
	return OK;
}


void GuiType::ControlGetCheck(ResultToken &aResultToken, GuiControlType &aControl)
{
	switch (SendMessage(aControl.hwnd, BM_GETCHECK, 0, 0))
	{
	case BST_CHECKED:
		_o_return(1);
	case BST_UNCHECKED:
		_o_return(0);
	case BST_INDETERMINATE:
		// Seems better to use a value other than blank because blank might sometimes represent the
		// state of an uninitialized or unfetched control.  In other words, a blank variable often
		// has an external meaning that transcends the more specific meaning often desirable when
		// retrieving the state of the control:
		_o_return(-1);
	}
	// Shouldn't be reached since ZERO(BST_UNCHECKED) is returned on failure.
}


FResult GuiType::ControlSetChoice(GuiControlType &aControl, LPTSTR aContents, bool aIsText)
{
	ExprTokenType choice;
	if (aIsText)
	{
		if (aControl.type == GUI_CONTROL_COMBOBOX)
		{
			// Use the simple SetWindowText() method, which works on the ComboBox's Edit.
			// Clearing the current selection isn't strictly necessary since ControlGetContents()
			// compensates for the index being inaccurate (due to other possible causes), but it
			// might help in some situations.  If it's not done, CB_GETCURSEL will return the
			// previous selection.  By contrast, a user typing in the box always sets it to -1.
			SendMessage(aControl.hwnd, CB_SETCURSEL, -1, 0);
			SetWindowText(aControl.hwnd, aContents);
			return OK;
		}
		choice.symbol = SYM_STRING;
		choice.marker = aContents;
	}
	else // (!aIsText)
	{
		if (!ParseInteger(aContents, choice.value_int64))
			return FError(ERR_INVALID_VALUE, nullptr, ErrorPrototype::Type);
		choice.symbol = SYM_INTEGER;
	}
	return ControlChoose(aControl, choice, TRUE) ? OK : FR_FAIL; // Pass TRUE to find an exact (not partial) match.
}


void GuiType::ControlGetDDL(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode)
{
	LRESULT index;
	if (aMode == Value_Mode || (aMode == Submit_Mode && (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT))) // Caller wants the position, not the text.
	{
		index = SendMessage(aControl.hwnd, CB_GETCURSEL, 0, 0); // Get index of currently selected item.
		// If there's no item selected (the default state, and also possible by Value:=0
		// or Text:=""), index is -1 (CB_ERR).  Since Value:=0 is valid and Value:=""
		// is not, return 0 in that case.  MSDN: "If no item is selected, it is CB_ERR."
		_o_return((int)index + 1);
	}
	ControlGetWindowText(aResultToken, aControl);
}


void GuiType::ControlGetComboBox(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode)
{
	LRESULT index = SendMessage(aControl.hwnd, CB_GETCURSEL, 0, 0); // Get index of currently selected item.

	//   Problem #1:
	// It seems that any text put into a ComboBox's edit field by typing/pasting or via Ctrl.Text
	// always resets the current index to -1, even if that text exactly matches an existing item.

	//   Problem #2:
	// CB_GETCURSEL sometimes returns an index which doesn't match what's in the Edit control.
	// This typically happens as a result of typing/pasting while the drop-down window is visible
	// and then canceling the dropdown without making a selection.
	// Note that when the CBN_SELENDCANCEL notification is received, CB_GETCURSEL correctly returns -1,
	// but at some point after that the control restores the index to a possibly incorrect value.  This
	// appears to happen in other programs as well, and has been confirmed on Windows 10 and earlier.

	int edit_length = GetWindowTextLength(aControl.hwnd); // Zero is valid (the control may be empty).
	// For simplicity and in case the text is very large, use tmalloc() unconditionally
	// even though the text is probably very small.  This allocation may be returned below.
	LPTSTR edit_text = tmalloc(edit_length + 1);
	if (!edit_text)
		_o_throw_oom;
	edit_length = GetWindowText(aControl.hwnd, edit_text, edit_length + 1);
	if (index != CB_ERR)
	{
		// The control returned a valid index, so verify that it matches the edit text.
		// If a search beginning at the indicated item does not return that item, it must
		// not be a match.  This approach avoids retrieving the item's text, which reduces
		// code size.  Performance hasn't been measured, but at a guess, is probably better
		// when the item matches (as it would 99.9% of the time), and worse in other cases.
		if (!edit_length)
		{
			// A different approach is needed in this case since CB_FINDSTRINGEXACT can't
			// search for an empty string.  This preserves the ability to detect selection
			// of a specific empty list item, though for simplicity, clearing the edit text
			// will not match up to an empty list item.
			if (0 != SendMessage(aControl.hwnd, CB_GETLBTEXTLEN, (WPARAM)index, 0))
				index = CB_ERR;
		}
		else
		{
			if (index != SendMessage(aControl.hwnd, CB_FINDSTRINGEXACT, index - 1, (LPARAM)edit_text))
				index = CB_ERR;
		}
	}
	if (index == CB_ERR)
	{
		// Check if the Edit field contains text that exactly matches one of the items in the drop-list.
		// If it does, that item's position should be retrieved in AltSubmit mode (even for non-AltSubmit
		// mode, this should be done because the case of the item in the drop-list is usually preferable
		// to any varying case the user may have manually typed).
		index = SendMessage(aControl.hwnd, CB_FINDSTRINGEXACT, -1, (LPARAM)edit_text); // It's not case sensitive.
		if (index == CB_ERR && aMode != Value_Mode) // Text or Submit mode.
		{
			// edit_text does not match any of the list items, so just return the text.
			aResultToken.AcceptMem(edit_text, edit_length);
			_o_return_retval;
		}
	}
	free(edit_text);

	if (aMode == Value_Mode || (aMode == Submit_Mode && (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT))) // Caller wants the position, not the text.
		_o_return((int)index + 1);

	LRESULT length = SendMessage(aControl.hwnd, CB_GETLBTEXTLEN, (WPARAM)index, 0);
	if (length == CB_ERR) // Given the way it was called, this should be impossible based on MSDN docs.
		_o_return_empty; // Return empty string (default).
	// In unusual cases, MSDN says the indicated length might be longer than it actually winds up
	// being when the item's text is retrieved.  This should be harmless, since there are many
	// other precedents where a variable is sized to something larger than it winds up carrying.
	if (!TokenSetResult(aResultToken, NULL, length))
		return; // It already displayed the error.
	length = SendMessage(aControl.hwnd, CB_GETLBTEXT, (WPARAM)index, (LPARAM)aResultToken.marker);
	if (length == CB_ERR) // Given the way it was called, this should be impossible based on MSDN docs.
		_o_return_p(_T(""), 0); // Must use this vs _o_return_empty to override TokenSetResult.
	_o_return_retval;
}


void GuiType::ControlGetListBox(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode)
{
	bool position_mode = (aMode == Value_Mode || (aMode == Submit_Mode && (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT)));
	if (GetWindowLong(aControl.hwnd, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
	{
		LRESULT sel_count = SendMessage(aControl.hwnd, LB_GETSELCOUNT, 0, 0);
		if (sel_count < 1)  // <=0 to check for LB_ERR too (but it should be impossible in this case).
			_o_return_empty;
		int *item = (int *)malloc(sel_count * sizeof(int)); // dynamic since there can be a very large number of items.
		if (!item)
			_o_throw_oom;
		sel_count = SendMessage(aControl.hwnd, LB_GETSELITEMS, (WPARAM)sel_count, (LPARAM)item);
		if (sel_count < 1)  // 0 or LB_ERR, but both these conditions should be impossible in this case.
		{
			free(item);
			_o_return_empty;
		}

		// Create object for storing the selected items.
		auto ret = Array::Create();
		if (!ret)
		{
			free(item);
			_o_throw_oom;
		}

		for (LRESULT length = sel_count - 1, i = 0; i < sel_count; ++i)
		{
			if (position_mode) // Caller wants the positions, not the text.
				ret->Append(item[i] + 1); // +1 to convert from zero-based to 1-based.
			else // Store item text vs. position.
			{
				LRESULT item_length = SendMessage(aControl.hwnd, LB_GETTEXTLEN, (WPARAM)item[i], 0);
				if (item_length == LB_ERR) // Realistically impossible based on MSDN.
				{
					free(item);
					ret->Release();
					_o_throw(_T("LB_GETTEXTLEN")); // Short msg since so rare.
				}
				LPTSTR temp = tmalloc(item_length+1);
				if (!temp)
				{
					free(item);
					ret->Release();
					_o_throw_oom; // Short msg since so rare.
				}
				LRESULT lr = SendMessage(aControl.hwnd, LB_GETTEXT, (WPARAM)item[i], (LPARAM)temp);
				if (lr > 0) // Given the way it was called, LB_ERR (-1) should be impossible based on MSDN docs.
					ret->Append(temp, item_length);
				free(temp);
			}
		}

		free(item);
		_o_return(ret);
	}
	else // Single-select ListBox style.
	{
		LRESULT index = SendMessage(aControl.hwnd, LB_GETCURSEL, 0, 0); // Get index of currently selected item.
		if (position_mode) // Caller wants the position, not the text.
			_o_return((int)index + 1); // 0 if no selection.
		if (index == LB_ERR) // There is no selection (or very rarely, some other type of problem).
			_o_return_empty;
		LRESULT length = SendMessage(aControl.hwnd, LB_GETTEXTLEN, (WPARAM)index, 0);
		if (length == LB_ERR // Given the way it was called, this should be impossible based on MSDN docs.
			|| length == 0)
			_o_return_empty;
		if (!TokenSetResult(aResultToken, NULL, length))
			return;  // It already displayed the error.
		length = SendMessage(aControl.hwnd, LB_GETTEXT, (WPARAM)index, (LPARAM)aResultToken.marker);
		if (length == LB_ERR) // Given the way it was called, this should be impossible based on MSDN docs.
		{
			*aResultToken.marker = 0; // Clear the buffer returned by TokenSetResult.
			aResultToken.marker_length = 0;
		}
	}
	_o_return_retval;
}


FResult GuiType::ControlSetEdit(GuiControlType &aControl, LPTSTR aContents, bool aIsText)
{
	// Auto-translate LF to CRLF if this is the "Value" property of a multi-line Edit control.
	// Note that TranslateLFtoCRLF() will return the original buffer we gave it if no translation
	// is needed.  Otherwise, it will return a new buffer which we are responsible for freeing
	// when done (or NULL if it failed to allocate the memory).
	LPTSTR malloc_buf = (!aIsText && (GetWindowLong(aControl.hwnd, GWL_STYLE) & ES_MULTILINE))
		? TranslateLFtoCRLF(aContents) : aContents; // Automatic translation, as documented.
	if (!malloc_buf)
		return FR_E_OUTOFMEM;
	// Users probably don't want or expect the control's event handler to be triggered by this
	// action, so suppress it.  This makes single-line Edits consistent with all other controls.
	aControl.attrib |= GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Disable events.
	// Set the new content.
	SetWindowText(aControl.hwnd,  malloc_buf);
	// WM_SETTEXT and WM_COMMAND are sent, not posted, so it's certain that the EN_CHANGE
	// notification (if any) has been received and discarded.  MsgSleep() is not needed.
	aControl.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Enable events.
	if (malloc_buf != aContents)
		free(malloc_buf);
	return OK;
}


void GuiType::ControlGetEdit(ResultToken &aResultToken, GuiControlType &aControl)
// Only called by Edit.Value, so aTranslateEOL == true is implied.
{
	if (!ControlGetWindowText(aResultToken, aControl))
		return;
	if (GetWindowLong(aControl.hwnd, GWL_STYLE) & ES_MULTILINE) // For consistency with ControlSetEdit().
	{
		// Auto-translate CRLF to LF for better compatibility with other script commands.
		// Since edit controls tend to have many hard returns in them, use SCS_SENSITIVE to
		// enhance performance.  This performance gain is extreme when the control contains
		// thousands of CRLFs:
		StrReplace(aResultToken.marker, _T("\r\n"), _T("\n"), SCS_SENSITIVE
			, -1, -1, NULL, &aResultToken.marker_length); // Pass &marker_length to improve performance and (more importantly) update it.
	}
}


FResult GuiType::ControlSetDateTime(GuiControlType &aControl, LPTSTR aContents)
{
	SYSTEMTIME st;
	if (*aContents)
	{
		if (YYYYMMDDToSystemTime(aContents, st, true))
			DateTime_SetSystemtime(aControl.hwnd, GDT_VALID, &st);
		else
			return FValueError(ERR_INVALID_VALUE, aContents);
	}
	else // User wants there to be no date selection.
	{
		// Ensure the DTS_SHOWNONE style is present, otherwise it won't work.  However,
		// it appears that this style cannot be applied after the control is created, so
		// this line is commented out:
		//SetWindowLong(aControl.hwnd, GWL_STYLE, GetWindowLong(aControl.hwnd, GWL_STYLE) | DTS_SHOWNONE);
		DateTime_SetSystemtime(aControl.hwnd, GDT_NONE, &st);  // Contents of st are ignored in this mode.
	}
	return OK;
}


FResult GuiControlType::DT_SetFormat(optl<StrArg> aFormat)
{
	CTRL_THROW_IF_DESTROYED;
	bool use_custom_format = false; // Set default.
	// Reset style to "pure" so that new style (or custom format) can take effect.
	DWORD style = GetWindowLong(hwnd, GWL_STYLE) // DTS_SHORTDATEFORMAT==0 so can be omitted below.
		& ~(DTS_LONGDATEFORMAT | DTS_SHORTDATECENTURYFORMAT | DTS_TIMEFORMAT);
	if (aFormat.has_nonempty_value())
	{
		// DTS_SHORTDATEFORMAT and DTS_SHORTDATECENTURYFORMAT
		// seem to produce identical results (both display 4-digit year), at least on XP.  Perhaps
		// DTS_SHORTDATECENTURYFORMAT is obsolete.  In any case, it's uncommon so for simplicity, is
		// not a named style.  It can always be applied numerically if desired.  Update:
		// DTS_SHORTDATECENTURYFORMAT is now applied by default upon creation, which can be overridden
		// explicitly via -0x0C in the control's options.
		if (!_tcsicmp(aFormat.value(), _T("LongDate"))) // LongDate seems more readable than "Long".  It also matches the keyword used by FormatTime.
			style |= DTS_LONGDATEFORMAT; // Competing styles were already purged above.
		else if (!_tcsicmp(aFormat.value(), _T("Time")))
			style |= DTS_TIMEFORMAT; // Competing styles were already purged above.
		else if (!_tcsicmp(aFormat.value(), _T("ShortDate")))
		{} // No change needed since DTS_SHORTDATEFORMAT is 0 and competing styles were already purged above.
		else // Custom format.
			use_custom_format = true;
	}
	//else aText is blank and use_custom_format==false, which will put DTS_SHORTDATEFORMAT into effect.
	if (!use_custom_format)
		SetWindowLong(hwnd, GWL_STYLE, style);
	//else leave style unchanged so that if format is later removed, the underlying named style will
	// not have been altered.
	DateTime_SetFormat(hwnd, use_custom_format ? aFormat.value() : NULL); // NULL removes any custom format so that the underlying style format is revealed.
	return OK;
}


void GuiType::ControlGetDateTime(ResultToken &aResultToken, GuiControlType &aControl)
{
	LPTSTR buf = aResultToken.buf;
	SYSTEMTIME st;
	_o_return_p(DateTime_GetSystemtime(aControl.hwnd, &st) == GDT_VALID
		? SystemTimeToYYYYMMDD(buf, st) : _T("")); // Blank string whenever GDT_NONE/GDT_ERROR.
}


FResult GuiType::ControlSetMonthCal(GuiControlType &aControl, LPTSTR aContents)
{
	SYSTEMTIME st[2];
	DWORD gdtr = YYYYMMDDToSystemTime2(aContents, st);
	if (!gdtr) // aContents is empty or invalid (it's not meaningful to set an empty value).
		return FValueError(ERR_INVALID_VALUE, aContents);
	if (GetWindowLong(aControl.hwnd, GWL_STYLE) & MCS_MULTISELECT) // Must use range-selection even if selection is only one date.
	{
		if (gdtr == GDTR_MIN) // No maximum is present, so set maximum to minimum.
			st[1] = st[0];
		//else both are present.
		MonthCal_SetSelRange(aControl.hwnd, st);
	}
	else
		MonthCal_SetCurSel(aControl.hwnd, st);
	ControlRedraw(aControl, true); // Confirmed necessary.
	return OK;
}


void GuiType::ControlGetMonthCal(ResultToken &aResultToken, GuiControlType &aControl)
{
	LPTSTR buf = aResultToken.buf;
	SYSTEMTIME st[2];
	if (GetWindowLong(aControl.hwnd, GWL_STYLE) & MCS_MULTISELECT)
	{
		// For code simplicity and due to the expected rarity of using the MonthCal control, much less
		// in its range-select mode, the range is returned with a dash between the min and max rather
		// than as an array or anything fancier.
		MonthCal_GetSelRange(aControl.hwnd, st);
		// Seems easier for script (due to consistency) to always return it in range format, even if
		// only one day is selected.
		SystemTimeToYYYYMMDD(buf, st[0]);
		buf[8] = '-'; // Retain only the first 8 chars to omit the time portion, which is unreliable (not relevant anyway).
		SystemTimeToYYYYMMDD(buf + 9, st[1]);
		buf[17] = 0; // Limit to 17 chars to omit the time portion of the second timestamp.
	}
	else
	{
		MonthCal_GetCurSel(aControl.hwnd, st);
		SystemTimeToYYYYMMDD(buf, st[0]);
		buf[8] = 0; // Limit to 8 chars to omit the time portion, which is unreliable (not relevant anyway).
	}
	_o_return_p(buf);
}


FResult GuiType::ControlSetHotkey(GuiControlType &aControl, LPTSTR aContents)
{
	SendMessage(aControl.hwnd, HKM_SETHOTKEY, TextToHotkey(aContents), 0); // This will set it to "None" if aContents is blank.
	return OK;
}


void GuiType::ControlGetHotkey(ResultToken &aResultToken, GuiControlType &aControl)
{
	LPTSTR buf = aResultToken.buf;
	// Testing shows that neither GetWindowText() nor WM_GETTEXT can pull anything out of a hotkey
	// control, so the only type of retrieval that can be offered is the HKM_GETHOTKEY method:
	HotkeyToText((WORD)SendMessage(aControl.hwnd, HKM_GETHOTKEY, 0, 0), buf);
	_o_return_p(buf);
}


FResult GuiType::ControlSetUpDown(GuiControlType &aControl, int new_pos)
{
	// MSDN: "If the parameter is outside the control's specified range, nPos will be set to the nearest
	// valid value."
	SendMessage(aControl.hwnd, UDM_SETPOS32, 0, new_pos);
	return OK;
}


void GuiType::ControlGetUpDown(ResultToken &aResultToken, GuiControlType &aControl)
{
	// Any out of range or non-numeric value in the buddy is ignored since error reporting is
	// left up to the script, which can compare contents of buddy to those of UpDown to check
	// validity if it wants.
	int pos = (int)SendMessage(aControl.hwnd, UDM_GETPOS32, 0, 0);
	_o_return(pos);
}


FResult GuiType::ControlSetSlider(GuiControlType &aControl, int aValue)
{
	// Confirmed this fact from MSDN: That the control automatically deals with out-of-range values
	// by setting slider to min or max:
	SendMessage(aControl.hwnd, TBM_SETPOS, TRUE, ControlInvertSliderIfNeeded(aControl, aValue));
	// Above msg has no return value.
	return OK;
}


void GuiType::ControlGetSlider(ResultToken &aResultToken, GuiControlType &aControl)
{
	_o_return(ControlInvertSliderIfNeeded(aControl, (int)SendMessage(aControl.hwnd, TBM_GETPOS, 0, 0)));
	// Above assigns it as a signed value because testing shows a slider can have part or all of its
	// available range consist of negative values.  32-bit values are supported if the range is set
	// with the right messages.
}


FResult GuiType::ControlSetProgress(GuiControlType &aControl, int aValue)
{
	// Confirmed through testing (PBM_DELTAPOS was also tested): The control automatically deals
	// with out-of-range values by setting bar to min or max.  
	SendMessage(aControl.hwnd, PBM_SETPOS, aValue, 0);
	return OK;
}


void GuiType::ControlGetProgress(ResultToken &aResultToken, GuiControlType &aControl)
{
	_o_return((int)SendMessage(aControl.hwnd, PBM_GETPOS, 0, 0));
}


void GuiType::ControlGetTab(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode)
{
	LPTSTR buf = aResultToken.buf;
	LRESULT index = TabCtrl_GetCurSel(aControl.hwnd); // Get index of currently selected item.
	if (aMode == Value_Mode || (aMode == Submit_Mode && (aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT))) // Caller wants the index, not the text.
		_o_return((int)index + 1); // 0 if no selection.
	if (index == -1) // There is no selection.
		_o_return_empty;
	// Otherwise: Get the stored name/caption of this tab:
	TCITEM tci;
	tci.mask = TCIF_TEXT;
	tci.pszText = buf;
	tci.cchTextMax = _f_retval_buf_size-1; // MSDN example uses -1.
	if (TabCtrl_GetItem(aControl.hwnd, index, &tci))
		_o_return_p(buf);
	_o_return_empty;
}


ResultType GuiType::ControlGetWindowText(ResultToken &aResultToken, GuiControlType &aControl)
{
	// Set up the result token.
	int length = GetWindowTextLength(aControl.hwnd); // Might be zero, which is properly handled below.
	if (TokenSetResult(aResultToken, NULL, length) != OK)
		return FAIL; // It already displayed the error.
	aResultToken.marker_length = GetWindowText(aControl.hwnd, aResultToken.marker, length+1);
	// MSDN: "If the function succeeds, the return value is the length, in characters,
	// of the copied string, not including the terminating null character."
	// However, GetWindowText() works by sending the control a WM_GETTEXT message, and some
	// controls don't respond correctly.  Link controls have been caught including the null
	// terminator in the count, so this works around it:
	if (aResultToken.marker_length && !aResultToken.marker[aResultToken.marker_length-1])
		--aResultToken.marker_length;
	return OK;
}


void GuiType::ControlRedraw(GuiControlType &aControl, bool aOnlyWithinTab)
{
	if (   !aOnlyWithinTab
		|| FindTabControl(aControl.tab_control_index) && IsWindowVisible(aControl.hwnd)   )
	{
		RECT rect;
		GetWindowRect(aControl.hwnd, &rect); // Limit it to only that part of the client area that is receiving the rect.
		MapWindowPoints(NULL, mHwnd, (LPPOINT)&rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
		InvalidateRect(mHwnd, &rect, TRUE); // Seems safer to use TRUE, not knowing all possible overlaps, etc.
		//Overkill: InvalidateRect(mHwnd, NULL, FALSE); // Erase doesn't seem to be necessary.
		// None of the following is enough:
		//Changes focused control, so no good: ControlUpdateCurrentTab(*tab_control, false);
		//RedrawWindow(tab_control->hwnd, NULL, NULL, 0 ..or.. RDW_INVALIDATE);
		//InvalidateRect(aControl.hwnd, NULL, TRUE);
		//InvalidateRect(tab_control->hwnd, NULL, TRUE);
	}
}



/////////////////
// Static members
/////////////////
FontType *GuiType::sFont = NULL; // An array of structs, allocated upon first use.
int GuiType::sFontCount = 0;
HWND GuiType::sTreeWithEditInProgress = NULL;



bool GuiType::Delete() // IObject::Delete()
{
	// Delete() should only be called when the script releases its last reference to the
	// object, or when the user closes the window while the script has no references.
	// See VisibilityChanged() for details.
	if (mHwnd)
		Destroy();
	else
		Dispose();
	// OnMessage() or perhaps some other callback may enable the script to regain a ref.
	// Although the object is now unusable, for program stability we must `delete this`
	// only if mRefCount is still 1 (it's never decremented to 0).
	if (mRefCount > 1)
		return false;
	return Object::Delete();
}



FResult GuiType::Destroy()
// Destroys the window and performs related cleanup which is only necessary for
// a successfully constructed Gui, then calls Dispose() for the remaining cleanup.
{
	if (!mHwnd)
		return OK; // We have already been destroyed.

	// First destroy any windows owned by this window, since they will be auto-destroyed
	// anyway due to their being owned.  By destroying them explicitly, the Destroy()
	// function is called recursively which keeps everything relatively neat.
	// Lexikos: I'm not really sure what Chris meant by "neat", since Destroy() is called
	// recursively via WM_DESTROY anyway.  It's currently left to WM_DESTROY because the
	// code below is actually broken in two ways (unique to the new object-oriented API):
	//  1) If there are no external references to gui, it will be deleted and gui->mNextGui
	//     will likely crash the program.
	//  2) Destroy() removes the gui from the list, which causes the loop to break at the
	//     first owned/child window.  Saving the next/previous link before-hand may not
	//     solve the problem since Destroy() could recursively destroy other windows.
	//for (GuiType* gui = g_firstGui; gui; gui = gui->mNextGui)
	//	if (gui->mOwner == mHwnd)
	//		gui->Destroy();

	// Testing shows that this must be done prior to calling DestroyWindow() later below, presumably
	// because the destruction immediately destroys the status bar, or prevents it from answering messages.
	// MSDN says "During the processing of [WM_DESTROY], it can be assumed that all child windows still exist".
	// However, for DestroyWindow() to have returned means that WM_DESTROY was already fully processed.
	if (mStatusBarHwnd) // IsWindow(gui.mStatusBarHwnd) isn't called because even if possible for it to have been destroyed, SendMessage below should return 0.
	{
		// This is done because the vast majority of people wouldn't want to have to worry about it.
		// They can always use DllCall() if they want to share the same HICON among multiple parts of
		// the same bar, or among different windows (fairly rare).
		HICON hicon;
		int part_count = (int)SendMessage(mStatusBarHwnd, SB_GETPARTS, 0, NULL); // MSDN: "This message always returns the number of parts in the status bar [regardless of how it is called]".
		for (int i = 0; i < part_count; ++i)
			if (hicon = (HICON)SendMessage(mStatusBarHwnd, SB_GETICON, i, 0))
				DestroyIcon(hicon);
	}

	if (IsWindow(mHwnd)) // If WM_DESTROY called us, the window might already be partially destroyed.
	{
		// If this window is using a menu bar but that menu is also used by some other window, first
		// detach the menu so that it doesn't get auto-destroyed with the window.  This is done
		// unconditionally since such a menu will be automatically destroyed when the script exits
		// or when the menu is destroyed explicitly via the Menu command.  It also prevents any
		// submenus attached to the menu bar from being destroyed, since those submenus might be
		// also used by other menus (however, this is not really an issue since menus destroyed
		// would be automatically re-created upon next use).  But in the case of a window that
		// is currently using a menu bar, destroying that bar in conjunction with the destruction
		// of some other window might cause bad side effects on some/all OSes.
		ShowWindow(mHwnd, SW_HIDE);  // Hide it to prevent re-drawing due to menu removal.
		::SetMenu(mHwnd, NULL);
		if (!mDestroyWindowHasBeenCalled)
		{
			mDestroyWindowHasBeenCalled = true;  // Signal the WM_DESTROY routine not to call us.
			DestroyWindow(mHwnd);  // The WindowProc is immediately called and it now destroys the window.
		}
		// else WM_DESTROY was called by a function other than this one (possibly auto-destruct due to
		// being owned by script's main window), so it would be bad to call DestroyWindow() again since
		// it's already in progress.
	}

	// This is done here rather than in Dispose() because a Gui is added to the list only
	// after being successfully constructed, and should remain there only until the window
	// is destroyed:
	RemoveGuiFromList(this);

	// It may be possible for an object released below to cause execution of script, and
	// although the window has been destroyed (making it impossible for FindGui() to return
	// this object) and the Gui has been removed from the list, the script might still have
	// a reference to it.  It is therefore important that this is done to mark the Gui as
	// invalid, so Invoke() raises an error:
	mHwnd = NULL;

	// Clean up final resources.
	Dispose();

	// The following might release the final reference to this object, thereby causing 'this'
	// to be deleted.  See VisibilityChanged() for details.
	if (mVisibleRefCounted)
		Release();
	// IT IS NOW UNSAFE TO REFER TO ANY NON-STATIC MEMBERS OF THIS OBJECT.

	// If this Gui was the last thing keeping the script running, exit the script:
	g_script.ExitIfNotPersistent(EXIT_CLOSE);
	return OK;
}



void GuiType::Dispose()
// Cleans up resources managed separately to the window.  This is separate from Destroy()
// for maintainability; i.e. to ensure resources are freed even if the window had not been
// successfully created.
{
	if (mDisposed)
		return;
	mDisposed = true;

	if (mBackgroundBrushWin)
		DeleteObject(mBackgroundBrushWin);
	if (mHdrop)
		DragFinish(mHdrop);
	
	// A partially constructed Gui can't have a non-zero mControlCount, but it seems more
	// appropriate to do this here than in Destroy():
	for (GuiIndexType u = 0; u < mControlCount; ++u)
	{
		mControl[u]->Dispose();
		mControl[u]->Release(); // Release() vs delete because the script may still have references to it.
	}
	// The following has been confirmed necessary to prevent the script from crashing if
	// the controls are being enumerated (e.g. by GuiTypeEnum) when the Gui is destroyed:
	mControlCount = 0;

	if (mMenu)
		mMenu->Release();

	mEvents.Dispose();
	if (mEventSink && this != mEventSink)
		mEventSink->Release();

	if (mIconEligibleForDestruction && mIconEligibleForDestruction != g_script.mCustomIcon) // v1.0.37.07.
		DestroyIconsIfUnused(mIconEligibleForDestruction, mIconEligibleForDestructionSmall);

	// For simplicity and performance, any fonts used *solely* by a destroyed window are destroyed
	// only when the program terminates.  Another reason for this is that sometimes a destroyed window
	// is soon recreated to use the same fonts it did before.

	RemoveAccelerators();
	free(mControl); // Free the control array, which was previously malloc'd.
}


void GuiControlType::Dispose()
// Clean up resources associated with this control.
// It's called Dispose vs Destroy because at the time it is called, the control window
// has already been destroyed as a side-effect of the parent window being destroyed.
{
	if (type == GUI_CONTROL_PIC && union_hbitmap)
	{
		if (attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR)
			DestroyIcon((HICON)union_hbitmap); // Works on cursors too.  See notes in LoadPicture().
		else // union_hbitmap is a bitmap rather than an icon or cursor.
			DeleteObject(union_hbitmap);
	}
	else if (type == GUI_CONTROL_LISTVIEW) // It was ensured at an earlier stage that union_lv_attrib != NULL.
		free(union_lv_attrib);
	else if (type == GUI_CONTROL_ACTIVEX && union_object)
		union_object->Release();
	//else do nothing, since this type has nothing more than a color stored in the union.
	if (background_brush)
		DeleteObject(background_brush);
	events.Dispose();
	if (name)
		free(name);
	name = NULL;
	hwnd = NULL;
	gui = NULL;
}



void GuiType::DestroyIconsIfUnused(HICON ahIcon, HICON ahIconSmall)
// Caller has ensured that the GUI window previously using ahIcon has been destroyed prior to calling
// this function.
{
	if (!ahIcon) // At least one caller relies on this check.
		return;
	for (GuiType* gui = g_firstGui; gui; gui = gui->mNextGui)
	{
		// If another window is using this icon, don't destroy the because that has been reported to disrupt
		// the window's display of the icon in some cases (apparently WM_SETICON doesn't make a copy of the
		// icon).  The windows still using the icon will be responsible for destroying it later.
		if (gui->mIconEligibleForDestruction == ahIcon)
			return;
	}
	// Since above didn't return, this icon is not currently in use by a GUI window.  The caller has
	// authorized us to destroy it.
	DestroyIcon(ahIcon);
	// L17: Small icon should always also be unused at this point.
	if (ahIconSmall != ahIcon)
		DestroyIcon(ahIconSmall);
}



FResult GuiType::Create(LPCTSTR aTitle)
{
	if (mHwnd) // It already exists
		return FR_E_FAILED; // Should be impossible since mHwnd is checked by caller.

	// Use a separate class for GUI, which gives it a separate WindowProc and allows it to be more
	// distinct when used with the ahk_class method of addressing windows.
	if (!sGuiWinClass)
	{
		WNDCLASSEX wc = {0};
		wc.cbSize = sizeof(wc);
		wc.lpszClassName = WINDOW_CLASS_GUI;
		wc.hInstance = g_hInstance;
		wc.lpfnWndProc = GuiWindowProc;
		wc.hIcon = g_IconLarge;
		wc.hIconSm = g_IconSmall;
		wc.style = CS_DBLCLKS; // v1.0.44.12: CS_DBLCLKS is accepted as a good default by nearly everyone.  It causes the window to receive WM_LBUTTONDBLCLK, WM_RBUTTONDBLCLK, and WM_MBUTTONDBLCLK (even without this, all windows receive WM_NCLBUTTONDBLCLK, WM_NCMBUTTONDBLCLK, and WM_NCRBUTTONDBLCLK).
			// CS_HREDRAW and CS_VREDRAW are not included above because they cause extra flickering.  It's generally better for a window to manage its own redrawing when it's resized.
		wc.hCursor = LoadCursor((HINSTANCE) NULL, IDC_ARROW);
		wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
		wc.cbWndExtra = DLGWINDOWEXTRA;  // So that it will be the type that uses DefDlgProc() vs. DefWindowProc().
		sGuiWinClass = RegisterClassEx(&wc);
		if (!sGuiWinClass)
			return FR_E_WIN32;
	}

	if (   !(mHwnd = CreateWindowEx(mExStyle, WINDOW_CLASS_GUI, aTitle
		, mStyle, 0, 0, 0, 0, mOwner, NULL, g_hInstance, NULL))   )
		return FR_E_WIN32;

	// Set the user pointer in the window to this GuiType object, so that it is possible to retrieve it back from the window handle.
	SetWindowLongPtr(mHwnd, GWLP_USERDATA, (LONG_PTR)this);

	// L17: Use separate big/small icons for best results.
	HICON big_icon, small_icon;
	if (g_script.mCustomIcon)
	{
		mIconEligibleForDestruction = big_icon = g_script.mCustomIcon;
		mIconEligibleForDestructionSmall = small_icon = g_script.mCustomIconSmall; // Should always be non-NULL if mCustomIcon is non-NULL.
	}
	else
	{
		big_icon = g_IconLarge;
		small_icon = g_IconSmall;
		// Unlike mCustomIcon, leave mIconEligibleForDestruction NULL because these
		// icons are used for various other purposes and should never be destroyed.
	}
	// Setting the small icon puts it in the upper left corner of the dialog window.
	// Setting the big icon makes the dialog show up correctly in the Alt-Tab menu (but big seems to
	// have no effect unless the window is unowned, i.e. it has a button on the task bar).
	// Change for v1.0.37.07: Set both icons unconditionally for code simplicity, and also in case
	// it's possible for the window to change after creation in a way that would make a custom icon
	// become relevant.  Set the big icon even if it's owned because there might be ways
	// an owned window can have an entry in the alt-tab menu.  The following ways come close
	// but don't actually succeed:
	// 1) It's owned by the main window but the main window isn't visible: It acquires the main window's icon
	//    in the alt-tab menu regardless of whether it was given a big icon of its own.
	// 2) It's owned by another GUI window but it has the WS_EX_APPWINDOW style (might force a taskbar button):
	//    Same effect as in #1.
	// 3) Possibly other ways.
	SendMessage(mHwnd, WM_SETICON, ICON_SMALL, (LPARAM)small_icon); // Testing shows that a zero is returned for both;
	SendMessage(mHwnd, WM_SETICON, ICON_BIG, (LPARAM)big_icon);   // i.e. there is no previous icon to destroy in this case.

	return OK;
}


FResult GuiType::OnEvent(StrArg aEventName, ExprTokenType &aCallback, optl<int> aAddRemove)
{
	GUI_MUST_HAVE_HWND;
	
	GuiEventType evt = ConvertEvent(aEventName);
	if (evt < GUI_EVENT_WINDOW_FIRST || evt > GUI_EVENT_WINDOW_LAST)
		return FR_E_ARG(0);
	
	return OnEvent(nullptr, evt, GUI_EVENTKIND_EVENT, aCallback, aAddRemove);
}


FResult GuiType::OnMessage(UINT aNumber, ExprTokenType &aCallback, optl<int> aAddRemove)
{
	GUI_MUST_HAVE_HWND;
	
	return OnEvent(nullptr, aNumber, GUI_EVENTKIND_MESSAGE, aCallback, aAddRemove);
}


FResult GuiType::OnEvent(GuiControlType *aControl, UINT aEvent, UCHAR aEventKind
	, ExprTokenType &aCallback, optl<int> aAddRemove)
{
	int max_threads = 1;
	if (aAddRemove.has_value())
		max_threads = aAddRemove.value();
	if (max_threads < -1 || max_threads > 1)
		// An event can currently only run one thread at a time.  By contrast, OnMessage
		// applies the thread limit per handler, not per message.  It's hard to say which
		// approach is better.
		return FR_E_ARG(2);

	TCHAR nbuf[MAX_NUMBER_SIZE];
	IObject *func = TokenToObject(aCallback);
	LPTSTR name = func ? nullptr : TokenToString(aCallback, nbuf);
	if (!func && !mEventSink)
		return FR_E_ARG(1);
	
	return OnEvent(aControl, aEvent, aEventKind, func, name, max_threads);
}


FResult GuiType::OnEvent(GuiControlType *aControl, UINT aEvent, UCHAR aEventKind
	, IObject *aFunc, LPTSTR aMethodName, int aMaxThreads)
{
	MsgMonitorList &handlers = aControl ? aControl->events : mEvents;
	MsgMonitorStruct *mon;
	if (aFunc)
		mon = handlers.Find(aEvent, aFunc, aEventKind);
	else
		mon = handlers.Find(aEvent, aMethodName, aEventKind);
	if (!aMaxThreads)
	{
		if (mon)
			handlers.Delete(mon);
		ApplyEventStyles(aControl, aEvent, false);
		return OK;
	}
	bool append = aMaxThreads >= 0;
	if (!append) // Negative values mean "call it last".
		aMaxThreads = -aMaxThreads; // Make it positive.
	if (aMaxThreads > MsgMonitorStruct::MAX_INSTANCES)
		aMaxThreads = MsgMonitorStruct::MAX_INSTANCES;
	if (!mon)
	{
		if (aFunc)
		{
			// Since this callback is being registered for the first time, validate.
			int param_count = 2;
			switch (aEventKind)
			{
			case GUI_EVENTKIND_MESSAGE:		param_count = 4; break;
			case GUI_EVENTKIND_COMMAND:		param_count = 1; break;
			case GUI_EVENTKIND_EVENT:
				switch (aEvent)
				{
				case GUI_EVENT_DROPFILES:	param_count = 5; break;
				case GUI_EVENT_CLOSE:
				case GUI_EVENT_ESCAPE:		param_count = 1; break;
				case GUI_EVENT_RESIZE:		param_count = 4; break;
				case GUI_EVENT_CONTEXTMENU:	param_count = 5 + (aControl == nullptr); break;
				case GUI_EVENT_CLICK:		param_count = 2 + (aControl->type == GUI_CONTROL_LINK); break;
				case GUI_EVENT_ITEMSELECT:
					if (aControl->type == GUI_CONTROL_TREEVIEW)
						break;
				case GUI_EVENT_ITEMCHECK:
				case GUI_EVENT_ITEMEXPAND:
					param_count = 3;
					break;
				}
			}
			auto fr = ValidateFunctor(aFunc, param_count);
			if (fr != OK)
				return fr;
			// Add the callback.
			mon = handlers.Add(aEvent, aFunc, append);
		}
		else
			mon = handlers.Add(aEvent, aMethodName, append);
		if (!mon)
			return FR_E_OUTOFMEM;
	}
	mon->instance_count = 0;
	mon->max_instances = aMaxThreads;
	mon->msg_type = aEventKind;
	ApplyEventStyles(aControl, aEvent, true);
	return OK;
}


void GuiType::ApplyEventStyles(GuiControlType *aControl, UINT aEvent, bool aAdded)
{
	int style_type;
	DWORD style_bit;
	HWND hwnd;
	if (aControl)
	{
		// These styles are only added automatically, not removed, for a few reasons:
		//  1) Reduces code size (by 224 bytes on x86 vs. just 64 bytes for this section).
		//     Note that it's necessary to check for handlers for every event in the group
		//     before removing the style.
		//  2) The script may have applied the style manually, in which case it should not
		//     be removed automatically (but we have no way to know).
		//  3) It's probably rare to remove a handler for one of these events.
		if (!aAdded)
			return;
		GuiControlType &control = *aControl;
		// This approach appears to produce the same code size as just checking for the
		// events which don't need the style; i.e. GUI_EVENT_CONTEXTMENU and (for Buttons)
		// GUI_EVENT_CLICK:
		static char sStaticNotify[] = {GUI_EVENT_CLICK, GUI_EVENT_DBLCLK, NULL};
		static char sButtonNotify[] = {GUI_EVENT_DBLCLK, GUI_EVENT_FOCUS, GUI_EVENT_LOSEFOCUS, NULL};
		char *event_group;
		switch (control.type)
		{
		case GUI_CONTROL_TEXT:
		case GUI_CONTROL_PIC:
			// Apply the SS_NOTIFY style *only* if the control actually has an associated action.
			// This is because otherwise the control would steal all clicks for any other controls
			// drawn on top of it (e.g. a picture control with some edit fields drawn on top of it).
			// See comments in the creation of GUI_CONTROL_PIC for more info:
			style_bit = SS_NOTIFY;
			event_group = sStaticNotify;
			break;
		
		case GUI_CONTROL_BUTTON:
		case GUI_CONTROL_CHECKBOX:
		case GUI_CONTROL_RADIO:
			style_bit = BS_NOTIFY;
			event_group = sButtonNotify;
			break;
		
		default:
			return;
		}
		if (!strchr(event_group, aEvent))
			return;
		style_type = GWL_STYLE;
		hwnd = control.hwnd;
	}
	else // aControl == NULL
	{
		if (aEvent == GUI_EVENT_DROPFILES)
		{
			// Apply the WS_EX_ACCEPTFILES style to ensure the window can accept dropped files.
			// Remove the style to ensure the mouse cursor correctly changes to show that the
			// window no longer accepts dropped files.
			style_type = GWL_EXSTYLE;
			style_bit = WS_EX_ACCEPTFILES;
		}
		else
			return;
		if (!aAdded && IsMonitoring(aEvent)) // Removed a handler but not the last.
			return;
		hwnd = mHwnd;
	}
	// Since above didn't return, there's a style to apply or remove.
	DWORD current_style = GetWindowLong(hwnd, style_type);
	if (aAdded != ((current_style & style_bit) != 0)) // Needs toggling.
	{
		current_style ^= style_bit; // Toggle it.
		SetWindowLong(hwnd, style_type, current_style);
	}
}



LPTSTR GuiType::sEventNames[] = GUI_EVENT_NAMES;

LPTSTR GuiType::ConvertEvent(GuiEventType evt)
{
	if (evt < _countof(sEventNames))
		return sEventNames[evt];

	// Else it's a character code - convert it to a string
	static TCHAR sBuf[2] = { 0, 0 };
	sBuf[0] = (TCHAR)(UCHAR)evt;
	return sBuf;
}

GuiEventType GuiType::ConvertEvent(LPCTSTR evt)
{
	for (GuiEventType i = 0; i < _countof(sEventNames); ++i)
		if (!_tcsicmp(sEventNames[i], evt))
			return i;
	if (*evt && !evt[1])
		return *evt; // Single-letter event name.
	return GUI_EVENT_NONE;
}



IObject* GuiType::CreateDropArray(HDROP hDrop)
{
	auto arr = Array::Create();
	// The added complexity/code size of handling out-of-memory here and in MsgSleep()
	// by aborting the new thread seems to outweigh the trivial benefits.  So assume
	// arr is non-NULL and let the access violation occur if otherwise.
	TCHAR buf[T_MAX_PATH];
	UINT file_count = DragQueryFile(hDrop, 0xFFFFFFFF, NULL, 0);
	for (UINT u = 0; u < file_count; u ++)
	{
		arr->Append(buf, DragQueryFile(hDrop, u, buf, _countof(buf)));
		// Result not checked for reasons described above.  The likely outcome at this
		// point would be for the event handler to either execute with fewer files than
		// expected, or for it to raise ERR_OUTOFMEM.
	}
	return arr;
}



void GuiType::UpdateMenuBars(HMENU aMenu)
// Caller has changed aMenu and wants the change visibly reflected in any windows that that
// use aMenu as a menu bar.  For example, if a menu item has been disabled, the grey-color
// won't show up immediately unless the window is refreshed.
{
	for (GuiType* gui = g_firstGui; gui; gui = gui->mNextGui)
	{
		//if (gui->mHwnd) // Always non-NULL for any item in the list.
		if (GetMenu(gui->mHwnd) == aMenu && IsWindowVisible(gui->mHwnd))
		{
			// Neither of the below two calls by itself is enough for all types of changes.
			// Thought it's possible that every type of change only needs one or the other, both
			// are done for simplicity:
			// This first line is necessary at least for cases where the height of the menu bar
			// (the number of rows needed to display all its items) has changed as a result
			// of the caller's change.  In addition, I believe SetWindowPos() must be called
			// before RedrawWindow() to prevent artifacts in some cases:
			SetWindowPos(gui->mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
			// This line is necessary at least when a single new menu item has been added:
			RedrawWindow(gui->mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
			// RDW_UPDATENOW: Seems best so that the window is in an visible updated state when function
			// returns.  This is because if the menu bar happens to be two rows or its size is changed
			// in any other way, the window dimensions themselves might change, and the caller might
			// rely on such a change being visibly finished for PixelGetColor, etc.
			//Not enough: UpdateWindow(gui->mHwnd);
		}
	}
}



ResultType GuiType::AddControl(GuiControls aControlType, LPCTSTR aOptions, LPCTSTR aText, GuiControlType*& apControl, Array *aObj)
// Caller must have ensured that mHwnd is non-NULL (i.e. that the parent window already exists).
{
	apControl = NULL; // Initialize return pointer.

	#define TOO_MANY_CONTROLS _T("Too many controls.") // Short msg since so rare.
	if (mControlCount >= MAX_CONTROLS_PER_GUI)
		return g_script.RuntimeError(TOO_MANY_CONTROLS);
	if (mControlCount >= mControlCapacity) // The section below on the above check already having been done.
	{
		// realloc() to keep the array contiguous, which allows better-performing methods to be
		// used to access the list of controls in various places.
		// Expand the array by one block:
		GuiControlType **realloc_temp;  // Needed since realloc returns NULL on failure but leaves original block allocated.
		if (   !(realloc_temp = (GuiControlType **)realloc(mControl  // If passed NULL, realloc() will do a malloc().
			, (mControlCapacity + GUI_CONTROL_BLOCK_SIZE) * sizeof(GuiControlType*)))   )
			return g_script.RuntimeError(TOO_MANY_CONTROLS); // A non-specific msg since this error is so rare.
		mControl = realloc_temp;
		mControlCapacity += GUI_CONTROL_BLOCK_SIZE;
	}

	////////////////////////////////////////////////////////////////////////////////////////
	// Set defaults for the various options, to be overridden individually by any specified.
	////////////////////////////////////////////////////////////////////////////////////////
	GuiControlType *pcontrol = new GuiControlType(this);
	mControl[mControlCount] = pcontrol;
	GuiControlType &control = *pcontrol;

	control.SetBase(GuiControlType::GetPrototype(aControlType));
	
	GuiControlOptionsType opt;
	ControlInitOptions(opt, control);
	// aOpt.checked is already okay since BST_UNCHECKED == 0
	// Similarly, the zero-init of "control" higher above set the right values for password_char, new_section, etc.

	if (aControlType == GUI_CONTROL_TAB2) // v1.0.47.05: Replace TAB2 with TAB at an early stage to simplify the code.  The only purpose of TAB2 is to flag this as the new type of tab that avoids redrawing issues but has a new z-order that would break some existing scripts.
	{
		aControlType = GUI_CONTROL_TAB;
		control.attrib |= GUI_CONTROL_ATTRIB_ALTBEHAVIOR; // v1.0.47.05: A means for new scripts to solve redrawing problems in tab controls at the cost of putting the tab control after its controls in the z-order.
	}
	else if (aControlType == GUI_CONTROL_TAB3)
	{
		aControlType = GUI_CONTROL_TAB;
		opt.tab_control_uses_dialog = true;
		opt.tab_control_autosize = TAB3_AUTOWIDTH | TAB3_AUTOHEIGHT;
	}
	if (aControlType == GUI_CONTROL_TAB)
	{
		if (mTabControlCount == MAX_TAB_CONTROLS)
		{
			delete pcontrol;
			return g_script.RuntimeError(_T("Too many tab controls.")); // Short msg since so rare.
		}
		// For now, don't allow a tab control to be create inside another tab control because it raises
		// doubt and probably would create complications.  If it ever is allowed, note that
		// control.tab_index stores this tab control's index (0 for the first tab control, 1 for the
		// second, etc.) -- this is done for performance reasons.
		control.tab_control_index = MAX_TAB_CONTROLS;
		control.tab_index = mTabControlCount; // Store its control-index to help look-up performance in other sections.
		// Autosize any previous tab control, since this new one is going to become current.
		// This is done here rather than when mCurrentTabControlIndex is actually changed later,
		// because this new control's positioning might be dependent on the previous control's size.
		if (GuiControlType *previous_tab_control = FindTabControl(mCurrentTabControlIndex))
		{
			// See "case GUI_CMD_TAB:" for more comments.
			AutoSizeTabControl(*previous_tab_control);
		}
	}
	else if (aControlType == GUI_CONTROL_STATUSBAR)
	{
		if (mStatusBarHwnd)
		{
			delete pcontrol;
			return g_script.RuntimeError(_T("Too many status bars.")); // Short msg since so rare.
		}
		control.tab_control_index = MAX_TAB_CONTROLS; // Indicate that bar isn't owned by any tab control.
		// No need to do the following because constructor did it:
		//control.tab_index = 0; // Ignored but set for maintainability/consistency.
	}
	else
	{
		control.tab_control_index = mCurrentTabControlIndex;
		control.tab_index = mCurrentTabIndex;
	}

	// If this is the first control, set the default margin for the window based on the size
	// of the current font, but only if the margins haven't already been set:
	if (!mControlCount)
	{
		SetDefaultMargins();
		mPrevX = mMarginX;  // This makes first control be positioned correctly if it lacks both X & Y coords.
	}

	control.type = aControlType; // Improves maintainability to do this early, but must be done after TAB2 vs. TAB is resolved higher above.

	/////////////////////////////////////////////////
	// Set control-specific defaults for any options.
	/////////////////////////////////////////////////
	opt.style_add |= WS_VISIBLE;  // Starting default for all control types.
	opt.use_theme = mUseTheme; // Set default.

	// Radio buttons are handled separately here, outside the switch() further below:
	if (aControlType == GUI_CONTROL_RADIO)
	{
		// The BS_NOTIFY style is probably better not applied by default to radios because although it
		// causes the control to send BN_DBLCLK messages, each double-click by the user is seen only
		// as one click for the purpose of cosmetically making the button appear like it is being
		// clicked rapidly.  It is applied by ApplyEventStyles() instead.
		// v1.0.47.04: Removed BS_MULTILINE from default because it is conditionally applied later below.
		// No WS_TABSTOP here since that is applied elsewhere depending on radio group nature.
		if (!mInRadioGroup)
			opt.style_add |= WS_GROUP; // Tabstop must be handled later below.
			// The mInRadioGroup flag will be changed accordingly after the control is successfully created.
		//else by default, no WS_TABSTOP or WS_GROUP.  However, WS_GROUP can be applied manually via the
		// options list to split this radio group off from one immediately prior to it.
	}
	else // Not a radio.
		if (mInRadioGroup) // Close out the prior radio group by giving this control the WS_GROUP style.
			opt.style_add |= WS_GROUP; // This might not be necessary on all OSes, but it seems traditional / best-practice.

	switch (aControlType) // Set starting defaults based on control type (the above also does some of that).
	{
	// Some controls also have the WS_EX_CLIENTEDGE exstyle by default because they look pretty strange
	// without them.  This seems to be the standard default used by most applications.
	// Note: It seems that WS_BORDER is hardly ever used in practice with controls, just parent windows.
	case GUI_CONTROL_DROPDOWNLIST:
		opt.style_add |= WS_TABSTOP|WS_VSCROLL;  // CBS_DROPDOWNLIST is forcibly applied later. WS_VSCROLL is necessary.
		break;
	case GUI_CONTROL_COMBOBOX:
		// CBS_DROPDOWN is set as the default here to allow the flexibility for it to be changed to
		// CBS_SIMPLE.  CBS_SIMPLE is allowed for ComboBox but not DropDownList because CBS_SIMPLE
		// has an edit control just like a combo, which DropDownList isn't equipped to handle via Submit().
		// Also, if CBS_AUTOHSCROLL is omitted, typed text cannot go beyond the visible width of the
		// edit control, so it seems best to have that as a default also:
		opt.style_add |= WS_TABSTOP|WS_VSCROLL|CBS_AUTOHSCROLL|CBS_DROPDOWN;  // WS_VSCROLL is necessary.
		break;
	case GUI_CONTROL_LISTBOX:
		// Omit LBS_STANDARD because it includes LBS_SORT, which we don't want as a default style.
		// However, as of v1.0.30.03, LBS_USETABSTOPS is included by default because:
		// 1) Not doing so seems to make it impossible to apply tab stops after the control has been created.
		// 2) Without this style, tabs appears as empty squares in the text, which seems undesirable for
		//    99.9% of applications.
		// 3) LBS_USETABSTOPS can be explicitly removed by specifying -0x80 in the options of "Gui Add".
		opt.style_add |= WS_TABSTOP|WS_VSCROLL|LBS_USETABSTOPS;  // WS_VSCROLL seems the most desirable default.
		opt.exstyle_add |= WS_EX_CLIENTEDGE;
		break;
	case GUI_CONTROL_LISTVIEW:
		// The ListView extended styles are actually an entirely separate class of styles that exist
		// separately from ExStyles.  This explains why Get/SetWindowLong doesn't work on them.
		// But keep in mind that some of the normal/classic extended styles can still be applied
		// to a ListView via Get/SetWindowLong.
		// The listview extended styles all require at least ComCtl32.dll 4.70 (some might require more)
		// and thus will have no effect in Win 95/NT unless they have MSIE 3.x or similar patch installed.
		// Thus, things like LVS_EX_FULLROWSELECT and LVS_EX_HEADERDRAGDROP will have no effect on those systems.
		opt.listview_style |= LVS_EX_FULLROWSELECT|LVS_EX_HEADERDRAGDROP; // LVS_AUTOARRANGE seems to disrupt the display of the column separators and have other weird effects in Report view.
		opt.style_add |= WS_TABSTOP|LVS_SHOWSELALWAYS; // LVS_REPORT is omitted to help catch bugs involving opt.listview_view.  WS_THICKFRAME allows the control itself to be drag-resized.
		opt.exstyle_add |= WS_EX_CLIENTEDGE; // WS_EX_STATICEDGE/WS_EX_WINDOWEDGE/WS_BORDER(non-ex) don't look as nice. WS_EX_DLGMODALFRAME is a weird but interesting effect.
		opt.listview_view = LV_VIEW_DETAILS; // Improves maintainability by avoiding the need to check if it's -1 in other places.
		break;
	case GUI_CONTROL_TREEVIEW:
		// Default style is somewhat debatable, but the familiarity of Explorer's own defaults seems best.
		// TVS_SHOWSELALWAYS seems preferable by most people, and is also consistent to what is used for ListView.
		// Lines and buttons also seem preferable because the main feature of a tree is its hierarchical nature,
		// and that nature isn't well revealed without buttons, and buttons can't be shown at the root level
		// without TVS_LINESATROOT, which in turn can't be active without TVS_HASLINES.
		opt.style_add |= WS_TABSTOP|TVS_SHOWSELALWAYS|TVS_HASLINES|TVS_LINESATROOT|TVS_HASBUTTONS; // TVS_LINESATROOT is necessary to get plus/minus buttons on root-level items.
		opt.exstyle_add |= WS_EX_CLIENTEDGE; // Debatable, but seems best for consistency with ListView.
		break;
	case GUI_CONTROL_EDIT:
		opt.style_add |= WS_TABSTOP;
		opt.exstyle_add |= WS_EX_CLIENTEDGE;
		break;		
	case GUI_CONTROL_LINK:
		opt.style_add |= WS_TABSTOP;
		break;
	case GUI_CONTROL_UPDOWN:
		// UDS_NOTHOUSANDS is debatable:
		// 1) The primary means by which a script validates whether the buddy contains an invalid
		//    or out-of-range value for its UpDown is to compare the contents of the two.  If one
		//    has commas and the other doesn't, the commas must first be removed before comparing.
		// 2) Presence of commas in numeric data is going to be a source of script bugs for those
		//    who take the buddy's contents rather than the UpDown's contents as the user input.
		//    However, you could argue that script is not proper if it does this blindly because
		//    the buddy could contain an out-of-range or non-numeric value.
		// 3) Display is more ergonomic if it has commas in it.
		// The above make it pretty hard to decide, so sticking with the default of have commas
		// seems ok.  Also, UDS_ALIGNRIGHT must be present by default because otherwise buddying
		// will not take effect correctly.
		opt.style_add |= UDS_SETBUDDYINT|UDS_ALIGNRIGHT|UDS_AUTOBUDDY|UDS_ARROWKEYS;
		break;
	case GUI_CONTROL_DATETIME: // Gets a tabstop even when it contains an empty checkbox indicating "no date".
		// DTS_SHORTDATECENTURYFORMAT is applied by default because it should make results more consistent
		// across old and new systems.  This is because new systems display a 4-digit year even without
		// this style, but older ones might display a two digit year.  This should make any system capable
		// of displaying a 4-digit year display it in the locale's customary format.  On systems that don't
		// support DTS_SHORTDATECENTURYFORMAT, it should be ignored, resulting in DTS_SHORTDATEFORMAT taking
		// effect automatically (untested).
		opt.style_add |= WS_TABSTOP|DTS_SHORTDATECENTURYFORMAT;
		break;
	case GUI_CONTROL_BUTTON: // v1.0.45: Removed BS_MULTILINE from default because it is conditionally applied later below.
	case GUI_CONTROL_CHECKBOX: // v1.0.47.04: Removed BS_MULTILINE from default because it is conditionally applied later below.
	case GUI_CONTROL_HOTKEY:
	case GUI_CONTROL_SLIDER:
		opt.style_add |= WS_TABSTOP;
		break;
	case GUI_CONTROL_TAB:
		// Override the normal default, requiring a manual +Theme in the control's options.  This is done
		// because themed tabs have a gradient background that is currently not well supported by the method
		// used here (controls' backgrounds do not match the gradient):
		if (!opt.tab_control_uses_dialog)
			opt.use_theme = false;
		opt.style_add |= WS_TABSTOP|TCS_MULTILINE;
		break;
	case GUI_CONTROL_ACTIVEX:
		opt.style_add |= WS_CLIPSIBLINGS;
		break;
	case GUI_CONTROL_CUSTOM:
		opt.style_add |= WS_TABSTOP;
		break;
	case GUI_CONTROL_STATUSBAR:
		// Although the following appears unnecessary, at least on XP, there's a good chance it's required
		// on older OSes such as Win 95/NT.  On newer OSes, apparently the control shows a grip on
		// its own even though it doesn't even give itself the SBARS_SIZEGRIP style.
		if (mStyle & WS_SIZEBOX) // Parent window is resizable.
			opt.style_add |= SBARS_SIZEGRIP; // Provide a grip by default.
		// Below: Seems best to provide SBARS_TOOLTIPS by default, since we're not bound by backward compatibility
		// like the OS is.  In theory, tips should be displayed only when the script has actually set some tip text
		// (e.g. via SendMessage).  In practice, tips are never displayed, even under the precise conditions
		// described at MSDN's SB_SETTIPTEXT, perhaps under certain OS versions and themes.  See bottom of
		// BIF_StatusBar() for more comments.
		opt.style_add |= SBARS_TOOLTIPS;
		break;
	case GUI_CONTROL_MONTHCAL:
		// Testing indicates that although the MonthCal attached to a DateTime control can be navigated using
		// the keyboard on Win2k/XP, standalone MonthCal controls don't support it, except on Vista and later.
		opt.style_add |= WS_TABSTOP;
		break;
	// Nothing extra for these currently:
	//case GUI_CONTROL_RADIO: This one is handled separately above the switch().
	//case GUI_CONTROL_TEXT:
	//case GUI_CONTROL_PIC:
	//case GUI_CONTROL_GROUPBOX:
	//case GUI_CONTROL_PROGRESS:
		// v1.0.44.11: The following was commented out for GROUPBOX to avoid unwanted wrapping of last letter when
		// the font is bold on XP Classic theme (other font styles and desktop themes may also be cause this).
		// Avoiding this problem seems to outweigh the breaking of old scripts that use GroupBoxes with more than
		// one line of text (which are likely to be very rare).
		//opt.style_add |= BS_MULTILINE;
		break;
	}

	/////////////////////////////
	// Parse the list of options.
	/////////////////////////////
	if (!ControlParseOptions(aOptions, opt, control))
	{
		delete pcontrol;
		return FAIL;  // It already displayed the error.
	}

	// The following is needed by ControlSetListViewOptions/ControlSetTreeViewOptions, and possibly others
	// in the future.  It applies mCurrentColor as a default if no color was specified for the control.
	bool uses_font_and_text_color = control.UsesFontAndTextColor();
	if (uses_font_and_text_color) // Must check this to avoid corrupting union_hbitmap for PIC controls.
	{
		if (opt.color == CLR_INVALID) // No color was specified in options param.
			opt.color = mCurrentColor;
		if (control.UsesUnionColor()) // Must check this to avoid corrupting other union members.
			control.union_color = opt.color;
	}
	// This is not done because it would be inconsistent with the other controls (Progress
	// does not react to changes in the window color) and because applying a background
	// color requires removing the control's theme:
	//if (aControlType == GUI_CONTROL_PROGRESS && opt.color_bk == CLR_INVALID)
	//	opt.color_bk = mBackgroundColorWin; // Use the current background color.
	// If either color is CLR_DEFAULT, that would be either because the options contained
	// cDefault/BackgroundDefault, or because the Gui's current text/background color has
	// not been set.  In either case, set it to CLR_INVALID to indicate "no change needed".
	if (opt.color == CLR_DEFAULT)
		opt.color = CLR_INVALID;
	if (opt.color_bk == CLR_DEFAULT)
	{
		if (aControlType != GUI_CONTROL_TREEVIEW) // v1.1.08: Always set the back-color of a TreeView, otherwise it sends WM_CTLCOLOREDIT on Win2k/XP (or if it is not themed).
			opt.color_bk = CLR_INVALID;
	}

	// Change for v1.0.45 (buttons) and v1.0.47.04 (checkboxes and radios): Under some desktop themes and
	// unusual DPI settings, it has been reported that the last letter of the control's text gets truncated
	// and/or causes an unwanted wrap that prevents proper display of the text.  To solve this, default to
	// "wrapping enabled" only when necessary.  One case it's usually necessary is when there's an explicit
	// width present because then the text can automatically word-wrap to the next line if it contains any
	// spaces/tabs/dashes (this also improves backward compatibility).
	DWORD contains_bs_multiline_if_applicable =
		(opt.width != COORD_UNSPECIFIED || opt.height != COORD_UNSPECIFIED
			|| opt.row_count > 1.5 || StrChrAny(aText, _T("\n\r"))) // Both LF and CR can start new lines.
		? (BS_MULTILINE & ~opt.style_remove) // Add BS_MULTILINE unless it was explicitly removed.
		: 0; // Otherwise: Omit BS_MULTILINE (unless it was explicitly added [the "0" is verified correct]) because on some unusual DPI settings (i.e. DPIs other than 96 or 120), DrawText() sometimes yields a width that is slightly too narrow, which causes unwanted wrapping in single-line checkboxes/radios/buttons.

	DWORD style = opt.style_add & ~opt.style_remove;
	DWORD exstyle = opt.exstyle_add & ~opt.exstyle_remove;

	//////////////////////////////////////////
	// Force any mandatory styles into effect.
	//////////////////////////////////////////
	style |= WS_CHILD;  // All control types must have this, even if script attempted to remove it explicitly.
	switch (aControlType)
	{
	case GUI_CONTROL_GROUPBOX:
		// There doesn't seem to be any flexibility lost by forcing the buttons to be the right type,
		// and doing so improves maintainability and peace-of-mind:
		style = (style & ~BS_TYPEMASK) | BS_GROUPBOX;  // Force it to be the right type of button.
		break;
	case GUI_CONTROL_BUTTON:
		if (style & BS_DEFPUSHBUTTON) // i.e. its single bit is present. BS_TYPEMASK is not involved in this line because it's a purity check.
			style = (style & ~BS_TYPEMASK) | BS_DEFPUSHBUTTON; // Done to ensure the lowest four bits are pure.
		else
			style &= ~BS_TYPEMASK;  // Force it to be the right type of button --> BS_PUSHBUTTON == 0
		style |= contains_bs_multiline_if_applicable;
		break;
	case GUI_CONTROL_CHECKBOX:
		// Note: BS_AUTO3STATE and BS_AUTOCHECKBOX are mutually exclusive due to their overlap within
		// the bit field:
		if ((style & BS_AUTO3STATE) == BS_AUTO3STATE) // Fixed for v1.0.45.03 to check if all the BS_AUTO3STATE bits are present, not just "any" of them. BS_TYPEMASK is not involved here because this is a purity check, and TYPEMASK would defeat the whole purpose.
			style = (style & ~BS_TYPEMASK) | BS_AUTO3STATE; // Done to ensure the lowest four bits are pure.
		else
			style = (style & ~BS_TYPEMASK) | BS_AUTOCHECKBOX;  // Force it to be the right type of button.
		style |= contains_bs_multiline_if_applicable; // v1.0.47.04: Added to avoid unwanted wrapping on systems with unusual DPI settings (DPIs other than 96 and 120 sometimes seem to cause a roundoff problem with DrawText()).
		break;
	case GUI_CONTROL_RADIO:
		style = (style & ~BS_TYPEMASK) | BS_AUTORADIOBUTTON;  // Force it to be the right type of button.
		// This below must be handled here rather than in the set-defaults section because this
		// radio might be the first of its group due to the script having explicitly specified the word
		// Group in options (useful to make two adjacent radio groups).
		if (style & WS_GROUP && !(opt.style_remove & WS_TABSTOP))
			style |= WS_TABSTOP;
		// Otherwise it lacks a tabstop by default.
		style |= contains_bs_multiline_if_applicable; // v1.0.47.04: Added to avoid unwanted wrapping on systems with unusual DPI settings (DPIs other than 96 and 120 sometimes seem to cause a roundoff problem with DrawText()).
		break;
	case GUI_CONTROL_DROPDOWNLIST:
		style |= CBS_DROPDOWNLIST;  // This works because CBS_DROPDOWNLIST == CBS_SIMPLE|CBS_DROPDOWN
		break;
	case GUI_CONTROL_COMBOBOX:
		if (style & CBS_SIMPLE) // i.e. CBS_SIMPLE has been added to the original default, so assume it is SIMPLE.
			style = (style & ~0x0F) | CBS_SIMPLE; // Done to ensure the lowest four bits are pure.
		else
			style = (style & ~0x0F) | CBS_DROPDOWN; // Done to ensure the lowest four bits are pure.
		break;
	case GUI_CONTROL_LISTBOX:
		style |= LBS_NOTIFY;  // There doesn't seem to be any flexibility lost by forcing this style.
		break;
	case GUI_CONTROL_EDIT:
		// This is done for maintainability and peace-of-mind, though it might not strictly be required
		// to be done at this stage:
		if (opt.row_count > 1.5 || _tcschr(aText, '\n')) // Multiple rows or contents contain newline.
			style |= (ES_MULTILINE & ~opt.style_remove); // Add multiline unless it was explicitly removed.
		// This next check is relied upon by other things.  If this edit has the multiline style either
		// due to the above check or any other reason, provide other default styles if those styles
		// weren't explicitly removed in the options list:
		if (style & ES_MULTILINE) // If allowed, enable vertical scrollbar and capturing of ENTER keystrokes.
			// Safest to include ES_AUTOVSCROLL, though it appears to have no effect on XP.  See also notes below:
			#define EDIT_MULTILINE_DEFAULT (WS_VSCROLL|ES_WANTRETURN|ES_AUTOVSCROLL)
			style |= EDIT_MULTILINE_DEFAULT & ~opt.style_remove;
			// In addition, word-wrapping is implied unless explicitly disabled via -wrap in options.
			// This is because -wrap adds the ES_AUTOHSCROLL style.
		// else: Single-line edit.  ES_AUTOHSCROLL will be applied later below if all the other checks
		// fail to auto-detect this edit as a multi-line edit.
		// Notes: ES_MULTILINE is required for any CRLFs in the default value to display correctly.
		// If ES_MULTILINE is in effect: "If you do not specify ES_AUTOHSCROLL, the control automatically
		// wraps words to the beginning of the next line when necessary."
		// Also, ES_AUTOVSCROLL seems to have no additional effect, perhaps because this window type
		// is considered to be a dialog. MSDN: "When the multiline edit control is not in a dialog box
		// and the ES_AUTOVSCROLL style is specified, the edit control shows as many lines as possible
		// and scrolls vertically when the user presses the ENTER key. If you do not specify ES_AUTOVSCROLL,
		// the edit control shows as many lines as possible and beeps if the user presses the ENTER key when
		// no more lines can be displayed."
		break;
	case GUI_CONTROL_TAB:
		style |= WS_CLIPSIBLINGS; // MSDN: Both the parent window and the tab control must have the WS_CLIPSIBLINGS window style.
		// TCS_OWNERDRAWFIXED is required to implement custom Text color in the tabs.
		// For some reason, it's also required for TabWindowProc's WM_ERASEBKGND to be able to
		// override the background color of the control's interior, at least when an XP theme is in effect.
		if (control.background_brush
			|| mBackgroundBrushWin && control.background_color != CLR_DEFAULT
			|| IS_AN_ACTUAL_COLOR(opt.color))
		{
			style |= TCS_OWNERDRAWFIXED;
			// Even if use_theme is true, the theme won't be applied due to the above style.
			opt.use_theme = false; // Let CreateTabDialog() know not to enable the theme.
		}
		else
			style &= ~TCS_OWNERDRAWFIXED;
		if (opt.width != COORD_UNSPECIFIED) // If a width was specified, don't autosize horizontally.
			opt.tab_control_autosize &= ~TAB3_AUTOWIDTH;
		if (opt.height != COORD_UNSPECIFIED)
			opt.tab_control_autosize &= ~TAB3_AUTOHEIGHT;
		break;

	// Nothing extra for these currently:
	//case GUI_CONTROL_TEXT:  Ensuring SS_BITMAP and such are absent seems too over-protective.
	//case GUI_CONTROL_LINK:
	//case GUI_CONTROL_PIC:   SS_BITMAP/SS_ICON are applied after the control isn't created so that it doesn't try to auto-load a resource.
	//case GUI_CONTROL_LISTVIEW:
	//case GUI_CONTROL_TREEVIEW:
	//case GUI_CONTROL_DATETIME:
	//case GUI_CONTROL_MONTHCAL:
	//case GUI_CONTROL_HOTKEY:
	//case GUI_CONTROL_UPDOWN:
	//case GUI_CONTROL_SLIDER:
	//case GUI_CONTROL_PROGRESS:
	//case GUI_CONTROL_ACTIVEX:
	//case GUI_CONTROL_STATUSBAR:
	}

	// The below will yield NULL for GUI_CONTROL_STATUSBAR because control.tab_control_index==OutOfBounds for it.
	GuiControlType *owning_tab_control = FindTabControl(control.tab_control_index); // For use in various places.
	GuiControlType *prev_control = mControlCount ? mControl[mControlCount - 1] : NULL; // For code size reduction, performance, and maintainability.

	////////////////////////////////////////////////////////////////////////////////////////////
	// Automatically set the control's position in the client area if no position was specified.
	////////////////////////////////////////////////////////////////////////////////////////////
	if (opt.x == COORD_UNSPECIFIED && opt.y == COORD_UNSPECIFIED)
	{
		if (owning_tab_control && !GetControlCountOnTabPage(control.tab_control_index, control.tab_index)) // Relies on short-circuit boolean.
		{
			// Since this control belongs to a tab control and that tab control already exists,
			// Position it relative to the tab control's client area upper-left corner if this
			// is the first control on this particular tab/page:
			POINT pt = GetPositionOfTabDisplayArea(*owning_tab_control);
			// Since both coords were unspecified, position this control at the upper left corner of its page.
			opt.x = pt.x + mMarginX;
			opt.y = pt.y + mMarginY;
		}
		else
		{
			// GUI_CONTROL_STATUSBAR ignores these:
			// Since both coords were unspecified, proceed downward from the previous control, using a default margin.
			opt.x = mPrevX;
			opt.y = mPrevY + mPrevHeight + mMarginY;  // Don't use mMaxExtentDown in this is a new column.
		}
		if ((aControlType == GUI_CONTROL_TEXT || aControlType == GUI_CONTROL_LINK) && prev_control // This is a text control and there is a previous control before it.
			&& (prev_control->type == GUI_CONTROL_TEXT || prev_control->type == GUI_CONTROL_LINK)
			&& prev_control->tab_control_index == control.tab_control_index  // v1.0.44.03: Don't do the adjustment if
			&& prev_control->tab_index == control.tab_index)                 // it's on another page or in another tab control.
			// Since this text control is being auto-positioned immediately below another, provide extra
			// margin space so that any edit control, dropdownlist, or other "tall input" control later added
			// to its right in "vertical progression" mode will line up with it.
			// v1.0.44: For code simplicity, this doesn't handle any status bar that might be present in between,
			// since that seems too rare and the consequences too mild.
			opt.y += GUI_CTL_VERTICAL_DEADSPACE;
	}
	// Can't happen due to the logic in the options-parsing section:
	//else if (opt.x == COORD_UNSPECIFIED)
	//	opt.x = mPrevX;
	//else if (y == COORD_UNSPECIFIED)
	//	opt.y = mPrevY;


	/////////////////////////////////////////////////////////////////////////////////////
	// For certain types of controls, provide a standard row_count if none was specified.
	/////////////////////////////////////////////////////////////////////////////////////

	// Set default for all control types.  GUI_CONTROL_MONTHCAL must be set up here because
	// an explicitly specified row-count must also avoid calculating height from row count
	// in the standard way.
	bool calc_height_later = aControlType == GUI_CONTROL_MONTHCAL || aControlType == GUI_CONTROL_LISTVIEW
		 || aControlType == GUI_CONTROL_TREEVIEW;
	bool calc_control_height_from_row_count = !calc_height_later; // Set default.

	if (opt.height == COORD_UNSPECIFIED && opt.row_count < 1)
	{
		switch(aControlType)
		{
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX: // For these 2, row-count is defined as the number of items to display in the list.
			// Update: Unfortunately, heights taller than the desktop do not work: pre-v6 common controls
			// misbehave when the height is too tall to fit on the screen.  So the below comment is
			// obsolete and kept only for reference:
			// Since no height or row-count was given, make the control very tall so that OSes older than
			// XP will behavior similar to XP: they will let the desktop height determine how tall the
			// control can be. One exception to this is a true CBS_SIMPLE combo, which has appearance
			// and functionality similar to a ListBox.  In that case, a default row-count is provided
			// since that is more appropriate than having a really tall box hogging the window.
			// Because CBS_DROPDOWNLIST includes both CBS_SIMPLE and CBS_DROPDOWN, a true "simple"
			// (which is similar to a listbox) must omit CBS_DROPDOWN:
			opt.row_count = 3;  // Actual height will be calculated below using this.
			// Avoid doing various calculations later below if the XP+ will ignore the height anyway.
			// CBS_NOINTEGRALHEIGHT is checked in case that style was explicitly applied to the control
			// by the script. This is because on XP+ it will cause the combo/DropDownList to obey the
			// specified default height set above.  Also, for a pure CBS_SIMPLE combo, the OS always
			// obeys height:
			if ((!(style & CBS_SIMPLE) || (style & CBS_DROPDOWN)) // Not a pure CBS_SIMPLE.
				&& !(style & CBS_NOINTEGRALHEIGHT)) // ... and it won't obey the height.
				calc_control_height_from_row_count = false; // Don't bother calculating the height (i.e. override the default).
			break;
		case GUI_CONTROL_LISTBOX:
			opt.row_count = 3;  // Actual height will be calculated below using this.
			break;
		case GUI_CONTROL_LISTVIEW:
		case GUI_CONTROL_TREEVIEW:
		case GUI_CONTROL_CUSTOM:
			opt.row_count = 5;  // Actual height will be calculated below using this.
			break;
		case GUI_CONTROL_GROUPBOX:
			// Seems more appropriate to give GUI_CONTROL_GROUPBOX exactly two rows: the first for the
			// title of the group-box and the second for its content (since it might contain controls
			// placed horizontally end-to-end, and thus only need one row).
			opt.row_count = 2;
			break;
		case GUI_CONTROL_EDIT:
			// If there's no default text in the control from which to later calc the height, use a basic default.
			if (!*aText)
				opt.row_count = (style & ES_MULTILINE) ? 3.0F : 1.0F;
			break;
		case GUI_CONTROL_DATETIME:
		case GUI_CONTROL_HOTKEY:
			opt.row_count = 1;
			break;
		case GUI_CONTROL_UPDOWN: // A somewhat arbitrary default in case it will lack a buddy to "snap to".
			// Make vertical up-downs tall by default.  If their width has also been omitted, they
			// will be made narrow by default in a later section.
			if (style & UDS_HORZ)
				// Height vs. row_count is specified to ensure the same thickness for both vertical
				// and horizontal up-downs:
				opt.height = PROGRESS_DEFAULT_THICKNESS; // Seems okay for up-down to use Progress's thickness.
			else
				opt.row_count = 5.0F;
			break;
		case GUI_CONTROL_SLIDER:
			// Make vertical trackbars tall by default.  If their width has also been omitted, they
			// will be made narrow by default in a later section.
			if (style & TBS_VERT)
				opt.row_count = 5.0F;
			else
				opt.height = ControlGetDefaultSliderThickness(style, opt.thickness);
			break;
		case GUI_CONTROL_PROGRESS:
			// Make vertical progress bars tall by default.  If their width has also been omitted, they
			// will be made narrow by default in a later section.
			if (style & PBS_VERTICAL)
				opt.row_count = 5.0F;
			else
				// Height vs. row_count is specified to ensure the same thickness for both vertical
				// and horizontal progress bars:
				opt.height = PROGRESS_DEFAULT_THICKNESS;
			break;
		case GUI_CONTROL_TAB:
			opt.row_count = 10;
			break;
		// Types not included
		// ------------------
		//case GUI_CONTROL_TEXT:      Rows are based on control's contents.
		//case GUI_CONTROL_LINK:      Rows are based on control's contents.
		//case GUI_CONTROL_PIC:       N/A
		//case GUI_CONTROL_BUTTON:    Rows are based on control's contents.
		//case GUI_CONTROL_CHECKBOX:  Same
		//case GUI_CONTROL_RADIO:     Same
		//case GUI_CONTROL_MONTHCAL:  Leave row-count unspecified so that an explicit r1 can be distinguished from "unspecified".
		//case GUI_CONTROL_ACTIVEX:   N/A
		//case GUI_CONTROL_STATUSBAR: For now, row-count is ignored/unused.
		}
	}
	else // Either a row_count or a height was explicitly specified.
		// If OS is XP+, must apply the CBS_NOINTEGRALHEIGHT style for these reasons:
		// 1) The app now has a manifest, which tells OS to use common controls v6.
		// 2) Common controls v6 will not obey the the user's specified height for the control's
		//    list portion unless the CBS_NOINTEGRALHEIGHT style is present.
		if (aControlType == GUI_CONTROL_DROPDOWNLIST || aControlType == GUI_CONTROL_COMBOBOX)
			style |= CBS_NOINTEGRALHEIGHT; // Forcibly applied, even if removed in options.

	////////////////////////////////////////////////////////////////////////////////////////////
	// In case the control being added requires an HDC to calculate its size, provide the means.
	////////////////////////////////////////////////////////////////////////////////////////////
	HDC hdc = NULL;
	HFONT hfont_old = NULL;
	TEXTMETRIC tm; // Used in more than one place.
	// To improve maintainability, always use this macro to deal with the above.
	// HDC will be released much further below when it is no longer needed.
	// Remember to release DC if ever need to return FAIL in the middle of auto-sizing/positioning.
	#define GUI_SET_HDC \
		if (!hdc)\
		{\
			hdc = GetDC(mHwnd);\
			hfont_old = (HFONT)SelectObject(hdc, sFont[mCurrentFontIndex].hfont);\
		}

	//////////////////////////////////////////////////////////////////////////////////////
	// If a row-count was specified or made available by the above defaults, calculate the
	// control's actual height (to be used when creating the window).  Note: If both
	// row_count and height were explicitly specified, row_count takes precedence.
	//////////////////////////////////////////////////////////////////////////////////////
	if (opt.row_count > 0)
	{
		// For GroupBoxes, add 1 to any row_count greater than 1 so that the title itself is
		// already included automatically.  In other words, the R-value specified by the user
		// should be the number of rows available INSIDE the box.
		// For DropDownLists and ComboBoxes, 1 is added because row_count is defined as the
		// number of rows shown in the drop-down portion of the control, so we need one extra
		// (used in later calculations) for the always visible portion of the control.
		switch (aControlType)
		{
		case GUI_CONTROL_GROUPBOX:
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX: // LISTVIEW has custom handling later on, so isn't listed here.
			++opt.row_count;
			break;
		}
		if (calc_control_height_from_row_count)
		{
			GUI_SET_HDC
			GetTextMetrics(hdc, &tm);
			// Calc the height by adding up the font height for each row, and including the space between lines
			// (tmExternalLeading) if there is more than one line.  0.5 is used in two places to prevent
			// negatives in one, and round the overall result in the other.
			opt.height = (int)((tm.tmHeight * opt.row_count) + (tm.tmExternalLeading * ((int)(opt.row_count + 0.5) - 1)) + 0.5);
			switch (aControlType)
			{
			case GUI_CONTROL_DROPDOWNLIST:
			case GUI_CONTROL_COMBOBOX:
			case GUI_CONTROL_LISTBOX:
			case GUI_CONTROL_EDIT:
			case GUI_CONTROL_DATETIME:
			case GUI_CONTROL_HOTKEY:
			case GUI_CONTROL_CUSTOM:
				opt.height += GUI_CTL_VERTICAL_DEADSPACE;
				if (style & WS_HSCROLL)
					opt.height += GetSystemMetrics(SM_CYHSCROLL);
				break;
			case GUI_CONTROL_BUTTON:
				// Provide a extra space for top/bottom margin together, proportional to the current font
				// size so that it looks better with very large or small fonts.  Logical font height - 1
				// looks about right on all font sizes, and with default font size on 100% DPI it gives a
				// height of 23, which is what Microsoft recommends: https://msdn.microsoft.com/library/dn742402.aspx#sizing
				opt.height += -sFont[mCurrentFontIndex].lfHeight - 1;
				break;
			case GUI_CONTROL_GROUPBOX: // Since groups usually contain other controls, the below sizing seems best.
				// Use row_count-2 because of the +1 added above for GUI_CONTROL_GROUPBOX.
				// The current font's height is added in to provide an upper/lower margin in the box
				// proportional to the current font size, which makes it look better in most cases:
				opt.height += mMarginY * (2 + ((int)(opt.row_count + 0.5) - 2));
				break;
			case GUI_CONTROL_TAB:
				opt.height += mMarginY * (2 + ((int)(opt.row_count + 0.5) - 1)); // -1 vs. the -2 used above.
				break;
			// Types not included
			// ------------------
			//case GUI_CONTROL_TEXT:     Uses basic height calculated above the switch().
			//case GUI_CONTROL_LINK:     Uses basic height calculated above the switch().
			//case GUI_CONTROL_PIC:      Uses basic height calculated above the switch() (seems OK even for pic).
			//case GUI_CONTROL_CHECKBOX: Uses basic height calculated above the switch().
			//case GUI_CONTROL_RADIO:    Same.
			//case GUI_CONTROL_UPDOWN:   Same.
			//case GUI_CONTROL_SLIDER:   Same.
			//case GUI_CONTROL_PROGRESS: Same.
			//case GUI_CONTROL_MONTHCAL: Not included at all in this section because it treats "rows" differently.
			//case GUI_CONTROL_ACTIVEX:  N/A
			//case GUI_CONTROL_STATUSBAR: N/A
			} // switch
		}
		else // calc_control_height_from_row_count == false
			// Assign a default just to allow the control to be created successfully. 13 is the default
			// height of a text/radio control for the typical 8 point font size, but the exact value
			// shouldn't matter (within reason) since calc_control_height_from_row_count is telling us this type of
			// control will not obey the height anyway.  Update: It seems better to use a small constant
			// value to help catch bugs while still allowing the control to be created:
			if (!calc_height_later)
				opt.height = DPIScale(30);
			//else MONTHCAL and others must keep their "unspecified height" value for later detection.
	}

	bool control_width_was_set_by_contents = false;

	if (opt.height == COORD_UNSPECIFIED || opt.width == COORD_UNSPECIFIED)
	{
		// Set defaults:
		int extra_width = 0, extra_height = 0;
		UINT draw_format = DT_CALCRECT;

		switch (aControlType)
		{
		case GUI_CONTROL_EDIT:
			if (!*aText) // Only auto-calculate edit's dimensions if there is text to do it with.
				break;
			// Since edit controls leave approximate 1 avg-char-width margin on the right side,
			// and probably exactly 4 pixels on the left counting its border and the internal
			// margin), adjust accordingly so that DrawText() will calculate the correct
			// control height based on word-wrapping.  Note: Can't use EM_GETRECT because
			// control doesn't exist yet (though that might be an alternative approach for
			// the future):
			GUI_SET_HDC
			GetTextMetrics(hdc, &tm);
			extra_width += 4 + tm.tmAveCharWidth;
			// Determine whether there will be a vertical scrollbar present.  If ES_MULTILINE hasn't
			// already been applied or auto-detected above, it's possible that a scrollbar will be
			// added later due to the text auto-wrapping.  In that case, the calculated height may
			// be incorrect due to the additional wrapping caused by the width taken up by the
			// scrollbar.  Since this combination of circumstances is rare, and since there are easy
			// workarounds, it's just documented here as a limitation:
			if (style & WS_VSCROLL)
				extra_width += GetSystemMetrics(SM_CXVSCROLL);
			// DT_EDITCONTROL: "the average character width is calculated in the same manner as for an edit control"
			// Although it's hard to say exactly what the quote above means, DT_EDITCONTROL definitely
			// affects the width of tab stops (making this consistent with Edit controls) and therefore
			// must be included.  It also excludes the last line if it is empty, which is undesirable,
			// so we need to compensate for that:
			if (*aText && aText[_tcslen(aText)-1] == '\n')
				extra_height += tm.tmHeight + tm.tmExternalLeading;
			// Also include DT_EXPANDTABS under the assumption that if there are tabs present, the user
			// intended for them to be there because a multiline edit would expand them (rather than trying
			// to worry about whether this control *might* become auto-multiline after this point.
			draw_format |= DT_EXPANDTABS|DT_EDITCONTROL|DT_NOPREFIX; // v1.0.44.10: Added DT_NOPREFIX because otherwise, if the text contains & or &&, the control won't be sized properly.
			// and now fall through and have the dimensions calculated based on what's in the control.
			// ABOVE FALLS THROUGH TO BELOW
		case GUI_CONTROL_TEXT:
		case GUI_CONTROL_BUTTON:
		case GUI_CONTROL_CHECKBOX:
		case GUI_CONTROL_RADIO:
		case GUI_CONTROL_LINK:
		{
			GUI_SET_HDC
			if (aControlType == GUI_CONTROL_TEXT)
			{
				draw_format |= DT_EXPANDTABS; // Buttons can't expand tabs, so don't add this for them.
				if (style & SS_NOPREFIX) // v1.0.44.10: This is necessary to auto-width the control properly if its contents include any ampersands.
					draw_format |= DT_NOPREFIX;
			}
			else if (aControlType == GUI_CONTROL_CHECKBOX || aControlType == GUI_CONTROL_RADIO)
			{
				// Both Checkbox and Radio seem to have the same spacing characteristics:
				// Expand to allow room for button itself, its border, and the space between
				// the button and the first character of its label (this space seems to
				// be the same as tmAveCharWidth).  +2 seems to be needed to make it work
				// for the various sizes of Courier New vs. Verdana that I tested.  The
				// alternative, (2 * GetSystemMetrics(SM_CXEDGE)), seems to add a little
				// too much width (namely 4 vs. 2).
				GetTextMetrics(hdc, &tm);
				extra_width += GetSystemMetrics(SM_CXMENUCHECK) + tm.tmAveCharWidth + 2; // v1.0.40.03: Reverted to +2 vs. +3 (it had been changed to +3 in v1.0.40.01).
			}
			if ((style & WS_BORDER) && (aControlType == GUI_CONTROL_TEXT || aControlType == GUI_CONTROL_LINK))
			{
				// This seems to be necessary only for Text and Link controls:
				extra_width = 2 * GetSystemMetrics(SM_CXBORDER);
				extra_height = 2 * GetSystemMetrics(SM_CYBORDER);
			}
			if (   opt.width != COORD_UNSPECIFIED // Since a width was given, auto-expand the height via word-wrapping,
				&& ( !(aControlType == GUI_CONTROL_BUTTON || aControlType == GUI_CONTROL_CHECKBOX || aControlType == GUI_CONTROL_RADIO)
					|| (style & BS_MULTILINE) )   ) // except when -Wrap is used on a Button, Checkbox or Radio.
				draw_format |= DT_WORDBREAK;

			RECT draw_rect;
			draw_rect.left = 0;
			draw_rect.top = 0;
			draw_rect.right = (opt.width == COORD_UNSPECIFIED) ? 0 : opt.width - extra_width; // extra_width
			draw_rect.bottom = (opt.height == COORD_UNSPECIFIED) ? 0 : opt.height;

			int draw_height;
			TCHAR last_char = 0;

			// Since a Link control's text contains markup which isn't actually rendered, we need
			// to strip it out before calling DrawText or it will put out the size calculation:
			if (aControlType == GUI_CONTROL_LINK)
			{
				// Unless the control has the LWS_NOPREFIX style, the first ampersand causes the character
				// following it to be rendered with an underline, and the ampersand itself is not rendered.
				// Since DrawText has this same behaviour by default, we can simply set DT_NOPREFIX when
				// appropriate and it should handle the ampersands correctly.
				if (style & LWS_NOPREFIX)
					draw_format |= DT_NOPREFIX;

				LPTSTR aTextCopy = new TCHAR[_tcslen(aText) + 1];

				LPCTSTR src;
				LPTSTR dst;
				for (src = aText, dst = aTextCopy; *src; ++src)
				{
					if (*src == '<' && ctoupper(src[1]) == 'A')
					{
						LPCTSTR linkText = NULL;
						LPCTSTR cp = omit_leading_whitespace(src + 2);

						// It is important to note that while the SysLink control's syntax is similar to HTML,
						// it is not HTML.  Some valid HTML won't work, and some invalid HTML will.  Testing
						// indicates that whitespace is not required between "<A" and the attribute, so there
						// is no check to see if the above actually omitted whitespace.
						for (;;)
						{
							if (*cp == '>')
							{
								linkText = cp + 1;
								break;
							}

							// Testing indicates that attribute names can only be alphanumeric chars.  This
							// method of testing for the attribute name does not treat <a=""> as an error,
							// which is perfect since the control tolerates it (the attribute is ignored).
							while (cisalnum(*cp)) 
								cp++;
							
							// What follows next must be the attribute value.  Spaces are not tolerated.
							if (*cp != '=' || '"' != cp[1])
								break; // Not valid.
								
							// Testing indicates that the attribute value can contain virtually anything,
							// including </a> but not quotation marks.
							cp = _tcschr(cp + 2, '"');
							if (!cp)
								break; // Not valid.
							
							// Allow whitespace after the attribute (before '>' or the next attribute).
							cp = omit_leading_whitespace(cp + 1);
						} // for (attribute-parsing loop)

						if (linkText)
						{
							// Testing indicates that spaces in the end tag are not tolerated:
							if (auto endTag = tcscasestr(linkText, _T("</a>")))
							{
								// Copy the link text and skip over the end tag.
								for (cp = linkText; cp < endTag; *dst++ = *cp++);
								src = endTag + 3; // The outer loop's increment makes it + 4.
								continue;
							}
							// Otherwise, there's no end tag so the begin tag is treated as text.
						}
						// Otherwise, it wasn't a valid tag, so it is treated as text.
					} // if (found a tag)
					*dst++ = *src;
				} // for (searching for tag, copying text)
				
				*dst = '\0';

				if (dst > aTextCopy)
					last_char = dst[-1];

				// If no text, "H" is used in case the function requires a non-empty string to give consistent results:
				draw_height = DrawText(hdc, *aTextCopy ? aTextCopy : _T("H"), -1, &draw_rect, draw_format);
				
				delete[] aTextCopy;
			}
			else
			{
				// If no text, "H" is used in case the function requires a non-empty string to give consistent results:
				draw_height = DrawText(hdc, *aText ? aText : _T("H"), -1, &draw_rect, draw_format);
				if (*aText)
					last_char = aText[_tcslen(aText)-1];
			}
			
			int draw_width = draw_rect.right - draw_rect.left;

			switch (aControlType)
			{
			case GUI_CONTROL_BUTTON:
				if (contains_bs_multiline_if_applicable) // BS_MULTILINE style prevents the control from utilizing the extra space.
					break;
			case GUI_CONTROL_TEXT:
			case GUI_CONTROL_EDIT:
				if (!last_char)
					break;
				// draw_rect doesn't include the overhang of the last character, so adjust for that
				// to avoid clipping of italic fonts and other fonts where characters overlap.
				// This is currently only done for Text, Edit and Button controls because the other
				// control types won't utilize the extra space (they are apparently clipped according
				// to what the OS erroneously calculates the width of the text to be).
				ABC abc;
				if (GetCharABCWidths(hdc, last_char, last_char, &abc))
					if (abc.abcC < 0)
						draw_width -= abc.abcC;
			}

			// Even if either height or width was already explicitly specified above, it seems best to
			// override it if DrawText() says it's not big enough.  REASONING: It seems too rare that
			// someone would want to use an explicit height/width to selectively hide part of a control's
			// contents, presumably for revelation later.  If that is truly desired, ControlMove or
			// similar can be used to resize the control afterward.  In addition, by specifying BOTH
			// width and height/rows, none of these calculations happens anyway, so that's another way
			// this override can be overridden.  UPDATE for v1.0.44.10: The override is now not done for Edit
			// controls because unlike the other control types enumerated above, it is much more common to
			// have an Edit not be tall enough to show all of it's initial text.  This fixes the following:
			//   Gui, Add, Edit, r2 ,Line1`nLine2`nLine3`nLine4
			// Since there's no explicit width above, the r2 option (or even an H option) would otherwise
			// be overridden in favor of making the edit tall enough to hold all 4 lines.
			// Another reason for not changing the other control types to be like Edit is that backward
			// compatibility probably outweighs any value added by changing them (and the added value is dubious
			// when the comments above are carefully considered).
			if (opt.height == COORD_UNSPECIFIED || (draw_height > opt.height && aControlType != GUI_CONTROL_EDIT))
			{
				opt.height = draw_height + extra_height;
				if (aControlType == GUI_CONTROL_EDIT)
				{
					opt.height += GUI_CTL_VERTICAL_DEADSPACE;
					if (style & WS_HSCROLL)
						opt.height += GetSystemMetrics(SM_CYHSCROLL);
				}
				else if (aControlType == GUI_CONTROL_BUTTON)
					opt.height += -sFont[mCurrentFontIndex].lfHeight - 1; // See calc_control_height_from_row_count section for comments.
			}
			if (opt.width == COORD_UNSPECIFIED || draw_width > opt.width)
			{
				// v1.0.44.08: Fixed the following line by moving it into this IF block.  This prevents
				// an up-down from widening its edit control when that edit control had an explicit width.
				control_width_was_set_by_contents = true; // Indicate that this control was auto-width'd.
				opt.width = draw_width + extra_width;  // See comments above for why specified width is overridden.
				if (aControlType == GUI_CONTROL_BUTTON)
					// Allow room for border and an internal margin proportional to the font height.
					// Button's border is 3D by default, so SM_CXEDGE vs. SM_CXBORDER is used?
					opt.width += 2 * GetSystemMetrics(SM_CXEDGE) + -sFont[mCurrentFontIndex].lfHeight; // Compared to the old way of using point_size, this is slightly wider (which looks better) and automatically scales by DPI.
			}
			break;
		} // case for text/button/checkbox/radio/link

		// Types not included
		// ------------------
		//case GUI_CONTROL_PIC:           If needed, it is given some default dimensions at the time of creation.
		//case GUI_CONTROL_GROUPBOX:      Seems too rare than anyone would want its width determined by its text.
		//case GUI_CONTROL_EDIT:          It is included, but only if it has default text inside it.
		//case GUI_CONTROL_TAB:           Seems too rare than anyone would want its width determined by tab-count.

		//case GUI_CONTROL_LISTVIEW:      Has custom handling later below.
		//case GUI_CONTROL_TREEVIEW:      Same.
		//case GUI_CONTROL_MONTHCAL:      Same.
		//case GUI_CONTROL_ACTIVEX:       N/A?
		//case GUI_CONTROL_STATUSBAR:     Ignores width/height, so no need to handle here.

		//case GUI_CONTROL_DROPDOWNLIST:  These last ones are given (later below) a standard width based on font size.
		//case GUI_CONTROL_COMBOBOX:      In addition, their height has already been determined further above.
		//case GUI_CONTROL_LISTBOX:
		//case GUI_CONTROL_DATETIME:
		//case GUI_CONTROL_HOTKEY:
		//case GUI_CONTROL_UPDOWN:
		//case GUI_CONTROL_SLIDER:
		//case GUI_CONTROL_PROGRESS:
		} // switch()
	}

	//////////////////////////////////////////////////////////////////////////////////////////
	// If the width was not specified and the above did not already determine it (which should
	// only be possible for the cases contained in the switch-stmt below), provide a default.
	//////////////////////////////////////////////////////////////////////////////////////////
	if (opt.width == COORD_UNSPECIFIED)
	{
		int gui_standard_width = GUI_STANDARD_WIDTH; // Resolve macro only once for performance/code size.
		switch(aControlType)
		{
		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
		case GUI_CONTROL_LISTBOX:
		case GUI_CONTROL_HOTKEY:
		case GUI_CONTROL_EDIT:
			opt.width = gui_standard_width;
			break;
		case GUI_CONTROL_LISTVIEW:
		case GUI_CONTROL_TREEVIEW:
		case GUI_CONTROL_DATETIME: // Seems better to have wider default to fit LongDate and because drop-down calendar is fairly wide (though the latter is a weak reason).
		case GUI_CONTROL_CUSTOM:
			opt.width = gui_standard_width * 2;
			break;
		case GUI_CONTROL_UPDOWN: // Iffy, but needs some kind of default?
			opt.width = (style & UDS_HORZ) ? gui_standard_width : PROGRESS_DEFAULT_THICKNESS; // Progress's default seems ok for up-down too.
			break;
		case GUI_CONTROL_SLIDER:
			// Make vertical trackbars narrow by default.  For vertical trackbars: there doesn't seem
			// to be much point in defaulting the width to something proportional to font size because
			// the thumb only seems to have two sizes and doesn't auto-grow any larger than that.
			opt.width = (style & TBS_VERT) ? ControlGetDefaultSliderThickness(style, opt.thickness) : gui_standard_width;
			break;
		case GUI_CONTROL_PROGRESS:
			opt.width = (style & PBS_VERTICAL) ? PROGRESS_DEFAULT_THICKNESS : gui_standard_width;
			break;
		case GUI_CONTROL_GROUPBOX:
			// Since groups and tabs contain other controls, allow room inside them for a margin based
			// on current font size.
			opt.width = gui_standard_width + (2 * mMarginX);
			break;
		case GUI_CONTROL_TAB:
			// Tabs tend to be wide so that that tabs can all fit on the top row, and because they
			// are usually used to fill up the entire window.  Therefore, default them to the ability
			// to hold two columns of standard-width controls:
			opt.width = (2 * gui_standard_width) + (3 * mMarginX);  // 3 vs. 2 to allow space in between columns.
			break;
		// Types not included
		// ------------------
		//case GUI_CONTROL_TEXT:      Exact width should already have been calculated based on contents.
		//case GUI_CONTROL_LINK:      Exact width should already have been calculated based on contents.
		//case GUI_CONTROL_PIC:       Calculated based on actual pic size if no explicit width was given.
		//case GUI_CONTROL_BUTTON:    Exact width should already have been calculated based on contents.
		//case GUI_CONTROL_CHECKBOX:  Same.
		//case GUI_CONTROL_RADIO:     Same.
		//case GUI_CONTROL_MONTHCAL:  Exact width will be calculated after the control is created (size to fit month).
		//case GUI_CONTROL_ACTIVEX:   Defaults to zero width/height, which could be useful in some cases.
		//case GUI_CONTROL_STATUSBAR: Ignores width, so no need to handle here.
		}
	}

	/////////////////////////////////////////////////////////////////////////////////////////
	// For edit controls: If the above didn't already determine how many rows it should have,
	// auto-detect that by comparing the current font size with the specified height. At this
	// stage, the above has already ensured that an Edit has at least a height or a row_count.
	/////////////////////////////////////////////////////////////////////////////////////////
	if (aControlType == GUI_CONTROL_EDIT && !(style & ES_MULTILINE))
	{
		if (opt.row_count < 1) // Determine the row-count to auto-detect multi-line vs. single-line.
		{
			GUI_SET_HDC
			GetTextMetrics(hdc, &tm);
			int height_beyond_first_row = opt.height - GUI_CTL_VERTICAL_DEADSPACE - tm.tmHeight;
			if (style & WS_HSCROLL)
				height_beyond_first_row -= GetSystemMetrics(SM_CYHSCROLL);
			if (height_beyond_first_row > 0)
			{
				opt.row_count = 1 + ((float)height_beyond_first_row / (tm.tmHeight + tm.tmExternalLeading));
				// This section is a near exact match for one higher above.  Search for comment
				// "Add multiline unless it was explicitly removed" for a full explanation and keep
				// the below in sync with that section above:
				if (opt.row_count > 1.5)
				{
					style |= (ES_MULTILINE & ~opt.style_remove); // Add multiline unless it was explicitly removed.
					// Do the below only if the above actually added multiline:
					if (style & ES_MULTILINE) // If allowed, enable vertical scrollbar and capturing of ENTER keystrokes.
						style |= EDIT_MULTILINE_DEFAULT & ~opt.style_remove;
					// else: Single-line edit.  ES_AUTOHSCROLL will be applied later below if all the other checks
					// fail to auto-detect this edit as a multi-line edit.
				}
			}
			else // there appears to be only one row.
				opt.row_count = 1;
				// And ES_AUTOHSCROLL will be applied later below if all the other checks
				// fail to auto-detect this edit as a multi-line edit.
		}
	}

	// If either height or width is still undetermined, leave it set to COORD_UNSPECIFIED since that
	// is a large negative number and should thus help catch bugs.  In other words, the above
	// heuristics should be designed to handle all cases and always resolve height/width to something,
	// with the possible exception of things that auto-size based on external content such as
	// GUI_CONTROL_PIC.

	//////////////////////
	//
	// CREATE THE CONTROL.
	//
	//////////////////////
	bool do_strip_theme = !opt.use_theme;   // Set defaults.
	bool retrieve_dimensions = false;       //
	int item_height, min_list_height;
	RECT rect;
	HMENU control_id = (HMENU)(size_t)GUI_INDEX_TO_ID(mControlCount); // Cast to size_t avoids compiler warning.
	HWND parent_hwnd = mHwnd;

	bool font_was_set = false;          // "
	bool is_parent_visible = IsWindowVisible(mHwnd) && !IsIconic(mHwnd);
	#define GUI_SETFONT \
	{\
		SendMessage(control.hwnd, WM_SETFONT, (WPARAM)sFont[mCurrentFontIndex].hfont, is_parent_visible);\
		font_was_set = true;\
	}

	// If a control is being added to a tab, even if the parent window is hidden (since it might
	// have been hidden by Gui, Cancel), make sure the control isn't visible unless it's on a
	// visible tab.
	// The below alters style vs. style_remove, since later below style_remove is checked to
	// find out if the control was explicitly hidden vs. hidden by the automatic action here:
	bool on_visible_page_of_tab_control = false;
	if (control.tab_control_index < MAX_TAB_CONTROLS) // This control belongs to a tab control (must check this even though FindTabControl() does too).
	{
		if (owning_tab_control) // Its tab control exists...
		{
			HWND tab_dialog = (HWND)GetProp(owning_tab_control->hwnd, _T("ahk_dlg"));
			if (tab_dialog)
			{
				// Controls in a Tab3 control are created as children of an invisible dialog window
				// rather than as siblings of the tab control.  This solves some visual glitches.
				parent_hwnd = tab_dialog;
				POINT pt = { opt.x, opt.y };
				MapWindowPoints(mHwnd, parent_hwnd, &pt, 1); // Translate from GUI client area to tab client area.
				opt.x = pt.x, opt.y = pt.y;
			}
			if (!(GetWindowLong(owning_tab_control->hwnd, GWL_STYLE) & WS_VISIBLE) // Don't use IsWindowVisible().
				|| TabCtrl_GetCurSel(owning_tab_control->hwnd) != control.tab_index)
				// ... but it's not set to the page/tab that contains this control, or the entire tab control is hidden.
				style &= ~WS_VISIBLE;
			else // Make the following true as long as the parent is also visible.
				on_visible_page_of_tab_control = is_parent_visible;  // For use later below.
		}
		else // Its tab control does not exist, so this control is kept hidden until such time that it does.
			style &= ~WS_VISIBLE;
	}
	// else do nothing.

	switch(aControlType)
	{
	case GUI_CONTROL_TEXT:
		// Seems best to omit SS_NOPREFIX by default so that ampersand can be used to create shortcut keys.
		control.hwnd = CreateWindowEx(exstyle, _T("static"), aText, style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL);
		break;

	case GUI_CONTROL_LINK:
		// Seems best to omit LWS_NOPREFIX by default so that ampersand can be used to create shortcut keys.
		control.hwnd = CreateWindowEx(exstyle, _T("SysLink"), aText, style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL);
		break;

	case GUI_CONTROL_PIC:
		if (opt.width == COORD_UNSPECIFIED)
			opt.width = 0;  // Use zero to tell LoadPicture() to keep original width.
		if (opt.height == COORD_UNSPECIFIED)
			opt.height = 0;  // Use zero to tell LoadPicture() to keep original height.
		// Must set its caption to aText so that documented ability to refer to a picture by its original
		// filename is possible:
		if (control.hwnd = CreateWindowEx(exstyle, _T("static"), aText, style
			, opt.x, opt.y, opt.width, opt.height  // OK if zero, control creation should still succeed.
			, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (!ControlLoadPicture(control, aText, opt.width, opt.height, opt.icon_number))
			{
				// GetLastError() isn't used because a variety of methods are used, and it is likely that
				// some failure cases haven't SetLastError().
				DestroyWindow(control.hwnd);
				control.hwnd = NULL;
				break;
			}
			// UPDATE ABOUT THE BELOW: Rajat says he can't get the Smart GUI working without
			// the controls retaining their original numbering/z-order.  This has to do with the fact
			// that TEXT controls and PIC controls are both static.  If only PIC controls were reordered,
			// that's not enough to make things work for him.  If only TEXT controls were ALSO reordered
			// to be at the top of the list (so that all statics retain their original ordering with
			// respect to each other) I have a 90% expectation (based on testing) that prefix/shortcut
			// keys inside static text controls would jump to the wrong control because their associated
			// control seems to be based solely on z-order.  ANOTHER REASON NOT to ever change the z-order
			// of controls automatically: Every time a picture is added, it would increment the z-order
			// number of all other control types by 1.  This is bad especially for text controls because
			// they are static just like pictures, and thus their Class+NN "unique ID" would change (as
			// seen by the ControlXXX commands) if a picture were ever added after a window was shown.
			// Older note: calling SetWindowPos() with HWND_TOPMOST doesn't seem to provide any useful
			// effect that I could discern (by contrast, HWND_TOP does actually move a control to the
			// top of the z-order).
			// The below is OBSOLETE and its code further below is commented out:
			// Facts about how overlapping controls are drawn vs. which one receives mouse clicks:
			// 1) The first control created is at the top of the Z-order (i.e. the lowest z-order number
			//    and the first in tab navigation), the second is next, and so on.
			// 2) Controls get drawn in ascending Z-order (i.e. the first control is drawn first
			//    and any later controls that overlap are drawn on top of it, except for controls that
			//    have WS_CLIPSIBLINGS).
			// 3) When a user clicks a point that contains two overlapping controls and each control is
			//    capable of capturing clicks, the one closer to the top captures the click even though it
			//    was drawn beneath (overlapped by) the other control.
			// Because of this behavior, the following policy seems best:
			// 1) Move all static images to the top of the Z-order so that other controls are always
			//    drawn on top of them.  This is done because it seems to be the behavior that would
			//    be desired at least 90% of the time.
			// 2) Do not do the same for static text and GroupBoxes because it seems too rare that
			//    overlapping would be done in such cases, and even if it is done it seems more
			//    flexible to allow the order in which the controls were created to determine how they
			//    overlap and which one get the clicks.
			//
			// Rather than push static pictures to the top in the reverse order they were created -- 
			// which might be a little more intuitive since the ones created first would then always
			// be "behind" ones created later -- for simplicity, we do it here at the time the control
			// is created.  This avoids complications such as a picture being added after the
			// window is shown for the first time and then not getting sent to the top where it
			// should be.  Update: The control is now kept in its original order for two reasons:
			// 1) The reason mentioned above (that later images should overlap earlier images).
			// 2) Allows Rajat's SmartGUI Creator to support picture controls (due to its grid background pic).
			// First find the last picture control in the array.  The last one should already be beneath
			// all the others in the z-order due to having been originally added using the below method.
			// Might eventually need to switch to EnumWindows() if it's possible for Z-order to be
			// altered, or controls to be inserted between others, by any commands in the future.
			//GuiIndexType index_of_last_picture = UINT_MAX;
			//for (u = 0; u < mControlCount; ++u)
			//{
			//	if (mControl[u].type == GUI_CONTROL_PIC)
			//		index_of_last_picture = u;
			//}
			//if (index_of_last_picture == UINT_MAX) // There are no other pictures, so put this one on top.
			//	SetWindowPos(control.hwnd, HWND_TOP, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_NOMOVE|SWP_NOSIZE);
			//else // Put this picture after the last picture in the z-order.
			//	SetWindowPos(control.hwnd, mControl[index_of_last_picture].hwnd, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_NOMOVE|SWP_NOSIZE);
			//// Adjust to control's actual size in case it changed for any reason (failure to load picture, etc.)
			retrieve_dimensions = true;
		}
		break;

	case GUI_CONTROL_GROUPBOX:
		// In this case, BS_MULTILINE will obey literal newlines in the text, but it does not automatically
		// wrap the text, at least on XP.  Since it's strange-looking to have multiple lines, newlines
		// should be rarely present anyway.  Also, BS_NOTIFY seems to have no effect on GroupBoxes (it
		// never sends any BN_CLICKED/BN_DBLCLK messages).  This has been verified twice.
		control.hwnd = CreateWindowEx(exstyle, _T("Button"), aText, style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL);
		break;

	case GUI_CONTROL_BUTTON:
		// For all "Button" type controls, BS_MULTILINE is included by default so that any literal
		// newlines in the button's name will start a new line of text as the user intended.
		// In addition, this causes automatic wrapping to occur if the user specified a width
		// too small to fit the entire line.
		if (control.hwnd = CreateWindowEx(exstyle, _T("Button"), aText, style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (style & BS_DEFPUSHBUTTON)
			{
				// First remove the style from the old default button, if there is one:
				if (mDefaultButtonIndex < mControlCount)
				{
					// MSDN says this is necessary in some cases:
					// Since the window might be visible at this point, send BM_SETSTYLE rather than
					// SetWindowLong() so that the button will get redrawn.  Update: The redraw doesn't
					// actually seem to happen (both the old and the new buttons retain the default-button
					// appearance until the window gets entirely redraw such as via alt-tab).  This is fairly
					// inexplicable since the exact same technique works with "GuiControl +Default".  In any
					// case, this is kept because it also serves to change the default button appearance later,
					// which is received in the WindowProc via WM_COMMAND:
					SendMessage(mControl[mDefaultButtonIndex]->hwnd, BM_SETSTYLE
						, (WPARAM)LOWORD((GetWindowLong(mControl[mDefaultButtonIndex]->hwnd, GWL_STYLE) & ~BS_DEFPUSHBUTTON))
						, MAKELPARAM(TRUE, 0)); // Redraw = yes. It's probably smart enough not to do it if the window is hidden.
					// The below attempts to get the old button to lose its default-border failed.  This might
					// be due to the fact that if the window hasn't yet been shown for the first time, its
					// client area isn't yet the right size, so the OS decides that no update is needed since
					// the control is probably outside the boundaries of the window:
					//InvalidateRect(mHwnd, NULL, TRUE);
					//GetClientRect(mControl[mDefaultButtonIndex].hwnd, &client_rect);
					//InvalidateRect(mControl[mDefaultButtonIndex].hwnd, &client_rect, TRUE);
					//ShowWindow(mHwnd, SW_SHOWNOACTIVATE); // i.e. don't activate it if it wasn't before.
					//ShowWindow(mHwnd, SW_HIDE);
					//UpdateWindow(mHwnd);
					//SendMessage(mHwnd, WM_NCPAINT, 1, 0);
					//RedrawWindow(mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
					//SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
				}
				mDefaultButtonIndex = mControlCount;
				SendMessage(mHwnd, DM_SETDEFID, (WPARAM)GUI_INDEX_TO_ID(mDefaultButtonIndex), 0);
				// Strangely, in spite of the control having been created with the BS_DEFPUSHBUTTON style,
				// need to send BM_SETSTYLE or else the default button will lack its visual style when the
				// dialog is first shown.  Also strange is that the following must be done *after*
				// removing the visual/default style from the old default button and/or after doing
				// DM_SETDEFID above.
				SendMessage(control.hwnd, BM_SETSTYLE, (WPARAM)LOWORD(style), MAKELPARAM(TRUE, 0)); // Redraw = yes. It's probably smart enough not to do it if the window is hidden.
			}
		}
		break;

	case GUI_CONTROL_CHECKBOX:
		// The BS_NOTIFY style is not a good idea for checkboxes because although it causes the control
		// to send BN_DBLCLK messages, any rapid clicks by the user on (for example) a tri-state checkbox
		// are seen only as one click for the purpose of changing the box's state.
		if (control.hwnd = CreateWindowEx(exstyle, _T("Button"), aText, style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (opt.checked != BST_UNCHECKED) // Set the specified state.
				SendMessage(control.hwnd, BM_SETCHECK, opt.checked, 0);
		}
		break;

	case GUI_CONTROL_RADIO:
		control.hwnd = CreateWindowEx(exstyle, _T("Button"), aText, style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL);
		// opt.checked is handled later below.
		break;

	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
		// It has been verified that that EM_LIMITTEXT has no effect when sent directly
		// to a ComboBox hwnd; however, it might work if sent to its edit-child. But for now,
		// a Combobox can only be limited to its visible width.  Later, there might
		// be a way to send a message to its child control to limit its width directly.
		if (opt.limit && control.type == GUI_CONTROL_COMBOBOX)
			style &= ~CBS_AUTOHSCROLL;
		// Since the control's variable can change, it seems best to pass in the empty string
		// as the control's caption, rather than the name of the variable.  The name of the variable
		// isn't that useful anymore anyway since GuiControl(Get) can access controls directly by
		// their current output-var names:
		if (control.hwnd = CreateWindowEx(exstyle, _T("ComboBox"), _T(""), style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			// Set font unconditionally to simplify calculations, which help ensure that at least one item
			// in the DropDownList/Combo is visible when the list drops down:
			GUI_SETFONT // Set font in preparation for asking it how tall each item is.
			if (calc_control_height_from_row_count)
			{
				item_height = (int)SendMessage(control.hwnd, CB_GETITEMHEIGHT, 0, 0);
				// Note that at this stage, height should contain a explicitly-specified height or height
				// estimate from the above, even if row_count is greater than 0.
				// The below calculation may need some fine tuning:
				int cbs_extra_height = ((style & CBS_SIMPLE) && !(style & CBS_DROPDOWN)) ? 4 : 2;
				min_list_height = (2 * item_height) + GUI_CTL_VERTICAL_DEADSPACE + cbs_extra_height;
				if (opt.height < min_list_height) // Adjust so that at least 1 item can be shown.
					opt.height = min_list_height;
				else if (opt.row_count > 0)
					// Now that we know the true item height (since the control has been created and we asked
					// it), resize the control to try to get it to the match the specified number of rows.
					// +2 seems to be the exact amount needed to prevent partial rows from showing
					// on all font sizes and themes when NOINTEGRALHEIGHT is in effect:
					opt.height = (int)(opt.row_count * item_height) + GUI_CTL_VERTICAL_DEADSPACE + cbs_extra_height;
			}
			MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint, since it might be visible.
			// Since combo's size is created to accommodate its drop-down height, adjust our true height
			// to its actual collapsed size.  This true height is used for auto-positioning the next
			// control, if it uses auto-positioning.  It might be possible for it's width to be different
			// also, such as if it snaps to a certain minimize width if one too small was specified,
			// so that is recalculated too:
			retrieve_dimensions = true;
		}
		break;

	case GUI_CONTROL_LISTBOX:
		// See GUI_CONTROL_COMBOBOX above for why empty string is passed in as the caption:
		if (control.hwnd = CreateWindowEx(exstyle, _T("ListBox"), _T(""), style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (opt.tabstop_count)
				SendMessage(control.hwnd, LB_SETTABSTOPS, opt.tabstop_count, (LPARAM)opt.tabstop);
			// For now, it seems best to always override a height that would cause zero items to be
			// displayed.  This is because there is a very thin control visible even if the height
			// is explicitly set to zero, which seems pointless (there are other ways to draw thin
			// looking objects for unusual purposes anyway).
			// Set font unconditionally to simplify calculations, which help ensure that at least one item
			// in the DropDownList/Combo is visible when the list drops down:
			GUI_SETFONT // Set font in preparation for asking it how tall each item is.
			item_height = (int)SendMessage(control.hwnd, LB_GETITEMHEIGHT, 0, 0);
			// Note that at this stage, height should contain a explicitly-specified height or height
			// estimate from the above, even if opt.row_count is greater than 0.
			min_list_height = item_height + GUI_CTL_VERTICAL_DEADSPACE;
			if (style & WS_HSCROLL)
				// Assume bar will be actually appear even though it won't in the rare case where
				// its specified pixel-width is smaller than the width of the window:
				min_list_height += GetSystemMetrics(SM_CYHSCROLL);
			if (opt.height < min_list_height) // Adjust so that at least 1 item can be shown.
				opt.height = min_list_height;
			else if (opt.row_count > 0)
			{
				// Now that we know the true item height (since the control has been created and we asked
				// it), resize the control to try to get it to the match the specified number of rows.
				opt.height = (int)(opt.row_count * item_height) + GUI_CTL_VERTICAL_DEADSPACE;
				if (style & WS_HSCROLL)
					// Assume bar will be actually appear even though it won't in the rare case where
					// its specified pixel-width is smaller than the width of the window:
				opt.height += GetSystemMetrics(SM_CYHSCROLL);
			}
			MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint, since it might be visible.
			// Since by default, the OS adjusts list's height to prevent a partial item from showing
			// (LBS_NOINTEGRALHEIGHT), fetch the actual height for possible use in positioning the
			// next control:
			retrieve_dimensions = true;
		}
		break;

	case GUI_CONTROL_LISTVIEW:
		if (opt.listview_view != LV_VIEW_TILE) // It was ensured earlier that listview_view can be set to LV_VIEW_TILE only for XP or later.
			style = (style & ~LVS_TYPEMASK) | opt.listview_view; // Create control in the correct view mode whenever possible (TILE is the exception because it can't be expressed via style).
		if (control.hwnd = CreateWindowEx(exstyle, WC_LISTVIEW, _T(""), style, opt.x, opt.y // exstyle does apply to ListViews.
			, opt.width, opt.height == COORD_UNSPECIFIED ? 200 : opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (   !(control.union_lv_attrib = (lv_attrib_type *)malloc(sizeof(lv_attrib_type)))   )
			{
				// Since mem alloc problem is so rare just get rid of the control and flag it to be reported
				// later below as "cannot create control".  Doing this avoids the need to every worry whether
				// control.union_lv_attrib is NULL in other places.
				DestroyWindow(control.hwnd);
				control.hwnd = NULL;
				break;
			}
			// Otherwise:
			ZeroMemory(control.union_lv_attrib, sizeof(lv_attrib_type));
			control.union_lv_attrib->sorted_by_col = -1; // Indicate that there is currently no sort order.
			control.union_lv_attrib->no_auto_sort = opt.listview_no_auto_sort;

			// v1.0.36.06: If this ListView is owned by a tab control, flag that tab control as needing
			// to stay after all of its controls in the z-order.  This solves ListView-inside-Tab redrawing
			// problems, namely the disappearance of the ListView or an incomplete drawing of it.
			if (owning_tab_control)
				owning_tab_control->attrib |= GUI_CONTROL_ATTRIB_ALTBEHAVIOR;

			// Seems best to put tile view into effect before applying any styles that might be dependent upon it:
			if (opt.listview_view == LV_VIEW_TILE) // An earlier stage has verified that this is true only if OS is XP or later.
				SendMessage(control.hwnd, LVM_SETVIEW, LV_VIEW_TILE, 0);
			if (opt.listview_style) // This is a third set of styles that exist in addition to normal & extended.
				ListView_SetExtendedListViewStyle(control.hwnd, opt.listview_style); // No return value. Will have no effect on Win95/NT that lack comctl32.dll 4.70+ distributed with MSIE 3.x.
			ControlSetListViewOptions(control, opt); // Relies on adjustments to opt.color and color_bk done higher above.

			if (opt.height == COORD_UNSPECIFIED) // Adjust the control's size to fit opt.row_count rows.
			{
				// Known limitation: This will be inaccurate if an ImageList is later assigned to the
				// ListView because that increases the height of each row slightly (or a lot if
				// a large-icon list is forced into Details/Report view and other views that are
				// traditionally small-icon).  The code size and complexity of trying to compensate
				// for this doesn't seem likely to be worth it.
				GUI_SETFONT  // Required before asking it for a height estimate.
				switch (opt.listview_view)
				{
				case LV_VIEW_DETAILS:
					// The following formula has been tested on XP with the point sizes 8, 9, 10, 12, 14, and 18 for:
					// Verdana
					// Courier New
					// Gui Default Font
					// Times New Roman
					opt.height = 4 + HIWORD(ListView_ApproximateViewRect(control.hwnd, -1, -1
						, (WPARAM)opt.row_count - 1)); // -1 seems to be needed to make it calculate right for LV_VIEW_DETAILS.
					// Above: It seems best to exclude any horiz. scroll bar from consideration, even though it will
					// block the last row if bar is present.  The bar can be dismissed by manually dragging the
					// column dividers or using the GuiControl auto-size methods.
					// Note that ListView_ApproximateViewRect() is not available on 95/NT4 that lack
					// comctl32.dll 4.70+ distributed with MSIE 3.x  Therefore, rather than having a possibly-
					// complicated work around in the code to detect DLL version, it will be documented in
					// the help file that the "rows" method will produce an incorrect height on those platforms.
					break;

				case LV_VIEW_TILE: // This one can be safely integrated with the LVS ones because it doesn't overlap with them.
					// The following approach doesn't seem to give back useful info about the total height of a
					// tile and the border beneath it, so it isn't used:
					//LVTILEVIEWINFO tvi;
					//tvi.cbSize = sizeof(LVTILEVIEWINFO);
					//tvi.dwMask = LVTVIM_TILESIZE | LVTVIM_LABELMARGIN;
					//ListView_GetTileViewInfo(control.hwnd, &tvi);
					// The following might not be perfect for integral scrolling purposes, but it does seem
					// correct in terms of allowing exactly the right number of rows to be visible when the
					// control is scrolled all the way to the top. It's also correct for eliminating a
					// vertical scroll bar if the icons all fit into the specified number of rows. Tested on
					// XP Theme and Classic theme.
					opt.height = 7 + (int)((HIWORD(ListView_GetItemSpacing(control.hwnd, FALSE)) - 3) * opt.row_count);
					break;

				default: // Namely the following:
				//case LV_VIEW_ICON:
				//case LV_VIEW_SMALLICON:
				//case LV_VIEW_LIST:
					// For these non-report views, it seems far better to define row_count as the number of
					// icons that can fit vertically rather than as the total number of icons, because the
					// latter can result in heights that vary based on too many factors, resulting in too
					// much inconsistency.
					GUI_SET_HDC
					GetTextMetrics(hdc, &tm);
					if (opt.listview_view == LV_VIEW_ICON)
					{
						// The vertical space between icons is not dependent upon font size.  In other words,
						// the control's total height to fit exactly N rows would be icon_height*N plus
						// icon_spacing*(N-1).  However, the font height is added so that the last row has
						// enough extra room to display one line of text beneath the icon.  The first/constant
						// number below is a combination of two components: 1) The control's internal margin
						// that it maintains to decide when to display scroll bars (perhaps 3 above and 3 below).
						// The space between the icon and its first line of text (which seems constant, perhaps 10).
						opt.height = 16 + tm.tmHeight + (int)(GetSystemMetrics(SM_CYICON) * opt.row_count
							+ HIWORD(ListView_GetItemSpacing(control.hwnd, FALSE) * (opt.row_count - 1)));
						// More complex and doesn't seem as accurate:
						//float half_icon_spacing = 0.5F * HIWORD(ListView_GetItemSpacing(control.hwnd, FALSE));
						//opt.height = (int)(((HIWORD(ListView_ApproximateViewRect(control.hwnd, 5, 5, 1)) 
						//	+ half_icon_spacing + 4) * opt.row_count) + ((half_icon_spacing - 17) * (opt.row_count - 1)));
					}
					else // SMALLICON or LIST. For simplicity, it's done the same way for both, though it doesn't work as well for LIST.
					{
						// Seems way too high even with "TRUE": HIWORD(ListView_GetItemSpacing(control.hwnd, TRUE)
						int cy_smicon = GetSystemMetrics(SM_CYSMICON);
						// 11 seems to be the right value to prevent unwanted vertical scroll bar in SMALLICON view:
						opt.height = 11 + (int)((cy_smicon > tm.tmHeight ? cy_smicon : tm.tmHeight) * opt.row_count
							+ 1 * (opt.row_count - 1));
						break;
					}
					break;
				} // switch()
				MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint should be smart enough not to do it if window is hidden.
			} // if (opt.height == COORD_UNSPECIFIED)
		} // CreateWindowEx() succeeded.
		break;

	case GUI_CONTROL_TREEVIEW:
		if (control.hwnd = CreateWindowEx(exstyle, WC_TREEVIEW, _T(""), style, opt.x, opt.y
			, opt.width, opt.height == COORD_UNSPECIFIED ? 200 : opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (opt.checked)
				// Testing confirms that unless the following advice is applied, an item's checkbox cannot
				// be checked immediately after the item is created:
				// MSDN: If you want to use the checkbox style, you must set the TVS_CHECKBOXES style (with
				// SetWindowLong) after you create the tree-view control and before you populate the tree.
				// Otherwise, the checkboxes might appear unchecked, depending on timing issues.
				SetWindowLong(control.hwnd, GWL_STYLE, style | TVS_CHECKBOXES);
			ControlSetTreeViewOptions(control, opt); // Relies on adjustments to opt.color and color_bk done higher above.
			if (opt.himagelist) // Currently only supported upon creation, not via GuiControl, since in that case the decision of whether to destroy the old imagelist would be uncertain.
				TreeView_SetImageList(control.hwnd, opt.himagelist, TVSIL_NORMAL); // Currently no error reporting.

			if (opt.height == COORD_UNSPECIFIED) // Adjust the control's size to fit opt.row_count rows.
			{
				// Known limitation (may exist for TreeViews the same as it does for ListViews):
				// The follow might be inaccurate if an ImageList is later assigned to the TreeView because
				// that may increase the height of each row. The code size and complexity of trying to
				// compensate for this doesn't seem likely to be worth it.
				GUI_SETFONT  // Required before asking it for a height estimate.
				opt.height = TreeView_GetItemHeight(control.hwnd);
				// The following formula has been tested on XP fonts DefaultGUI, Verdana, Courier (for a few
				// point sizes).
				opt.height = 4 + (int)(opt.row_count * opt.height);
				// Above: It seems best to exclude any horiz. scroll bar from consideration, even though it will
				// block the last row if bar is present.  The bar can be dismissed by manually dragging the
				// column dividers or using the GuiControl auto-size methods.
				// Note that ListView_ApproximateViewRect() is not available on 95/NT4 that lack
				// comctl32.dll 4.70+ distributed with MSIE 3.x  Therefore, rather than having a possibly-
				// complicated work around in the code to detect DLL version, it will be documented in
				// the help file that the "rows" method will produce an incorrect height on those platforms.
				MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint should be smart enough not to do it if window is hidden.
			} // if (opt.height == COORD_UNSPECIFIED)
		} // CreateWindowEx() succeeded.
	break;

	case GUI_CONTROL_EDIT:
	{
		if (!(style & ES_MULTILINE)) // ES_MULTILINE was not explicitly or automatically specified.
		{
			if (opt.limit < 0) // This is the signal to limit input length to visible width of field.
				// But it can only work if the Edit isn't a multiline.
				style &= ~(WS_HSCROLL|ES_AUTOHSCROLL); // Enable the limiting style.
			else // Since this is a single-line edit, add AutoHScroll if it wasn't explicitly removed.
				style |= ES_AUTOHSCROLL & ~opt.style_remove;
				// But no attempt is made to turn off WS_VSCROLL or ES_WANTRETURN since those might have some
				// usefulness even in a single-line edit?  In any case, it seems too overprotective to do so.
		}
		// malloc() is done because edit controls in NT/2k/XP support more than 64K.
		// Mem alloc errors are so rare (since the text is usually less than 32K/64K) that no error is displayed.
		// Instead, the un-translated text is put in directly.  Also, translation is not done for
		// single-line edits since they can't display line breaks correctly anyway.
		// Note that TranslateLFtoCRLF() will return the original buffer we gave it if no translation
		// is needed.  Otherwise, it will return a new buffer which we are responsible for freeing
		// when done (or NULL if it failed to allocate the memory).
		auto malloc_buf = (*aText && (style & ES_MULTILINE)) ? TranslateLFtoCRLF(aText) : nullptr;
		if (control.hwnd = CreateWindowEx(exstyle, _T("Edit"), malloc_buf ? malloc_buf : aText, style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			// As documented in MSDN, setting a password char will have no effect for multi-line edits
			// since they do not support password/mask char.
			// It seems best to allow password_char to be a literal asterisk so that there's a way to
			// have asterisk vs. bullet/closed-circle on OSes that default to bullet.
			if ((style & ES_PASSWORD) && opt.password_char) // Override default.
				SendMessage(control.hwnd, EM_SETPASSWORDCHAR, (WPARAM)opt.password_char, 0);
			if (opt.limit < 0)
				opt.limit = 0;
			//else leave it as the zero (unlimited) or positive (restricted) limit already set.
			// Now set the limit. Specifying a limit of zero opens the control to its maximum text capacity,
			// which removes the 32K size restriction.  Testing shows that this does not increase the actual
			// amount of memory used for controls containing small amounts of text.  All it does is allow
			// the control to allocate more memory as the user enters text.
			SendMessage(control.hwnd, EM_LIMITTEXT, (WPARAM)opt.limit, 0); // EM_LIMITTEXT == EM_SETLIMITTEXT
			if (opt.tabstop_count)
				SendMessage(control.hwnd, EM_SETTABSTOPS, opt.tabstop_count, (LPARAM)opt.tabstop);
		}
		if (malloc_buf != aText)
			free(malloc_buf);
		break;
	}

	case GUI_CONTROL_DATETIME:
	{
		bool use_custom_format = false;
		if (*aText)
		{
			// DTS_SHORTDATEFORMAT and DTS_SHORTDATECENTURYFORMAT seem to produce identical results
			// (both display 4-digit year), at least on XP.  Perhaps DTS_SHORTDATECENTURYFORMAT is
			// obsolete.  In any case, it's uncommon so for simplicity, is not a named style.  It
			// can always be applied numerically if desired. Update: DTS_SHORTDATECENTURYFORMAT is
			// now applied by default higher above, which can be overridden explicitly via -0x0C
			// in the control's options.
			if (!_tcsicmp(aText, _T("LongDate"))) // LongDate seems more readable than "Long".  It also matches the keyword used by FormatTime.
				style = (style & ~(DTS_SHORTDATECENTURYFORMAT | DTS_TIMEFORMAT)) | DTS_LONGDATEFORMAT; // Purify.
			else if (!_tcsicmp(aText, _T("Time")))
				style = (style & ~(DTS_SHORTDATECENTURYFORMAT | DTS_LONGDATEFORMAT)) | DTS_TIMEFORMAT;  // Purify.
			else // Custom format. Don't purify (to retain the underlying default in case custom format is ever removed).
				use_custom_format = true;
		}
		if (opt.choice == 2) // "ChooseNone" was present, so ensure DTS_SHOWNONE is present to allow it.
			style |= DTS_SHOWNONE;
		//else it's blank, so retain the default DTS_SHORTDATEFORMAT (0x0000).
		if (control.hwnd = CreateWindowEx(exstyle, DATETIMEPICK_CLASS, _T(""), style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (use_custom_format)
				DateTime_SetFormat(control.hwnd, aText);
			//else keep the default or the format set via style higher above.
			// Feels safer to do range prior to selection even though unlike GUI_CONTROL_MONTHCAL,
			// GUI_CONTROL_DATETIME tolerates them in the reverse order when one doesn't fit the other.
			if (opt.gdtr_range) // If the date/time set above is invalid in light of the following new range, the date will be automatically to the closest valid date.
				DateTime_SetRange(control.hwnd, opt.gdtr_range, opt.sys_time_range);
			//else keep default range, which is "unrestricted range".
			if (opt.choice) // The option "ChooseYYYYMMDD" was present and valid (or ChooseNone was present, choice==2)
				DateTime_SetSystemtime(control.hwnd, opt.choice == 1 ? GDT_VALID : GDT_NONE, opt.sys_time);
			//else keep default, which is although undocumented appears to be today's date+time, which certainly is the expected default.
			if (IS_AN_ACTUAL_COLOR(opt.color))
				DateTime_SetMonthCalColor(control.hwnd, MCSC_TEXT, opt.color);
			// Note: The DateTime_SetMonthCalFont() macro is never used because apparently it's not required
			// to set the font, or even to repaint.
		}
		break;
	}

	case GUI_CONTROL_MONTHCAL:
		if (*aText)
		{
			opt.gdtr = YYYYMMDDToSystemTime2(aText, opt.sys_time);
			if (!opt.gdtr)
			{
				delete pcontrol;
				return g_script.RuntimeError(ERR_INVALID_VALUE, aText);
			}
			if (opt.gdtr == (GDTR_MIN | GDTR_MAX)) // When range is present, multi-select is automatically put into effect.
				style |= MCS_MULTISELECT;  // Must be applied during control creation since it can't be changed afterward.
		}
		// Create the control with arbitrary width/height if no width/height were explicitly specified.
		// It will be resized after creation by querying the control:
		if (control.hwnd = CreateWindowEx(exstyle, MONTHCAL_CLASS, _T(""), style, opt.x, opt.y
			, opt.width < 0 ? 100 : opt.width  // Negative width has special meaning upon creation (see below).
			, opt.height == COORD_UNSPECIFIED ? 100 : opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (style & MCS_MULTISELECT) // Must do this prior to setting initial contents in case contents is a range greater than 7 days.
				MonthCal_SetMaxSelCount(control.hwnd, 366); // 7 days seems too restrictive a default, so expand.
			if (opt.gdtr_range) // If the date/time set above is invalid in light of the following new range, the date will be automatically to the closest valid date.
				MonthCal_SetRange(control.hwnd, opt.gdtr_range, opt.sys_time_range);
			//else keep default range, which is "unrestricted range".
			if (opt.gdtr) // An explicit selection, either a range or single date, is present.
			{
				if (style & MCS_MULTISELECT) // Must use range-selection even if selection is only one date.
				{
					if (opt.gdtr == GDTR_MIN) // No maximum is present, so set maximum to minimum.
						opt.sys_time[1] = opt.sys_time[0];
					//else just max, or both are present.  Assume both for code simplicity.
					MonthCal_SetSelRange(control.hwnd, opt.sys_time);
				}
				else
					MonthCal_SetCurSel(control.hwnd, opt.sys_time);
			}
			//else keep default, which is although undocumented appears to be today's date+time, which certainly is the expected default.
			if (IS_AN_ACTUAL_COLOR(opt.color))
				MonthCal_SetColor(control.hwnd, MCSC_TEXT, opt.color);
			GUI_SETFONT  // Required before asking it about its month size.
			if ((opt.width == COORD_UNSPECIFIED || opt.height == COORD_UNSPECIFIED)
				&& MonthCal_GetMinReqRect(control.hwnd, &rect))
			{
				// Autosize width and/or height by asking the control how big each month is.
				// MSDN: "The top and left members of lpRectInfo will always be zero. The right and bottom
				// members represent the minimum cx and cy required for the control."
				if (opt.width < 0) // Negative width vs. COORD_UNSPECIFIED are different in this case.
				{
					// MSDN: "The rectangle returned by MonthCal_GetMinReqRect does not include the width
					// of the "Today" string, if it is present. If the MCS_NOTODAY style is not set,
					// retrieve the rectangle that defines the "Today" string width by calling the
					// MonthCal_GetMaxTodayWidth macro. Use the larger of the two rectangles to ensure
					// that the "Today" string is not clipped.
					int month_width;
					if (style & MCS_NOTODAY) // No today-string, so width is always that from GetMinReqRect.
						month_width = rect.right;
					else // There will be a today-string present, so use the greater of the two widths.
					{
						month_width = MonthCal_GetMaxTodayWidth(control.hwnd);
						if (month_width < rect.right)
							month_width = rect.right;
					}
					if (opt.width == COORD_UNSPECIFIED) // Use default, which is to provide room for a single month.
						opt.width = month_width;
					else // It's some explicit negative number.  Use it as a multiplier to provide multiple months.
					{
						// Multiple months must need a little extra room for border between: 0.02 but 0.03 is okay.
						// For safety, a larger value is used.
						opt.width = -opt.width;
						// Provide room for each separator.  There's one separator for each month after the
						// first, and the separator always seems to be exactly 6 regardless of font face/size.
						// This has been tested on both Classic and XP themes.
						opt.width = opt.width*month_width + (opt.width - 1)*6;
					}
				}
				if (opt.height == COORD_UNSPECIFIED)
				{
					opt.height = rect.bottom; // Init for default and for use below (room for only a single month's height).
					if (opt.row_count > 0 && opt.row_count != 1.0) // row_count was explicitly specified by the script, so use its exact value, even if it isn't a whole number (for flexibility).
					{
						// Unlike horizontally stacked calendars, vertically stacking them produces no separator
						// between them.
						GUI_SET_HDC
						GetTextMetrics(hdc, &tm);
						// If there will be no today string, the height reported by MonthCal_GetMinReqRect
						// is not correct for use in calculating the height of more than one month stacked
						// vertically.  Must adjust it to make it display properly.
						if (style & MCS_NOTODAY) // No today string, but space is still reserved for it, so must compensate for that.
							opt.height += tm.tmHeight + 4; // Formula tested with Courier New and Verdana 8/10/12 with row counts between 1 and 5.
						opt.height = (int)(opt.height * opt.row_count); // Calculate height of all months.
						// Regardless of whether MCS_NOTODAY is present, the below is still the right formula.
						// Room for the today-string is reserved only once at the bottom (even without MCS_NOTODAY),
						// so need to subtract that (also note that some months have 6 rows and others only 5,
						// but there is whitespace padding in the case of 5 to make all months the same height).
						opt.height = (int)(opt.height - ((opt.row_count - 1) * (tm.tmHeight - 2)));
						// Above: -2 seems to work for Verdana and Courier 8/10/12/14/18.
					}
					//else opt.row_count was unspecified, so stay with the default set above of exactly
					// one month tall.
				}
				MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint should be smart enough not to do it if window is hidden.
			} // Width or height was unspecified.
		} // Control created OK.
		break;

	case GUI_CONTROL_HOTKEY:
		// In this case, not only doesn't the caption appear anywhere, it's not set either (or at least
		// not retrievable via GetWindowText()):
		if (control.hwnd = CreateWindowEx(exstyle, HOTKEY_CLASS, _T(""), style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (*aText)
				SendMessage(control.hwnd, HKM_SETHOTKEY, TextToHotkey(aText), 0);
			if (opt.limit > 0)
				SendMessage(control.hwnd, HKM_SETRULES, opt.limit, MAKELPARAM(HOTKEYF_CONTROL|HOTKEYF_ALT, 0));
				// Above must also specify Ctrl+Alt or some other default, otherwise the restriction will have
				// no effect.
		}
		break;

	case GUI_CONTROL_UPDOWN:
		// The buddy of an up-down can meaningfully be one of the following:
		//case GUI_CONTROL_EDIT:
		//case GUI_CONTROL_TEXT:
		//case GUI_CONTROL_GROUPBOX:
		//case GUI_CONTROL_BUTTON:
		//case GUI_CONTROL_CHECKBOX:
		//case GUI_CONTROL_RADIO:
		//case GUI_CONTROL_LISTBOX:
		// (But testing shows, not these):
		//case GUI_CONTROL_UPDOWN: An up-down will snap onto another up-down, but not have any meaningful effect.
		//case GUI_CONTROL_DROPDOWNLIST:
		//case GUI_CONTROL_COMBOBOX:
		//case GUI_CONTROL_LISTVIEW:
		//case GUI_CONTROL_TREEVIEW:
		//case GUI_CONTROL_HOTKEY:
		//case GUI_CONTROL_UPDOWN:
		//case GUI_CONTROL_SLIDER:
		//case GUI_CONTROL_PROGRESS:
		//case GUI_CONTROL_TAB:
		//case GUI_CONTROL_STATUSBAR: As expected, it doesn't work properly.

		// v1.0.44: Don't allow buddying of UpDown to StatusBar (this must be done prior to the next section).
		// UPDATE: Due to rarity and user-should-know-better, this is not checked for (to reduce code size):
		//if (prev_control && prev_control->type == GUI_CONTROL_STATUSBAR)
		//	style &= ~UDS_AUTOBUDDY;
		// v1.0.42.02: The below is a fix for tab controls that contain a ListView so that up-downs in the
		// tab control don't snap onto the tab control (due to the z-order change done by the ListView creation
		// section whenever a ListView exists inside a tab control).
		bool provide_buddy_manually;
		if (   provide_buddy_manually = (style & UDS_AUTOBUDDY)
			&& (mStatusBarHwnd // Added for v1.0.44.01 (fixed in v1.0.44.04): Since the status bar is pushed to the bottom of the z-order after adding each other control, must do manual buddying whenever an UpDown is added after the status bar (to prevent it from attaching to the status bar).
			|| (owning_tab_control // mControlCount is greater than zero whenever owning_tab_control!=NULL
				&& (owning_tab_control->attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR)))   )
			style &= ~UDS_AUTOBUDDY; // Remove it during control creation to avoid up-down snapping onto tab control.

		// Suppress any Edit control Change events which are generated by the UpDown
		// control implicitly changing the contents of its buddy Edit control.
		if (prev_control && prev_control->type == GUI_CONTROL_EDIT)
			prev_control->attrib |= GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS;

		// The control is created unconditionally because if UDS_AUTOBUDDY is in effect, need to create the
		// control to find out its position and size (since it snaps to its buddy).  That size can then be
		// retrieved and used to figure out how to resize the buddy in cases where its width-set-automatically
		// -based-on-contents should not be squished as a result of buddying.
		if (control.hwnd = CreateWindowEx(exstyle, UPDOWN_CLASS, _T(""), style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (provide_buddy_manually) // v1.0.42.02 (see comment where provide_buddy_manually is initialized).
				SendMessage(control.hwnd, UDM_SETBUDDY, (WPARAM)prev_control->hwnd, 0); // See StatusBar notes above.  Also, mControlCount>0 whenever provide_buddy_manually==true.
			if (   mControlCount // Ensure there is a previous control to snap onto (below relies on this check).
				&& ((style & UDS_AUTOBUDDY) || provide_buddy_manually)   )
			{
				// Since creation of a buddied up-down ignored the specified x/y and width/height, update them
				// for use here and also later below for updating mMaxExtentRight, etc.
				GetWindowRect(control.hwnd, &rect);
				MapWindowPoints(NULL, parent_hwnd, (LPPOINT)&rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
				opt.x = rect.left;
				opt.y = rect.top;
				opt.width = rect.right - rect.left;
				opt.height = rect.bottom - rect.top;
				// Get its buddy's rectangle for use in two places:
				RECT buddy_rect;
				GetWindowRect(prev_control->hwnd, &buddy_rect);
				MapWindowPoints(NULL, parent_hwnd, (LPPOINT)&buddy_rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
				// Note: It does not matter if UDS_HORZ is in effect because strangely, the up-down still
				// winds up on the left or right side of the buddy, not the top/bottom.
				if (mControlWidthWasSetByContents)
				{
					// Since the previous control's width was determined solely by the size of its contents,
					// enlarge the control to undo the narrowing just done by the buddying process.
					// This relies on the fact that during buddying, the UpDown was auto-sized and positioned
					// to fit its buddy.
					if (style & UDS_ALIGNRIGHT)
					{
						// Since moving an up-down's buddy is not enough to move the up-down,
						// so that must be shifted too:
						opt.x += opt.width;
						rect.right += opt.width; // Updated for use further below.
						MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint this control separately from its buddy below.
					}
					// Enlarge the buddy control to restore it to the size it had prior to being reduced by the
					// buddying process:
					buddy_rect.right += opt.width; // Must be updated for use in two places.
					MoveWindow(prev_control->hwnd, buddy_rect.left, buddy_rect.top
						, buddy_rect.right - buddy_rect.left, buddy_rect.bottom - buddy_rect.top, TRUE);
				}
				// Set x/y and width/height to be that of combined/buddied control so that auto-positioning
				// of future controls will see it as a single control:
				if (style & UDS_ALIGNRIGHT)
				{
					opt.x = buddy_rect.left;
					// Must calculate the total width of the combined control not as a sum of their two widths,
					// but as the different between right and left.  Otherwise, the width will be off by either
					// 2 or 3 because of the slight overlap between the two controls.
					opt.width = rect.right - buddy_rect.left;
				}
				else
					opt.width = buddy_rect.right - rect.left;
					//and opt.x set to the x position of the up-down, since it's on the leftmost side.
				// Leave opt.y and opt.height as-is.
				if (!opt.range_changed && prev_control->type == GUI_CONTROL_LISTBOX)
				{
					// ListBox buddy needs an inverted UpDown (if the UpDown is vertical) to work the way
					// you'd expect.
					opt.range_changed = true;
					if (style & UDS_HORZ) // Use MAXVAL because otherwise ListBox select will be restricted to the first 100 entries.
						opt.range_max = UD_MAXVAL;
					else
						opt.range_min = UD_MAXVAL;  // ListBox needs an inverted UpDown to work the way you'd expect.
				}
			} // The up-down snapped onto a buddy control.
			if (!opt.range_changed) // Control's default is wacky inverted negative, so provide 0-100 as a better/traditional default.
			{
				opt.range_changed = true;
				opt.range_max = 100;
			}
			ControlSetUpDownOptions(control, opt); // This must be done prior to the below.
			// Set the position unconditionally, even if aText is blank.  This causes a blank aText to be
			// see as zero, which ensures the buddy has a legal starting value (which won't happen if the
			// range does not include zero, since it would default to zero then).
			// MSDN: "If the parameter is outside the control's specified range, nPos will be set to the nearest
			// valid value."
			SendMessage(control.hwnd, UDM_SETPOS32, 0, ATOI(aText));
		} // Control was successfully created.
		if (prev_control && prev_control->type == GUI_CONTROL_EDIT)
			prev_control->attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable Change event.
		break;

	case GUI_CONTROL_SLIDER:
		if (control.hwnd = CreateWindowEx(exstyle, TRACKBAR_CLASS, _T(""), style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			ControlSetSliderOptions(control, opt); // Fix for v1.0.25.08: This must be done prior to the below.
			// The control automatically deals with out-of-range values by setting slider to min or max.
			// MSDN: "If this value is outside the control's maximum and minimum range, the position
			// is set to the maximum or minimum value."
			if (*aText)
				SendMessage(control.hwnd, TBM_SETPOS, TRUE, ControlInvertSliderIfNeeded(control, ATOI(aText)));
				// Above msg has no return value.
			//else leave it at the OS's default starting position (probably always the far left or top of the range).
		}
		break;

	case GUI_CONTROL_PROGRESS:
		if (control.hwnd = CreateWindowEx(exstyle, PROGRESS_CLASS, _T(""), style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			ControlSetProgressOptions(control, opt, style); // Fix for v1.0.27.01: This must be done prior to the below.
			// This has been confirmed though testing, even when the range is dynamically changed
			// after the control is created to something that no longer includes the bar's current
			// position: The control automatically deals with out-of-range values by setting bar to
			// min or max.
			if (*aText)
				SendMessage(control.hwnd, PBM_SETPOS, ATOI(aText), 0);
			//else leave it at the OS's default starting position (probably always the far left or top of the range).
			do_strip_theme = false;  // The above would have already stripped it if needed, so don't do it again.
		}
		break;

	case GUI_CONTROL_TAB:
		if (control.hwnd = CreateWindowEx(exstyle, WC_TABCONTROL, _T(""), style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			if (opt.tab_control_uses_dialog) // It's a Tab3 control.
			{
				// Create a child dialog to host the controls for all tabs.
				if (!CreateTabDialog(control, opt))
				{
					DestroyWindow(control.hwnd);
					control.hwnd = NULL;
					break;
				}
			}
			if (opt.tab_control_autosize)
				SetProp(control.hwnd, _T("ahk_autosize"), (HANDLE)opt.tab_control_autosize);
			// After a new tab control is created, default all subsequently created controls to belonging
			// to the first tab of this tab control: 
			mCurrentTabControlIndex = mTabControlCount;
			mCurrentTabIndex = 0;
			++mTabControlCount;
			// Override the tab's window-proc so that custom background color becomes possible:
			g_TabClassProc = (WNDPROC)SetWindowLongPtr(control.hwnd, GWLP_WNDPROC, (LONG_PTR)TabWindowProc);
			// Doesn't work to remove theme background from tab:
			//EnableThemeDialogTexture(control.hwnd, ETDT_DISABLE);
		}
		break;
		
	case GUI_CONTROL_ACTIVEX:
	{
		static bool sAtlAxInitialized = false;
		if (!sAtlAxInitialized)
		{
			typedef BOOL (WINAPI *MyAtlAxWinInit)();
			if (HMODULE hmodAtl = LoadLibrary(_T("atl")))
			{
				if (MyAtlAxWinInit fnAtlAxWinInit = (MyAtlAxWinInit)GetProcAddress(hmodAtl, "AtlAxWinInit"))
					sAtlAxInitialized = fnAtlAxWinInit();
				if (!sAtlAxInitialized)
					FreeLibrary(hmodAtl);
				// Otherwise, init was successful so don't unload it.
			}
			// If any of the above calls failed, attempt to create the window anyway:
		}
		if (control.hwnd = CreateWindowEx(exstyle, _T("AtlAxWin"), aText, style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL))
		{
			// A reference to the object is kept for two reasons:
			//  1) So Ctrl.Value can return the same wrapper object each time.
			//  2) To ensure that if the script subscribes to events with ComObjConnect(),
			//     they will continue to work as long as the GUI exists, rather than being
			//     disconnected the moment the script releases its reference.
			if (  !(control.union_object = ControlGetActiveX(control.hwnd))  )
			{
				DestroyWindow(control.hwnd);
				control.hwnd = NULL;
				break;
			}
		}
		break;
	}

	case GUI_CONTROL_CUSTOM:
		if (opt.customClassAtom == 0)
		{
			delete pcontrol;
			return g_script.RuntimeError(_T("A window class is required."));
		}
		control.hwnd = CreateWindowEx(exstyle, MAKEINTATOM(opt.customClassAtom), aText, style
			, opt.x, opt.y, opt.width, opt.height, parent_hwnd, control_id, g_hInstance, NULL);
		break;

	case GUI_CONTROL_STATUSBAR:
		if (control.hwnd = CreateStatusWindow(style, aText, mHwnd, (UINT)(size_t)control_id))
		{
			mStatusBarHwnd = control.hwnd;
			if (opt.color_bk != CLR_INVALID) // Explicit color change was requested.
				SendMessage(mStatusBarHwnd, SB_SETBKCOLOR, 0, opt.color_bk);
		}
		break;
	} // switch() for control creation.

	////////////////////////////////
	// Release the HDC if necessary.
	////////////////////////////////
	if (hdc)
	{
		if (hfont_old)
		{
			SelectObject(hdc, hfont_old); // Necessary to avoid memory leak.
			hfont_old = NULL;
		}
		ReleaseDC(mHwnd, hdc);
		hdc = NULL;
	}

	// Below also serves as a bug check, i.e. GUI_CONTROL_INVALID or some unknown type.
	if (!control.hwnd)
	{
		delete pcontrol;
		return g_script.RuntimeError(_T("Can't create control."));
	}
	// Otherwise the above control creation succeeded.
	++mControlCount;
	apControl = pcontrol;
	mControlWidthWasSetByContents = control_width_was_set_by_contents; // Set for use by next control, if any.

	if (control.type == GUI_CONTROL_RADIO)
	{
		if (opt.checked != BST_UNCHECKED)
			ControlCheckRadioButton(control, mControlCount - 1, opt.checked); // Also handles alteration of the group's tabstop, etc.
		//else since the control has just been created, there's no need to uncheck it or do any actions
		// related to unchecking, such as tabstop adjustment.
		mInRadioGroup = true; // Set here, only after creation was successful.
	}
	else // For code simplicity and due to rarity, even GUI_CONTROL_STATUSBAR starts a new radio group.
		mInRadioGroup = false;

	// Check style_remove vs. style because this control might be hidden just because it was added
	// to a tab that isn't active:
	if (opt.style_remove & WS_VISIBLE)
		control.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN;  // For use with tab controls.
	if (opt.style_add & WS_DISABLED)
		control.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED;

	// Strip theme from the control if called for:
	// It is stripped for radios, checkboxes, and groupboxes if they have a custom text color.
	// Otherwise the transparency and/or custom text color will not be obeyed on XP, at least when
	// a non-Classic theme is active.  For GroupBoxes, when a theme is active, it will obey 
	// custom background color but not a custom text color.  The same is true for radios and checkboxes.
	if (do_strip_theme || (control.union_color != CLR_DEFAULT && (control.type == GUI_CONTROL_CHECKBOX
		|| control.type == GUI_CONTROL_RADIO || control.type == GUI_CONTROL_GROUPBOX)) // GroupBox needs it too.
		|| (control.type == GUI_CONTROL_GROUPBOX && control.background_color == CLR_TRANSPARENT) // Tested and found to be necessary.
		|| (control.type == GUI_CONTROL_DROPDOWNLIST && control.background_color != CLR_INVALID))
		SetWindowTheme(control.hwnd, L"", L"");

	// Must set the font even if mCurrentFontIndex > 0, otherwise the bold SYSTEM_FONT will be used.
	// Note: Neither the slider's buddies nor itself are affected by the font setting, so it's not applied.
	// However, the buddies are affected at the time they are created if they are a type that uses a font.
	if (!font_was_set && uses_font_and_text_color)
		GUI_SETFONT

	if (opt.redraw == CONDITION_FALSE)
		SendMessage(control.hwnd, WM_SETREDRAW, FALSE, 0); // Disable redrawing for this control to allow contents to be added to it more quickly.
		// It's not necessary to do the following because by definition the control has just been created
		// and thus redraw can't have been off for it previously:
		//if (opt.redraw == CONDITION_TRUE) // Since redrawing is being turned back on, invalidate the control so that it updates itself.
		//	InvalidateRect(control.hwnd, NULL, TRUE);

	///////////////////////////////////////////////////
	// Add any content to the control and set its font.
	///////////////////////////////////////////////////
	ControlAddItems(control, aObj); // Must be done after font-set above so that ListView columns can be auto-sized to fit their text.
	if (opt.choice > 0)
		ControlSetChoice(control, opt.choice);

	if (control.type == GUI_CONTROL_TAB && opt.row_count > 0)
	{
		// Now that the tabs have been added (possibly creating more than one row of tabs), resize so that
		// the interior of the control has the actual number of rows specified.
		GetClientRect(control.hwnd, &rect); // MSDN: "the coordinates of the upper-left corner are (0,0)"
		// This is a workaround for the fact that TabCtrl_AdjustRect() seems to give an invalid
		// height when the tabs are at the bottom, at least on XP.  Unfortunately, this workaround
		// does not work when the tabs or on the left or right side, so don't even bother with that
		// adjustment (it's very rare that a tab control would have an unspecified width anyway).
		bool bottom_is_in_effect = (style & TCS_BOTTOM) && !(style & TCS_VERTICAL);
		if (bottom_is_in_effect)
			SetWindowLong(control.hwnd, GWL_STYLE, style & ~TCS_BOTTOM);
		// Insist on a taller client area (or same height in the case of TCS_VERTICAL):
		TabCtrl_AdjustRect(control.hwnd, TRUE, &rect); // Calculate new window height.
		if (bottom_is_in_effect)
			SetWindowLong(control.hwnd, GWL_STYLE, style);
		opt.height = rect.bottom - rect.top;  // Update opt.height for here and for later use below.
		// The below is commented out because TabCtrl_AdjustRect() is unable to cope with tabs on
		// the left or right sides.  It would be rarely used anyway.
		//if (style & TCS_VERTICAL && width_was_originally_unspecified)
		//	// Also make the interior wider in this case, to make the interior as large as intended.
		//	// It is a known limitation that this adjustment does not occur when the script did not
		//	// specify a row_count or omitted height and row_count.
		//	opt.width = rect.right - rect.left;
		MoveWindow(control.hwnd, opt.x, opt.y, opt.width, opt.height, TRUE); // Repaint, since parent might be visible.
	}
	else if (control.type == GUI_CONTROL_TAB && opt.tab_control_uses_dialog)
	{
		// Update this tab control's dialog to match the tab's display area.  This should be done
		// whenever the display area changes, such as when tabs are added for the first time or the
		// number of tab rows changes.  It's not done if (opt.row_count > 0), because the section
		// above would have already triggered an update via MoveWindow/WM_WINDOWPOSCHANGED.
		UpdateTabDialog(control.hwnd);
	}

	if (retrieve_dimensions) // Update to actual size for use later below.
	{
		GetWindowRect(control.hwnd, &rect);
		opt.height = rect.bottom - rect.top;
		opt.width = rect.right - rect.left;

		if (aControlType == GUI_CONTROL_LISTBOX && (style & WS_HSCROLL))
		{
			if (opt.hscroll_pixels < 0) // Calculate a default based on control's width.
				// Since horizontal scrollbar is relatively rarely used, no fancy method
				// such as calculating scrolling-width via LB_GETTEXTLEN & current font's
				// average width is used.
				opt.hscroll_pixels = 3 * opt.width;
			// If hscroll_pixels is now zero or smaller than the width of the control, the
			// scrollbar will not be shown.  But the message is still sent unconditionally
			// in case it has some desirable side-effects:
			SendMessage(control.hwnd, LB_SETHORIZONTALEXTENT, (WPARAM)opt.hscroll_pixels, 0);
		}
	}

	// v1.0.36.06: If this tab control contains a ListView, keep the tab control after all of its controls
	// in the z-order.  This solves ListView-inside-Tab redrawing problems, namely the disappearance of
	// the ListView or an incomplete drawing of it.  Doing it this way preserves the tab-navigation
	// order of controls inside the tab control, both those above and those beneath the ListView.
	// The only thing it alters is the tab navigation to the tab control itself, which will now occur
	// after rather than before all the controls inside it. For most uses, this a very minor difference,
	// especially given the rarity of having ListViews inside tab controls.  If this solution ever
	// proves undesirable, one alternative is to somehow force the ListView to properly redraw whenever
	// its inside a tab control.  Perhaps this could be done by subclassing the ListView or Tab control
	// and having it do something different or additional in response to WM_ERASEBKGND.  It might
	// also be done in the parent window's proc in response to WM_ERASEBKGND.
	if (owning_tab_control && parent_hwnd == mHwnd) // The control is on a tab control which is its sibling.
	{
		// Fix for v1.0.35: Probably due to clip-siblings, adding a control within the area of a tab control
		// does not properly draw the control.  This seems to apply to most/all control types.
		if (on_visible_page_of_tab_control)
		{
			// Not enough for GUI_CONTROL_DATETIME (it's border is not drawn):
			//InvalidateRect(control.hwnd, NULL, TRUE);  // TRUE is required, at least for GUI_CONTROL_DATETIME.
			GetWindowRect(control.hwnd, &rect);
			MapWindowPoints(NULL, mHwnd, (LPPOINT)&rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
			InvalidateRect(mHwnd, &rect, FALSE);
		}
		if (owning_tab_control->attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR) // Put the tab control after the newly added control. See comment higher above.
			SetWindowPos(owning_tab_control->hwnd, control.hwnd, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);
	}

	// v1.0.44: Must keep status bar at the bottom of the z-order so that it gets drawn last.  This alleviates
	// (but does not completely prevent) other controls from overlapping it and getting drawn on top. This is
	// done each time a control is added -- rather than at some single time such as when the parent window is
	// first shown -- in case the script adds more controls later.
	if (mStatusBarHwnd) // Seems harmless to do it even if the just-added control IS the status bar. Also relies on the fact that that only one status bar is allowed.
		SetWindowPos(mStatusBarHwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE);

	if (aControlType != GUI_CONTROL_STATUSBAR) // i.e. don't let status bar affect positioning of controls relative to each other.
	{
		if (parent_hwnd != mHwnd) // This control is in a tab dialog.
		{
			// Translate the coordinates back to Gui-relative client coordinates.
			POINT pt = { opt.x, opt.y };
			MapWindowPoints(parent_hwnd, mHwnd, &pt, 1);
			opt.x = pt.x, opt.y = pt.y;
		}

		////////////////////////////////////////////////////////////////////////////////////////////////////
		// Save the details of this control's position for possible use in auto-positioning the next control.
		////////////////////////////////////////////////////////////////////////////////////////////////////
		mPrevX = opt.x;
		mPrevY = opt.y;
		mPrevWidth = opt.width;
		mPrevHeight = opt.height;
		int right = opt.x + opt.width;
		int bottom = opt.y + opt.height;
		if (parent_hwnd == mHwnd // Skip this for controls which are clipped by a parent control.
			&& !opt.tab_control_autosize) // For auto-sizing Tab3, don't adjust mMaxExtent until later.
		{
			if (right > mMaxExtentRight)
				mMaxExtentRight = right;
			if (bottom > mMaxExtentDown)
				mMaxExtentDown = bottom;
		}

		// As documented, always start new section for very first control, but never if this control is GUI_CONTROL_STATUSBAR.
		if (opt.start_new_section || mControlCount == 1 // aControlType!=GUI_CONTROL_STATUSBAR due to check higher above.
			|| (mControlCount == 2 && mControl[0]->type == GUI_CONTROL_STATUSBAR)) // This is the first non-statusbar control.
		{
			mSectionX = opt.x;
			mSectionY = opt.y;
			mMaxExtentRightSection = right;
			mMaxExtentDownSection = bottom;
		}
		else
		{
			if (right > mMaxExtentRightSection)
				mMaxExtentRightSection = right;
			if (bottom > mMaxExtentDownSection)
				mMaxExtentDownSection = bottom;
		}
	}

	return OK;
}



FResult GuiType::Opt(StrArg aOptions)
{
	GUI_MUST_HAVE_HWND;
	bool set_last_found_window = false;
	ToggleValueType own_dialogs = TOGGLE_INVALID;
	if (!ParseOptions(aOptions, set_last_found_window, own_dialogs))
		return FR_FAIL; // Above already displayed an error message.
	if (set_last_found_window)
		g->hWndLastUsed = mHwnd;
	SetOwnDialogs(own_dialogs);
	return OK;
}



ResultType GuiType::ParseOptions(LPCTSTR aOptions, bool &aSetLastFoundWindow, ToggleValueType &aOwnDialogs)
// This function is similar to ControlParseOptions() further below, so should be maintained alongside it.
// Caller must have already initialized aSetLastFoundWindow/, bool &aOwnDialogs with desired starting values.
// Caller must ensure that aOptions is a modifiable string, since this method temporarily alters it.
{
	if (mHwnd)
	{
		// Since window already exists, its mStyle and mExStyle members might be out-of-date due to
		// "WinSet Transparent", etc.  So update them:
		mStyle = GetWindowLong(mHwnd, GWL_STYLE);
		mExStyle = GetWindowLong(mHwnd, GWL_EXSTYLE);
	}
	DWORD style_orig = mStyle;
	DWORD exstyle_orig = mExStyle;

	TCHAR option[MAX_NUMBER_SIZE + 1]; // Enough for any single option word or number, with room to avoid issues due to truncation.
	LPCTSTR next_option, option_end;
	DWORD option_dword;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = aOptions; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		for (option_end = next_option; *option_end && !IS_SPACE_OR_TAB(*option_end); ++option_end);
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.
		
		// Make a null-terminated copy to simplify comparisons below.
		tcslcpy(option, next_option, min((option_end - next_option) + 1, _countof(option)));
		
		// Attributes and option words:

		bool set_owner;
		if ((set_owner = !_tcsnicmp(option, _T("Owner"), 5))
			  || !_tcsnicmp(option, _T("Parent"), 6))
		{
			if (!adding)
				mOwner = NULL;
			else
			{
				auto name = option + 5 + !set_owner; // 6 for "Parent"
				if (*name || !set_owner) // i.e. "+Parent" on its own is invalid (and should not default to g_hWnd).
				{
					HWND new_owner = NULL;
					UINT hwnd_uint;
					if (ParseInteger(name, hwnd_uint))
						if (IsWindow((HWND)(UINT_PTR)hwnd_uint))
							new_owner = (HWND)(UINT_PTR)hwnd_uint;
					if (!new_owner || new_owner == mHwnd) // Window can't own itself!
					{
						if (!ValueError(_T("Invalid or nonexistent owner or parent window."), option, FAIL_OR_OK))
							return FAIL;
						// Otherwise, user wants to continue.
					}
					mOwner = new_owner;
				}
				else
					mOwner = g_hWnd; // Make a window owned (by script's main window) omits its taskbar button.
			}
			if (set_owner) // +/-Owner
			{
				if (mStyle & WS_CHILD)
				{
					// Since Owner and Parent are mutually-exclusive by nature, it seems appropriate for
					// +Owner to also apply -Parent.  If this wasn't done, +Owner would merely change the
					// parent window (see comments further below).
					mStyle = mStyle & ~WS_CHILD | WS_POPUP;
					if (mHwnd)
					{
						// This seems to be necessary even if SetWindowLong() is called to update the
						// style before attempting to change the owner.  By contrast, there doesn't seem
						// to be any problem with delaying the style update until after all of the other
						// options are parsed.
						SetParent(mHwnd, NULL);
					}
				}
				// Although MSDN doesn't explicitly document any way to change the owner of an existing
				// window, the following method was mentioned in a PDC talk by Raymond Chen.  MSDN does
				// say it shouldn't be used to change the parent of a child window, but maybe what it
				// actually means is that "HWNDPARENT" is a misnomer; it should've been "HWNDOWNER".
				// On the other hand, this method ACTUALLY DOES CHANGE THE PARENT WINDOW if mHwnd is a
				// child window and the check above is disabled (at least it did during testing).
				if (mHwnd)
					SetWindowLongPtr(mHwnd, GWLP_HWNDPARENT, (LONG_PTR)mOwner);
			}
			else // +/-Parent
			{
				if (mHwnd)
					SetParent(mHwnd, mOwner);
				if (mOwner)
					mStyle = mStyle & ~WS_POPUP | WS_CHILD;
				else
					mStyle = mStyle & ~WS_CHILD | WS_POPUP;
				// The new style will be applied by a later section.
			}
		}

		else if (!_tcsicmp(option, _T("AlwaysOnTop")))
		{
			// If the window already exists, SetWindowLong() isn't enough.  Must use SetWindowPos()
			// to make it take effect.
			if (mHwnd)
			{
				SetWindowPos(mHwnd, adding ? HWND_TOPMOST : HWND_NOTOPMOST, 0, 0, 0, 0
					, SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE); // SWP_NOACTIVATE prevents the side-effect of activating the window, which is undesirable if only its style is changing.
				// Fix for v1.0.41.01: Update the original style too, so that the call to SetWindowLong() later below
				// is made only if multiple styles are being changed on the same line, e.g. Gui +Disabled -SysMenu
				if (adding) exstyle_orig |= WS_EX_TOPMOST; else exstyle_orig &= ~WS_EX_TOPMOST;
			}
			// Fix for v1.0.41.01: The following line is now done unconditionally.  Previously, it wasn't
			// done if the window already existed, which caused an example such as the following to first
			// set the window always on top and then immediately afterward try to unset it via SetWindowLong
			// (because mExStyle hadn't been updated to reflect the change made by SetWindowPos):
			// Gui, +AlwaysOnTop +Disabled -SysMenu
			if (adding) mExStyle |= WS_EX_TOPMOST; else mExStyle &= ~WS_EX_TOPMOST;
		}

		else if (!_tcsicmp(option, _T("Border")))
			if (adding) mStyle |= WS_BORDER; else mStyle &= ~WS_BORDER;

		else if (!_tcsicmp(option, _T("Caption")))
			if (adding) mStyle |= WS_CAPTION; else mStyle = mStyle & ~WS_CAPTION;

		else if (!_tcsicmp(option, _T("Disabled")))
		{
			if (mHwnd)
			{
				EnableWindow(mHwnd, adding ? FALSE : TRUE);  // Must not not apply WS_DISABLED directly because that breaks the window.
				// Fix for v1.0.41.01: Update the original style too, so that the call to SetWindowLong() later below
				// is made only if multiple styles are being changed on the same line, e.g. Gui +Disabled -SysMenu
				if (adding) style_orig |= WS_DISABLED; else style_orig &= ~WS_DISABLED;
			}
			// Fix for v1.0.41.01: The following line is now done unconditionally.  Previously, it wasn't
			// done if the window already existed, which caused an example such as the following to first
			// disable the window and then immediately afterward try to enable it via SetWindowLong
			// (because mStyle hadn't been updated to reflect the change made by SetWindowPos):
			// Gui, +AlwaysOnTop +Disabled -SysMenu
			if (adding) mStyle |= WS_DISABLED; else mStyle &= ~WS_DISABLED;
		}
		
		else if (!_tcsicmp(option, _T("LastFound")))
			aSetLastFoundWindow = true; // Regardless of whether "adding" is true or false.

		else if (!_tcsicmp(option, _T("MaximizeBox"))) // See above comment.
			if (adding) mStyle |= WS_MAXIMIZEBOX|WS_SYSMENU; else mStyle &= ~WS_MAXIMIZEBOX;

		else if (!_tcsicmp(option, _T("MinimizeBox")))
			// WS_MINIMIZEBOX requires WS_SYSMENU to take effect.  It can be explicitly omitted
			// via "+MinimizeBox -SysMenu" if that functionality is ever needed.
			if (adding) mStyle |= WS_MINIMIZEBOX|WS_SYSMENU; else mStyle &= ~WS_MINIMIZEBOX;

		else if (!_tcsnicmp(option, _T("MinSize"), 7)) // v1.0.44.13: Added for use with WM_GETMINMAXINFO.
		{
			if (adding)
			{
				ParseMinMaxSizeOption(option + 7, mMinWidth, mMinHeight);
			}
			else // "-MinSize", so tell the WM_GETMINMAXINFO handler to use system defaults.
			{
				mMinWidth = COORD_UNSPECIFIED;
				mMinHeight = COORD_UNSPECIFIED;
			}
		}

		else if (!_tcsnicmp(option, _T("MaxSize"), 7)) // v1.0.44.13: Added for use with WM_GETMINMAXINFO.
		{
			if (adding)
			{
				ParseMinMaxSizeOption(option + 7, mMaxWidth, mMaxHeight);
			}
			else // "-MaxSize", so tell the WM_GETMINMAXINFO handler to use system defaults.
			{
				mMaxWidth = COORD_UNSPECIFIED;
				mMaxHeight = COORD_UNSPECIFIED;
			}
		}

		else if (!_tcsicmp(option, _T("OwnDialogs")))
			aOwnDialogs = (adding ? TOGGLED_ON : TOGGLED_OFF);

		else if (!_tcsicmp(option, _T("Resize"))) // Minus removes either or both.
			if (adding) mStyle |= WS_SIZEBOX|WS_MAXIMIZEBOX; else mStyle &= ~(WS_SIZEBOX|WS_MAXIMIZEBOX);

		else if (!_tcsicmp(option, _T("SysMenu")))
			if (adding) mStyle |= WS_SYSMENU; else mStyle &= ~WS_SYSMENU;

		else if (!_tcsicmp(option, _T("Theme")))
			mUseTheme = adding;
			// But don't apply/remove theme from parent window because that is usually undesirable.
			// This is because even old apps running on XP still have the new parent window theme,
			// at least for their title bar and title bar buttons (except console apps, maybe).

		else if (!_tcsicmp(option, _T("ToolWindow")))
			// WS_EX_TOOLWINDOW provides narrower title bar, omits task bar button, and omits
			// entry in the alt-tab menu.
			if (adding) mExStyle |= WS_EX_TOOLWINDOW; else mExStyle &= ~WS_EX_TOOLWINDOW;

		else if (!_tcsicmp(option, _T("DPIScale")))
			mUsesDPIScaling = adding;

		// This one should be near the bottom since "E" is fairly vague and might be contained at the start
		// of future option words such as Edge, Exit, etc.
		else if (ctoupper(*option) == 'E' && ParsePositiveInteger(option + 1, option_dword)) // Extended style
		{
			if (adding)
				mExStyle |= option_dword;
			else
				mExStyle &= ~option_dword;
		}

		else if (ParsePositiveInteger(option, option_dword)) // Above has already verified that *option can't be whitespace.
		{
			// Pure numbers are assumed to be style additions or removals:
			if (adding)
				mStyle |= option_dword;
			else
				mStyle &= ~option_dword;
		}

		else // v1.1.04: Validate Gui options.
		{
			if (!ValueError(ERR_INVALID_OPTION, option, FAIL_OR_OK))
				return FAIL;
			// Otherwise, user wants to continue.
		}

	} // for() each item in option list

	// Besides reducing the code size and complexity, another reason all changes to style are made
	// here rather than above is that multiple changes might have been made above to the style,
	// and there's no point in redrawing/updating the window for each one:
	if (mHwnd && (mStyle != style_orig || mExStyle != exstyle_orig))
	{
		// v1.0.27.01: Must do this prior to SetWindowLong() because sometimes SetWindowLong()
		// traumatizes the window (such as "Gui -Caption"), making it effectively invisible
		// even though its non-functional remnant is still on the screen:
		bool is_visible = IsWindowVisible(mHwnd) && !IsIconic(mHwnd);

		// Since window already exists but its style has changed, attempt to update it dynamically.
		if (mStyle != style_orig)
			SetWindowLong(mHwnd, GWL_STYLE, mStyle);
		if (mExStyle != exstyle_orig)
			SetWindowLong(mHwnd, GWL_EXSTYLE, mExStyle);

		if (is_visible)
		{
			// Hiding then showing is the only way I've discovered to make it update.  If the window
			// is not updated, a strange effect occurs where the window is still visible but can no
			// longer be used at all (clicks pass right through it).  This show/hide method is less
			// desirable due to possible side effects caused to any script that happens to be watching
			// for its existence/non-existence, so it would be nice if some better way can be discovered
			// to do this.
			// Update by Lexikos: I've been unable to produce the "strange effect" described above using
			// WinSet on Win2k/XP/7, so changing this section to use the same method as WinSet. If there
			// are any problems, at least WinSet and Gui will be consistent.
			// SetWindowPos is also necessary, otherwise the frame thickness entirely around the window
			// does not get updated (just parts of it):
			SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
			// ShowWindow(mHwnd, SW_HIDE);
			// ShowWindow(mHwnd, SW_SHOWNA); // i.e. don't activate it if it wasn't before. Note that SW_SHOWNA avoids restoring the window if it is currently minimized or maximized (unlike SW_SHOWNOACTIVATE).
			InvalidateRect(mHwnd, NULL, TRUE); // Quite a few styles require this to become visibly manifest.
			// None of the following methods alone is enough, at least not when the window is currently active:
			// 1) InvalidateRect(mHwnd, NULL, TRUE);
			// 2) SendMessage(mHwnd, WM_NCPAINT, 1, 0);  // 1 = Repaint entire frame.
			// 3) RedrawWindow(mHwnd, NULL, NULL, RDW_INVALIDATE|RDW_FRAME|RDW_UPDATENOW);
			// 4) SetWindowPos(mHwnd, NULL, 0, 0, 0, 0, SWP_DRAWFRAME|SWP_FRAMECHANGED|SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE);
		}
		// Continue on to create the window so that code is simplified in other places by
		// using the assumption that "if gui[i] object exists, so does its window".
		// Another important reason this is done is that if an owner window were to be destroyed
		// before the window it owns is actually created, the WM_DESTROY logic would have to check
		// for any windows owned by the window being destroyed and update them.
	}

	return OK;
}



void GuiType::GetNonClientArea(LONG &aWidth, LONG &aHeight)
// Added for v1.0.44.13.
// Yields only the *extra* width/height added by the windows non-client area.
// If the window hasn't been shown for the first time, the caller wants zeros.
// The reason for making the script specify size of client area rather than entire window is that it
// seems far more useful.  For example, a script might know exactly how much minimum height its
// controls require in the client area, but would find it inconvenient to have to take into account
// the height of the title bar and menu bar (which vary depending on theme and other settings).
{
	if (mGuiShowHasNeverBeenDone) // In this case, the script might not yet have added the menu bar and other styles that affect the size of the non-client area.  So caller wants to do these calculations later.
	{
		aWidth = 0;
		aHeight = 0;
		return;
	}
	// Otherwise, mGuiShowHasNeverBeenDone==false, which should mean that mHwnd!=NULL.
	RECT rect, client_rect;
	GetWindowRect(mHwnd, &rect);
	GetClientRect(mHwnd, &client_rect); // Client rect's left & top are always zero.
	aWidth = (rect.right - rect.left) - client_rect.right;
	aHeight = (rect.bottom - rect.top) - client_rect.bottom;
}



void GuiType::GetTotalWidthAndHeight(LONG &aWidth, LONG &aHeight)
// Added for v1.0.44.13.
// Yields total width and height of entire window.
// If the window hasn't been shown for the first time, the caller wants COORD_CENTERED.
{
	if (mGuiShowHasNeverBeenDone)
	{
		aWidth = COORD_CENTERED;
		aHeight = COORD_CENTERED;
		return;
	}
	// Otherwise, mGuiShowHasNeverBeenDone==false, which should mean that mHwnd!=NULL.
	RECT rect;
	GetWindowRect(mHwnd, &rect);
	aWidth = rect.right - rect.left;
	aHeight = rect.bottom - rect.top;
}



void GuiType::ParseMinMaxSizeOption(LPCTSTR option_value, LONG &aWidth, LONG &aHeight)
{
	if (*option_value)
	{
		// The following will retrieve zeros if window hasn't yet been shown for the first time,
		// in which case the first showing will do the NC adjustment for us.  The overall approach
		// used here was chose to avoid any chance for Min/MaxSize to be adjusted more than once
		// to convert client size to entire-size, which would be wrong since the adjustment must be
		// applied only once.  Examples of such situations are when one of the coordinates is omitted,
		// or when +MinSize is specified prior to the first "Gui Show" but +MaxSize is specified after.
		LONG nc_width, nc_height;
		GetNonClientArea(nc_width, nc_height);
		// _ttoi() vs. ATOI() is used below to avoid ambiguity of "x" being hex 0x vs. a delimiter.
		auto pos_of_the_x = StrChrAny(option_value, _T("Xx"));
		if (pos_of_the_x && pos_of_the_x[1]) // Kept simple due to rarity of transgressions and their being inconsequential.
			aHeight = Scale(_ttoi(pos_of_the_x + 1)) + nc_height;
		//else it's "MinSize333" or "MinSize333x", so leave height unchanged as documented.
		if (pos_of_the_x != option_value) // There's no 'x' or it lies to the right of option_value.
			aWidth = Scale(_ttoi(option_value)) + nc_width; // _ttoi() automatically stops converting when it reaches non-numeric character.
		//else it's "MinSizeX333", so leave width unchanged as documented.
	}
	else // Since no width or height was specified:
		// Use the window's current size. But if window hasn't yet been shown for the
		// first time, this will set the values to COORD_CENTERED, which tells the
		// first-show routine to get the total width/height upon first showing (since
		// that's where the window's initial size is determined).
		GetTotalWidthAndHeight(aWidth, aHeight);
}



ResultType GuiType::ControlParseOptions(LPCTSTR aOptions, GuiControlOptionsType &aOpt, GuiControlType &aControl
	, GuiIndexType aControlIndex)
// Caller must have already initialized aOpt with zeroes or any other desired starting values.
// Caller must ensure that aOptions is a modifiable string, since this method temporarily alters it.
{
	TCHAR option[MAX_NUMBER_SIZE + 1]; // Enough for any single option word or number, with room to avoid issues due to truncation.
	LPCTSTR next_option, option_end;
	LPCTSTR error_message; // Used by "return_error:" when aControl.hwnd == NULL.
	DWORD option_dword;
	bool adding; // Whether this option is being added (+) or removed (-).
	GuiControlType *tab_control;
	RECT rect;
	POINT pt;
	bool do_invalidate_rect = false; // Set default.

	for (next_option = aOptions; *next_option; next_option = omit_leading_whitespace(option_end))
	{
		if (*next_option == '-')
		{
			adding = false;
			// omit_leading_whitespace() is not called, which enforces the fact that the option word must
			// immediately follow the +/- sign.  This is done to allow the flexibility to have options
			// omit the plus/minus sign, and also to reserve more flexibility for future option formats.
			++next_option;  // Point it to the option word itself.
		}
		else
		{
			// Assume option is being added in the absence of either sign.  However, when we were
			// called by GuiControl(), the first option in the list must begin with +/- otherwise the cmd
			// would never have been properly detected as GUICONTROL_CMD_OPTIONS in the first place.
			adding = true;
			if (*next_option == '+')
				++next_option;  // Point it to the option word itself.
			//else do not increment, under the assumption that the plus has been omitted from a valid
			// option word and is thus an implicit plus.
		}

		if (!*next_option) // In case the entire option string ends in a naked + or -.
			break;
		// Find the end of this option item:
		for (option_end = next_option; *option_end && !IS_SPACE_OR_TAB(*option_end); ++option_end);
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.
		
		// Make a null-terminated copy to simplify comparisons below.
		tcslcpy(option, next_option, min((option_end - next_option) + 1, _countof(option)));

		// Attributes:
		if (!_tcsicmp(option, _T("Section"))) // Adding and removing are treated the same in this case.
			aOpt.start_new_section = true;    // Ignored by caller when control already exists.
		else if (!_tcsicmp(option, _T("AltSubmit")) && aControl.type != GUI_CONTROL_EDIT)
		{
			// v1.0.44: Don't allow control's AltSubmit bit to be set unless it's valid option for
			// that type.  This protects the GUI_CONTROL_ATTRIB_ALTSUBMIT bit from being corrupted
			// in control types that use it for other/internal purposes.  Update: For code size reduction
			// and performance, only exclude control types that use the ALTSUBMIT bit for an internal
			// purpose vs. allowing the script to set it via "AltSubmit".
			if (adding) aControl.attrib |= GUI_CONTROL_ATTRIB_ALTSUBMIT; else aControl.attrib &= ~GUI_CONTROL_ATTRIB_ALTSUBMIT;
			//switch(aControl.type)
			//{
			//case GUI_CONTROL_TAB:
			//case GUI_CONTROL_PIC:
			//case GUI_CONTROL_DROPDOWNLIST:
			//case GUI_CONTROL_COMBOBOX:
			//case GUI_CONTROL_LISTBOX:
			//case GUI_CONTROL_LISTVIEW:
			//case GUI_CONTROL_TREEVIEW:
			//case GUI_CONTROL_MONTHCAL:
			//case GUI_CONTROL_SLIDER:
			//	if (adding) aControl.attrib |= GUI_CONTROL_ATTRIB_ALTSUBMIT; else aControl.attrib &= ~GUI_CONTROL_ATTRIB_ALTSUBMIT;
			//	break;
			//// All other types either use the bit for some internal purpose or want it reserved for possible
			//// future use.  So don't allow the presence of "AltSubmit" to change the bit.
			//}
		}

		// Content of control (these are currently only effective if the control is being newly created):
		else if (!_tcsnicmp(option, _T("Checked"), 7)) // Caller knows to ignore if inapplicable. Applicable for ListView too.
		{
			auto option_value = option + 7;
			if (!_tcsicmp(option_value, _T("Gray"))) // Radios can't have the 3rd/gray state, but for simplicity it's permitted.
				if (adding) aOpt.checked = BST_INDETERMINATE; else aOpt.checked = BST_UNCHECKED;
			else
			{
				if (aControl.type == GUI_CONTROL_LISTVIEW)
					if (adding) aOpt.listview_style |= LVS_EX_CHECKBOXES; else aOpt.listview_style &= ~LVS_EX_CHECKBOXES;
				else
				{
					// As of v1.0.26, Checked/Hidden/Disabled can be followed by an optional 1/0/-1 so that
					// there is a way for a script to set the starting state by reading from an INI or registry
					// entry that contains 1 or 0 instead of needing the literal word "checked" stored in there.
					// Otherwise, a script would have to do something like the following before every "Gui Add":
					// if Box1Enabled
					//    Enable = Enabled
					// else
					//    Enable =
					// Gui Add, checkbox, %Enable%, My checkbox.
					if (*option_value) // There's more after the word, namely a 1, 0, or -1.
					{
						aOpt.checked = ATOI(option_value);
						if (aOpt.checked == -1)
							aOpt.checked = BST_INDETERMINATE;
					}
					else // Below is also used for GUI_CONTROL_TREEVIEW creation because its checkboxes must be added AFTER the control is created.
						aOpt.checked = adding; // BST_CHECKED == 1, BST_UNCHECKED == 0
				}
			} // Non-checkedGRAY
		} // Checked.
		else if (!_tcsnicmp(option, _T("Choose"), 6))
		{
			// "CHOOSE" provides an easier way to conditionally select a different item at the time
			// the control is added.  Example: gui, add, ListBox, vMyList Choose%choice%, %MyItemList%
			// Caller should ignore aOpt.choice if it isn't applicable for this control type.
			if (adding)
			{
				auto option_value = option + 6;
				switch (aControl.type)
				{
				case GUI_CONTROL_DATETIME:
					if (!_tcsicmp(option_value, _T("None")))
						aOpt.choice = 2; // Special flag value to indicate "none".
					else // See if it's a valid date-time.
						if (YYYYMMDDToSystemTime(option_value, aOpt.sys_time[0], true)) // Date string is valid.
							aOpt.choice = 1; // Overwrite 0 to flag sys_time as both present and valid.
						else
						{
							error_message = ERR_INVALID_OPTION;
							goto return_error;
						}
					break;
				default:
					aOpt.choice = ATOI(option_value);
					if (aOpt.choice < 1) // Invalid: number should be 1 or greater.
						aOpt.choice = 0; // Flag it as invalid.
				}
			}
			//else do nothing (not currently implemented)
		}

		// Styles (general):
		else if (!_tcsicmp(option, _T("Border")))
			if (adding) aOpt.style_add |= WS_BORDER; else aOpt.style_remove |= WS_BORDER;
		else if (!_tcsicmp(option, _T("VScroll"))) // Seems harmless in this case not to check aControl.type to ensure it's an input-capable control.
			if (adding) aOpt.style_add |= WS_VSCROLL; else aOpt.style_remove |= WS_VSCROLL;
		else if (!_tcsnicmp(option, _T("HScroll"), 7)) // Seems harmless in this case not to check aControl.type to ensure it's an input-capable control.
		{
			if (aControl.type == GUI_CONTROL_TREEVIEW)
				// Testing shows that Tree doesn't seem to fully support removal of hscroll bar after creation.
				if (adding) aOpt.style_remove |= TVS_NOHSCROLL; else aOpt.style_add |= TVS_NOHSCROLL;
			else
				if (adding)
				{
					// MSDN: "To respond to the LB_SETHORIZONTALEXTENT message, the list box must have
					// been defined with the WS_HSCROLL style."
					aOpt.style_add |= WS_HSCROLL;
					auto option_value = option + 7;
					aOpt.hscroll_pixels = *option_value ? ATOI(option_value) : -1;  // -1 signals it to use a default based on control's width.
				}
				else
					aOpt.style_remove |= WS_HSCROLL;
		}
		else if (!_tcsicmp(option, _T("Tabstop"))) // Seems harmless in this case not to check aControl.type to ensure it's an input-capable control.
			if (adding) aOpt.style_add |= WS_TABSTOP; else aOpt.style_remove |= WS_TABSTOP;
		else if (!_tcsicmp(option, _T("NoTab"))) // Supported for backward compatibility and it might be more ergonomic for "Gui Add".
			if (adding) aOpt.style_remove |= WS_TABSTOP; else aOpt.style_add |= WS_TABSTOP;
		else if (!_tcsicmp(option, _T("group")))
			if (adding) aOpt.style_add |= WS_GROUP; else aOpt.style_remove |= WS_GROUP;
		else if (!_tcsicmp(option, _T("Redraw")))  // Seems a little more intuitive/memorable than "Draw".
			aOpt.redraw = adding ? CONDITION_TRUE : CONDITION_FALSE; // Otherwise leave it at its default of 0.
		else if (!_tcsnicmp(option, _T("Disabled"), 8))
		{
			// As of v1.0.26, Checked/Hidden/Disabled can be followed by an optional 1/0/-1 so that
			// there is a way for a script to set the starting state by reading from an INI or registry
			// entry that contains 1 or 0 instead of needing the literal word "checked" stored in there.
			// Otherwise, a script would have to do something like the following before every "Gui Add":
			// if Box1Enabled
			//    Enable = Enabled
			// else
			//    Enable =
			// Gui Add, checkbox, %Enable%, My checkbox.
			if (option[8] && !ATOI(option + 8)) // If it's Disabled0, invert the mode to become "enabled".
				adding = !adding;
			if (aControl.hwnd) // More correct to call EnableWindow and let it set the style.  Do not set the style explicitly in this case since that might break it.
				EnableWindow(aControl.hwnd, adding ? FALSE : TRUE);
			else
				if (adding) aOpt.style_add |= WS_DISABLED; else aOpt.style_remove |= WS_DISABLED;
			// Update the "explicitly" flags so that switching tabs won't reset the control's state.
			if (adding)
				aControl.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED;
			else
				aControl.attrib &= ~GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED;
		}
		else if (!_tcsnicmp(option, _T("Hidden"), 6))
		{
			// As of v1.0.26, Checked/Hidden/Disabled can be followed by an optional 1/0/-1 so that
			// there is a way for a script to set the starting state by reading from an INI or registry
			// entry that contains 1 or 0 instead of needing the literal word "checked" stored in there.
			// Otherwise, a script would have to do something like the following before every "Gui Add":
			// if Box1Enabled
			//    Enable = Enabled
			// else
			//    Enable =
			// Gui Add, checkbox, %Enable%, My checkbox.
			if (option[6] && !ATOI(option + 6)) // If it's Hidden0, invert the mode to become "show".
				adding = !adding;
			if (aControl.hwnd) // More correct to call ShowWindow() and let it set the style.  Do not set the style explicitly in this case since that might break it.
				ShowWindow(aControl.hwnd, adding ? SW_HIDE : SW_SHOWNOACTIVATE);
			else
				if (adding) aOpt.style_remove |= WS_VISIBLE; else aOpt.style_add |= WS_VISIBLE;
			// Update the "explicitly" flags so that switching tabs won't reset the control's state.
			if (adding)
				aControl.attrib |= GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN;
			else
				aControl.attrib &= ~GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN;
		}
		else if (!_tcsicmp(option, _T("Wrap")))
		{
			switch(aControl.type)
			{
			case GUI_CONTROL_TEXT: // This one is a little tricky but the below should be appropriate in most cases:
				if (adding) aOpt.style_remove |= SS_TYPEMASK; else aOpt.style_add = (aOpt.style_add & ~SS_TYPEMASK) | SS_LEFTNOWORDWRAP; // v1.0.44.10: Added SS_TYPEMASK to "else" section to provide more graceful handling for cases like "-Wrap +Center", which would otherwise put an unexpected style like SS_OWNERDRAW into effect.
				break;
			case GUI_CONTROL_GROUPBOX:
			case GUI_CONTROL_BUTTON:
			case GUI_CONTROL_CHECKBOX:
			case GUI_CONTROL_RADIO:
				if (adding) aOpt.style_add |= BS_MULTILINE; else aOpt.style_remove |= BS_MULTILINE;
				break;
			case GUI_CONTROL_UPDOWN:
				if (adding) aOpt.style_add |= UDS_WRAP; else aOpt.style_remove |= UDS_WRAP;
				break;
			case GUI_CONTROL_EDIT: // Must be a multi-line now or shortly in the future or these will have no effect.
				if (adding) aOpt.style_remove |= WS_HSCROLL|ES_AUTOHSCROLL; else aOpt.style_add |= ES_AUTOHSCROLL;
				// WS_HSCROLL is removed because with it, wrapping is automatically off.
				break;
			case GUI_CONTROL_TAB:
				if (adding) aOpt.style_add |= TCS_MULTILINE; else aOpt.style_remove |= TCS_MULTILINE;
				// WS_HSCROLL is removed because with it, wrapping is automatically off.
				break;
			// N/A for these:
			//case GUI_CONTROL_PIC:
			//case GUI_CONTROL_DROPDOWNLIST:
			//case GUI_CONTROL_COMBOBOX:
			//case GUI_CONTROL_LISTBOX:
			//case GUI_CONTROL_LISTVIEW:
			//case GUI_CONTROL_TREEVIEW:
			//case GUI_CONTROL_DATETIME:
			//case GUI_CONTROL_MONTHCAL:
			//case GUI_CONTROL_HOTKEY:
			//case GUI_CONTROL_SLIDER:
			//case GUI_CONTROL_PROGRESS:
			//case GUI_CONTROL_LINK:
			}
		}
		else if (!_tcsnicmp(option, _T("Background"), 10))
		{
			auto option_value = option + 10;

			// Reset background properties to simplify the next section.
			aControl.background_color = CLR_INVALID;
			if (aControl.background_brush)
			{
				DeleteObject(aControl.background_brush);
				aControl.background_brush = NULL;
			}

			if (!_tcsicmp(option_value, _T("Trans")))
			{
				if (adding)
				{
					if (!aControl.SupportsBackgroundTrans())
					{
						error_message = ERR_GUI_NOT_FOR_THIS_TYPE;
						goto return_error;
					}
					aControl.background_color = CLR_TRANSPARENT;
				}
				//else do nothing, because the attrib was already removed above.
			}
			else if (adding)
			{
				if (*option_value)
				{
					if (!aControl.SupportsBackgroundColor())
					{
						error_message = ERR_GUI_NOT_FOR_THIS_TYPE;
						goto return_error;
					}
					if (!ColorToBGR(option_value, aOpt.color_bk))
					{
						error_message = ERR_INVALID_OPTION;
						goto return_error;
					}
					if (aOpt.color_bk != CLR_DEFAULT && aControl.RequiresBackgroundBrush())
						aControl.background_brush = CreateSolidBrush(aOpt.color_bk);
					aControl.background_color = aOpt.color_bk;
				}
				//else +Background (the traditional way to revert -Background).
				// background_brush and background_color were already reset above.
			}
			else // -Background
			{
				aControl.background_color = CLR_DEFAULT;
				aOpt.color_bk = CLR_DEFAULT;
			}
		} // Option "Background".
		else if (!_tcsicmp(option, _T("group"))) // This overlaps with g-label, but seems well worth it in this case.
			if (adding) aOpt.style_add |= WS_GROUP; else aOpt.style_remove |= WS_GROUP;
		else if (!_tcsicmp(option, _T("Theme")))
			aOpt.use_theme = adding;

		// Picture / ListView
		else if (!_tcsnicmp(option, _T("Icon"), 4)) // Caller should ignore aOpt.icon_number if it isn't applicable for this control type.
		{
			auto option_value = option + 4;
			if (aControl.type == GUI_CONTROL_LISTVIEW) // Unconditional regardless of the value of "adding".
				aOpt.listview_view = _tcsicmp(option_value, _T("Small")) ? LV_VIEW_ICON : LV_VIEW_SMALLICON;
			else
				if (adding)
					aOpt.icon_number = ATOI(option_value);
				//else do nothing (not currently implemented)
		}
		else if (!_tcsicmp(option, _T("Report")))
			aOpt.listview_view = LV_VIEW_DETAILS; // Unconditional regardless of the value of "adding".
		else if (!_tcsicmp(option, _T("List")))
			aOpt.listview_view = LV_VIEW_LIST; // Unconditional regardless of the value of "adding".
		else if (!_tcsicmp(option, _T("Tile"))) // Fortunately, subsequent changes to the control's style do not pop it out of Tile mode. It's apparently smart enough to do that only when the LVS_TYPEMASK bits change.
			aOpt.listview_view = LV_VIEW_TILE;
		else if (aControl.type == GUI_CONTROL_LISTVIEW && !_tcsicmp(option, _T("Hdr")))
			if (adding) aOpt.style_remove |= LVS_NOCOLUMNHEADER; else aOpt.style_add |= LVS_NOCOLUMNHEADER;
		else if (aControl.type == GUI_CONTROL_LISTVIEW && !_tcsnicmp(option, _T("NoSort"), 6))
		{
			if (!_tcsicmp(option + 6, _T("Hdr"))) // Prevents the header from being clickable like a set of buttons.
				if (adding) aOpt.style_add |= LVS_NOSORTHEADER; else aOpt.style_remove |= LVS_NOSORTHEADER; // Testing shows it can't be changed after the control is created.
			else // Header is still clickable (unless above is *also* specified), but has no automatic sorting.
				aOpt.listview_no_auto_sort = adding;
		}
		else if (aControl.type == GUI_CONTROL_LISTVIEW && !_tcsicmp(option, _T("Grid")))
			if (adding) aOpt.listview_style |= LVS_EX_GRIDLINES; else aOpt.listview_style &= ~LVS_EX_GRIDLINES;
		else if (!_tcsnicmp(option, _T("Count"), 5)) // Script should only provide the option for ListViews.
			aOpt.limit = ATOI(option + 5); // For simplicity, the value of "adding" is ignored.
		else if (!_tcsnicmp(option, _T("LV"), 2) && ParsePositiveInteger(option + 2, option_dword))
			if (adding) aOpt.listview_style |= option_dword; else aOpt.listview_style &= ~option_dword;
		else if (!_tcsnicmp(option, _T("ImageList"), 9))
		{
			if (adding)
				aOpt.himagelist = (HIMAGELIST)(size_t)ATOU(option + 9);
			//else removal not currently supported, since that would require detection of whether
			// to destroy the old imagelist, which is difficult to know because it might be in use
			// by other types of controls?
		}

		// Button
		else if (aControl.type == GUI_CONTROL_BUTTON && !_tcsicmp(option, _T("Default")))
			if (adding) aOpt.style_add |= BS_DEFPUSHBUTTON; else aOpt.style_remove |= BS_DEFPUSHBUTTON;
		else if (aControl.type == GUI_CONTROL_CHECKBOX && !_tcsicmp(option, _T("Check3"))) // Radios can't have the 3rd/gray state.
			if (adding) aOpt.style_add |= BS_AUTO3STATE; else aOpt.style_remove |= BS_AUTO3STATE;

		// Edit (and upper/lowercase for combobox/ddl, and others)
		else if (!_tcsicmp(option, _T("ReadOnly")))
		{
			switch (aControl.type)
			{
			case GUI_CONTROL_EDIT:
				if (aControl.hwnd) // Update the existing edit.  Must use SendMessage() vs. changing the style.
					SendMessage(aControl.hwnd, EM_SETREADONLY, adding, 0);
				else
					if (adding) aOpt.style_add |= ES_READONLY; else aOpt.style_remove |= ES_READONLY;
				break;
			case GUI_CONTROL_LISTBOX:
				if (adding) aOpt.style_add |= LBS_NOSEL; else aOpt.style_remove |= LBS_NOSEL;
				break;
			case GUI_CONTROL_LISTVIEW:
				if (adding) aOpt.style_remove |= LVS_EDITLABELS; else aOpt.style_add |= LVS_EDITLABELS;
				break;
			case GUI_CONTROL_TREEVIEW:
				if (adding) aOpt.style_remove |= TVS_EDITLABELS; else aOpt.style_add |= TVS_EDITLABELS;
				break;
			}
		}
		else if (!_tcsicmp(option, _T("Multi")))
		{
			// It was named "multi" vs. multiline and/or "MultiSel" because it seems easier to
			// remember in these cases.  In fact, any time two styles can be combined into one
			// name whose actual function depends on the control type, it seems likely to make
			// things easier to remember.
			switch(aControl.type)
			{
			case GUI_CONTROL_EDIT:
				if (adding) aOpt.style_add |= ES_MULTILINE; else aOpt.style_remove |= ES_MULTILINE;
				break;
			case GUI_CONTROL_LISTBOX:
				if (adding) aOpt.style_add |= LBS_EXTENDEDSEL; else aOpt.style_remove |= LBS_EXTENDEDSEL;
				break;
			case GUI_CONTROL_LISTVIEW:
				if (adding) aOpt.style_remove |= LVS_SINGLESEL; else aOpt.style_add |= LVS_SINGLESEL;
				break;
			case GUI_CONTROL_MONTHCAL:
				if (adding) aOpt.style_add |= MCS_MULTISELECT; else aOpt.style_remove |= MCS_MULTISELECT;
				break;
			}
		}
		else if (aControl.type == GUI_CONTROL_EDIT && !_tcsicmp(option, _T("WantReturn")))
			if (adding) aOpt.style_add |= ES_WANTRETURN; else aOpt.style_remove |= ES_WANTRETURN;
		else if (aControl.type == GUI_CONTROL_EDIT && !_tcsicmp(option, _T("WantTab")))
			if (adding) aControl.attrib |= GUI_CONTROL_ATTRIB_ALTBEHAVIOR; else aControl.attrib &= ~GUI_CONTROL_ATTRIB_ALTBEHAVIOR;
		else if (aControl.type == GUI_CONTROL_EDIT && !_tcsicmp(option, _T("WantCtrlA"))) // v1.0.44: Presence of AltSubmit bit means DON'T want Ctrl-A.
			if (adding) aControl.attrib &= ~GUI_CONTROL_ATTRIB_ALTSUBMIT; else aControl.attrib |= GUI_CONTROL_ATTRIB_ALTSUBMIT;
		else if ((aControl.type == GUI_CONTROL_LISTVIEW || aControl.type == GUI_CONTROL_TREEVIEW)
			&& !_tcsicmp(option, _T("WantF2"))) // v1.0.44: All an F2 keystroke to edit the focused item.
			// Since WantF2 is the initial default, a script will almost never specify WantF2.  Therefore, it's
			// probably not worth the code size to put -ReadOnly into effect automatically for +WantF2.
			if (adding) aControl.attrib &= ~GUI_CONTROL_ATTRIB_ALTBEHAVIOR; else aControl.attrib |= GUI_CONTROL_ATTRIB_ALTBEHAVIOR;
		else if (aControl.type == GUI_CONTROL_EDIT && !_tcsicmp(option, _T("Number")))
			if (adding) aOpt.style_add |= ES_NUMBER; else aOpt.style_remove |= ES_NUMBER;
		else if (!_tcsicmp(option, _T("Lowercase")))
		{
			if (aControl.type == GUI_CONTROL_EDIT)
				if (adding) aOpt.style_add |= ES_LOWERCASE; else aOpt.style_remove |= ES_LOWERCASE;
			else if (aControl.type == GUI_CONTROL_COMBOBOX || aControl.type == GUI_CONTROL_DROPDOWNLIST)
				if (adding) aOpt.style_add |= CBS_LOWERCASE; else aOpt.style_remove |= CBS_LOWERCASE;
		}
		else if (!_tcsicmp(option, _T("Uppercase")))
		{
			if (aControl.type == GUI_CONTROL_EDIT)
				if (adding) aOpt.style_add |= ES_UPPERCASE; else aOpt.style_remove |= ES_UPPERCASE;
			else if (aControl.type == GUI_CONTROL_COMBOBOX || aControl.type == GUI_CONTROL_DROPDOWNLIST)
				if (adding) aOpt.style_add |= CBS_UPPERCASE; else aOpt.style_remove |= CBS_UPPERCASE;
		}
		else if (aControl.type == GUI_CONTROL_EDIT && !_tcsnicmp(option, _T("Password"), 8))
		{
			// Allow a space to be the masking character, since it's conceivable that might
			// be wanted in cases where someone doesn't want anyone to know they're typing a password.
			// Simplest to assign unconditionally, regardless of whether adding or removing:
			aOpt.password_char = option[8];  // Can be '\0', which indicates "use OS default".
			if (adding)
			{
				aOpt.style_add |= ES_PASSWORD;
				if (aControl.hwnd) // Update the existing edit.
				{
					// Provide default since otherwise pass-char will be removed vs. added.
					// The encoding of the message's parameter depends on whether SendMessageW or SendMessageA is called.
					// 9679 corresponds to the default bullet character on XP and later (verified with EM_GETPASSWORDCHAR).
					if (!aOpt.password_char)
						SendMessageW(aControl.hwnd, EM_SETPASSWORDCHAR, 9679, 0);
					else
						SendMessage(aControl.hwnd, EM_SETPASSWORDCHAR, (WPARAM)aOpt.password_char, 0);
				}
			}
			else
			{
				aOpt.style_remove |= ES_PASSWORD;
				if (aControl.hwnd) // Update the existing edit.
					SendMessage(aControl.hwnd, EM_SETPASSWORDCHAR, 0, 0);
			}
			// Despite what MSDN seems to say, the control is not redrawn when the EM_SETPASSWORDCHAR
			// message is received.  This is required for the control to visibly update immediately:
			do_invalidate_rect = true;
		}
		else if (!_tcsnicmp(option, _T("Limit"), 5)) // This is used for Hotkey controls also.
		{
			if (adding)
			{
				auto option_value = option + 5;
				aOpt.limit = *option_value ? ATOI(option_value) : -1;  // -1 signals it to limit input to visible width of field.
				// aOpt.limit will later be ignored for some control types.
			}
			else
				aOpt.limit = INT_MIN; // Signal it to remove the limit.
		}

		// Combo/DropDownList/ListBox/ListView
		else if (aControl.type == GUI_CONTROL_COMBOBOX && !_tcsicmp(option, _T("Simple"))) // DDL is not equipped to handle this style.
			if (adding) aOpt.style_add |= CBS_SIMPLE; else aOpt.style_remove |= CBS_SIMPLE;
		else if (!_tcsnicmp(option, _T("Sort"), 4))
		{
			switch(aControl.type)
			{
			case GUI_CONTROL_LISTBOX:
				if (adding) aOpt.style_add |= LBS_SORT; else aOpt.style_remove |= LBS_SORT;
				break;
			case GUI_CONTROL_LISTVIEW: // LVS_SORTDESCENDING is not a named style due to rarity of use.
				if (adding)
					aOpt.style_add |= _tcsicmp(option + 4, _T("Desc")) ? LVS_SORTASCENDING : LVS_SORTDESCENDING;
				else
					aOpt.style_remove |= LVS_SORTASCENDING|LVS_SORTDESCENDING;
				break;
			case GUI_CONTROL_DROPDOWNLIST:
			case GUI_CONTROL_COMBOBOX:
				if (adding) aOpt.style_add |= CBS_SORT; else aOpt.style_remove |= CBS_SORT;
				break;
			}
		}

		// UpDown
		else if (aControl.type == GUI_CONTROL_UPDOWN && !_tcsicmp(option, _T("Horz")))
			if (adding)
			{
				aOpt.style_add |= UDS_HORZ;
				aOpt.style_add &= ~UDS_AUTOBUDDY; // Doing it this way allows "Horz +0x10" to override Horz's lack of buddy.
			}
			else
				aOpt.style_remove |= UDS_HORZ; // But don't add UDS_AUTOBUDDY since it seems undesirable most of the time.

		// Slider
		else if (aControl.type == GUI_CONTROL_SLIDER && !_tcsicmp(option, _T("Invert"))) // Not called "Reverse" to avoid confusion with the non-functional style of that name.
			if (adding) aControl.attrib |= GUI_CONTROL_ATTRIB_ALTBEHAVIOR; else aControl.attrib &= ~GUI_CONTROL_ATTRIB_ALTBEHAVIOR;
		else if (aControl.type == GUI_CONTROL_SLIDER && !_tcsicmp(option, _T("NoTicks")))
			if (adding) aOpt.style_add |= TBS_NOTICKS; else aOpt.style_remove |= TBS_NOTICKS;
		else if (aControl.type == GUI_CONTROL_SLIDER && !_tcsnicmp(option, _T("TickInterval"), 12))
		{
			aOpt.tick_interval_changed = true;
			if (adding)
			{
				aOpt.style_add |= TBS_AUTOTICKS;
				auto option_value = option + 12;
				aOpt.tick_interval_specified = *option_value;
				aOpt.tick_interval = ATOI(option_value);
			}
			else
			{
				aOpt.style_remove |= TBS_AUTOTICKS;
				aOpt.tick_interval = -1;  // Signal it to remove the ticks later below (if the window exists).
			}
		}
		else if (!_tcsnicmp(option, _T("Line"), 4))
		{
			auto option_value = option + 4;
			if (aControl.type == GUI_CONTROL_SLIDER)
			{
				if (adding)
					aOpt.line_size = ATOI(option_value);
				//else removal not supported.
			}
			else if (aControl.type == GUI_CONTROL_TREEVIEW && ctoupper(*option_value) == 'S')
				// Seems best to consider TVS_HASLINES|TVS_LINESATROOT to be an inseparable group since
				// one without the other is rare (script can always be overridden by specifying numeric styles):
				if (adding) aOpt.style_add |= TVS_HASLINES|TVS_LINESATROOT; else aOpt.style_remove |= TVS_HASLINES|TVS_LINESATROOT;
		}
		else if (aControl.type == GUI_CONTROL_SLIDER && !_tcsnicmp(option, _T("Page"), 4))
		{
			if (adding)
				aOpt.page_size = ATOI(option + 4);
			//else removal not supported.
		}
		else if (aControl.type == GUI_CONTROL_SLIDER && !_tcsnicmp(option, _T("Thick"), 5))
		{
			if (adding)
			{
				aOpt.style_add |= TBS_FIXEDLENGTH;
				aOpt.thickness = ATOI(option + 5);
			}
			else // Removing the style is enough to reset its appearance on both XP Theme and Classic Theme.
				aOpt.style_remove |= TBS_FIXEDLENGTH;
		}
		else if (!_tcsnicmp(option, _T("ToolTip"), 7))
		{
			auto option_value = option + 7;
			// Below was commented out because the SBARS_TOOLTIPS doesn't seem to do much, if anything.
			// See bottom of BIF_StatusBar() for more comments.
			//if (aControl.type == GUI_CONTROL_STATUSBAR)
			//{
			//	if (!*option_value)
			//		if (adding) aOpt.style_add |= SBARS_TOOLTIPS; else aOpt.style_remove |= SBARS_TOOLTIPS;
			//}
			//else
			if (aControl.type == GUI_CONTROL_SLIDER)
			{
				if (adding)
				{
					aOpt.tip_side = -1;  // Set default.
					switch(ctoupper(*option_value))
					{
					case 'T': aOpt.tip_side = TBTS_TOP; break;
					case 'L': aOpt.tip_side = TBTS_LEFT; break;
					case 'B': aOpt.tip_side = TBTS_BOTTOM; break;
					case 'R': aOpt.tip_side = TBTS_RIGHT; break;
					}
					if (aOpt.tip_side < 0)
						aOpt.tip_side = 0; // Restore to the value that means "use default side".
					else
						++aOpt.tip_side; // Offset by 1, since zero is reserved as "use default side".
					aOpt.style_add |= TBS_TOOLTIPS;
				}
				else
					aOpt.style_remove |= TBS_TOOLTIPS;
			}
		}
		else if (aControl.type == GUI_CONTROL_SLIDER && !_tcsnicmp(option, _T("Buddy"), 5))
		{
			if (adding)
			{
				auto option_value = option + 5;
				TCHAR which_buddy = *option_value;
				if (which_buddy) // i.e. it's not the zero terminator
				{
					GuiIndexType u = FindControl(option_value + 1);
					if (u < mControlCount)
					{
						if (which_buddy == '1')
							aOpt.buddy1 = mControl[u];
						else // assume '2'
							aOpt.buddy2 = mControl[u];
					}
				}
			}
			//else removal not supported.
		}

		// Progress and Slider
		else if (!_tcsicmp(option, _T("Vertical")))
		{
			// Seems best not to recognize Vertical for Tab controls since Left and Right
			// already cover it very well.
			if (aControl.type == GUI_CONTROL_SLIDER)
				if (adding) aOpt.style_add |= TBS_VERT; else aOpt.style_remove |= TBS_VERT;
			else if (aControl.type == GUI_CONTROL_PROGRESS)
				if (adding) aOpt.style_add |= PBS_VERTICAL; else aOpt.style_remove |= PBS_VERTICAL;
			//else do nothing, not a supported type
		}
		else if (!_tcsnicmp(option, _T("Range"), 5)) // Caller should ignore aOpt.range_min/max if it isn't applicable for this control type.
		{
			if (adding)
			{
				auto option_value = option + 5; // Helps with omitting the first minus sign, if any, below.
				if (*option_value) // Prevent reading beyond the zero terminator due to option_value+1 in some places below.
				{
					if (aControl.type == GUI_CONTROL_DATETIME || aControl.type == GUI_CONTROL_MONTHCAL)
					{
						// Note: aOpt.range_changed is not set for these control types. aOpt.gdtr_range is used instead.
						aOpt.gdtr_range = YYYYMMDDToSystemTime2(option_value, aOpt.sys_time_range);
						if (!aOpt.gdtr_range) // Must be invalid, since option_value was verified to be non-empty.
						{
							error_message = ERR_INVALID_OPTION;
							goto return_error;
						}
						if (aControl.hwnd) // Caller relies on us doing it now.
						{
							SendMessage(aControl.hwnd // MCM_SETRANGE != DTM_SETRANGE
								, aControl.type == GUI_CONTROL_DATETIME ? DTM_SETRANGE : MCM_SETRANGE
								, aOpt.gdtr_range, (LPARAM)aOpt.sys_time_range);
							// Unlike GUI_CONTROL_DATETIME, GUI_CONTROL_MONTHCAL doesn't visibly update selection
							// when new range is applied, not even with InvalidateRect().  In fact, it doesn't
							// internally update either.  This might be worked around by getting the selected
							// date (or range of dates) and reapplying them, but the need is rare enough that
							// code size reduction seems more important.
						}
					}
					else // Control types other than datetime/monthcal.
					{
						aOpt.range_changed = true;
						aOpt.range_min = ATOI(option_value);
						if (auto cp = _tcschr(option_value + 1, '-')) // +1 to omit the min's minus sign, if it has one.
							aOpt.range_max = ATOI(cp + 1);
					}
				}
				//else the Range word is present but has nothing after it.  Ignore it.
			}
			//else removing.  Do nothing (not currently implemented).
		}

		// Progress
		else if (aControl.type == GUI_CONTROL_PROGRESS && !_tcsicmp(option, _T("Smooth")))
			if (adding) aOpt.style_add |= PBS_SMOOTH; else aOpt.style_remove |= PBS_SMOOTH;

		// Tab control
		else if (!_tcsicmp(option, _T("Buttons")))
		{
			if (aControl.type == GUI_CONTROL_TAB)
				if (adding) aOpt.style_add |= TCS_BUTTONS; else aOpt.style_remove |= TCS_BUTTONS;
			else if (aControl.type == GUI_CONTROL_TREEVIEW)
				if (adding) aOpt.style_add |= TVS_HASBUTTONS; else aOpt.style_remove |= TVS_HASBUTTONS;
		}
		else if (aControl.type == GUI_CONTROL_TAB && !_tcsicmp(option, _T("Bottom")))
			if (adding)
			{
				aOpt.style_add |= TCS_BOTTOM;
				aOpt.style_remove |= TCS_VERTICAL;
			}
			else
				aOpt.style_remove |= TCS_BOTTOM;

		else if (aControl.type == GUI_CONTROL_CUSTOM && !_tcsnicmp(option, _T("Class"), 5))
		{
			WNDCLASSEX wc;
			// Retrieve the class atom (http://blogs.msdn.com/b/oldnewthing/archive/2004/10/11/240744.aspx)
			aOpt.customClassAtom = (ATOM) GetClassInfoEx(g_hInstance, option + 5, &wc);
			if (aOpt.customClassAtom == 0)
			{
				error_message = _T("Unrecognized window class.");
				goto return_error;
			}
		}

		// Styles (alignment/justification):
		else if (!_tcsicmp(option, _T("Center")))
			if (adding)
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_SLIDER:
					if (adding) aOpt.style_add |= TBS_BOTH;
					aOpt.style_remove |= TBS_LEFT;
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_add |= SS_CENTER;
					aOpt.style_remove |= (SS_TYPEMASK & ~SS_CENTER); // i.e. Zero out all type-bits except SS_CENTER's bit.
					break;
				case GUI_CONTROL_GROUPBOX: // Changes alignment of its label.
				case GUI_CONTROL_BUTTON:   // Probably has no effect in this case, since it's centered by default?
				case GUI_CONTROL_CHECKBOX: // Puts gap between box and label.
				case GUI_CONTROL_RADIO:
					aOpt.style_add |= BS_CENTER;
					// But don't remove BS_LEFT or BS_RIGHT since BS_CENTER is defined as a combination of them.
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_add |= ES_CENTER;
					aOpt.style_remove |= ES_RIGHT; // Mutually exclusive since together they are (probably) invalid.
					break;
				// Not applicable for:
				//case GUI_CONTROL_PIC: SS_CENTERIMAGE is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				//case GUI_CONTROL_LISTVIEW:
				//case GUI_CONTROL_TREEVIEW:
				//case GUI_CONTROL_UPDOWN:
				//case GUI_CONTROL_DATETIME:
				//case GUI_CONTROL_MONTHCAL:
				//case GUI_CONTROL_LINK:
				}
			}
			else // Removing.
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_SLIDER:
					aOpt.style_remove |= TBS_BOTH;
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_remove |= SS_TYPEMASK; // Revert to SS_LEFT because there's no way of knowing what the intended or previous value was.
					break;
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_RADIO:
					// BS_CENTER is a tricky one since it is a combination of BS_LEFT and BS_RIGHT.
					// If the control exists and has either BS_LEFT or BS_RIGHT (but not both), do
					// nothing:
					if (aControl.hwnd)
					{
						// v1.0.44.08: Fixed the following by adding "== BS_CENTER":
						if ((GetWindowLong(aControl.hwnd, GWL_STYLE) & BS_CENTER) == BS_CENTER) // i.e. it has both BS_LEFT and BS_RIGHT
							aOpt.style_remove |= BS_CENTER;
						//else nothing needs to be done.
					}
					else // v1.0.44.08: Fixed the following by adding "== BS_CENTER":
						if ((aOpt.style_add & BS_CENTER) == BS_CENTER) // i.e. Both BS_LEFT and BS_RIGHT are set to be added.
							aOpt.style_add &= ~BS_CENTER; // Undo it, which later helps avoid the need to apply style_add prior to style_remove.
						//else nothing needs to be done.
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_remove |= ES_CENTER;
					break;
				// Not applicable for:
				//case GUI_CONTROL_PIC: SS_CENTERIMAGE is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				//case GUI_CONTROL_LISTVIEW:
				//case GUI_CONTROL_TREEVIEW:
				//case GUI_CONTROL_UPDOWN:
				//case GUI_CONTROL_DATETIME:
				//case GUI_CONTROL_MONTHCAL:
				//case GUI_CONTROL_LINK:
				}
			}

		else if (!_tcsicmp(option, _T("Right")))
			if (adding)
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_UPDOWN:
					aOpt.style_remove |= UDS_ALIGNLEFT;
					aOpt.style_add |= UDS_ALIGNRIGHT;
					break;
				case GUI_CONTROL_DATETIME:
					aOpt.style_add |= DTS_RIGHTALIGN;
					break;
				case GUI_CONTROL_SLIDER:
					aOpt.style_remove |= TBS_LEFT|TBS_BOTH;
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_add |= SS_RIGHT;
					aOpt.style_remove |= (SS_TYPEMASK & ~SS_RIGHT); // i.e. Zero out all type-bits except SS_RIGHT's bit.
					break;
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_RADIO:
					aOpt.style_add |= BS_RIGHT;
					// Doing this indirectly removes BS_CENTER, and does it in a way that makes unimportant
					// the order in which style_add and style_remove are applied later:
					aOpt.style_remove |= BS_LEFT;
					// And by default, put button itself to the right of its label since that seems
					// likely to be far more common/desirable (there can be a more obscure option
					// later to change this default):
					if (aControl.type == GUI_CONTROL_CHECKBOX || aControl.type == GUI_CONTROL_RADIO)
						aOpt.style_add |= BS_RIGHTBUTTON;
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_add |= ES_RIGHT;
					aOpt.style_remove |= ES_CENTER; // Mutually exclusive since together they are (probably) invalid.
					break;
				case GUI_CONTROL_TAB:
					aOpt.style_add |= TCS_VERTICAL|TCS_MULTILINE|TCS_RIGHT;
					break;
				case GUI_CONTROL_LINK:
					aOpt.style_add |= LWS_RIGHT;
					break;
				// Not applicable for:
				//case GUI_CONTROL_MONTHCAL:
				//case GUI_CONTROL_PIC: SS_RIGHTJUST is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				//case GUI_CONTROL_LISTVIEW:
				//case GUI_CONTROL_TREEVIEW:
				}
			}
			else // Removing.
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_DATETIME:
					aOpt.style_remove |= DTS_RIGHTALIGN;
					break;
				case GUI_CONTROL_SLIDER:
					aOpt.style_add |= TBS_LEFT;
					aOpt.style_remove |= TBS_BOTH; // Debatable.
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_remove |= SS_TYPEMASK; // Revert to SS_LEFT because there's no way of knowing what the intended or previous value was.
					break;
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_RADIO:
					// BS_RIGHT is a tricky one since it is included inside BS_CENTER.
					// Thus, if the control exists and has BS_CENTER, do nothing since
					// BS_RIGHT can't be in effect if BS_CENTER already is:
					if (aControl.hwnd)
					{
						if ((GetWindowLong(aControl.hwnd, GWL_STYLE) & BS_CENTER) != BS_CENTER) // v1.0.44.08: Fixed by adding "!= BS_CENTER".
							aOpt.style_remove |= BS_RIGHT;
					}
					else
						if ((aOpt.style_add & BS_CENTER) != BS_CENTER) // v1.0.44.08: Fixed by adding "!= BS_CENTER".
							aOpt.style_add &= ~BS_RIGHT;  // A little strange, but seems correct since control hasn't even been created yet.
						//else nothing needs to be done because BS_RIGHT is already in effect removed since
						//BS_CENTER makes BS_RIGHT impossible to manifest.
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_remove |= ES_RIGHT;
					break;
				case GUI_CONTROL_TAB:
					aOpt.style_remove |= TCS_VERTICAL|TCS_RIGHT;
					break;
				// Not applicable for:
				//case GUI_CONTROL_UPDOWN: Removing "right" doesn't make much sense, so only adding "left" is supported.
				//case GUI_CONTROL_MONTHCAL:
				//case GUI_CONTROL_PIC: SS_RIGHTJUST is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				//case GUI_CONTROL_LISTVIEW:
				//case GUI_CONTROL_TREEVIEW:
				//case GUI_CONTROL_LINK:
				}
			}

		else if (!_tcsicmp(option, _T("Left")))
		{
			if (adding)
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_UPDOWN:
					aOpt.style_remove |= UDS_ALIGNRIGHT;
					aOpt.style_add |= UDS_ALIGNLEFT;
					break;
				case GUI_CONTROL_DATETIME:
					aOpt.style_remove |= DTS_RIGHTALIGN;
					break;
				case GUI_CONTROL_SLIDER:
					aOpt.style_add |= TBS_LEFT;
					aOpt.style_remove |= TBS_BOTH;
					break;
				case GUI_CONTROL_TEXT:
					aOpt.style_remove |= SS_TYPEMASK; // i.e. Zero out all type-bits to expose the default of 0, which is SS_LEFT.
					break;
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_RADIO:
					aOpt.style_add |= BS_LEFT;
					// Doing this indirectly removes BS_CENTER, and does it in a way that makes unimportant
					// the order in which style_add and style_remove are applied later:
					aOpt.style_remove |= BS_RIGHT;
					// And by default, put button itself to the left of its label since that seems
					// likely to be far more common/desirable (there can be a more obscure option
					// later to change this default):
					if (aControl.type == GUI_CONTROL_CHECKBOX || aControl.type == GUI_CONTROL_RADIO)
						aOpt.style_remove |= BS_RIGHTBUTTON;
					break;
				case GUI_CONTROL_EDIT:
					aOpt.style_remove |= ES_RIGHT|ES_CENTER;  // Removing these exposes the default of 0, which is LEFT.
					break;
				case GUI_CONTROL_TAB:
					aOpt.style_add |= TCS_VERTICAL|TCS_MULTILINE;
					aOpt.style_remove |= TCS_RIGHT;
					break;
				// Not applicable for:
				//case GUI_CONTROL_MONTHCAL:
				//case GUI_CONTROL_PIC: SS_CENTERIMAGE is currently not used due to auto-pic-scaling/fitting.
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				//case GUI_CONTROL_LISTVIEW:
				//case GUI_CONTROL_TREEVIEW:
				//case GUI_CONTROL_LINK:
				}
			}
			else // Removing.
			{
				switch (aControl.type)
				{
				case GUI_CONTROL_SLIDER:
					aOpt.style_remove |= TBS_LEFT|TBS_BOTH; // Removing TBS_BOTH is debatable, but "-left" is pretty rare/obscure anyway.
					break;
				case GUI_CONTROL_GROUPBOX:
				case GUI_CONTROL_BUTTON:
				case GUI_CONTROL_CHECKBOX:
				case GUI_CONTROL_RADIO:
					// BS_LEFT is a tricky one since it is included inside BS_CENTER.
					// Thus, if the control exists and has BS_CENTER, do nothing since
					// BS_LEFT can't be in effect if BS_CENTER already is:
					if (aControl.hwnd)
					{
						if ((GetWindowLong(aControl.hwnd, GWL_STYLE) & BS_CENTER) != BS_CENTER) // v1.0.44.08: Fixed by adding "!= BS_CENTER".
							aOpt.style_remove |= BS_LEFT;
					}
					else
						if ((aOpt.style_add & BS_CENTER) != BS_CENTER) // v1.0.44.08: Fixed by adding "!= BS_CENTER".
							aOpt.style_add &= ~BS_LEFT;  // A little strange, but seems correct since control hasn't even been created yet.
						//else nothing needs to be done because BS_LEFT is already in effect removed since
						//BS_CENTER makes BS_LEFT impossible to manifest.
					break;
				case GUI_CONTROL_TAB:
					aOpt.style_remove |= TCS_VERTICAL;
					break;
				// Not applicable for these since their LEFT attributes are zero and thus cannot be removed:
				//case GUI_CONTROL_UPDOWN: Removing "left" doesn't make much sense, so only adding "right" is supported.
				//case GUI_CONTROL_DATETIME: Removing "left" is not supported since it seems counterintuitive and too rarely needed.
				//case GUI_CONTROL_MONTHCAL:
				//case GUI_CONTROL_TEXT:
				//case GUI_CONTROL_PIC:
				//case GUI_CONTROL_DROPDOWNLIST:
				//case GUI_CONTROL_COMBOBOX:
				//case GUI_CONTROL_LISTBOX:
				//case GUI_CONTROL_LISTVIEW:
				//case GUI_CONTROL_TREEVIEW:
				//case GUI_CONTROL_EDIT:
				//case GUI_CONTROL_LINK:
				}
			}
		} // else if

		else
		{
			// THE BELOW SHOULD BE DONE LAST so that they don't steal phrases/words that should be detected
			// as option words above.  An existing example is H for Hidden (above) or Height (below).
			// Additional examples:
			// if "visible" and "resize" ever become valid option words, the below would otherwise wrongly
			// detect them as variable=isible and row_count=esize, respectively.

			if (ParsePositiveInteger(option, option_dword)) // Above has already verified that *option can't be whitespace.
			{
				// Pure numbers are assumed to be style additions or removals:
				if (adding) aOpt.style_add |= option_dword; else aOpt.style_remove |= option_dword;
				continue;
			}

			TCHAR option_char = ctoupper(*option);
			auto option_value = option + 1; // Above has already verified that option isn't the empty string.

			if (!*option_value)
			{
				// The option word consists of only one character, so consider it valid only if it doesn't
				// require an arg.  Example: An isolated "H" should not cause the height to be set to zero.
				switch (option_char)
				{
				case 'C':
					if (!adding)
						aOpt.color = CLR_DEFAULT; // i.e. use system default even if mCurrentColor is non-default.
					break;
				case 'V':
					ControlSetName(aControl, NULL);
					break;
				default:
					// Anything else is invalid.
					error_message = ERR_INVALID_OPTION;
					goto return_error;
				}
				continue;
			}
			// Since above didn't "continue", there is text after the option letter, so take action accordingly.

			int option_int = 0; // Only valid for [XYWHTER].
			float option_float; // Only valid for R.
			TCHAR option_char2; // Only valid for [XYWH].
			bool use_margin_offset; // Only valid for [XY].

			if (_tcschr(_T("XYWHTER"), option_char))
			{
				option_char2 = ctoupper(*option_value);
				switch (option_char)
				{
				case 'W':
				case 'H':
					if (option_char2 == 'P')
						++option_value;
					break;
				case 'X':
				case 'Y':
					if (use_margin_offset = (*option_value == '+' && 'M' == ctoupper(option_value[1]))) // x+m or y+m.
						option_value += 2;
					else if (_tcschr(_T("MPS+"), option_char2)) // Any other non-digit char should be picked up as an error via IsNumeric().
						++option_value;							// Also skips over the '+' to allow, eg, x+-n
					break;
				}
				// Allow both integer and floating-point for any numeric options.
				switch (IsNumeric(option_value, TRUE, FALSE, TRUE))
				{
				case PURE_INTEGER:
					option_float = float(option_int = ATOI(option_value)); // 'E' depends on this supporting the full range of DWORD.
					break;
				case PURE_FLOAT:
					option_int = int(option_float = (float)_tstof(option_value)); // _tstof() vs. ATOF() because PURE_FLOAT is never hexadecimal.
					break;
				default:
					if (*option_value || !option_char2) // i.e. allow blank option_value if option_char2 was recognized above.
						option_char = 0; // Mark it as invalid for switch() below.
				}
				if ((option_char == 'X' || option_char == 'Y')
					|| (option_char == 'W' || option_char == 'H') && (option_int != -1 || option_char2 == 'P')) // Scale W/H unless it's W-1 or H-1.
					option_int = Scale(option_int);
			}

			switch (option_char)
			{
			case 'T': // Tabstop (the kind that exists inside a multi-line edit control or ListBox).
				if (aOpt.tabstop_count < GUI_MAX_TABSTOPS)
					aOpt.tabstop[aOpt.tabstop_count++] = (UINT)option_int;
				//else ignore ones beyond the maximum.
				break;

			case 'V': // Name (originally: Variable)
				switch (ControlSetName(aControl, option_value))
				{
				case FR_FAIL:
					return FAIL;
				case FR_E_OUTOFMEM:
					return MemoryError();
				}
				break;

			case 'C':  // Color
				if (!ColorToBGR(option_value, aOpt.color))
				{
					error_message = ERR_INVALID_OPTION;
					goto return_error;
				}
				break;

			case 'W':
				if (option_char2 == 'P') // Use the previous control's value.
					aOpt.width = mPrevWidth + option_int;
				else
					aOpt.width = option_int;
				break;

			case 'H':
				if (option_char2 == 'P') // Use the previous control's value.
					aOpt.height = mPrevHeight + option_int;
				else
					aOpt.height = option_int;
				break;

			case 'X':
				if (option_char2 == '+')
				{
					int offset = (use_margin_offset ? mMarginX : 0) + option_int;
					if (tab_control = FindTabControl(aControl.tab_control_index)) // Assign.
					{
						// Since this control belongs to a tab control and that tab control already exists,
						// Position it relative to the tab control's client area upper-left corner if this
						// is the first control on this particular tab/page:
						if (!GetControlCountOnTabPage(aControl.tab_control_index, aControl.tab_index))
						{
							pt = GetPositionOfTabDisplayArea(*tab_control);
							aOpt.x = pt.x + offset;
							if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
								aOpt.y = pt.y + mMarginY;
							break;
						}
						// else fall through and do it the standard way.
					}
					// Since above didn't break, do it the standard way.
					aOpt.x = mPrevX + mPrevWidth + offset;
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = mPrevY;  // Since moving in the X direction, retain the same Y as previous control.
				}
				// For the M and P sub-options, note that the +/- prefix is optional.  The number is simply
				// read in as-is (though the use of + is more self-documenting in this case than omitting
				// the sign entirely).
				else if (option_char2 == 'M') // Use the X margin
				{
					aOpt.x = mMarginX + option_int;
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = mMaxExtentDown + mMarginY;
				}
				else if (option_char2 == 'P') // Use the previous control's X position.
				{
					aOpt.x = mPrevX + option_int;
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = aOpt.x != mPrevX ? mPrevY  // Since moving in the X direction, retain the same Y as previous control.
							: mPrevY + mPrevHeight + mMarginY; // Same X, so default to below the previous control.
				}
				else if (option_char2 == 'S') // Use the saved X position
				{
					aOpt.x = mSectionX + option_int;
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = mMaxExtentDownSection + mMarginY;  // In this case, mMarginY is the padding between controls.
				}
				else
				{
					aOpt.x = option_int;
					if (aOpt.y == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.y = mMaxExtentDown + mMarginY;
				}
				break;

			case 'Y':
				if (option_char2 == '+')
				{
					int offset = (use_margin_offset ? mMarginY : 0) + option_int;
					if (tab_control = FindTabControl(aControl.tab_control_index)) // Assign.
					{
						// Since this control belongs to a tab control and that tab control already exists,
						// Position it relative to the tab control's client area upper-left corner if this
						// is the first control on this particular tab/page:
						if (!GetControlCountOnTabPage(aControl.tab_control_index, aControl.tab_index))
						{
							pt = GetPositionOfTabDisplayArea(*tab_control);
							aOpt.y = pt.y + offset;
							if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
								aOpt.x = pt.x + mMarginX;
							break;
						}
						// else fall through and do it the standard way.
					}
					// Since above didn't break, do it the standard way.
					aOpt.y = mPrevY + mPrevHeight + offset;
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = mPrevX;  // Since moving in the Y direction, retain the same X as previous control.
				}
				// For the M and P sub-options, not that the +/- prefix is optional.  The number is simply
				// read in as-is (though the use of + is more self-documenting in this case than omitting
				// the sign entirely).
				else if (option_char2 == 'M') // Use the Y margin
				{
					aOpt.y = mMarginY + option_int;
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = mMaxExtentRight + mMarginX;
				}
				else if (option_char2 == 'P') // Use the previous control's Y position.
				{
					aOpt.y = mPrevY + option_int;
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = aOpt.y != mPrevY ? mPrevX  // Since moving in the Y direction, retain the same X as previous control.
							: mPrevX + mPrevWidth + mMarginY; // Same Y, so default to below the previous control.
				}
				else if (option_char2 == 'S') // Use the saved Y position
				{
					aOpt.y = mSectionY + option_int;
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = mMaxExtentRightSection + mMarginX; // In this case, mMarginX is the padding between controls.
				}
				else
				{
					aOpt.y = option_int;
					if (aOpt.x == COORD_UNSPECIFIED) // Not yet explicitly set, so use default.
						aOpt.x = mMaxExtentRight + mMarginX;
				}
				break;

			case 'R': // The number of rows desired in the control (can be fractional).
				aOpt.row_count = option_float;
				break;

			case 'E': // Extended style additions or removals.
				if (adding)
					aOpt.exstyle_add |= (DWORD)option_int;
				else
					aOpt.exstyle_remove |= (DWORD)option_int;
				break;

			default:
				// Anything not already handled above is invalid.
				error_message = ERR_INVALID_OPTION;
				goto return_error;
			} // switch()
		} // Final "else" in the "else if" ladder.

		continue;
	// This "subroutine" is used to reduce code size (needs to be re-verified):
	return_error:
		if (!ValueError(error_message, option, FAIL_OR_OK))
			return FAIL;
		// Otherwise, user wants to continue.
	} // for() each item in option list
	
	// If the control has already been created, apply the new style and exstyle here, if any:
	if (aControl.hwnd)
	{
		DWORD current_style = GetWindowLong(aControl.hwnd, GWL_STYLE);
		DWORD new_style = (current_style | aOpt.style_add) & ~aOpt.style_remove; // Some things such as GUI_CONTROL_TEXT+SS_TYPEMASK might rely on style_remove being applied *after* style_add.
		UINT current_value_uint; // Currently only used for Progress controls.

		// Fix for v1.0.24:
		// Certain styles can't be applied with a simple bit-or.  The below section is a subset of
		// a similar section in AddControl() to make sure that these styles are properly handled:
		switch (aControl.type)
		{
		case GUI_CONTROL_PIC:
			// Fixed for v1.0.25.11 to prevent SS_ICON from getting changed to SS_BITMAP:
			new_style = (new_style & ~0x0F) | (current_style & 0x0F); // Done to ensure the lowest four bits are pure.
			break;
		case GUI_CONTROL_GROUPBOX:
			// There doesn't seem to be any flexibility lost by forcing the buttons to be the right type,
			// and doing so improves maintainability and peace-of-mind:
			new_style = (new_style & ~BS_TYPEMASK) | BS_GROUPBOX;  // Force it to be the right type of button.
			break;
		case GUI_CONTROL_BUTTON:
			if (new_style & BS_DEFPUSHBUTTON)
				new_style = (new_style & ~BS_TYPEMASK) | BS_DEFPUSHBUTTON; // Done to ensure the lowest four bits are pure.
			else
				new_style &= ~BS_TYPEMASK;  // Force it to be the right type of button --> BS_PUSHBUTTON == 0
			// Fixed for v1.0.33.01:
			// The following must be done here rather than later because it's possible for the
			// button to lack the BS_DEFPUSHBUTTON style even when it is the default button (such as when
			// keyboard focus is on some other button).  Consequently, no difference in style might be
			// detected further below, which is why it's done here:
			if (aControlIndex == mDefaultButtonIndex)
			{
				if (aOpt.style_remove & BS_DEFPUSHBUTTON)
				{
					// Remove the default button (rarely needed so that's why there is current no
					// "Gui, NoDefaultButton" command:
					mDefaultButtonIndex = -1;
					// This will alter the control id received via WM_COMMAND when the user presses ENTER:
					SendMessage(mHwnd, DM_SETDEFID, (WPARAM)IDOK, 0); // restore to default
					// Sometimes button visually has the default style even when GetWindowLong() says
					// it doesn't.  Given how rare it is to not have a default button at all after having
					// one, rather than try to analyze exactly what circumstances this happens in,
					// unconditionally reset the style and redraw by indicating that current style is
					// different from the new style:
					current_style |= BS_DEFPUSHBUTTON;
				}
			}
			else // This button isn't the default button yet, but is it becoming it?
			{
				// Remember that the default button doesn't always have the BS_DEFPUSHBUTTON, namely at
				// times when it shouldn't visually appear to be the default, such as when keyboard focus
				// is on some other button.  Therefore, don't rely on new_style or current_style's
				// having or not having BS_DEFPUSHBUTTON as being a correct indicator.
				if (aOpt.style_add & BS_DEFPUSHBUTTON)
				{
					// First remove the style from the old default button, if there is one:
                    if (mDefaultButtonIndex < mControlCount)
						SendMessage(mControl[mDefaultButtonIndex]->hwnd, BM_SETSTYLE
							, (WPARAM)LOWORD((GetWindowLong(mControl[mDefaultButtonIndex]->hwnd, GWL_STYLE) & ~BS_DEFPUSHBUTTON))
							, MAKELPARAM(TRUE, 0)); // Redraw = yes. It's probably smart enough not to do it if the window is hidden.
					mDefaultButtonIndex = aControlIndex;
					// This will alter the control id received via WM_COMMAND when the user presses ENTER:
					SendMessage(mHwnd, DM_SETDEFID, (WPARAM)GUI_INDEX_TO_ID(mDefaultButtonIndex), 0);
					// This button's visual/default appearance will be updated further below, if warranted.
				}
			}
			break;
		case GUI_CONTROL_CHECKBOX:
			// Note: BS_AUTO3STATE and BS_AUTOCHECKBOX are mutually exclusive due to their overlap within
			// the bit field:
			if ((new_style & BS_AUTO3STATE) == BS_AUTO3STATE) // Fixed for v1.0.45.03 to check if all the BS_AUTO3STATE bits are present, not just "any" of them. BS_TYPEMASK is not involved here because this is a purity check, and TYPEMASK would defeat the whole purpose.
				new_style = (new_style & ~BS_TYPEMASK) | BS_AUTO3STATE; // Done to ensure the lowest four bits are pure.
			else
				new_style = (new_style & ~BS_TYPEMASK) | BS_AUTOCHECKBOX;  // Force it to be the right type of button.
			break;
		case GUI_CONTROL_RADIO:
			new_style = (new_style & ~BS_TYPEMASK) | BS_AUTORADIOBUTTON;  // Force it to be the right type of button.
			break;
		case GUI_CONTROL_DROPDOWNLIST:
			new_style |= CBS_DROPDOWNLIST;  // This works because CBS_DROPDOWNLIST == CBS_SIMPLE|CBS_DROPDOWN
			break;
		case GUI_CONTROL_COMBOBOX:
			if (new_style & CBS_SIMPLE) // i.e. CBS_SIMPLE has been added to the original default, so assume it is SIMPLE.
				new_style = (new_style & ~0x0F) | CBS_SIMPLE; // Done to ensure the lowest four bits are pure.
			else
				new_style = (new_style & ~0x0F) | CBS_DROPDOWN; // Done to ensure the lowest four bits are pure.
			break;
		case GUI_CONTROL_LISTVIEW: // Being in the switch serves to verify control's type because it isn't verified in places where listview_view is set.
			if (aOpt.listview_view != -1) // A new view was explicitly specified.
			{
				// Fix for v1.0.36.04:
				// For XP, must always use ListView_SetView() because otherwise switching from Tile view back to
				// the *same* view that was in effect prior to tile view wouldn't work (since the the control's
				// style LVS_TYPEMASK bits would not have changed).  This is because LV_VIEW_TILE is a special
				// view that cannot be set via style change.
				ListView_SetView(aControl.hwnd, aOpt.listview_view);
				// Regardless of whether SetView was called above, adjust the style too so that the upcoming
				// style change won't undo what was just done above:
				if (aOpt.listview_view != LV_VIEW_TILE) // It was ensured earlier that listview_view can be set to LV_VIEW_TILE only for XP or later.
					new_style = (new_style & ~LVS_TYPEMASK) | aOpt.listview_view;
			}
			break;
		// Nothing extra for these currently:
		//case GUI_CONTROL_LISTBOX: i.e. allow LBS_NOTIFY to be removed in case anyone really wants to do that.
		//case GUI_CONTROL_TREEVIEW:
		//case GUI_CONTROL_EDIT:
		//case GUI_CONTROL_TEXT:  Ensuring SS_BITMAP and such are absent seems too over-protective.
		//case GUI_CONTROL_LINK:
		//case GUI_CONTROL_DATETIME:
		//case GUI_CONTROL_MONTHCAL:
		//case GUI_CONTROL_HOTKEY:
		//case GUI_CONTROL_UPDOWN:
		//case GUI_CONTROL_SLIDER:
		//case GUI_CONTROL_PROGRESS:
		//case GUI_CONTROL_TAB: i.e. allow WS_CLIPSIBLINGS to be removed (Rajat needs this) and also TCS_OWNERDRAWFIXED in case anyone really wants to.
		}

		// This needs to be done prior to applying the updated style since it sometimes adds
		// more style attributes:
		if (aOpt.limit) // A char length-limit was specified or de-specified for an edit/combo field.
		{
			// These styles are applied last so that multiline vs. singleline will already be resolved
			// and known, since all options have now been processed.
			if (aControl.type == GUI_CONTROL_EDIT)
			{
				// For the below, note that EM_LIMITTEXT == EM_SETLIMITTEXT.
				if (aOpt.limit < 0)
				{
					// Either limit to visible width of field, or remove existing limit.
					// In both cases, first remove the control's internal limit in case it
					// was something really small before:
					SendMessage(aControl.hwnd, EM_LIMITTEXT, 0, 0);
					// Limit > INT_MIN but less than zero is the signal to limit input length to visible
					// width of field. But it can only work if the edit isn't a multiline.
					if (aOpt.limit != INT_MIN && !(new_style & ES_MULTILINE))
						new_style &= ~(WS_HSCROLL|ES_AUTOHSCROLL); // Enable the limit-to-visible-width style.
				}
				else // greater than zero, since zero itself it checked in one of the enclosing IFs above.
					SendMessage(aControl.hwnd, EM_LIMITTEXT, aOpt.limit, 0); // Set a hard limit.
			}
			else if (aControl.type == GUI_CONTROL_HOTKEY)
			{
				if (aOpt.limit < 0) // This is the signal to remove any existing limit.
					SendMessage(aControl.hwnd, HKM_SETRULES, 0, 0);
				else // greater than zero, since zero itself it checked in one of the enclosing IFs above.
					SendMessage(aControl.hwnd, HKM_SETRULES, aOpt.limit, MAKELPARAM(HOTKEYF_CONTROL|HOTKEYF_ALT, 0));
					// Above must also specify Ctrl+Alt or some other default, otherwise the restriction will have
					// no effect.
			}
			// Altering the limit after the control exists appears to be ineffective, so this is commented out:
			//else if (aControl.type == GUI_CONTROL_COMBOBOX)
			//	// It has been verified that that EM_LIMITTEXT has no effect when sent directly
			//	// to a ComboBox hwnd; however, it might work if sent to its edit-child.
			//	// For now, a Combobox can only be limited to its visible width.  Later, there might
			//	// be a way to send a message to its child control to limit its width directly.
			//	if (aOpt.limit == INT_MIN) // remove existing limit
			//		new_style |= CBS_AUTOHSCROLL;
			//	else
			//		new_style &= ~CBS_AUTOHSCROLL;
			//	// i.e. SetWindowLong() cannot manifest the above style after the window exists.
		}

		bool style_change_ok;
		bool style_needed_changing = false; // Set default.

		if (current_style != new_style)
		{
			style_needed_changing = true; // Either style or exstyle is changing.
			style_change_ok = false; // Starting assumption.
			switch (aControl.type)
			{
			case GUI_CONTROL_BUTTON:
				// For some reason, the following must be done *after* removing the visual/default style from
				// the old default button and/or after doing DM_SETDEFID above.
				// BM_SETSTYLE is much more likely to have an effect for buttons than SetWindowLong().
				// Redraw = yes, though it seems to be ineffective sometimes. It's probably smart enough not to do
				// it if the window is hidden.
				// Fixed for v1.0.33.01: The HWND should be aControl.hwnd not mControl[mDefaultButtonIndex].hwnd
				SendMessage(aControl.hwnd, BM_SETSTYLE, (WPARAM)LOWORD(new_style), MAKELPARAM(TRUE, 0));
				break;
			case GUI_CONTROL_LISTBOX:
				if ((new_style & WS_HSCROLL) && !(current_style & WS_HSCROLL)) // Scroll bar being added.
				{
					if (aOpt.hscroll_pixels < 0) // Calculate a default based on control's width.
					{
						// Since horizontal scrollbar is relatively rarely used, no fancy method
						// such as calculating scrolling-width via LB_GETTEXTLEN and current font's
						// average width is used.
						GetWindowRect(aControl.hwnd, &rect);
						aOpt.hscroll_pixels = 3 * (rect.right - rect.left);
					}
					// If hscroll_pixels is now zero or smaller than the width of the control,
					// the scrollbar will not be shown:
					SendMessage(aControl.hwnd, LB_SETHORIZONTALEXTENT, (WPARAM)aOpt.hscroll_pixels, 0);
				}
				else if (!(new_style & WS_HSCROLL) && (current_style & WS_HSCROLL)) // Scroll bar being removed.
					SendMessage(aControl.hwnd, LB_SETHORIZONTALEXTENT, 0, 0);
				break;
			case GUI_CONTROL_PROGRESS:
				// SetWindowLong with GWL_STYLE appears to reset the position,
				// so save the current position here to be restored later:
				current_value_uint = (UINT)SendMessage(aControl.hwnd, PBM_GETPOS, 0, 0);
				break;
			} // switch()
			SetLastError(0); // Prior to SetWindowLong(), as recommended by MSDN.
			// Call this even for buttons because BM_SETSTYLE only handles the LOWORD part of the style:
			if (SetWindowLong(aControl.hwnd, GWL_STYLE, new_style) || !GetLastError()) // This is the precise way to detect success according to MSDN.
			{
				// Even if it indicated success, sometimes it failed anyway.  Find out for sure:
				if (GetWindowLong(aControl.hwnd, GWL_STYLE) != current_style)
					style_change_ok = true; // Even a partial change counts as a success.
			}
		}

		DWORD current_exstyle = GetWindowLong(aControl.hwnd, GWL_EXSTYLE);
		DWORD new_exstyle = (current_exstyle | aOpt.exstyle_add) & ~aOpt.exstyle_remove;
		if (current_exstyle != new_exstyle)
		{
			if (!style_needed_changing)
			{
				style_needed_changing = true; // Either style or exstyle is changing.
				style_change_ok = false; // Starting assumption.
			}
			//ELSE don't change the value of style_change_ok because we want it to retain the value set by
			// the GWL_STYLE change above; i.e. a partial success on either style or exstyle counts as a full
			// success.
			SetLastError(0); // Prior to SetWindowLong(), as recommended by MSDN.
			if (SetWindowLong(aControl.hwnd, GWL_EXSTYLE, new_exstyle) || !GetLastError()) // This is the precise way to detect success according to MSDN.
			{
				// Even if it indicated success, sometimes it failed anyway.  Find out for sure:
				if (GetWindowLong(aControl.hwnd, GWL_EXSTYLE) != current_exstyle)
					style_change_ok = true; // Even a partial change counts as a success.
			}
		}

		if (aControl.type == GUI_CONTROL_LISTVIEW)
		{
			// These ListView "extended styles" exist entirely separate from normal extended styles.
			// In other words, a ListView may have three types of styles: Normal, Extended, and its
			// own set of LV Extended styles.
			DWORD current_lv_style = ListView_GetExtendedListViewStyle(aControl.hwnd);
			if (current_lv_style != aOpt.listview_style)
			{
				if (!style_needed_changing)
				{
					style_needed_changing = true; // Either style, exstyle, or lv_exstyle is changing.
					style_change_ok = false; // Starting assumption.
				}
				//ELSE don't change the value of style_change_ok because we want it to retain the value set by
				// the GWL_STYLE or EXSTYLE change above; i.e. a partial success on either style or exstyle
				// counts as a full success.
				ListView_SetExtendedListViewStyle(aControl.hwnd, aOpt.listview_style); // Has no return value.
				if (ListView_GetExtendedListViewStyle(aControl.hwnd) != current_lv_style)
					style_change_ok = true; // Even a partial change counts as a success.
			}
		}

  		// Redrawing the controls is required in some cases, such as a checkbox losing its 3-state
		// style while it has a gray checkmark in it (which incidentally in this case only changes
		// the appearance of the control, not the internal stored value in this case).
		if (style_needed_changing && style_change_ok)
			do_invalidate_rect = true;

		// Do the below only after applying the styles above since part of it requires that the style be
		// updated and applied above.
		switch (aControl.type)
		{
		case GUI_CONTROL_UPDOWN:
			ControlSetUpDownOptions(aControl, aOpt);
			break;
		case GUI_CONTROL_SLIDER:
			ControlSetSliderOptions(aControl, aOpt);
			if (aOpt.style_remove & TBS_TOOLTIPS)
				SendMessage(aControl.hwnd, TBM_SETTOOLTIPS, NULL, 0); // i.e. removing the TBS_TOOLTIPS style is not enough.
			break;
		case GUI_CONTROL_LISTVIEW:
			ControlSetListViewOptions(aControl, aOpt);
			break;
		case GUI_CONTROL_TREEVIEW:
			ControlSetTreeViewOptions(aControl, aOpt);
			break;
		case GUI_CONTROL_PROGRESS:
			ControlSetProgressOptions(aControl, aOpt, new_style);
			// Above strips theme if required by new options.  It also applies new colors.
			if (current_style != new_style)
				// SetWindowLong with GWL_STYLE resets the position, so restore it here.
				SendMessage(aControl.hwnd, PBM_SETPOS, current_value_uint, 0);
			break;
		case GUI_CONTROL_EDIT:
			if (aOpt.tabstop_count)
			{
				SendMessage(aControl.hwnd, EM_SETTABSTOPS, aOpt.tabstop_count, (LPARAM)aOpt.tabstop);
				// MSDN: "If the application is changing the tab stops for text already in the edit control,
				// it should call the InvalidateRect function to redraw the edit control window."
				do_invalidate_rect = true; // Override the default.
			}
			break;
		case GUI_CONTROL_LISTBOX:
			if (aOpt.tabstop_count)
			{
				SendMessage(aControl.hwnd, LB_SETTABSTOPS, aOpt.tabstop_count, (LPARAM)aOpt.tabstop);
				do_invalidate_rect = true; // This is done for the same reason that EDIT (above) does it.
			}
			break;
		case GUI_CONTROL_STATUSBAR:
			if (aOpt.color_bk != CLR_INVALID) // Explicit color change was requested.
				SendMessage(aControl.hwnd, SB_SETBKCOLOR, 0, aOpt.color_bk);
			break;
		}

		if (aOpt.color != CLR_INVALID) // Explicit color change was requested.
		{
			// ListView and TreeView controls were already handled above, but are also supported here.
			ControlSetTextColor(aControl, aOpt.color);
			do_invalidate_rect = true;
		}

		if (aOpt.redraw)
		{
			SendMessage(aControl.hwnd, WM_SETREDRAW, aOpt.redraw == CONDITION_TRUE, 0);
			if (aOpt.redraw == CONDITION_TRUE // Since redrawing is being turned back on, invalidate the control so that it updates itself.
				&& aControl.type != GUI_CONTROL_TREEVIEW) // This type is documented not to need it; others like ListView are not, so might need it on some OSes or under some conditions.
				do_invalidate_rect = true;
		}

		if (do_invalidate_rect)
			InvalidateRect(aControl.hwnd, NULL, TRUE); // Assume there's text in the control.

		if (style_needed_changing && !style_change_ok)
			return g_script.ThrowRuntimeException(ERR_FAILED);
	} // aControl.hwnd is not NULL

	return OK;
}



void GuiType::ControlInitOptions(GuiControlOptionsType &aOpt, GuiControlType &aControl)
// Not done as class to avoid code-size overhead of initializer list, etc.
{
	ZeroMemory(&aOpt, sizeof(GuiControlOptionsType));
	if (aControl.type == GUI_CONTROL_LISTVIEW) // Since this doesn't have the _add and _remove components, must initialize.
	{
		if (aControl.hwnd)
			aOpt.listview_style = ListView_GetExtendedListViewStyle(aControl.hwnd); // Will have no effect on 95/NT4 that lack comctl32.dll 4.70+ distributed with MSIE 3.x
		aOpt.listview_view = -1;  // Indicate "unspecified" so that changes can be detected.
	}
	aOpt.x = aOpt.y = aOpt.width = aOpt.height = COORD_UNSPECIFIED;
	aOpt.color = CLR_INVALID;
	aOpt.color_bk = CLR_INVALID;
	// Above: If it stays unaltered, CLR_INVALID means "leave color as it is".  This is for
	// use with "GuiControl, +option" so that ControlSetProgressOptions() and others know that
	// the "+/-Background" item was not among the options in the list.  The reason this is needed
	// for background color but not bar color is that bar_color is stored as a control attribute,
	// but to save memory, background color is not.  In addition, there is no good way to ask a
	// progress control what its background color currently is.
}



void GuiType::ControlAddItems(GuiControlType &aControl, Array *aObj)
{
	if (!aObj)
		return;

	UINT msg_add;
	int requested_index = 0;

	switch (aControl.type)
	{
	case GUI_CONTROL_TAB: // These cases must be listed anyway to do a break vs. return, so might as well init conditionally rather than unconditionally.
		requested_index = TabCtrl_GetItemCount(aControl.hwnd); // So that tabs are "appended at the end of the control's list", as documented for GuiControl.
		// Fall through:
	case GUI_CONTROL_LISTVIEW:
		msg_add = 0;
		break;
	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
		msg_add = CB_ADDSTRING;
		break;
	case GUI_CONTROL_LISTBOX:
		msg_add = LB_ADDSTRING;
		break;
	default: // Do nothing for any other control type that somehow makes its way here.
		return;
	}

	LRESULT item_index;
	TCHAR num_buf[MAX_NUMBER_SIZE];
	Array::index_t obj_index;
	ExprTokenType tok;

	// For tab controls:
	TCITEM tci;
	tci.mask = TCIF_TEXT | TCIF_IMAGE; // Simpler just to init unconditionally rather than checking control type.
	tci.iImage = -1;

	// For ListView:
	LVCOLUMN lvc;
	lvc.mask = LVCF_TEXT; // Simpler just to init unconditionally rather than checking control type.

	// Check *this_field at the top too, in case list ends in delimiter.
	for (obj_index = 0; aObj->ItemToToken(obj_index, tok); ++obj_index)
	{
		LPTSTR this_field = TokenToString(tok, num_buf);

		// Add the item:
		switch (aControl.type)
		{
		case GUI_CONTROL_TAB:
			if (requested_index > MAX_TABS_PER_CONTROL - 1) // Unlikely, but indicate failure if so.
				item_index = -1;
			else
			{
				tci.pszText = this_field;
				item_index = TabCtrl_InsertItem(aControl.hwnd, requested_index, &tci);
				if (item_index != -1) // item_index is used later below as an indicator of success.
					++requested_index;
			}
			break;
		case GUI_CONTROL_LISTVIEW:
			lvc.pszText = this_field;
			item_index = ListView_InsertColumn(aControl.hwnd, requested_index, &lvc);
			if (item_index != -1) // item_index is used later below as an indicator of success.
				++requested_index;
			break;
		default:
			item_index = SendMessage(aControl.hwnd, msg_add, 0, (LPARAM)this_field); // In this case, ignore any errors, namely CB_ERR/LB_ERR and CB_ERRSPACE).
			// For the above, item_index must be retrieved and used as the item's index because it might
			// be different than expected if the control's SORT style is in effect.
		}
	} // for()

	if (aControl.type == GUI_CONTROL_LISTVIEW)
	{
		// Fix for v1.0.36.03: requested_index is already one beyond the number of columns that were added
		// because it's always set up for the next column that would be added if there were any.
		// Therefore, there is no need to add one to it to get the column count.
		aControl.union_lv_attrib->col_count = requested_index; // Keep track of column count, mostly so that LV_ModifyCol and such can properly maintain the array of columns).
		// It seems a useful default to do a basic auto-size upon creation, even though it won't take into
		// account contents of the rows or the later presence of a vertical scroll bar. The last column
		// will be overlapped when/if a v-scrollbar appears, which would produce an h-scrollbar too.
		// It seems best to retain this behavior rather than trying to shrink the last column to allow
		// room for a scroll bar because: 1) Having it use all available width is desirable at least
		// some of the time (such as times when there will be only a few rows; 2) It simplifies the code.
		// This method of auto-sizing each column to fit its text works much better than setting
		// lvc.cx to ListView_GetStringWidth upon creation of the column.
		if (ListView_GetView(aControl.hwnd) == LV_VIEW_DETAILS)
			for (int i = 0; i < requested_index; ++i) // Auto-size each column.
				ListView_SetColumnWidth(aControl.hwnd, i, LVSCW_AUTOSIZE_USEHEADER);
	}
}



void GuiType::ControlSetChoice(GuiControlType &aControl, int aChoice)
{
	--aChoice; // Convert one-based to zero-based.
	switch (aControl.type)
	{
	case GUI_CONTROL_TAB:
		// MSDN: "A tab control does not send a TCN_SELCHANGING or TCN_SELCHANGE notification message
		// when a tab is selected using the TCM_SETCURSEL message."
		TabCtrl_SetCurSel(aControl.hwnd, aChoice);
		break;
	case GUI_CONTROL_DROPDOWNLIST:
	case GUI_CONTROL_COMBOBOX:
		SendMessage(aControl.hwnd, CB_SETCURSEL, (WPARAM)aChoice, 0);
		break;
	case GUI_CONTROL_LISTBOX:
		// Multi-select box requires a different message to have a cumulative effect:
		if (GetWindowLong(aControl.hwnd, GWL_STYLE) & (LBS_EXTENDEDSEL|LBS_MULTIPLESEL))
			SendMessage(aControl.hwnd, LB_SETSEL, (WPARAM)TRUE, (LPARAM)aChoice);
		else
			SendMessage(aControl.hwnd, LB_SETCURSEL, (WPARAM)aChoice, 0);
		break;
	}
}



ResultType GuiType::ControlLoadPicture(GuiControlType &aControl, LPCTSTR aFilename, int aWidth, int aHeight, int aIconNumber)
{
	// LoadPicture() uses CopyImage() to scale the image, which seems to provide better scaling
	// quality than using MoveWindow() (followed by redrawing the parent window) on the static
	// control that contains the image.
	int image_type;
	HBITMAP new_image = LoadPicture(aFilename, aWidth, aHeight, image_type, aIconNumber
		, aControl.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT);
	if (!new_image && *aFilename)
		return FAIL; // Caller will report the error.

	// In light of the below, it seems best to delete the bitmaps whenever the control changes
	// to a new image or whenever the control is destroyed.  Otherwise, if a control or its
	// parent window is destroyed and recreated many times, memory allocation would continue
	// to grow from all of the abandoned bitmaps/icons.
	// MSDN: "When you are finished using a bitmap...loaded without specifying the LR_SHARED flag,
	// you can release its associated memory by calling...DeleteObject."
	// MSDN: "The system automatically deletes these resources when the process that created them
	// terminates, however, calling the appropriate function saves memory and decreases the size
	// of the process's working set."
	// v1.0.40.12: For maintainability, destroy the handle returned by STM_SETIMAGE, even though it
	// should be identical to control.union_hbitmap (due to a call to STM_GETIMAGE in another section).
	// UPDATE: This is now done after LoadPicture() is called to reduce the length of time between
	// the background being erased and the new picture being drawn, which might decrease flicker.
	// It's still done as a separate step to setting the new bitmap due to uncertainty about how
	// the control handles changing the image type, and the notes below about animation.
	if (aControl.union_hbitmap)
		if (aControl.attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR) // union_hbitmap is an icon or cursor.
			// The control's image is set to NULL for the following reasons:
			// 1) It turns off the control's animation timer in case the new image is not animated.
			// 2) It feels a little bit safer to destroy the image only after it has been removed
			//    from the control.
			// NOTE: IMAGE_ICON or IMAGE_CURSOR must be passed, not IMAGE_BITMAP.  Otherwise the
			// animated property of the control (via a timer that the control created) will remain
			// in effect for the next image, even if it isn't animated, which results in a
			// flashing/redrawing effect:
			DestroyIcon((HICON)SendMessage(aControl.hwnd, STM_SETIMAGE, IMAGE_CURSOR, NULL));
			// DestroyIcon() works on cursors too.  See notes in LoadPicture().
		else // union_hbitmap is a bitmap
			DeleteObject((HGDIOBJ)SendMessage(aControl.hwnd, STM_SETIMAGE, IMAGE_BITMAP, NULL));
	aControl.union_hbitmap = new_image;
	if (!new_image)
	{
		// By design, no error is reported for a blank aFilename.
		// For simplicity, any existing SS_BITMAP/SS_ICON style bit is not removed.
		return OK;
	}
	if (image_type == IMAGE_ICON && aControl.background_color == CLR_TRANSPARENT)
	{
		// Static controls don't appear to support background transparency with icons,
		// so convert the icon to a bitmap.
		if (HBITMAP hbitmap = IconToBitmap32((HICON)aControl.union_hbitmap, false))
		{
			DestroyIcon((HICON)aControl.union_hbitmap);
			aControl.union_hbitmap = hbitmap;
			image_type = IMAGE_BITMAP;
		}
	}
	// For image to display correctly, must apply SS_ICON for cursors/icons and SS_BITMAP for bitmaps.
	// This style change is made *after* the control is created so that the act of creation doesn't
	// attempt to load the image from a resource (which as documented by SS_ICON/SS_BITMAP, would happen
	// since text is interpreted as the name of a bitmap in the resource file).
	DWORD style = GetWindowLong(aControl.hwnd, GWL_STYLE);
	DWORD style_image_type = style & 0x0F;
	style &= ~0x0F;  // Purge the low-order four bits in case style-image-type needs to be altered below.
	if (image_type == IMAGE_BITMAP)
	{
		if (style_image_type != SS_BITMAP)
			SetWindowLong(aControl.hwnd, GWL_STYLE, style | SS_BITMAP);
	}
	else // Icon or Cursor.
		if (style_image_type != SS_ICON) // Must apply SS_ICON or such handles cannot be displayed.
			SetWindowLong(aControl.hwnd, GWL_STYLE, style | SS_ICON);
	// Above uses ~0x0F to ensure the lowest four/five bits are pure.
	// Also note that it does not seem correct to use SS_TYPEMASK if bitmaps/icons can also have
	// any of the following styles.  MSDN confirms(?) this by saying that SS_TYPEMASK is out of date
	// and should not be used:
	//#define SS_ETCHEDHORZ       0x00000010L
	//#define SS_ETCHEDVERT       0x00000011L
	//#define SS_ETCHEDFRAME      0x00000012L
	SendMessage(aControl.hwnd, STM_SETIMAGE, (WPARAM)image_type, (LPARAM)aControl.union_hbitmap);
	// Fix for 1.0.40.12: The below was added because STM_SETIMAGE above may have caused the control to
	// create a new hbitmap (possibly only for alpha channel bitmaps on XP, but might also apply to icons),
	// in which case we now have two handles: the one inside the control and the one from which
	// it was copied.  Task Manager confirms that the control does not delete the original
	// handle when it creates a new handle.  Rather than waiting until later to delete the handle,
	// it seems best to do it here so that:
	// 1) The script uses less memory during the time that the picture control exists.
	// 2) Don't have to delete two handles (control.union_hbitmap and the one returned by STM_SETIMAGE)
	//    when the time comes to change the image inside the control.
	//
	// MSDN: "With Microsoft Windows XP, if the bitmap passed in the STM_SETIMAGE message contains pixels
	// with non-zero alpha, the static control takes a copy of the bitmap. This copied bitmap is returned
	// by the next STM_SETIMAGE message... if it does not check and release the bitmaps returned from
	// STM_SETIMAGE messages, the bitmaps are leaked."
	HBITMAP hbitmap_actual;
	if ((hbitmap_actual = (HBITMAP)SendMessage(aControl.hwnd, STM_GETIMAGE, image_type, 0)) // Assign
		&& hbitmap_actual != aControl.union_hbitmap)  // The control decided to make a new handle.
	{
		if (image_type == IMAGE_BITMAP)
			DeleteObject(aControl.union_hbitmap);
		else // Icon or cursor.
			DestroyIcon((HICON)aControl.union_hbitmap); // Works on cursors too.
		// In additional to improving maintainability, the following might also be necessary to allow
		// Gui::Destroy() to avoid a memory leak when the picture control is destroyed as a result
		// of its parent being destroyed (though I've read that the control is supposed to destroy its
		// hbitmap when it was directly responsible for creating it originally [but not otherwise]):
		aControl.union_hbitmap = hbitmap_actual;
	}
	if (image_type == IMAGE_BITMAP)
		aControl.attrib &= ~GUI_CONTROL_ATTRIB_ALTBEHAVIOR;  // Flag it as a bitmap so that DeleteObject vs. DestroyIcon will be called for it.
	else // Cursor or Icon, which are functionally identical for our purposes.
		aControl.attrib |= GUI_CONTROL_ATTRIB_ALTBEHAVIOR;
	return OK;
}



FResult GuiType::Show(optl<StrArg> aOptions)
{
	GUI_MUST_HAVE_HWND;

	// In the future, it seems best to rely on mShowIsInProgress to prevent the Window Proc from ever
	// doing a MsgSleep() to launch a script subroutine.  This is because if anything we do in this
	// function results in a launch of the Window Proc (such as MoveWindow and ShowWindow), our
	// activity here might be interrupted in a destructive way.  For example, if a script subroutine
	// is launched while we're in the middle of something here, our activity is suspended until
	// the subroutine completes and the call stack collapses back to here.  But if that subroutine
	// recursively calls us while the prior call is still in progress, the mShowIsInProgress would
	// be set to false when that layer completes, leaving it false when it really should be true
	// because our layer isn't done yet.
	mShowIsInProgress = true; // Signal WM_SIZE to queue the GuiSize launch.  We'll unqueue via MsgSleep() when we're done.

	// Set defaults, to be overridden by the presence of zero or more options:
	int x = COORD_UNSPECIFIED;
	int y = COORD_UNSPECIFIED;
	int width = COORD_UNSPECIFIED;
	int height = COORD_UNSPECIFIED;
	bool auto_size = false;

	BOOL is_maximized = IsZoomed(mHwnd); // Safer to assume it's possible for both to be true simultaneously in
	BOOL is_minimized = IsIconic(mHwnd); // future/past OSes.
	int show_mode;
	// There is evidence that SW_SHOWNORMAL might be better than SW_SHOW for the first showing because
	// someone reported that a window appears centered on the screen for its first showing even if some
	// other position was specified.  In addition, MSDN says (without explanation): "An application should
	// specify [SW_SHOWNORMAL] when displaying the window for the first time."  However, SW_SHOWNORMAL is
	// avoided after the first showing of the window because that would probably also do a "restore" on the
	// window if it was maximized previously.  Note that the description of SW_SHOWNORMAL is virtually the
	// same as that of SW_RESTORE in MSDN.  UPDATE: mGuiShowHasNeverBeenDone is used here instead of mFirstActivation
	// because it seems more flexible to have "Gui Show" behave consistently (SW_SHOW) every time after
	// the first use of "Gui Show".  UPDATE: Since SW_SHOW seems to have no effect on minimized windows,
	// at least on XP, and since such a minimized window will be restored by action of SetForegroundWindowEx(),
	// it seems best to unconditionally use SW_SHOWNORMAL, rather than "mGuiShowHasNeverBeenDone ? SW_SHOWNORMAL : SW_SHOW".
	// This is done so that the window will be restored and thus have a better chance of being successfully
	// activated (and thus not requiring the call to SetForegroundWindowEx()).
	// v1.0.44.08: Fixed to default to SW_SHOW for currently-maximized windows so that they don't get unmaximized
	// by "Gui Show" (unless other options require it).  Also, it's been observed that SW_SHOWNORMAL differs from
	// SW_RESTORE in at least one way: When the target is a minimized window, SW_SHOWNORMAL will both restore it
	// and unmaximize it.  But SW_RESTORE unminimizes the window in a way that retains its maximized state (if
	// it was previously maximized).  Therefore, SW_RESTORE is now the default if the window is currently minimized.
	if (is_minimized)
		show_mode = SW_RESTORE; // See above comments. For backward compatibility, window is unminimized even if it was previously hidden (rather than simply showing its taskbar button and keeping it minimized).
	else if (is_maximized)
		show_mode = SW_SHOW; // See above.
	else
		show_mode = SW_SHOWNORMAL;

	for (auto cp_end = aOptions.value_or_empty(), cp = cp_end; *cp; cp = cp_end)
	{
		switch(ctoupper(*cp))
		{
		// For options such as W, H, X and Y: Use _ttoi() vs. ATOI() to avoid interpreting something like 0x01B
		// as hex when in fact the B was meant to be an option letter.
		case ' ':
		case '\t':
			++cp_end;
			break;
		case 'A':
			if (!_tcsnicmp(cp, _T("AutoSize"), 8))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				cp_end += 8;
				auto_size = true;
			}
			break;
		case 'C':
			if (!_tcsnicmp(cp, _T("Center"), 6))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				cp_end += 6;
				x = COORD_CENTERED;
				y = COORD_CENTERED;
				// If the window is currently maximized, show_mode isn't set to SW_RESTORE unconditionally here
				// due to obscurity and because it might reduce flexibility.  If the window is currently minimized,
				// above has already set the default show_mode to SW_RESTORE to ensure correct operation of
				// something like "Gui, Show, Center".
			}
			break;
		case 'M':
			if (!_tcsnicmp(cp, _T("Minimize"), 8)) // Seems best to reserve "Min" for other things, such as Min W/H. "Minimize" is also more self-documenting.
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				cp_end += 8;
				show_mode = SW_MINIMIZE;  // Seems more typically useful/desirable than SW_SHOWMINIMIZED.
			}
			else if (!_tcsnicmp(cp, _T("Maximize"), 8))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				cp_end += 8;
				show_mode = SW_MAXIMIZE;  // SW_MAXIMIZE == SW_SHOWMAXIMIZED
			}
			break;
		case 'N':
			if (!_tcsnicmp(cp, _T("NA"), 2))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				cp_end += 2;
				show_mode = SW_SHOWNA;
			}
			else if (!_tcsnicmp(cp, _T("NoActivate"), 10))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				cp_end += 10;
				show_mode = SW_SHOWNOACTIVATE;
			}
			break;
		case 'R':
			if (!_tcsnicmp(cp, _T("Restore"), 7))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				cp_end += 7;
				show_mode = SW_RESTORE;
			}
			break;
		case 'X':
		case 'Y':
			if (!_tcsnicmp(cp + 1, _T("Center"), 6))
			{
				if (ctoupper(*cp) == 'X')
					x = COORD_CENTERED;
				else
					y = COORD_CENTERED;
				cp_end += 7; // 7 in this case since we're working with cp + 1
				continue;
			}
			// OTHERWISE, FALL THROUGH:
		case 'W':
		case 'H':
			if (!_tcsnicmp(cp, _T("Hide"), 4))
			{
				// Skip over the text of the name so that it isn't interpreted as option letters.
				cp_end += 4;
				show_mode = SW_HIDE;
				continue;
			}
			// Use _tcstod vs _tcstol to support common cases like % "w" A_ScreenWidth/4 even
			// when the current float format is something unusual.  Something like "w400.5"
			// could be considered invalid, but is allowed for simplicity and convenience.
			int n = (int)_tcstod(cp + 1, const_cast<LPTSTR*>(&cp_end));
			if (cp_end == cp + 1) // No number.
			{
				cp_end = cp; // Flag it as invalid.
				break;
			}
			switch (ctoupper(*cp))
			{
			case 'X': x = n; break;
			case 'Y': y = n; break;
			// Allow any width/height to be specified so that the window can be "rolled up" to its title bar:
			case 'W': width = Scale(n); break;
			case 'H': height = Scale(n); break;
			}
			break;
		} // switch()
		if (cp_end == cp)
		{
			if (!ValueError(ERR_INVALID_OPTION, cp, FAIL_OR_OK))
				return FAIL;
			// Otherwise, user wants to continue.  Ignore any invalid characters up to the next valid one.
			if (!(cp_end = StrChrAny(cp_end, _T("ACMNRXYWH"))))
				break;
		}
	} // for()

	int width_orig = width;
	int height_orig = height;

	// The following section must be done prior to any calls to GetWindow/ClientRect(mHwnd) because
	// neither of them can retrieve the correct dimensions of a minimized window.  Similarly, if
	// the window is maximized but is about to be restored, do that prior to getting any of mHwnd's
	// rectangles because we want to use the restored size as the basis for centering, resizing, etc.
	// If show_mode is "hide", move the window only after hiding it (to reduce screen flicker).
	// If the window is being restored from a minimized or maximized state, move the window only
	// after restoring it; otherwise, any resize to be done by the MoveWindow() might not take effect.
	// Note that SW_SHOWNOACTIVATE is very similar to SW_RESTORE in its effects.
	bool show_was_done = false;
	if (show_mode == SW_HIDE // Hiding a window or restoring a window known to be minimized/maximized.
		|| (show_mode == SW_RESTORE || show_mode == SW_SHOWNOACTIVATE) && (is_maximized || is_minimized)) // v1.0.44.08: Fixed missing "show_mode ==".
	{
		ShowWindow(mHwnd, show_mode);
		show_was_done = true;
	}
	// Note that SW_RESTORE and SW_SHOWNOACTIVATE will show a window if it's hidden.  Therefore, just
	// because the window is not in need of restoring doesn't mean the ShowWindow() call is skipped.
	// That is why show_was_done is left false in such cases.

	if (mGuiShowHasNeverBeenDone)
	{
		// Ensure all Tab3 controls have been sized prior to autosizing or showing the GUI.
		// This has been confirmed to work even if IsIconic(), although that would require
		// calling WinMinimize or option +0x20000000, since SW_MINIMIZE hasn't been applied.
		for (GuiIndexType u = 0; u < mControlCount; ++u)
		{
			if (mControl[u]->type == GUI_CONTROL_TAB)
				AutoSizeTabControl(*mControl[u]);
		}
		// For the rare case that Show is called before any controls are added or MarginX/Y
		// is queried, the following ensures mMarginx and mMarginY have usable values:
		SetDefaultMargins();
	}

	// Due to the checking above, if the window is minimized/maximized now, that means it will still be
	// minimized/maximized when this function is done.  As a result, it's not really valid to call
	// MoveWindow() for any purpose (auto-centering, auto-sizing, new position, new size, etc.).
	// The below is especially necessary for minimized windows because it avoid calculating window
	// dimensions, auto-centering, etc. based on incorrect values returned by GetWindow/ClientRect(mHwnd).
	// Update: For flexibility, it seems best to allow a maximized window to be moved, which might be
	// valid on a multi-monitor system.  This maintains flexibility and doesn't appear to give up
	// anything because the script can do an explicit "Gui, Show, Restore" prior to a
	// "Gui, Show, x33 y44 w400" to be sure the window is restored before the operation (or combine
	// both of those commands into one).
	bool allow_move_window;
	RECT rect;

	if (allow_move_window = !IsIconic(mHwnd)) // Call IsIconic() again in case above changed the window's state.
	{
		if (auto_size) // Check this one first so that it takes precedence over mGuiShowHasNeverBeenDone below.
		{
			// Find out a different set of max extents rather than using mMaxExtentRight/Down, which should
			// not be altered because they are used to position any subsequently added controls.
			width = 0;
			height = 0;
			for (GuiIndexType u = 0; u < mControlCount; ++u)
			{
				GuiControlType &control = *mControl[u];
				if (control.type != GUI_CONTROL_STATUSBAR // Status bar is compensated for in a diff. way.
					&& (control.tab_control_index == MAX_TAB_CONTROLS || GetParent(control.hwnd) == mHwnd) // Skip controls which are clipped by a parent control.
					&& GetWindowLong(control.hwnd, GWL_STYLE) & WS_VISIBLE) // Don't use IsWindowVisible() in case parent window is hidden.
				{
					GetWindowRect(control.hwnd, &rect);
					MapWindowPoints(NULL, mHwnd, (LPPOINT)&rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
					if (rect.right > width)
						width = rect.right;
					if (rect.bottom > height)
						height = rect.bottom;
				}
			}
			if (width > 0)
				width += mMarginX;
			if (height > 0)
				height += mMarginY;
			// Don't use IsWindowVisible() because that would say that bar is hidden just because the parent
			// window is hidden.  We want to know if the bar is truly hidden, separately from the window:
			if (mStatusBarHwnd && GetWindowLong(mStatusBarHwnd, GWL_STYLE) & WS_VISIBLE)
			{
				GetWindowRect(mStatusBarHwnd, &rect); // GetWindowRect vs. GetClientRect to include any borders it might have.
				height += rect.bottom - rect.top;
			}
		}
		else if (width == COORD_UNSPECIFIED || height == COORD_UNSPECIFIED)
		{
			if (mGuiShowHasNeverBeenDone) // By default, center the window if this is the first use of "Gui Show" (even "Gui Show, Hide").
			{
				if (width == COORD_UNSPECIFIED)
					width = mMaxExtentRight + mMarginX;
				if (height == COORD_UNSPECIFIED)
				{
					height = mMaxExtentDown + mMarginY;
					if (mStatusBarHwnd && GetWindowLong(mStatusBarHwnd, GWL_STYLE) & WS_VISIBLE) // See comments in similar section above.
					{
						GetWindowRect(mStatusBarHwnd, &rect); // GetWindowRect vs. GetClientRect to include any borders it might have.
						height += rect.bottom - rect.top;
					}
				}
			}
			else
			{
				GetClientRect(mHwnd, &rect);
				if (width == COORD_UNSPECIFIED) // Keep the current client width, as documented.
					width = rect.right - rect.left;
				if (height == COORD_UNSPECIFIED) // Keep the current client height, as documented.
					height = rect.bottom - rect.top;
			}
		}

		// v1.0.44.13: For code size reasons and due to rarity-of-need, the following isn't done (also, it
		// can't catch all such situations; e.g. when "Gui +MinSize" can be used after "Gui Show".
		// Plus it might add a bit of flexibility to allow "Gui Show" to override min/max:
		// Older: The following prevents situations in which the window starts off at a size that's
		// too big or too small, which in turn causes it to snap to the min/max size the moment
		// the user tries to drag-move or drag-resize it:
		//if (mMinWidth >= 0 && width < mMinWidth) // mMinWidth >= 0 covers both COORD_UNSPECIFIED and COORD_CENTERED.
		//	width = mMinWidth;
		//else if (mMaxWidth >= 0 && width > mMaxWidth)
		//	width = mMaxWidth;
		//if (mMinHeight >= 0 && height < mMinHeight)
		//	height = mMinHeight;
		//else if (mMaxHeight >= 0 && height > mMaxHeight)
		//	height = mMaxHeight;

	} // if (allow_move_window)

	if (mGuiShowHasNeverBeenDone)
	{
		// Update any tab controls to show only their correct pane.  This should only be necessary
		// upon the first "Gui Show" (even "Gui, Show, Hide") of the window because subsequent switches
		// of the control's tab should result in a TCN_SELCHANGE notification.
		if (mTabControlCount)
			for (GuiIndexType u = 0; u < mControlCount; ++u)
				if (mControl[u]->type == GUI_CONTROL_TAB)
					ControlUpdateCurrentTab(*mControl[u], false); // Pass false so that default/z-order focus is used across entire window.
		// By default, center the window if this is the first time it's being shown:
		if (x == COORD_UNSPECIFIED)
			x = COORD_CENTERED;
		if (y == COORD_UNSPECIFIED)
			y = COORD_CENTERED;
	}

	BOOL is_visible = IsWindowVisible(mHwnd);

	if (allow_move_window)
	{
		// The above has determined the height/width of the client area.  From that area, determine
		// the window's new screen rect, including title bar, borders, etc.
		// If the window has a border or caption this also changes top & left *slightly* from zero.
		RECT rect = {0, 0, width, height}; // left,top,right,bottom
		LONG style = GetWindowLong(mHwnd, GWL_STYLE);
		BOOL has_menu = GetMenu(mHwnd) ? TRUE : FALSE;
		AdjustWindowRectEx(&rect, style, has_menu, GetWindowLong(mHwnd, GWL_EXSTYLE));
		// MSDN: "The AdjustWindowRectEx function does not take the WS_VSCROLL or WS_HSCROLL styles into
		// account. To account for the scroll bars, call the GetSystemMetrics function with SM_CXVSCROLL
		// or SM_CYHSCROLL."
		if (style & WS_HSCROLL)
			rect.bottom += GetSystemMetrics(SM_CYHSCROLL);
		if (style & WS_VSCROLL)
			rect.right += GetSystemMetrics(SM_CXVSCROLL);
		// Compensate for menu wrapping: https://blogs.msdn.microsoft.com/oldnewthing/20030911-00/?p=42553/
		if (has_menu)
		{
			RECT rcTemp = rect;
			rcTemp.bottom = 0x7FFF;
			SendMessage(mHwnd, WM_NCCALCSIZE, FALSE, (LPARAM)&rcTemp);
			rect.bottom += rcTemp.top;
		}

		int nc_width = rect.right - rect.left - width;
		int nc_height = rect.bottom - rect.top - height;

		width = rect.right - rect.left;  // rect.left might be slightly less than zero.
		height = rect.bottom - rect.top; // rect.top might be slightly less than zero. A status bar is properly handled since it's inside the window's client area.

		RECT work_rect;
		bool is_child_window = mOwner && (style & WS_CHILD);
		if (is_child_window)
			GetClientRect(mOwner, &work_rect); // Center within parent window (our position is set relative to mOwner's client area, not in screen coordinates).
		else
			SystemParametersInfo(SPI_GETWORKAREA, 0, &work_rect, 0);  // Get desktop rect excluding task bar.
		int work_width = work_rect.right - work_rect.left;  // Note that "left" won't be zero if task bar is on left!
		int work_height = work_rect.bottom - work_rect.top; // Note that "top" won't be zero if task bar is on top!

		// Seems best to restrict window size to the size of the desktop whenever explicit sizes
		// weren't given, since most users would probably want that.  But only on first use of
		// "Gui Show" (even "Gui, Show, Hide"):
		if (mGuiShowHasNeverBeenDone && !is_child_window)
		{
			if (width_orig == COORD_UNSPECIFIED && width > work_width)
				width = work_width;
			if (height_orig == COORD_UNSPECIFIED && height > work_height)
				height = work_height;
		}

		if (x == COORD_CENTERED || y == COORD_CENTERED) // Center it, based on its dimensions determined above.
		{
			// This does not handle multi-monitor systems explicitly, and has no need to do so since
			// SPI_GETWORKAREA "Retrieves the size of the work area on the primary display monitor".
			if (x == COORD_CENTERED)
				x = work_rect.left + ((work_width - width) / 2);
			if (y == COORD_CENTERED)
				y = work_rect.top + ((work_height - height) / 2);
		}

		RECT old_rect;
		GetWindowRect(mHwnd, &old_rect);
		int old_width = old_rect.right - old_rect.left;
		int old_height = old_rect.bottom - old_rect.top;

		// Added for v1.0.44.13:
		// Below is done inside this block (allow_move_window) because that way, it should always
		// execute whenever mGuiShowHasNeverBeenDone (since the window shouldn't be iconic prior to
		// its first showing).  In addition, below must be down prior to any ShowWindow() that does
		// a minimize or maximize because that would prevent GetWindowRect/GetClientRect calculations
		// below from working properly.
		// v1.1.34.03: It's now done before the MoveWindow call below, so that the initial size is
		// limited by the correct values.  nc_width and nc_height are expected to be accurate, and
		// must be used rather than calling GetClientRect and calculating the difference, as that's
		// likely to be 0x0 prior to calling MoveWindow.
		if (mGuiShowHasNeverBeenDone) // This is the first showing of this window.
		{
			// Now that the window's style, edge type, title bar, menu bar, and other non-client attributes have
			// likely (but not certainly) been determined, adjust MinMaxSize values from client size to
			// entire-window size for use with WM_GETMINMAXINFO.

			if (mMinWidth == COORD_CENTERED) // COORD_CENTERED is the flag that means, "use window's current, total width."
				mMinWidth = width;
			else if (mMinWidth != COORD_UNSPECIFIED)
				mMinWidth += nc_width;

			if (mMinHeight == COORD_CENTERED)
				mMinHeight = height;
			else if (mMinHeight != COORD_UNSPECIFIED)
				mMinHeight += nc_height;

			if (mMaxWidth == COORD_CENTERED)
				mMaxWidth = width;
			else if (mMaxWidth != COORD_UNSPECIFIED)
				mMaxWidth += nc_width;

			if (mMaxHeight == COORD_CENTERED)
				mMaxHeight = height;
			else if (mMaxHeight != COORD_UNSPECIFIED)
				mMaxHeight += nc_height;
		} // if (mGuiShowHasNeverBeenDone)

		// Avoid calling MoveWindow() if nothing changed because it might repaint/redraw even if window size/pos
		// didn't change:
		if (width != old_width || height != old_height || (x != COORD_UNSPECIFIED && x != old_rect.left)
			|| (y != COORD_UNSPECIFIED && y != old_rect.top)) // v1.0.45: Fixed to be old_rect.top not old_rect.bottom.
		{
			// v1.0.44.08: Window state gets messed up if it's resized without first unmaximizing it (for example,
			// it can't be resized by dragging its lower-right corner).  So it seems best to unmaximize, perhaps
			// even when merely the position is changing rather than the size (even on a multimonitor system,
			// it might not be valid to reposition a maximized window without unmaximizing it?)
			if (IsZoomed(mHwnd)) // Call IsZoomed() again in case above changed the state. No need to check IsIconic() because above already set default show-mode to SW_RESTORE for such windows.
				ShowWindow(mHwnd, SW_RESTORE); // But restore isn't done for something like "Gui, Show, Center" because it's too obscure and might reduce flexibility (debatable).
			if (is_child_window)
				ScreenToClient(mOwner, (LPPOINT)&old_rect); // Make the old coords relative to parent.
			MoveWindow(mHwnd, x == COORD_UNSPECIFIED ? old_rect.left : x, y == COORD_UNSPECIFIED ? old_rect.top : y
				, width, height, is_visible);  // Do repaint if window is visible.
		}
	} // if (allow_move_window)

	// Note that for SW_MINIMIZE and SW_MAXIMIZE, the MoveWindow() above should be done prior to ShowWindow()
	// so that the window will "remember" its new size upon being restored later.
	if (!show_was_done)
	{
		if (show_mode == SW_MINIMIZE && GetForegroundWindow() == mHwnd)
		{
			// Before minimizing the window, call DefDlgProc to save the handle of the focused control.
			// This is only necessary when minimizing an active GUI by calling ShowWindow(), as it is
			// normally done by DefDlgProc when WM_SYSCOMMAND, SC_MINIMIZE is received.  Without this,
			// focus is reset to the first control with WS_TABSTOP when the window is restored.
			// We don't use WM_SYSCOMMAND to minimize the window because it would cause the "minimize"
			// system sound to be played, and that should probably only happen as a result of user action.
			DefDlgProc(mHwnd, WM_ACTIVATE, WA_INACTIVE, 0);
		}
		ShowWindow(mHwnd, show_mode);
	}
	VisibilityChanged(); // AddRef() to keep the object alive while the window is visible.

	bool we_did_the_first_activation = false; // Set default.

	switch(show_mode)
	{
	case SW_SHOW:
	case SW_SHOWNORMAL:
	case SW_MAXIMIZE:
	case SW_RESTORE:
		if (GetAncestor(mHwnd, GA_ROOT) != mHwnd) // It's a child window, perhaps due to the +Parent option.
			// Since this is a child window, SetForegroundWindowEx will fail and fall back to the "two-alts"
			// method of forcing activation, which causes some messages to be discarded due to the temporary
			// modal menu loop.  So don't even try it.
			break;
		if (mHwnd != GetForegroundWindow())
			SetForegroundWindowEx(mHwnd);   // In the above modes, try to force it to the foreground as documented.
		if (mFirstActivation)
		{
			// Since the window has never before been active, any of the above qualify as "first activation".
			// Thus, we are no longer at the first activation:
			mFirstActivation = false;
			we_did_the_first_activation = true; // And we're the ones who did the first activation.
		}
		break;
	// No action for these:
	//case SW_MINIMIZE:
	//case SW_SHOWNA:
	//case SW_SHOWNOACTIVATE:
	//case SW_HIDE:
	}

	// No attempt is made to handle the fact that Gui windows can be shown or activated via WinShow and
	// WinActivate.  In such cases, if the tab control itself is focused, mFirstActivation will still focus
	// a control inside the tab rather than leaving the tab control focused.  Similarly, if the window
	// was shown with NA or NOACTIVATE or MINIMIZE, when the first use of an activation mode of "Gui Show"
	// is done, even if it's far into the future, long after the user has activated and navigated in the
	// window, the same "first activation" behavior will be done anyway.  This is documented here as a
	// known limitation, since fixing it would probably add an unreasonable amount of complexity.
	if (we_did_the_first_activation)
	{
		HWND focused_hwnd = GetFocus(); // Window probably must be visible and active for GetFocus() to work.
		if (focused_hwnd)
		{
			if (GuiControlType *focused_control = FindControl(focused_hwnd))
			{
				// Since this is the first activation, if the focus wound up on a tab control itself as a result
				// of the above, focus the first control of that tab since that is traditional.  HOWEVER, do not
				// instead default tab controls to lacking WS_TABSTOP since it is traditional for them to have
				// that property, probably to aid accessibility.
				if (focused_control->type == GUI_CONTROL_TAB)
				{
					// v1.0.27: The following must be done, at least in some cases, because otherwise
					// controls outside of the tab control will not get drawn correctly.  I suspect this
					// is because at the exact moment execution reaches the line below, the window is in
					// a transitional state, with some WM_PAINT and other messages waiting in the queue
					// for it.  If those messages are not processed prior to ControlUpdateCurrentTab()'s
					// use of WM_SETREDRAW, they might get dropped out of the queue and lost.
					UpdateWindow(mHwnd);
					ControlUpdateCurrentTab(*focused_control, true);
				}
				else if (focused_control->type == GUI_CONTROL_BUTTON)
				{
					if (mDefaultButtonIndex < mControlCount && mControl[mDefaultButtonIndex] != focused_control)
					{
						// There is a Default button, but some other button was focused instead (because it was
						// the first input-capable control).  There are a few reasons to override this:
						//  1) It is counter-intuitive to have two different kinds of "default button".
						//  2) MsgBox (and maybe other system dialogs that have only buttons) focus the Default
						//     button by default, even if it is not the first button.
						//  3) In most other cases, when a button is focused it also temporarily becomes Default,
						//     so pressing Enter or Space will activate it; but due to the way the default message
						//     processing focuses the initial control, it ends up being focused but not Default,
						//     so Space and Enter click two different buttons.  This lasts until the user changes
						//     focus, even by switching to another window and back.
						//     This inconsistency can be avoided by using WM_NEXTDLGCTL to "refocus" the button,
						//     but then the button with "Default" option would only become the default once some
						//     non-button control is focused, which seems unlikely to be what the author intended.
						SetFocus(mControl[mDefaultButtonIndex]->hwnd);
						// This is only done if a button was focused by default, as if some other type of control
						// is focused, the Default button can often still be activated with Enter.  Most of the
						// time if there is an input field and a confirmation button, the input field should have
						// the initial focus.  There are numerous examples of system and third-party dialogs that
						// behave this way.
					}
				}
				//else focus has already been set appropriately, and nothing needs to be done.
			}
		}
		else // No window/control has keyboard focus (see comment below).
			SetFocus(mHwnd);
			// The above was added in v1.0.46.05 to fix the fact that a GUI window could be both active and
			// foreground yet not have keyboard focus.  This occurs under the following circumstances (and
			// possibly others):
			// 1) A script with a menu item that shows a GUI window is reloaded via its tray menu item "Reload".
			// 2) The GUI window is shown via its custom tray menu item.
			// 3) The window becomes active and foreground, but doesn't have keyboard focus (not even its
			//    GuiEscape label will work until you switch away from that window then back to it).
			// Note: Default handling of WM_SETFOCUS moves the focus to the first input-capable control.
			// This case should be unusual since default handling of WM_ACTIVATE is also to SetFocus().
	}

	mGuiShowHasNeverBeenDone = false;
	// It seems best to reset this prior to SLEEP below, but after the above line (for code clarity) since
	// otherwise it might get stuck in a true state if the SLEEP results in the launch of a script
	// subroutine that takes a long time to complete:
	mShowIsInProgress = false;

	// Update for v1.0.25: The below is now done last to prevent the GuiSize label (if any) from launching
	// while this function is still incomplete; in other words, don't allow the GuiSize label to launch
	// until after all of the above members and actions have been completed.
	// If this weren't done, whenever a command that blocks (fully uses) the main thread such as "Drive Eject"
	// immediately follows "Gui Show", the GUI window might not appear until afterward because our thread
	// never had a chance to call its WindowProc with all the messages needed to actually show the window:
	MsgSleep(-1);
	// UpdateWindow() would probably achieve the same effect as the above, but it feels safer to do
	// the above because it ensures that our message queue is empty prior to returning to our caller.

	return OK;
}



FResult GuiType::Minimize()
{
	GUI_MUST_HAVE_HWND;
	ShowWindow(mHwnd, SW_MINIMIZE);
	return OK;
}

FResult GuiType::Maximize()
{
	GUI_MUST_HAVE_HWND;
	ShowWindow(mHwnd, SW_MAXIMIZE);
	return OK;
}

FResult GuiType::Restore()
{
	GUI_MUST_HAVE_HWND;
	ShowWindow(mHwnd, SW_RESTORE);
	return OK;
}

FResult GuiType::Flash(optl<BOOL> aBlink)
{
	GUI_MUST_HAVE_HWND;
	FlashWindow(mHwnd, aBlink.value_or(TRUE));
	return OK;
}

FResult GuiType::Hide()
{
	GUI_MUST_HAVE_HWND;
	Cancel();
	return OK;
}



void GuiType::Cancel()
{
	if (mHwnd)
	{
		ShowWindow(mHwnd, SW_HIDE);
		VisibilityChanged(); // This may Release() and indirectly Destroy() the Gui.
	}
	// If this Gui was the last thing keeping the script running, exit the script:
	g_script.ExitIfNotPersistent(EXIT_CLOSE);
}



void GuiType::Close()
// If there is an OnClose event handler defined, launch it as a new thread.
// In this case, don't close or hide the window.  It's up to the handler to do that
// if it wants to.
// If there is no handler, treat it the same as Hide().
{
	if (!IsMonitoring(GUI_EVENT_CLOSE))
		return (void)Cancel();
	POST_AHK_GUI_ACTION(mHwnd, NO_CONTROL_INDEX, GUI_EVENT_CLOSE, NO_EVENT_INFO);
	// MsgSleep() is not done because "case AHK_GUI_ACTION" in GuiWindowProc() takes care of it.
	// See its comments for why.
}



void GuiType::Escape() // Similar to close, except typically called when the user presses ESCAPE.
// If there is an OnEscape event handler defined, launch it as a new thread.
{
	if (!IsMonitoring(GUI_EVENT_ESCAPE)) // The user preference (via votes on forum poll) is to do nothing by default.
		return;
	// See lengthy comments in Event() about this section:
	POST_AHK_GUI_ACTION(mHwnd, NO_CONTROL_INDEX, GUI_EVENT_ESCAPE, NO_EVENT_INFO);
	// MsgSleep() is not done because "case AHK_GUI_ACTION" in GuiWindowProc() takes care of it.
	// See its comments for why.
}



void GuiType::VisibilityChanged()
{
	// Adjust the ref count to reflect the fact that the user has a "reference" to the GUI;
	// that is, the user can interact with it, raising events which may cause callbacks into
	// script.  The script doesn't need any other references to the GUI, since one will be
	// provided in the event callback.
	// In other words, the script is not required to keep a reference to the object while the
	// GUI is visible, but also isn't required to Destroy() it explicitly.  Destroy() will be
	// called automatically when the last reference is released (which may be below).
	bool visible = IsWindowVisible(mHwnd);
	if (visible != mVisibleRefCounted) // Visibility really has changed.
	{
		mVisibleRefCounted = visible; // Change this first in case of recursion.
		if (visible)
			AddRef();
		else
			Release();
	}
}



FResult GuiType::Submit(optl<BOOL> aHideIt, IObject *&aRetVal)
// Caller has ensured that all controls have valid, non-NULL hwnds:
{
	GUI_MUST_HAVE_HWND;

	Object* ret = Object::Create();
	if (!ret)
		return FR_E_OUTOFMEM;

	// Handle all non-radio controls:
	GuiIndexType u;
	for (u = 0; u < mControlCount; ++u)
	{
		GuiControlType& control = *mControl[u]; // For performance and readability.
		if (control.name == NULL) // Skip the control if it does not have a name
			continue;
		
		if (control.HasSubmittableValue())
		{
			TCHAR temp_buf[MAX_NUMBER_SIZE];
			ResultToken value;
			value.InitResult(temp_buf);

			ControlGetContents(value, control, Submit_Mode);
			if (value.Exited())
			{
				ret->Release();
				return FR_FAIL;
			}
			if (!ret->SetOwnProp(control.name, value))
			{
				value.Free();
				goto outofmem;
			}
			value.Free();
		}
	}

	// Handle GUI_CONTROL_RADIO separately so that any radio group that has a single name
	// to share among all its members can be given special treatment:
	int group_radios = 0;           // The number of radio buttons in the current group.
	int group_radios_with_name = 0; // The number of the above that have a name.
	LPTSTR group_name = NULL;       // The last-found name of the current group.
	int selection_number = 0;       // Which radio in the current group is selected (0 if none).
	LPTSTR output_name;

	// The below uses <= so that it goes one beyond the limit.  This allows the final radio group
	// (if any) to be noticed in the case where the very last control in the window is a radio button.
	// This is because in such a case, there is no "terminating control" having the WS_GROUP style:
	for (u = 0; u <= mControlCount; ++u)
	{
		GuiControlType& control = *mControl[u]; // For performance and readability.

		// WS_GROUP is used to determine where one group ends and the next begins -- rather than using
		// seeing if the control's type is radio -- because in the future it may be possible for a radio
		// group to have other controls interspersed within it and yet still be a group for the purpose
		// of "auto radio button (only one selected at a time)" behavior:
		if (u == mControlCount || GetWindowLong(control.hwnd, GWL_STYLE) & WS_GROUP) // New group. Relies on short-circuit boolean order.
		{
			// If the prior group had more than one radio in it but only one had a name, that
			// name is shared among all radios (as of v1.0.20).  Otherwise:
			// 1) If it has zero radios and/or zero names: already fully handled by other logic.
			// 2) It has multiple names: the default values assigned in the loop are simply retained.
			// 3) It has exactly one radio in it and that radio has a name: same as above.
			if (group_radios_with_name == 1 && group_radios > 1)
			{
				// Multiple buttons selected.  Since this is so rare, don't give it a distinct value.
				// Instead, treat this the same as "none selected".  Update for v1.0.24: It is no longer
				// directly possible to have multiple radios selected by having the word "checked" in
				// more than one of their AddCtrl commands.  However, there are probably other ways
				// to get multiple buttons checked (perhaps the Control command), so this handling
				// for multiple selections is left intact.
				if (selection_number == -1)
					selection_number = 0;
				if (!ret->SetOwnProp(group_name, selection_number)) // group_var should not be NULL since group_radios_with_name == 1
					goto outofmem;
			}
			if (u == mControlCount) // The last control in the window is a radio and its group was just processed.
				break;
			group_radios = group_radios_with_name = selection_number = 0;
		}
		if (control.type == GUI_CONTROL_RADIO)
		{
			++group_radios;
			if (output_name = control.name) // Assign.
			{
				++group_radios_with_name;
				group_name = output_name; // If this group winds up having only one name, this will be it.
			}
			int control_is_checked = SendMessage(control.hwnd, BM_GETCHECK, 0, 0) == BST_CHECKED;
			if (control_is_checked)
			{
				if (selection_number) // Multiple buttons selected, so flag this as an indeterminate state.
					selection_number = -1;
				else
					selection_number = group_radios;
			}
			if (output_name)
				if (!ret->SetOwnProp(output_name, control_is_checked))
					goto outofmem;
		}
	} // for()

	if (aHideIt.value_or(TRUE))
		Cancel();
	aRetVal = ret;
	return OK;

outofmem:
	ret->Release();
	return FR_E_OUTOFMEM;
}



void GuiType::ControlGetContents(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode)
{
	// Control types for which Text and Value both have special handling:
	switch (aControl.type)
	{
	case GUI_CONTROL_DROPDOWNLIST: return ControlGetDDL(aResultToken, aControl, aMode);
	case GUI_CONTROL_COMBOBOX: return ControlGetComboBox(aResultToken, aControl, aMode);
	case GUI_CONTROL_LISTBOX: return ControlGetListBox(aResultToken, aControl, aMode);
	case GUI_CONTROL_TAB: return ControlGetTab(aResultToken, aControl, aMode);
	
	case GUI_CONTROL_TEXT: // See case GUI_CONTROL_TEXT in ControlSetContents() for the reasoning behind this.
	case GUI_CONTROL_PIC: // The control's window text is usually the original filename.
		aMode = Text_Mode; // Value and Text both return the window text.
		break;
	}
	// The remaining control types have no need for special handling of the Text property,
	// either because they have normal caption text, or because they have no "text version"
	// of the value, or no "value" at all.
	if (aMode == Text_Mode)
	{
		ControlGetWindowText(aResultToken, aControl);
		_o_return_retval;
	}
	// Otherwise, it's the Value property or Submit method.
	switch (aControl.type)
	{
	case GUI_CONTROL_RADIO: // Same as below.
	case GUI_CONTROL_CHECKBOX: return ControlGetCheck(aResultToken, aControl);
	case GUI_CONTROL_EDIT: return ControlGetEdit(aResultToken, aControl);
	case GUI_CONTROL_DATETIME: return ControlGetDateTime(aResultToken, aControl);
	case GUI_CONTROL_MONTHCAL: return ControlGetMonthCal(aResultToken, aControl);
	case GUI_CONTROL_HOTKEY: return ControlGetHotkey(aResultToken, aControl);
	case GUI_CONTROL_UPDOWN: return ControlGetUpDown(aResultToken, aControl);
	case GUI_CONTROL_SLIDER: return ControlGetSlider(aResultToken, aControl);
	case GUI_CONTROL_PROGRESS: return ControlGetProgress(aResultToken, aControl);

	case GUI_CONTROL_ACTIVEX:
		aControl.union_object->AddRef();
		_o_return(aControl.union_object);
		
	// Already handled in the switch() above:
	//case GUI_CONTROL_DATETIME:
	//case GUI_CONTROL_DROPDOWNLIST:
	//case GUI_CONTROL_COMBOBOX:
	//case GUI_CONTROL_LISTBOX:
	//case GUI_CONTROL_TAB:
	//case GUI_CONTROL_PIC:
	//case GUI_CONTROL_TEXT:

	case GUI_CONTROL_LINK:
	case GUI_CONTROL_GROUPBOX:
	case GUI_CONTROL_BUTTON:
	case GUI_CONTROL_STATUSBAR:
		// Above: The control has no "value".  GuiControl.Text should be used to get its text.
	case GUI_CONTROL_LISTVIEW: // More flexible methods are available.
	case GUI_CONTROL_TREEVIEW: // More flexible methods are available.
	case GUI_CONTROL_CUSTOM: // Don't know what this control's "value" is or how to get it.
	default: // Should help code size by allowing the compiler to omit checks for the above cases.
		// The concept of a "value" does not apply to these controls.  Value could be
		// an alias for Text in these cases, but that doesn't seem helpful/meaningful.
		// This also "reserves" Value for other possible future uses, such as assigning
		// an image to a Button control.
		_o_return_empty;
	}
}



GuiIndexType GuiType::FindControl(LPCTSTR aControlID)
// Find the index of the control that matches the string, which can be either:
// 1) HWND
// 2) Name
// 3) Class+NN
// 4) Control's title/caption.
// Returns -1 if not found.
{
	// v1.0.44.08: Added the following check.  Without it, ControlExist() (further below) would retrieve the
	// topmost child, which isn't very useful or intuitive.  This currently affects only the following commands:
	// 1) GuiControl (but not GuiControlGet because it has special handling for a blank ControlID).
	// 2) Gui, ListView|TreeView, MyTree|MyList
	if (!*aControlID)
		return -1;
	GuiIndexType u;
	UINT hwnd_uint;
	if (ParseInteger(aControlID, hwnd_uint)) // Allow negatives, for flexibility, but truncate to filter out sign-extended values.
	{
		// v1.1.04: Allow Gui controls to be referenced by HWND.  There is some risk of breaking
		// scripts, but only if the text of one control contains the HWND of another control.
		u = FindControlIndex((HWND)(UINT_PTR)hwnd_uint);
		if (u < mControlCount)
			return u;
		// Otherwise: no match was found, so fall back to considering it as text.
	}
	// Try to find the control with the specified name.
	for (u = 0; u < mControlCount; ++u)
		if (mControl[u]->name && _tcsicmp(mControl[u]->name, aControlID) == 0)
			return u;
	// Otherwise: No match found, so fall back to standard control class and/or text finding method.
	HWND control_hwnd = ControlExist(mHwnd, aControlID);
	return FindControlIndex(control_hwnd);
}



int GuiType::FindGroup(GuiIndexType aControlIndex, GuiIndexType &aGroupStart, GuiIndexType &aGroupEnd)
// Caller must provide a valid aControlIndex for an existing control.
// Returns the number of radio buttons inside the group. In addition, it provides start and end
// values to the caller via aGroupStart/End, where Start is the index of the first control in
// the group and End is the index of the control *after* the last control in the group (if none,
// aGroupEnd is set to mControlCount).
// NOTE: This returns the range covering the entire group, and it is possible for the group
// to contain non-radio type controls.  Thus, the caller should check each control inside the
// returned range to make sure it's a radio before operating upon it.
{
	// Work backwards in the control array until the first member of the group is found or the
	// first array index, whichever comes first (the first array index is the top control in the
	// Z-Order and thus treated by the OS as an implicit start of a new group):
	int group_radios = 0; // This and next are both init'd to 0 not 1 because the first loop checks aControlIndex itself.
	for (aGroupStart = aControlIndex;; --aGroupStart)
	{
		if (mControl[aGroupStart]->type == GUI_CONTROL_RADIO)
			++group_radios;
		if (!aGroupStart || GetWindowLong(mControl[aGroupStart]->hwnd, GWL_STYLE) & WS_GROUP)
			break;
	}
	// Now find the control after the last control (or mControlCount if none).  Must start at +1
	// because if aControlIndex's control has the WS_GROUP style, we don't want to count that
	// as the end of the group (because in fact that would be the beginning of the group).
	for (aGroupEnd = aControlIndex + 1; aGroupEnd < mControlCount; ++aGroupEnd)
	{
		// Unlike the previous loop, this one must do this check prior to the next one:
		if (GetWindowLong(mControl[aGroupEnd]->hwnd, GWL_STYLE) & WS_GROUP)
			break;
		if (mControl[aGroupEnd]->type == GUI_CONTROL_RADIO)
			++group_radios;
	}
	return group_radios;
}



FResult GuiType::SetFont(optl<StrArg> aOptions, optl<StrArg> aFontName)
{
	GUI_MUST_HAVE_HWND;
	COLORREF color;
	int font_index = FindOrCreateFont(aOptions.value_or_empty(), aFontName.value_or_empty()
		, &sFont[mCurrentFontIndex], &color);
	if (color != CLR_NONE) // Even if the above call failed, it returns a color if one was specified.
		mCurrentColor = color;
	if (font_index > -1) // Success.
	{
		mCurrentFontIndex = font_index;
		return OK;
	}
	// Failure of the above is rare because it falls back to default typeface if the one specified
	// isn't found.  It will have already displayed the error:
	return FAIL;
}



int GuiType::FindOrCreateFont(LPCTSTR aOptions, LPCTSTR aFontName, FontType *aFoundationFont, COLORREF *aColor)
// Returns the index of existing or new font within the sFont array (an index is returned so that
// caller can see that this is the default-gui-font whenever index 0 is returned).  Returns -1
// on error, but still sets *aColor to be the color name, if any was specified in aOptions.
// To prevent a large number of font handles from being created (such as one for each control
// that uses something other than GUI_DEFAULT_FONT), it seems best to conserve system resources
// by creating new fonts only when called for.  Therefore, this function will first check if
// the specified font already exists within the array of fonts.  If not found, a new font will
// be added to the array.
{
	// Set default output parameter in case of early return:
	if (aColor) // Caller wanted color returned in an output parameter.
		*aColor = CLR_NONE; // Because we want CLR_DEFAULT to indicate a real color.

	if (!*aOptions && !*aFontName)
	{
		// Relies on the fact that first item in the font array is always the default font.
		// If there are fonts, the default font should be the first one (index 0).
		// If not, we create it here:
		if (!sFontCount)
		{
			// sFont is created upon first use to conserve ~20 KB memory in non-GUI scripts.
			if (  !(sFont || (sFont = (FontType *)malloc(sizeof(FontType) * MAX_GUI_FONTS)))  )
				// The allocation is small enough that failure would indicate it's unlikely the
				// program could continue for long, so it doesn't seem worth the added complication
				// to support recovering from this error:
				g_script.CriticalError(ERR_OUTOFMEM);
			// For simplifying other code sections, create an entry in the array for the default font
			// (GUI constructor relies on at least one font existing in the array).
			// Doesn't seem likely that DEFAULT_GUI_FONT face/size will change while a script is running,
			// or even while the system is running for that matter.  I think it's always an 8 or 9 point
			// font regardless of desktop's appearance/theme settings.
			// MSDN: "It is not necessary (but it is not harmful) to delete stock objects by calling DeleteObject."
			sFont[sFontCount].hfont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
			// Get attributes of DEFAULT_GUI_FONT (name, size, etc.)
			FontGetAttributes(sFont[sFontCount]);
			++sFontCount;
		}
		// Tell caller to return to default color, since this is documented behavior when
		// returning to default font:
		if (aColor) // Caller wanted color returned in an output parameter.
			*aColor = CLR_DEFAULT;
		return 0;  // Always returns 0 since that is always the index of the default font.
	}

	// Otherwise, a non-default name/size, etc. is being requested.  Find or create a font to match it.
	if (!aFoundationFont) // Caller didn't specify a font whose attributes should be used as default.
	{
		if (sFontCount > 0)
			aFoundationFont = &sFont[0]; // Use default if it exists.
		else
			return -1; // No error displayed since shouldn't happen if things are designed right.
	}

	// Copy the current default font's attributes into our local font structure.
	// The caller must ensure that mCurrentFontIndex array element exists:
	FontType font = *aFoundationFont;
	if (*aFontName)
		tcslcpy(font.lfFaceName, aFontName, _countof(font.lfFaceName));
	COLORREF color = CLR_NONE; // Because we want to treat CLR_DEFAULT as a real color.

	int point_size = 0;

	TCHAR option[MAX_NUMBER_SIZE + 1];
	LPCTSTR next_option, option_end;
	for (next_option = aOptions; *(next_option = omit_leading_whitespace(next_option)); next_option = option_end)
	{
		// Find the end of this option item:
		for (option_end = next_option; *option_end && !IS_SPACE_OR_TAB(*option_end); ++option_end);

		// Make a copy to simplify comparisons below.
		tcslcpy(option, next_option, min((option_end - next_option) + 1, _countof(option)));

		if (!_tcsicmp(option, _T("Bold")))
			font.lfWeight = FW_BOLD;
		else if (!_tcsicmp(option, _T("italic")))
			font.lfItalic = true;
		else if (!_tcsicmp(option, _T("norm")))
		{
			font.lfItalic = false;
			font.lfUnderline = false;
			font.lfStrikeOut = false;
			font.lfWeight = FW_NORMAL;
		}
		else if (!_tcsicmp(option, _T("underline")))
			font.lfUnderline = true;
		else if (!_tcsicmp(option, _T("strike")))
			font.lfStrikeOut = true;
		else
		{
			// All of the remaining options are single-letter followed by a value:
			TCHAR option_char = ctoupper(*option);
			LPTSTR option_value = option + 1;
			if (!*option_value // Does not have a value.
				|| _tcschr(_T("SWQ"), option_char) // Or is a specific option...
				&& !IsNumeric(option_value, FALSE, FALSE, TRUE)) // and does not have a number.
				goto invalid_option;

			switch (option_char)
			{
			case 'C':
				if (!ColorToBGR(option_value, color))
					goto invalid_option;
				break;
			case 'W': font.lfWeight = ATOI(option_value); break;
			case 'S': point_size = (int)(_tstof(option_value) + 0.5); break; // Round to nearest int.
			case 'Q': font.lfQuality = ATOI(option_value); break;
			default: goto invalid_option;
			}
		}
	}

	if (aColor) // Caller wanted color returned in an output parameter.
		*aColor = color;

	HDC hdc = GetDC(HWND_DESKTOP);
	// Fetch the value every time in case it can change while the system is running (e.g. due to changing
	// display to TV-Out, etc).
	int pixels_per_point_y = GetDeviceCaps(hdc, LOGPIXELSY);
	
	// MulDiv() is usually better because it has automatic rounding, getting the target font
	// closer to the size specified.  This must be done prior to calling FindFont below:
	if (point_size)
		font.lfHeight = -MulDiv(point_size, pixels_per_point_y, 72);

	// The reason it's done this way is that CreateFont() does not always (ever?) fail if given a
	// non-existent typeface:
	if (!FontExist(hdc, font.lfFaceName)) // Fall back to foundation font's type face, as documented.
		_tcscpy(font.lfFaceName, aFoundationFont ? aFoundationFont->lfFaceName : sFont[0].lfFaceName);

	ReleaseDC(HWND_DESKTOP, hdc);

	// Now that the attributes of the requested font are known, see if such a font already
	// exists in the array:
	int font_index = FindFont(font);
	if (font_index != -1) // Match found.
		return font_index;

	// Since above didn't return, create the font if there's room.
	if (sFontCount >= MAX_GUI_FONTS)
	{
		g_script.ScriptError(_T("Too many fonts."));  // Short msg since so rare.
		return -1;
	}
	
	// Set these parameters to defaults to ensure consistent results:
	font.lfWidth = 0;
	font.lfEscapement = 0;
	font.lfOrientation = 0;
	font.lfCharSet = DEFAULT_CHARSET;
	font.lfOutPrecision = OUT_TT_PRECIS;
	font.lfClipPrecision = CLIP_DEFAULT_PRECIS;
	font.lfPitchAndFamily = FF_DONTCARE;
	
	if (  !(font.hfont = CreateFontIndirect(&font))  )
	{
		g_script.ScriptError(_T("Can't create font."));  // Short msg since so rare.
		return -1;
	}

	sFont[sFontCount++] = font; // Copy the newly created font's attributes into the next array element.
	return sFontCount - 1; // The index of the newly created font.

	// This "subroutine" is used to reduce code size:
invalid_option:
	ValueError(ERR_INVALID_OPTION, next_option, FAIL);
	return -1;
}



void GuiType::FontGetAttributes(FontType &aFont)
{
	GetObject(aFont.hfont, sizeof(LOGFONT), &aFont);
}



int GuiType::FindFont(FontType &aFont)
{
	for (int i = 0; i < sFontCount; ++i)
		if (!_tcsicmp(sFont[i].lfFaceName, aFont.lfFaceName) // lstrcmpi() is not used: 1) avoids breaking existing scripts; 2) provides consistent behavior across multiple locales.
			&& sFont[i].lfHeight == aFont.lfHeight // There's no risk of mismatch between cell height (positive) and character height (negative), since we always set it to the latter.
			&& sFont[i].lfWeight == aFont.lfWeight
			&& sFont[i].lfItalic == aFont.lfItalic
			&& sFont[i].lfUnderline == aFont.lfUnderline
			&& sFont[i].lfStrikeOut == aFont.lfStrikeOut
			&& sFont[i].lfQuality == aFont.lfQuality) // Match found.
			return i;
	return -1;  // Indicate failure.
}



LRESULT CALLBACK GuiWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	// If a message pump other than our own is running -- such as that of a dialog like MsgBox -- it will
	// dispatch messages directly here.  This is detected by means of g->CalledByIsDialogMessageOrDispatch==false.
	// Such messages need to be checked here because MsgSleep hasn't seen the message and thus hasn't
	// done the check. The g->CalledByIsDialogMessageOrDispatch method relies on the fact that we never call
	// MsgSleep here for the types of messages dispatched from MsgSleep, which seems true.  Also, if
	// we do launch a monitor thread here via MsgMonitor, that means g->CalledByIsDialogMessageOrDispatch==false.
	// Therefore, any calls to MsgSleep made by the new thread can't corrupt our caller's settings of
	// g->CalledByIsDialogMessageOrDispatch because in that case, our caller isn't MsgSleep's IsDialog/Dispatch.
	// As an added precaution against the complexity of these message issues (only one of several such scenarios
	// is described above), CalledByIsDialogMessageOrDispatch is put into the g-struct rather than being
	// a normal global.  That way, a thread's calls to MsgSleep can't interfere with the value of
	// CalledByIsDialogMessageOrDispatch for any threads beneath it.  Although this may technically be
	// unnecessary, it adds maintainability.
	LRESULT msg_reply;
	if (g->CalledByIsDialogMessageOrDispatch && g->CalledByIsDialogMessageOrDispatchMsg == iMsg)
		g->CalledByIsDialogMessageOrDispatch = false; // Suppress this one message, not any other messages that could be sent due to recursion.
	else if (g_MsgMonitor.Count() && MsgMonitor(hWnd, iMsg, wParam, lParam, NULL, msg_reply)) // Count is checked here to avoid function-call overhead.
		return msg_reply; // MsgMonitor has returned "true", indicating that this message should be omitted from further processing.
	//g->CalledByIsDialogMessageOrDispatch = false; // Now done conditionally above.
	// Fixed for v1.0.40.01: The above line was added to resolve a case where our caller did make the value
	// true but the message it sent us results in a recursive call to us (such as when the user resizes a
	// window by dragging its borders: that apparently starts a loop in DefDlgProc that calls this
	// function recursively).  This fixes OnMessage(0x24, "WM_GETMINMAXINFO") and probably others.
	// Known limitation: If the above launched a thread but the thread didn't cause it to return,
	// and iMsg is something like AHK_GUI_ACTION that will be reposted via PostMessage(), the monitor
	// will be launched again when MsgSleep is called in conjunction with the repost. Given the rarity
	// and the minimal consequences of this, no extra code (such as passing a new parameter to MsgSleep)
	// is added to handle this.

	GuiType *pgui;
	GuiControlType *pcontrol;
	GuiIndexType control_index;
	bool text_color_was_changed;
	TCHAR buf[1024];

	pgui = GuiType::FindGui(hWnd);
	if (!pgui) // This avoids numerous checks below.
		return DefDlgProc(hWnd, iMsg, wParam, lParam);

	if (pgui->mEvents.IsMonitoring(iMsg, GUI_EVENTKIND_MESSAGE))
	{
		ExprTokenType param[] = { pgui, (__int64)wParam, (__int64)(DWORD_PTR)lParam, (__int64)iMsg };
		InitNewThread(0, false, true);
		INT_PTR retval;
		auto result = pgui->mEvents.Call(param, 4, iMsg, GUI_EVENTKIND_MESSAGE, pgui, &retval);
		ResumeUnderlyingThread();
		if (result == EARLY_RETURN)
			return retval;
	}

	switch (iMsg)
	{
	//case WM_CREATE:
		// The DefDlgProc() call above is sufficient.
		// This case wouldn't be executed because FindGui() won't work until after CreateWindowEx() has returned.
		// If it ever becomes necessary, the GuiType pointer could be passed as lpParam of CreateWindowEx().

	case WM_SIZE: // Listed first for performance.
		pgui->mIsMinimized = wParam == SIZE_MINIMIZED; // See "case WM_SETFOCUS" for comments.
		if (pgui->mStatusBarHwnd)
			// Send the msg even if the bar is hidden because the OS typically knows not to do extra drawing work for
			// hidden controls.  In addition, when the bar is shown again, it might be the wrong size if this isn't done.
			// Known/documented limitation: In spite of being in the right z-order position, any control that
			// overlaps the status bar might sometimes get drawn on top of it.
			SendMessage(pgui->mStatusBarHwnd, WM_SIZE, wParam, lParam); // It apparently ignores wParam and lParam, but just in case send it the actuals.
		switch (wParam)
		{
		case SIZE_RESTORED: wParam = 0; break;
		case SIZE_MAXIMIZED: wParam = 1; break;
		case SIZE_MINIMIZED: wParam = -1; break;
		//default:
		// Note that SIZE_MAXSHOW/SIZE_MAXHIDE don't seem to ever be received under the conditions
		// described at MSDN, even if the window has WS_POPUP style.  They are probably relics from
		// 16-bit Windows.  For simplicity (and because the window may have actually been resized)
		// let wParam have its current (unexpected) value, if that's ever possible.
		}
		// probably never contain those values, and as a result they are not documented in the help file.
		if (pgui->IsMonitoring(GUI_EVENT_RESIZE))
			POST_AHK_GUI_ACTION(hWnd, NO_CONTROL_INDEX, MAKEWORD(GUI_EVENT_RESIZE, wParam), lParam);
			// MsgSleep() is not done because "case AHK_GUI_ACTION" in GuiWindowProc() takes care of it.
			// See its comments for why.
		return 0; // "If an application processes this message, it should return zero."
		// Testing shows that the window still resizes correctly (controls are revealed as the window
		// is expanded) even if the event isn't passed on to the default proc.

	case WM_GETMINMAXINFO: // Added for v1.0.44.13.
	{
		MINMAXINFO &mmi = *(LPMINMAXINFO)lParam;
		if (pgui->mMinWidth >= 0) // This check covers both COORD_UNSPECIFIED and COORD_CENTERED.
			mmi.ptMinTrackSize.x = pgui->mMinWidth;
		if (pgui->mMinHeight >= 0)
			mmi.ptMinTrackSize.y = pgui->mMinHeight;
		// mmi.ptMaxSize.x/y aren't changed because it seems the OS automatically uses ptMaxTrackSize
		// for them, at least when ptMaxTrackSize is smaller than the system's default for mmi.ptMaxSize.
		if (pgui->mMaxWidth >= 0)
		{
			mmi.ptMaxTrackSize.x = pgui->mMaxWidth;
			// If the user applies a MaxSize less than MinSize, MinSize appears to take precedence except
			// that the window behaves oddly when dragging the left or top edge outward.  This adjustment
			// removes the odd behaviour but does not visibly affect the size limits.  The user-specified
			// MaxSize and MinSize are retained, so applying -MinSize lets MaxSize take effect.
			if (mmi.ptMaxTrackSize.x < mmi.ptMinTrackSize.x)
				mmi.ptMaxTrackSize.x = mmi.ptMinTrackSize.x;
		}
		if (pgui->mMaxHeight >= 0)
		{
			mmi.ptMaxTrackSize.y = pgui->mMaxHeight;
			if (mmi.ptMaxTrackSize.y < mmi.ptMinTrackSize.y) // See above for comments.
				mmi.ptMaxTrackSize.y = mmi.ptMinTrackSize.y;
		}
		return 0; // "If an application processes this message, it should return zero."
	}

	case WM_COMMAND:
	{
		int id = LOWORD(wParam);
		// For maintainability, this is checked first because "code" (the HIWORD) is sometimes or always 0,
		// which falsely indicates that the message is from a menu:
		if (id == IDCANCEL) // IDCANCEL is a special Control ID.  The user pressed esc.
		{
			// Known limitation:
			// Example:
			//Gui, Add, Text,, Gui1
			//Gui, Add, Text,, Gui2
			//Gui, Show, w333
			//GuiControl, Disable, Gui1
			//return
			//
			//GuiEscape:
			//MsgBox GuiEscape
			//return
			// It appears that in cases like the above, the OS doesn't send the WM_COMMAND+IDCANCEL message
			// to the program when you press Escape. Although it could be fixed by having the escape keystroke
			// unconditionally call the GuiEscape label, that might break existing features and scripts that
			// rely on escape's ability to perform other functions in a window. 
			// I'm not sure whether such functions exist and how many of them there are. Examples might include
			// using escape to close a menu, drop-list, or other pop-up attribute of a control inside the window.
			// Escape also cancels a user's drag-move of the window (restoring the window to its pre-drag location).
			// If pressing escape were to unconditionally call the GuiEscape label, features like these might be
			// broken.  So currently this behavior is documented in the help file as a known limitation.
			pgui->Escape();
			return 0; // Might be necessary to prevent auto-window-close.
			// Note: It is not necessary to check for IDOK because:
			// 1) If there is no default button, the IDOK message is ignored.
			// 2) If there is a default button, we should never receive IDOK because BM_SETSTYLE (sent earlier)
			// will have altered the message we receive to be the ID of the actual default button.
		}
		// Since above didn't return:
		if (id >= ID_USER_FIRST)
		{
			// Since all control id's are less than ID_USER_FIRST, this message is either
			// a user defined menu item ID or a bogus message due to it corresponding to
			// a non-existent menu item or a main/tray menu item (which should never be
			// received or processed here).
			HandleMenuItem(hWnd, id, pgui->mHwnd);
			return 0; // Indicate fully handled.
		}
		// Otherwise id should contain the ID of an actual control.  Validate that in case of bogus msg.
		// Perhaps because this is a DialogProc rather than a WindowProc, the following does not appear
		// to be true: MSDN: "The high-order word [of wParam] specifies the notification code if the message
		// is from a control. If the message is from an accelerator, [high order word] is 1. If the message
		// is from a menu, [high order word] is zero."
		GuiIndexType control_index = GUI_ID_TO_INDEX(id); // Convert from ID to array index. Relies on unsigned to flag as out-of-bounds.
		if (control_index < pgui->mControlCount // Relies on short-circuit boolean order.
			&& pgui->mControl[control_index]->hwnd == (HWND)lParam) // Handles match (this filters out bogus msgs).
			pgui->Event(control_index, HIWORD(wParam), GUI_EVENT_WM_COMMAND);
			// v1.0.35: And now pass it on to DefDlgProc() in case it needs to see certain types of messages.
		break;
	}

	case WM_SYSCOMMAND:
		if (wParam == SC_CLOSE)
		{
			pgui->Close();
			return 0;
		}
		break;

	case WM_NOTIFY:
	{
		NMHDR &nmhdr = *(LPNMHDR)lParam;
		if (!nmhdr.idFrom)
		{
			// This might be the header of a ListView control.  Although there may be some
			// risk of notification codes being ambiguous for custom controls (between the
			// actual control and any child controls), the ability to monitor a ListView's
			// header notifications seems useful enough to justify the risk.
			control_index = pgui->FindControlIndex(nmhdr.hwndFrom);
		}
		else
		{
			control_index = (GuiIndexType)GUI_ID_TO_INDEX(nmhdr.idFrom); // Convert from ID to array index.  Relies on unsigned to flag as out-of-bounds.
		}
		if (control_index >= pgui->mControlCount)
			break;  // Invalid to us, but perhaps meaningful DefDlgProc(), so let it handle it.
		GuiControlType &control = *pgui->mControl[control_index]; // For performance and convenience.
		if (control.hwnd != nmhdr.hwndFrom && nmhdr.idFrom) // Ensure handles match (this filters out bogus msgs).
			break;

		INT_PTR retval;
		if (pgui->ControlWmNotify(control, &nmhdr, retval))
			// A handler was found for this notification and it returned a value.
			return retval;

		UINT_PTR event_info = NO_EVENT_INFO; // Set default, to be possibly overridden below.
		USHORT gui_event = GUI_EVENT_NONE;
		bool explicitly_return_true = false;

		switch (control.type)
		{
		/////////////////////
		// LISTVIEW WM_NOTIFY
		/////////////////////
		case GUI_CONTROL_LISTVIEW:
			switch (nmhdr.code)
			{
			// MSDN: LVN_HOTTRACK: "Return zero to allow the list view to perform its normal track select processing."
			// Also, LVN_HOTTRACK is listed first for performance since it arrives far more often than any other notification.
			case LVN_HOTTRACK:  // v1.0.36.04: No longer an event because it occurs so often: Due to single-thread limit, it was decreasing the reliability of AltSubmit ListViews' receipt of other events such as "I", such as Toralf's Icon Viewer.
			case NM_CUSTOMDRAW: // Return CDRF_DODEFAULT (0).  Occurs for every redraw, such as mouse cursor sliding over control or window activation.
			case LVN_ITEMCHANGING: // Not yet supported (seems rarely needed), so always allow the change by returning 0 (FALSE).
			case LVN_INSERTITEM: // Any ways other than ListView_InsertItem() to insert items?
			case LVN_DELETEITEM: // Might be received for each individual (non-DeleteAll) deletion).
			case LVN_GETINFOTIPW: // v1.0.44: Received even without LVS_EX_INFOTIP?. In any case, there's currently no point
			case LVN_GETINFOTIPA: // in notifying the script because it would have no means of changing the tip (by altering the struct), except perhaps OnMessage.
				return 0; // Return immediately to avoid calling Event() and DefDlgProc(). A return value of 0 is suitable for all of the above.

			case LVN_ITEMCHANGED:
			{
				// This is received for selection/deselection, which means clicking a new item generates
				// at least two of them (in practice, it generates between 1 and 3 but not sure why).
				// It's also received for checking/unchecking an item.  Extending a selection via Shift-ArrowKey
				// generates between 1 and 3 of them, perhaps at random?  Maybe all we can count on is that you
				// get at least one when the selection has changed or a box is (un)checked.
				NMLISTVIEW &lv = *(LPNMLISTVIEW)lParam;
				event_info = 1 + lv.iItem; // MSDN: If iItem is -1, the change has been applied to all items in the list view.

				// Although the OS sometimes generates focus+select together, that's only when the focus is
				// moving and selecting a single item; i.e. not by Ctrl/Shift and the mouse or arrow keys.
				// Also, de-focus and de-select seem to always be separate.
				UINT newly_changed =  lv.uNewState ^ lv.uOldState; // uChanged doesn't seem accurate: it's always 8?  So derive the "correct" value of which flags have actually changed.
				UINT newly_on = newly_changed & lv.uNewState;
				UINT newly_off = newly_changed & lv.uOldState;
				
				// If the focus and selection is changed with Shift and the arrow keys, the control sends the
				// selection change first, then focus change.  So for consistency, report the selection change
				// first when the notification includes both.
				if (newly_changed & LVIS_SELECTED)
					// Pass the new selected state in the second byte of the event code.  Passing 1 or 2 vs. boolean allows shared handling for ITEMSELECT and ITEMCHECK.
					pgui->Event(control_index, 0, GUI_EVENT_ITEMSELECT | ((newly_on & LVIS_SELECTED) ? 0x200 : 0x100), event_info);

				if (newly_on & LVIS_FOCUSED)
					pgui->Event(control_index, 0, GUI_EVENT_ITEMFOCUS, event_info);

				if (newly_changed & LVIS_STATEIMAGEMASK) // State image changed.
				{
					if (lv.uOldState & LVIS_STATEIMAGEMASK) // Image is changing from a non-blank image to a different non-blank image.
					{
						// Pass the new state image index in the second byte of the event code.
						// If the control has checkboxes, these will be 1 (unchecked) or 2 (checked),
						// but doing it this way provides extra utility in case of other state images.
						pgui->Event(control_index, 0, GUI_EVENT_ITEMCHECK | ((lv.uNewState & LVIS_STATEIMAGEMASK) >> 4), event_info);
					}
					//else state image changed from blank/none to some new image.  Omit this event because it seems to do more harm than good in 99% of cases (especially since it typically only occurs when the script calls LV_Add/Insert).
				}
				return 0; // All recognised events for this notification have been handled above.
			}

			case NM_SETFOCUS: gui_event = GUI_EVENT_FOCUS; break;
			case NM_KILLFOCUS: gui_event = GUI_EVENT_LOSEFOCUS; break;

			case LVN_KEYDOWN:
				if (((LPNMLVKEYDOWN)lParam)->wVKey == VK_F2 && !(control.attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR)) // WantF2 is in effect.
				{
					int focused_index = ListView_GetNextItem(control.hwnd, -1, LVNI_FOCUSED);
					if (focused_index != -1)
						SendMessage(control.hwnd, LVM_EDITLABEL, focused_index, 0);  // Has no effect if the control is read-only.
				}
				break;

			// When alt-submit mode isn't in effect, it seems best to ignore all clicks except double-clicks, since
			// right-click should normally be handled via GuiContenxtMenu instead (to allow AppsKey to work, etc.);
			// and since left-clicks can be used to extend a selection (ctrl-click or shift-click), so are pretty
			// vague events that most scripts probably wouldn't have explicit handling for.  A script that needs
			// to know when the selection changes can turn on AltSubmit to catch a wide variety of ways the
			// selection can change, the most all-encompassing of which is probably LVN_ITEMCHANGED.
			case NM_CLICK:
				// v1.0.36.03: For NM_CLICK/NM_RCLICK, it's somewhat debatable to set event_info when the
				// ListView isn't single-select, but the usefulness seems to outweigh any confusion it might cause.
				gui_event = GUI_EVENT_CLICK;
				event_info = 1 + ((LPNMITEMACTIVATE)lParam)->iItem;
				break;
			case NM_RCLICK:
			case NM_RDBLCLK:
				gui_event = GUI_EVENT_CONTEXTMENU;
				event_info = 1 + ((LPNMITEMACTIVATE)lParam)->iItem;
				// Post the event unconditionally rather than calling Event(), so that the Gui's
				// ContextMenu handler can be called even if the control doesn't have one:
				POST_AHK_GUI_ACTION(pgui->mHwnd, control_index, gui_event, event_info);
				return TRUE; // To prevent a second ContextMenu event (which would lack the row number).
			case NM_DBLCLK:
				gui_event = GUI_EVENT_DBLCLK;
				event_info = 1 + ((LPNMITEMACTIVATE)lParam)->iItem;
				break;

			case LVN_COLUMNCLICK:
			{
				gui_event = GUI_EVENT_COLCLK;
				NMLISTVIEW &lv = *(LPNMLISTVIEW)lParam;
				event_info = 1 + lv.iSubItem; // The one-based column number that was clicked.
				// The following must be done here rather than in Event() in case the control has no g-label:
				if (!(control.union_lv_attrib->no_auto_sort)) // Automatic sorting is in effect.
					GuiType::LV_Sort(control, lv.iSubItem, true); // -1 to convert column index back to zero-based.
				break;
			}

			// ItemEdit: It's seems more useful to filter out canceled edits by default, since
			// most scripts probably aren't interested in begin-edit and therefore would not be
			// expecting an end-edit without actual change.  On top of begin-edit being rarely
			// needed, it seems best to require the script to use OnNotify() for it because it
			// might help avoid the case where script makes a change in response to begin-edit
			// but then fails to revert it because the user canceled.  OnNotify() also allows
			// the script to indicate which rows are editable or which edits to accept.
			//case LVN_BEGINLABELEDITW: // Received even for non-Unicode apps, at least on XP.  Even so, the text contained it the struct is apparently always ANSI vs. Unicode.
			//case LVN_BEGINLABELEDITA: // Never received, at least not on XP?
			//	gui_event = 'E';
			//	event_info = 1 + ((NMLVDISPINFO *)lParam)->item.iItem;
			//	break;
			case LVN_ENDLABELEDITW: // See comment above.
			case LVN_ENDLABELEDITA:
				if (!((NMLVDISPINFO *)lParam)->item.pszText) // Edit was cancelled (no change made).
					// Unless the script also subscribes to LVN_BEGINLABELEDIT (unlikely), it's very
					// unlikely that a script will want to know when a user has entered edit mode and
					// then canceled without actually changing the item.
					break;
				gui_event = GUI_EVENT_ITEMEDIT;
				event_info = 1 + ((NMLVDISPINFO *)lParam)->item.iItem;
				// Must return TRUE explicitly because apparently DefDlgProc() would return FALSE,
				// which would cause the change to be rejected:
				explicitly_return_true = true;
				break;

			case LVN_DELETEALLITEMS:
				return TRUE; // For performance, tell it not to notify us as each individual item is deleted.
			} // switch(nmhdr.code).
			break;

		/////////////////////
		// TREEVIEW WM_NOTIFY
		/////////////////////
		case GUI_CONTROL_TREEVIEW:
			switch (nmhdr.code)
			{
			case NM_SETCURSOR:  // Received very often, every time the mouse moves while over the control.
			case NM_CUSTOMDRAW: // Return CDRF_DODEFAULT (0). Occurs for every redraw, such as mouse cursor sliding over control or window activation.
			case TVN_DELETEITEMW:
			case TVN_DELETEITEMA:
			// TVN_SELCHANGING, TVN_ITEMEXPANDING, and TVN_SINGLEEXPAND are not reported to the script as events
			// because there is currently no support for vetoing the selection-change or expansion; plus these
			// notifications each have an "-ED" counterpart notification that is reported to the script (even
			// TVN_SINGLEEXPAND is followed by a TVN_ITEMEXPANDED notification).
			case TVN_SELCHANGINGW:   // Received even for non-Unicode apps, at least on XP.
			case TVN_SELCHANGINGA:
			case TVN_ITEMEXPANDINGW: // Received even for non-Unicode apps, at least on XP.
			case TVN_ITEMEXPANDINGA:
			case TVN_SINGLEEXPAND: // Note that TVNRET_DEFAULT==0. This is received only when style contains TVS_SINGLEEXPAND.
			case TVN_GETINFOTIPA: // Received when TVS_INFOTIP is present. However, there's currently no point
			case TVN_GETINFOTIPW: // in notifying the script because it would have no means of changing the tip (by altering the struct), except perhaps OnMessage.
				return 0; // Return immediately to avoid calling Event() and DefDlgProc(). A return value of 0 is suitable for all of the above.

			case TVN_SELCHANGEDW:
			case TVN_SELCHANGEDA:
				// TreeViews are always single-select, so unlike ListView, each event always represents one
				// selection and (if there was a selection before) one deselection.  It seems better to omit
				// the "selected?" parameter altogether than to always pass 1, so TreeView's ItemSelect is
				// like ListView's ItemFocus.  Although focus and selection are the same for a TreeView,
				// it's probably best to stick with "select" because it's a more common term than "focus"
				// within the context of TreeView controls, and also for consistency with notifications/msgs
				// documented on MSDN (in case the user needs them).
				gui_event = GUI_EVENT_ITEMSELECT;
				event_info = (UINT_PTR)((LPNMTREEVIEW)lParam)->itemNew.hItem;
				break;

			case TVN_ITEMEXPANDEDW: // Received even for non-Unicode apps, at least on XP.
			case TVN_ITEMEXPANDEDA:
				// The "action" flag is a bitwise value that should always contain either TVE_COLLAPSE or
				// TVE_EXPAND (testing shows that TVE_TOGGLE never occurs, as expected).
				gui_event = GUI_EVENT_ITEMEXPAND | ((((LPNMTREEVIEW)lParam)->action & TVE_EXPAND) ? 0x200 : 0x100);
				// It is especially important to store the HTREEITEM of this event for the TVS_SINGLEEXPAND style
				// because an item that wasn't even clicked on is collapsed to allow the new one to expand.
				// There might be no way to find out which item collapsed other than taking note of it here.
				event_info = (UINT_PTR)((LPNMTREEVIEW)lParam)->itemNew.hItem;
				break;

			case TVN_ITEMCHANGEDW:
			case TVN_ITEMCHANGEDA:
			{
				NMTVITEMCHANGE &tv = *(NMTVITEMCHANGE*)lParam;
				if (tv.uChanged == TVIF_STATE && ((tv.uStateOld ^ tv.uStateNew) & TVIS_STATEIMAGEMASK)) // State image changed.
				{
					// Pass the new state image index in the second byte of the event code.
					// If the control has checkboxes, these will be 1 (unchecked) or 2 (checked),
					// but doing it this way provides extra utility in case of other state images.
					gui_event = GUI_EVENT_ITEMCHECK | ((tv.uStateNew & TVIS_STATEIMAGEMASK) >> 4);
					event_info = (UINT)(size_t)tv.hItem;
				}
				break;
			}

			// ItemEdit: See comments in ListView section.
			case TVN_BEGINLABELEDITW: // Received even for non-Unicode apps, at least on XP.  Even so, the text contained it the struct is apparently always ANSI vs. Unicode.
			case TVN_BEGINLABELEDITA: // Never received, at least not on XP?
				//gui_event = 'E';
				//event_info = (UINT_PTR)((LPNMTVDISPINFO)lParam)->item.hItem;
				GuiType::sTreeWithEditInProgress = control.hwnd;
				break;
			case TVN_ENDLABELEDITW: // See comment above.
			case TVN_ENDLABELEDITA:
				if (!((LPNMTVDISPINFO)lParam)->item.pszText) // Edit was cancelled (no change made).
					// Unless the script also subscribes to TVN_BEGINLABELEDIT (unlikely), it's very
					// unlikely that a script will want to know when a user has entered edit mode and
					// then canceled without actually changing the item.
					break;
				gui_event = GUI_EVENT_ITEMEDIT;
				event_info = (UINT_PTR)((LPNMTVDISPINFO)lParam)->item.hItem;
				GuiType::sTreeWithEditInProgress = NULL;
				// Must return TRUE explicitly because apparently DefDlgProc() would return FALSE,
				// which would cause the change to be rejected:
				explicitly_return_true = true;
				break;

			// Since a left-click is just one method of changing selection (keyboard navigation is another),
			// it seems desirable for performance not to report such clicks except in alt-submit mode.
			// Similarly, right-clicks are reported only in alt-submit mode because GuiContextMenu should be used
			// to catch right-clicks (due to its additional handling for the AppsKey).
			case NM_CLICK:
			case NM_DBLCLK:
			//case NM_RCLICK: // Not handled because the default processing generates WM_CONTEXTMENU, which is what we want.
			//case NM_RDBLCLK: // As above, but NM_RDBLCLK is never actually generated by a TreeView (at least in some OS versions).
				switch(nmhdr.code)
				{
				case NM_CLICK: gui_event = GUI_EVENT_CLICK; break;
				case NM_DBLCLK: gui_event = GUI_EVENT_DBLCLK; break;
				}
				// Since testing shows that none of the NMHDR members contains the HTREEITEM, must use
				// another method to discover it for the various mouse-click events.
				TVHITTESTINFO ht;
				// GetMessagePos() is used because it should be more accurate than GetCursorPos() in case a
				// the message was in the queue a long time.  There is some concern due to GetMessagePos() being
				// documented to be valid only for GetMessage(): there's no certainty that all message pumps
				// (such as that of MsgBox) use GetMessage vs. PeekMessage, but its hard to imagine that
				// GetMessagePos() doesn't work for PeekMessage().  In any case, all message pumps by built-in
				// OS dialogs like MsgBox probably use GetMessage().  There's another concern: that this WM_NOTIFY
				// msg was sent (vs. posted) from our own thread somehow, in which case it never got queued so
				// GetMessagePos() might yield an inaccurate value.  But until that is proven to be an actual
				// problem, it seems best to do it the "correct" way.
				DWORD pos;
				pos = GetMessagePos();
				ht.pt.x = (short)LOWORD(pos);
				ht.pt.y = (short)HIWORD(pos);
				ScreenToClient(control.hwnd, &ht.pt);
				event_info = (UINT_PTR)TreeView_HitTest(control.hwnd, &ht);
				break;

			case NM_SETFOCUS: gui_event = GUI_EVENT_FOCUS; break;
			case NM_KILLFOCUS: gui_event = GUI_EVENT_LOSEFOCUS; break;

			case TVN_KEYDOWN:
				if (((LPNMTVKEYDOWN)lParam)->wVKey == VK_F2 && !(control.attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR)) // WantF2 is in effect.
				{
					HTREEITEM hitem;
					if (hitem = TreeView_GetSelection(control.hwnd))
						SendMessage(control.hwnd, TVM_EDITLABEL, 0, (LPARAM)hitem); // Has no effect if the control is read-only.
				}
				break;
			} // switch(nmhdr.code).
			break;

		//////////////////////
		// OTHER CONTROL TYPES
		//////////////////////
		case GUI_CONTROL_DATETIME: // NMDATETIMECHANGE struct contains an NMHDR as it's first member.
			// Both MonthCal's year spinner (when year is clicked on) and DateTime's drop-down calendar
			// seem to start a new message pump.  This is one of the reason things were redesigned to
			// avoid doing a MsgSleep(-1) after posting AHK_GUI_ACTION at the bottom of Event().
			// See its comments for details.
			switch (nmhdr.code)
			{
			case DTN_DATETIMECHANGE: gui_event = GUI_EVENT_CHANGE; break;
			case NM_SETFOCUS: gui_event = GUI_EVENT_FOCUS; break;
			case NM_KILLFOCUS: gui_event = GUI_EVENT_LOSEFOCUS; break;
			}
			break;

		case GUI_CONTROL_MONTHCAL:
			// Although the NMSELCHANGE notification struct contains the control's current date/time,
			// it simplifies the code to fetch it again (performance is probably good since the control
			// almost certainly just passes back a pointer to its self-maintained structs).
			// MCN_SELECT isn't used because it isn't sent when the user scrolls to a new month via the
			// calendar's arrow buttons, even though doing so sets a new date inside the control, or when
			// a new year or month is chosen by clicking directly on the month or year.  MCN_SELCHANGE
			// seems to be superior in every way.
			if (nmhdr.code == MCN_SELCHANGE)
				gui_event = GUI_EVENT_CHANGE;
			break;

		case GUI_CONTROL_UPDOWN:
			// Now it just returns 0 for simplicity, but the following are kept for reference.
			//if (nmhdr.code == UDN_DELTAPOS) // No script control/intervention over this currently.
			//	return 0; // MSDN: "Return nonzero to prevent the change in the control's position, or zero to allow the change."
			// Strangely, NM_RELEASEDCAPTURE never seems to be received.  In fact, nothing other than
			// UDN_DELTAPOS is ever received.  Therefore, WM_VSCROLL/WM_HSCROLL are relied upon instead.
			return 0;  // 0 is appropriate for all notifications in this case (no need to let DefDlgProc handle it).

		case GUI_CONTROL_TAB:
			if (nmhdr.code == TCN_SELCHANGE)
			{
				// For code reduction and simplicity (and due to rarity of script needing it), A_EventInfo
				// is not set to the newly selected tab name (or number in the case of AltSubmit).
				pgui->ControlUpdateCurrentTab(control, true);
				gui_event = GUI_EVENT_CHANGE;
			}
			break;

		case GUI_CONTROL_LINK:
			if (nmhdr.code == NM_CLICK || nmhdr.code == NM_RETURN)
			{
				NMLINK &nmLink = *(PNMLINK)lParam;
				LITEM &item = nmLink.item;
				// If (and only if) the script is not monitoring the Click event,
				// attempt to execute the link's HREF attribute:
				if (*item.szUrl && !control.events.IsMonitoring(GUI_EVENT_CLICK))
				{
					g_script.ActionExec(item.szUrl, NULL, NULL, false);
					return 0;
				}
				gui_event = GUI_EVENT_CLICK;
				event_info = item.iLink + 1; // Convert to 1-based.
			}
			break;

		case GUI_CONTROL_STATUSBAR:
			switch(nmhdr.code)
			{
			case NM_CLICK:
				gui_event = GUI_EVENT_CLICK;
				// Pass the one-based part number that was clicked.  If the user clicked near the size grip,
				// apparently a large number is returned (at least on some OSes).
				event_info = ((LPNMMOUSE)lParam)->dwItemSpec + 1;
				break;
			case NM_DBLCLK:
				gui_event = GUI_EVENT_DBLCLK;
				event_info = ((LPNMMOUSE)lParam)->dwItemSpec + 1;
				break;
			case NM_RCLICK:
			case NM_RDBLCLK:
				// Handling it here allows the ContextMenu event to receive the SB part number.
				gui_event = GUI_EVENT_CONTEXTMENU;
				event_info = ((LPNMMOUSE)lParam)->dwItemSpec + 1;
				// Post the event unconditionally rather than calling Event(), so that the Gui's
				// ContextMenu handler can be called even if the control doesn't have one:
				POST_AHK_GUI_ACTION(pgui->mHwnd, control_index, gui_event, event_info);
				return TRUE; // To prevent a second ContextMenu event (which would lack the part number).
			} // switch(nmhdr.code)
		} // switch(control.type) within case WM_NOTIFY.

		// Raise any event identified above.
		if (gui_event)
		{
			pgui->Event(control_index, nmhdr.code, gui_event, event_info);
			if (explicitly_return_true)
				return TRUE;
		}

		break; // outermost switch()
	}

	case WM_VSCROLL: // These two should only be received for sliders and up-downs.
	case WM_HSCROLL:
		pgui->Event(GUI_HWND_TO_INDEX((HWND)lParam), LOWORD(wParam), GUI_EVENT_NONE, HIWORD(wParam));
		return 0; // "If an application processes this message, it should return zero."
	
	//case WM_ERASEBKGND:
	//	if (!pgui->mBackgroundBrushWin) // Let default proc handle it.
	//		break;
	//	// Can't use SetBkColor(), need an real brush to fill it.
	//	GetClipBox((HDC)wParam, &rect);
	//	FillRect((HDC)wParam, &rect, pgui->mBackgroundBrushWin);
	//	return 1; // "An application should return nonzero if it erases the background."

	// The below seems to be the equivalent of the above (but MSDN indicates it will only work
	// if there is no WM_ERASEBKGND handler).  It might perform a little better.
	case WM_CTLCOLORDLG:
		if (!pgui->mBackgroundBrushWin) // Let default proc handle it.
			break;
		SetBkColor((HDC)wParam, pgui->mBackgroundColorWin);
		return (LRESULT)pgui->mBackgroundBrushWin;

	// It seems that scrollbars belong to controls (such as Edit and ListBox) do not send us
	// WM_CTLCOLORSCROLLBAR (unlike the static messages we receive for radio and checkbox).
	// Therefore, this section is commented out since it has no effect (it might be useful
	// if a control's class window-proc is ever overridden with a new proc):
	//case WM_CTLCOLORSCROLLBAR:
	//	if (pgui->mBackgroundBrushWin)
	//	{
	//		// Since we're processing this msg rather than passing it on to the default proc, must set
	//		// background color unconditionally, otherwise plain white will likely be used:
	//		SetTextColor((HDC)wParam, pgui->mBackgroundColorWin);
	//		SetBkColor((HDC)wParam, pgui->mBackgroundColorWin);
	//		// Always return a real HBRUSH so that Windows knows we altered the HDC for it to use:
	//		return (LRESULT)pgui->mBackgroundBrushWin;
	//	}
	//	break;

	case WM_CTLCOLORSTATIC:
	case WM_CTLCOLORLISTBOX:
	case WM_CTLCOLOREDIT:
	case WM_CTLCOLORBTN:
		// MSDN: Buttons with the BS_PUSHBUTTON, BS_DEFPUSHBUTTON, or BS_PUSHLIKE styles do not use the
		// returned brush. Buttons with these styles are always drawn with the default system colors.
		// This is because "drawing push buttons requires several different brushes-face, highlight and
		// shadow". In short, to provide a custom appearance for push buttons, use an owner-drawn button.
		// Thus, WM_CTLCOLORBTN not handled here because it doesn't seem to have any effect on the
		// types of buttons used so far.  This has been confirmed: Even when a theme is in effect,
		// checkboxes, radios, and groupboxes do not receive WM_CTLCOLORBTN, but they do receive
		// WM_CTLCOLORSTATIC.
		if (   !(pcontrol = pgui->FindControl((HWND)lParam))   )
			break;
		if (text_color_was_changed = (pcontrol->type != GUI_CONTROL_PIC && pcontrol->union_color != CLR_DEFAULT))
			SetTextColor((HDC)wParam, pcontrol->union_color);
		else
			// Always set a text color in case a brush will be returned below.  If this isn't done and
			// a brush is returned, the system will use whatever color was set before; probably black.
			SetTextColor((HDC)wParam, GetSysColor(COLOR_WINDOWTEXT));

		if (pcontrol->background_color == CLR_TRANSPARENT)
		{
			SetBkMode((HDC)wParam, TRANSPARENT);
			return (LRESULT)GetStockObject(NULL_BRUSH);
		}

		HBRUSH bk_brush;
		COLORREF bk_color;
		bool use_bg_color;
		// Use mBackgroundColorWin only for specific controls that have no apparent frame.
		// WM_CTLCOLORSTATIC can be received for Edit, DDL/ComboBox and TreeView (-Theme)
		// controls when they are disabled, but we use system default colors (or the
		// control's own background) in those cases for the following reasons:
		//  - Consistency with ListBox, Hotkey, DateTime, MonthCal and ListView.
		//  - TreeView looks bad (item backgrounds don't match the control).
		//  - To match the control's text, which usually takes on the "disabled" color.
		use_bg_color = iMsg == WM_CTLCOLORSTATIC && pcontrol->UsesGuiBgColor();
		pgui->ControlGetBkColor(*pcontrol, use_bg_color, bk_brush, bk_color);
		
		if (bk_brush)
		{
			// Since we're processing this msg rather than passing it on to the default proc, must set
			// background color unconditionally, otherwise plain white will likely be used:
			SetBkColor((HDC)wParam, bk_color);
			// Always return a real HBRUSH so that Windows knows we altered the HDC for it to use:
			return (LRESULT)bk_brush;
		}

		if (!text_color_was_changed) // No need to return a brush since no changes are needed.  Let def. proc. handle it.
			break;
		if (iMsg == WM_CTLCOLORSTATIC)
		{
			SetBkColor((HDC)wParam, GetSysColor(COLOR_BTNFACE)); // Use default window color for static controls.
			return (LRESULT)GetSysColorBrush(COLOR_BTNFACE);
		}
		else
		{
			SetBkColor((HDC)wParam, GetSysColor(COLOR_WINDOW)); // Use default control-bkgnd color for others.
			return (LRESULT)GetSysColorBrush(COLOR_WINDOW);
		}
		break;


	case WM_DRAWITEM:
	{
		// WM_DRAWITEM msg is received by GUI windows that contain a tab control with custom tab
		// colors.  The TCS_OWNERDRAWFIXED style is what causes this message to be received.
		LPDRAWITEMSTRUCT lpdis = (LPDRAWITEMSTRUCT)lParam;
		// Otherwise, it might be a tab control with custom tab colors:
		control_index = (GuiIndexType)GUI_ID_TO_INDEX(lpdis->CtlID); // Convert from ID to array index. Relies on unsigned to flag as out-of-bounds.
		if (control_index >= pgui->mControlCount // Relies on short-circuit eval order.
			|| pgui->mControl[control_index]->hwnd != lpdis->hwndItem  // Handles do not match (this filters out bogus msgs).
			|| pgui->mControl[control_index]->type != GUI_CONTROL_TAB) // In case this msg can be received for other types.
			break;
		GuiControlType &control = *pgui->mControl[control_index]; // For performance & convenience.
		HBRUSH bk_brush;
		COLORREF bk_color;
		pgui->ControlGetBkColor(control, true, bk_brush, bk_color);
		if (bk_brush)
		{
			FillRect(lpdis->hDC, &lpdis->rcItem, bk_brush); // Fill the tab itself.
			SetBkColor(lpdis->hDC, bk_color); // Set the text's background color.
		}
		else // Must do this anyway, otherwise there is an unwanted thin white line and possibly other problems.
			FillRect(lpdis->hDC, &lpdis->rcItem, (HBRUSH)(size_t)GetClassLongPtr(control.hwnd, GCLP_HBRBACKGROUND));
		// else leave background colors to default, in the case where only the text itself has a custom color.
		// Get the stored name/caption of this tab:
		TCITEM tci;
		tci.mask = TCIF_TEXT;
		tci.pszText = buf;
		tci.cchTextMax = _countof(buf) - 1; // MSDN example uses -1.
		// Set text color if needed:
        COLORREF prev_color = CLR_INVALID;
		if (control.union_color != CLR_DEFAULT)
			prev_color = SetTextColor(lpdis->hDC, control.union_color);
		// Draw the text.  Note that rcItem contains the dimensions of a tab that has already been sized
		// to handle the amount of text in the tab at the specified WM_SETFONT font size, which makes
		// this much easier.
		if (TabCtrl_GetItem(lpdis->hwndItem, lpdis->itemID, &tci))
		{
			// The text is centered horizontally and vertically because that seems to be how the
			// control acts without the TCS_OWNERDRAWFIXED style.  DT_NOPREFIX is not specified
			// because that is not how the control acts without the TCS_OWNERDRAWFIXED style
			// (ampersands do cause underlined letters, even though they currently have no effect).
			if (TabCtrl_GetCurSel(control.hwnd) != lpdis->itemID)
				lpdis->rcItem.top += 3; // For some reason, the non-current tabs' rects are a little off.
			DrawText(lpdis->hDC, tci.pszText, (int)_tcslen(tci.pszText), &lpdis->rcItem
				, DT_CENTER|DT_VCENTER|DT_SINGLELINE); // DT_VCENTER requires DT_SINGLELINE.
			// Cruder method, probably not always accurate depending on theme/display-settings/etc.:
			//TextOut(lpdis->hDC, lpdis->rcItem.left + 5, lpdis->rcItem.top + 3, tci.pszText, (int)_tcslen(tci.pszText));
		}
		if (prev_color != CLR_INVALID) // Put the previous color back into effect for this DC.
			SetTextColor(lpdis->hDC, prev_color);
		break;
	}

	case WM_SETFOCUS:
		// DefDlgProc "Sets the input focus to the control identified by a previously saved control window
		// handle. If no such handle exists, the procedure sets the input focus to the first control ..."
	case WM_ACTIVATE:
		// DefDlgProc "Restores the input focus to the control identified by the previously saved handle
		// if the dialog box is activated."
		// "If the window is being activated and is not minimized, the DefWindowProc function sets the
		// keyboard focus to the window."
		if (pgui->mIsMinimized)
			return 0;
		// The simple check above solves a problem caused by bad OS behaviour that I couldn't find any
		// mention of in all my searching, and that took hours to debug; so I'll be verbose:
		//
		// Without the suppression check, restoring a minimized GUI window (without activating it first)
		// causes focus to reset to the first control in the GUI, instead of the control that had focus
		// previously.  The sequence of events is as follows:
		//  - WM_ACTIVATE, WA_ACTIVE (window is active and already NOT MINIMIZED according to wParam)
		//    DefDlgProc restores focus to the correct control.
		//  - WM_SIZE, SIZE_RESTORED (now we're told the window has been restored)
		//  - WM_SETFOCUS (the control has lost focus and the window itself has gained it)
		//    DefDlgProc resets focus to the window's FIRST control.
		//  - WM_ACTIVATE, WA_ACTIVE (again)
		//
		// My conclusion after debugging with Spy++, logging breakpoints and OnMessage, was that these
		// WM_ACTIVATE and WM_SETFOCUS messages are directly caused by whatever system function restores
		// the window.  If only the first WM_ACTIVATE message is suppressed, the following occurs:
		//  - WM_ACTIVATE, WA_ACTIVE (we don't call DefDlgProc)
		//  - WM_SETFOCUS (this is sent after WM_ACTIVATE returns)
		//    DefDlgProc restores focus to the correct control.
		//  - WM_SIZE, SIZE_RESTORED
		//  - WM_SETFOCUS
		//    DefDlgProc resets focus to the window's FIRST control (as before).
		//  - WM_ACTIVATE, WA_ACTIVE (again)
		//
		// So apparently if the window doesn't set focus in response to WM_ACTIVATE, the system calls
		// SetFocus(hWnd).  But even though it does that, it calls SetFocus(hWnd) again afterward!
		//
		// Another attempted workaround was to have WM_SETFOCUS do SetFocus(wParam) if wParam was this
		// window's own control.  This worked, but the focus changed several times each time the window
		// was restored, and for some controls this generates redundant focus and lost-focus events.
		//
		// To solve this we suppress both WM_ACTIVATE and WM_SETFOCUS in between WM_SIZE, SIZE_MINIMIZED
		// and the next WM_SIZE with any other parameter value.  The end result is that when the window
		// is activated, the window itself is focused (which normally happens, temporarily) and then
		// focus changes only once, when WM_ACTIVATE is received after WM_SIZE.  There is no second
		// WM_SETFOCUS, because hWnd already has the focus.
		//
		// WM_SIZE is tracked with mIsMinimized rather than using IsIconic(hWnd) because the iconic state
		// of the window changes before the transition is completed (which is probably why WM_ACTIVATE's
		// wParam doesn't indicate that the window is minimized).
		//
		// Note that this issue didn't occur in at least the following cases:
		//  1) If the window is activated before restoring it, WM_ACTIVATE is received while the window
		//     is definitely in the minimized state (and not in transition), therefore DefDlgProc does
		//     not handle it by changing the focus.
		//  2) If the window is restored via the taskbar button, it loses focus when the user clicks the
		//     taskbar, so see above.
		//  3) If the window is minimized via the taskbar button, it can be restored even without first
		//     activating it.  This seems to be a consequence of some shonky handling by the system, as
		//     it sends WM_ACTIVATE, 0x00200001 (WA_ACTIVE with upper word non-zero, meaning minimized),
		//     even though the window IS NOT ACTIVE, and then when the window is activated, WM_ACTIVATE
		//     is not sent (and without that message, there is no premature focus restoration).
		//     I came across several mentions of this issue with WM_ACTIVATE and the taskbar in my searches.
		//     Note that sending the bogus message manually produces a different result; i.e. WM_ACTIVATE
		//     is still sent by the system when the window is restored.
		break;

	case WM_CONTEXTMENU:
		{
			HWND clicked_hwnd = (HWND)wParam;
			bool from_keyboard; // Whether Context Menu was generated from keyboard (AppsKey or Shift-F10).
			if (   !(from_keyboard = (lParam == -1))   ) // Mouse click vs. keyboard event.
			{
				// If the click occurred above the client area, assume it was in title/menu bar or border.
				// Let default proc handle it.
				point_and_hwnd_type pah = {0};
				pah.pt.x = LOWORD(lParam);
				pah.pt.y = HIWORD(lParam);
				POINT client_pt = pah.pt;
				if (!ScreenToClient(hWnd, &client_pt) || client_pt.y < 0)
					break; // Allows default proc to display standard system context menu for title bar.
				// v1.0.38.01: Recognize clicks on pictures and text controls as occurring in that control
				// (via A_GuiControl) rather than generically in the window:
				if (clicked_hwnd == pgui->mHwnd)
				{
					// v1.0.40.01: Rather than doing "ChildWindowFromPoint(clicked_hwnd, client_pt)" -- which fails to
					// detect text and picture controls (and perhaps others) when they're inside GroupBoxes and
					// Tab controls -- use the MouseGetPos() method, which seems much more accurate.
					EnumChildWindows(clicked_hwnd, EnumChildFindPoint, (LPARAM)&pah); // Find topmost control containing point.
					clicked_hwnd = pah.hwnd_found; // Okay if NULL; the next stage will handle it.
				}
			}
			// Finding control_index requires only GUI_HWND_TO_INDEX (not FindControl) since context menu message
			// never arrives for a ComboBox's Edit control (since that control has its own context menu).
			control_index = GUI_HWND_TO_INDEX(clicked_hwnd); // Yields a small negative value on failure, which due to unsigned is seen as a large positive number.
			if (control_index >= pgui->mControlCount) // The user probably clicked the parent window rather than inside one of its controls.
				control_index = NO_CONTROL_INDEX;
				// Above flags it as a non-control event. Must use NO_CONTROL_INDEX rather than something
				// like 0xFFFFFFFF so that high-order bit is preserved for use below.
			if (pgui->IsMonitoring(GUI_EVENT_CONTEXTMENU)
				|| control_index != NO_CONTROL_INDEX && pgui->mControl[control_index]->events.IsMonitoring(GUI_EVENT_CONTEXTMENU))
				POST_AHK_GUI_ACTION(hWnd, control_index, MAKEWORD(GUI_EVENT_CONTEXTMENU, from_keyboard), NO_EVENT_INFO);
			return 0; // Return value doesn't matter.
		}
		//else it's for some non-GUI window (probably impossible).  Let DefDlgProc() handle it.
		break;

	case WM_DROPFILES:
	{
		HDROP hdrop = (HDROP)wParam;
		if (pgui->mHdrop || !pgui->IsMonitoring(GUI_EVENT_DROPFILES))
		{
			// There is no event handler in the script, or this window is still processing a prior drop.
			// Ignore this drop and free its memory.
			DragFinish(hdrop);
			return 0; // "An application should return zero if it processes this message."
		}
		// Otherwise: Indicate that this window is now processing the drop.  DragFinish() will be called later.
		pgui->mHdrop = hdrop;
		point_and_hwnd_type pah = {0};
		// DragQueryPoint()'s return value is non-zero if the drop occurred in the client area.
		// However, that info seems too rarely needed to justify storing it anywhere:
		DragQueryPoint(hdrop, &pah.pt);
		ClientToScreen(hWnd, &pah.pt); // EnumChildFindPoint() requires screen coords.
		EnumChildWindows(hWnd, EnumChildFindPoint, (LPARAM)&pah); // Find topmost control containing point.
		// Look up the control in case the drop occurred in a child of a child, such as the edit portion
		// of a ComboBox (FindControl will take that into account):
		pcontrol = pah.hwnd_found ? pgui->FindControl(pah.hwnd_found) : NULL;
		// Finding control_index requires only GUI_HWND_TO_INDEX (not FindControl) since EnumChildFindPoint (above)
		// already properly resolves the Edit control of a ComboBox to be the ComboBox itself.
		control_index = pcontrol ? GUI_HWND_TO_INDEX(pcontrol->hwnd) : NO_CONTROL_INDEX;
		// Above: NO_CONTROL_INDEX indicates to GetGuiControl() that there is no control in this case.
		POST_AHK_GUI_ACTION(hWnd, control_index, GUI_EVENT_DROPFILES, NO_EVENT_INFO); // The HDROP is not passed via message so that it can be released (via the destructor) if the program closes during the drop operation.
		// MsgSleep() is not done because "case AHK_GUI_ACTION" in GuiWindowProc() takes care of it.
		// See its comments for why.
		return 0; // "An application should return zero if it processes this message."
	}

	case AHK_GUI_ACTION:
	case AHK_USER_MENU:
		// v1.0.36.03: The g_MenuIsVisible check was added as a means to discard the message. Otherwise
		// MSG_FILTER_MAX would result in a bouncing effect or something else that disrupts a popup menu,
		// namely a context menu shown by an AltSubmit ListView (regardless of whether it's shown by
		// GuiContextMenu or in response to a RightClick event).  This is because a ListView apparently
		// generates the following notifications while the context menu is displayed:
		// C: release mouse capture
		// H: hottrack
		// f: lost focus
		// I think the issue here is that there are times when messages should be reposted and
		// other times when they should not be.  The MonthCal and DateTime cases mentioned below are
		// times when they should, because the MSG_FILTER_MAX filter is not in effect then.  But when a
		// script's own popup menu is displayed, any message subject to filtering (which includes
		// AHK_GUI_ACTION and AHK_USER_MENU) should probably never be reposted because that disrupts
		// the ability to select an item in the menu or dismiss it (possibly due to the theoretical
		// bouncing-around effect described below).
		// OLDER: I don't think the below is a complete explanation since it doesn't take into account
		// the fact that the message filter might be in effect due to a menu being visible, which if
		// true would prevent MsgSleep from processing the message.
		// OLDEST: MsgSleep() is the critical step.  It forces our thread msg pump to handle the message now
		// because otherwise it would probably become a CPU-maxing loop wherein the dialog or MonthCal
		// msg pump that called us dispatches the above message right back to us, causing it to
		// bounce around thousands of times until that other msg pump finally finishes.
		// UPDATE: This will backfire if the script is uninterruptible.
		if (IsInterruptible())
		{
			// Handling these messages here by reposting them to our thread relieves the one who posted them
			// from ever having to do a MsgSleep(-1), which in turn allows it or its caller to acknowledge
			// its message in a timely fashion, which in turn prevents undesirable side-effects when a
			// g-labeled DateTime's drop-down is navigated via its arrow buttons (jumps ahead two months
			// instead of one, infinite loop with mouse button stuck down on some systems, etc.). Another
			// side-effect is the failure of a g-labeled MonthCal to be able to notify of date change when
			// the user clicks the year and uses the spinner to select a new year.  This solves both of
			// those issues and almost certainly others:
			PostMessage(hWnd, iMsg, wParam, lParam);
			MsgSleep(-1, RETURN_AFTER_MESSAGES_SPECIAL_FILTER);
		}
		return 0;

	case WM_CLOSE: // For now, take the same action as SC_CLOSE.
		pgui->Close();
		return 0;

	case WM_DESTROY:
		// Update to below: If a GUI window is owned by the script's main window (via "gui +owner"),
		// it can be destroyed automatically.  Because of this and the fact that it's difficult to
		// keep track of all the ways a window can be destroyed, it seems best for peace-of-mind to
		// have this WM_DESTROY handler 
		// Older Note: Let default-proc handle WM_DESTROY because with the current design, it should
		// be impossible for a window to be destroyed without the object "knowing about it" and
		// updating itself (then destroying itself) accordingly.  The object methods always
		// destroy (recursively) any windows it owns, so once again it "knows about it".
		if (!pgui->mDestroyWindowHasBeenCalled)
		{
			pgui->mDestroyWindowHasBeenCalled = true; // Tell it not to call DestroyWindow(), just clean up everything else.
			pgui->Destroy();
		}
		// Above: if mDestroyWindowHasBeenCalled==true, we were called by Destroy(), so don't call Destroy() again recursively.
		// And in any case, pass it on to DefDlgProc() in case it does any extra cleanup:
		break;

	// For WM_ENTERMENULOOP/WM_EXITMENULOOP, there is similar code in MainWindowProc(), so maintain them together.
	// g_MenuIsVisible is tracked for the purpose of preventing thread interruptions while the program
	// is in a modal menu loop, since the menu would become unresponsive until the interruption returns.
	// WM_ENTERMENULOOP is received both when entering a modal menu loop (due to a menu bar or popup menu)
	// and when a modeless menu is being displayed (in which case there's actually no menu loop).
	// WM_INITMENUPOPUP is handled to verify which situation this is, so that new threads can launch
	// while a modeless menu is being displayed.  We use this combination of messages rather than just
	// the INIT messages because some scripts already suppress WM_ENTERMENULOOP to allow new threads.
	// "break" is used rather than "return 0" to let DefWindowProc()/DefDlgProc() take whatever action
	// it needs to do for these.
	case WM_ENTERMENULOOP:
		g_MenuIsVisible = true;
		break;
	case WM_EXITMENULOOP:
		g_MenuIsVisible = false; // See comments above.
		break;
	case WM_INITMENUPOPUP:
		InitMenuPopup((HMENU)wParam);
		break;
	case WM_UNINITMENUPOPUP:
		// This currently isn't needed in GuiWindowProc() because notifications for temp-modeless
		// menus are always sent to g_hWnd, but it is kept here for symmetry/maintainability:
		UninitMenuPopup((HMENU)wParam);
		break;
	case WM_NCACTIVATE:
		if (!wParam && g_MenuIsTempModeless)
		{
			// The documentation is quite ambiguous, but it appears the return value of WM_NCACTIVATE
			// affects whether the actual change of foreground window goes ahead, whereas the appearance
			// of the nonclient area is controlled by whether or not DefDlgProc() is called.
			return TRUE; // Allow change of foreground window, but appear to remain active.
		}
		break;
	} // switch()
	
	// This will handle anything not already fully handled and returned from above:
	return DefDlgProc(hWnd, iMsg, wParam, lParam);
}



LRESULT CALLBACK TabWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	// Variables are kept separate up here for future expansion of this function (to handle
	// more iMsgs/cases, etc.):
	GuiType *pgui;
	GuiControlType *pcontrol;
	HWND parent_window;
	HBRUSH bk_brush;
	COLORREF bk_color;

	switch (iMsg)
	{
	case WM_ERASEBKGND:
	case WM_WINDOWPOSCHANGED:
		parent_window = GetParent(hWnd);
		// Relies on short-circuit boolean order:
		if (  !(pgui = GuiType::FindGui(parent_window)) || !(pcontrol = pgui->FindControl(hWnd))  )
			break;
		switch (iMsg)
		{
		case WM_ERASEBKGND:
			pgui->ControlGetBkColor(*pcontrol, true, bk_brush, bk_color);
			if (!bk_brush)
				break; // Let default proc handle it.
			// Can't use SetBkColor(), need a real brush to fill it.
			RECT clipbox;
			GetClipBox((HDC)wParam, &clipbox);
			FillRect((HDC)wParam, &clipbox, bk_brush);
			return 1; // "An application should return nonzero if it erases the background."
		case WM_WINDOWPOSCHANGED:
			if ((LPWINDOWPOS(lParam)->flags & (SWP_NOMOVE | SWP_NOSIZE)) != (SWP_NOMOVE | SWP_NOSIZE)) // i.e. moved, resized or both.
			{
				// This appears to allow the tab control to update its row count:
				LRESULT r = CallWindowProc(g_TabClassProc, hWnd, iMsg, wParam, lParam);
				// Update this tab control's dialog, if it has one, to match the new size/position.
				pgui->UpdateTabDialog(hWnd);
				return r;
			}
		}
	}

	// This will handle anything not already fully handled and returned from above:
	return CallWindowProc(g_TabClassProc, hWnd, iMsg, wParam, lParam);
}



void GuiType::Event(GuiIndexType aControlIndex, UINT aNotifyCode, USHORT aGuiEvent, UINT_PTR aEventInfo)
// Caller should pass GUI_EVENT_NONE (zero) for aGuiEvent if it wants us to determine aGuiEvent based on the
// type of control and the incoming aNotifyCode.
// This function handles events within a GUI window that caused one of its controls to change in a meaningful
// way, or that is an event that could trigger an external action, such as clicking a button or icon.
{
	if (aControlIndex >= mControlCount) // Caller probably already checked, but just to be safe.
		return;
	GuiControlType &control = *mControl[aControlIndex];
	if (!control.events.Count()) // But don't check IsMonitoring(aGuiEvent) yet, because aGuiEvent may be undetermined.
		return; // No event handlers associated with this control, so no action.
	if (control.attrib & GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS)
		return;

	// Update: The below is now checked by MsgSleep() at the time the launch actually would occur because
	// g_nThreads will be more accurate/timely then:
	// If this control already has a thread running in its label, don't create a new thread to avoid
	// problems of buried threads, or a stack of suspended threads that might be resumed later
	// at an unexpected time. Users of timer subs that take a long time to run should be aware, as
	// documented in the help file, that long interruptions are then possible.
	//if (g_nThreads >= g_MaxThreadsTotal || (aControl->attrib & GUI_CONTROL_ATTRIB_HANDLER_IS_RUNNING))
	//	continue

	if (aGuiEvent == GUI_EVENT_WM_COMMAND) // Called by WM_COMMAND.
	{
		if (control.events.IsMonitoring(aNotifyCode, GUI_EVENTKIND_COMMAND))
		{
			POST_AHK_GUI_ACTION(mHwnd, aControlIndex, aGuiEvent, aNotifyCode);
			// Do the rest as well, just in case this command corresponds to an event.
		}
		aGuiEvent = GUI_EVENT_NONE; // Determine the actual event below.
	}

	if (aGuiEvent == GUI_EVENT_NONE) // Caller wants us to determine aGuiEvent based on control type and aNotifyCode.
	{
		switch(control.type)
		{
		case GUI_CONTROL_BUTTON:
		case GUI_CONTROL_CHECKBOX:
		case GUI_CONTROL_RADIO:
			switch (aNotifyCode)
			{
			case BN_CLICKED: // Must explicitly list this case since the default label below does a return.
				// Fix for v1.0.24: The below excludes from consideration messages from radios that are
				// being unchecked.  This prevents a radio group's g-label from being fired twice when the
				// user navigates to a new radio via the arrow keys.  It also filters out the BN_CLICKED that
				// occurs when the user tabs over to a radio group that lacks a selected button.  This new
				// behavior seems like it would be desirable most of the time.
				if (control.type == GUI_CONTROL_RADIO && SendMessage(control.hwnd, BM_GETCHECK, 0, 0) == BST_UNCHECKED)
					return;
				aGuiEvent = GUI_EVENT_CLICK;
				break;
			// All three of these require BS_NOTIFY:
			case BN_DBLCLK: aGuiEvent = GUI_EVENT_DBLCLK; break;
			case BN_SETFOCUS: aGuiEvent = GUI_EVENT_FOCUS; break;
			case BN_KILLFOCUS: aGuiEvent = GUI_EVENT_LOSEFOCUS; break;
			default:
				return;
			}
			break;

		case GUI_CONTROL_DROPDOWNLIST:
		case GUI_CONTROL_COMBOBOX:
			switch (aNotifyCode)
			{
			case CBN_SELCHANGE:  // Must explicitly list this case since the default label does a return.
			case CBN_EDITCHANGE: // Added for v1.0.24 to support detection of changes in a ComboBox's edit portion.
				aGuiEvent = GUI_EVENT_CHANGE;
				break;
			case CBN_DBLCLK: // Used by CBS_SIMPLE (i.e. list always visible).
				aGuiEvent = GUI_EVENT_DBLCLK; // But due to rarity of use, the focused row number is not stored in aEventInfo.
				break;
			case CBN_SETFOCUS: aGuiEvent = GUI_EVENT_FOCUS; break;
			case CBN_KILLFOCUS: aGuiEvent = GUI_EVENT_LOSEFOCUS; break;
			default:
				return;
			}
			break;

		case GUI_CONTROL_LISTBOX:
			switch (aNotifyCode)
			{
			case LBN_SELCHANGE: // Must explicitly list this case since the default label does a return.
				aGuiEvent = GUI_EVENT_CHANGE;
				break;
			case LBN_DBLCLK:
				aGuiEvent = GUI_EVENT_DBLCLK;
				aEventInfo = 1 + (UINT_PTR)SendMessage(control.hwnd, LB_GETCARETINDEX, 0, 0); // +1 to convert to one-based index.
				break;
			case LBN_SETFOCUS: aGuiEvent = GUI_EVENT_FOCUS; break;
			case LBN_KILLFOCUS: aGuiEvent = GUI_EVENT_LOSEFOCUS; break;
			default:
				return;
			}
			break;

		case GUI_CONTROL_EDIT:
			// Seems more appropriate to check EN_CHANGE vs. EN_UPDATE since EN_CHANGE occurs only after
			// any redrawing of the control.
			switch (aNotifyCode)
			{
			case EN_CHANGE: aGuiEvent = GUI_EVENT_CHANGE; break;
			case EN_SETFOCUS: aGuiEvent = GUI_EVENT_FOCUS; break;
			case EN_KILLFOCUS: aGuiEvent = GUI_EVENT_LOSEFOCUS; break;
			default:
				return;
			}
			break;

		case GUI_CONTROL_HOTKEY: // The only notification sent by the hotkey control is EN_CHANGE.
			aGuiEvent = GUI_EVENT_CHANGE;
			break; // No action.

		case GUI_CONTROL_TEXT:
		case GUI_CONTROL_PIC:
			// Update: Unlike buttons, it's all-or-none for static controls.  Testing shows that if
			// STN_DBLCLK is not checked for and the user clicks rapidly, half the clicks will be
			// ignored:
			// Based on experience with BN_DBLCLK, it's likely that STN_DBLCLK must be included or else
			// these control types won't be responsive to rapid consecutive clicks:
			switch (aNotifyCode)
			{
			case STN_CLICKED: // Must explicitly list this case since the default label does a return.
				aGuiEvent = GUI_EVENT_CLICK;
				break;
			case STN_DBLCLK:
				aGuiEvent = GUI_EVENT_DBLCLK;
				break;
			default:
				return;
			}
			break;

		case GUI_CONTROL_UPDOWN:
			// Due to the difficulty in distinguishing between clicking an arrow button and pressing an
			// arrow key on the keyboard, there is currently no GUI_CONTROL_ATTRIB_ALTSUBMIT mode for
			// up-downs.  That mode could be reserved to allow the script to override the user's position
			// change of the up-down by means of the script returning 1 or 0 in response to UDN_DELTAPOS.
			if (aNotifyCode == SB_THUMBPOSITION)
			{
				// User has pressed arrow keys or clicked down on the mouse on one of the arrows.
				aGuiEvent = GUI_EVENT_CHANGE;
				break;
			}
			// Otherwise, ignore all others.  SB_ENDSCROLL is received when user has released mouse after
			// scrolling one of the arrows (never arrives for arrow keys, even when holding them down).
			// That event (and any others, but especially that one) is ignored because it would launch
			// the g-label twice: once for lbutton-down and once for up.
			return;

		case GUI_CONTROL_SLIDER:
			switch (aNotifyCode)
			{
			case TB_THUMBTRACK:
				// TB_THUMBTRACK is sent very frequently while the user is dragging the slider.
				// It seems best to disable it by default, and only call the g-label once the user
				// releases the mouse button (i.e. when TB_THUMBPOSITION is received).
			case TB_ENDTRACK:
				// TB_ENDTRACK always follows some other event. Since it isn't generated when the user
				// uses the mouse wheel, filter it out by default so that the g-label is always called
				// only once when the slider moves, by default.
				if (!(control.attrib & GUI_CONTROL_ATTRIB_ALTSUBMIT)) // Ignore this event.
					return;
				// Otherwise, continue:
			default:
				// Namely the following:
				//case TB_THUMBPOSITION: // Mouse wheel or WM_LBUTTONUP following a TB_THUMBTRACK notification message
				//case TB_LINEUP:        // VK_LEFT or VK_UP
				//case TB_LINEDOWN:      // VK_RIGHT or VK_DOWN
				//case TB_PAGEUP:        // VK_PRIOR (the user clicked the channel above or to the left of the slider)
				//case TB_PAGEDOWN:      // VK_NEXT (the user clicked the channel below or to the right of the slider)
				//case TB_TOP:           // VK_HOME
				//case TB_BOTTOM:        // VK_END
				// In most cases the script simply wants to know when the value changes,
				// so this is one event with a parameter rather than multiple events:
				aGuiEvent = GUI_EVENT_CHANGE;
				aEventInfo = aNotifyCode;
			}
			break;

		// The following need no handling in this section because aGuiEvent and aEventInfo
		// were given appropriate values by our caller (execution never reaches this point):
		//case GUI_CONTROL_TREEVIEW:
		//case GUI_CONTROL_LISTVIEW:
		//case GUI_CONTROL_TAB: // aNotifyCode == TCN_SELCHANGE should be the only possibility.
		//case GUI_CONTROL_DATETIME:
		//case GUI_CONTROL_MONTHCAL:
		//case GUI_CONTROL_LINK:
		//
		// The following are not needed because execution never reaches this point.  This is because these types
		// are forbidden from having event handlers, since they don't support raising any events.
		//case GUI_CONTROL_GROUPBOX:
		//case GUI_CONTROL_PROGRESS:

		} // switch(control.type)
	} // if (aGuiEvent == GUI_EVENT_NONE)

	if (!control.events.IsMonitoring(LOBYTE(aGuiEvent))) // Checked after possibly determining aGuiEvent above.  LOBYTE() is needed for ListView 'I'.
		return;

	POST_AHK_GUI_ACTION(mHwnd, aControlIndex, aGuiEvent, (LPARAM)aEventInfo);
	// MsgSleep(-1, RETURN_AFTER_MESSAGES_SPECIAL_FILTER) is not done because "case AHK_GUI_ACTION" in GuiWindowProc()
	// takes care of it.  See its comments for why.

	// BACKGROUND ABOUT THE ABOVE:
	// Rather than launching the thread directly from here, it seems best to always post it to our
	// thread to be handled there.  Here are the reasons:
	// 1) We don't want to be in the situation where a thread launched here would return first
	//    to a dialog's message pump rather than MsgSleep's pump.  That's because our thread
	//    might have queued messages that would then be misrouted or lost because the dialog's
	//    dispatched them incorrectly, or didn't know what to do with them because they had a
	//    NULL hwnd.
	// 2) If the script happens to be uninterruptible, we would want to "re-queue" the messages
	//    in this fashion anyway (doing so avoids possible conflict if the current quasi-thread
	//    is in the middle of a critical operation, such as trying to open the clipboard for 
	//    a command it is right in the middle of executing).
	// 3) "Re-queuing" this event *only* for case #2 might cause problems with losing the original
	//     sequence of events that occurred in the GUI.  For example, if some events were re-queued
	//     due to uninterruptibility, but other more recent ones were not (because the thread became
	//     interruptible again at a critical moment), the more recent events would take effect before
	//     the older ones.  Requeuing all events ensures that when they do take effect, they do so
	//     in their original order.
	//
	// More explanation about Case #1 above.  Consider this order of events: 1) Current thread is
	// waiting for dialog, thus that dialog's msg pump is running. 2) User presses a button on GUI
	// window, and then another while the prev. button's thread is still uninterruptible (perhaps
	// this happened due to automating the form with the Send command).  3) Due to uninterruptibility,
	// this event would be re-queued to the thread's msg queue.  4) If the first thread ends before any
	// call to MsgSleep(), we'll be back at the dialog's msg pump again, and thus the requeued message
	// would be misrouted or discarded due to automatic dispatching.
	//
	// Info about why events are buffered when script is uninterruptible:
 	// It seems best to buffer GUI events that would trigger a new thread, rather than ignoring
	// them or allowing them to launch unconditionally.  Ignoring them is bad because lost events
	// might cause a GUI to get out of sync with how its controls are designed to update and
	// interact with each other.  Allowing them to launch anyway is bad in the case where
	// a critical operation for another thread is underway (such as attempting to open the
	// clipboard).  Thus, post this event back to our thread so that even if a msg pump
	// other than our own is running (such as MsgBox or another dialog), the msg will stay
	// buffered (the msg might bounce around fiercely if we kept posting it to our WindowProc,
	// since any other msg pump would not have the buffering filter in place that ours has).
	//
	// UPDATE: I think some of the above might be obsolete now.  Consider this order of events:
	// 1) Current thread is waiting for dialog, thus that dialog's msg pump is running. 2) User presses
	// a button on GUI window. 3) We receive that message here and post it below. 4) When we exit,
	// the dialog's msg pump receives the message and dispatches it to GuiWindowProc. 5) GuiWindowProc
	// posts the message and immediately does a MsgSleep to force our msg pump to handle the message.
	// 6) Everything is fine unless our msg pump leaves the message queued due to uninterruptibility,
	// in which case the msg will bounce around, possibly maxing the CPU, until either the dialog's msg
	// pump ends or the other thread becomes interruptible (which usually doesn't take more than 20 ms).
	// If the script is uninterruptible due to some long operation such as sending keystrokes or trying
	// to open the clipboard when it is locked, any dialog message pump underneath on the call stack
	// shouldn't be an issue because the long-operation (clipboard/Send) does not return to that
	// msg pump until the operation is complete; in other words, the message to stay queued rather than
	// bouncing around.

	// Although an additional WORD of info could be squeezed into the message by passing the control's HWND
	// instead of the parent window's (and the msg pump could then look up the parent window via
	// GetNonChildParent), it's probably not feasible because if some other message pump is running, it would
	// route AHK_GUI_ACTION messages to the window proc. of the control rather than the parent window, which
	// would prevent them from being re-posted back to the queue (see "case AHK_GUI_ACTION" in GuiWindowProc()).
}


bool GuiType::ControlWmNotify(GuiControlType &aControl, LPNMHDR aNmHdr, INT_PTR &aRetVal)
{
	// See MsgMonitor() in application.cpp for comments, as this method is based on the beforementioned function.

	if (!INTERRUPTIBLE_IN_EMERGENCY)
		return false;

	if (!aControl.events.IsMonitoring(aNmHdr->code, GUI_EVENTKIND_NOTIFY))
		return false;

	if (g_nThreads >= g_MaxThreadsTotal)
		return false;

	if (g->Priority > 0)
		return false;

	InitNewThread(0, false, true);
	g_script.mLastPeekTime = GetTickCount();
	AddRef();
	
	ExprTokenType param[] = { &aControl, (__int64)(DWORD_PTR)aNmHdr };
	ResultType result = aControl.events.Call(param, _countof(param)
		, aNmHdr->code, GUI_EVENTKIND_NOTIFY, this, &aRetVal);

	Release();
	ResumeUnderlyingThread();

	// Consider this notification fully handled only if a non-empty value was returned.
	return result == EARLY_RETURN;
}


WORD GuiType::TextToHotkey(LPCTSTR aText)
// Returns a WORD compatible with the HKM_SETHOTKEY message:
// LOBYTE is the virtual key.
// HIBYTE is a set of modifiers:
// HOTKEYF_ALT ALT key
// HOTKEYF_CONTROL CONTROL key
// HOTKEYF_SHIFT SHIFT key
// HOTKEYF_EXT Extended key
{
	if (!*aText)
		return 0;

	BYTE modifiers = 0; // Set default.
	for (; aText[1]; ++aText) // For each character except the last.
	{
		switch (*aText)
		{
		case '!': modifiers |= HOTKEYF_ALT; continue;
		case '^': modifiers |= HOTKEYF_CONTROL; continue;
		case '+': modifiers |= HOTKEYF_SHIFT; continue;
		//default: // Some other character type, so it marks the end of the modifiers.
		}
		break;
	}

	// For translating the virtual key below, the following notes apply:
	// The following extended keys are unlikely, and in any case don't have a non-extended counterpart,
	// so no special handling:
	// VK_CANCEL (Ctrl-break)
	// VK_SNAPSHOT (PrintScreen).
	//
	// These do not have a non-extended counterpart, i.e. their VK has only one possible scan code:
	// VK_DIVIDE (NumpadDivide/slash)
	// VK_NUMLOCK
	//
	// All of the following are handled properly via the scan code logic below:
	// VK_INSERT
	// VK_PRIOR
	// VK_NEXT
	// VK_HOME
	// VK_END
	// VK_UP
	// VK_DOWN
	// VK_LEFT
	// VK_RIGHT
	//
	// Same note as above but these cannot be typed by the user, only programmatically inserted via
	// initial value of "Gui Add" or via "GuiControl,, MyHotkey, ^Delete":
	// VK_DELETE
	// VK_RETURN
	// Note: NumpadEnter (not Enter) is extended, unlike Home/End/Pgup/PgDn/Arrows, which are
	// NON-extended on the keypad.

	modLR_type mods = 0;
	BYTE vk = TextToVK(aText, &mods);
	if (!vk)
		return 0;  // Indicate total failure because the key text is invalid.
	if (mods & (MOD_LALT | MOD_RALT))
		modifiers |= HOTKEYF_ALT;
	if (mods & (MOD_LCONTROL | MOD_RCONTROL))
		modifiers |= HOTKEYF_CONTROL;
	if (mods & (MOD_LSHIFT | MOD_RSHIFT))
		modifiers |= HOTKEYF_SHIFT;
	// Find out if the HOTKEYF_EXT flag should be set.
	sc_type sc = TextToSC(aText); // Better than vk_to_sc() since that has both an primary and secondary scan codes to choose from.
	if (!sc) // Since not found above, default to the primary scan code.
		sc = vk_to_sc(vk);
	if (sc & 0x100) // The scan code derived above is extended.
		modifiers |= HOTKEYF_EXT;
	return MAKEWORD(vk, modifiers);
}



LPTSTR GuiType::HotkeyToText(WORD aHotkey, LPTSTR aBuf)
// Caller has ensured aBuf is large enough to hold any hotkey name.
// Returns aBuf.
{
	BYTE modifiers = HIBYTE(aHotkey); // In this case, both the VK and the modifiers are bytes, not words.
	LPTSTR cp = aBuf;
	if (modifiers & HOTKEYF_SHIFT)
		*cp++ = '+';
	if (modifiers & HOTKEYF_CONTROL)
		*cp++ = '^';
	if (modifiers & HOTKEYF_ALT)
		*cp++ = '!';
	BYTE vk = LOBYTE(aHotkey);

	if (modifiers & HOTKEYF_EXT) // Try to find the extended version of this VK if it has two versions.
	{
		// Fix for v1.0.37.03: A virtual key that has only one scan code should be resolved by VK
		// rather than SC.  Otherwise, Numlock will wind up being SC145 and NumpadDiv something similar.
		// If a hotkey control could capture AppsKey, PrintScreen, Ctrl-Break (VK_CANCEL), which it can't, this
		// would also apply to them.
		// Fix for v1.1.26.02: Check sc2 != 0, not sc1 != sc2, otherwise the fix described above doesn't work.
		sc_type sc2 = vk_to_sc(vk, true); // Secondary scan code.
		if (sc2) // Non-zero means this key has two scan codes.
		{
			sc_type sc1 = vk_to_sc(vk); // Primary scan code for this virtual key.
			sc_type sc = (sc2 & 0x100) ? sc2 : sc1;
			if (sc & 0x100)
			{
				SCtoKeyName(sc, cp, 100);
				return aBuf;
			}
		}
	}
	// Since above didn't return, use a simple lookup on VK, since it gives preference to non-extended keys.
	return VKtoKeyName(vk, cp, 100);
}



void GuiType::ControlCheckRadioButton(GuiControlType &aControl, GuiIndexType aControlIndex, WPARAM aCheckType)
{
	GuiIndexType radio_start, radio_end;
	FindGroup(aControlIndex, radio_start, radio_end); // Even if the return value is 1, do the below because it ensures things like tabstop are in the right state.
	if (aCheckType == BST_CHECKED)
	{
		// Testing shows that controls with different parent windows never behave as though they
		// are part of the same radio group.  CheckRadioButton has been confirmed as working with
		// both non-dialog windows (such as Static/Text) and Tab3's tab dialog.  This is required
		// for this function to work on Radio buttons within a Tab3 control:
		HWND hDlg = GetParent(aControl.hwnd);
		// This will check the specified button and uncheck all the others in the group.
		// There is at least one other reason to call CheckRadioButton() rather than doing something
		// manually: It prevents an unwanted firing of the radio's g-label upon WM_ACTIVATE,
		// at least when a radio group is first in the window's z-order and the radio group has
		// an initially selected button:
		CheckRadioButton(hDlg, GUI_INDEX_TO_ID(radio_start), GUI_INDEX_TO_ID(radio_end - 1), GUI_INDEX_TO_ID(aControlIndex));
	}
	else // Uncheck it.
	{
		// If the group was originally created with the tabstop style, unchecking the button that currently
		// has that style would also remove the tabstop style because apparently that's how radio buttons
		// respond to being unchecked.  Compensate for this by giving the first radio in the group the
		// tabstop style. Update: The below no longer checks to see if the radio group has the tabstop style
		// because it's fairly pointless.  This is because when the user checks/clicks a radio button,
		// the control automatically acquires the tabstop style.  In other words, even though -Tabstop is
		// allowed in a radio's options, it will be overridden the first time a user selects a radio button.
		HWND first_radio_in_group = NULL;
		// Find the first radio in this control group:
		for (GuiIndexType u = radio_start; u < radio_end; ++u)
			if (mControl[u]->type == GUI_CONTROL_RADIO) // Since a group can have non-radio controls in it.
			{
				first_radio_in_group = mControl[u]->hwnd;
				break;
			}
		// The below can't be done until after the above because it would remove the tabstop style
		// if the specified radio happens to be the one that has the tabstop style.  This is usually
		// the case since the button is usually on (since the script is now turning it off here),
		// and other logic has ensured that the on-button is the one with the tabstop style:
		SendMessage(aControl.hwnd, BM_SETCHECK, BST_UNCHECKED, 0);
		if (first_radio_in_group) // Apply the tabstop style to it.
			SetWindowLong(first_radio_in_group, GWL_STYLE, WS_TABSTOP | GetWindowLong(first_radio_in_group, GWL_STYLE));
	}
}



void GuiType::ControlSetUpDownOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt)
// Caller has ensured that aControl.type is an UpDown.
{
	if (aOpt.range_changed)
		SendMessage(aControl.hwnd, UDM_SETRANGE32, aOpt.range_min, aOpt.range_max);
}



int GuiType::ControlGetDefaultSliderThickness(DWORD aStyle, int aThumbThickness)
{
	if (aThumbThickness < 1)
		aThumbThickness = 20;  // Set default.
	// Provide a small margin on both sides, otherwise the bar is sometimes truncated.
	aThumbThickness += 5; // 5 looks better than 4 in most styles/themes.
	if (aStyle & TBS_NOTICKS) // This takes precedence over TBS_BOTH (which if present will still make the thumb flat vs. pointed).
		return DPIScale(aThumbThickness);
	if (aStyle & TBS_BOTH)
		return DPIScale(aThumbThickness + 16);
	return DPIScale(aThumbThickness + 8);
}



int GuiType::ControlInvertSliderIfNeeded(GuiControlType &aControl, int aPosition)
// Caller has ensured that aControl.type is slider.
{
	return (aControl.attrib & GUI_CONTROL_ATTRIB_ALTBEHAVIOR)
		? ((int)SendMessage(aControl.hwnd, TBM_GETRANGEMAX, 0, 0) - aPosition) + (int)SendMessage(aControl.hwnd, TBM_GETRANGEMIN, 0, 0)
		: aPosition;  // No inversion necessary.
}



void GuiType::ControlSetSliderOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt)
// Caller has ensured that aControl.type is slider.
{
	if (aOpt.range_changed)
	{
		// Don't use TBM_SETRANGE because then only 16-bit values are supported:
		SendMessage(aControl.hwnd, TBM_SETRANGEMIN, FALSE, aOpt.range_min); // No redraw
		SendMessage(aControl.hwnd, TBM_SETRANGEMAX, TRUE, aOpt.range_max); // Redraw.
	}
	if (aOpt.tick_interval_changed)
	{
		if (aOpt.tick_interval < 0) // This is the signal to remove the existing tickmarks.
			SendMessage(aControl.hwnd, TBM_CLEARTICS, TRUE, 0);
		else if (aOpt.tick_interval_specified)
			SendMessage(aControl.hwnd, TBM_SETTICFREQ, aOpt.tick_interval, 0);
		else
			// +TickInterval without a value.  TBS_AUTOTICKS was added, but doesn't take effect
			// until TBM_SETRANGE' or TBM_SETTICFREQ is sent.  Since the script might have already
			// set an interval and there's no TBM_GETTICFREQ, use TBM_SETRANGEMAX to set the ticks.
			if (!aOpt.range_changed) // TBM_SETRANGEMAX wasn't already sent.
			{
				SendMessage(aControl.hwnd, TBM_SETRANGEMAX, TRUE
					, SendMessage(aControl.hwnd, TBM_GETRANGEMAX, 0, 0));
			}
	}
	if (aOpt.line_size > 0) // Removal is not supported, so only positive values are considered.
		SendMessage(aControl.hwnd, TBM_SETLINESIZE, 0, aOpt.line_size);
	if (aOpt.page_size > 0) // Removal is not supported, so only positive values are considered.
		SendMessage(aControl.hwnd, TBM_SETPAGESIZE, 0, aOpt.page_size);
	if (aOpt.thickness > 0)
		SendMessage(aControl.hwnd, TBM_SETTHUMBLENGTH, aOpt.thickness, 0);
	if (aOpt.tip_side)
		SendMessage(aControl.hwnd, TBM_SETTIPSIDE, aOpt.tip_side - 1, 0); // -1 to convert back to zero base.

	// Buddy positioning is left primitive and automatic even when auto-position of this slider
	// or the controls that come after it is in effect.  This is because buddy controls seem too
	// rarely used (due to their lack of positioning options), which is the reason why extra code
	// isn't added here and in "GuiControl Move" to treat the buddies as part of the control rect
	// (i.e. as an entire unit), nor any code for auto-positioning the entire unit when the control
	// is created.  If such code were added, it would require even more code if the slider itself
	// has an automatic position, because the positions of its buddies would have to be recorded,
	// then after they are set as buddies (which moves them) the slider is moved up or left to the
	// position where its buddies used to be.  Otherwise, there would be a gap left during
	// auto-layout.
	// For these, removal is not supported, only changing, since removal seems too rarely needed:
	if (aOpt.buddy1)
		SendMessage(aControl.hwnd, TBM_SETBUDDY, TRUE, (LPARAM)aOpt.buddy1->hwnd);  // Left/top
	if (aOpt.buddy2)
		SendMessage(aControl.hwnd, TBM_SETBUDDY, FALSE, (LPARAM)aOpt.buddy2->hwnd); // Right/bottom
}



void GuiType::ControlSetListViewOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt)
// Caller has ensured that aControl.type is ListView.
// Caller has ensured that aOpt.color_bk is CLR_INVALID if no change should be made to the
// current background color.
{
	if (aOpt.limit)
	{
		if (ListView_GetItemCount(aControl.hwnd) > 0)
			SendMessage(aControl.hwnd, LVM_SETITEMCOUNT, aOpt.limit, 0); // Last parameter should be 0 for LVS_OWNERDATA (verified if you look at the definition of ListView_SetItemCount macro).
		else
			// When the control has no rows, work around the fact that LVM_SETITEMCOUNT delivers less than 20%
			// of its full benefit unless done after the first row is added (at least on XP SP1).  The message
			// is deferred until later by setting this flag:
			aControl.union_lv_attrib->row_count_hint = aOpt.limit;
	}
	if (aOpt.color != CLR_INVALID || aOpt.color_bk != CLR_INVALID)
	{
		if (aOpt.color != CLR_INVALID) // Explicit color change was requested.
		{
			// Although MSDN doesn't say, ListView_SetTextColor() appears to support CLR_DEFAULT.
			ListView_SetTextColor(aControl.hwnd, aOpt.color);
			aOpt.color = CLR_INVALID; // Let ControlParseOptions() know we handled it.
		}
		if (aOpt.color_bk != CLR_INVALID) // Explicit color change was requested.
		{
			// Making both the same seems the best default because BkColor only applies to the portion
			// of the control that doesn't have text in it, which is typically very little.
			// Unlike ListView_SetTextBkColor, ListView_SetBkColor() treats CLR_DEFAULT as black.
			// therefore, make them both GetSysColor(COLOR_WINDOW) for consistency.  This color is
			// probably the default anyway:
			COLORREF color = (aOpt.color_bk == CLR_DEFAULT) ? GetSysColor(COLOR_WINDOW) : aOpt.color_bk;
			ListView_SetTextBkColor(aControl.hwnd, color);
			ListView_SetBkColor(aControl.hwnd, color);
		}
		// It used to work without this; I don't know what conditions changed, but apparently it's needed
		// at least sometimes.  The last param must be TRUE otherwise an space not filled by rows or columns
		// doesn't get updated:
		InvalidateRect(aControl.hwnd, NULL, TRUE);
	}
}



void GuiType::ControlSetTreeViewOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt)
// Caller has ensured that aControl.type is TreeView.
// Caller has ensured that aOpt.color_bk is CLR_INVALID if no change should be made to the
// current background color.
{
	if (aOpt.color != CLR_INVALID) // Explicit color change was requested.
	{
		TreeView_SetTextColor(aControl.hwnd, aOpt.color == CLR_DEFAULT ? -1 : aOpt.color); // Must pass -1 rather than CLR_DEFAULT for TreeView controls.
		aOpt.color = CLR_INVALID; // Let ControlParseOptions() know we handled it.
	}
	if (aOpt.color_bk != CLR_INVALID)
		TreeView_SetBkColor(aControl.hwnd, aOpt.color_bk == CLR_DEFAULT ? -1 : aOpt.color_bk);
	// Disabled because it apparently is not supported on XP:
	//if (aOpt.tabstop_count)
	//	// Although TreeView_GetIndent() confirms that the following takes effect, it doesn't seem to 
	//	// cause any visible change, at least under XP SP2 (tried various style changes as well as Classic vs.
	//	// XP theme, as well as the "-Theme" option.
	//	TreeView_SetIndent(aControl.hwnd, aOpt.tabstop[0]);

	// Unlike ListView, seems not to be needed:
	//InvalidateRect(aControl.hwnd, NULL, TRUE);
}



void GuiType::ControlSetProgressOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt, DWORD aStyle)
// Caller has ensured that aControl.type is Progress.
// Caller has ensured that aOpt.color_bk is CLR_INVALID if no change should be made to the
// bar's current background color.
{
	// If any options are present that cannot be manifest while a visual theme is in effect, ensure any
	// such theme is removed from the control (currently, once removed it is never put back on):
	// Override the default so that colors/smooth can be manifest even when non-classic theme is in effect.
	if (   IS_AN_ACTUAL_COLOR(aOpt.color)
		|| IS_AN_ACTUAL_COLOR(aOpt.color_bk)
		|| (aStyle & PBS_SMOOTH))
		SetWindowTheme(aControl.hwnd, L"", L""); // Remove theme if options call for something theme can't show.

	if (aOpt.range_min || aOpt.range_max) // Must check like this because although it valid for one to be zero, both should not be.
		SendMessage(aControl.hwnd, PBM_SETRANGE32, aOpt.range_min, aOpt.range_max);

	if (aOpt.color != CLR_INVALID) // Explicit color change was requested.
		SendMessage(aControl.hwnd, PBM_SETBARCOLOR, 0, aOpt.color);

	if (aOpt.color_bk != CLR_INVALID) // Explicit color change was requested.
		SendMessage(aControl.hwnd, PBM_SETBKCOLOR, 0, aOpt.color_bk);
}



GuiControlType *GuiType::ControlOverrideBkColor(GuiControlType &aControl)
// Caller has ensured that aControl.type is something for which the window's or tab control's background
// should apply (e.g. Progress or Text).
{
	GuiControlType *ptab_control;
	if (!mTabControlCount || !(ptab_control = FindTabControl(aControl.tab_control_index))
		|| ptab_control->background_color == CLR_INVALID) // Relies on short-circuit boolean order.
		return NULL;  // Override not needed because control isn't on a tab, or it's tab has same color as window.
	if (GetParent(aControl.hwnd) != mHwnd)
		// Even if the control doesn't lie entirely within the tab control's display area,
		// it is only visible within that area due to the parent-child relationship.
		return ptab_control;
	// Does this control lie mostly inside the tab?  Note that controls can belong to a tab page even though
	// they aren't physically located inside the page.
	RECT overlap_rect, tab_rect, control_rect;
	GetWindowRect(ptab_control->hwnd, &tab_rect);
	GetWindowRect(aControl.hwnd, &control_rect);
	IntersectRect(&overlap_rect, &tab_rect, &control_rect);
	// Returns true if more than 50% of control's area is inside the tab:
	return (overlap_rect.right - overlap_rect.left) * (overlap_rect.bottom - overlap_rect.top)
		> 0.5 * (control_rect.right - control_rect.left) * (control_rect.bottom - control_rect.top)
		? ptab_control : NULL;
}



void GuiType::ControlGetBkColor(GuiControlType &aControl, bool aUseWindowColor, HBRUSH &aBrush, COLORREF &aColor)
{
	if (aControl.background_color != CLR_INVALID)
	{
		aBrush = aControl.background_brush;
		aColor = aControl.background_color;
	}
	else if (aUseWindowColor)
	{
		// If this static control both belongs to a tab control and is within its physical boundaries,
		// match its background to the tab control's.  This is only necessary if the tab control has
		// the +/-Background option, since otherwise its background would be the same as the window's:
		if (GuiControlType *tab_control = ControlOverrideBkColor(aControl))
		{
			aBrush = tab_control->background_brush;
			aColor = tab_control->background_color;
		}
		else
		{
			aBrush = mBackgroundBrushWin;
			aColor = mBackgroundColorWin;
		}
	}
	else
	{
		// Caller requires these initialised.
		aBrush = NULL;
		aColor = CLR_DEFAULT;
	}
}



void GuiType::ControlUpdateCurrentTab(GuiControlType &aTabControl, bool aFocusFirstControl)
// Handles the selection of a new tab in a tab control.
{
	int curr_tab_index = TabCtrl_GetCurSel(aTabControl.hwnd);
	// If there's no tab selected, any controls contained by the tab control should be hidden.
	// This can happen if all tabs are removed, such as by GuiControl,,TabCtl,|
	//if (curr_tab_index == -1) // No tab is selected.
	//	return;

	// Fix for v1.0.23:
	// If the tab control lacks the visible property, hide all its controls on all its tabs.
	// Don't use IsWindowVisible() because that would say that tab is hidden just because the parent
	// window is hidden, which is not desirable because then when the parent is shown, the shower
	// would always have to remember to call us.  This "hide all" behavior is done here rather than
	// attempting to "rely on everyone else to do their jobs of keeping the controls in the right state"
	// because it improves maintainability:
	DWORD tab_style = GetWindowLong(aTabControl.hwnd, GWL_STYLE);
	bool hide_all = !(tab_style & WS_VISIBLE); // Regardless of whether mHwnd is visible or not.
	bool disable_all = (tab_style & WS_DISABLED); // Don't use IsWindowEnabled() because it might return false if parent is disabled?
	// Say that the focus was already set correctly if the entire tab control is hidden or caller said
	// not to focus it:
	bool focus_was_set;
	bool parent_is_visible = IsWindowVisible(mHwnd);
	bool parent_is_visible_and_not_minimized = parent_is_visible && !IsIconic(mHwnd);
	if (hide_all || disable_all)
		focus_was_set = true;  // Tell the below not to set focus, since all tab controls are hidden or disabled.
	else if (aFocusFirstControl)  // Note that SetFocus() has an effect even if the parent window is hidden. i.e. next time the window is shown, the control will be focused.
		focus_was_set = false; // Tell it to focus the first control on the new page.
	else
	{
		HWND focused_hwnd;
		GuiControlType *focused_control;
		// If the currently focused control is somewhere in this tab control (but not the tab control
		// itself, because arrow-key navigation relies on tabs stay focused while the user is pressing
		// left and right-arrow), override the fact that aFocusFirstControl is false so that when the
		// page changes, its first control will be focused:
		focus_was_set = !(   parent_is_visible && (focused_hwnd = GetFocus())
			&& (focused_control = FindControl(focused_hwnd))
			&& focused_control->tab_control_index == aTabControl.tab_index   );
	}

	bool will_be_visible, will_be_enabled, has_visible_style, has_enabled_style, member_of_current_tab, control_state_altered;
	DWORD style;
	RECT rect, tab_rect;
	POINT *rect_pt = (POINT *)&rect; // i.e. have rect_pt be an array of the two points already within rect.

	GetWindowRect(aTabControl.hwnd, &tab_rect);

	// Update: Don't do the below because it causes a tab to look focused even when it isn't in cases
	// where a control was focused while drawing was suspended.  This is because the below omits the
	// tab rows themselves from the InvalidateRect() further below:
	// Tabs on left (TCS_BUTTONS only) require workaround, at least on XP.  Otherwise tab_rect.left will be
	// much too large.  Because of this, include entire tab rect if it can't be "deflated" reliably:
	//if (!(tab_style & TCS_VERTICAL) || (tab_style & TCS_RIGHT) || !(tab_style & TCS_BUTTONS))
	//	TabCtrl_AdjustRect(aTabControl.hwnd, FALSE, &tab_rect); // Reduce it to just the area without the tabs, since the tabs have already been redrawn.

	HWND tab_dialog = (HWND)GetProp(aTabControl.hwnd, _T("ahk_dlg"));
	if (tab_dialog)
		// Hiding the tab dialog and showing after re-enabling redraw solves an issue with ActiveX
		// WebBrowser controls not being redrawn (seems to be a WM_SETREDRAW glitch).
		ShowWindow(tab_dialog, SW_HIDE);

	// For a likely cleaner transition between tabs, disable redrawing until the switch is complete.
	// Doing it this way also serves to refresh a tab whose controls have just been disabled via
	// something like "GuiControl, Disable, MyTab", which would otherwise not happen because unlike
	// ShowWindow(), EnableWindow() apparently does not cause a repaint to occur.
	// Fix for v1.0.25.14: Don't send the message below (and its counterpart later on) because that
	// sometimes or always, as a side-effect, shows the window if it's hidden:
	// Fix for v1.1.27.07: WM_SETREDRAW appears to have the effect of clearing the update region
	// (causing any pending redraw, such as by a previous call to this function for a different tab
	// control, to be discarded).  To work around this, save the current update region.
	RECT update_rect = {0}; // For simplicity, use GetUpdateRect() vs GetUpdateRgn().
	if (parent_is_visible_and_not_minimized)
	{
		GetUpdateRect(mHwnd, &update_rect, FALSE);
		SendMessage(mHwnd, WM_SETREDRAW, FALSE, 0);
	}
	bool invalidate_entire_parent = false; // Set default.

	// Even if mHwnd is hidden, set styles to Show/Hide and Enable/Disable any controls that need it.
	for (GuiIndexType u = 0; u < mControlCount; ++u)
	{
		// Note aTabControl.tab_index stores aTabControl's tab_control_index (true only for type GUI_CONTROL_TAB).
		if (mControl[u]->tab_control_index != aTabControl.tab_index) // This control is not in this tab control.
			continue;
		GuiControlType &control = *mControl[u]; // Probably helps performance; certainly improves conciseness.
		member_of_current_tab = (control.tab_index == curr_tab_index);
		will_be_visible = !hide_all && member_of_current_tab && !(control.attrib & GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN);
		will_be_enabled = !disable_all && member_of_current_tab && !(control.attrib & GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED);
		// Above: Disabling the control when it's not on the current tab means that it can't be
		// activated, even by a keyboard shorcut - for instance, a button with text "&Click me"
		// could be activated by pressing Alt+C if the control was enabled but hidden.
		// Don't use IsWindowVisible() because if the parent window is hidden, I think that will
		// always say that the controls are hidden too.  In any case, IsWindowVisible() does not
		// work correctly for this when the window is first shown:
		style = GetWindowLong(control.hwnd, GWL_STYLE);
		has_visible_style =  style & WS_VISIBLE;
		has_enabled_style = !(style & WS_DISABLED);
		// Showing/hiding/enabling/disabling only when necessary might cut down on redrawing:
		control_state_altered = false;  // Set default.
		if (will_be_visible != has_visible_style)
		{
			ShowWindow(control.hwnd, will_be_visible ? SW_SHOWNOACTIVATE : SW_HIDE);
			control_state_altered = true;
		}
		if (will_be_enabled != has_enabled_style)
		{
			// Note that it seems to make sense to disable even text/pic/groupbox controls because
			// they can receive clicks and double clicks (except GroupBox).
			EnableWindow(control.hwnd, will_be_enabled);
			control_state_altered = true;
		}
		if (control_state_altered)
		{
			// If this altered control lies at least partially outside the tab's interior,
			// set it up to do the full repaint of the parent window:
			GetWindowRect(control.hwnd, &rect);
			if (!(PtInRect(&tab_rect, rect_pt[0]) && PtInRect(&tab_rect, rect_pt[1])))
				invalidate_entire_parent = true;
		}
		// The above's use of show/hide across a wide range of controls may be necessary to support things
		// such as the dynamic removal of tabs via "GuiControl,, MyTab, |NewTabSet1|NewTabSet2", i.e. if the
		// newly added removed tab was active, it's controls should now be hidden.
		// The below sets focus to the first input-capable control, which seems standard for the tab-control
		// dialogs I've seen.
		if (!focus_was_set && member_of_current_tab && will_be_visible && will_be_enabled)
		{
			switch(control.type)
			{
			case GUI_CONTROL_TEXT:
			case GUI_CONTROL_PIC:
			case GUI_CONTROL_GROUPBOX:
			case GUI_CONTROL_PROGRESS:
			case GUI_CONTROL_MONTHCAL:
			case GUI_CONTROL_UPDOWN: // It appears that not even non-buddied up-downs can be focused.
				break; // Do nothing for the above types because they cannot be focused.
			default:
			//case GUI_CONTROL_STATUSBAR: Nothing needs to be done because other logic has ensured it can't be a member of any tab.
			//case GUI_CONTROL_BUTTON:
			//case GUI_CONTROL_CHECKBOX:
			//case GUI_CONTROL_RADIO:
			//case GUI_CONTROL_DROPDOWNLIST:
			//case GUI_CONTROL_COMBOBOX:
			//case GUI_CONTROL_LISTBOX:
			//case GUI_CONTROL_LISTVIEW:
			//case GUI_CONTROL_TREEVIEW:
			//case GUI_CONTROL_EDIT:
			//case GUI_CONTROL_DATETIME:
			//case GUI_CONTROL_HOTKEY:
			//case GUI_CONTROL_SLIDER:
			//case GUI_CONTROL_TAB:
			//case GUI_CONTROL_LINK:
				// Fix for v1.0.24: Don't check the return value of SetFocus() because sometimes it returns
				// NULL even when the call will wind up succeeding.  For example, if the user clicks on
				// the second tab in a tab control, SetFocus() will probably return NULL because there
				// is not previously focused control at the instant the call is made.  This is because
				// the control that had focus has likely already been hidden and thus lost focus before
				// we arrived at this stage:
				SetFocus(control.hwnd); // Note that this has an effect even if the parent window is hidden. i.e. next time the parent is shown, this control will be focused.
				focus_was_set = true; // i.e. SetFocus() only for the FIRST control that meets the above criteria.
			}
		}
	}

	if (parent_is_visible_and_not_minimized) // Fix for v1.0.25.14.  See further above for details.
		SendMessage(mHwnd, WM_SETREDRAW, TRUE, 0); // Re-enable drawing before below so that tab can be focused below.

	if (tab_dialog && !hide_all)
		ShowWindow(tab_dialog, SW_SHOW);

	// In case tab is empty or there is no control capable of receiving focus, focus the tab itself
	// instead.  This allows the Ctrl-Pgdn/Pgup keyboard shortcuts to continue to navigate within
	// this tab control rather than having the focus get kicked backed outside the tab control
	// -- which I think happens when the tab contains no controls or only text controls (pic controls
	// seem okay for some reason), i.e. if the control with focus is hidden, the dialog falls back to
	// giving the focus to the the first focus-capable control in the z-order.
	if (!focus_was_set)
		SetFocus(aTabControl.hwnd); // Note that this has an effect even if the parent window is hidden. i.e. next time the parent is shown, this control will be focused.

	// UPDATE: Below is now only done when necessary to cut down on flicker:
	// Seems best to invalidate the entire client area because otherwise, if any of the tab's controls lie
	// outside of its interior (this is common for TCS_BUTTONS style), they would not get repainted properly.
	// In addition, tab controls tend to occupy the majority of their parent's client area anyway:
	if (parent_is_visible_and_not_minimized)
	{
		if (invalidate_entire_parent)
			InvalidateRect(mHwnd, NULL, TRUE); // TRUE seems safer.
		else
		{
			if (update_rect.left != update_rect.right)
				// Reset the window's update region, which was likely cleared by WM_SETREDRAW.
				// Doing this separately to tab_rect might perform better than using UnionRect(),
				// which could cover a much larger area than necessary.  The system should handle
				// the case where the two rects overlap (i.e. there won't be redundant repainting).
				InvalidateRect(mHwnd, &update_rect, TRUE);
			MapWindowPoints(NULL, mHwnd, (LPPOINT)&tab_rect, 2); // Convert rect to client coordinates (not the same as GetClientRect()).
			InvalidateRect(mHwnd, &tab_rect, TRUE); // Seems safer to use TRUE, not knowing all possible overlaps, etc.
		}
	}
}



GuiControlType *GuiType::FindTabControl(TabControlIndexType aTabControlIndex)
{
	if (aTabControlIndex == MAX_TAB_CONTROLS)
		// This indicates it's not a member of a tab control. Callers rely on this check.
		return NULL;
	TabControlIndexType tab_control_index = 0;
	for (GuiIndexType u = 0; u < mControlCount; ++u)
		if (mControl[u]->type == GUI_CONTROL_TAB)
			if (tab_control_index == aTabControlIndex)
				return mControl[u];
			else
				++tab_control_index;
	return NULL; // Since above didn't return, indicate failure.
}



int GuiType::FindTabIndexByName(GuiControlType &aTabControl, LPTSTR aName, bool aExactMatch)
// Find the first tab in this tab control whose leading-part-of-name matches aName.
// Return int vs. TabIndexType so that failure can be indicated.
{
	int tab_count = TabCtrl_GetItemCount(aTabControl.hwnd);
	// Although strictly performing a leading-part-of-name match should technically cause
	// an empty string to result in the first item (index 0), it seems unlikely that a user
	// would expect that.  The caller should verify *aName != 0.
	if (!tab_count || !*aName)
		return -1; // No match.

	TCITEM tci;
	tci.mask = TCIF_TEXT;
	TCHAR buf[1024];
	tci.pszText = buf;
	tci.cchTextMax = _countof(buf) - 1; // MSDN example uses -1.

	size_t aName_length = _tcslen(aName);
	if (aName_length >= _countof(buf)) // Checking this early avoids having to check it in the loop.
		return -1; // No match possible.

	for (int i = 0; i < tab_count; ++i)
	{
		if (TabCtrl_GetItem(aTabControl.hwnd, i, &tci))
		{
			if (aExactMatch)
			{
				if (!_tcsicmp(tci.pszText, aName))  // Match found.
					return i;
			}
			else
			{
				tci.pszText[aName_length] = '\0'; // Facilitates checking of only the leading part like strncmp(). Buffer overflow is impossible due to a check higher above.
				if (!lstrcmpi(tci.pszText, aName)) // Match found.
					return i;
			}
		}
	}

	// Since above didn't return, no match found.
	return -1;
}



int GuiType::GetControlCountOnTabPage(TabControlIndexType aTabControlIndex, TabIndexType aTabIndex)
{
	int count = 0;
	for (GuiIndexType u = 0; u < mControlCount; ++u)
		if (mControl[u]->tab_index == aTabIndex && mControl[u]->tab_control_index == aTabControlIndex) // This boolean order helps performance.
			++count;
	return count;
}



void GuiType::GetTabDisplayAreaRect(HWND aTabControlHwnd, RECT &aRect)
// Gets coordinates of tab control's display area relative to parent window's client area.
{
	RECT rect;
	GetClientRect(aTabControlHwnd, &rect); // Used because the coordinates of its upper-left corner are (0,0).
	DWORD style = GetWindowLong(aTabControlHwnd, GWL_STYLE);
	if (style & TCS_BUTTONS) // Workarounds are needed for TCS_BUTTONS.
	{
		RECT item_rect;
		TabCtrl_GetItemRect(aTabControlHwnd, 0, &item_rect);
		int row_count = TabCtrl_GetRowCount(aTabControlHwnd);
		const int offset = 3; // The margin between buttons seems to be 3 pixels.
		if (style & TCS_VERTICAL)
		{
			// TabCtrl_AdjustRect would not work for these cases.
			int width = (item_rect.right - item_rect.left + offset) * row_count;
			if (style & TCS_RIGHT)
				rect.right -= width;
			else
				rect.left += width;
		}
		else
		{
			int height = (item_rect.bottom - item_rect.top + offset) * row_count;
			if (style & TCS_BOTTOM)
				// TabCtrl_AdjustRect would not work for this case.
				rect.bottom -= height;
			else
				// TabCtrl_AdjustRect plus an adjustment of bottom += 3 would work,
				// but it's easier to handle this case alongside TCS_BOTTOM (above).
				rect.top += height;
		}
	}
	else
	{
		TabCtrl_AdjustRect(aTabControlHwnd, FALSE, &rect); // Retrieve the area beneath the tabs.
		rect.left -= 2;  // Testing shows that X (but not Y) is off by exactly 2.
	}
	MapWindowPoints(aTabControlHwnd, mHwnd, (LPPOINT)&rect, 2); // Convert tab-control-relative coordinates to Gui-relative.
	aRect.left = rect.left;
	aRect.right = rect.right;
	aRect.top = rect.top;
	aRect.bottom = rect.bottom;
}



POINT GuiType::GetPositionOfTabDisplayArea(GuiControlType &aTabControl)
// Gets position of tab control's display area relative to parent window's client area.
{
	RECT rect;
	GetTabDisplayAreaRect(aTabControl.hwnd, rect);
	POINT pt = { rect.left, rect.top };
	return pt;
}



ResultType GuiType::SelectAdjacentTab(GuiControlType &aTabControl, bool aMoveToRight, bool aFocusFirstControl
	, bool aWrapAround)
{
	int tab_count = TabCtrl_GetItemCount(aTabControl.hwnd);
	if (!tab_count)
		return FAIL;

	int selected_tab = TabCtrl_GetCurSel(aTabControl.hwnd);
	if (selected_tab == -1) // Not sure how this can happen in this case (since it has at least one tab).
		selected_tab = aMoveToRight ? 0 : tab_count - 1; // Select the first or last tab.
	else
	{
		if (aMoveToRight) // e.g. Ctrl-PgDn or Ctrl-Tab, right-arrow
		{
			++selected_tab;
			if (selected_tab >= tab_count) // wrap around to the start
			{
				if (!aWrapAround)
					return FAIL; // Indicate that tab was not selected due to non-wrap.
				selected_tab = 0;
			}
		}
		else // Ctrl-PgUp or Ctrl-Shift-Tab
		{
			--selected_tab;
			if (selected_tab < 0) // wrap around to the end
			{
				if (!aWrapAround)
					return FAIL; // Indicate that tab was not selected due to non-wrap.
				selected_tab = tab_count - 1;
			}
		}
	}
	// MSDN: "A tab control does not send a TCN_SELCHANGING or TCN_SELCHANGE notification message
	// when a tab is selected using the TCM_SETCURSEL message."
	TabCtrl_SetCurSel(aTabControl.hwnd, selected_tab);
	ControlUpdateCurrentTab(aTabControl, aFocusFirstControl);

	// Fix for v1.0.35: Keyboard navigation of a tab control should still launch the tab's event handler
	// if it has one:
	Event(GUI_HWND_TO_INDEX(aTabControl.hwnd), TCN_SELCHANGE, GUI_EVENT_CHANGE);

	return OK;
}



void GuiType::AutoSizeTabControl(GuiControlType &aTabControl)
{
	int autosize = (int)(INT_PTR)GetProp(aTabControl.hwnd, _T("ahk_autosize"));
	if (!autosize)
		return;
	RemoveProp(aTabControl.hwnd, _T("ahk_autosize"));

	// Below: MUST NOT initialize to zero, since these are screen coordinates
	// and it is possible to position the GUI prior to this function being called.
	int tab_index = aTabControl.tab_index, right = INT_MIN, bottom = INT_MIN;
	RECT rc;

	for (GuiIndexType i = 0; i < mControlCount; ++i)
	{
		if (mControl[i]->tab_control_index != tab_index)
			continue;
		GetWindowRect(mControl[i]->hwnd, &rc);
		if (right < rc.right)
			right = rc.right;
		if (bottom < rc.bottom)
			bottom = rc.bottom;
	}

	// Testing shows that -32768 is the minimum screen coordinate, but even if
	// that's not always the case, a right edge of INT_MIN would mean that the
	// controls have 0 width, since window positions are 32-bit at most.
	BOOL at_least_one_control = right != INT_MIN;

	RECT tab_rect;
	GetWindowRect(aTabControl.hwnd, &tab_rect);
	if ((autosize & TAB3_AUTOWIDTH) && at_least_one_control)
		tab_rect.right = right + mMarginX + 4;
	if ((autosize & TAB3_AUTOHEIGHT) && at_least_one_control)
		tab_rect.bottom = bottom + mMarginY + 4;
	MapWindowPoints(NULL, mHwnd, (LPPOINT)&tab_rect, 2); // Screen to GUI client.
	int width = tab_rect.right - tab_rect.left
		, height = tab_rect.bottom - tab_rect.top;

	DWORD style = GetWindowLong(aTabControl.hwnd, GWL_STYLE);
	
	// The number of "rows" of vertical buttons affect the tab control's width,
	// while horizontal buttons affect its height; so check the appropriate flag:
	bool adjust_for_rows = 0 != (autosize & ((style & TCS_VERTICAL) ? TAB3_AUTOWIDTH : TAB3_AUTOHEIGHT));
	// For TCS_RIGHT, we need to adjust even when there's only one row before and after:
	int rows_before = (!adjust_for_rows || (style & TCS_RIGHT)) ? 0 : TabCtrl_GetRowCount(aTabControl.hwnd);
	
	// Apply new size.
	MoveWindow(aTabControl.hwnd, tab_rect.left, tab_rect.top, width, height, TRUE);

	if (adjust_for_rows)
	{
		int rows_after = TabCtrl_GetRowCount(aTabControl.hwnd);
		if (rows_before != rows_after)
		{
			// Adjust for the newly added tab rows/columns.
			RECT item_rect;
			TabCtrl_GetItemRect(aTabControl.hwnd, 0, &item_rect);
			int offset = (style & TCS_BUTTONS) ? 3 : 0; // The margin between buttons seems to be 3 pixels.
			if (style & TCS_VERTICAL) // Tabs on left/right.
			{
				width += (rows_after - rows_before) * (item_rect.right - item_rect.left + offset);
				tab_rect.right = tab_rect.left + width;
			}
			else // Tabs on top/bottom.
			{
				height += (rows_after - rows_before) * (item_rect.bottom - item_rect.top + offset);
				tab_rect.bottom = tab_rect.top + height;
			}
			// Apply size adjustment.
			MoveWindow(aTabControl.hwnd, tab_rect.left, tab_rect.top, width, height, TRUE);
		}
	}
	
	if (mControl[mControlCount-1]->tab_control_index == tab_index // The previous control belongs to this tab control.
		|| !at_least_one_control) // Empty Tab3. See below.
	{
		// Position next control relative to the tab control rather than its last child control.
		mPrevX = tab_rect.left;
		mPrevY = tab_rect.top;
		mPrevWidth = width;
		mPrevHeight = height;
		// This is used by some positioning modes, but is also required for auto-sizing of the Gui
		// to account for the auto-sized Tab3 control (even if it has retained its default size):
		if (mMaxExtentRight < tab_rect.right)
			mMaxExtentRight = tab_rect.right;
		if (mMaxExtentDown < tab_rect.bottom)
			mMaxExtentDown = tab_rect.bottom;
	}
}



ResultType GuiType::CreateTabDialog(GuiControlType &aTabControl, GuiControlOptionsType &aOpt)
// Creates a dialog to host controls for a tab control.
{
	HWND tab_dialog;
	struct MYDLG : DLGTEMPLATE
	{
		WORD wNoMenu;
		WORD wStdClass;
		WORD wNoTitle;
	};
	MYDLG dlg;
	ZeroMemory(&dlg, sizeof(dlg));
	dlg.style = WS_VISIBLE | WS_CHILD | DS_CONTROL; // DS_CONTROL enables proper tab navigation and possibly other things.
	//dlg.dwExtendedStyle = WS_EX_CONTROLPARENT; // Redundant; DS_CONTROL causes this to be applied.
	// Create the dialog.
	tab_dialog = CreateDialogIndirect(g_hInstance, &dlg, mHwnd, TabDialogProc);
	if (!tab_dialog)
		return FAIL;
	if (!SetProp(aTabControl.hwnd, _T("ahk_dlg"), tab_dialog))
	{
		DestroyWindow(tab_dialog);
		return FAIL;
	}
	LONG attrib = aTabControl.tab_index;
	if (aOpt.use_theme)
	{
		// Apply a tab control background to the tab dialog.  This also affects
		// the backgrounds of some controls, such as Radio and Groupbox controls.
		EnableThemeDialogTexture(tab_dialog, 6); // ETDT_ENABLETAB = 6
		attrib |= TABDIALOG_ATTRIB_THEMED;
	}
	SetWindowLongPtr(tab_dialog, GWLP_USERDATA, attrib);
	if (!((mExStyle = GetWindowLong(mHwnd, GWL_EXSTYLE)) & WS_EX_CONTROLPARENT))
		// WS_EX_CONTROLPARENT must be applied to the parent window as well as the
		// tab dialog to allow tab navigation to work correctly.  With it applied
		// only to the tab dialog, tabbing out of a multi-line edit control won't
		// work if it is the last control on a tab.
		SetWindowLong(mHwnd, GWL_EXSTYLE, mExStyle |= WS_EX_CONTROLPARENT);
	return OK;
}



void GuiType::UpdateTabDialog(HWND aTabControlHwnd)
// Positions a tab control's container dialog to fill the tab control's display area.
{
	HWND tab_dialog = (HWND)GetProp(aTabControlHwnd, _T("ahk_dlg"));
	if (tab_dialog) // It's a Tab3 control.
	{
		RECT rect;
		GetTabDisplayAreaRect(aTabControlHwnd, rect);
		MoveWindow(tab_dialog, rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top, TRUE);
	}
}



HBRUSH CreateTabDialogBrush(HWND aDlg, HDC aCompatibleDC)
{
	HBRUSH brush = NULL;
	RECT rect;
	GetClientRect(aDlg, &rect);
	if (HDC bitmap_dc = CreateCompatibleDC(aCompatibleDC))
	{
		if (HBITMAP bitmap = CreateCompatibleBitmap(aCompatibleDC, rect.right, rect.bottom))
		{
			HBITMAP prev_bitmap = (HBITMAP)SelectObject(bitmap_dc, bitmap);

			// Print the tab control into the bitmap.
			SendMessage(aDlg, WM_PRINTCLIENT, (WPARAM)bitmap_dc, PRF_CLIENT);

			// Create a brush from the bitmap.
			brush = CreatePatternBrush(bitmap);

			SelectObject(bitmap_dc, prev_bitmap);
			DeleteObject(bitmap);
		}
		DeleteDC(bitmap_dc);
	}
	return brush;
}



INT_PTR CALLBACK TabDialogProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	// Messages to be forwarded to the main GUI window:
	case WM_COMMAND:
	case WM_NOTIFY:
	case WM_VSCROLL: // These two should only be received for sliders and up-downs.
	case WM_HSCROLL: //
	case WM_CONTEXTMENU:
	//case WM_DROPFILES: // Not sent to the tab dialog since it doesn't accept drop.
	case WM_CTLCOLOREDIT:
	case WM_CTLCOLORLISTBOX:
	case WM_CTLCOLORSTATIC:
	case WM_CTLCOLORBTN: // Not used directly, but might be used by script.
	case WM_CTLCOLORSCROLLBAR: // As above.
		if (GuiType *pgui = GuiType::FindGui(GetParent(hDlg)))
		{
			// For Slider and Link controls to look right in a themed tab control, this custom
			// handling is needed.  This also handles other controls for simplicity, although
			// there are easier methods available for them:
			//  - Text/Picture: BackgroundTrans works (and overrides this section)
			//  - Check/Radio/Group: EnableThemeDialogTexture() is enough
			// Unlike IsThemeActive, IsAppThemed will return false if the "Disable visual styles"
			// compatibility setting is turned on for this script's EXE.
			GuiControlType *pcontrol;
			if (   uMsg == WM_CTLCOLORSTATIC
				&& (TABDIALOG_ATTRIB_THEMED & GetWindowLongPtr(hDlg, GWLP_USERDATA))
				&& (pcontrol = pgui->FindControl((HWND)lParam))
				&& (pcontrol->background_color != CLR_TRANSPARENT)
				&& IsAppThemed()   )
			{
				HDC hdc = (HDC)wParam;
				HBRUSH brush = (HBRUSH)GetProp(hDlg, _T("br"));
				if (!brush)
				{
					// Create a pattern brush based on a bitmap of this tab dialog's background.
					if (brush = CreateTabDialogBrush(hDlg, hdc))
						SetProp(hDlg, _T("br"), brush);
				}
				if (brush)
				{
					if (pcontrol->type != GUI_CONTROL_PIC && pcontrol->union_color != CLR_DEFAULT)
						SetTextColor((HDC)wParam, pcontrol->union_color);
					else
						SetTextColor((HDC)wParam, GetSysColor(COLOR_WINDOWTEXT));

					// Tell the control to draw its background using our pattern brush.
					POINT pt = { 0,0 };
					MapWindowPoints(hDlg, (HWND)lParam, &pt, 1);
					SetBrushOrgEx(hdc, pt.x, pt.y, NULL);
					SetBkMode(hdc, TRANSPARENT); // Affects text rendering.
					return (INT_PTR)brush;
				}
			}
			// Forward this message to the GUI:
			LRESULT result = SendMessage(pgui->mHwnd, uMsg, wParam, lParam);
			if (uMsg >= WM_CTLCOLOREDIT && uMsg <= WM_CTLCOLORSTATIC)
				return (INT_PTR)result;
			// Except for WM_CTLCOLOR messages and a few others, DialogProcs are expected to
			// return more specific results with SetWindowLong.
			SetWindowLongPtr(hDlg, DWLP_MSGRESULT, result);
			return TRUE;
		}
		break;
	// Messages handled by the tab dialog:
	case WM_CTLCOLORDLG:
		if (GuiType *pgui = GuiType::FindGui(GetParent(hDlg)))
		{
			// Apply the tab control's background color (which might be inherited from the Gui)
			// if any.  A non-NULL brush implies -Theme, even if the background was inherited.
			LONG attrib = (LONG)GetWindowLongPtr(hDlg, GWLP_USERDATA);
			GuiControlType *ptab_control = pgui->FindTabControl(TABDIALOG_ATTRIB_INDEX(attrib));
			HBRUSH bk_brush;
			COLORREF bk_color;
			pgui->ControlGetBkColor(*ptab_control, TRUE, bk_brush, bk_color);
			if (bk_brush)
			{
				SetBkColor((HDC)wParam, bk_color);
				return (INT_PTR)bk_brush;
			}
		}
		break;
	case WM_WINDOWPOSCHANGED: // Recreate the brush when the tab dialog is resized.
	case WM_DESTROY: // Delete the brush when the window is destroyed.
		if (HBRUSH brush = (HBRUSH)GetProp(hDlg, _T("br"))) // br was set by WM_CTLCOLORSTATIC above.
		{
			RemoveProp(hDlg, _T("br"));
			DeleteObject(brush);
		}
		break;
	}
	return FALSE;
}



void GuiType::ControlGetPosOfFocusedItem(GuiControlType &aControl, POINT &aPoint)
// Caller has ensured that aControl is the focused control if the window has one.  If not,
// aControl can be any other control.
// Based on the control type, the position of the focused subitem within the control is
// returned (in screen coords) If the control has no focused item, the position of the
// control's caret (which seems to work okay on all control types, even pictures) is returned.
{
	LRESULT index;
	RECT rect;
	rect.left = COORD_UNSPECIFIED; // Init to detect whether rect has been set yet.

	switch (aControl.type)
	{
	case GUI_CONTROL_LISTBOX: // Testing shows that GetCaret() doesn't report focused row's position.
		index = SendMessage(aControl.hwnd, LB_GETCARETINDEX, 0, 0); // Testing shows that only one item at a time can have focus, even when multiple items are selected.
		if (index != LB_ERR) //  LB_ERR == -1
			SendMessage(aControl.hwnd, LB_GETITEMRECT, index, (LPARAM)&rect);
		// If above didn't get the rect for either reason, a default method is used later below.
		break;

	case GUI_CONTROL_LISTVIEW: // Testing shows that GetCaret() doesn't report focused row's position.
		index = ListView_GetNextItem(aControl.hwnd, -1, LVNI_FOCUSED); // Testing shows that only one item at a time can have focus, even when multiple items are selected.
		if (index != -1)
		{
			// If the focused item happens to be beneath the viewable area, the context menu gets
			// displayed beneath the ListView, but this behavior seems okay because of the rarity
			// and because Windows Explorer behaves the same way.
			// Don't use the ListView_GetItemRect macro in this case (to cut down on its code size).
			rect.left = LVIR_LABEL; // Seems better than LVIR_ICON in case icon is on right vs. left side of item.
			SendMessage(aControl.hwnd, LVM_GETITEMRECT, index, (LPARAM)&rect);
		}
		//else a default method is used later below, flagged by rect.left==COORD_UNSPECIFIED.
		break;

	case GUI_CONTROL_TREEVIEW: // Testing shows that GetCaret() doesn't report focused row's position.
		HTREEITEM hitem;
		if (hitem = TreeView_GetSelection(aControl.hwnd)) // Same as SendMessage(aControl.hwnd, TVM_GETNEXTITEM, TVGN_CARET, NULL).
			// If the focused item happens to be beneath the viewable area, the context menu gets
			// displayed beneath the ListView, but this behavior seems okay because of the rarity
			// and because Windows Explorer behaves the same way.
			// Don't use the ListView_GetItemRect macro in this case (to cut down on its code size).
			TreeView_GetItemRect(aControl.hwnd, hitem, &rect, TRUE); // Pass TRUE because caller typically wants to display a context menu, and this gives a more precise location for it.
		//else a default method is used later below, flagged by rect.left==COORD_UNSPECIFIED.
		break;

	case GUI_CONTROL_SLIDER: // GetCaretPos() doesn't retrieve thumb position, so it seems best to do so in case slider is very tall or long.
		SendMessage(aControl.hwnd, TBM_GETTHUMBRECT, 0, (WPARAM)&rect); // No return value.
		break;
	}

	// Notes about control types not handled above:
	//case GUI_CONTROL_STATUSBAR: For this and many others below, caller should never call it for this type.
	//case GUI_CONTROL_TEXT:
	//case GUI_CONTROL_LINK:
	//case GUI_CONTROL_PIC:
	//case GUI_CONTROL_GROUPBOX:
	//case GUI_CONTROL_BUTTON:
	//case GUI_CONTROL_CHECKBOX:
	//case GUI_CONTROL_RADIO:
	//case GUI_CONTROL_DROPDOWNLIST:
	//case GUI_CONTROL_COMBOBOX:
	//case GUI_CONTROL_EDIT:         Has it's own context menu.
	//case GUI_CONTROL_DATETIME:
	//case GUI_CONTROL_MONTHCAL:     Has it's own context menu. Can't be focused anyway.
	//case GUI_CONTROL_HOTKEY:
	//case GUI_CONTROL_UPDOWN:
	//case GUI_CONTROL_PROGRESS:
	//case GUI_CONTROL_TAB:  For simplicity, just do basic reporting rather than trying to find pos. of focused tab.

	if (rect.left == COORD_UNSPECIFIED) // Control's rect hasn't yet been fetched, so fall back to default method.
	{
		GetWindowRect(aControl.hwnd, &rect);
		// Decided against this since it doesn't seem to matter for any current control types.  If a custom
		// context menu is ever supported for Edit controls, maybe use GetCaretPos() for them.
		//GetCaretPos(&aPoint); // For some control types, this might give a more precise/appropriate position than GetWindowRect().
		//ClientToScreen(aControl.hwnd, &aPoint);
		//// A little nicer for most control types (such as DateTime) to shift it down a little so that popup/context
		//// menu or tooltip doesn't fully obstruct the control's contents.
		//aPoint.y += 10;  // A constant 10 is used because varying it by font doesn't seem worthwhile given that menu/tooltip fonts are of fixed size.
	}
	else
		MapWindowPoints(aControl.hwnd, NULL, (LPPOINT)&rect, 2); // Convert rect from client coords to screen coords.

	aPoint.x = rect.left;
	aPoint.y = rect.top + 2 + (rect.bottom - rect.top)/2;  // +2 to shift it down a tad, revealing more of the selected item.
	// Above: Moving it down a little by default seems desirable 95% of the time to prevent it
	// from covering up the focused row, the slider's thumb, a datetime's single row, etc.
}



struct LV_SortType
{
	LVFINDINFO lvfi;
	LVITEM lvi;
	HWND hwnd;
	lv_col_type col;
	TCHAR buf1[LV_TEXT_BUF_SIZE];
	TCHAR buf2[LV_TEXT_BUF_SIZE];
	bool sort_ascending;
	bool incoming_is_index;
};



int CALLBACK LV_GeneralSort(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
// ListView sorting by field's text or something derived from the text for each call.
{
	LV_SortType &lvs = *(LV_SortType *)lParamSort;

	// v1.0.44.12: Testing shows that LVM_GETITEMW automatically converts the ANSI contents of our ListView
	// into Unicode, which is nice because it avoids the overhead and code size of having to call
	// ToWideChar(), along with the extra/temp buffers it requires to receive the wide version.
	UINT msg_lvm_getitem = (lvs.col.case_sensitive == SCS_INSENSITIVE_LOGICAL && lvs.col.type == LV_COL_TEXT)
		? LVM_GETITEMW : LVM_GETITEM; // Both items above are checked so that SCS_INSENSITIVE_LOGICAL can be effect even for non-text columns because it allows a column to be later changed to TEXT and retain its "logical-sort" setting.
	// NOTE: It's safe to send a LVITEM struct rather than an LVITEMW with the LVM_GETITEMW message because
	// the only difference between them is the type "LPWSTR pszText", which is no problem as long as caller
	// has properly halved cchTextMax to reflect that wide-chars are twice as wide as 8-bit characters.

	// MSDN: "During the sorting process, the list-view contents are unstable. If the [ListView_SortItems]
	// callback function sends any messages to the list-view control, the results are unpredictable (aside
	// from LVM_GETITEM, which is allowed by ListView_SortItemsEx but not ListView_SortItems)."
	// Since SortItemsEx has become so much more common/available, the doubt about whether the non-Ex
	// ListView_SortItems actually allows LVM_GETITEM (which it probably does in spite of not being
	// documented) much less of a concern.
	// Older: It seems hard to believe that you shouldn't send ANY kind of message because how could you
	// ever use ListView_SortItems() without either having LVS_OWNERDATA or allocating temp memory for the
	// entire column (to whose rows lParam would point)?
	// UPDATE: The following seems to be one alternative:
	// Do a "virtual qsort" on this column's contents by having qsort() sort an array of row numbers according
	// to the contents of each particular row's field in that column (i.e. qsort's callback would call LV_GETITEM).
	// In other words, the array would start off in order (1,2,3) but afterward would contain the proper sort
	// (e.g. 3,1,2). Next, traverse the array and store the correct "order number" in the corresponding row's
	// special "lParam container" (for example, 3,1,2 would store 1 in row 3, 2 in row 1, and 3 in row 2, and 4 in...).
	// Then the ListView can be sorted via a method like the high performance LV_Int32Sort.
	// However, since the above would require TWO SORTS, it would probably be slower (though the second sort would
	// require only a tiny fraction of the time of the first).
	lvs.lvi.pszText = lvs.buf1; // lvi's other members were already set by the caller.
	if (lvs.incoming_is_index) // Serves to avoid the potentially high performance overhead of ListView_FindItem() where possible.
	{
		lvs.lvi.iItem = (int)lParam1;
		SendMessage(lvs.hwnd, msg_lvm_getitem, 0, (LPARAM)&lvs.lvi); // Use LVM_GETITEM vs. LVM_GETITEMTEXT because MSDN says that only LVM_GETITEM is safe during the sort.
	}
	else
	{
		// Unfortunately, lParam cannot be used as the index itself because apparently, the sorting
		// process puts the item indices into a state of flux.  In other words, the indices are
		// changing while the sort progresses, so it's not possible to use an item's original index
		// as a way to uniquely identify it.
		lvs.lvfi.lParam = lParam1;
		lvs.lvi.iItem = ListView_FindItem(lvs.hwnd, -1, &lvs.lvfi);
		if (lvs.lvi.iItem < 0) // Not found.  Impossible if caller set the LParam to a unique value.
			*lvs.buf1 = '\0';
		else
			SendMessage(lvs.hwnd, msg_lvm_getitem, 0, (LPARAM)&lvs.lvi);
	}

	// Must use lvi.pszText vs. buf because MSDN says (for LVM_GETITEM, but it might also apply to
	// LVM_GETITEMTEXT even though it isn't documented): "Applications should not assume that the text will
	// necessarily be placed in the specified buffer. The control may instead change the pszText member
	// of the structure to point to the new text rather than place it in the buffer."
	LPTSTR field1 = lvs.lvi.pszText; // Save value of pszText in case it no longer points to lvs.buf1.

	// Fetch Item #2 (see comments in #1 above):
	lvs.lvi.pszText = lvs.buf2; // lvi's other members were already set by the caller.
	if (lvs.incoming_is_index)
	{
		lvs.lvi.iItem = (int)lParam2;
		SendMessage(lvs.hwnd, msg_lvm_getitem, 0, (LPARAM)&lvs.lvi); // Use LVM_GETITEM vs. LVM_GETITEMTEXT because MSDN says that only LVM_GETITEM is safe during the sort.
	}
	else
	{
		// Set any lvfi members not already set by the caller.  Note that lvi.mask was set to LVIF_TEXT by the caller.
		lvs.lvfi.lParam = lParam2;
		lvs.lvi.iItem = ListView_FindItem(lvs.hwnd, -1, &lvs.lvfi);
		if (lvs.lvi.iItem < 0) // Not found.  Impossible if caller set the LParam to a unique value.
			*lvs.buf2 = '\0';
		else
			SendMessage(lvs.hwnd, msg_lvm_getitem, 0, (LPARAM)&lvs.lvi);
	}

	// MSDN: "return a negative value if the first item should precede the second"
	int result;
	if (lvs.col.type == LV_COL_TEXT)
	{
		if (lvs.col.case_sensitive == SCS_INSENSITIVE_LOGICAL)
			result = StrCmpLogicalW((LPCWSTR)field1, (LPCWSTR)lvs.lvi.pszText);
		else
			result = tcscmp2(field1, lvs.lvi.pszText, lvs.col.case_sensitive); // Must not refer to buf1/buf2 directly, see above.
	}
	else
	{
		// Unlike ACT_SORT, supporting hex for an explicit-floating point column seems far too rare to
		// justify, hence _tstof() is used vs. ATOF().  v1.0.46.03: Fixed to sort properly (formerly, it just case the difference to an int, which isn't right).
		double f1 = _tstof(field1), f2 = _tstof(lvs.lvi.pszText); // Must not refer to buf1/buf2 directly, see above.
		result = (f1 > f2) ? 1 : (f1 == f2 ? 0 : -1);
	}
	return lvs.sort_ascending ? result : -result;
}



int CALLBACK LV_Int32Sort(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	// Caller-provided value of lParamSort is TRUE (non-zero) when ascending order is desired.
	// MSDN: "return a negative value if the first item should precede the second"
	return (int)(lParamSort ? (lParam1 - lParam2) : (lParam2 - lParam1));
}



void GuiType::LV_Sort(GuiControlType &aControl, int aColumnIndex, bool aSortOnlyIfEnabled, TCHAR aForceDirection)
// aForceDirection should be 'A' to force ascending, 'D' to force ascending, or '\0' to use the column's
// current default direction.
{
	if (aColumnIndex < 0 || aColumnIndex >= LV_MAX_COLUMNS) // Invalid (avoids array access violation).
		return;
	lv_attrib_type &lv_attrib = *aControl.union_lv_attrib;
	lv_col_type &col = lv_attrib.col[aColumnIndex];

	int item_count = ListView_GetItemCount(aControl.hwnd);
	if ((col.sort_disabled && aSortOnlyIfEnabled) || item_count < 2) // This column cannot be sorted or doesn't need to be.
		return; // Below relies on having returned here when control is empty or contains 1 item.

	// Init any lvs members that are needed by both LV_Int32Sort and the other sorting functions.
	// The new sort order is determined by the column's primary order unless the user clicked the current
	// sort-column, in which case the direction is reversed (unless the column is unidirectional):
	LV_SortType lvs;
	if (aForceDirection)
		lvs.sort_ascending = (aForceDirection == 'A');
	else
		lvs.sort_ascending = (aColumnIndex == lv_attrib.sorted_by_col && !col.unidirectional)
			? !lv_attrib.is_now_sorted_ascending : !col.prefer_descending;

	// Init those members needed for LVM_GETITEM if it turns out to be needed.  This section
	// also serves to permanently init cchTextMax for use by the sorting functions too:
	lvs.lvi.pszText = lvs.buf1;
	lvs.lvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // Set default. Subtracts 1 because of that nagging doubt about size vs. length. Some MSDN examples subtract one, such as TabCtrl_GetItem()'s cchTextMax.

	if (col.type == LV_COL_INTEGER)
	{
		// Testing indicates that the following approach is 25 times faster than the general-sort method.
		// Assign the 32-bit integer as the items lParam at this early stage rather than getting the text
		// and converting it to an integer for every call of the sort proc.
		for (lvs.lvi.lParam = 0, lvs.lvi.iItem = 0; lvs.lvi.iItem < item_count; ++lvs.lvi.iItem)
		{
			lvs.lvi.mask = LVIF_TEXT;
			lvs.lvi.iSubItem = aColumnIndex; // Which field to fetch (must be reset each time since it's set to 0 below).
			lvs.lvi.lParam = SendMessage(aControl.hwnd, LVM_GETITEM, 0, (LPARAM)&lvs.lvi)
				? ATOI(lvs.lvi.pszText) : 0; // Must not refer to lvs.buf1 directly because MSDN says LVM_GETITEMTEXT might have changed pszText to point to some other string.
			lvs.lvi.mask = LVIF_PARAM;
			lvs.lvi.iSubItem = 0; // Indicate that an item vs. subitem is being operated on (subitems can't have an lParam).
			ListView_SetItem(aControl.hwnd, &lvs.lvi);
		}
		// Always use non-Ex() for this one because it's likely to perform best due to the lParam setup above.
		// The value of iSubItem is not reset to aColumnIndex because LV_Int32Sort() doesn't use it.
        SendMessage(aControl.hwnd, LVM_SORTITEMS, lvs.sort_ascending, (LPARAM)LV_Int32Sort); // Should always succeed since it uses non-Ex() version.
	}
	else // It's LV_COL_TEXT or LV_COL_FLOAT.
	{
		if (col.type == LV_COL_TEXT && col.case_sensitive == SCS_INSENSITIVE_LOGICAL) // SCS_INSENSITIVE_LOGICAL can be in effect even when type isn't LV_COL_TEXT because it allows a column to be later changed to TEXT and retain its "logical-sort" setting.
		{
			// v1.0.44.12: Support logical sorting, which treats numeric strings as true numbers like Windows XP
			// Explorer's sorting.  This is done here rather than in LV_ModifyCol() because it seems more
			// maintainable/robust (plus LV_GeneralSort() relies on us to do this check).
			lvs.lvi.cchTextMax = lvs.lvi.cchTextMax/2 - 1; // Buffer can hold only half as many Unicode characters as non-Unicode (subtract 1 for the extra-wide NULL terminator).
		}
		// Since LVM_SORTITEMSEX requires comctl32.dll version 5.80+, the non-Ex version is used
		// whenever the EX version fails to work.  One reason to strongly prefer the Ex version
		// is that MSDN says the non-Ex version shouldn't query the control during the sort,
		// which although hard to believe, is a concern.  Therefore:
		// Try to use the SortEx() method first. If it doesn't work, fall back to the non-Ex method under
		// the assumption that the OS doesn't support the Ex() method.
		// Initialize struct members as much as possible so that the sort callback function doesn't have to do it
		// the many times it's called. Some of the others were already initialized higher above for internal use here.
		lvs.hwnd = aControl.hwnd;
		lvs.lvi.iSubItem = aColumnIndex; // Zero-based column index to indicate whether the item or one of its sub-items should be retrieved.
		lvs.col = col; // Struct copy, which should enhance sorting performance over a pointer.
		lvs.incoming_is_index = true;
		lvs.lvi.pszText = NULL; // Serves to detect whether the sort-proc actually ran (it won't if this is Win95 or some other OS that lacks SortEx).
		lvs.lvi.mask = LVIF_TEXT;
		SendMessage(aControl.hwnd, LVM_SORTITEMSEX, (WPARAM)&lvs, (LPARAM)LV_GeneralSort);
		if (!lvs.lvi.pszText)
		{
			// Since SortEx() didn't run (above has already ensured that this wasn't because the control is empty),
			// fall back to Sort() method.
			// Use a simple sequential lParam to guarantee uniqueness. This must be done every time in case
			// rows have been inserted/deleted since the last time, in which case uniqueness would not be
			// certain otherwise:
			lvs.lvi.iSubItem = 0; // Indicate that an item vs. subitem is to be updated (subitems can't have an lParam).
			lvs.lvi.mask = LVIF_PARAM; // Indicate which member is to be updated.
			for (lvs.lvi.lParam = 0, lvs.lvi.iItem = 0; lvs.lvi.iItem < item_count; ++lvs.lvi.lParam, ++lvs.lvi.iItem)
				ListView_SetItem(aControl.hwnd, &lvs.lvi);
			// Initialize struct members as much as possible so that the sort callback function doesn't have to do it
			// each time it's called.   Some of the others were already initialized higher above for internal use here.
			lvs.incoming_is_index = false;
			lvs.lvfi.flags = LVFI_PARAM; // This is the find-method; i.e. the sort function will find each item based on its LPARAM.
			lvs.lvi.mask = LVIF_TEXT; // Sort proc. uses LVIF_TEXT internally because the PARAM mask only applies to lvfi vs. lvi.
			lvs.lvi.iSubItem = aColumnIndex; // Zero-based column index to indicate whether the item or one of its sub-items should be retrieved.
			SendMessage(aControl.hwnd, LVM_SORTITEMS, (WPARAM)&lvs, (LPARAM)LV_GeneralSort);
		}
	}

	// For simplicity, ListView_SortItems()'s return value (TRUE/FALSE) is ignored since it shouldn't
	// realistically fail.  Just update things to indicate the current sort-column and direction:
	lv_attrib.sorted_by_col = aColumnIndex;
	lv_attrib.is_now_sorted_ascending = lvs.sort_ascending;
}



#define MAX_ACCELERATORS 128

void GuiType::UpdateAccelerators(UserMenu &aMenu)
{
	// Destroy the old accelerator table, if any.
	RemoveAccelerators();

	int accel_count = 0;
	ACCEL accel[MAX_ACCELERATORS];

	// Call recursive function to handle submenus.
	UpdateAccelerators(aMenu, accel, accel_count);

	if (accel_count)
		mAccel = CreateAcceleratorTable(accel, accel_count);
}

void GuiType::UpdateAccelerators(UserMenu &aMenu, LPACCEL aAccel, int &aAccelCount)
{
	UserMenuItem *item;
	for (item = aMenu.mFirstMenuItem; item && aAccelCount < MAX_ACCELERATORS; item = item->mNextMenuItem)
	{
		if (item->mSubmenu)
		{
			// Recursively process accelerators in submenus.
			UpdateAccelerators(*item->mSubmenu, aAccel, aAccelCount);
		}
		else if (LPTSTR tab = _tcschr(item->mName, '\t'))
		{
			if (ConvertAccelerator(tab + 1, aAccel[aAccelCount]))
			{
				// This accelerator is valid.
				aAccel[aAccelCount++].cmd = item->mMenuID;
			}
			// Otherwise, the text following '\t' was an invalid accelerator or not an accelerator
			// at all. For simplicity, flexibility and backward-compatibility, just ignore it.
		}
	}
}

void GuiType::RemoveAccelerators()
{
	if (mAccel)
	{
		DestroyAcceleratorTable(mAccel);
		mAccel = NULL;
	}
}

bool GuiType::ConvertAccelerator(LPTSTR aString, ACCEL &aAccel)
{
	aString = omit_leading_whitespace(aString);
	if (!*aString) // Multiple points below rely on this check.
		return false;

	// Omitting FVIRTKEY and storing a character code instead of a vk code allows the
	// accelerator to be triggered by Alt+Numpad and perhaps IME or other means. However,
	// testing shows that it also causes the modifier key flags to be ignored, so this
	// approach can only be used if aString is a lone character:
	if (!aString[1])
	{
		aAccel.key = (WORD)(TBYTE)*aString;
		aAccel.fVirt = 0;
		return true;
	}

	aAccel.fVirt = FVIRTKEY; // Init.

	modLR_type modLR = 0;

	LPTSTR cp;
	while (cp = _tcschr(aString + 1, '+')) // For each modifier.  +1 is used in case "+" is the key.
	{
		LPTSTR cp_end = omit_trailing_whitespace(aString, cp - 1);
		size_t len = cp_end - aString + 1;
		if (!_tcsnicmp(aString, _T("Ctrl"), len))
			modLR |= MOD_LCONTROL;
		else if (!_tcsnicmp(aString, _T("Alt"), len))
			modLR |= MOD_LALT;
		else if (!_tcsnicmp(aString, _T("Shift"), len))
			modLR |= MOD_LSHIFT;
		else
			return false;
		aString = omit_leading_whitespace(cp + 1);
		if (!*aString) // Below relies on this check.
			return false; // "Modifier+", but no key name.
	}

	// Now that any modifiers have been parsed and removed, convert the key name.
	// If only a single character remains, this may set additional modifiers.
	if (!aString[1])
	{
		// It seems preferable to treat "Ctrl+O" as ^o and not ^+o, so convert the character
		// to lower-case so that MOD_LSHIFT is added only for keys like "@". Seems best to use
		// the locale-dependent method, though somewhat debatable.
		aAccel.key = CharToVKAndModifiers(ltolower(*aString), &modLR, GetKeyboardLayout(0));
	}
	else
		aAccel.key = TextToVK(aString);

	if (modLR & MOD_LCONTROL)
		aAccel.fVirt |= FCONTROL;
	if (modLR & MOD_LALT) // Not MOD_RALT, which would mean AltGr (unsupported).
		aAccel.fVirt |= FALT;
	if (modLR & MOD_LSHIFT)
		aAccel.fVirt |= FSHIFT;

	return aAccel.key; // i.e. false if not a valid key name.
}



void GuiType::SetDefaultMargins()
{
	// The original default mMarginX was documented as "1.25 times font-height", which is
	// confusing because font-height is in points while mMarginX is in pixels.  It's kept
	// at this value anyway since it seems to work well.  The calculation below uses a
	// factor relative to 72 vs. denominator of 96 to emulate pixel-to-point conversion.
	// 96 is used even when DPI != 100% so that the scaling performed by point-to-pixel
	// conversion is kept (lfHeight is already proportionate to screen DPI).
	if (mMarginX == COORD_UNSPECIFIED)
		mMarginX = MulDiv(sFont[mCurrentFontIndex].lfHeight, -90, 96); // Seems to be a good rule of thumb.  Originally 1.25 * point_size.
	if (mMarginY == COORD_UNSPECIFIED)
		mMarginY = MulDiv(sFont[mCurrentFontIndex].lfHeight, -54, 96); // Also seems good.  Originally 0.75 * point_size.
}
