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

#pragma once

#include "stdafx.h" // pre-compiled headers
#include "script.h"


#define GUI_INDEX_TO_ID(index) (index + CONTROL_ID_FIRST)
#define GUI_ID_TO_INDEX(id) (id - CONTROL_ID_FIRST) // Returns a small negative if "id" is invalid, such as 0.
#define GUI_HWND_TO_INDEX(hwnd) GUI_ID_TO_INDEX(GetDlgCtrlID(hwnd)) // Returns a small negative on failure (e.g. HWND not found).
// Notes about above:
// 1) Callers should call GuiType::FindControl() instead of GUI_HWND_TO_INDEX() if the hwnd might be a combobox's
//    edit control.
// 2) Testing shows that GetDlgCtrlID() is much faster than looping through a GUI window's control array to find
//    a matching HWND.


#define ERR_GUI_NO_WINDOW _T("Gui has no window.")
#define ERR_GUI_NOT_FOR_THIS_TYPE _T("Not supported for this control type.") // Used by GuiControl object and Control functions.


INT_PTR CALLBACK TabDialogProc(HWND hDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
#define TABDIALOG_ATTRIB_INDEX(a) (TabControlIndexType)(a & 0xFF)
#define TABDIALOG_ATTRIB_THEMED 0x100


typedef UCHAR GuiControls;
enum GuiControlTypes {GUI_CONTROL_INVALID // GUI_CONTROL_INVALID must be zero due to things like ZeroMemory() on the struct.
	, GUI_CONTROL_TEXT, GUI_CONTROL_PIC, GUI_CONTROL_GROUPBOX
	, GUI_CONTROL_BUTTON, GUI_CONTROL_CHECKBOX, GUI_CONTROL_RADIO
	, GUI_CONTROL_DROPDOWNLIST, GUI_CONTROL_COMBOBOX
	, GUI_CONTROL_LISTBOX, GUI_CONTROL_LISTVIEW, GUI_CONTROL_TREEVIEW
	, GUI_CONTROL_EDIT, GUI_CONTROL_DATETIME, GUI_CONTROL_MONTHCAL, GUI_CONTROL_HOTKEY
	, GUI_CONTROL_UPDOWN, GUI_CONTROL_SLIDER, GUI_CONTROL_PROGRESS, GUI_CONTROL_TAB, GUI_CONTROL_TAB2, GUI_CONTROL_TAB3
	, GUI_CONTROL_ACTIVEX, GUI_CONTROL_LINK, GUI_CONTROL_CUSTOM, GUI_CONTROL_STATUSBAR // Kept last to reflect it being bottommost in switch()s (for perf), since not too often used.
	, GUI_CONTROL_TYPE_COUNT};

#define GUI_CONTROL_TYPE_NAMES  _T(""), \
	_T("Text"), _T("Pic"), _T("GroupBox"), \
	_T("Button"), _T("CheckBox"), _T("Radio"), \
	_T("DDL"), _T("ComboBox"), \
	_T("ListBox"), _T("ListView"), _T("TreeView"), \
	_T("Edit"), _T("DateTime"), _T("MonthCal"), _T("Hotkey"), \
	_T("UpDown"), _T("Slider"), _T("Progress"), _T("Tab"), _T("Tab2"), _T("Tab3"), \
	_T("ActiveX"), _T("Link"), _T("Custom"), _T("StatusBar")


struct FontType : public LOGFONT
{
	HFONT hfont;
};

#define LV_TEXT_BUF_SIZE 8192  // Max amount of text in a ListView sub-item.  Somewhat arbitrary: not sure what the real limit is, if any.
enum LVColTypes {LV_COL_TEXT, LV_COL_INTEGER, LV_COL_FLOAT}; // LV_COL_TEXT must be zero so that it's the default with ZeroMemory.
struct lv_col_type
{
	UCHAR type;             // UCHAR vs. enum LVColTypes to save memory.
	bool sort_disabled;     // If true, clicking the column will have no automatic sorting effect.
	UCHAR case_sensitive;   // Ignored if type isn't LV_COL_TEXT.  SCS_INSENSITIVE is the default.
	bool unidirectional;    // Sorting cannot be reversed/toggled.
	bool prefer_descending; // Whether this column defaults to descending order (on first click or for unidirectional).
};

struct lv_attrib_type
{
	int sorted_by_col; // Index of column by which the control is currently sorted (-1 if none).
	bool is_now_sorted_ascending; // The direction in which the above column is currently sorted.
	bool no_auto_sort; // Automatic sorting disabled.
	#define LV_MAX_COLUMNS 200
	lv_col_type col[LV_MAX_COLUMNS];
	int col_count; // Number of columns currently in the above array.
	int row_count_hint;
};

typedef UCHAR TabControlIndexType;
typedef UCHAR TabIndexType;
// Keep the below in sync with the size of the types above:
#define MAX_TAB_CONTROLS 255  // i.e. the value 255 itself is reserved to mean "doesn't belong to a tab".
#define MAX_TABS_PER_CONTROL 256
struct GuiControlType : public Object
{
	GuiType* gui;
	HWND hwnd = NULL;
	LPTSTR name = nullptr;
	MsgMonitorList events;
	// Keep any fields that are smaller than 4 bytes adjacent to each other.  This conserves memory
	// due to byte-alignment.  It has been verified to save 4 bytes per struct in this case:
	GuiControls type = GUI_CONTROL_INVALID;
	#define GUI_CONTROL_ATTRIB_DPI_RESIZE          0x01
	#define GUI_CONTROL_ATTRIB_ALTSUBMIT           0x02
	// Unused: 0x04
	#define GUI_CONTROL_ATTRIB_EXPLICITLY_HIDDEN   0x08
	#define GUI_CONTROL_ATTRIB_EXPLICITLY_DISABLED 0x10
	#define GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS     0x20
	// Unused: 0x40
	#define GUI_CONTROL_ATTRIB_ALTBEHAVIOR         0x80 // Slider +Invert, ListView/TreeView +WantF2, Edit +WantTab
	UCHAR attrib = 0; // A field of option flags/bits defined above.
	TabControlIndexType tab_control_index = 0; // Which tab control this control belongs to, if any.
	TabIndexType tab_index = 0; // For type==TAB, this stores the tab control's index.  For other types, it stores the page.
	#define CLR_TRANSPARENT 0xFF000001L
	#define IS_AN_ACTUAL_COLOR(color) !((color) & ~0xffffff) // Produces smaller code than checking against CLR_DEFAULT || CLR_INVALID.
	COLORREF background_color = CLR_INVALID;
	HBRUSH background_brush = NULL;
	union
	{
		COLORREF union_color;  // Color of the control's text.
		HBITMAP union_hbitmap = NULL; // For PIC controls, stores the bitmap.
		// Note: Pic controls cannot obey the text color, but they can obey the window's background
		// color if the picture's background is transparent (at least in the case of icons on XP).
		lv_attrib_type *union_lv_attrib; // For ListView: Some attributes and an array of columns.
		IObject *union_object; // For ActiveX.
	};

	GuiControlType() = delete;
	GuiControlType(GuiType* owner) : gui(owner) {}

	static LPTSTR sTypeNames[GUI_CONTROL_TYPE_COUNT];
	static GuiControls ConvertTypeName(LPCTSTR aTypeName);
	LPTSTR GetTypeName();

	// An array of these attributes is used in place of several long switch() statements,
	// to reduce code size and possibly improve performance:
	enum TypeAttribType
	{
		TYPE_SUPPORTS_BGTRANS = 0x01, // Supports +BackgroundTrans.
		TYPE_SUPPORTS_BGCOLOR = 0x02, // Supports +Background followed by a color value.
		TYPE_REQUIRES_BGBRUSH = 0x04, // Requires a brush to be created to implement background color.
		TYPE_MSGBKCOLOR = TYPE_SUPPORTS_BGCOLOR | TYPE_REQUIRES_BGBRUSH, // Supports background color by responding to WM_CTLCOLOR, WM_ERASEBKGND or WM_DRAWITEM.
		TYPE_SETBKCOLOR = TYPE_SUPPORTS_BGCOLOR, // Supports setting a background color by sending it a message.
		TYPE_NO_SUBMIT = 0x08, // Doesn't accept user input, or is excluded from Submit() for some other reason.
		TYPE_HAS_NO_TEXT = 0x10, // Has no text and therefore doesn't use the font or text color.
		TYPE_RESERVE_UNION = 0x20, // Uses the union for some other purpose, so union_color must not be set.
		TYPE_USES_BGCOLOR = 0x40, // Uses Gui.BackColor.
		TYPE_HAS_ITEMS = 0x80, // Add() accepts an Array of items rather than text.
		TYPE_STATICBACK = TYPE_MSGBKCOLOR | TYPE_USES_BGCOLOR, // For brevity in the attrib array.
	};
	typedef UCHAR TypeAttribs;
	static TypeAttribs TypeHasAttrib(GuiControls aType, TypeAttribs aAttrib);
	TypeAttribs TypeHasAttrib(TypeAttribs aAttrib) { return TypeHasAttrib(type, aAttrib); }

	static UCHAR **sRaisesEvents;
	bool SupportsEvent(GuiEventType aEvent);

	bool SupportsBackgroundTrans()
	{
		return TypeHasAttrib(TYPE_SUPPORTS_BGTRANS);
		//switch (type)
		//{
		// Supported via WM_CTLCOLORSTATIC:
		//case GUI_CONTROL_TEXT:
		//case GUI_CONTROL_PIC:
		//case GUI_CONTROL_GROUPBOX:
		//case GUI_CONTROL_BUTTON:
		//	return true;
		//case GUI_CONTROL_CHECKBOX:     Checkbox and radios with trans background have problems with
		//case GUI_CONTROL_RADIO:        their focus rects being drawn incorrectly.
		//case GUI_CONTROL_LISTBOX:      These are also a problem, at least under some theme settings.
		//case GUI_CONTROL_EDIT:
		//case GUI_CONTROL_DROPDOWNLIST:
		//case GUI_CONTROL_SLIDER:       These are a problem under both classic and non-classic themes.
		//case GUI_CONTROL_COMBOBOX:
		//case GUI_CONTROL_LINK:         BackgroundTrans would have no effect.
		//case GUI_CONTROL_LISTVIEW:     Can't reach this point because WM_CTLCOLORxxx is never received for it.
		//case GUI_CONTROL_TREEVIEW:     Same (verified).
		//case GUI_CONTROL_PROGRESS:     Same (verified).
		//case GUI_CONTROL_UPDOWN:       Same (verified).
		//case GUI_CONTROL_DATETIME:     Same (verified).
		//case GUI_CONTROL_MONTHCAL:     Same (verified).
		//case GUI_CONTROL_HOTKEY:       Same (verified).
		//case GUI_CONTROL_TAB:          Same.
		//case GUI_CONTROL_STATUSBAR:    Its text fields (parts) are its children, not ours, so its window proc probably receives WM_CTLCOLORSTATIC, not ours.
		//default:
		//	return false; // Prohibit the TRANS setting for the above control types.
		//}
	}

	bool SupportsBackgroundColor()
	{
		return TypeHasAttrib(TYPE_SUPPORTS_BGCOLOR);
	}

	bool RequiresBackgroundBrush()
	{
		return TypeHasAttrib(TYPE_REQUIRES_BGBRUSH);
	}

	bool HasSubmittableValue()
	{
		return !TypeHasAttrib(TYPE_NO_SUBMIT);
	}

	bool UsesFontAndTextColor()
	{
		return !TypeHasAttrib(TYPE_HAS_NO_TEXT);
	}

	bool UsesUnionColor()
	{
		// It's easier to exclude those which require the union for some other purpose
		// than to whitelist all controls which could potentially cause a WM_CTLCOLOR
		// message (or WM_ERASEBKGND/WM_DRAWITEM in the case of Tab).
		return !TypeHasAttrib(TYPE_RESERVE_UNION);
	}

	bool UsesGuiBgColor()
	{
		return TypeHasAttrib(TYPE_USES_BGCOLOR);
	}

	static ObjectMemberMd sMembers[];
	static ObjectMemberMd sMembersList[]; // Tab, ListBox, ComboBox, DDL
	static ObjectMemberMd sMembersTab[];
	static ObjectMemberMd sMembersDate[];
	static ObjectMemberMd sMembersEdit[];
	static ObjectMemberMd sMembersLV[];
	static ObjectMemberMd sMembersTV[];
	static ObjectMemberMd sMembersSB[];
	static ObjectMemberMd sMembersCB[];

	static Object *sPrototype, *sPrototypeList;
	static Object *sPrototypes[GUI_CONTROL_TYPE_COUNT];
	static void DefineControlClasses();
	static Object *GetPrototype(GuiControls aType);

	FResult Focus();
	FResult GetPos(int *aX, int *aY, int *aWidth, int *aHeight);
	FResult Move(optl<int> aX, optl<int> aY, optl<int> aWidth, optl<int> aHeight);
	FResult OnCommand(int aNotifyCode, ExprTokenType &aCallback, optl<int> aAddRemove);
	FResult OnEvent(StrArg aEventName, ExprTokenType &aCallback, optl<int> aAddRemove);
	FResult OnMessage(UINT aNumber, ExprTokenType &aCallback, optl<int> aAddRemove);
	FResult OnNotify(int aNotifyCode, ExprTokenType &aCallback, optl<int> aAddRemove);
	FResult Opt(StrArg aOptions);
	FResult Redraw();
	FResult SetFont(optl<StrArg> aOptions, optl<StrArg> aFontName);
	
	FResult get_ClassNN(StrRet &aRetVal);
	FResult get_Enabled(BOOL &aRetVal);
	FResult set_Enabled(BOOL aValue);
	FResult get_Focused(BOOL &aRetVal);
	FResult get_Gui(IObject *&aRetVal);
	FResult get_Hwnd(UINT &aRetVal);
	FResult get_Name(StrRet &aRetVal);
	FResult set_Name(StrArg aValue);
	FResult get_Text(ResultToken &aResultToken);
	FResult set_Text(ExprTokenType &aValue);
	FResult get_Type(StrRet &aRetVal);
	FResult get_Value(ResultToken &aResultToken);
	FResult set_Value(ExprTokenType &aValue);
	FResult get_Visible(BOOL &aRetVal);
	FResult set_Visible(BOOL aValue);
	
	FResult DT_SetFormat(optl<StrArg> aFormat);

	FResult Edit_SetCue(StrArg aCueText, optl<BOOL> aActivate);
	
	FResult List_Add(ExprTokenType &aItems);
	FResult List_Choose(ExprTokenType &aValue);
	FResult List_Delete(optl<int> aIndex);
	
	FResult LV_AddInsertModify(optl<int> aRow, optl<StrArg> aOptions, VariantParams &aCol, int *aRetVal, bool aModify);
	FResult LV_Add(optl<StrArg> aOptions, VariantParams &aCol, int &aRetVal)				{ return LV_AddInsertModify(nullptr, aOptions, aCol, &aRetVal, false); }
	FResult LV_Insert(int aRow, optl<StrArg> aOptions, VariantParams &aCol, int &aRetVal)	{ return LV_AddInsertModify(aRow, aOptions, aCol, &aRetVal, false); }
	FResult LV_Modify(int aRow, optl<StrArg> aOptions, VariantParams &aCol)					{ return LV_AddInsertModify(aRow, aOptions, aCol, nullptr, true); }
	
	FResult LV_InsertModifyCol(optl<int> aColumn, optl<StrArg> aOptions, optl<StrArg> aTitle, int *aRetVal, bool aModify);
	FResult LV_InsertCol(optl<int> aColumn, optl<StrArg> aOptions, optl<StrArg> aTitle, int &aRetVal)	{ return LV_InsertModifyCol(aColumn, aOptions, aTitle, &aRetVal, false); }
	FResult LV_ModifyCol(optl<int> aColumn, optl<StrArg> aOptions, optl<StrArg> aTitle)			{ return LV_InsertModifyCol(aColumn, aOptions, aTitle, nullptr, true); }
	void RescaleListViewColumns(int aNumerator, int aDenominator);

	FResult LV_Delete(optl<int> aRow);
	FResult LV_DeleteCol(int aColumn);
	FResult LV_GetCount(optl<StrArg> aMode, int &aRetVal);
	FResult LV_GetNext(optl<int> aStartIndex, optl<StrArg> aRowType, int &aRetVal);
	FResult LV_GetText(int aRow, optl<int> aColumn, StrRet &aRetVal);
	FResult LV_SetImageList(UINT_PTR aImageListID, optl<int> aIconType, UINT_PTR &aRetVal);
	
	FResult SB_SetIcon(StrArg aFilename, optl<int> aIconNumber, optl<UINT> aPartNumber, UINT_PTR &aRetVal);
	FResult SB_SetParts(VariantParams &aParam, UINT& aRetVal);
	FResult SB_SetText(StrArg aNewText, optl<UINT> aPartNumber, optl<UINT> aStyle);
	
	FResult CB_SetCue(StrArg aCueText);

	FResult Tab_UseTab(ExprTokenType *aTab, optl<BOOL> aExact);
	
	FResult TV_AddModify(bool aAdd, UINT_PTR aItemID, UINT_PTR aParentItemID, optl<StrArg> aOptions, optl<StrArg> aName, UINT_PTR &aRetVal);
	FResult TV_Add(StrArg aName, optl<UINT_PTR> aParentItemID, optl<StrArg> aOptions, UINT_PTR &aRetVal) { return TV_AddModify(true, 0, aParentItemID.value_or(0), aOptions, aName, aRetVal); }
	FResult TV_Modify(UINT_PTR aItemID, optl<StrArg> aOptions, optl<StrArg> aNewName, UINT_PTR &aRetVal) { return TV_AddModify(false, aItemID, 0, aOptions, aNewName, aRetVal); }
	FResult TV_Delete(optl<UINT_PTR> aItemID);
	FResult TV_Get(UINT_PTR aItemID, StrArg aAttribute, UINT_PTR &aRetVal);
	FResult TV_GetChild(UINT_PTR aItemID, UINT_PTR &aRetVal);
	FResult TV_GetCount(UINT &aRetVal);
	FResult TV_GetNext(optl<UINT_PTR> aItemID, optl<StrArg> aItemType, UINT_PTR &aRetVal);
	FResult TV_GetParent(UINT_PTR aItemID, UINT_PTR &aRetVal);
	FResult TV_GetPrev(UINT_PTR aItemID, UINT_PTR &aRetVal);
	FResult TV_GetSelection(UINT_PTR &aRetVal);
	FResult TV_GetText(UINT_PTR aItemID, StrRet &aRetVal);
	FResult TV_SetImageList(UINT_PTR aImageListID, optl<int> aIconType, UINT_PTR &aRetVal);

	void Dispose(); // Called by GuiType::Dispose().
};

struct GuiControlOptionsType
{
	DWORD style_add, style_remove, exstyle_add, exstyle_remove, listview_style;
	int listview_view; // Viewing mode, such as LV_VIEW_ICON, LV_VIEW_DETAILS.  Int vs. DWORD to more easily use any negative value as "invalid".
	HIMAGELIST himagelist;
	int x, y, width, height;  // Position info.
	float row_count;
	int choice;  // Which item of a DropDownList/ComboBox/ListBox to initially choose.
	int range_min, range_max;  // Allowable range, such as for a slider control.
	int tick_interval; // The interval at which to draw tickmarks for a slider control.
	int line_size, page_size; // Also for slider.
	int thickness;  // Thickness of slider's thumb.
	int tip_side; // Which side of the control to display the tip on (0 to use default side).
	GuiControlType *buddy1, *buddy2;
	COLORREF color; // Control's text color.
	COLORREF color_bk; // Control's background color.
	int limit;   // The max number of characters to permit in an edit or combobox's edit (also used by ListView).
	int hscroll_pixels;  // The number of pixels for a listbox's horizontal scrollbar to be able to scroll.
	int checked; // When zeroed, struct contains default starting state of checkbox/radio, i.e. BST_UNCHECKED.
	int icon_number; // Which icon of a multi-icon file to use.  Zero means use-default, i.e. the first icon.
	#define GUI_MAX_TABSTOPS 50
	UINT tabstop[GUI_MAX_TABSTOPS]; // Array of tabstops for the interior of a multi-line edit control.
	UINT tabstop_count;  // The number of entries in the above array.
	SYSTEMTIME sys_time[2]; // Needs to support 2 elements for MONTHCAL's multi/range mode.
	SYSTEMTIME sys_time_range[2];
	DWORD gdtr, gdtr_range; // Used in connection with sys_time and sys_time_range.
	ResultType redraw;  // Whether the state of WM_REDRAW should be changed.
	TCHAR password_char; // When zeroed, indicates "use default password" for an edit control with the password style.
	bool range_changed;
	bool tick_interval_changed, tick_interval_specified;
	bool start_new_section;
	bool use_theme; // v1.0.32: Provides the means for the window's current setting of mUseTheme to be overridden.
	bool listview_no_auto_sort; // v1.0.44: More maintainable and frees up GUI_CONTROL_ATTRIB_ALTBEHAVIOR for other uses.
	bool tab_control_uses_dialog;
	#define TAB3_AUTOWIDTH 1
	#define TAB3_AUTOHEIGHT 2
	CHAR tab_control_autosize;
	ATOM customClassAtom;
};

LRESULT CALLBACK GuiWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK TabWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK ControlWindowProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);

class GuiType : public Object
{
	// The use of 72 and 96 below comes from v1, using the font's point size in the
	// calculation.  It's really just 11.25 * font height in pixels.  Could use 12
	// or 11 * font height, but keeping the same defaults as v1 seems worthwhile.
	#define GUI_STANDARD_WIDTH_MULTIPLIER 15 // This times font size = width, if all other means of determining it are exhausted.
	#define GUI_STANDARD_WIDTH GUI_STANDARD_WIDTH_MULTIPLIER * (MulDiv(sFont[mCurrentFontIndex].lfHeight, -72, 96)) // 96 vs. g_ScreenDPI since lfHeight already accounts for DPI.  Don't combine GUI_STANDARD_WIDTH_MULTIPLIER with -72 as it changes the result (due to rounding).
	// Update for v1.0.21: Reduced it to 8 vs. 9 because 8 causes the height each edit (with the
	// default style) to exactly match that of a Combo or DropDownList.  This type of spacing seems
	// to be what other apps use too, and seems to make edits stand out a little nicer:
	#define GUI_CTL_VERTICAL_DEADSPACE DPIScale(8)
	#define PROGRESS_DEFAULT_THICKNESS MulDiv(sFont[mCurrentFontIndex].lfHeight, -2 * 72, 96) // 96 vs. g_ScreenDPI to preserve DPI scale.

public:
	// Ensure fields of the same size are grouped together to avoid padding between larger types
	// and smaller types.  On last check, this saved 8 bytes per GUI on x64 (where pointers are
	// of course 64-bit, so a sequence like HBRUSH, COLORREF, HBRUSH would cause padding).
	// POINTER-SIZED FIELDS:
	GuiType *mNextGui, *mPrevGui; // For global Gui linked list.
	HWND mHwnd;
	LPTSTR mName;
	HWND mStatusBarHwnd;
	HWND mOwner;  // The window that owns this one, if any.  Note that Windows provides no way to change owners after window creation.
	// Control IDs are higher than their index in the array by +CONTROL_ID_FIRST.  This offset is
	// necessary because windows that behave like dialogs automatically return IDOK and IDCANCEL in
	// response to certain types of standard actions:
	GuiControlType **mControl; // Will become an array of controls when the window is first created.
	HBRUSH mBackgroundBrushWin;   // Brush corresponding to mBackgroundColorWin.
	HDROP mHdrop;                 // Used for drag and drop operations.
	HICON mIconEligibleForDestruction; // The window's icon, which can be destroyed when the window is destroyed if nothing else is using it.
	HICON mIconEligibleForDestructionSmall; // L17: A window may have two icons: ICON_SMALL and ICON_BIG.
	HACCEL mAccel; // Keyboard accelerator table.
	UserMenu *mMenu;
	IObject* mEventSink;
	MsgMonitorList mEvents;
	// 32-BIT FIELDS:
	GuiIndexType mControlCount;
	GuiIndexType mControlCapacity; // How many controls can fit into the current memory size of mControl.
	GuiIndexType mDefaultButtonIndex; // Index vs. pointer is needed for some things.
	DWORD mStyle, mExStyle; // Style of window.
	int mCurrentFontIndex;
	COLORREF mCurrentColor;       // The default color of text in controls.
	COLORREF mBackgroundColorWin; // The window's background color itself.
	// Keep the following position/size fields in sequence from mMarginX to mMaxHeight for RescaleForDPI():
	int mMarginX, mMarginY, mPrevX, mPrevY, mPrevWidth, mPrevHeight, mMaxExtentRight, mMaxExtentDown
		, mSectionX, mSectionY, mMaxExtentRightSection, mMaxExtentDownSection;
	LONG mMinWidth, mMinHeight, mMaxWidth, mMaxHeight;
	int mDPI;
	// 8-BIT FIELDS:
	TabControlIndexType mTabControlCount;
	TabControlIndexType mCurrentTabControlIndex; // Which tab control of the window.
	TabIndexType mCurrentTabIndex;// Which tab of a tab control is currently the default for newly added controls.
	bool mInRadioGroup; // Whether the control currently being created is inside a prior radio-group.
	bool mUseTheme;  // Whether XP theme and styles should be applied to the parent window and subsequently added controls.
	bool mGuiShowHasNeverBeenDone, mFirstActivation, mShowIsInProgress, mDestroyWindowHasBeenCalled;
	bool mControlWidthWasSetByContents; // Whether the most recently added control was auto-width'd to fit its contents.
	bool mUsesDPIScaling; // Whether the GUI uses DPI scaling.
	bool mDefaultDPIResize; // Default DPIResize setting to apply to new controls.
	bool mIsMinimized; // Workaround for bad OS behaviour; see "case WM_SETFOCUS".
	bool mDisposed; // Simplifies Dispose().
	bool mVisibleRefCounted; // Whether AddRef() has been done as a result of the window being shown.

	#define MAX_GUI_FONTS 200  // v1.0.44.14: Increased from 100 to 200 due to feedback that 100 wasn't enough.  But to alleviate memory usage, the array is now allocated upon first use.
	static FontType *sFont; // An array of structs, allocated upon first use.
	static int sFontCount;
	static HWND sTreeWithEditInProgress; // Needed because TreeView's edit control for label-editing conflicts with IDOK (default button).

	// Don't overload new and delete operators in this case since we want to use real dynamic memory
	// (since GUIs can be destroyed and recreated, over and over).
	
	FResult Add(StrArg aCtrlType, optl<StrArg> aOptions, ExprTokenType *aContent, IObject *&aRetVal);
	#define ADD_METHOD(NAME, CTRLTYPE) \
		FResult NAME(optl<StrArg> aOptions, ExprTokenType *aContent, IObject *&aRetVal) { \
			return AddControl(aOptions, aContent, aRetVal, CTRLTYPE); \
		}
	ADD_METHOD(AddActiveX,		GUI_CONTROL_ACTIVEX)
	ADD_METHOD(AddButton,		GUI_CONTROL_BUTTON)
	ADD_METHOD(AddCheckBox,		GUI_CONTROL_CHECKBOX)
	ADD_METHOD(AddComboBox,		GUI_CONTROL_COMBOBOX)
	ADD_METHOD(AddCustom,		GUI_CONTROL_CUSTOM)
	ADD_METHOD(AddDateTime,		GUI_CONTROL_DATETIME)
	ADD_METHOD(AddDDL,			GUI_CONTROL_DROPDOWNLIST)
	ADD_METHOD(AddDropDownList,	GUI_CONTROL_DROPDOWNLIST)
	ADD_METHOD(AddEdit,			GUI_CONTROL_EDIT)
	ADD_METHOD(AddGroupBox,		GUI_CONTROL_GROUPBOX)
	ADD_METHOD(AddHotkey,		GUI_CONTROL_HOTKEY)
	ADD_METHOD(AddLink,			GUI_CONTROL_LINK)
	ADD_METHOD(AddListBox,		GUI_CONTROL_LISTBOX)
	ADD_METHOD(AddListView,		GUI_CONTROL_LISTVIEW)
	ADD_METHOD(AddMonthCal,		GUI_CONTROL_MONTHCAL)
	ADD_METHOD(AddPic,			GUI_CONTROL_PIC)
	ADD_METHOD(AddPicture,		GUI_CONTROL_PIC)
	ADD_METHOD(AddProgress,		GUI_CONTROL_PROGRESS)
	ADD_METHOD(AddRadio,		GUI_CONTROL_RADIO)
	ADD_METHOD(AddSlider,		GUI_CONTROL_SLIDER)
	ADD_METHOD(AddStatusBar,	GUI_CONTROL_STATUSBAR)
	ADD_METHOD(AddTab,			GUI_CONTROL_TAB)
	ADD_METHOD(AddTab2,			GUI_CONTROL_TAB2)
	ADD_METHOD(AddTab3,			GUI_CONTROL_TAB3)
	ADD_METHOD(AddText,			GUI_CONTROL_TEXT)
	ADD_METHOD(AddTreeView,		GUI_CONTROL_TREEVIEW)
	ADD_METHOD(AddUpDown,		GUI_CONTROL_UPDOWN)
	#undef ADD_METHOD
	
	FResult __New(optl<StrArg> aOptions, optl<StrArg> aTitle, optl<IObject*> aEventObj);
	FResult Destroy();
	FResult Show(optl<StrArg> aOptions);
	FResult Hide();
	FResult Minimize();
	FResult Maximize();
	FResult Restore();
	FResult Flash(optl<BOOL> aBlink);
	
	FResult GetPos(int *aX, int *aY, int *aWidth, int *aHeight);
	FResult GetClientPos(int *aX, int *aY, int *aWidth, int *aHeight);
	FResult Move(optl<int> aX, optl<int> aY, optl<int> aWidth, optl<int> aHeight);
	
	FResult OnEvent(StrArg aEventName, ExprTokenType &aCallback, optl<int> aAddRemove);
	FResult OnMessage(UINT aNumber, ExprTokenType &aCallback, optl<int> aAddRemove);
	FResult Opt(StrArg aOptions);
	FResult SetFont(optl<StrArg> aOptions, optl<StrArg> aFontName);
	FResult Submit(optl<BOOL> aHideIt, IObject *&aRetVal);
	FResult __Enum(optl<int> aVarCount, IObject *&aRetVal);
	
	FResult get_Hwnd(UINT &aRetVal);
	FResult get_Title(StrRet &aRetVal);
	FResult set_Title(StrArg aValue);
	FResult get_MenuBar(ResultToken &aRetVal);
	FResult set_MenuBar(ExprTokenType *aValue);
	FResult get___Item(ExprTokenType &aIndex, ResultToken &aRetVal);
	FResult get_FocusedCtrl(IObject *&aRetVal);
	FResult get_MarginX(int &aRetVal) { return get_Margin(aRetVal, mMarginX); }
	FResult get_MarginY(int &aRetVal) { return get_Margin(aRetVal, mMarginY); }
	FResult set_MarginX(int aValue) { mMarginX = Scale(aValue); return OK; }
	FResult set_MarginY(int aValue) { mMarginY = Scale(aValue); return OK; }
	FResult get_BackColor(ResultToken &aResultToken);
	FResult set_BackColor(ExprTokenType &aValue);
	FResult set_Name(StrArg aName);
	FResult get_Name(StrRet &aRetVal);

	static ObjectMemberMd sMembers[];
	static int sMemberCount;
	static Object *sPrototype;

	GuiType() // Constructor
		: mHwnd(NULL), mOwner(NULL), mName(NULL)
		, mPrevGui(NULL), mNextGui(NULL)
		, mControl(NULL), mControlCount(0), mControlCapacity(0)
		, mStatusBarHwnd(NULL)
		, mDefaultButtonIndex(-1), mEventSink(NULL)
		, mMenu(NULL)
		// The styles DS_CENTER and DS_3DLOOK appear to be ineffectual in this case.
		// Also note that WS_CLIPSIBLINGS winds up on the window even if unspecified, which is a strong hint
		// that it should always be used for top level windows across all OSes.  Usenet posts confirm this.
		// Also, it seems safer to have WS_POPUP under a vague feeling that it seems to apply to dialog
		// style windows such as this one, and the fact that it also allows the window's caption to be
		// removed, which implies that POPUP windows are more flexible than OVERLAPPED windows.
		, mStyle(WS_POPUP|WS_CLIPSIBLINGS|WS_CAPTION|WS_SYSMENU|WS_MINIMIZEBOX) // WS_CLIPCHILDREN (doesn't seem helpful currently)
		, mExStyle(0) // This and the above should not be used once the window has been created since they might get out of date.
		, mInRadioGroup(false), mUseTheme(true)
		, mCurrentFontIndex(FindOrCreateFont()) // Must call this in constructor to ensure sFont array is never empty while a GUI object exists.  Omit params to tell it to find or create DEFAULT_GUI_FONT.
		, mTabControlCount(0), mCurrentTabControlIndex(MAX_TAB_CONTROLS), mCurrentTabIndex(0)
		, mCurrentColor(CLR_DEFAULT)
		, mBackgroundColorWin(CLR_DEFAULT), mBackgroundBrushWin(NULL)
		, mHdrop(NULL), mIconEligibleForDestruction(NULL), mIconEligibleForDestructionSmall(NULL)
		, mAccel(NULL)
		, mMarginX(COORD_UNSPECIFIED), mMarginY(COORD_UNSPECIFIED) // These will be set when the first control is added.
		, mPrevX(0), mPrevY(0)
		, mPrevWidth(0), mPrevHeight(0) // Needs to be zero for first control to start off at right offset.
		, mMaxExtentRight(0), mMaxExtentDown(0)
		, mSectionX(COORD_UNSPECIFIED), mSectionY(COORD_UNSPECIFIED)
		, mMaxExtentRightSection(COORD_UNSPECIFIED), mMaxExtentDownSection(COORD_UNSPECIFIED)
		, mMinWidth(COORD_UNSPECIFIED), mMinHeight(COORD_UNSPECIFIED)
		, mMaxWidth(COORD_UNSPECIFIED), mMaxHeight(COORD_UNSPECIFIED)
		, mGuiShowHasNeverBeenDone(true), mFirstActivation(true), mShowIsInProgress(false)
		, mDestroyWindowHasBeenCalled(false), mControlWidthWasSetByContents(false)
		, mUsesDPIScaling(true), mDPI(96), mDefaultDPIResize(true)
		, mDisposed(false)
		, mVisibleRefCounted(false)
	{
		SetBase(sPrototype);
	}

	~GuiType()
	{
		// This is done here rather than in Dispose() to allow get_Name() and set_Name()
		// to omit the mHwnd checks, which allows the Gui to be identified after Destroy()
		// is called, and also reduces code size marginally.
		free(mName);
	}
	
	void Dispose();
	static void DestroyIconsIfUnused(HICON ahIcon, HICON ahIconSmall); // L17: Renamed function and added parameter to also handle the window's small icon.
	static GuiType *Create() { return new GuiType(); } // For Object::New<GuiType>().
	FResult AddControl(optl<StrArg> aOptions, ExprTokenType *aContent, IObject *&aRetVal, GuiControls aCtrlType);
	FResult Create(LPCTSTR aTitle);
	FResult OnEvent(GuiControlType *aControl, UINT aEvent, UCHAR aEventKind, ExprTokenType &aCallback, optl<int> aAddRemove);
	FResult OnEvent(GuiControlType *aControl, UINT aEvent, UCHAR aEventKind, IObject *aFunc, LPTSTR aMethodName, int aMaxThreads);
	void ApplyEventStyles(GuiControlType *aControl, UINT aEvent, bool aAdded);
	void ApplySubclassing(GuiControlType *aControl);
	static LPTSTR sEventNames[];
	static LPTSTR ConvertEvent(GuiEventType evt);
	static GuiEventType ConvertEvent(LPCTSTR evt);
	BOOL IsMonitoring(GuiEventType aEvent) { return mEvents.IsMonitoring(aEvent); }

	ResultType GetEnumItem(UINT &aIndex, Var *, Var *, int);

	static IObject* CreateDropArray(HDROP hDrop);
	static void UpdateMenuBars(HMENU aMenu);
	ResultType AddControl(GuiControls aControlType, LPCTSTR aOptions, LPCTSTR aText, GuiControlType*& apControl, Array *aObj = NULL);
	void MethodGetPos(int *aX, int *aY, int *aWidth, int *aHeight, RECT &aRect, HWND aOrigin);

	ResultType ParseOptions(LPCTSTR aOptions, bool &aSetLastFoundWindow, ToggleValueType &aOwnDialogs);
	void SetOwnDialogs(ToggleValueType state)
	{
		if (state == TOGGLE_INVALID)
			return;
		g->DialogOwner = state == TOGGLED_ON ? mHwnd : NULL;
	}
	void GetNonClientArea(LONG &aWidth, LONG &aHeight);
	void GetTotalWidthAndHeight(LONG &aWidth, LONG &aHeight);
	void ParseMinMaxSizeOption(LPCTSTR aOptionValue, LONG &aWidth, LONG &aHeight);

	ResultType ControlParseOptions(LPCTSTR aOptions, GuiControlOptionsType &aOpt, GuiControlType &aControl
		, GuiIndexType aControlIndex = -1); // aControlIndex is not needed upon control creation.
	void ControlInitOptions(GuiControlOptionsType &aOpt, GuiControlType &aControl);
	void ControlAddItems(GuiControlType &aControl, Array *aObj);
	void ControlSetChoice(GuiControlType &aControl, int aChoice);
	ResultType ControlLoadPicture(GuiControlType &aControl, LPCTSTR aFilename, int aWidth, int aHeight, int aIconNumber);
	void Cancel();
	void Close(); // Due to SC_CLOSE, etc.
	void Escape(); // Similar to close, except typically called when the user presses ESCAPE.
	void VisibilityChanged();

	static GuiType *FindGui(HWND aHwnd, bool aVerifyProcess = false);
	static GuiType *FindGuiParent(HWND aHwnd, bool aVerifyProcess = false);

	GuiIndexType FindControl(LPCTSTR aControlID);
	GuiIndexType FindControlIndex(HWND aHwnd)
	{
		for (;;)
		{
			GuiIndexType index = GUI_HWND_TO_INDEX(aHwnd); // Retrieves a small negative on failure, which will be out of bounds when converted to unsigned.
			if (index < mControlCount && mControl[index]->hwnd == aHwnd) // A match was found.  Fix for v1.1.09.03: Confirm it is actually one of our controls.
				return index;
			// Not found yet; try again with parent.  ComboBox and ListView only need one iteration,
			// but multiple iterations may be needed by ActiveX/Custom or other future control types.
			// Note that a ComboBox's drop-list (class ComboLBox) is apparently a direct child of the
			// desktop, so this won't help us in that case.
			aHwnd = GetParent(aHwnd);
			if (!aHwnd || aHwnd == mHwnd) // No match, so indicate failure.
				return NO_CONTROL_INDEX;
		}
	}
	GuiControlType *FindControl(HWND aHwnd)
	{
		GuiIndexType index = FindControlIndex(aHwnd);
		return index == NO_CONTROL_INDEX ? NULL : mControl[index];
	}

	int FindGroup(GuiIndexType aControlIndex, GuiIndexType &aGroupStart, GuiIndexType &aGroupEnd);

	int FindOrCreateFont(LPCTSTR aOptions = _T(""), LPCTSTR aFontName = _T(""), FontType *aFoundationFont = NULL
		, COLORREF *aColor = NULL);
	static int FindFont(FontType &aFont);
	static void FontGetAttributes(FontType &aFont);

	void Event(GuiIndexType aControlIndex, UINT aNotifyCode, USHORT aGuiEvent = GUI_EVENT_NONE, UINT_PTR aEventInfo = 0);
	bool ControlWmNotify(GuiControlType &aControl, LPNMHDR aNmHdr, INT_PTR &aRetVal);
	bool MsgMonitor(GuiControlType *aControl, UINT aMsg, WPARAM awParam, LPARAM alParam, MSG *apMsg, INT_PTR *aRetVal);

	static WORD TextToHotkey(LPCTSTR aText);
	static LPTSTR HotkeyToText(WORD aHotkey, LPTSTR aBuf);
	FResult ControlSetName(GuiControlType &aControl, LPCTSTR aName);
	void ControlSetEnabled(GuiControlType &aControl, bool aEnabled);
	void ControlSetVisible(GuiControlType &aControl, bool aVisible);
	FResult ControlMove(GuiControlType &aControl, int aX, int aY, int aWidth, int aHeight);
	ResultType ControlSetFont(GuiControlType &aControl, LPCTSTR aOptions, LPCTSTR aFontName);
	void ControlSetTextColor(GuiControlType &aControl, COLORREF aColor);
	void ControlSetMonthCalColor(GuiControlType &aControl, COLORREF aColor, UINT aMsg);
	ResultType ControlChoose(GuiControlType &aControl, ExprTokenType &aParam, BOOL aOneExact = FALSE);
	void ControlCheckRadioButton(GuiControlType &aControl, GuiIndexType aControlIndex, WPARAM aCheckType);
	void ControlSetUpDownOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	int ControlGetDefaultSliderThickness(DWORD aStyle, int aThumbThickness);
	void ControlSetSliderOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	int ControlInvertSliderIfNeeded(GuiControlType &aControl, int aPosition);
	void ControlSetListViewOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	void ControlSetTreeViewOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt);
	void ControlSetProgressOptions(GuiControlType &aControl, GuiControlOptionsType &aOpt, DWORD aStyle);
	GuiControlType *ControlOverrideBkColor(GuiControlType &aControl);
	void ControlGetBkColor(GuiControlType &aControl, bool aUseWindowColor, HBRUSH &aBrush, COLORREF &aColor);
	
	FResult ControlSetContents(GuiControlType &aControl, ExprTokenType &aContents, bool aIsText);
	FResult ControlSetPic(GuiControlType &aControl, LPTSTR aContents);
	FResult ControlSetChoice(GuiControlType &aControl, LPTSTR aContents, bool aIsText); // DDL, ComboBox, ListBox, Tab
	FResult ControlSetEdit(GuiControlType &aControl, LPTSTR aContents, bool aIsText);
	FResult ControlSetDateTime(GuiControlType &aControl, LPTSTR aContents);
	FResult ControlSetMonthCal(GuiControlType &aControl, LPTSTR aContents);
	FResult ControlSetHotkey(GuiControlType &aControl, LPTSTR aContents);
	FResult ControlSetCheck(GuiControlType &aControl, int aValue); // CheckBox, Radio
	FResult ControlSetUpDown(GuiControlType &aControl, int aValue);
	FResult ControlSetSlider(GuiControlType &aControl, int aValue);
	FResult ControlSetProgress(GuiControlType &aControl, int aValue);

	enum ValueModeType { Value_Mode, Text_Mode, Submit_Mode };

	void ControlGetContents(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode = Value_Mode);
	void ControlGetCheck(ResultToken &aResultToken, GuiControlType &aControl); // CheckBox, Radio
	void ControlGetDDL(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode);
	void ControlGetComboBox(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode);
	void ControlGetListBox(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode);
	void ControlGetEdit(ResultToken &aResultToken, GuiControlType &aControl);
	void ControlGetDateTime(ResultToken &aResultToken, GuiControlType &aControl);
	void ControlGetMonthCal(ResultToken &aResultToken, GuiControlType &aControl);
	void ControlGetHotkey(ResultToken &aResultToken, GuiControlType &aControl);
	void ControlGetUpDown(ResultToken &aResultToken, GuiControlType &aControl);
	void ControlGetSlider(ResultToken &aResultToken, GuiControlType &aControl);
	void ControlGetProgress(ResultToken &aResultToken, GuiControlType &aControl);
	void ControlGetTab(ResultToken &aResultToken, GuiControlType &aControl, ValueModeType aMode);
	
	ResultType ControlGetWindowText(ResultToken &aResultToken, GuiControlType &aControl);
	void ControlRedraw(GuiControlType &aControl, bool aOnlyWithinTab = false);

	void ControlUpdateCurrentTab(GuiControlType &aTabControl, bool aFocusFirstControl);
	GuiControlType *FindTabControl(TabControlIndexType aTabControlIndex);
	int FindTabIndexByName(GuiControlType &aTabControl, LPTSTR aName, bool aExactMatch = false);
	int GetControlCountOnTabPage(TabControlIndexType aTabControlIndex, TabIndexType aTabIndex);
	void GetTabDisplayAreaRect(HWND aTabControlHwnd, RECT &aRect);
	POINT GetPositionOfTabDisplayArea(GuiControlType &aTabControl);
	ResultType SelectAdjacentTab(GuiControlType &aTabControl, bool aMoveToRight, bool aFocusFirstControl
		, bool aWrapAround);
	void AutoSizeTabControl(GuiControlType &aTabControl);
	ResultType CreateTabDialog(GuiControlType &aTabControl, GuiControlOptionsType &aOpt);
	void UpdateTabDialog(HWND aTabControlHwnd);
	void ControlGetPosOfFocusedItem(GuiControlType &aControl, POINT &aPoint);
	static void LV_Sort(GuiControlType &aControl, int aColumnIndex, bool aSortOnlyIfEnabled, TCHAR aForceDirection = '\0');
	static IObject *ControlGetActiveX(HWND aWnd);
	
	void UpdateAccelerators(UserMenu &aMenu);
	void UpdateAccelerators(UserMenu &aMenu, LPACCEL aAccel, int &aAccelCount);
	void RemoveAccelerators();
	static bool ConvertAccelerator(LPTSTR aString, ACCEL &aAccel);

	void SetDefaultMargins();
	FResult get_Margin(int &aRetVal, int &aMargin);

	// See DPIScale() and DPIUnscale() for more details.
	int Scale(int x) { return mUsesDPIScaling ? MulDiv(x, mDPI, 96) : x; }
	int Unscale(int x) { return mUsesDPIScaling ? MulDiv(x, 96, mDPI) : x; }
	// The following is a workaround for the "w-1" and "h-1" options:
	int ScaleSize(int x) { return mUsesDPIScaling && x != -1 ? DPIScale(x) : x; }

	void RescaleForDPI(int aDPI, RECT &aRect);
	static int RescaleFontForDPI(HFONT aFont, int aOldDPI, int aNewDPI);

	int GetSystemMetrics(int nIndex);

protected:
	bool Delete() override;
};
