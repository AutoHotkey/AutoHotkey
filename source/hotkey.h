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

#ifndef hotkey_h
#define hotkey_h

#include "keyboard_mouse.h"
#include "script.h"  // For which label (and in turn which line) in the script to jump to.
EXTERN_SCRIPT;  // For g_script.

// Due to control/alt/shift modifiers, quite a lot of hotkey combinations are possible, so support any
// conceivable use.  This value is used for initial sizing of the shk array, which doubles in size each
// time its capacity is reached (making the need for expansion rare).
#define INITIAL_MAX_HOTKEYS 256

// Note: 0xBFFF is the largest ID that can be used with RegisterHotkey().
// But further limit this to 0x7FFF (32,767) so that the highest order bit
// is reserved for our other use:
#define HOTKEY_KEY_UP                  0x8000
#define HOTKEY_ID_MASK                 0x7FFF
#define HOTKEY_ID_INVALID              HOTKEY_ID_MASK
#define HOTKEY_ID_ALT_TAB              0x7FFE
#define HOTKEY_ID_ALT_TAB_SHIFT        0x7FFD
#define HOTKEY_ID_ALT_TAB_MENU         0x7FFC
#define HOTKEY_ID_ALT_TAB_AND_MENU     0x7FFB
#define HOTKEY_ID_ALT_TAB_MENU_DISMISS 0x7FFA
#define HOTKEY_ID_MAX                  0x7FF9  // 32762 hotkeys
#define HOTKEY_ID_ON                   0x01  // This and the next 2 are used only for convenience by ConvertAltTab().
#define HOTKEY_ID_OFF                  0x02
#define HOTKEY_ID_TOGGLE               0x03
#define IS_ALT_TAB(id) (id > HOTKEY_ID_MAX && id < HOTKEY_ID_INVALID)

#define COMPOSITE_DELIMITER _T(" & ")
#define COMPOSITE_DELIMITER_LENGTH 3

// Smallest workable size: to save mem in some large arrays that use this:
typedef USHORT HotkeyIDType; // This is relied upon to be unsigned; e.g. many places omit a check for ID < 0.
typedef HotkeyIDType HookActionType;

typedef UCHAR HotkeyTypeType;
enum HotkeyTypeEnum {HK_NORMAL, HK_KEYBD_HOOK, HK_MOUSE_HOOK, HK_BOTH_HOOKS, HK_JOYSTICK};
// If above numbers are ever changed/reshuffled, update the macros below.
#define HK_TYPE_CAN_BECOME_KEYBD_HOOK(type) (type == HK_NORMAL)
#define HK_TYPE_IS_HOOK(type) (type > HK_NORMAL && type < HK_JOYSTICK)

HWND HotCriterionAllowsFiring(HotkeyCriterion *aCriterion, LPTSTR aHotkeyName); // Used by hotkeys and hotstrings.
bool HotInputLevelAllowsFiring(SendLevelType inputLevel, ULONG_PTR aEventExtraInfo, LPTSTR aKeyHistoryChar);
ResultType SetHotkeyCriterion(HotCriterionType aType, LPTSTR aWinTitle, LPTSTR aWinText);
HotkeyCriterion *AddHotkeyIfExpr();
HotkeyCriterion *FindHotkeyIfExpr(LPTSTR aExpr);



struct HotkeyVariant
{
	LabelRef mJumpToLabel;
	HotkeyCriterion *mHotCriterion;
	HotkeyVariant *mNextVariant;
	DWORD mRunAgainTime;
	int mPriority;
	// Keep members that are less than 32-bit adjacent to each other to conserve memory in with the default
	// 4-byte alignment:
	USHORT mIndex;
	UCHAR mExistingThreads, mMaxThreads;
	SendLevelType mInputLevel;
	bool mNoSuppress; // v1.0.44: This became a per-variant attribute because it's more useful/flexible that way.
	bool mMaxThreadsBuffer;
	bool mRunAgainAfterFinished;
	bool mEnabled; // Whether this variant has been disabled via the Hotkey command.
};



class Hotkey
{
private:
	// These are done as static, rather than having an outer-class to contain all the hotkeys, because
	// the hotkey ID is used as the array index for performance reasons.  Having an outer class implies
	// the potential future use of more than one set of hotkeys, which could still be implemented
	// within static data and methods to retain the indexing/performance method:
	static HookType sWhichHookNeeded;
	static HookType sWhichHookAlways;
	static DWORD sTimePrev;
	static DWORD sTimeNow;
	static HotkeyIDType sNextID;

	bool Enable(HotkeyVariant &aVariant) // Returns true if the variant needed to be disabled, in which case caller should generally call ManifestAllHotkeysHotstringsHooks().
	{
		if (aVariant.mEnabled) // Added for v1.0.23 to greatly improve performance when hotkey is already in the right state.
			return false; // Indicate that it's already enabled.
		aVariant.mEnabled = true;
		return true;
	}

	bool Disable(HotkeyVariant &aVariant) // Returns true if the variant needed to be disabled, in which case caller should generally call ManifestAllHotkeysHotstringsHooks().
	{
		if (!aVariant.mEnabled) // Added for v1.0.23 to greatly improve performance when hotkey is already in the right state.
			return false; // Indicate that it's already disabled.
		aVariant.mEnabled = false;
		aVariant.mRunAgainAfterFinished = false; // ManifestAllHotkeysHotstringsHooks() won't do this unless the entire hotkey is disabled/unregistered.
		return true;
	}

	bool EnableParent() // Returns true if the hotkey needed to be disabled, in which case caller should generally call ManifestAllHotkeysHotstringsHooks().
	{
		if (mParentEnabled)
			return false; // Indicate that it's already enabled.
		mParentEnabled = true;
		return true;
	}

	bool DisableParent() // Returns true if the hotkey needed to be disabled, in which case caller should generally call ManifestAllHotkeysHotstringsHooks().
	{
		if (!mParentEnabled)
			return false; // Indicate that it's already disabled.
		mParentEnabled = false;
		return true;
	}

	ResultType Register();
	ResultType Unregister();

	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {SimpleHeap::Delete(aPtr);}  // Deletes aPtr if it was the most recently allocated.
	void operator delete[](void *aPtr) {SimpleHeap::Delete(aPtr);}

	// For now, constructor & destructor are private so that only static methods can create new
	// objects.  This allow proper tracking of which OS hotkey IDs have been used.
	Hotkey(HotkeyIDType aID, IObject *aJumpToLabel, HookActionType aHookAction, LPTSTR aName, bool aSuffixHasTilde, bool aUseErrorLevel);
	~Hotkey() {if (mIsRegistered) Unregister();}

public:
	static Hotkey **shk;
	static int shkMax;

	// 32-bit members:
	mod_type mModifiers;  // MOD_ALT, MOD_CONTROL, MOD_SHIFT, MOD_WIN, or some additive or bitwise-or combination of these.

	// 16-bit members:
	HotkeyIDType mID;  // Must be unique for each hotkey of a given thread.
	HookActionType mHookAction;
	sc_type mSC; // Scan code.  All vk's have a scan code, but not vice versa.
	sc_type mModifierSC; // If mModifierVK is zero, this scan code, if non-zero, will be used as the modifier.
	HotkeyIDType mNextCustomCombo; // ID of the next custom combo with the same suffix as this one (initialized by the hook, but only for combos).

	// Keep single-byte attributes adjacent to each other to conserve memory within byte-aligned class/struct:
	modLR_type mModifiersLR;  // Left-right centric versions of the above.
	modLR_type mModifiersConsolidatedLR; // The combination of mModifierVK, mModifierSC, mModifiersLR, modifiers
	vk_type mVK; // virtual-key code, e.g. VK_TAB, VK_LWIN, VK_LMENU, VK_APPS, VK_F10.  If zero, use sc below.
	vk_type mModifierVK; // Any other virtual key that must be pressed down in order to activate "vk" itself.
	HotkeyTypeType mType;
	#define NO_SUPPRESS_PREFIX 0x01               // Bitwise: Bit #1
	#define AT_LEAST_ONE_VARIANT_HAS_TILDE 0x02   // Bitwise: Bit #2
	#define AT_LEAST_ONE_VARIANT_LACKS_TILDE 0x04 // Bitwise: Bit #3
	#define NO_SUPPRESS_NEXT_UP_EVENT 0x08        // Bitwise: Bit #4
	#define NO_SUPPRESS_SUFFIX_VARIES (AT_LEAST_ONE_VARIANT_HAS_TILDE | AT_LEAST_ONE_VARIANT_LACKS_TILDE) // i.e. a hotkey that has variants of both types.
	#define NO_SUPPRESS_STATES NO_SUPPRESS_NEXT_UP_EVENT  // This is a bitwise union (currently only one item) of those of the above that represent a the key's dynamically changing state as the user types.
	UCHAR mNoSuppress;  // Uses the flags above.  Normally 0, but can be overridden by using the hotkey tilde (~) prefix).
	bool mKeybdHookMandatory;
	bool mAllowExtraModifiers;  // False if the hotkey should not fire when extraneous modifiers are held down.
	bool mKeyUp; // A hotkey that should fire on key-up.
	bool mVK_WasSpecifiedByNumber; // A hotkey defined by something like "VK24::" or "Hotkey, VK24, ..."
#ifdef CONFIG_WIN9X
	bool mUnregisterDuringThread; // Win9x: Whether this hotkey should be unregistered during its own subroutine (to prevent its own Send command from firing itself).  Seems okay to apply this to all variants.
#endif
	bool mIsRegistered;  // Whether this hotkey has been successfully registered.
	bool mParentEnabled; // When true, the individual variants' mEnabled flags matter. When false, the entire hotkey is disabled.
	bool mConstructedOK;
	
	// 64- or 32-bit members:
	LPTSTR mName; // Points to the label name for static hotkeys, or a dynamically-allocated string for dynamic hotkeys.
	HotkeyVariant *mFirstVariant, *mLastVariant; // v1.0.42: Linked list of variant hotkeys created via #IfWin directives.

	// Make sHotkeyCount an alias for sNextID.  Make it const to enforce modifying the value in only one way:
	static const HotkeyIDType &sHotkeyCount;
	static bool sJoystickHasHotkeys[MAX_JOYSTICKS];
	static DWORD sJoyHotkeyCount;

	static void AllDestructAndExit(int exit_code);

	#define HOTKEY_EL_BADLABEL           _T("1") // Set as strings so that SetFormat doesn't affect their appearance (for use with "If ErrorLevel in 5,6").
	#define HOTKEY_EL_INVALID_KEYNAME    _T("2")
	#define HOTKEY_EL_UNSUPPORTED_PREFIX _T("3")
	#define HOTKEY_EL_ALTTAB             _T("4")
	#define HOTKEY_EL_NOTEXIST           _T("5")
	#define HOTKEY_EL_NOTEXISTVARIANT    _T("6")
	#define HOTKEY_EL_WIN9X              _T("50")
	#define HOTKEY_EL_NOREG              _T("51")
	#define HOTKEY_EL_MAXCOUNT           _T("98") // 98 allows room for other ErrorLevels to be added in between.
	#define HOTKEY_EL_MEM                _T("99")
	static ResultType Dynamic(LPTSTR aHotkeyName, LPTSTR aLabelName, LPTSTR aOptions, IObject *aJumpToLabel, Var *aJumpToLabelVar);

	static Hotkey *AddHotkey(IObject *aJumpToLabel, HookActionType aHookAction, LPTSTR aName, bool aSuffixHasTilde, bool aUseErrorLevel);
	HotkeyVariant *FindVariant();
	HotkeyVariant *AddVariant(IObject *aJumpToLabel, bool aSuffixHasTilde);
	static bool PrefixHasNoEnabledSuffixes(int aVKorSC, bool aIsSC);
	HotkeyVariant *CriterionAllowsFiring(HWND *aFoundHWND = NULL);
	static HotkeyVariant *CriterionAllowsFiring(HotkeyIDType aHotkeyID, HWND &aFoundHWND);
	static HotkeyVariant *CriterionFiringIsCertain(HotkeyIDType &aHotkeyIDwithFlags, bool aKeyUp, ULONG_PTR aExtraInfo
		, UCHAR &aNoSuppress, bool &aFireWithNoSuppress, LPTSTR aSingleChar);
	static modLR_type HotkeyRequiresModLR(HotkeyIDType aHotkeyIDwithoutflags, modLR_type aModLR);
	static void TriggerJoyHotkeys(int aJoystickID, DWORD aButtonsNewlyDown);
	void PerformInNewThreadMadeByCaller(HotkeyVariant &aVariant);
	static void ManifestAllHotkeysHotstringsHooks();
	static void RequireHook(HookType aWhichHook) {sWhichHookAlways |= aWhichHook;}
	static ResultType TextInterpret(LPTSTR aName, Hotkey *aThisHotkey, bool aUseErrorLevel);

	struct HotkeyProperties // Struct used by TextToModifiers() and its callers.
	{
		mod_type modifiers;
		modLR_type modifiersLR;
		TCHAR prefix_text[32];  // Has to be large enough to hold the largest key name in g_key_to_vk,
		TCHAR suffix_text[32];  // which is probably "Browser_Favorites" (17).
		bool suffix_has_tilde; // As opposed to "prefix has tilde".
		bool has_asterisk;
		bool is_key_up;
		bool hook_is_mandatory;
	};
	static LPTSTR TextToModifiers(LPTSTR aText, Hotkey *aThisHotkey, HotkeyProperties *aProperties = NULL);
	static ResultType TextToKey(LPTSTR aText, LPTSTR aHotkeyName, bool aIsModifier, Hotkey *aThisHotkey, bool aUseErrorLevel);

	static void InstallKeybdHook();
	static void InstallMouseHook();

	bool PerformIsAllowed(HotkeyVariant &aVariant)
	{
		// For now, attempts to launch another simultaneous instance of this subroutine
		// are ignored if MaxThreadsPerHotkey (for this particular hotkey) has been reached.
		// In the future, it might be better to have this user-configurable, i.e. to devise
		// some way for the hotkeys to be kept queued up so that they take effect only when
		// the number of currently active threads drops below the max.  But doing such
		// might make "infinite key loops" harder to catch because the rate of incoming hotkeys
		// would be slowed down to prevent the subroutines from running concurrently:
		ActionTypeType act = aVariant.mJumpToLabel->TypeOfFirstLine();
		return aVariant.mExistingThreads < aVariant.mMaxThreads || (ACT_IS_ALWAYS_ALLOWED(act)); // See below.
		// Although our caller may have already called ACT_IS_ALWAYS_ALLOWED(), it was for a different reason.
	}

	bool IsExemptFromSuspend() // A hotkey is considered exempt if even one of its variants is exempt.
	{
		// It's the caller's responsibility to check vp->mEnabled; that isn't done here.
		if (mHookAction) // An alt-tab hotkey (which overrides all its variants) is never exempt.
			return false;
		for (HotkeyVariant *vp = mFirstVariant; vp; vp = vp->mNextVariant)
			if (vp->mJumpToLabel->IsExemptFromSuspend()) // If it has no label, it's never exempt.
				return true; // Even a single exempt variant makes the entire hotkey exempt.
		// Otherwise, since the above didn't find any exempt variants, the entire hotkey is non-exempt:
		return false;
	}

	bool IsCompletelyDisabled()
	{
		if (mHookAction) // Alt tab hotkeys are disabled completely if and only if the parent is disabled.
			return !mParentEnabled;
		for (HotkeyVariant *vp = mFirstVariant; vp; vp = vp->mNextVariant)
			if (vp->mEnabled)
				return false;
		return true;
	}

	static void RunAgainAfterFinished(HotkeyVariant &aVariant)
	{
		if (aVariant.mMaxThreadsBuffer)
			aVariant.mRunAgainAfterFinished = true;
		aVariant.mRunAgainTime = GetTickCount();
		// Above: The time this event was buffered, to make sure it doesn't get too old.
	}

	static void ResetRunAgainAfterFinished()  // For all hotkeys and all variants of each.
	{
		HotkeyVariant *vp;
		for (int i = 0; i < sHotkeyCount; ++i)
			for (vp = shk[i]->mFirstVariant; vp; vp = vp->mNextVariant)
				vp->mRunAgainAfterFinished = false;
	}

	static HookActionType ConvertAltTab(LPTSTR aBuf, bool aAllowOnOff)
	{
		if (!aBuf || !*aBuf) return 0;
		if (!_tcsicmp(aBuf, _T("AltTab"))) return HOTKEY_ID_ALT_TAB;
		if (!_tcsicmp(aBuf, _T("ShiftAltTab"))) return HOTKEY_ID_ALT_TAB_SHIFT;
		if (!_tcsicmp(aBuf, _T("AltTabMenu"))) return HOTKEY_ID_ALT_TAB_MENU;
		if (!_tcsicmp(aBuf, _T("AltTabAndMenu"))) return HOTKEY_ID_ALT_TAB_AND_MENU;
		if (!_tcsicmp(aBuf, _T("AltTabMenuDismiss"))) return HOTKEY_ID_ALT_TAB_MENU_DISMISS;
		if (aAllowOnOff)
		{
			if (!_tcsicmp(aBuf, _T("On"))) return HOTKEY_ID_ON;
			if (!_tcsicmp(aBuf, _T("Off"))) return HOTKEY_ID_OFF;
			if (!_tcsicmp(aBuf, _T("Toggle"))) return HOTKEY_ID_TOGGLE;
		}
		return 0;
	}

	static Hotkey *FindHotkeyByTrueNature(LPTSTR aName, bool &aSuffixHasTilde, bool &aHookIsMandatory);
	static Hotkey *FindHotkeyContainingModLR(modLR_type aModifiersLR);  //, HotkeyIDType hotkey_id_to_omit);
	//static Hotkey *FindHotkeyWithThisModifier(vk_type aVK, sc_type aSC);
	//static Hotkey *FindHotkeyBySC(sc2_type aSC2, mod_type aModifiers, modLR_type aModifiersLR);

	static LPTSTR ListHotkeys(LPTSTR aBuf, int aBufSize);
	LPTSTR ToText(LPTSTR aBuf, int aBufSize, bool aAppendNewline);

private:
	static HotkeyVariant *CriterionFiringIsCertainHelper(HotkeyIDType &aHotkeyIDwithFlags, bool aKeyUp, UCHAR &aNoSuppress
		, bool &aFireWithNoSuppress, LPTSTR aSingleChar);
};


///////////////////////////////////////////////////////////////////////////////////

#define MAX_HOTSTRING_LENGTH 40  // Hard to imagine a need for more than this, and most are only a few chars long.
#define MAX_HOTSTRING_LENGTH_STR _T("40")  // Keep in sync with the above.
#define HOTSTRING_BLOCK_SIZE 1024
typedef UINT HotstringIDType;

enum CaseConformModes {CASE_CONFORM_NONE, CASE_CONFORM_ALL_CAPS, CASE_CONFORM_FIRST_CAP};


class Hotstring
{
public:
	static Hotstring **shs;  // An array to be allocated on first use (performs better than linked list).
	static HotstringIDType sHotstringCount;
	static HotstringIDType sHotstringCountMax;
	static UINT sEnabledCount; // v1.1.28.00: For performance, such as avoiding calling ToAsciiEx() in the hook.

	LabelRef mJumpToLabel;
	LPTSTR mName;
	LPTSTR mString, mReplacement;
	HotkeyCriterion *mHotCriterion;
	int mPriority, mKeyDelay;

	// Keep members that are smaller than 32-bit adjacent with each other to conserve memory (due to 4-byte alignment).
	SendModes mSendMode;
	SendRawType mSendRaw;
	SendLevelType mInputLevel;
	UCHAR mStringLength;
	UCHAR mSuspended; // FALSE or a combination of one of the following:
	#define HS_SUSPENDED 0x01
	#define HS_TURNED_OFF 0x02
	#define HS_TEMPORARILY_DISABLED 0x04
	UCHAR mExistingThreads, mMaxThreads;
	bool mCaseSensitive, mConformToCase, mDoBackspace, mOmitEndChar, mEndCharRequired
		, mDetectWhenInsideWord, mDoReset, mConstructedOK;

	static void SuspendAll(bool aSuspend);
	ResultType PerformInNewThreadMadeByCaller();
	void DoReplace(LPARAM alParam);
	static Hotstring *FindHotstring(LPTSTR aHotstring, bool aCaseSensitive, bool aDetectWhenInsideWord, HotkeyCriterion *aHotCriterion);
	static ResultType AddHotstring(LPTSTR aName, LabelPtr aJumpToLabel, LPTSTR aOptions, LPTSTR aHotstring
		, LPTSTR aReplacement, bool aHasContinuationSection, UCHAR aSuspend = FALSE);
	static void ParseOptions(LPTSTR aOptions, int &aPriority, int &aKeyDelay, SendModes &aSendMode
		, bool &aCaseSensitive, bool &aConformToCase, bool &aDoBackspace, bool &aOmitEndChar, SendRawType &aSendRaw
		, bool &aEndCharRequired, bool &aDetectWhenInsideWord, bool &aDoReset, bool &aExecuteAction);
	void ParseOptions(LPTSTR aOptions);

	// Constructor & destructor:
	Hotstring(LPTSTR aName, LabelPtr aJumpToLabel, LPTSTR aOptions, LPTSTR aHotstring, LPTSTR aReplacement
		, bool aHasContinuationSection, UCHAR aSuspend);
	~Hotstring() {}  // Note that mReplacement is sometimes malloc'd, sometimes from SimpleHeap, and sometimes the empty string.

	void *operator new(size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void *operator new[](size_t aBytes) {return SimpleHeap::Malloc(aBytes);}
	void operator delete(void *aPtr) {SimpleHeap::Delete(aPtr);}  // Deletes aPtr if it was the most recently allocated.
	void operator delete[](void *aPtr) {SimpleHeap::Delete(aPtr);}
};


#endif
