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
#include "hotkey.h"
#include "globaldata.h"  // For g_os and other global vars.
#include "window.h" // For MsgBox()
//#include "application.h" // For ExitApp()
#include "script_func_impl.h"

// Initialize static members:
HookType Hotkey::sWhichHookNeeded = 0;
HookType Hotkey::sWhichHookAlways = 0;
DWORD Hotkey::sTimePrev = {0};
DWORD Hotkey::sTimeNow = {0};
Hotkey **Hotkey::shk = NULL;
int Hotkey::shkMax = 0;
HotkeyIDType Hotkey::sNextID = 0;
const HotkeyIDType &Hotkey::sHotkeyCount = Hotkey::sNextID;
bool Hotkey::sJoystickHasHotkeys[MAX_JOYSTICKS] = {false};
DWORD Hotkey::sJoyHotkeyCount = 0;



HWND HotCriterionAllowsFiring(HotkeyCriterion *aCriterion, LPTSTR aHotkeyName)
// This is a global function because it's used by both hotkeys and hotstrings.
// In addition to being called by the hook thread, this can now be called by the main thread.
// That happens when a WM_HOTKEY message arrives (for non-hook hotkeys, i.e. RegisterHotkey).
// Returns a non-NULL HWND if firing is allowed.  However, if it's a global criterion or
// a "not-criterion" such as #HotIf Not WinActive(), (HWND)1 is returned rather than a genuine HWND.
{
	HWND found_hwnd;
	if (!aCriterion)
		return (HWND)1; // Always allow hotkey to fire.
	switch(aCriterion->Type)
	{
	case HOT_IF_ACTIVE:
	case HOT_IF_NOT_ACTIVE:
		found_hwnd = WinActive(g_default, aCriterion->WinTitle, aCriterion->WinText, _T(""), _T(""), false); // Thread-safe.
		break;
	case HOT_IF_EXIST:
	case HOT_IF_NOT_EXIST:
		found_hwnd = WinExist(g_default, aCriterion->WinTitle, aCriterion->WinText, _T(""), _T(""), false, false); // Thread-safe.
		break;
	// L4: Handling of #HotIf (expression) hotkey variants.
	case HOT_IF_CALLBACK:
		// Expression evaluation must be done in the main thread. If the message times out, the hotkey/hotstring is not allowed to fire.
		DWORD_PTR res;
		return (SendMessageTimeout(g_hWnd, AHK_HOT_IF_EVAL, (WPARAM)aCriterion, (LPARAM)aHotkeyName, SMTO_BLOCK | SMTO_ABORTIFHUNG, g_HotExprTimeout, &res) && res == CONDITION_TRUE) ? (HWND)1 : NULL;
	}
	return (aCriterion->Type == HOT_IF_ACTIVE || aCriterion->Type == HOT_IF_EXIST) ? found_hwnd : (HWND)!found_hwnd;
}



FResult SetHotkeyCriterion(HotCriterionType aType, LPCTSTR aWinTitle, LPCTSTR aWinText)
// Returns FR_FAIL if memory couldn't be allocated (and an error was raised), or OK otherwise.
// This is a global function because it's used by both hotkeys and hotstrings.
{
	HotkeyCriterion *cp = nullptr;
	if (  (*aWinTitle || *aWinText)
		&& !(cp = FindHotkeyCriterion(aType, aWinTitle, aWinText))
		&& !(cp = AddHotkeyCriterion(aType, aWinTitle, aWinText))  )
		return FR_FAIL;
	g->HotCriterion = cp;
	return OK;
}


HotkeyCriterion *FindHotkeyCriterion(HotCriterionType aType, LPCTSTR aWinTitle, LPCTSTR aWinText)
{
	// Storing combinations of WinTitle+WinText doesn't have as good a best-case memory savings as
	// have a separate linked list for Title vs. Text.  But it does save code size, and in the vast
	// majority of scripts, the memory savings would be insignificant.
	HotkeyCriterion *cp;
	for (cp = g_FirstHotCriterion; cp; cp = cp->NextCriterion)
		if (cp->Type == aType && !_tcscmp(cp->WinTitle, aWinTitle) && !_tcscmp(cp->WinText, aWinText)) // Case insensitive.
			return cp;

	return NULL;
}


HotkeyCriterion *AddHotkeyCriterion(HotCriterionType aType, LPCTSTR aWinTitle, LPCTSTR aWinText)
{
	HotkeyCriterion *cp;
	cp = SimpleHeap::Alloc<HotkeyCriterion>();
	cp->Type = aType;
	cp->OriginalExpr = nullptr;
	if (*aWinTitle)
	{
		if (   !(cp->WinTitle = SimpleHeap::Malloc(aWinTitle))   )
			return NULL;
	}
	else
		cp->WinTitle = _T("");
	if (*aWinText)
	{
		if (   !(cp->WinText = SimpleHeap::Malloc(aWinText))   )
			return NULL;
	}
	else
		cp->WinText = _T("");
	return AddHotkeyCriterion(cp);
}


HotkeyCriterion *AddHotkeyCriterion(HotkeyCriterion *cp)
{
	cp->NextCriterion = NULL;
	if (!g_FirstHotCriterion)
		g_FirstHotCriterion = g_LastHotCriterion = cp;
	else
	{
		g_LastHotCriterion->NextCriterion = cp;
		// This must be done after the above:
		g_LastHotCriterion = cp;
	}
	return cp;
}


HotkeyCriterion *AddHotkeyIfExpr()
{
	HotkeyCriterion *cp = SimpleHeap::Alloc<HotkeyCriterion>();
	cp->NextExpr = NULL;
	cp->OriginalExpr = nullptr;
	if (g_LastHotExpr)
		g_LastHotExpr->NextExpr = cp;
	else
		g_FirstHotExpr = cp;
	g_LastHotExpr = cp;
	return cp;
}


HotkeyCriterion *FindHotkeyIfExpr(LPCTSTR aExpr)
{
	for (HotkeyCriterion* cp = g_FirstHotExpr; cp; cp = cp->NextExpr)
		if (cp->OriginalExpr && !_tcscmp(aExpr, cp->OriginalExpr)) // Case-sensitive since the expression might be.
			return cp;
	return NULL;
}


void Script::PreparseHotkeyIfExpr(Line* aLine)
// Optimize simple #HotIf expressions into the more specific HOT_IF_ types so that they can be
// evaluated by the hook directly, without synchronizing with the main thread.
{
	ExprTokenType *postfix = aLine->mArg[0].postfix;
	if (postfix[0].symbol != SYM_OBJECT)
		return;
	auto fn = dynamic_cast<BuiltInFunc*>(postfix[0].object);
	if (!fn || fn->mBIF != &BIF_WinExistActive)
		return; // Not WinExist() or WinActive().
	++postfix;
	int param_count = 0;
	while (postfix[param_count].symbol == SYM_STRING)
		++param_count;
	if (postfix[param_count].symbol != SYM_FUNC // Not a function call, or it doesn't only accept strings.
		|| param_count > 2) // Too many parameters.
		return;
	bool invert = postfix[param_count+1].symbol == SYM_LOWNOT || postfix[param_count+1].symbol == SYM_HIGHNOT;
	if (postfix[param_count+1+invert].symbol != SYM_INVALID)
		return; // There's more to the expression.
	// Otherwise, it was a single call to WinExist() or WinActive() where each parameter
	// was exactly one literal string and the result was optionally inverted with "not" or "!".
	HotkeyCriterion *hc = (HotkeyCriterion *)aLine->mAttribute;
	// Change the parameters of this criterion.  FindHotkeyIfExpr() will still be able to
	// find it based on the expression text since it only relies on ExprLine.
	if (ctoupper(fn->mName[3]) == 'A')
		hc->Type = invert ? HOT_IF_NOT_ACTIVE : HOT_IF_ACTIVE;
	else
		hc->Type = invert ? HOT_IF_NOT_EXIST : HOT_IF_EXIST;
	hc->WinTitle = param_count > 0 ? postfix[0].marker : _T("");
	hc->WinText = param_count > 1 ? postfix[1].marker : _T("");
	// The following adds a duplicate in the event that there are two different expressions
	// which resolve to the same criterion, such as WinExist("x","") and WinExist("x", "").
	// In that case, only the first criterion can be referenced by its WinTitle & WinText
	// (but each can be referenced by its unique expression text).  This seems unavoidable
	// since variants are only unique to a given expression, and trying to work around that
	// here would cause inconsistency since this only applies to very specific expressions.
	// At this stage, the only criterion in the list are those added by the following line:
	AddHotkeyCriterion(hc);
}



void Hotkey::ManifestAllHotkeysHotstringsHooks()
// This function examines all hotkeys and hotstrings to determine:
// - Which hotkeys to register/unregister, or activate/deactivate in the hook.
// - Which hotkeys to be changed from HK_NORMAL to HK_KEYBD_HOOK (or vice versa).
// - In pursuit of the above, also assess the interdependencies between hotkeys: the presence or
//   absence of a given hotkey can sometimes impact whether other hotkeys need to be converted from
//   HK_NORMAL to HK_KEYBD_HOOK.  For example, a newly added/enabled global hotkey variant can
//   cause a HK_KEYBD_HOOK hotkey to become HK_NORMAL, and the converse is also true.
// - Based on the above, decide whether the keyboard and/or mouse hooks need to be (de)activated.
{
	// v1.0.37.05: A prefix key such as "a" in "a & b" should cause any use of "a" as a suffix
	// (such as ^!a) also to be a hook hotkey.  Otherwise, the ^!a hotkey won't fire because the
	// hook prevents the OS's hotkey monitor from seeing that the hotkey was pressed.  NOTE:
	// This is done only for virtual keys because prefix keys that are done by scan code (mModifierSC)
	// should already be hook hotkeys when used as suffix keys (there may be a few unusual exceptions,
	// but they seem too rare to justify the extra code size).
	// Update for v1.0.40: This first pass through the hotkeys now also checks things for hotkeys
	// that can affect other hotkeys. If this weren't done in the first pass, it might be possible
	// for a hotkey to make some other hotkey into a hook hotkey, but then the hook might not be
	// installed if that hotkey had already been processed earlier in the second pass.  Similarly,
	// a hotkey processed earlier in the second pass might have been registered when in fact it
	// should have become a hook hotkey due to something learned only later in the second pass.
	// Doing these types of things in the first pass resolves such situations.
	// Update for v1.1.27: Doing the above in the first pass doesn't work correctly, as mType is
	// reset to default during the first pass (even if a previous iteration might has set it to
	// HK_KEYBD_HOOK, such as when it is eclipsed by a wildcard hotkey).  One workaround would
	// be to set mKeybdHookMandatory = true, but that would prevent the hotkey from reverting to
	// HK_NORMAL when it no longer needs the hook.  Instead, there are now three passes.
	bool vk_is_prefix[VK_ARRAY_COUNT] = {false};
	bool *hk_is_inactive = (bool *)_alloca(sHotkeyCount * sizeof(bool)); // No init needed.  Currently limited to around 16k (HOTKEY_ID_MAX).
	HotkeyVariant *vp;
	int i, j;

	// FIRST PASS THROUGH THE HOTKEYS:
	for (i = 0; i < sHotkeyCount; ++i)
	{
		Hotkey &hot = *shk[i]; // For performance and convenience.
		if (   hk_is_inactive[i] = ((g_IsSuspended && !hot.IsExemptFromSuspend())
			|| hot.IsCompletelyDisabled())   ) // Listed last for short-circuit performance.
		{
			// In the cases above, nothing later below can change the fact that this hotkey should
			// now be in an unregistered state.
			if (hot.mIsRegistered)
			{
				hot.Unregister();
				// In case the hotkey's thread is already running, it seems best to cancel any repeat-run
				// that has already been scheduled.  Older comment: CT_SUSPEND, at least, relies on us to do this.
				for (vp = hot.mFirstVariant; vp; vp = vp->mNextVariant)
					vp->mRunAgainAfterFinished = false; // Applies to all hotkey types, not just registered ones.
			}
			continue;
		}
		// Otherwise, this hotkey will be in effect, so check its attributes.

		if (hot.mKeybdHookMandatory)
		{
			// v1.0.44: The following is relied upon by some things like the Hotkey constructor and the tilde prefix
			// (the latter can set mKeybdHookMandatory for a hotkey sometime after the first variant is added [such
			// as for a subsequent variant]).  This practice also improves maintainability.
			if (HK_TYPE_CAN_BECOME_KEYBD_HOOK(hot.mType)) // To ensure it hasn't since become a joystick/mouse/mouse-and-keyboard hotkey.
				hot.mType = HK_KEYBD_HOOK;
		}
		else // Hook isn't mandatory, so set any non-mouse/joystick/both hotkey to normal for possibly overriding later below.
		{
			// v1.0.42: The following is done to support situations in which a hotkey can be a hook hotkey sometimes,
			// but after a (de)suspend or after a change to other hotkeys via the Hotkey command, might no longer
			// require the hook.  Examples include:
			// 1) A hotkey can't be registered because some other app is using it, but later
			//    that condition changes.
			// 2) Suspend or the Hotkey command changes wildcard hotkeys so that non-wildcard
			//    hotkeys that have the same suffix are no longer eclipsed, and thus don't need
			//    the hook.  The same type of thing can happen if a key-up hotkey is disabled,
			//    which would allow it's key-down hotkey to become non-hook.  Similarly, if a
			//    all of a prefix key's hotkeys become disabled, and that prefix is also a suffix,
			//    those suffixes no longer need to be hook hotkeys.
			// 3) There may be other ways, especially in the future involving #HotIf WinActive/Exist
			//    keys whose criteria change.
			if (hot.mType == HK_KEYBD_HOOK)
				hot.mType = HK_NORMAL; // To possibly be overridden back to HK_KEYBD_HOOK later below; but if not, it will be registered later below.
		}

		if (hot.mModifierVK)
			vk_is_prefix[hot.mModifierVK] = true;
	} // End of first pass loop.
	
	// SECOND PASS THROUGH THE HOTKEYS:
	// Check for hotkeys that can affect other hotkeys, such as wildcard or key-up hotkeys.
	// This is separate to the other passes for reasons described at the top of the function.
	for (i = 0; i < sHotkeyCount; ++i)
	{
		if (hk_is_inactive[i])
			continue;
		Hotkey &hot = *shk[i]; // For performance and convenience.

		if (hot.mKeyUp && hot.mVK) // No need to do the below for mSC hotkeys since their down hotkeys would already be handled by the hook.
		{
			// For each key-up hotkey, search for any its counterpart that's a down-hotkey (if any).
			// Such a hotkey should also be handled by the hook because if registered, such as
			// "#5" (reg) and "#5 up" (hook), the hook would suppress the down event because it
			// is unaware that down-hotkey exists (it's suppressed to prevent the key from being
			// stuck in a logically down state).
			for (j = 0; j < sHotkeyCount; ++j)
			{
				// No need to check the following because they are already hook hotkeys:
				// mModifierVK/SC
				// mAllowExtraModifiers
				// mNoSuppress
				// In addition, there is no need to check shk[j]->mKeyUp because that can't be
				// true if it's mType is HK_NORMAL:
				// Also, g_IsSuspended and IsCompletelyDisabled() aren't checked
				// because it's harmless to operate on disabled hotkeys in this way.
				if (shk[j]->mVK == hot.mVK && HK_TYPE_CAN_BECOME_KEYBD_HOOK(shk[j]->mType) // Ordered for short-circuit performance.
					&& shk[j]->mModifiersConsolidatedLR == hot.mModifiersConsolidatedLR)
				{
					shk[j]->mType = HK_KEYBD_HOOK;
					// And if it's currently registered, it will be unregistered later below.
				}
			}
		}

		// v1.0.40: If this is a wildcard hotkey, any hotkeys it eclipses (i.e. includes as subsets)
		// should be made into hook hotkeys also, because otherwise they would be overridden by hook.
		// The following criteria are checked:
		// 1) Exclude those that have a ModifierSC/VK because in those cases, mAllowExtraModifiers is
		//    ignored.
		// 2) Exclude those that lack an mVK because those with mSC can't eclipse registered hotkeys
		//   (since any would-be eclipsed mSC hotkey is already a hook hotkey due to is SC nature).
		// 3) It must not have any mModifiersLR because such hotkeys can't completely eclipse
		//    registered hotkeys since they always have neutral vs. left/right-specific modifiers.
		//    For example, if *<^a is a hotkey, ^a can still be a registered hotkey because it could
		//    still be activated by pressing RControl+a.
		// 4) For maintainability, it doesn't check mNoSuppress because the hook is needed anyway,
		//    so might as well handle eclipsed hotkeys with it too.
		if (hot.mAllowExtraModifiers && hot.mVK && !hot.mModifiersLR && !(hot.mModifierSC || hot.mModifierVK))
		{
			for (j = 0; j < sHotkeyCount; ++j)
			{
				// If it's not of type HK_NORMAL, there's no need to change its type regardless
				// of the values of its other members.  Also, if the wildcard hotkey (hot) has
				// any neutral modifiers, this hotkey must have at least those neutral modifiers
				// too or else it's not eclipsed (and thus registering it is okay).  In other words,
				// the wildcard hotkey's neutral modifiers must be a perfect subset of this hotkey's
				// modifiers for this one to be eclipsed by it. Note: Neither mModifiersLR nor
				// mModifiersConsolidated is checked for simplicity and also because it seems to add
				// flexibility.  For example, *<^>^a would require both left AND right ctrl to be down,
				// not EITHER. In other words, mModifiersLR can never in effect contain a neutral modifier.
				if (shk[j]->mVK == hot.mVK && HK_TYPE_CAN_BECOME_KEYBD_HOOK(shk[j]->mType) // Ordered for short-circuit performance.
					&& (hot.mModifiers & shk[j]->mModifiers) == hot.mModifiers)
				{
					// Note: No need to check mModifiersLR because it would already be a hook hotkey in that case;
					// that is, the check of shk[j]->mType precludes it.  It also precludes the possibility
					// of shk[j] being a key-up hotkey, wildcard hotkey, etc.
					shk[j]->mType = HK_KEYBD_HOOK;
					// And if it's currently registered, it will be unregistered later below.
				}
			}
		}
	} // End of second pass loop.

	// THIRD PASS THROUGH THE HOTKEYS:
	// v1.0.42: Reset sWhichHookNeeded because it's now possible that the hook was on before but no longer
	// needed due to changing of a hotkey from hook to registered (for various reasons described above):
	sWhichHookNeeded = 0;
	for (i = 0; i < sHotkeyCount; ++i)
	{
		if (hk_is_inactive[i])
			continue; // v1.0.40: Treat disabled hotkeys as though they're not even present.
		Hotkey &hot = *shk[i]; // For performance and convenience.

		// HK_MOUSE_HOOK hotkeys, and most HK_KEYBD_HOOK hotkeys, are handled by the hotkey constructor.
		// What we do here upgrade any NORMAL/registered hotkey to HK_KEYBD_HOOK if there are other
		// hotkeys that interact or overlap with it in such a way that the hook is preferred.
		// This evaluation is done here because only now that hotkeys are about to be activated do
		// we know which ones are disabled or suspended, and thus don't need to be taken into account.
		if (HK_TYPE_CAN_BECOME_KEYBD_HOOK(hot.mType))
		{
			if (vk_is_prefix[hot.mVK])
				// If it's a suffix that is also used as a prefix, use hook (this allows ^!a to work without $ when "a & b" is a hotkey).
				// v1.0.42: This was fixed so that mVK_WasSpecifiedByNumber dosn't affect it.  That is, a suffix that's
				// also used as a prefix should become a hook hotkey even if the suffix is specified as "vkNNN::".
				hot.mType = HK_KEYBD_HOOK;
				// And if it's currently registered, it will be unregistered later below.
			else
			{
				// v1.0.42: Any #HotIf keyboard hotkey must use the hook if it lacks an enabled,
				// non-suspended, global variant.  Under those conditions, the hotkey is either:
				// 1) Single-variant hotkey that has criteria (non-global).
				// 2) Multi-variant hotkey but all variants have criteria (non-global).
				// 3) A hotkey with a non-suppressed (~) variant (always, for code simplicity): already handled by AddVariant().
				// In both cases above, the hook must handle the hotkey because there can be
				// situations in which the hook should let the hotkey's keystroke pass through
				// to the active window (i.e. the hook is needed to dynamically disable the hotkey).
				// mHookAction isn't checked here since those hotkeys shouldn't reach this stage (since they're always hook hotkeys).
				for (hot.mType = HK_KEYBD_HOOK, vp = hot.mFirstVariant; vp; vp = vp->mNextVariant)
				{
					if (   !vp->mHotCriterion && vp->mEnabled // It's a global variant (no criteria) and it's enabled...
						&& (!g_IsSuspended || vp->mSuspendExempt)   )
						// ... and this variant isn't suspended (we already know IsCompletelyDisabled()==false from an earlier check).
					{
						hot.mType = HK_NORMAL; // Reset back to how it was before this loop started.  Hook not needed.
						break;
					}
				}
				// If the above promoted it from NORMAL to HOOK but the hotkey is currently registered,
				// it will be unregistered later below.
			}
		}

		// Check if this mouse hotkey also requires the keyboard hook (e.g. #LButton).
		// Some mouse hotkeys, such as those with normal modifiers, don't require it
		// since the mouse hook has logic to handle that situation.  But those that
		// are composite hotkeys such as "RButton & Space" or "Space & RButton" need
		// the keyboard hook:
		if (hot.mType == HK_MOUSE_HOOK && (
			hot.mModifierSC || hot.mSC // i.e. since it's an SC, the modifying key isn't a mouse button.
			|| hot.mHookAction // v1.0.25.05: At least some alt-tab actions require the keyboard hook. For example, a script consisting only of "MButton::AltTabAndMenu" would not work properly otherwise.
			// v1.0.25.05: The line below was added to prevent the Start Menu from appearing, which
			// requires the keyboard hook. ALT hotkeys don't need it because the mouse hook sends
			// a CTRL keystroke to disguise them, a trick that is unfortunately not reliable for
			// when it happens while the while key is down (though it does disguise a Win-up).
			|| ((hot.mModifiersConsolidatedLR & (MOD_LWIN|MOD_RWIN)) && !(hot.mModifiersConsolidatedLR & (MOD_LALT|MOD_RALT)))
			// For v1.0.30, above has been expanded to include Win+Shift and Win+Control modifiers.
			|| (hot.mVK && !IsMouseVK(hot.mVK)) // e.g. "RButton & Space"
			|| (hot.mModifierVK && !IsMouseVK(hot.mModifierVK)))   ) // e.g. "Space & RButton"
			hot.mType = HK_BOTH_HOOKS;  // Needed by ChangeHookState().
			// For the above, the following types of mouse hotkeys do not need the keyboard hook:
			// 1) mAllowExtraModifiers: Already handled since the mouse hook fetches the modifier state
			//    manually when the keyboard hook isn't installed.
			// 2) mModifiersConsolidatedLR (i.e. the mouse button is modified by a normal modifier
			//    such as CTRL): Same reason as #1.
			// 3) As a subset of #2, mouse hotkeys that use WIN as a modifier will not have the
			//    Start Menu suppressed unless the keyboard hook is installed.  It's debatable,
			//    but that seems a small price to pay (esp. given how rare it is just to have
			//    the mouse hook with no keyboard hook) to avoid the overhead of the keyboard hook.
		
		// If the hotkey is normal, try to register it.  If the register fails, use the hook to try
		// to override any other script or program that might have it registered (as documented):
		if (hot.mType == HK_NORMAL)
		{
			if (!hot.Register()) // Can't register it, usually due to some other application or the OS using it.
				hot.mType = HK_KEYBD_HOOK;
		}
		else // mType isn't NORMAL (possibly due to something above changing it), so ensure it isn't registered.
			if (hot.mIsRegistered) // Improves typical performance since this hotkey could be mouse, joystick, etc.
				// Although the hook effectively overrides registered hotkeys, they should be unregistered anyway
				// to prevent the Send command from triggering the hotkey, and perhaps other side-effects.
				hot.Unregister();

		switch (hot.mType)
		{
		case HK_KEYBD_HOOK: sWhichHookNeeded |= HOOK_KEYBD; break;
		case HK_MOUSE_HOOK: sWhichHookNeeded |= HOOK_MOUSE; break;
		case HK_BOTH_HOOKS: sWhichHookNeeded |= HOOK_KEYBD|HOOK_MOUSE; break;
		}
	} // for()

	// Check if anything else requires the hook.
	// But do this part outside of the above block because these values may have changed since
	// this function was first called.  By design, the Num/Scroll/CapsLock AlwaysOn/Off setting
	// stays in effect even when Suspend in ON.
	if (   Hotstring::sEnabledCount
		|| g_input // v1.0.91: Hook is needed for collecting input.
		|| !(g_ForceNumLock == NEUTRAL && g_ForceCapsLock == NEUTRAL && g_ForceScrollLock == NEUTRAL)   )
		sWhichHookNeeded |= HOOK_KEYBD;
	if (g_BlockMouseMove || (g_HSResetUponMouseClick && Hotstring::sEnabledCount))
		sWhichHookNeeded |= HOOK_MOUSE;

	// Install or deinstall either or both hooks, if necessary, based on these param values.
	ChangeHookState(shk, sHotkeyCount, sWhichHookNeeded, sWhichHookAlways);

	// Fix for v1.0.34: If the auto-execute section uses the Hotkey command but returns before doing
	// something that calls MsgSleep, the main timer won't have been turned on.  For example:
	// Hotkey, Joy1, MySubroutine
	// ;Sleep 1  ; This was a workaround before this fix.
	// return
	// By putting the following check here rather than in AutoHotkey.cpp, that problem is resolved.
	// In addition...
	if (sJoyHotkeyCount)  // Joystick hotkeys require the timer to be always on.
		SET_MAIN_TIMER
}



void Hotkey::MaybeUninstallHook()
// Caller knows that one of the users of the keyboard hook no longer requires it,
// and wants it uninstalled if it is no longer needed by anything else.
{
	// Do some quick checks to avoid scanning all hotkeys unnecessarily:
	if (g_input || Hotstring::sEnabledCount || (sWhichHookAlways & HOOK_KEYBD))
		return;
	// Do more thorough checking to determine whether the hook is still needed:
	ManifestAllHotkeysHotstringsHooks();
}



void Hotkey::AllDestruct()
{
	// MSDN: "Before terminating, an application must call the UnhookWindowsHookEx function to free
	// system resources associated with the hook."
	AddRemoveHooks(0); // Remove all hooks. By contrast, registered hotkeys are unregistered below.
	if (g_PlaybackHook) // Would be unusual for this to be installed during exit, but should be checked for completeness.
		UnhookWindowsHookEx(g_PlaybackHook);
	for (int i = 0; i < sHotkeyCount; ++i)
		delete shk[i]; // Unregisters before destroying.
}



bool Hotkey::PrefixHasNoEnabledSuffixes(int aVKorSC, bool aIsSC, bool &aSuppress)
// aVKorSC contains the virtual key or scan code of the specified prefix key (it's a scan code if aIsSC is true).
// Returns true if this prefix key has no suffixes that can possibly fire.  Each such suffix is prevented from
// firing by one or more of the following:
// 1) Hotkey is completely disabled via IsCompletelyDisabled().
// 2) Hotkey has criterion and those criterion do not allow the hotkey to fire.
// Caller is expected to set aSuppress to a default of false.
{
	// v1.0.44: Added aAsModifier so that a pair of hotkeys such as:
	//   LControl::tooltip LControl
	//   <^c::tooltip ^c
	// ...works as it did in versions prior to 1.0.41, namely that LControl fires on key-up rather than
	// down because it is considered a prefix key for the <^c hotkey .
	modLR_type aAsModifier = KeyToModifiersLR(aIsSC ? 0 : aVKorSC, aIsSC ? aVKorSC : 0, NULL);

	bool has_enabled_suffix = false;

	for (int i = 0; i < sHotkeyCount; ++i)
	{
		Hotkey &hk = *shk[i];
		if (aVKorSC != (aIsSC ? hk.mModifierSC : hk.mModifierVK) && !(aAsModifier & hk.mModifiersLR)
			|| hk.IsCompletelyDisabled())
			continue; // This hotkey isn't enabled or it doesn't use the specified key as a prefix.  No further checking for it.
		if (hk.mHookAction)
		{
			if (g_IsSuspended)
				// An alt-tab hotkey (non-NULL mHookAction) is always suspended when g_IsSuspended==true because
				// alt-tab hotkeys have no subroutine capable of making them exempt.  So g_IsSuspended is checked
				// for alt-tab hotkeys here; and for other types of hotkeys, it's checked further below.
				continue;
			//else // This alt-tab hotkey is currently active.
			if ((hk.mNoSuppress & NO_SUPPRESS_PREFIX) || aSuppress)
				return false; // Since any stored mHotCriterion are ignored for alt-tab hotkeys, no further checking is needed.
			has_enabled_suffix = true;
			continue; // Still need to check other hotkeys for NO_SUPPRESS_PREFIX.
		}
		if (has_enabled_suffix && !(hk.mNoSuppress & NO_SUPPRESS_PREFIX))
			continue; // No need to evaluate this hotkey's variants.
		// Otherwise, find out if any of its variants is eligible to fire.  If so, immediately return
		// false because even one eligible hotkey means this prefix is enabled.
		for (HotkeyVariant *vp = hk.mFirstVariant; vp; vp = vp->mNextVariant)
			// v1.0.42: Fixed to take into account whether the hotkey is suspended (previously it only checked
			// whether the hotkey was enabled (above), which isn't enough):
			if (   vp->mEnabled // This particular variant within its parent hotkey is enabled.
				&& (!g_IsSuspended || vp->mSuspendExempt) // This variant isn't suspended...
				&& (!vp->mHotCriterion || HotCriterionAllowsFiring(vp->mHotCriterion, hk.mName))   ) // ... and its criteria allow it to fire.
			{
				if ((vp->mNoSuppress & NO_SUPPRESS_PREFIX) || aSuppress)
					return false; // At least one of this prefix's suffixes is eligible for firing.
				has_enabled_suffix = true;
				if (!(hk.mNoSuppress & NO_SUPPRESS_PREFIX))
					break; // None of this hotkey's variants have NO_SUPPRESS_PREFIX.
				// Keep checking to ensure no other enabled variants have NO_SUPPRESS_PREFIX.
			}
	}
	// Since above didn't return, either no hotkeys were found for this prefix that are capable of firing,
	// or no variants were found with the NO_SUPPRESS_PREFIX flag.
	aSuppress = has_enabled_suffix;
	return !has_enabled_suffix;
}



HotkeyVariant *Hotkey::CriterionAllowsFiring(HWND *aFoundHWND, ULONG_PTR aExtraInfo, LPTSTR aSingleChar)
// Caller must not call this for AltTab hotkeys IDs because this will always return NULL in such cases.
// Returns the address of the first matching non-global hotkey variant that is allowed to fire.
// If there is no non-global one eligible, the global one is returned (or NULL if none).
// If non-NULL, aFoundHWND is an output variable for the caller, but it is only set if a
// non-global/criterion variant is found; that is, it isn't changed when no match is found or
// when the match is a global variant.  Even when set, aFoundHWND will be (HWND)1 for
// "not-criteria" such as #HotIf Not WinActive().
{
	// Check mParentEnabled in case the hotkey became disabled between the time the message was posted
	// and the time it arrived.  A similar check is done for "suspend" later below (since "suspend"
	// is a per-variant attribute).
	if (!mParentEnabled) // IsCompletelyDisabled() isn't called because the loop below checks all the mEnabled flags, no need to do it twice.
		return NULL;

	HWND unused;
	HWND &found_hwnd = aFoundHWND ? *aFoundHWND : unused;  // To simplify other things.
	found_hwnd = NULL;  // Set default output parameter for caller (in case loop never sets it).
	HotkeyVariant *vp, *vp_to_fire;

	// aHookAction isn't checked because this should never be called for alt-tab hotkeys (see other comments above).
	for (vp_to_fire = NULL, vp = mFirstVariant; vp; vp = vp->mNextVariant)
	{
		// Technically, g_IsSuspended needs to be checked only if our caller is TriggerJoyHotkeys()
		// because other callers would never have received the hotkey message in the first place.
		// However, since it's possible for a hotkey to become suspended between the time its hotkey
		// message is posted and the time it is fetched and processed, aborting the firing seems
		// like the best choice for the widest variety of circumstances (even though it's a departure
		// from the behavior in previous versions).  Another reason to check g_IsSuspended unconditionally
		// is for maintainability and code size reduction.  Finally, it's unlikely to significantly
		// impact performance since the vast majority of hotkeys have either one or just a few variants.
		if (   vp->mEnabled // This particular variant within its parent hotkey is enabled.
			&& (!g_IsSuspended || vp->mSuspendExempt) // This variant isn't suspended...
			&& HotInputLevelAllowsFiring(vp->mInputLevel, aExtraInfo, aSingleChar) // ... its #InputLevel allows it to fire...
			&& (!vp->mHotCriterion || (found_hwnd = HotCriterionAllowsFiring(vp->mHotCriterion, mName)))   ) // ... and its criteria allow it to fire.
		{
			if (vp->mHotCriterion) // Since this is the first criteria hotkey, it takes precedence.
				return vp;
			//else since this is variant has no criteria, let the first criteria variant in the list
			// take precedence over it (if there is one).  If none is found, the vp_to_fire will stay
			// set as the non-criterion variant.
			vp_to_fire = vp;
		}
	}
	return vp_to_fire; // Either NULL or the variant found by the loop.
}

bool HotInputLevelAllowsFiring(SendLevelType inputLevel, ULONG_PTR aEventExtraInfo, LPTSTR aKeyHistoryChar)
{
	if (InputLevelFromInfo(aEventExtraInfo) <= inputLevel)
	{
		if (aKeyHistoryChar)
			*aKeyHistoryChar = 'i'; // Mark as ignored in KeyHistory
		return false;
	}
	return true;
}


HotkeyVariant *Hotkey::CriterionFiringIsCertain(HotkeyIDType &aHotkeyIDwithFlags, bool aKeyUp, ULONG_PTR aExtraInfo
	, UCHAR &aNoSuppress, bool &aFireWithNoSuppress, LPTSTR aSingleChar)
// v1.0.44: Caller has ensured that aFireWithNoSuppress is true if has already been decided and false if undecided.
// Upon return, caller can assume that the value in it is now decided rather than undecided.
// v1.0.42: Caller must not call this for AltTab hotkeys IDs, but this will always return NULL in such cases.
// aHotkeyToFireUponRelease is sometimes modified for the caller here, as is *aSingleChar (if aSingleChar isn't NULL).
// Caller has ensured that aHotkeyIDwithFlags contains a valid/existing hotkey ID.
// Technically, aHotkeyIDwithMask can be with or without the flags in the high bits.
// If present, they're removed.
{
	// aHookAction isn't checked because this should never be called for alt-tab hotkeys (see other comments above).
	HotkeyIDType hotkey_id = aHotkeyIDwithFlags & HOTKEY_ID_MASK;
	// The following check is for maintainability, since caller should have already checked and
	// handled HOTKEY_ID_ALT_TAB and similar.  Less-than-zero check not necessary because it's unsigned.
	if (hotkey_id >= sHotkeyCount)
		return NULL; // Special alt-tab hotkey quasi-ID used by the hook.
	Hotkey &hk = *shk[hotkey_id]; // For convenience and performance.

	if (aFireWithNoSuppress // Caller has already determined its value with certainty...
		|| (hk.mNoSuppress & NO_SUPPRESS_SUFFIX_VARIES) != NO_SUPPRESS_SUFFIX_VARIES) // ...or its value is easy to determine, so do it now (compare to itself since it's a bitwise union).
	{
		// Since aFireWithNoSuppress can now be easily determined for the caller (or was already determined by the caller
		// itself), it's possible to take advantage of the following optimization, which is especially important in cases
		// where TitleMatchMode is "slow":
		// For performance, the following returns without having called WinExist/Active if it sees that one of this
		// hotkey's variant's will certainly fire due to the fact that it has a non-suspended global variant.
		// This reduces the number of situations in which double the number of WinExist/Active() calls are made
		// (once in the hook to determine whether the hotkey keystroke should be passed through to the active window,
		// and again upon receipt of the message for reasons explained there).
		for (HotkeyVariant *vp = hk.mFirstVariant; vp; vp = vp->mNextVariant)
			if (!vp->mHotCriterion && vp->mEnabled && (!g_IsSuspended || vp->mSuspendExempt)
				&& HotInputLevelAllowsFiring(vp->mInputLevel, aExtraInfo, aSingleChar))
			{
				// Fix for v1.0.47.02: The following section (above "return") was moved into this block
				// from above the for() because only when this for() returns is it certain that this
				// hk/hotkey_id is actually the right one, and thus its attributes can be used to determine
				// aFireWithNoSuppress for the caller.
				// Since this hotkey has variants only of one type (tilde or non-tilde), this variant must be of that type.
				if (!aFireWithNoSuppress) // Caller hasn't yet determined its value with certainty (currently, this statement might always be true).
					aFireWithNoSuppress = (hk.mNoSuppress & AT_LEAST_ONE_VARIANT_HAS_TILDE); // Due to other checks, this means all variants are tilde.
				return vp; // Caller knows this isn't necessarily the variant that will fire since !vp->mHotCriterion.
			}
	}

	// Since above didn't return, a slower method is needed to find out which variant of this hotkey (if any)
	// should fire.
	HotkeyVariant *vp;
	if (vp = hk.CriterionAllowsFiring(NULL, aExtraInfo, aSingleChar))
	{
		if (!aFireWithNoSuppress) // Caller hasn't yet determined its value with certainty (currently, this statement might always be true).
			aFireWithNoSuppress = (vp->mNoSuppress & AT_LEAST_ONE_VARIANT_HAS_TILDE);
		return vp; // It found an eligible variant to fire.
	}

	// Since above didn't find any variant of the hotkey than can fire, check for other eligible hotkeys.
	if (!hk.mHookAction) // Rule out those that aren't susceptible to the bug.
	{
		// Custom combos are no longer ruled out by the above since they allow extra modifiers and
		// are capable of obscuring non-custom combos; e.g. LCtrl & a:: obscures <^a::, ^+a:: and so on.
		// Fix for v1.0.46.13: Although the section higher above found no variant to fire for the
		// caller-specified hotkey ID, it's possible that some other hotkey (one with a wildcard) is
		// eligible to fire due to the eclipsing behavior of wildcard hotkeys.  For example:
		//    #HotIf Not WinActive("Untitled")
		//    q::tooltip ThisHotkey . " Non-notepad"
		//    #HotIf WinActive("Untitled")
		//    *q::tooltip ThisHotkey . " Notepad"
		// However, the logic here might not be a perfect solution because it fires the first available
		// hotkey that has a variant whose criteria are met (which might not be exactly the desired rules
		// of precedence).  However, I think it's extremely rare that there would be more than one hotkey
		// that matches the original hotkey (VK, SC, has-wildcard) etc.  Even in the rare cases that there
		// is more than one, the rarity is compounded by the rarity of the bug even occurring, which probably
		// makes the odds vanishingly small.  That's why the following simple, high-performance loop is used
		// rather than more a more complex one that "locates the smallest (most specific) eclipsed wildcard
		// hotkey", or "the uppermost variant among all eclipsed wildcards that is eligible to fire".
		// UPDATE: This now uses a linked list of hotkeys which share the same suffix key, in the order of
		// sort_most_general_before_least, which might solve the concern about precedence.
		mod_type modifiers = ConvertModifiersLR(g_modifiersLR_logical_non_ignored); // Neutral modifiers.
		for (HotkeyIDType candidate_id = hk.mNextHotkey; candidate_id != HOTKEY_ID_INVALID; )
		{
			Hotkey &hk2 = *shk[candidate_id]; // For performance and convenience.
			candidate_id = hk2.mNextHotkey;
			// Non-wildcard hotkeys are eligible for the workaround in cases like ^+a vs <^+a vs ^<+a, where
			// the neutral modifier acts as a sort of wildcard (it permits left, right or both).  This also
			// increases support for varying names, such as Esc vs. Escape vs. vk1B (which already partially
			// worked if wildcards were used).
			// However, must ensure only the allowed modifiers are down when !mAllowExtraModifiers.
			// mVK and mSC aren't checked since the linked list only includes hotkeys for this same suffix key.
			// This also allows the workaround to be partially applied to LCtrl vs. Ctrl and similar (as suffixes).
			if (  (hk2.mAllowExtraModifiers || !(~hk2.mModifiersConsolidatedLR & g_modifiersLR_logical_non_ignored))
				&& hk2.mKeyUp == hk.mKeyUp // Seems necessary that up/down nature is the same in both.
				&& !hk2.mModifierVK // Avoid accidental matching of normal hotkeys with custom-combo "&"
				&& !hk2.mModifierSC // hotkeys that happen to have the same mVK/SC.
				&& !hk2.mHookAction // Might be unnecessary to check this; but just in case.
				&& hk2.mID != hotkey_id // Don't consider the original hotkey because it was already found ineligible.
				&& !(hk2.mModifiers & ~modifiers) // All neutral modifiers required by the candidate are pressed.
				&& !(hk2.mModifiersLR & ~g_modifiersLR_logical_non_ignored) // All left-right specific modifiers required by the candidate are pressed.
				//&& hk2.mType != HK_JOYSTICK // Seems unnecessary since joystick hotkeys don't call us and even if they did, probably shouldn't be included.
				//&& hk2.mParentEnabled   ) // CriterionAllowsFiring() will check this for us.
				)
			{
				// The following section is similar to one higher above, so maintain them together:
				if (vp = hk2.CriterionAllowsFiring(NULL, aExtraInfo, aSingleChar))
				{
					if (!aFireWithNoSuppress) // Caller hasn't yet determined its value with certainty (currently, this statement might always be true).
						aFireWithNoSuppress = (vp->mNoSuppress & AT_LEAST_ONE_VARIANT_HAS_TILDE);
					aHotkeyIDwithFlags = hk2.mID; // Caller currently doesn't need the flags put onto it, so they're omitted.
					return vp; // It found an eligible variant to fire.
				}
			}
		}
	}

	// Otherwise, this hotkey has no variants that can fire.  Caller wants a few things updated in that case.
	if (!aFireWithNoSuppress) // Caller hasn't yet determined its value with certainty.
		aFireWithNoSuppress = true; // Fix for v1.0.47.04: Added this line and the one above to fix the fact that a context-sensitive hotkey like "a UP::" would block the down-event of that key even when the right window/criteria aren't met.
	// If this is a key-down hotkey:
	// Leave aHotkeyToFireUponRelease set to whatever it was so that the criteria are
	// evaluated later, at the time of release.  It seems more correct that way, though the actual
	// change (hopefully improvement) in usability is unknown.
	// Since the down-event of this key won't be suppressed, it seems best never to suppress the
	// key-up hotkey (if it has one), if nothing else than to be sure the logical key state of that
	// key as shown by GetAsyncKeyState() returns the correct value (for modifiers, this is even more
	// important since them getting stuck down causes undesirable behavior).  If it doesn't have a
	// key-up hotkey, the up-keystroke should wind up being non-suppressed anyway due to default
	// processing).
	if (!aKeyUp)
		aNoSuppress |= NO_SUPPRESS_NEXT_UP_EVENT;  // Update output parameter for the caller.
	if (aSingleChar && *aSingleChar != 'i') // 'i' takes precedence because it's used to detect when #InputLevel prevented the hotkey from firing, to prevent it from being suppressed.
		*aSingleChar = '#'; // '#' in KeyHistory to indicate this hotkey is disabled due to #HotIf WinActive/Exist() criterion.
	return NULL;
}



HotkeyIDType Hotkey::FindPairedHotkey(HotkeyIDType aFirstID, modLR_type aModsLR, bool aKeyUp)
{
	mod_type modifiers = ConvertModifiersLR(aModsLR); // Neutral modifiers.
	for (HotkeyIDType candidate_id = aFirstID; candidate_id != HOTKEY_ID_INVALID; )
	{
		Hotkey &hk2 = *shk[candidate_id]; // For performance and convenience.
		candidate_id = hk2.mNextHotkey;
		if (  (hk2.mAllowExtraModifiers || !(~hk2.mModifiersConsolidatedLR & aModsLR))
			&& hk2.mKeyUp == aKeyUp
			&& !hk2.mModifierVK // Avoid accidental matching of normal hotkeys with custom-combo "&"
			&& !hk2.mModifierSC // hotkeys that happen to have the same mVK/SC.
			&& !hk2.mHookAction // Might be unnecessary to check this; but just in case.
			&& !(hk2.mModifiers & ~modifiers) // All neutral modifiers required by the candidate are pressed.
			&& !(hk2.mModifiersLR & ~aModsLR) // All left-right specific modifiers required by the candidate are pressed.
			//&& hk2.mParentEnabled // CriterionAllowsFiring() will check this for us.
			)
			return aKeyUp ? (hk2.mID | HOTKEY_KEY_UP) : hk2.mID;
	}
	return HOTKEY_ID_INVALID;
}



modLR_type Hotkey::HotkeyRequiresModLR(HotkeyIDType aHotkeyID, modLR_type aModLR)
{
	if (aHotkeyID >= sHotkeyCount)
		return 0;
	return shk[aHotkeyID]->mModifiersConsolidatedLR & aModLR;
}



void Hotkey::TriggerJoyHotkeys(int aJoystickID, DWORD aButtonsNewlyDown)
{
	for (int i = 0; i < sHotkeyCount; ++i)
	{
		Hotkey &hk = *shk[i]; // For performance and convenience.
		// Fix for v1.0.34: If hotkey isn't enabled, or hotkeys are suspended and this one isn't
		// exempt, don't fire it.  These checks are necessary only for joystick hotkeys because 
		// normal hotkeys are completely deactivated when turned off or suspended, but the joystick
		// is still polled even when some joystick hotkeys are disabled.  UPDATE: In v1.0.42, Suspend
		// is checked upon receipt of the message, not here upon sending.
		if (hk.mType == HK_JOYSTICK && hk.mVK == aJoystickID
			&& (aButtonsNewlyDown & ((DWORD)0x01 << (hk.mSC - JOYCTRL_1)))) // This hotkey's button is among those newly pressed.
		{
			// Criteria are checked, and variant determined, upon arrival of message rather than when sending
			// ("suspend" is also checked then).  This is because joystick button presses are never hidden
			// from the active window (the concept really doesn't apply), so not checking here avoids the
			// performance loss of a second check (the loss can be significant in the case of
			// "SetTitleMatchMode Slow").
			//
			// Post it to the thread because the message pump itself (not the WindowProc) will handle it.
			// UPDATE: Posting to NULL would have a risk of discarding the message if a MsgBox pump or
			// pump other than MsgSleep() were running.  The only reason it doesn't is that this function
			// is only ever called from MsgSleep(), which is careful to process all messages (at least
			// those that aren't kept queued due to the message filter) prior to returning to its caller.
			// But for maintainability, it seems best to change this to g_hWnd vs. NULL to make joystick
			// hotkeys behave more like standard hotkeys.
			PostMessage(g_hWnd, WM_HOTKEY, (WPARAM)i, 0);
		}
		//else continue the loop in case the user has newly pressed more than one joystick button.
	}
}



void Hotkey::PerformInNewThreadMadeByCaller(HotkeyVariant &aVariant)
// Caller is responsible for having called PerformIsAllowed() before calling us.
// Caller must have already created a new thread for us, and must close the thread when we return.
{
	static bool sDialogIsDisplayed = false;  // Prevents double-display caused by key buffering.
	if (sDialogIsDisplayed) // Another recursion layer is already displaying the warning dialog below.
		return; // Don't allow new hotkeys to fire during that time.

	// Help prevent runaway hotkeys (infinite loops due to recursion in bad script files):
	static UINT throttled_key_count = 0;  // This var doesn't belong in struct since it's used only here.
	UINT time_until_now;
	int display_warning;
	if (!sTimePrev)
		sTimePrev = GetTickCount();

	++throttled_key_count;
	sTimeNow = GetTickCount();
	// Calculate the amount of time since the last reset of the sliding interval.
	// Note: A tickcount in the past can be subtracted from one in the future to find
	// the true difference between them, even if the system's uptime is greater than
	// 49 days and the future one has wrapped but the past one hasn't.  This is
	// due to the nature of DWORD subtraction.  The only time this calculation will be
	// unreliable is when the true difference between the past and future
	// tickcounts itself is greater than about 49 days:
	time_until_now = (sTimeNow - sTimePrev);
	if (display_warning = (throttled_key_count > g_MaxHotkeysPerInterval
		&& time_until_now < g_HotkeyThrottleInterval))
	{
		// The moment any dialog is displayed, hotkey processing is halted since this
		// app currently has only one thread.
		TCHAR error_text[2048];
		// Using %f with wsprintf() yields a floating point runtime error dialog.
		// UPDATE: That happens if you don't cast to float, or don't have a float var
		// involved somewhere.  Avoiding floats altogether may reduce EXE size
		// and maybe other benefits (due to it not being "loaded")?
		sntprintf(error_text, _countof(error_text), _T("%u hotkeys have been received in the last %ums.\n\n")
			_T("Do you want to continue?\n(see A_MaxHotkeysPerInterval in the help file)")  // In case its stuck in a loop.
			, throttled_key_count, time_until_now);

		// Turn off any RunAgain flags that may be on, which in essence is the same as de-buffering
		// any pending hotkey keystrokes that haven't yet been fired:
		ResetRunAgainAfterFinished();

		// This is now needed since hotkeys can still fire while a messagebox is displayed.
		// Seems safest to do this even if it isn't always necessary:
		sDialogIsDisplayed = true;
		g_AllowInterruption = FALSE;
		if (MsgBox(error_text, MB_YESNO) == IDNO)
			g_script.ExitApp(EXIT_CLOSE); // Might not actually Exit if there's an OnExit function.
		g_AllowInterruption = TRUE;
		sDialogIsDisplayed = false;
	}
	// The display_warning var is needed due to the fact that there's an OR in this condition:
	if (display_warning || time_until_now > g_HotkeyThrottleInterval)
	{
		// Reset the sliding interval whenever it expires.  Doing it this way makes the
		// sliding interval more sensitive than alternate methods might be.
		// Also reset it if a warning was displayed, since in that case it didn't expire.
		throttled_key_count = 0;
		sTimePrev = sTimeNow;
	}
	if (display_warning)
		// At this point, even though the user chose to continue, it seems safest
		// to ignore this particular hotkey event since it might be WinClose or some
		// other command that would have unpredictable results due to the displaying
		// of the dialog itself.
		return;


	// This is stored as an attribute of the script (semi-globally) rather than passed
	// as a parameter to ExecUntil (and from their on to any calls to SendKeys() that it
	// makes) because it's possible for SendKeys to be called asynchronously, namely
	// by a timed subroutine, while A_HotkeyModifierTimeout is still in effect,
	// in which case we would want SendKeys() to take note of these modifiers even
	// if it was called from an ExecUntil() other than ours here:
	g_script.mThisHotkeyModifiersLR = mModifiersConsolidatedLR;

	// LAUNCH HOTKEY SUBROUTINE:
	++aVariant.mExistingThreads;  // This is the thread count for this particular hotkey only.

	ExprTokenType params = { mName };
	ResultType result = aVariant.mCallback->ExecuteInNewThread(g_script.mThisHotkeyName, &params, 1);
	
	--aVariant.mExistingThreads;

	if (result == FAIL)
		aVariant.mRunAgainAfterFinished = false;  // Ensure this is reset due to the error.
	else if (aVariant.mRunAgainAfterFinished)
	{
		// But MsgSleep() can change it back to true again, when called by the above call
		// to ExecUntil(), to keep it auto-repeating:
		aVariant.mRunAgainAfterFinished = false;  // i.e. this "run again" ticket has now been used up.
		if (GetTickCount() - aVariant.mRunAgainTime <= 1000)
		{
			// v1.0.44.14: Post a message rather than directly running the above ExecUntil again.
			// This fixes unreported bugs in previous versions where the thread isn't reinitialized before
			// the launch of one of these buffered hotkeys, which caused settings such as SetKeyDelay
			// not to start off at their defaults.  Also, there are quite a few other things that the main
			// event loop does to prep for the launch of a hotkey.  Rather than copying them here or
			// trying to put them into a shared function (which would be difficult due to their nature),
			// it's much more maintainable to post a message, and in most cases, it shouldn't measurably
			// affect response time (this feature is rarely used anyway).
			PostMessage(g_hWnd, WM_HOTKEY, (WPARAM)mID, 0);
		}
		//else it was posted too long ago, so don't do it.  This is because most users wouldn't
		// want a buffered hotkey to stay pending for a long time after it was pressed, because
		// that might lead to unexpected behavior.
	}
}



FResult Hotkey::IfExpr(IObject *aExprObj)
// HotIf ; Set null criterion.
// HotIf FunctionObject
{
	if (aExprObj)
	{
		HotkeyCriterion *cp;
		for (cp = g_FirstHotExpr; ; cp = cp->NextExpr)
		{
			if (!cp) // End of the list and it wasn't found.
			{
				auto fr = ValidateFunctor(aExprObj, 1);
				if (fr != OK)
					return fr;
				if (  !(cp = AddHotkeyIfExpr())  )
					return FR_E_OUTOFMEM;
				aExprObj->AddRef();
				cp->Type = HOT_IF_CALLBACK;
				cp->Callback = aExprObj;
				cp->WinTitle = _T("");
				cp->WinText = _T("");
				break;
			}
			if (cp->Type == HOT_IF_CALLBACK && cp->Callback == aExprObj)
				break;
		}
		g->HotCriterion = cp;
	}
	else
	{
		g->HotCriterion = nullptr;
	}
	return OK;
}



FResult Hotkey::IfExpr(LPCTSTR aExpr)
// HotIf ; Set null criterion.
// HotIf "Exact-expression-text"
{
	if (!aExpr || !*aExpr)
	{
		g->HotCriterion = nullptr;
	}
	else
	{
		HotkeyCriterion *cp = FindHotkeyIfExpr(aExpr);
		if (!cp) // Expression not found.
			return FValueError(ERR_HOTKEY_IF_EXPR, aExpr);
		g->HotCriterion = cp;
	}
	return OK;
}



FResult Hotkey::Dynamic(LPCTSTR aHotkeyName, LPCTSTR aOptions, IObject *aCallback, HookActionType aHookAction)
// Creates, updates, enables, or disables a hotkey dynamically (while the script is running).
// Returns OK or FAIL.
{
	// Caller has ensured that aCallback and aHookAction can't both be non-zero.  Furthermore,
	// both can be zero/NULL only when the caller is updating an existing hotkey to have new options
	// (i.e. it's retaining its current callback).
	if (aCallback)
	{
		auto fr = ValidateFunctor(aCallback, 1);
		if (fr != OK)
			return fr;
	}

	UCHAR no_suppress;
	bool hook_is_mandatory;
	Hotkey *hk = FindHotkeyByTrueNature(aHotkeyName, no_suppress, hook_is_mandatory); // NULL if not found.
	HotkeyVariant *variant = hk ? hk->FindVariant() : NULL;
	bool update_all_hotkeys = false;  // This method avoids multiple calls to ManifestAllHotkeysHotstringsHooks() (which is high-overhead).
	bool variant_was_just_created = false;

	switch (aHookAction)
	{
	case HOTKEY_ID_ON:
	case HOTKEY_ID_OFF:
	case HOTKEY_ID_TOGGLE:
		if (!hk)
			return FError(ERR_NONEXISTENT_HOTKEY, aHotkeyName, ErrorPrototype::Target);
		if (!(variant || hk->mHookAction)) // mHookAction (alt-tab) hotkeys don't need a variant that matches the current criteria.
			// To avoid ambiguity and also allow the script to use error handling to detect whether a variant
			// already exists, it seems best to strictly require a matching variant rather than falling back
			// onto some "default variant" such as the global variant (if any).
			return FError(ERR_NONEXISTENT_VARIANT, aHotkeyName, ErrorPrototype::Target);
		if (aHookAction == HOTKEY_ID_TOGGLE)
			aHookAction = hk->mHookAction
				? (hk->mParentEnabled ? HOTKEY_ID_OFF : HOTKEY_ID_ON) // Enable/disable parent hotkey (due to alt-tab being a global hotkey).
				: (variant->mEnabled ? HOTKEY_ID_OFF : HOTKEY_ID_ON); // Enable/disable individual variant.
		if (aHookAction == HOTKEY_ID_ON)
		{
			if (hk->mHookAction ? hk->EnableParent() : hk->Enable(*variant))
				update_all_hotkeys = true; // Do it this way so that any previous "true" value isn't lost.
		}
		else
			if (hk->mHookAction ? hk->DisableParent() : hk->Disable(*variant))
				update_all_hotkeys = true; // Do it this way so that any previous "true" value isn't lost.
		break;

	default: // aHookAction is 0 or an AltTab action.  COMMAND: Hotkey, Name, Callback|AltTabAction
		if (!hk) // No existing hotkey of this name, so create a new hotkey.
		{
			if (aHookAction) // Create hotkey: Hotkey Name, AltTabAction
				hk = AddHotkey(NULL, aHookAction, aHotkeyName, no_suppress);
			else // Create hotkey: Hotkey Name, Callback [, Options]
			{
				if (!aCallback) // Caller is trying to set new aOptions for a nonexistent hotkey.
					return FError(ERR_NONEXISTENT_HOTKEY, aHotkeyName, ErrorPrototype::Target);
				hk = AddHotkey(aCallback, 0, aHotkeyName, no_suppress);
			}
			if (!hk)
				return FAIL; // AddHotkey() already displayed the error.
			variant = hk->mLastVariant; // Update for use with the options-parsing section further below.
			update_all_hotkeys = true;
			variant_was_just_created = true;
		}
		else // Hotkey already exists (though possibly not the required variant).  Update the hotkey if appropriate.
		{
			if (hk->mHookAction != aHookAction) // COMMAND: Change to/from alt-tab hotkey.
			{
				// LoadIncludedFile() contains logic and comments similar to this, so maintain them together.
				// If aHookAction isn't zero, the caller is converting this hotkey into a global alt-tab
				// hotkey (alt-tab hotkeys are never subject to #HotIf, as documented).  Thus, variant can
				// be NULL because making a hotkey become alt-tab doesn't require the creation or existence
				// of a variant matching the current #HotIf criteria.  However, continue on to process the
				// Options parameter in case it contains "On" or some other keyword applicable to alt-tab.
				hk->mHookAction = aHookAction;
				if (!aHookAction)
					// Since this hotkey is going from alt-tab to non-alt-tab, make sure it's not disabled
					// because currently, mParentEnabled is only actually used by alt-tab hotkeys (though it
					// may have other uses in the future, which is why it's implemented and named the way it is).
					hk->mParentEnabled = true;
				else // This hotkey is becoming alt-tab.
				{
					// Make the hook mandatory for this hotkey. Known limitation: For maintainability and code size,
					// this is never undone (even if the hotkey is changed back to non-alt-tab) because there are
					// many other reasons a hotkey's mKeybdHookMandatory could be true, so can't change it back to
					// false without checking all those other things.
					if (HK_TYPE_CAN_BECOME_KEYBD_HOOK(hk->mType))
						hk->mKeybdHookMandatory = true; // Causes mType to be set to HK_KEYBD_HOOK by ManifestAllHotkeysHotstringsHooks().
				}
				// Even if it's still an alt-tab action (just a different one), hook's data structures still
				// need to be updated.  Otherwise, this is a change from an alt-tab type to a non-alt-tab type,
				// or vice versa: Due to the complexity of registered vs. hook hotkeys, for now just start from
				// scratch so that there is high confidence that the hook and all registered hotkeys, including
				// their interdependencies, will be re-initialized correctly.
				update_all_hotkeys = true;
			}
			
			// If the above changed the action from an Alt-tab type to non-alt-tab, there may be a callback
			// to be applied to the existing variant (or created as a new variant).
			if (aCallback) // Update hotkey: Hotkey Name, Callback [, Options]
			{
				// If there's a matching variant, update its callback. Otherwise, create a new variant.
				if (variant) // There's an existing variant...
				{
					if (aCallback != variant->mCallback) // ...and its callback is being changed.
						variant->mCallback = aCallback;
				}
				else // No existing variant matching current criteria, so create a new variant.
				{
					if (   !(variant = hk->AddVariant(aCallback, no_suppress))   ) // Out of memory.
						return FR_E_OUTOFMEM;
					variant_was_just_created = true;
					update_all_hotkeys = true;
					// It seems undesirable for #UseHook to be applied to a hotkey just because it's options
					// were updated with the Hotkey command; therefore, #UseHook is only applied for newly
					// created variants such as this one.  For others, the $ prefix can be applied.
					if (g_ForceKeybdHook)
						hook_is_mandatory = true;
				}
			}
			else
				// NULL callback, so either it just became an alt-tab hotkey above, or it's "Hotkey Name,, Options".
				if (!variant) // Below relies on this check.
					break; // Let the error-catch below report it as an error.

			// v1.1.15: Allow the ~tilde prefix to be added/removed from an existing hotkey variant.
			// v1.1.19: Apply this change even if aCallback is omitted.  This is redundant if
			// variant_was_just_created, but checking that condition seems counter-productive.
			variant->mNoSuppress = no_suppress;
			// hk->mNoSuppress might be inaccurate if a no-suppress flag was just removed from this variant,
			// but that just means a slight reduction in efficiency if tilde is removed from all variants.
			hk->mNoSuppress |= no_suppress; // Apply both AT_LEAST_ONE_VARIANT_HAS_TILDE and NO_SUPPRESS_PREFIX, if present.
			if (!(no_suppress & AT_LEAST_ONE_VARIANT_HAS_TILDE))
				hk->mNoSuppress |= AT_LEAST_ONE_VARIANT_LACKS_TILDE;
				
			// v1.1.19: Allow the $UseHook prefix to be added to an existing hotkey.
			if (!hk->mKeybdHookMandatory && (hook_is_mandatory || no_suppress))
			{
				// Require the hook for all variants of this hotkey if any variant requires it.
				// This seems more intuitive than the old behaviour, which required $ or #UseHook
				// to be used on the *first* variant, even though it affected all variants.
				update_all_hotkeys = true; // Since it may be switching from reg to k-hook.
				hk->mKeybdHookMandatory = true;
			}
		} // Hotkey already existed.
		break;
	} // switch(aHookAction)

	// Above has ensured that hk is not NULL.

	// The following check catches:
	// Hotkey, Name,, Options  ; Where name exists as a hotkey, but the right variant doesn't yet exist.
	// If it catches anything else, that could be a bug, so this error message will help spot it.
	if (!(variant || hk->mHookAction)) // mHookAction (alt-tab) hotkeys don't need a variant that matches the current criteria.
		return FError(ERR_NONEXISTENT_VARIANT, aHotkeyName, ErrorPrototype::Target);
	// Below relies on the fact that either variant or hk->mHookAction (or both) is now non-zero.
	// Specifically, when an existing hotkey was changed to become an alt-tab hotkey, above, there will sometimes
	// be a NULL variant (depending on whether there happens to be a variant in the hotkey that matches the current criteria).

	// If aOptions is blank, any new hotkey or variant created above will have used the current values of
	// g_MaxThreadsBuffer, etc.
	if (*aOptions)
	{
		for (auto cp = aOptions; *cp; ++cp)
		{
			switch(ctoupper(*cp))
			{
			case 'O': // v1.0.38.02.
				if (ctoupper(cp[1]) == 'N') // Full validation for maintainability.
				{
					++cp; // Omit the 'N' from further consideration in case it ever becomes a valid option letter.
					if (hk->mHookAction ? hk->EnableParent() : hk->Enable(*variant)) // Under these conditions, earlier logic has ensured variant is non-NULL.
						update_all_hotkeys = true; // Do it this way so that any previous "true" value isn't lost.
				}
				else if (!_tcsnicmp(cp, _T("Off"), 3))
				{
					cp += 2; // Omit the letters of the word from further consideration in case "f" ever becomes a valid option letter.
					if (hk->mHookAction ? hk->DisableParent() : hk->Disable(*variant)) // Under these conditions, earlier logic has ensured variant is non-NULL.
						update_all_hotkeys = true; // Do it this way so that any previous "true" value isn't lost.
					if (variant_was_just_created) // This variant (and possibly its parent hotkey) was just created above.
						update_all_hotkeys = false; // Override the "true" that was set (either right above *or* anywhere earlier) because this new hotkey/variant won't affect other hotkeys.
				}
				break;
			case 'B':
				if (variant)
					variant->mMaxThreadsBuffer = (cp[1] != '0');  // i.e. if the char is NULL or something other than '0'.
				break;
			// For options such as P & T: Use atoi() vs. ATOI() to avoid interpreting something like 0x01B
			// as hex when in fact the B was meant to be an option letter:
			case 'P':
				if (variant)
					variant->mPriority = _ttoi(cp + 1);
				break;
			case 'S':
				if (variant)
					variant->mSuspendExempt = (cp[1] != '0');
				break;
			case 'T':
				if (variant)
				{
					variant->mMaxThreads = _ttoi(cp + 1);
					if (variant->mMaxThreads > g_MaxThreadsTotal) // To avoid array overflow, this limit must by obeyed except where otherwise documented.
						// Older comment: Keep this limited to prevent stack overflow due to too many pseudo-threads.
						variant->mMaxThreads = g_MaxThreadsTotal;
					else if (variant->mMaxThreads < 1)
						variant->mMaxThreads = 1;
				}
				break;
			case 'I':
				if (variant)
				{
					int new_input_level = _ttoi(cp + 1);
					if (SendLevelIsValid(new_input_level))
					{
						if (new_input_level && !hk->mKeybdHookMandatory)
						{
							// For simplicity, a hotkey requires the hook if any of its variants have a non-zero
							// input level, even if those variants are disabled.  The same is done for the tilde
							// prefix above and in AddVariant(); see there for more comments.
							hk->mKeybdHookMandatory = true;
							update_all_hotkeys = true;
						}
						variant->mInputLevel = (SendLevelType)new_input_level;
					}
				}
				break;
			// Otherwise: Ignore other characters, such as the digits that comprise the number after the T option.
			}
		} // for()
	} // if (*aOptions)
		
	if (update_all_hotkeys)
		ManifestAllHotkeysHotstringsHooks(); // See its comments for why it's done in so many of the above situations.

	return OK;
}



Hotkey *Hotkey::AddHotkey(IObject *aCallback, HookActionType aHookAction, LPCTSTR aName, UCHAR aNoSuppress)
// aCallback can be NULL if the caller is creating a dynamic hotkey that has an aHookAction.
// aName must not be NULL.
// Returns the address of the new hotkey on success, or NULL otherwise.
// The caller is responsible for calling ManifestAllHotkeysHotstringsHooks(), if appropriate.
{
	if (   (shkMax <= sNextID && !HookAdjustMaxHotkeys(shk, shkMax, shkMax ? shkMax * 2 : INITIAL_MAX_HOTKEYS)) // Allocate or expand shk if needed.
		|| !(shk[sNextID] = new Hotkey(sNextID, aCallback, aHookAction, aName, aNoSuppress))   )
	{
		MemoryError();
		return NULL;
	}
	if (!shk[sNextID]->mConstructedOK)
	{
		delete shk[sNextID];  // SimpleHeap allows deletion of most recently added item.
		return NULL;  // The constructor already displayed the error.
	}
	++sNextID;
	return shk[sNextID - 1]; // Indicate success by returning the new hotkey.
}



Hotkey::Hotkey(HotkeyIDType aID, IObject *aCallback, HookActionType aHookAction, LPCTSTR aName
	, UCHAR aNoSuppress)
	: mID(HOTKEY_ID_INVALID)  // Default until overridden.
	// Caller must ensure that either aName or aCallback isn't NULL.
	, mVK(0)
	, mSC(0)
	, mModifiers(0)
	, mModifiersLR(0)
	, mKeybdHookMandatory(false)
	, mAllowExtraModifiers(false)
	, mKeyUp(false)
	, mNoSuppress(0)  // Default is to suppress both prefixes and suffixes.
	, mModifierVK(0)
	, mModifierSC(0)
	, mModifiersConsolidatedLR(0)
	, mType(HK_NORMAL) // Default unless overridden to become mouse, joystick, hook, etc.
	, mVK_WasSpecifiedByNumber(false)
	, mIsRegistered(false)
	, mParentEnabled(true)
	, mHookAction(aHookAction)   // Alt-tab and possibly other uses.
	, mFirstVariant(NULL), mLastVariant(NULL)  // Init linked list early for maintainability.
	, mConstructedOK(false)

// It's better to receive the hotkey_id as a param, since only the caller has better knowledge and
// verification of the fact that this hotkey's id is always set equal to it's index in the array
// (for performance reasons).
{
	if (sNextID > HOTKEY_ID_MAX)
	{
		// This will actually cause the script to terminate if this hotkey is a static (load-time)
		// hotkey.  In the future, some other behavior is probably better:
		g_script.ScriptError(_T("Max hotkeys."));  // Brief msg since so rare.
		return;
	}

	if (!TextInterpret(aName, this)) // The called function already displayed the error.
		return;

	if (mType == HK_JOYSTICK)
	{
		if (mModifiers || mModifierVK || mModifierSC)
		{
			ValueError(ERR_INVALID_HOTKEY, aName, FAIL);
			return;
		}
	}
	else // Perform modifier adjustment and other activities that don't apply to joysticks.
	{
		// Remove any modifiers that are obviously redundant from keys (even NORMAL/registered ones
		// due to cases where RegisterHotkey() fails and the key is then auto-enabled via the hook).
		// No attempt is currently made to correct a silly hotkey such as "lwin & lwin".  In addition,
		// weird hotkeys such as <^Control and ^LControl are not currently validated and might yield
		// unpredictable results.
		bool is_neutral;
		modLR_type modifiers_lr;
		if (modifiers_lr = KeyToModifiersLR(mVK, mSC, &is_neutral))
		{
			// This hotkey's action-key is itself a modifier, so ensure that it's not defined
			// to modify itself.  Other sections might rely on us doing this:
			if (is_neutral)
				// Since the action-key is a neutral modifier (not left or right specific),
				// turn off any neutral modifiers that may be on:
				mModifiers &= ~ConvertModifiersLR(modifiers_lr);
			else
				mModifiersLR &= ~modifiers_lr;
		}

		if (   (mHookAction == HOTKEY_ID_ALT_TAB || mHookAction == HOTKEY_ID_ALT_TAB_SHIFT)
			&& !mModifierVK && !mModifierSC   )
		{
			if (mModifiers)
			{
				// Neutral modifier has been specified.  Future enhancement: improve this
				// to try to guess which key, left or right, should be used based on the
				// location of the suffix key on the keyboard.  Lexikos: Better not do that
				// since a wrong guess will leave the user wondering why it doesn't work.
				ValueError(ERR_ALTTAB_MODLR, aName, FAIL);
				return;  // Key is invalid so don't give it an ID.
			}
			if (mModifiersLR)
			{
				// If mModifiersLR contains only a single modifier key, that is valid
				// so we convert it here to its corresponding mModifierVK for use by
				// the hook.
				switch (mModifiersLR)
				{
				case MOD_LCONTROL: mModifierVK = VK_LCONTROL; break;
				case MOD_RCONTROL: mModifierVK = VK_RCONTROL; break;
				case MOD_LSHIFT: mModifierVK = VK_LSHIFT; break;
				case MOD_RSHIFT: mModifierVK = VK_RSHIFT; break;
				case MOD_LALT: mModifierVK = VK_LMENU; break;
				case MOD_RALT: mModifierVK = VK_RMENU; break;
				case MOD_LWIN: mModifierVK = VK_LWIN; break;
				case MOD_RWIN: mModifierVK = VK_RWIN; break;
				default:
					ValueError(ERR_ALTTAB_ONEMOD, aName, FAIL);
					return;  // Key is invalid so don't give it an ID.
				}
				// Since above didn't return:
				mModifiersLR = 0;  // Since ModifierVK/SC is now its substitute.
			}
			// Update: This is no longer needed because the hook attempts to compensate.
			// However, leaving it enabled may improve performance and reliability.
			// Update#2: No, it needs to be disabled, otherwise alt-tab won't work right
			// in the rare case where an ALT key itself is defined as "AltTabMenu":
			//else
				// It has no ModifierVK/SC and no modifiers, so it's a hotkey that is defined
				// to fire only when the Alt-Tab menu is visible.  Since the ALT key must be
				// down for that menu to be visible (on all OSes?), add the ALT key to this
				// keys modifiers so that it will be detected as a hotkey whenever the
				// Alt-Tab menu is visible:
			//	modifiers |= MOD_ALT;
		}

		if (HK_TYPE_CAN_BECOME_KEYBD_HOOK(mType)) // Added in v1.0.39 to make a hotkey such as "LButton & LCtrl" install the mouse hook.
		{
			switch (mVK)
			{
			case 0: // Scan codes having no available virtual key must always be handled by the hook.
			// In addition, to support preventing the toggleable keys from toggling, handle those
			// with the hook also.
			case VK_NUMLOCK:
			case VK_CAPITAL:
			case VK_SCROLL:
			// When the AppsKey used as a suffix, always use the hook to handle it because registering
			// such keys with RegisterHotkey() will fail to suppress(hide) the key-up events from the system,
			// and the key-up for Apps key, at least in apps like Explorer, is a special event that results in
			// the context menu appearing (though most other apps seem to use the key-down event rather than
			// the key-up, so they would probably work okay).  Note: Of possible future use is the fact that
			// if the Alt key is held down before pressing Appskey, it's native function does
			// not occur.  This may be similar to the fact that LWIN and RWIN don't cause the
			// start menu to appear if a shift key is held down.
			case VK_APPS:
			// Finally, the non-neutral (left-right) modifier keys (except LWin and RWin) must also
			// be done with the hook because even if RegisterHotkey() claims to succeed on them,
			// I'm 99% sure I tried it and the hotkeys don't actually work with that method:
			case VK_LCONTROL:
			case VK_RCONTROL:
			case VK_LSHIFT:
			case VK_RSHIFT:
			case VK_LMENU:
			case VK_RMENU:
				mKeybdHookMandatory = true;
				break;

			// To prevent the Start Menu from appearing for a naked LWIN or RWIN, must
			// handle this key with the hook (the presence of a normal modifier makes
			// this unnecessary, at least under WinXP, because the Start Menu is
			// never invoked when a modifier key is held down with lwin/rwin).
			case VK_LWIN:
			case VK_RWIN:
			// If this hotkey is an unmodified modifier (e.g. Control::) and there
			// are any other hotkeys that rely specifically on this modifier,
			// have the hook handle this hotkey so that it will only fire on key-up
			// rather than key-down.  Note: cases where this key's modifiersLR or
			// ModifierVK/SC are non-zero -- as well as hotkeys that use sc vs. vk
			// -- have already been set to use the keybd hook, so don't need to be
			// handled here.  UPDATE: All the following cases have been already set
			// to be HK_KEYBD_HOOK:
			// - left/right ctrl/alt/shift (since RegisterHotkey() doesn't support them).
			// - Any key with a ModifierVK/SC
			// - The naked lwin or rwin key (due to the check above)
			// Therefore, the only case left to be detected by this next line is the
			// one in which the user configures the naked neutral key VK_SHIFT,
			// VK_MENU, or VK_CONTROL.  As a safety precaution, always handle those
			// neutral keys with the hook so that their action will only fire
			// when the key is released (thus allowing each key to be used for its
			// normal modifying function):
			case VK_CONTROL:
			case VK_MENU:
			case VK_SHIFT:
				if (!mModifiers && !mModifiersLR) // Modifier key as suffix and has no modifiers (or only a ModifierVK/SC).
					mKeybdHookMandatory = true;
				//else keys modified by CTRL/SHIFT/ALT/WIN can always be registered normally because these
				// modifiers are never used (are overridden) when that key is used as a ModifierVK
				// for another key.  Example: if key1 is a ModifierVK for another key, ^key1
				// (or any other modified versions of key1) can be registered as a hotkey because
				// that doesn't affect the hook's ability to use key1 as a prefix:
				break;
			}
		} // if HK_TYPE_CAN_BECOME_KEYBD_HOOK(mType)

		if (HK_TYPE_CAN_BECOME_KEYBD_HOOK(mType))
			if (   (mModifiersLR || aHookAction || mKeyUp || mModifierVK || mModifierSC) // mSC is handled higher above.
				|| (g_ForceKeybdHook || mAllowExtraModifiers // mNoSuppress must be checked later because it can be changed by any variants added after *this* one.
					|| (mVK && !mVK_WasSpecifiedByNumber && vk_to_sc(mVK, true)))   ) // Its mVK corresponds to two scan codes (such as "ENTER").
				mKeybdHookMandatory = true;
			// v1.0.38.02: The check of mVK_WasSpecifiedByNumber above was added so that an explicit VK hotkey such
			// as "VK24::" (which is VK_HOME) can be handled via RegisterHotkey() vs. the hook.  Someone asked for
			// this ability, but even if it weren't for that it seems more correct to recognize an explicitly-specified
			// VK as a "neutral VK" (i.e. one that fires for both scan codes if the VK has two scan codes). The user
			// can always specify "SCnnn::" as a hotkey to avoid this fire-on-both-scan-codes behavior.

		// Currently, these take precedence over each other in the following order, so don't
		// just bitwise-or them together in case there's any ineffectual stuff stored in
		// the fields that have no effect (e.g. modifiers have no effect if there's a mModifierVK):
		if (mModifierVK)
			mModifiersConsolidatedLR = KeyToModifiersLR(mModifierVK);
		else if (mModifierSC)
			mModifiersConsolidatedLR = KeyToModifiersLR(0, mModifierSC);
		else
		{
			mModifiersConsolidatedLR = mModifiersLR;
			if (mModifiers)
				mModifiersConsolidatedLR |= ConvertModifiers(mModifiers);
		}
	} // if (mType != HK_JOYSTICK)

	// If mKeybdHookMandatory==true, ManifestAllHotkeysHotstringsHooks() will set mType to HK_KEYBD_HOOK for us.

	// To avoid memory leak, this is done only when it is certain the hotkey will be created:
	if (   !(mName = SimpleHeap::Malloc(aName))
		|| !(AddVariant(aCallback, aNoSuppress))   ) // Too rare to worry about freeing the other if only one fails.
	{
		MemoryError();
		return;
	}
	// Above has ensured that both mFirstVariant and mLastVariant are non-NULL, so callers can rely on that.

	// Always assign the ID last, right before a successful return, so that the caller is notified
	// that the constructor succeeded:
	mConstructedOK = true;
	mID = aID;
	// Don't do this because the caller still needs the old/unincremented value:
	//++sHotkeyCount;  // Hmm, seems best to do this here, but revisit this sometime.
}



HotkeyVariant *Hotkey::FindVariant()
// Returns the address of the variant in this hotkey whose criterion matches the current #HotIf criterion.
// If no match, it returns NULL.
{
	for (HotkeyVariant *vp = mFirstVariant; vp; vp = vp->mNextVariant)
		if (vp->mHotCriterion == g->HotCriterion)
			return vp;
	return NULL;
}



HotkeyVariant *Hotkey::AddVariant(IObject *aCallback, UCHAR aNoSuppress)
// Returns NULL upon out-of-memory; otherwise, the address of the new variant.
// The caller is responsible for calling ManifestAllHotkeysHotstringsHooks(), if appropriate.
{
	HotkeyVariant *vp = SimpleHeap::Alloc<HotkeyVariant>();
	ZeroMemory(vp, sizeof(HotkeyVariant));
	// The following members are left at 0/NULL by the above:
	// mNextVariant
	// mExistingThreads
	// mRunAgainAfterFinished
	// mRunAgainTime
	// mPriority (default priority is always 0)
	HotkeyVariant &v = *vp;
	v.mCallback = aCallback;
	v.mOriginalCallback = g_script.mLastHotFunc;
	v.mMaxThreads = g_MaxThreadsPerHotkey;    // The values of these can vary during load-time.
	v.mMaxThreadsBuffer = g_MaxThreadsBuffer; //
	v.mInputLevel = g_InputLevel;
	v.mHotCriterion = g->HotCriterion; // If this hotkey is an alt-tab one (mHookAction), this is stored but ignored until/unless the Hotkey command converts it into a non-alt-tab hotkey.
	v.mEnabled = true;
	v.mSuspendExempt = g_SuspendExempt;
	if (v.mInputLevel > 0)
	{
		// A non-zero InputLevel only works when using the hook
		mKeybdHookMandatory = true;
	}
	v.mNoSuppress = aNoSuppress;
	mNoSuppress |= aNoSuppress; // Apply both AT_LEAST_ONE_VARIANT_HAS_TILDE and NO_SUPPRESS_PREFIX, if present.
	if (aNoSuppress & AT_LEAST_ONE_VARIANT_HAS_TILDE)
	{
		// For simplicity, make the hook mandatory for any hotkey that has at least one non-suppressed variant.
		// Otherwise, ManifestAllHotkeysHotstringsHooks() would have to do a loop to check if any
		// non-suppressed variants are actually enabled & non-suspended to decide if the hook is actually needed
		// for a hotkey that has a global variant.  Due to rarity and code size, it doesn't seem worth it.
		mKeybdHookMandatory = true;
	}
	else
		mNoSuppress |= AT_LEAST_ONE_VARIANT_LACKS_TILDE;

	// Update the linked list:
	if (!mFirstVariant)
	{
		vp->mIndex = 1; // Start at 1 since 0 means "undetermined variant".
		mFirstVariant = mLastVariant = vp;
	}
	else
	{
		vp->mIndex = mLastVariant->mIndex + 1;
		mLastVariant->mNextVariant = vp;
		// This must be done after the above:
		mLastVariant = vp;
	}

	return vp;  // Indicate success by returning the new object.
}



ResultType Hotkey::TextInterpret(LPCTSTR aName, Hotkey *aThisHotkey, bool aSyntaxCheckOnly)
// Returns OK or FAIL.  This function is static and aThisHotkey is passed in as a parameter
// so that aThisHotkey can be NULL. NULL signals that aName should be checked as a valid
// hotkey only rather than populating the members of the new hotkey aThisHotkey. This function
// and those it calls should avoid showing any error dialogs in validation mode.  Instead,
// it should simply return OK if aName is a valid hotkey and FAIL otherwise.
{
	// Make a copy that can be modified:
	TCHAR hotkey_name[256];
	tcslcpy(hotkey_name, aName, _countof(hotkey_name));
	LPTSTR term1 = hotkey_name;
	LPTSTR term2 = _tcsstr(term1, COMPOSITE_DELIMITER);
	if (!term2)
		return TextToKey(TextToModifiers(term1, aThisHotkey), false, aThisHotkey, aSyntaxCheckOnly);
	if (*term1 == '~')
		++term1; // Some other stage handles this modifier, so just ignore it here.
    LPTSTR end_of_term1 = omit_trailing_whitespace(term1, term2) + 1;
	// Temporarily terminate the string so that the 2nd term is hidden:
	TCHAR ctemp = *end_of_term1;
	*end_of_term1 = '\0';
	ResultType result = TextToKey(term1, true, aThisHotkey, aSyntaxCheckOnly);
	*end_of_term1 = ctemp;  // Undo the termination.
	if (result == FAIL || result == CONDITION_FALSE)
		return result;
	term2 += COMPOSITE_DELIMITER_LENGTH;
	term2 = omit_leading_whitespace(term2);
	// Even though modifiers on keys already modified by a mModifierVK are not supported, call
	// TextToModifiers() anyway to use its output (for consistency).  The modifiers it sets
	// are currently ignored because the mModifierVK takes precedence.
	// UPDATE: Treat any modifier other than '~' as an error, since otherwise users expect
	// hotkeys like "' & +e::Send È" to work.
	//term2 = TextToModifiers(term2, aThisHotkey);
	if (*term2 == '~')
		++term2; // Some other stage handles this modifier, so just ignore it here.
	ResultType result2 = TextToKey(term2, false, aThisHotkey, aSyntaxCheckOnly);
	return result2 != OK ? result2 : result;
}



LPCTSTR Hotkey::TextToModifiers(LPCTSTR aText, Hotkey *aThisHotkey, HotkeyProperties *aProperties)
// This function and those it calls should avoid showing any error dialogs when caller passes NULL for aThisHotkey.
// Takes input param <text> to support receiving only a subset of object.text.
// Returns the location in <text> of the first non-modifier key.
// Checks only the first char(s) for modifiers in case these characters appear elsewhere (e.g. +{+}).
// But come to think of it, +{+} isn't valid because + itself is already shift-equals.  So += would be
// used instead, e.g. +==action.  Similarly, all the others, except =, would be invalid as hotkeys also.
// UPDATE: On some keyboard layouts, the + key and various others don't require the shift key to be
// manifest.  Thus, on these systems a hotkey such as ^+:: is now supported as meaning Ctrl-Plus.
{
	// Init output parameter for caller if it gave one:
	if (aProperties)
		ZeroMemory(aProperties, sizeof(HotkeyProperties));

	if (!*aText)
		return aText; // Below relies on this having ensured that aText isn't blank.

	// Explicitly avoids initializing modifiers to 0 because the caller may have already included
	// some set some modifiers in there.
	LPCTSTR marker;
	bool key_left, key_right;

	// Simplifies and reduces code size below:
	mod_type temp_modifiers;
	mod_type &modifiers = aProperties ? aProperties->modifiers : (aThisHotkey ? aThisHotkey->mModifiers : temp_modifiers);
	modLR_type temp_modifiersLR;
	modLR_type &modifiersLR = aProperties ? aProperties->modifiersLR: (aThisHotkey ? aThisHotkey->mModifiersLR : temp_modifiersLR);

	// Improved for v1.0.37.03: The loop's condition is now marker[1] vs. marker[0] so that
	// the last character is never considered a modifier.  This allows a modifier symbol
	// to double as the name of a suffix key.  It also fixes issues on layouts where the
	// symbols +^#! do not require the shift key to be held down, such as the German layout.
	//
	// Improved for v1.0.40.01: The loop's condition now stops when it reaches a single space followed
	// by the word "Up" so that hotkeys like "< up" and "+ up" are supported by seeing their '<' or '+' as
	// a key name rather than a modifier symbol.
	// Fix for v1.1.27.05: Stop at any space, not just " up", so that " & " is also covered.
	// This fixes "> & v" to not interpret ">" as a modifier.  This also causes "+ ::" to be
	// seen as invalid, where previously TextToModifiers() saw it as Shift+Space but a later
	// stage trimmed the space and registered "+::".  This is best since trailing spaces are
	// not allowed in any other hotkeys, and even "+  ::" (two spaces) was not allowed.
	for (marker = aText, key_left = false, key_right = false; marker[1] && marker[1] != ' '; ++marker)
	{
		switch (*marker)
		{
		case '>':
			key_right = true;
			break;
		case '<':
			key_left = true;
			break;
		case '*':
			if (aThisHotkey)
				aThisHotkey->mAllowExtraModifiers = true;
			if (aProperties)
				aProperties->has_asterisk = true;
			break;
		case '~':
			if (aProperties)
				aProperties->suffix_has_tilde = true; // If this is the prefix's tilde rather than the suffix, it will be overridden later below.
			break;
		case '$':
			if (aThisHotkey)
				aThisHotkey->mKeybdHookMandatory = true; // This flag will be ignored if TextToKey() decides this is a JOYSTICK or MOUSE hotkey.
			if (aProperties)
				aProperties->hook_is_mandatory = true;
			break;
		case '!':
			if ((!key_right && !key_left))
			{
				modifiers |= MOD_ALT;
				break;
			}
			// Both left and right may be specified, e.g. ><+a means both shift keys must be held down:
			if (key_left)
			{
				modifiersLR |= MOD_LALT;
				key_left = false;
			}
			if (key_right)
			{
				modifiersLR |= MOD_RALT;
				key_right = false;
			}
			break;
		case '^':
			if ((!key_right && !key_left))
			{
				modifiers |= MOD_CONTROL;
				break;
			}
			if (key_left)
			{
				modifiersLR |= MOD_LCONTROL;
				key_left = false;
			}
			if (key_right)
			{
				modifiersLR |= MOD_RCONTROL;
				key_right = false;
			}
			break;
		case '+':
			if ((!key_right && !key_left))
			{
				modifiers |= MOD_SHIFT;
				break;
			}
			if (key_left)
			{
				modifiersLR |= MOD_LSHIFT;
				key_left = false;
			}
			if (key_right)
			{
				modifiersLR |= MOD_RSHIFT;
				key_right = false;
			}
			break;
		case '#':
			if ((!key_right && !key_left))
			{
				modifiers |= MOD_WIN;
				break;
			}
			if (key_left)
			{
				modifiersLR |= MOD_LWIN;
				key_left = false;
			}
			if (key_right)
			{
				modifiersLR |= MOD_RWIN;
				key_right = false;
			}
			break;
		default:
			goto break_loop; // Stop immediately whenever a non-modifying char is found.
		} // switch (*marker)
	} // for()
break_loop:

	// Now *marker is the start of the key's name.  In addition, one of the following is now true:
	// 1) marker[0] is a non-modifier symbol; that is, the loop stopped because it found the first non-modifier symbol.
	// 2) marker[1] is '\0'; that is, the loop stopped because it reached the next-to-last char (the last char itself is never a modifier; e.g. ^+ is Ctrl+Plus on some keyboard layouts).
	// 3) marker[1] is the start of the string " Up", in which case marker[0] is considered the suffix key even if it happens to be a modifier symbol (see comments at for-loop's control stmt).
	if (aProperties)
	{
		// When caller passes non-NULL aProperties, it didn't omit the prefix portion of a composite hotkey
		// (e.g. the "a & " part of "a & b" is present).  So parse these and all other types of hotkeys when in this mode.
		if (auto composite = _tcsstr(marker, COMPOSITE_DELIMITER))
		{
			tcslcpy(aProperties->prefix_text, marker, _countof(aProperties->prefix_text)); // Protect against overflow case script ultra-long (and thus invalid) key name.
			if (auto temp = _tcsstr(aProperties->prefix_text, COMPOSITE_DELIMITER)) // Check again in case it tried to overflow.
				omit_trailing_whitespace(aProperties->prefix_text, temp)[1] = '\0'; // Truncate prefix_text so that the suffix text is omitted.
			composite = omit_leading_whitespace(composite + COMPOSITE_DELIMITER_LENGTH);
			aProperties->prefix_has_tilde = aProperties->suffix_has_tilde;
			if (aProperties->suffix_has_tilde = (*composite == '~')) // Override any value of no_suppress set higher above.
				++composite; // For simplicity, no skipping of leading whitespace between tilde and the suffix key name.
			tcslcpy(aProperties->suffix_text, composite, _countof(aProperties->suffix_text)); // Protect against overflow case script ultra-long (and thus invalid) key name.
		}
		else // A normal (non-composite) hotkey, so no_suppress was already set properly (higher above).
			tcslcpy(aProperties->suffix_text, omit_leading_whitespace(marker), _countof(aProperties->suffix_text)); // Protect against overflow case script ultra-long (and thus invalid) key name.
		if (auto temp = tcscasestr(aProperties->suffix_text, _T(" Up"))) // Should be reliable detection method because leading spaces have been omitted and it's unlikely a legitimate key name will ever contain a space followed by "Up".
		{
			omit_trailing_whitespace(aProperties->suffix_text, temp)[1] = '\0'; // Omit " Up" from suffix_text since caller wants that.
			aProperties->is_key_up = true; // Override the default set earlier.
		}
	}
	return marker;
}



ResultType Hotkey::TextToKey(LPCTSTR aText, bool aIsModifier, Hotkey *aThisHotkey, bool aSyntaxCheckOnly)
// This function and those it calls should avoid showing any error dialogs when caller passes NULL for
// aThisHotkey (however, there is at least one exception explained in comments below where it occurs).
// Caller must ensure that aText is a modifiable string.
// Takes input param aText to support receiving only a subset of mName.
// In private members, sets the values of vk/sc or ModifierVK/ModifierSC depending on aIsModifier.
// It may also merge new modifiers into the existing value of modifiers, so the caller
// should never reset modifiers after calling this.
// Returns OK or FAIL.
{
	vk_type temp_vk; // No need to initialize this one.
	sc_type temp_sc = 0;
	modLR_type modifiersLR = 0;
	bool is_mouse = false;
	int joystick_id;

	TCHAR buf[32]; // Large enough that truncation wouldn't yield a valid key name.

	// Previous steps should make it unnecessary to call omit_leading_whitespace(aText).
	auto keyname_end = find_identifier_end(aText);
	if (keyname_end == aText && *aText) // Any single character except '\0' can be a key name.
		++keyname_end;

	if (!aIsModifier && IS_SPACE_OR_TAB(*keyname_end) && !_tcsicmp(omit_leading_whitespace(keyname_end), _T("Up")))
	{
		if (aSyntaxCheckOnly)
			return OK; // It's a word or single character followed by " up" -- looks valid.
		// This is a key-up hotkey, such as "Ctrl Up::".
		if (aThisHotkey)
			aThisHotkey->mKeyUp = true;
		tcslcpy(buf, aText, _countof(buf)); // Make a null-terminated copy to simplify various checks below.
		if (keyname_end - aText < _countof(buf))
			buf[keyname_end - aText] = '\0'; // Terminate at the first space so that the word "up" is removed from further consideration.
		aText = buf;
	}
	else
	{
		// If there's something after the first word/character, it's not a hotkey.
		if (aSyntaxCheckOnly)
			return *keyname_end ? FAIL : OK;
		if (*keyname_end)
			return CONDITION_FALSE;
	}

	HotkeyTypeType hotkey_type_temp;
	HotkeyTypeType &hotkey_type = aThisHotkey ? aThisHotkey->mType : hotkey_type_temp; // Simplifies and reduces code size below.

	if (temp_vk = TextToVK(aText, &modifiersLR, true)) // Assign.
	{
		if (aIsModifier)
		{
			if (IS_WHEEL_VK(temp_vk))
				return ValueError(ERR_UNSUPPORTED_PREFIX, aText, FAIL);
		}
		else
			// This is done here rather than at some later stage because we have access to the raw
			// name of the suffix key (with any leading modifiers such as ^ omitted from the beginning):
			if (aThisHotkey)
				aThisHotkey->mVK_WasSpecifiedByNumber = !_tcsnicmp(aText, _T("VK"), 2);
		is_mouse = IsMouseVK(temp_vk);
		if (modifiersLR & (MOD_LSHIFT | MOD_RSHIFT))
			if (temp_vk >= 'A' && temp_vk <= 'Z')  // VK of an alpha char is the same as the ASCII code of its uppercase version.
				modifiersLR &= ~(MOD_LSHIFT | MOD_RSHIFT);
				// Above: Making alpha chars case insensitive seems much more friendly.  In other words,
				// if the user defines ^Z as a hotkey, it will really be ^z, not ^+z.  By removing SHIFT
				// from the modifiers here, we're only removing it from our modifiers, not the global
				// modifiers that have already been set elsewhere for this key (e.g. +Z will still be +z).
	}
	else // No virtual key was found.  Is there a scan code?
		if (   !(temp_sc = TextToSC(aText))   )
			if (   !(temp_sc = (sc_type)ConvertJoy(aText, &joystick_id, true))   )  // Is there a joystick control/button?
			{
				if (!aText[1] && !g_script.mIsReadyToExecute)
				{
					// At load time, single-character key names are always considered valid but show a
					// warning if they can't be registered as hotkeys on the current keyboard layout.
					if (!aThisHotkey) // First stage: caller wants to differentiate this case from others.
						return CONDITION_TRUE; // It's valid, but won't be active.
					return FAIL; // Second stage: return FAIL to avoid creating an invalid hotkey.
				}
				// It's more appropriate to say "key name" than "hotkey" in this message because it's only
				// showing the one bad key name when it's a composite hotkey such as "Capslock & y".
				return ValueError(ERR_INVALID_KEYNAME, aText, FAIL);
			}
			else
			{
				// Block joystick buttons as prefix keys at this stage in case hotkey_type would be overridden
				// by the suffix key.  For example, the hotkey `Joy1 & LButton::` would reinterpret Joy1 as sc0C.
				if (aIsModifier)
					return ValueError(ERR_UNSUPPORTED_PREFIX, aText, FAIL);

				++sJoyHotkeyCount;
				hotkey_type = HK_JOYSTICK;
				temp_vk = (vk_type)joystick_id;  // 0 is the 1st joystick, 1 the 2nd, etc.
				sJoystickHasHotkeys[joystick_id] = true;
			}


/*
If ever do this, be sure to avoid doing it for keys that must be tracked by scan code (e.g. those in the
scan code array).
	if (!temp_vk && !is_mouse)  // sc must be non-zero or else it would have already returned above.
		if (temp_vk = sc_to_vk(temp_sc))
		{
			snprintf(error_text, sizeof(error_text), "DEBUG: \"%s\" (scan code %X) was successfully mapped to virtual key %X", text, temp_sc, temp_vk);
			MsgBox(error_text);
			temp_sc = 0; // Maybe set this just for safety, even though a non-zero vk should always take precedence over it.
		}
*/
	if (is_mouse)
		hotkey_type = HK_MOUSE_HOOK;

	if (aIsModifier)
	{
		if (aThisHotkey)
		{
			aThisHotkey->mModifierVK = temp_vk;
			aThisHotkey->mModifierSC = temp_sc;
		}
	}
	else
	{
		if (aThisHotkey)
		{
			aThisHotkey->mVK = temp_vk;
            aThisHotkey->mSC = temp_sc;
			// Turn on any additional modifiers.  e.g. SHIFT to realize '#'.
			// Fix for v1.0.37.03: To avoid using the keyboard hook for something like "+::", which in
			// turn would allow the hotkey fire only for LShift+Equals rather than RShift+Equals, convert
			// modifiers from left-right to neutral.  But exclude right-side modifiers (except RWin) so that
			// things like AltGr are more precisely handled (the implications of this policy could use
			// further review).  Currently, right-Alt (via AltGr) is the only possible right-side key.
			aThisHotkey->mModifiers |= ConvertModifiersLR(modifiersLR & (MOD_RWIN|MOD_LWIN|MOD_LCONTROL|MOD_LALT|MOD_LSHIFT));
			aThisHotkey->mModifiersLR |= (modifiersLR & (MOD_RSHIFT|MOD_RALT|MOD_RCONTROL)); // Not MOD_RWIN since it belongs above.
		}
	}
	return OK;
}



ResultType Hotkey::Register()
// Returns OK or FAIL.  Caller is responsible for having checked whether the hotkey is suspended or disabled.
{
	if (mIsRegistered)
		return OK;  // It's normal for a disabled hotkey to return OK.

	if (mType != HK_NORMAL) // Caller normally checks this for performance, but it's checked again for maintainability.
		return FAIL; // Don't attempt to register joystick or hook hotkeys, since their VK/modifiers aren't intended for that.

	// Indicate that the key modifies itself because RegisterHotkey() requires that +SHIFT,
	// for example, be used to register the naked SHIFT key.  So what we do here saves the
	// user from having to specify +SHIFT in the script:
	mod_type modifiers_to_register = mModifiers;
	switch (mVK)
	{
	case VK_LWIN:
	case VK_RWIN: modifiers_to_register |= MOD_WIN; break;
	case VK_CONTROL: modifiers_to_register |= MOD_CONTROL; break;
	case VK_SHIFT: modifiers_to_register |= MOD_SHIFT; break;
	case VK_MENU: modifiers_to_register |= MOD_ALT; break;
	}

	// Must register them to our main window (i.e. don't use NULL to indicate our thread),
	// otherwise any modal dialogs, such as MessageBox(), that call DispatchMessage()
	// internally wouldn't be able to find anyone to send hotkey messages to, so they
	// would probably be lost:
	if (mIsRegistered = RegisterHotKey(g_hWnd, mID, modifiers_to_register, mVK))
		return OK;
	return FAIL;
	// Above: On failure, reset the modifiers in case this function changed them.  This is
	// done in case this hotkey will now be handled by the hook, which doesn't want any
	// extra modifiers that were added above.
	// UPDATE: In v1.0.42, the mModifiers value is never changed here because a hotkey that
	// gets registered here one time might fail to be registered some other time (perhaps due
	// to Suspend, followed by some other app taking ownership that hotkey, followed by
	// de-suspend, which would then have a hotkey with the wrong modifiers in it for the hook.
}



ResultType Hotkey::Unregister()
// Returns OK or FAIL.
{
	if (!mIsRegistered)
		return OK;
	// Don't report any errors in here, at least not when we were called in conjunction
	// with cleanup and exit.  Such reporting might cause an infinite loop, leading to
	// a stack overflow if the reporting itself encounters an error and tries to exit,
	// which in turn would call us again:
	if (mIsRegistered = !UnregisterHotKey(g_hWnd, mID))  // I've see it fail in one rare case.
		return FAIL;
	return OK;
}



void Hotkey::InstallKeybdHook()
// Caller must ensure that sWhichHookNeeded and sWhichHookAlways contain HOOK_MOUSE, if appropriate.
// Generally, this is achieved by having ensured that ManifestAllHotkeysHotstringsHooks() was called at least
// once since the last time any hotstrings or hotkeys were disabled/enabled/changed, but in any case at least
// once since the program started.
{
	sWhichHookNeeded |= HOOK_KEYBD;
	if (!g_KeybdHook)
		ChangeHookState(shk, sHotkeyCount, sWhichHookNeeded, sWhichHookAlways);
}



void Hotkey::InstallMouseHook()
// Same comment as for InstallKeybdHook() above.
{
	sWhichHookNeeded |= HOOK_MOUSE;
	if (!g_MouseHook)
		ChangeHookState(shk, sHotkeyCount, sWhichHookNeeded, sWhichHookAlways);
}



Hotkey *Hotkey::FindHotkeyByTrueNature(LPCTSTR aName, UCHAR &aNoSuppress, bool &aHookIsMandatory)
// Returns the address of the hotkey if found, NULL otherwise.
// In v1.0.42, it tries harder to find a match so that the order of modifier symbols doesn't affect the true nature of a hotkey.
// For example, ^!c should be the same as !^c, primarily because RegisterHotkey() and the hook would consider them the same.
// Primary benefits to the above:
// 1) Catches script bugs, such as unintended duplicates.
// 2) Allows a script to use the Hotkey command more precisely and with greater functionality.
// 3) Allows hotkey variants to work correctly even when the order of modifiers varies.  For example, if ^!c is a hotkey that fires
//    only when window type 1 is active and !^c (reversed modifiers) is a hotkey that fires only when window type 2 is active,
//    one of them would never fire because the hook isn't capable or storing two hotkey IDs for the same combination of
//    modifiers+VK/SC.
{
	HotkeyProperties prop_candidate, prop_existing;
	TextToModifiers(aName, NULL, &prop_candidate);
	aNoSuppress = (prop_candidate.prefix_has_tilde ? NO_SUPPRESS_PREFIX : 0) // Set for caller.
				| (prop_candidate.suffix_has_tilde ? AT_LEAST_ONE_VARIANT_HAS_TILDE : 0);
	aHookIsMandatory = prop_candidate.hook_is_mandatory; // Set for caller.
	// Both suffix_has_tilde and prefix_has_tilde are ignored during dupe-checking below.
	// See comments inside the loop for details.

	for (int i = 0; i < sHotkeyCount; ++i)
	{
		TextToModifiers(shk[i]->mName, NULL, &prop_existing);
		if (   prop_existing.modifiers == prop_candidate.modifiers
			&& prop_existing.modifiersLR == prop_candidate.modifiersLR
			&& prop_existing.is_key_up == prop_candidate.is_key_up
			// Treat wildcard (*) as an entirely separate hotkey from one without a wildcard.  This is because
			// the hook has special handling for wildcards that allow non-wildcard hotkeys that overlap them to
			// take precedence, sort of like "clip children".  The logic that builds the eclipsing array would
			// need to be redesigned, which might not even be possible given the complexity of interactions
			// between variant-precedence and hotkey/wildcard-precedence.
			// By contrast, in v1.0.44 pass-through (~) is considered an attribute of each variant of a
			// particular hotkey, not something that makes an entirely new hotkey.
			// This was done because the old method of having them distinct appears to have only one advantage:
			// the ability to dynamically enable/disable ~x separately from x (since if both were in effect
			// simultaneously, one would override the other due to two different hotkey IDs competing for the same
			// ID slot within the VK/SC hook arrays).  The advantages of allowing tilde to be a per-variant attribute
			// seem substantial, namely to have some variant/siblings pass-through while others do not.
			&& prop_existing.has_asterisk == prop_candidate.has_asterisk
			// v1.0.43.05: Use stricmp not lstrcmpi because an uppercase high ANSI letter isn't necessarily
			// produced by holding down the shift key and pressing the lowercase letter.  In addition, it
			// preserves backward compatibility and may improve flexibility.
			&& !_tcsicmp(prop_existing.prefix_text, prop_candidate.prefix_text)
			&& !_tcsicmp(prop_existing.suffix_text, prop_candidate.suffix_text)   )
			return shk[i]; // Match found.
	}

	return NULL;  // No match found.
}



Hotkey *Hotkey::FindHotkeyContainingModLR(modLR_type aModifiersLR) // , HotkeyIDType hotkey_id_to_omit)
// Returns the address of the hotkey if found, NULL otherwise.
// Find the first hotkey whose modifiersLR contains *any* of the modifiers shows in the parameter value.
// Obsolete: The caller tells us the ID of the hotkey to omit from the search because that one
// would always be found (since something like "lcontrol=calc.exe" in the script
// would really be defines as  "<^control=calc.exe".
// Note: By intent, this function does not find hotkeys whose normal/neutral modifiers
// contain <modifiersLR>.
{
	if (!aModifiersLR)
		return NULL;
	for (int i = 0; i < sHotkeyCount; ++i)
		// Bitwise set-intersection: indicates if anything in common:
		if (shk[i]->mModifiersLR & aModifiersLR)
		//if (i != hotkey_id_to_omit && shk[i]->mModifiersLR & modifiersLR)
			return shk[i];
	return NULL;  // No match found.
}


//Hotkey *Hotkey::FindHotkeyWithThisModifier(vk_type aVK, sc_type aSC)
//// Returns the address of the hotkey if found, NULL otherwise.
//// Answers the question: What is the first hotkey with mModifierVK or mModifierSC equal to those given?
//// A non-zero vk param will take precedence over any non-zero value for sc.
//{
//	if (!aVK & !aSC)
//		return NULL;
//	for (int i = 0; i < sHotkeyCount; ++i)
//		if (   (aVK && aVK == shk[i]->mModifierVK) || (aSC && aSC == shk[i]->mModifierSC)   )
//			return shk[i];
//	return NULL;  // No match found.
//}
//
//
//
//Hotkey *Hotkey::FindHotkeyBySC(sc2_type aSC2, mod_type aModifiers, modLR_type aModifiersLR)
//// Returns the address of the hotkey if found, NULL otherwise.
//// Answers the question: What is the first hotkey with the given sc & modifiers *regardless* of
//// any non-zero mModifierVK or mModifierSC it may have?  The mModifierSC/vk is ignored because
//// the caller wants to know whether this key would be blocked if its counterpart were registered.
//// For example, the hook wouldn't see "MEDIA_STOP & NumpadENTER" at all if NumPadENTER was
//// already registered via RegisterHotkey(), since RegisterHotkey() doesn't honor any modifiers
//// other than the standard ones.
//{
//	for (int i = 0; i < sHotkeyCount; ++i)
//		if (!shk[i]->mVK && (shk[i]->mSC == aSC2.a || shk[i]->mSC == aSC2.b))
//			if (shk[i]->mModifiers == aModifiers && shk[i]->mModifiersLR == aModifiersLR)  // Ensures an exact match.
//				return shk[i];
//	return NULL;  // No match found.
//}



LPTSTR Hotkey::ListHotkeys(LPTSTR aBuf, int aBufSize)
// Translates this script's list of variables into text equivalent, putting the result
// into aBuf and returning the position in aBuf of its new string terminator.
{
	LPTSTR aBuf_orig = aBuf;
	// Save vertical space by limiting newlines here:
	aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("Type\tOff?\tLevel\tRunning\tName\r\n")
							 _T("-------------------------------------------------------------------\r\n"));
	// Start at the oldest and continue up through the newest:
	for (int i = 0; i < sHotkeyCount; ++i)
		aBuf = shk[i]->ToText(aBuf, BUF_SPACE_REMAINING, true);
	return aBuf;
}



LPTSTR Hotkey::ToText(LPTSTR aBuf, int aBufSize, bool aAppendNewline)
// aBufSize is an int so that any negative values passed in from caller are not lost.
// Caller has ensured that aBuf isn't NULL.
// Translates this var into its text equivalent, putting the result into aBuf and
// returning the position in aBuf of its new string terminator.
{
	LPTSTR aBuf_orig = aBuf;

	HotkeyVariant *vp;
	int existing_threads;
	for (existing_threads = 0, vp = mFirstVariant; vp; vp = vp->mNextVariant)
		existing_threads += vp->mExistingThreads;

	TCHAR existing_threads_str[128];
	if (existing_threads)
		_itot(existing_threads, existing_threads_str, 10);
	else
		*existing_threads_str = '\0'; // Make it blank to avoid clutter in the hotkey display.

	TCHAR htype[32];
	switch (mType)
	{
	case HK_NORMAL: _tcscpy(htype, _T("reg")); break;
	case HK_KEYBD_HOOK: _tcscpy(htype, _T("k-hook")); break;
	case HK_MOUSE_HOOK: _tcscpy(htype, _T("m-hook")); break;
	case HK_BOTH_HOOKS: _tcscpy(htype, _T("2-hooks")); break;
	case HK_JOYSTICK: _tcscpy(htype, _T("joypoll")); break;
	default: *htype = '\0'; // For maintainability; should never happen.
	}

	LPTSTR enabled_str;
	if (IsCompletelyDisabled()) // Takes into account alt-tab vs. non-alt-tab, etc.
		enabled_str = _T("OFF");
	else if (mHookAction && mParentEnabled) // It's completely "on" in this case.
		enabled_str = _T("");
	else // It's on or if all individual variants are on, otherwise it's partial.
	{
		// Set default: Empty string means "ON" because it reduces clutter in the displayed list.
		for (enabled_str = _T(""), vp = mFirstVariant; vp; vp = vp->mNextVariant)
			if (!vp->mEnabled)
			{
				enabled_str = _T("PART");
				break;
			}
	}

	TCHAR level_str[7]; // Room for "99-100".
	int min_level = 100, max_level = -1;
	for (vp = mFirstVariant; vp; vp = vp->mNextVariant)
	{
		if (min_level > vp->mInputLevel)
			min_level = vp->mInputLevel;
		if (max_level < vp->mInputLevel)
			max_level = vp->mInputLevel;
	}
	if (min_level != max_level)
		_stprintf(level_str, _T("%i-%i"), min_level, max_level);
	else if (min_level)
		ITOA(min_level, level_str);
	else // Show nothing for level 0.
		*level_str = '\0';

	aBuf += sntprintf(aBuf, BUF_SPACE_REMAINING, _T("%s%s\t%s\t%s\t%s\t%s")
		, htype, (mType == HK_NORMAL && !mIsRegistered) ? _T("(no)") : _T("")
		, enabled_str
		, level_str
		, existing_threads_str
		, mName);
	if (aAppendNewline && BUF_SPACE_REMAINING >= 2)
	{
		*aBuf++ = '\r';
		*aBuf++ = '\n';
		*aBuf = '\0';
	}
	return aBuf;
}


///////////////
// Hot Strings
///////////////

// Init static variables:
Hotstring **Hotstring::shs = NULL;
HotstringIDType Hotstring::sHotstringCount = 0;
HotstringIDType Hotstring::sHotstringCountMax = 0;
UINT Hotstring::sEnabledCount = 0;


void Hotstring::SuspendAll(bool aSuspend)
{
	if (sHotstringCount < 1) // At least one part below relies on this check.
		return;

	UINT u;
	if (aSuspend) // Suspend all those that aren't exempt.
	{
		// Recalculating sEnabledCount might perform better in the average case since most aren't exempt.
		for (sEnabledCount = 0, u = 0; u < sHotstringCount; ++u)
			if (shs[u]->mSuspendExempt)
			{
				shs[u]->mSuspended &= ~HS_SUSPENDED;
				if (!shs[u]->mSuspended) // Not turned off.
					++sEnabledCount;
			}
			else
				shs[u]->mSuspended |= HS_SUSPENDED;
	}
	else // Unsuspend all.
	{
		UINT previous_count = sEnabledCount;
		// Recalculating sEnabledCount is probably best since we otherwise need to both remove HS_SUSPENDED
		// and determine if the final suspension status has changed (i.e. no other bits were set).
		for (sEnabledCount = 0, u = 0; u < sHotstringCount; ++u)
		{
			shs[u]->mSuspended &= ~HS_SUSPENDED;
			if (!shs[u]->mSuspended) // Not turned off.
				++sEnabledCount;
		}
		// v1.0.44.08: Added the following section.  Also, the HS buffer is reset, but only when hotstrings
		// are newly enabled after having been entirely disabled.  This is because CollectInput() would not
		// have been called in a long time, making the contents of g_HSBuf obsolete, which in turn might
		// otherwise cause accidental firings based on old keystrokes coupled with new ones.
		if (previous_count == 0 && sEnabledCount > 0)
		{
			*g_HSBuf = '\0';
			g_HSBufLength = 0;
		}
	}
}



ResultType Hotstring::PerformInNewThreadMadeByCaller()
// Returns OK or FAIL.  Caller has already ensured that the backspacing (if specified by mDoBackspace)
// has been done.  Caller must have already created a new thread for us, and must close the thread when
// we return.
{
	if (mExistingThreads >= mMaxThreads)
		return FAIL;
	// See Hotkey::Perform() for details about this.  For hot strings -- which also use the
	// g_script.mThisHotkeyStartTime value to determine whether g_script.mThisHotkeyModifiersLR
	// is still timely/accurate -- it seems best to set to "no modifiers":
	g_script.mThisHotkeyModifiersLR = 0;
	++mExistingThreads;  // This is the thread count for this particular hotstring only.
	
	ResultType result;
	ExprTokenType params = { mName };
	result = mCallback->ExecuteInNewThread(g_script.mThisHotkeyName, &params, 1);
	
	--mExistingThreads;
	return result ? OK : FAIL;	// Return OK on all non-failure results.
}



void Hotstring::DoReplace(LPARAM alParam)
// alParam contains details about how the hotstring was triggered.
{
	global_struct &g = *::g; // Reduces code size and may improve performance.
	// The below buffer allows room for the longest replacement text plus MAX_HOTSTRING_LENGTH for the
	// optional backspaces, +10 for the possible presence of {Raw} and a safety margin.
	TCHAR SendBuf[LINE_SIZE + MAX_HOTSTRING_LENGTH + 10];
	*SendBuf = '\0';
	LPTSTR start_of_replacement = SendBuf;  // Set default.

	if (mDoBackspace)
	{
		int backspace_count = mStringLength;
#ifdef UNICODE
		for (LPCTSTR cp = mString; *cp; ++cp)
			if (IS_SURROGATE_PAIR(cp[0], cp[1]))
				++cp, --backspace_count; // Treat this surrogate pair as a single character (which it is).
#endif
		// Subtract 1 from backspaces because the final key pressed by the user to make a
		// match was already suppressed by the hook (it wasn't sent through to the active
		// window).  So what we do is backspace over all the other keys prior to that one,
		// put in the replacement text (if applicable), then send the EndChar through
		// (if applicable) to complete the sequence.
		if (!mEndCharRequired)
			--backspace_count;
		for (int i = 0; i < backspace_count; ++i)
			*start_of_replacement++ = '\b';  // Use raw backspaces, not {BS n}, in case the send will be raw.
		*start_of_replacement = '\0'; // Terminate the string created above.
	}

	if (mReplacement)
	{
		_tcscpy(start_of_replacement, mReplacement);
		CaseConformModes case_conform_mode = (CaseConformModes)HIWORD(alParam);
		if (case_conform_mode == CASE_CONFORM_ALL_CAPS)
			CharUpper(start_of_replacement);
		else if (case_conform_mode == CASE_CONFORM_FIRST_CAP)
			*start_of_replacement = ltoupper(*start_of_replacement);
		if (!mOmitEndChar) // The ending character (if present) needs to be sent too.
		{
			// Send the final character in raw mode so that chars such as !{} are sent properly.
			// v1.0.43: Avoid two separate calls to SendKeys because:
			// 1) It defeats the uninterruptibility of the hotstring's replacement by allowing the user's
			//    buffered keystrokes to take effect in between the two calls to SendKeys.
			// 2) Performance: Avoids having to install the playback hook twice, etc.
			TCHAR end_char;
			if (mEndCharRequired && (end_char = (TCHAR)LOWORD(alParam))) // Must now check mEndCharRequired because LOWORD has been overloaded with context-sensitive meanings.
			{
				LPTSTR end = start_of_replacement + _tcslen(start_of_replacement);
				// v1.0.43.02: Don't send "{Raw}" if already in raw mode!
				// v1.1.27: Avoid adding {Raw} if it gets switched on within the replacement text.
				if (mSendRaw || tcscasestr(start_of_replacement, _T("{Raw}")) || tcscasestr(start_of_replacement, _T("{Text}")))
					*end++ = end_char, *end = '\0';
				else
					_stprintf(end, _T("%s%c"), _T("{Raw}"), end_char);
			}
		}
	}

	if (!*SendBuf) // No keys to send.
		return;

	// For the following, mSendMode isn't checked because the backup/restore is needed to varying extents
	// by every mode.
	int old_delay = g.KeyDelay;
	int old_press_duration = g.PressDuration;
	int old_delay_play = g.KeyDelayPlay;
	int old_press_duration_play = g.PressDurationPlay;
	SendLevelType old_send_level = g.SendLevel;

	g.KeyDelay = mKeyDelay; // This is relatively safe since SendKeys() normally can't be interrupted by a new thread.
	g.PressDuration = -1;   // Always -1, since Send command can be used in body of hotstring to have a custom press duration.
	g.KeyDelayPlay = -1;
	g.PressDurationPlay = mKeyDelay; // Seems likely to be more useful (such as in games) to apply mKeyDelay to press duration rather than above.
	// Setting the SendLevel to 0 rather than this->mInputLevel since auto-replace hotstrings are used for text replacement rather than
	// key remapping, which means the user almost always won't want the generated input to trigger other hotkeys or hotstrings.
	// Action hotstrings (not using auto-replace) do get their thread's SendLevel initialized to the hotstring's InputLevel.
	g.SendLevel = 0;

	// v1.0.43: The following section gives time for the hook to pass the final keystroke of the hotstring to the
	// system.  This is necessary only for modes other than the original/SendEvent mode because that one takes
	// advantage of the serialized nature of the keyboard hook to ensure the user's final character always appears
	// on screen before the replacement text can appear.
	// By contrast, when the mode is SendPlay (and less frequently, SendInput), the system and/or hook needs
	// another timeslice to ensure that AllowKeyToGoToSystem() actually takes effect on screen (SuppressThisKey()
	// doesn't seem to have this problem).
	if (!(mDoBackspace || mOmitEndChar) && mSendMode != SM_EVENT) // The final character of the abbreviation (or its EndChar) was not suppressed by the hook.
		Sleep(0);

	SendKeys(SendBuf, (SendRawModes)mSendRaw, mSendMode); // Send the backspaces and/or replacement.

	// Restore original values.
	g.KeyDelay = old_delay;
	g.PressDuration = old_press_duration;
	g.KeyDelayPlay = old_delay_play;
	g.PressDurationPlay = old_press_duration_play;
	g.SendLevel = old_send_level;
}



ResultType Hotstring::AddHotstring(LPCTSTR aName, IObjectPtr aCallback, LPCTSTR aOptions, LPCTSTR aHotstring
		, LPCTSTR aReplacement, bool aHasContinuationSection, UCHAR aSuspend)
// Returns OK or FAIL.
// Caller has ensured that aHotstringOptions is blank if there are no options.  Otherwise, aHotstringOptions
// should end in a colon, which marks the end of the options list.  aHotstring is the hotstring itself
// (e.g. "ahk"), which does not have to be unique, unlike aName, which was made unique by also including
// any options (e.g. ::ahk:: has a different aName than :c:ahk::).
// Caller has also ensured that aHotstring is not blank.
{
	// The length is limited for performance reasons, notably so that the hook does not have to move
	// memory around in the buffer it uses to watch for hotstrings:
	if (_tcslen(aHotstring) > MAX_HOTSTRING_LENGTH)
		return ValueError(_T("Hotstring max abbreviation length is ") MAX_HOTSTRING_LENGTH_STR _T("."), aHotstring, FAIL);

	if (!shs)
	{
		if (   !(shs = (Hotstring **)malloc(HOTSTRING_BLOCK_SIZE * sizeof(Hotstring *)))   )
			return MemoryError(); // Short msg. since so rare.
		sHotstringCountMax = HOTSTRING_BLOCK_SIZE;
	}
	else if (sHotstringCount >= sHotstringCountMax) // Realloc to preserve contents and keep contiguous array.
	{
		// Expand the array by one block.  Use a temp var. because realloc() returns NULL on failure
		// but leaves original block allocated.
		void *realloc_temp = realloc(shs, (sHotstringCountMax + HOTSTRING_BLOCK_SIZE) * sizeof(Hotstring *));
		if (!realloc_temp)
			return MemoryError();  // Short msg. since so rare.
		shs = (Hotstring **)realloc_temp;
		sHotstringCountMax += HOTSTRING_BLOCK_SIZE;
	}

	if (   !(shs[sHotstringCount] = new Hotstring(aName, aCallback, aOptions, aHotstring, aReplacement, aHasContinuationSection, aSuspend))   )
		return MemoryError(); // Short msg. since so rare.
	if (!shs[sHotstringCount]->mConstructedOK)
	{
		delete shs[sHotstringCount];  // SimpleHeap allows deletion of most recently added item.
		return FAIL;  // The constructor already displayed the error.
	}

	++sHotstringCount;
	if (!g_script.mIsReadyToExecute) // Caller is LoadIncludedFile(); allow BIF_Hotstring to manage this at runtime.
		++sEnabledCount; // This works because the script can't be suspended during startup (aSuspend is always FALSE).
	return OK;
}



Hotstring::Hotstring(LPCTSTR aName, IObjectPtr aCallback, LPCTSTR aOptions, LPCTSTR aHotstring, LPCTSTR aReplacement
	, bool aHasContinuationSection, UCHAR aSuspend)
	: mCallback(aCallback)  // Any NULL value will cause failure further below.
	, mName(NULL)
	, mString(NULL), mReplacement(NULL), mStringLength(0)
	, mHotCriterion(g->HotCriterion)
	, mSuspended(aSuspend)
	, mExistingThreads(0)
	, mMaxThreads(g_MaxThreadsPerHotkey)  // The value of g_MaxThreadsPerHotkey can vary during load-time.
	, mPriority(g_HSPriority), mKeyDelay(g_HSKeyDelay), mSendMode(g_HSSendMode)  // And all these can vary too.
	, mCaseSensitive(g_HSCaseSensitive), mConformToCase(g_HSConformToCase), mDoBackspace(g_HSDoBackspace)
	, mOmitEndChar(g_HSOmitEndChar), mSendRaw(aHasContinuationSection ? SCM_RAW_TEXT : g_HSSendRaw)
	, mEndCharRequired(g_HSEndCharRequired), mDetectWhenInsideWord(g_HSDetectWhenInsideWord), mDoReset(g_HSDoReset)
	, mInputLevel(g_InputLevel)
	, mSuspendExempt(g_SuspendExempt || g_SuspendExemptHS)
	, mConstructedOK(false)
{
	bool unused_x = false;
	ParseOptions(aOptions, mPriority, mKeyDelay, mSendMode, mCaseSensitive, mConformToCase, mDoBackspace
		, mOmitEndChar, mSendRaw, mEndCharRequired, mDetectWhenInsideWord, mDoReset, unused_x, mSuspendExempt);
	
	// To avoid memory leak, this is done only when it is certain the hotstring will be created:
	if (   !(mString = SimpleHeap::Malloc(aHotstring))   )
		return; // ScriptError() was already called by Malloc().
	if (   !(mName = SimpleHeap::Malloc(aName))   )
	{
		SimpleHeap::Delete(mString); // SimpleHeap allows deletion of most recently added item.
		return;
	}
	mStringLength = (UCHAR)_tcslen(mString);
	if (*aReplacement)
	{
		// SimpleHeap is not used for the replacement as it can be changed at runtime by Hotstring().
		if (   !(mReplacement = _tcsdup(aReplacement))   )
		{
			MemoryError(); // Short msg since very rare.
			return;
		}
	}
	else // Leave mReplacement NULL, but make this false so that the hook doesn't do extra work.
		mConformToCase = false;

	mConstructedOK = true; // Done at the very end.
}



void Hotstring::ParseOptions(LPCTSTR aOptions)
{
	bool unused_X_option;
	ParseOptions(aOptions, mPriority, mKeyDelay, mSendMode, mCaseSensitive, mConformToCase, mDoBackspace
		, mOmitEndChar, mSendRaw, mEndCharRequired, mDetectWhenInsideWord, mDoReset, unused_X_option, mSuspendExempt);
}



void Hotstring::ParseOptions(LPCTSTR aOptions, int &aPriority, int &aKeyDelay, SendModes &aSendMode
	, bool &aCaseSensitive, bool &aConformToCase, bool &aDoBackspace, bool &aOmitEndChar, SendRawType &aSendRaw
	, bool &aEndCharRequired, bool &aDetectWhenInsideWord, bool &aDoReset, bool &aExecuteAction, bool &aSuspendExempt)
{
	// In this case, colon rather than zero marks the end of the string.  However, the string
	// might be empty so check for that too.  In addition, this is now called from
	// IsDirective(), so that's another reason to check for normal string termination.
	for (auto cp = aOptions; *cp && *cp != ':'; ++cp)
	{
		auto cp1 = cp + 1;
		switch(ctoupper(*cp))
		{
		case '*':
			aEndCharRequired = (*cp1 == '0');
			break;
		case '?':
			aDetectWhenInsideWord = (*cp1 != '0');
			break;
		case 'B': // Do backspacing.
			aDoBackspace = (*cp1 != '0');
			break;
		case 'C':
			if (*cp1 == '0') // restore both settings to default.
			{
				aConformToCase = true;
				aCaseSensitive = false;
			}
			else if (*cp1 == '1')
			{
				aConformToCase = false;
				aCaseSensitive = false;
			}
			else // treat as plain "C"
			{
				aConformToCase = false;  // No point in conforming if its case sensitive.
				aCaseSensitive = true;
			}
			break;
		case 'O':
			aOmitEndChar = (*cp1 != '0');
			break;
		// For options such as K & P: Use atoi() vs. ATOI() to avoid interpreting something like 0x01C
		// as hex when in fact the C was meant to be an option letter:
		case 'K':
			aKeyDelay = _ttoi(cp1);
			break;
		case 'P':
			aPriority = _ttoi(cp1);
			break;
		case 'R':
			aSendRaw = (*cp1 != '0') ? SCM_RAW : SCM_NOT_RAW;
			break;
		case 'T':
			aSendRaw = (*cp1 != '0') ? SCM_RAW_TEXT : SCM_NOT_RAW;
			break;
		case 'S':
			if (*cp1)
				++cp; // Skip over S's sub-letter (if any) to exclude it from  further consideration.
			switch (ctoupper(*cp1))
			{
			// There is no means to choose SM_INPUT because it seems too rarely desired (since auto-replace
			// hotstrings would then become interruptible, allowing the keystrokes of fast typists to get
			// interspersed with the replacement text).
			case 'I': aSendMode = SM_INPUT_FALLBACK_TO_PLAY; break;
			case 'E': aSendMode = SM_EVENT; break;
			case 'P': aSendMode = SM_PLAY; break;
			default: aSuspendExempt = (*cp1 != '0'); break;
			}
			break;
		case 'Z':
			aDoReset = (*cp1 != '0');
			break;
		case 'X':
			aExecuteAction = (*cp1 != '0');
			break;
		// Otherwise: Ignore other characters, such as the digits that comprise the number after the P option.
		}
	}
}



Hotstring *Hotstring::FindHotstring(LPCTSTR aHotstring, bool aCaseSensitive, bool aDetectWhenInsideWord, HotkeyCriterion *aHotCriterion)
{
	for (UINT u = 0; u < sHotstringCount; ++u)
	{
		Hotstring &hs = *shs[u];
		// hs.mEndCharRequired is not checked because although it affects the conditions for activating
		// the hotstring, ::abbrev:: and :*:abbrev:: cannot co-exist (the latter would always take over).
		if (   hs.mHotCriterion == aHotCriterion // Same #HotIf criterion.
			&& hs.mCaseSensitive == aCaseSensitive // ::BTW:: and :C:BTW:: can co-exist.
			&& hs.mDetectWhenInsideWord == aDetectWhenInsideWord // :?:ion:: and ::ion:: can co-exist.
			&& (aCaseSensitive ? !_tcscmp(hs.mString, aHotstring) : !lstrcmpi(hs.mString, aHotstring))   ) // :C:BTW:: and :C:btw:: can co-exist, but not ::BTW:: and ::btw::.
			return &hs;
	}
	return NULL;
}



bif_impl FResult BIF_Hotstring(StrArg name, ExprTokenType *aReplacement, optl<StrArg> aOnOff, ResultToken &aResultToken)
{
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T("");

	TCHAR number_buf[MAX_NUMBER_SIZE];
	auto action = aReplacement ? TokenToString(*aReplacement, number_buf) : _T("");

	if (!_tcsicmp(name, _T("EndChars"))) // Equivalent to #Hotstring EndChars <action>
	{
		TokenSetResult(aResultToken, g_EndChars); // Return the old value.
		if (aReplacement)
			// There is some concern of a race condition with the hook thread, but since g_EndChars
			// has static storage duration and is therefore zero-initialized at startup (in particular,
			// the last char is always zero), the only consequence would be that some old end chars may
			// be used if a hotstring is evaluated during the tcslcpy() call.
			tcslcpy(g_EndChars, action, _countof(g_EndChars));
		return OK;
	}
	else if (!_tcsicmp(name, _T("MouseReset"))) // "MouseReset, true" seems more intuitive than "NoMouse, false"
	{
		bool previous_value = g_HSResetUponMouseClick;
		if (aReplacement)
		{
			g_HSResetUponMouseClick = TokenToBOOL(*aReplacement);
			if (g_HSResetUponMouseClick != previous_value && Hotstring::sEnabledCount) // No need if there aren't any hotstrings.
				Hotkey::ManifestAllHotkeysHotstringsHooks(); // Install the hook if needed, or uninstall if no longer needed.
		}
		aResultToken.SetValue(previous_value);
		return OK;
	}
	else if (!_tcsicmp(name, _T("Reset")))
	{
		*g_HSBuf = '\0';
		g_HSBufLength = 0;
		return OK;
	}
	else if (!aReplacement && !aOnOff.has_value() && *name != ':') // Equivalent to #Hotstring <name>
	{
		// TODO: Build string of current options and return it?
		bool unused_X_option; // 'X' option is required to be passed for each Hotstring() call, for clarity.
		Hotstring::ParseOptions(name, g_HSPriority, g_HSKeyDelay, g_HSSendMode, g_HSCaseSensitive
			, g_HSConformToCase, g_HSDoBackspace, g_HSOmitEndChar, g_HSSendRaw, g_HSEndCharRequired
			, g_HSDetectWhenInsideWord, g_HSDoReset, unused_X_option, g_SuspendExemptHS);
		return OK;
	}

	// Parse the hotstring name (this is similar to a section in LoadIncludedFile()):
	LPCTSTR hotstring_start = NULL;
	LPCTSTR hotstring_options = _T(""); // Set default as "no options were specified for this hotstring".
	if (name[0] == ':' && name[1])
	{
		if (name[1] != ':')
		{
			hotstring_options = name + 1; // Point it to the hotstring's option letters.
			// The following relies on the fact that options should never contain a literal colon.
			if (   !(hotstring_start = _tcschr(hotstring_options, ':'))   )
				hotstring_start = NULL; // Indicate that this isn't a hotstring after all.
			else
				++hotstring_start; // Points to the hotstring itself.
		}
		else // Double-colon, so it's a hotstring if there's more after this (but this means no options are present).
			if (name[2])
				hotstring_start = name + 2;
			//else it's just a naked "::", which is invalid.
	}
	if (!hotstring_start)
		return FR_E_ARG(0);
	
	// Determine options which affect hotstring identity/uniqueness.
	bool case_sensitive = g_HSCaseSensitive;
	bool detect_inside_word = g_HSDetectWhenInsideWord;
	bool un; int iun; SendModes sm; SendRawType sr; // Unused.
	bool execute_action = false;
	if (*hotstring_options)
		Hotstring::ParseOptions(hotstring_options, iun, iun, sm, case_sensitive, un, un, un, sr, un, detect_inside_word, un, execute_action, un);

	IObject *action_obj = NULL;
	if (aReplacement)
	{
		if (action_obj = TokenToObject(*aReplacement))
			action_obj->AddRef();
		else if (execute_action)
			return FR_E_ARG(1); // The 'X' option must be used together with a function object.
		// Otherwise, it's always replacement text
	}
	
	ToggleValueType toggle = NEUTRAL;
	if (  aOnOff.has_value() && !(toggle = Line::ConvertOnOffToggle(aOnOff.value()))  )
	{
		if (action_obj)
			action_obj->Release();
		return FR_E_ARG(2);
	}

	bool was_already_enabled;
	Hotstring *existing = Hotstring::FindHotstring(hotstring_start, case_sensitive, detect_inside_word, g->HotCriterion);
	if (existing)
	{
		was_already_enabled = !existing->mSuspended;

		// Update the replacement string or function, if specified.
		if (action_obj || *action)
		{
			LPTSTR new_replacement = NULL; // Set default: not auto-replace.
			if (!action_obj) // Caller specified a replacement string ('E' option was handled above).
			{
				if (existing->mReplacement && !_tcscmp(action, existing->mReplacement))
				{
					// Caller explicitly passed the same string it already had, which might be common,
					// such as if a single Hotstring() call site is used to both create and update.
					new_replacement = existing->mReplacement; // Avoid reallocating it.
				}
				else if (   !(new_replacement = _tcsdup(action))   )
					return FR_E_OUTOFMEM;
			}
			existing->mSuspended |= HS_TEMPORARILY_DISABLED;
			WaitHookIdle();
			// At this point it is certain the hook thread is not in the middle of reading this
			// hotstring's other properties, such as mReplacement (which we may be about to free).
			if (new_replacement != existing->mReplacement)
			{
				if (existing->mReplacement)
					free(existing->mReplacement);
				existing->mReplacement = new_replacement;
			}
			if (action_obj != existing->mCallback)
			{
				existing->mCallback = action_obj;
			}
		}
		// Update the hotstring's options.  Note that mCaseSensitive and mDetectWhenInsideWord
		// can't be changed this way since FindHotstring() would not have found it if they differed.
		// This is done after the above to avoid *partial* updates in the event of a failure.
		existing->ParseOptions(hotstring_options);
		switch (toggle)
		{
		case TOGGLE:      existing->mSuspended ^=  HS_TURNED_OFF; break;
		case TOGGLED_ON:  existing->mSuspended &= ~HS_TURNED_OFF; break;
		case TOGGLED_OFF: existing->mSuspended |=  HS_TURNED_OFF; break;
		}
		existing->mSuspended &= ~HS_TEMPORARILY_DISABLED; // Re-enable if it was disabled above.
	}
	else // No matching hotstring yet.
	{
		if (!action_obj && !*action)
			return FError(ERR_NONEXISTENT_HOTSTRING, name, ErrorPrototype::Target);

		UCHAR initial_suspend_state = (toggle == TOGGLED_OFF) ? HS_TURNED_OFF : FALSE;
		if (g_IsSuspended)
			initial_suspend_state |= HS_SUSPENDED;

		if (!Hotstring::AddHotstring(name, action_obj, hotstring_options, hotstring_start, action, false, initial_suspend_state))
		{
			if (action_obj)
				action_obj->Release();
			return FR_FAIL;
		}

		existing = Hotstring::shs[Hotstring::sHotstringCount-1];
		was_already_enabled = false; // Because it didn't exist.
	}

	if (action_obj)
		action_obj->Release();

	// Note that mSuspended must be 0 to count as enabled, meaning the hotstring was neither
	// turned off by us nor suspended by SuspendAll().  If it was suspended, there's no change
	// in status.
	bool is_enabled = !existing->mSuspended; // Important to avoid direct comparison with mSuspended becauses it isn't pure bool.
	if (is_enabled != was_already_enabled)
	{
		// One of the following just happened:
		//  - a hotstring was created and enabled
		//  - an existing disabled hotstring was just enabled
		//  - an existing enabled hotstring was just disabled
		UINT previously_enabled = Hotstring::sEnabledCount;
		Hotstring::sEnabledCount += is_enabled ? +1 : -1;
		if ((Hotstring::sEnabledCount > 0) != (previously_enabled > 0)) // Change in status of whether the hotstring recognizer is needed.
		{
			if (is_enabled)
			{
				*g_HSBuf = '\0'; // See SuspendAll() for comments.
				g_HSBufLength = 0;
			}
			if (!is_enabled || !g_KeybdHook) // Hook may not be needed anymore || hook is needed but not present.
				Hotkey::ManifestAllHotkeysHotstringsHooks();
		}
	}
	return OK;
}
