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

#include "stdafx.h"
#include "script.h"
#include "globaldata.h"
#include "script_func_impl.h"



bif_impl FResult CoordMode(StrArg aCommand, optl<StrArg> aMode)
{
	CoordModeType mode = aMode.has_value() ? Line::ConvertCoordMode(aMode.value()) : COORD_MODE_SCREEN;
	CoordModeType shift = Line::ConvertCoordModeCmd(aCommand);
	if (shift == COORD_MODE_INVALID)
		return FR_E_ARG(0);
	if (mode == COORD_MODE_INVALID)
		return FR_E_ARG(1);
	g->CoordMode = (g->CoordMode & ~(COORD_MODE_MASK << shift)) | (mode << shift);
	return OK;
}


bif_impl FResult SendMode(StrArg aMode)
{
	auto new_mode = Line::ConvertSendMode(aMode, SM_INVALID);
	if (new_mode == SM_INVALID)
		return FR_E_ARG(0);
	g->SendMode = new_mode;
	return OK;
}


bif_impl FResult SendLevel(int aLevel)
{
	if (!SendLevelIsValid(aLevel))
		return FR_E_ARG(0);
	g->SendLevel = aLevel;
	return OK;
}


bif_impl FResult SetDefaultMouseSpeed(int aSpeed)
{
	if (aSpeed < 0 || aSpeed > MAX_MOUSE_SPEED)
		return FR_E_ARG(0);
	g->DefaultMouseSpeed = (UCHAR)aSpeed;
	return OK;
}


bif_impl FResult SetKeyDelay(optl<int> aDelay, optl<int> aDuration, optl<StrArg> aMode)
{
	if (aDelay.has_value() && *aDelay < -1)
		return FR_E_ARG(0);
	if (aDuration.has_value() && *aDuration < -1)
		return FR_E_ARG(1);
	if (aMode.has_value() && !_tcsicmp(aMode.value(), _T("Play")))
	{
		if (aDelay.has_value())
			g->KeyDelayPlay = *aDelay;
		if (aDuration.has_value())
			g->PressDurationPlay = *aDuration;
	}
	else
	{
		if (aMode.has_nonempty_value()) // Anything other than "Play" or "" is invalid.
			return FR_E_ARG(2);
		if (aDelay.has_value())
			g->KeyDelay = *aDelay;
		if (aDuration.has_value())
			g->PressDuration = *aDuration;
	}
	return OK;
}


bif_impl FResult SetMouseDelay(int aDelay, optl<StrArg> aMode)
{
	if (aDelay < -1)
		return FR_E_ARG(0);
	if (aMode.has_value() && !_tcsicmp(aMode.value(), _T("Play")))
		g->MouseDelayPlay = aDelay;
	else if (aMode.has_nonempty_value()) // Anything other than "Play" or "" is invalid.
		return FR_E_ARG(1);
	else
		g->MouseDelay = aDelay;
	return OK;
}


bif_impl FResult SetWinDelay(int aDelay)
{
	if (aDelay < -1)
		return FR_E_ARG(0);
	g->WinDelay = aDelay;
	return OK;
}


bif_impl FResult SetControlDelay(int aDelay)
{
	if (aDelay < -1)
		return FR_E_ARG(0);
	g->ControlDelay = aDelay;
	return OK;
}



BIV_DECL_R(BIV_Clipboard)
{
	auto length = g_clip.Get();
	if (TokenSetResult(aResultToken, nullptr, length))
	{
		aResultToken.marker_length = g_clip.Get(aResultToken.marker);
		if (aResultToken.marker_length == CLIPBOARD_FAILURE)
			aResultToken.SetExitResult(FAIL);
		aResultToken.symbol = SYM_STRING;
	}
	g_clip.Close();
}

BIV_DECL_W(BIV_Clipboard_Set)
{
	if (auto *obj = BivRValueToObject())
	{
		if (ClipboardAll *cba = dynamic_cast<ClipboardAll *>(obj))
		{
			if (!Var::SetClipboardAll(cba->Data(), cba->Size()))
				_f_return_FAIL;
			return;
		}
		_f_throw_type(_T("ClipboardAll"), obj->Type());
	}
	size_t aLength;
	LPTSTR aBuf = BivRValueToString(&aLength);
	if (!g_clip.Set(aBuf, aLength))
		_f_return_FAIL;
}


BIV_DECL_R(BIV_MMM_DDD)
{
	LPTSTR format_str;
	switch(ctoupper(aVarName[2]))
	{
	// Use the case-sensitive formats required by GetDateFormat():
	case 'M': format_str = (aVarName[5] ? _T("MMMM") : _T("MMM")); break;
	case 'D': format_str = (aVarName[5] ? _T("dddd") : _T("ddd")); break;
	}
	// Confirmed: The below will automatically use the local time (not UTC) when 3rd param is NULL.
	int len = GetDateFormatEx(LOCALE_NAME_USER_DEFAULT, 0, nullptr, format_str, _f_retval_buf, _f_retval_buf_size, nullptr);
	if (len && !_f_retval_buf[len - 1])
		--len; // "Returns the number of characters written" apparently includes the terminator.
	_f_return_p(_f_retval_buf, len);
}

VarSizeType GetDateTimeBIV(LPTSTR aBuf, LPTSTR aVarName)
{
	if (!aBuf)
		return 6; // Since only an estimate is needed in this mode, return the maximum length of any item.

	aVarName += 2; // Skip past the "A_".

	// The current time is refreshed only if it's been a certain number of milliseconds since
	// the last fetch of one of these built-in time variables.  This keeps the variables in
	// sync with one another when they are used consecutively such as this example:
	// Var := A_Hour ':' A_Min ':' A_Sec
	// Using GetTickCount() because it's very low overhead compared to the other time functions:
	static DWORD sLastUpdate = 0; // Static should be thread + recursion safe in this case.
	static SYSTEMTIME sST = {0}; // Init to detect when it's empty.
	BOOL is_msec = !_tcsicmp(aVarName, _T("MSec")); // Always refresh if it's milliseconds, for better accuracy.
	DWORD now_tick = GetTickCount();
	if (is_msec || now_tick - sLastUpdate > 50 || !sST.wYear) // See comments above.
	{
		GetLocalTime(&sST);
		sLastUpdate = now_tick;
	}

	if (is_msec)
		return _stprintf(aBuf, _T("%03d"), sST.wMilliseconds);

	TCHAR second_letter = ctoupper(aVarName[1]);
	switch(ctoupper(aVarName[0]))
	{
	case 'Y':
		switch(second_letter)
		{
		case 'D': // A_YDay
			return _stprintf(aBuf, _T("%d"), GetYDay(sST.wMonth, sST.wDay, IS_LEAP_YEAR(sST.wYear)));
		case 'W': // A_YWeek
			return GetISOWeekNumber(aBuf, sST.wYear
				, GetYDay(sST.wMonth, sST.wDay, IS_LEAP_YEAR(sST.wYear))
				, sST.wDayOfWeek);
		default:  // A_Year/A_YYYY
			return _stprintf(aBuf, _T("%d"), sST.wYear);
		}
		// No break because all cases above return:
		//break;
	case 'M':
		switch(second_letter)
		{
		case 'D': // A_MDay (synonymous with A_DD)
			return _stprintf(aBuf, _T("%02d"), sST.wDay);
		case 'I': // A_Min
			return _stprintf(aBuf, _T("%02d"), sST.wMinute);
		default: // A_MM and A_Mon (A_MSec was already completely handled higher above).
			return _stprintf(aBuf, _T("%02d"), sST.wMonth);
		}
		// No break because all cases above return:
		//break;
	case 'D': // A_DD (synonymous with A_MDay)
		return _stprintf(aBuf, _T("%02d"), sST.wDay);
	case 'W': // A_WDay
		return _stprintf(aBuf, _T("%d"), sST.wDayOfWeek + 1);
	case 'H': // A_Hour
		return _stprintf(aBuf, _T("%02d"), sST.wHour);
	case 'S': // A_Sec (A_MSec was already completely handled higher above).
		return _stprintf(aBuf, _T("%02d"), sST.wSecond);
	}
	return 0; // Never reached, but avoids compiler warning.
}

BIV_DECL_R(BIV_DateTime)
{
	GetDateTimeBIV(_f_retval_buf, aVarName);
	_f_return_p(_f_retval_buf);
}

BIV_DECL_R(BIV_ListLines)
{
	_f_return_b(g->ListLinesIsEnabled);
}

BIV_DECL_W(BIV_ListLines_Set)
{
	g->ListLinesIsEnabled = BivRValueToBOOL();
}

BIV_DECL_R(BIV_TitleMatchMode)
{
	if (g->TitleMatchMode == FIND_REGEX) // v1.0.45.
		// For backward compatibility (due to possible use of case-sensitive comparison), never change the case used here:
		_f_return_p(_T("RegEx"));
	// Otherwise, it's a numerical mode:
	_f_return_i(g->TitleMatchMode);
}

BIV_DECL_W(BIV_TitleMatchMode_Set)
{
	LPTSTR aBuf = BivRValueToString();
	switch (Line::ConvertTitleMatchMode(aBuf))
	{
	case FIND_IN_LEADING_PART: g->TitleMatchMode = FIND_IN_LEADING_PART; break;
	case FIND_ANYWHERE: g->TitleMatchMode = FIND_ANYWHERE; break;
	case FIND_REGEX: g->TitleMatchMode = FIND_REGEX; break;
	case FIND_EXACT: g->TitleMatchMode = FIND_EXACT; break;
	// For simplicity, this function handles both variables.
	case FIND_FAST: g->TitleFindFast = true; break;
	case FIND_SLOW: g->TitleFindFast = false; break;
	default:
		_f_throw_value(ERR_INVALID_VALUE, aBuf);
	}
}

BIV_DECL_R(BIV_TitleMatchModeSpeed)
{
	// For backward compatibility (due to possible use of case-sensitive comparison), never change the case used here:
	_f_return_p(g->TitleFindFast ? _T("Fast") : _T("Slow"));
}

BIV_DECL_R(BIV_DetectHiddenWindows)
{
	_f_return_b(g->DetectHiddenWindows);
}

BIV_DECL_W(BIV_DetectHiddenWindows_Set)
{
	g->DetectHiddenWindows = BivRValueToBOOL();
}

BIV_DECL_R(BIV_DetectHiddenText)
{
	_f_return_b(g->DetectHiddenText);
}

BIV_DECL_W(BIV_DetectHiddenText_Set)
{
	g->DetectHiddenText = BivRValueToBOOL();
}

int& BIV_xDelay(LPTSTR aVarName)
{
	global_struct &g = *::g; // Reduces code size.
	switch (ctoupper(aVarName[2])) // a_X...
	{
	case 'K':
		if (ctolower(aVarName[6]) == 'e') // a_keydE...
		{
			if (aVarName[10]) // a_keydelayP...
				return g.KeyDelayPlay;
			else
				return g.KeyDelay;
		}
		else // a_keydU...
		{
			if (aVarName[13]) // a_keydurationP...
				return g.PressDurationPlay;
			else
				return g.PressDuration;
		}
	case 'M':
		if (aVarName[12]) // a_mousedelayP...
			return g.MouseDelayPlay;
		else
			return g.MouseDelay;
	case 'W':
		return g.WinDelay;
	//case 'C':
	default:
		return g.ControlDelay;
	}
}

BIV_DECL_R(BIV_xDelay)
{
	_f_return_i(BIV_xDelay(aVarName));
}

BIV_DECL_W(BIV_xDelay_Set)
{
	BIV_xDelay(aVarName) = (int)BivRValueToInt64();
}

BIV_DECL_R(BIV_DefaultMouseSpeed)
{
	_f_return_i(g->DefaultMouseSpeed);
}

BIV_DECL_W(BIV_DefaultMouseSpeed_Set)
{
	g->DefaultMouseSpeed = (int)BivRValueToInt64();
}

BIV_DECL_R(BIV_CoordMode)
{
	static LPTSTR sCoordModes[] = COORD_MODES;
	_f_return_p(sCoordModes[(g->CoordMode >> Line::ConvertCoordModeCmd(aVarName + 11)) & COORD_MODE_MASK]);
}

BIV_DECL_W(BIV_CoordMode_Set)
{
	auto value = BivRValueToString();
	if (FAILED(CoordMode(aVarName + 11, value))) // A_CoordMode is 11 chars.
		_f_throw_value(ERR_INVALID_VALUE, value);
}

BIV_DECL_R(BIV_SendMode)
{
	static LPTSTR sSendModes[] = SEND_MODES;
	_f_return_p(sSendModes[g->SendMode]);
}

BIV_DECL_W(BIV_SendMode_Set)
{
	auto value = BivRValueToString();
	if (FAILED(SendMode(value)))
		_f_throw_value(ERR_INVALID_VALUE, value);
}

BIV_DECL_R(BIV_SendLevel)
{
	_f_return_i(g->SendLevel);
}

BIV_DECL_W(BIV_SendLevel_Set)
{
	Throw_if_RValue_NaN();
	if (FAILED(SendLevel((int)BivRValueToInt64())))
		_f_throw_value(ERR_INVALID_VALUE, BivRValueToString());
}

BIV_DECL_R(BIV_StoreCapsLockMode)
{
	_f_return_b(g->StoreCapslockMode);
}

BIV_DECL_W(BIV_StoreCapsLockMode_Set)
{
	g->StoreCapslockMode = BivRValueToBOOL();
}

BIV_DECL_R(BIV_Hotkey)
{
	if (aVarName[2] == 'M') // A_MaxHotkeysPerInterval
		_f_return_i(g_MaxHotkeysPerInterval);
	if (aVarName[8] == 'I') // A_HotkeyInterval
		_f_return_i(g_HotkeyThrottleInterval);
	// A_HotkeyModifierTimeout
	_f_return_i(g_HotkeyModifierTimeout);
}

BIV_DECL_W(BIV_Hotkey_Set)
{
	Throw_if_RValue_NaN();
	int value = (int)BivRValueToInt64();
	if (aVarName[2] == 'M') // A_MaxHotkeysPerInterval
	{
		if (value < 1)
			_f_throw_value(ERR_INVALID_VALUE);
		g_MaxHotkeysPerInterval = value;
	}
	else if (aVarName[8] == 'I') // A_HotkeyInterval
	{
		if (value < 0)
			_f_throw_value(ERR_INVALID_VALUE);
		g_HotkeyThrottleInterval = value;
	}
	else // A_HotkeyModifierTimeout
		g_HotkeyModifierTimeout = value;
}

BIV_DECL_R(BIV_MenuMaskKey)
{
	if (!g_MenuMaskKeyVK && !g_MenuMaskKeySC)
		_f_return_empty; // Return a "false" value to indicate there is no masking.
	// For uniformity, simplicity and to avoid any loss of information, always return vkNNscNNN.
	auto len = sntprintf(_f_retval_buf, _f_retval_buf_size, _T("vk%02Xsc%03X"), g_MenuMaskKeyVK, g_MenuMaskKeySC);
	_f_return_p(_f_retval_buf, len);
}

BIV_DECL_W(BIV_MenuMaskKey_Set)
{
	auto keyname = BivRValueToString();
	if (!*keyname) // Allow "" to mean "no masking".
	{
		g_MenuMaskKeyVK = 0;
		g_MenuMaskKeySC = 0;
		return;
	}
	vk_type vk;
	sc_type sc;
	// Testing shows that sending an event with zero VK but non-zero SC fails to suppress
	// the Start menu (although it does suppress the window menu).  However, allowing all
	// valid key strings seems more correct than requiring g_MenuMaskKeyVK != 0, and adds
	// flexibility at very little cost.  Note that this use of TextToVKandSC()'s return
	// value (vs. checking VK|SC) allows vk00sc000 to turn off masking altogether.
	if (!TextToVKandSC(keyname, vk, sc))
		_f_throw_value(ERR_INVALID_VALUE);
	g_MenuMaskKeyVK = vk;
	g_MenuMaskKeySC = sc;
}

BIV_DECL_R(BIV_IsPaused) // v1.0.48: Lexikos: Added BIV_IsPaused and BIV_IsCritical.
{
	// Although A_IsPaused could indicate how many threads are paused beneath the current thread,
	// that would be a problem because it would yield a non-zero value even when the underlying thread
	// isn't paused (i.e. other threads below it are paused), which would defeat the original purpose.
	// In addition, A_IsPaused probably won't be commonly used, so it seems best to keep it simple.
	// NAMING: A_IsPaused seems to be a better name than A_Pause or A_Paused due to:
	//    Better readability.
	//    Consistent with A_IsSuspended, which is strongly related to pause/unpause.
	//    The fact that it wouldn't be likely for a function to turn off pause then turn it back on
	//      (or vice versa), which was the main reason for storing "Off" and "On" in things like
	//      A_DetectHiddenWindows.
	// Checking g>g_array avoids any chance of underflow, which might otherwise happen if this is
	// called by the AutoExec section or a threadless callback running in thread #0.
	_f_return_b(g > g_array && g[-1].IsPaused);
}

BIV_DECL_R(BIV_IsCritical) // v1.0.48: Lexikos: Added BIV_IsPaused and BIV_IsCritical.
{
	// It seems more useful to return g->PeekFrequency than "On" or "Off" (Critical() ensures that
	// g->PeekFrequency!=0 whenever g->ThreadIsCritical==true).  Also, the word "Is" in "A_IsCritical"
	// implies a value that can be used as a boolean such as "if A_IsCritical".
	if (g->ThreadIsCritical)
		_f_return_i(g->PeekFrequency); // Critical() ensures that g->PeekFrequency > 0 when critical is on.
	// Otherwise:
	_f_return_i(0);
}

BIV_DECL_R(BIV_IsSuspended)
{
	_f_return_b(g_IsSuspended);
}



BIV_DECL_R(BIV_IsCompiled)
{
#ifdef AUTOHOTKEYSC
	_f_return_b(true);
#else
	_f_return_b(g_script.mKind == Script::ScriptKindResource);
#endif
}



BIV_DECL_R(BIV_FileEncoding)
{
	// A similar section may be found under "case Encoding:" in FileObject::Invoke.  Maintain that with this:
	LPTSTR enc;
	switch (g->Encoding)
	{
	// Returning readable strings for these seems more useful than returning their numeric values, especially with CP_AHKNOBOM:
	case CP_UTF8:                enc = _T("UTF-8");      break;
	case CP_UTF8 | CP_AHKNOBOM:  enc = _T("UTF-8-RAW");  break;
	case CP_UTF16:               enc = _T("UTF-16");     break;
	case CP_UTF16 | CP_AHKNOBOM: enc = _T("UTF-16-RAW"); break;
	default:
	  {
		enc = _f_retval_buf;
		enc[0] = _T('C');
		enc[1] = _T('P');
		_itot(g->Encoding, enc + 2, 10);
	  }
	}
	_f_return_p(enc);
}

BIV_DECL_W(BIV_FileEncoding_Set)
{
	LPTSTR aBuf = BivRValueToString();
	UINT new_encoding = Line::ConvertFileEncoding(aBuf);
	if (new_encoding == -1)
		_f_throw_value(ERR_INVALID_VALUE, aBuf);
	g->Encoding = new_encoding;
}



BIV_DECL_R(BIV_RegView)
{
	LPTSTR value;
	switch (g->RegView)
	{
	case KEY_WOW64_32KEY: value = _T("32"); break;
	case KEY_WOW64_64KEY: value = _T("64"); break;
	default: value = _T("Default"); break;
	}
	_f_return_p(value);
}

BIV_DECL_W(BIV_RegView_Set)
{
	LPTSTR aBuf = BivRValueToString();
	DWORD reg_view = Line::RegConvertView(aBuf);
	// Validate the parameter even if it's not going to be used.
	if (reg_view == -1)
		_f_throw_value(ERR_INVALID_VALUE, aBuf);
	// Since these flags cause the registry functions to fail on Win2k and have no effect on
	// any later 32-bit OS, ignore this command when the OS is 32-bit.  Leave A_RegView = "Default".
	if (IsOS64Bit())
		g->RegView = reg_view;
}



BIV_DECL_R(BIV_LastError)
{
	_f_return_i(g->LastError);
}

BIV_DECL_W(BIV_LastError_Set)
{
	SetLastError(g->LastError = (DWORD)BivRValueToInt64());
}



BIV_DECL_R(BIV_PtrSize)
{
	_f_return_i(sizeof(void *));
}



BIV_DECL_R(BIV_ScreenDPI)
{
	_f_return_i(g_ScreenDPI);
}



BIV_DECL_R(BIV_TrayMenu)
{
	// Currently ExpandExpression() does not automatically release objects returned
	// by BIVs since this is the only one, and we're retaining a reference to it.
	//g_script.mTrayMenu->AddRef();
	_f_return(g_script.mTrayMenu);
}



BIV_DECL_R(BIV_AllowMainWindow)
{
	_f_return_b(g_AllowMainWindow);
}

BIV_DECL_W(BIV_AllowMainWindow_Set)
{
	g_script.AllowMainWindow(BivRValueToBOOL());
}



BIV_DECL_R(BIV_IconHidden)
{
	_f_return_b(g_NoTrayIcon);
}

BIV_DECL_W(BIV_IconHidden_Set)
{
	g_script.ShowTrayIcon(!BivRValueToBOOL());
}

void Script::ShowTrayIcon(bool aShow)
{
	if (g_NoTrayIcon = !aShow) // Assign.
	{
		if (mNIC.hWnd) // Since it exists, destroy it.
		{
			Shell_NotifyIcon(NIM_DELETE, &mNIC); // Remove it.
			mNIC.hWnd = NULL;  // Set this as an indicator that tray icon is not installed.
			// but don't do DestroyMenu() on mTrayMenu->mMenu (if non-NULL) since it may have been
			// changed by the user to have the custom items on top of the standard items,
			// for example, and we don't want to lose that ordering in case the script turns
			// the icon back on at some future time during this session.  Also, the script
			// may provide some other means of displaying the tray menu.
		}
	}
	else
	{
		if (!mNIC.hWnd) // The icon doesn't exist, so create it.
		{
			CreateTrayIcon();
			UpdateTrayIcon(true);  // Force the icon into the correct pause/suspend state.
		}
	}
}

BIV_DECL_R(BIV_IconTip)
{
	// Return the custom tip if any, otherwise the default tip.
	_f_return_p(g_script.mTrayIconTip ? g_script.mTrayIconTip : g_script.mFileName);
}

BIV_DECL_W(BIV_IconTip_Set)
{
	g_script.SetTrayTip(BivRValueToString());
}

void Script::SetTrayTip(LPTSTR aText)
{
	// Allocate mTrayIconTip on first use even if aText is empty, so that
	// it will override the use of mFileName as the tray tip text.
	// This allows the script to completely disable the tray tooltip.
	if (!mTrayIconTip)
		mTrayIconTip = SimpleHeap::Alloc<TCHAR>(_countof(mNIC.szTip)); // SimpleHeap improves avg. case mem load.
	if (mTrayIconTip)
		tcslcpy(mTrayIconTip, aText, _countof(mNIC.szTip));
	if (mNIC.hWnd) // i.e. only update the tip if the tray icon exists (can't work otherwise).
	{
		UPDATE_TIP_FIELD
		Shell_NotifyIcon(NIM_MODIFY, &mNIC);  // Currently not checking its result (e.g. in case a shell other than Explorer is running).
	}
}

BIV_DECL_R(BIV_IconFile)
{
	_f_return_p(g_script.mCustomIconFile ? g_script.mCustomIconFile : _T(""));
}

BIV_DECL_R(BIV_IconNumber)
{
	_f_return_i(g_script.mCustomIconNumber);
}



BIV_DECL_R(BIV_PriorKey)
{
	int validEventCount = 0;
	// Start at the current event (offset 1)
	for (int iOffset = 1; iOffset <= g_MaxHistoryKeys; ++iOffset)
	{
		// Get index for circular buffer
		int i = (g_KeyHistoryNext + g_MaxHistoryKeys - iOffset) % g_MaxHistoryKeys;
		// Keep looking until we hit the second valid event
		if (g_KeyHistory[i].event_type != _T('i') // Not an ignored event.
			&& g_KeyHistory[i].event_type != _T('U') // Not a Unicode packet (SendInput/VK_PACKET).
			&& ++validEventCount > 1)
		{
			// Find the next most recent key-down
			if (!g_KeyHistory[i].key_up)
			{
				GetKeyName(g_KeyHistory[i].vk, g_KeyHistory[i].sc, _f_retval_buf, _f_retval_buf_size);
				_f_return_p(_f_retval_buf);
			}
		}
	}
	_f_return_empty;
}



LPTSTR GetExitReasonString(ExitReasons aExitReason)
{
	LPTSTR str;
	switch(aExitReason)
	{
	case EXIT_LOGOFF: str = _T("Logoff"); break;
	case EXIT_SHUTDOWN: str = _T("Shutdown"); break;
	// Since the below are all relatively rare, except WM_CLOSE perhaps, they are all included
	// as one word to cut down on the number of possible words (it's easier to write OnExit
	// functions to cover all possibilities if there are fewer of them).
	case EXIT_CRITICAL:
	case EXIT_DESTROY:
	case EXIT_CLOSE: str = _T("Close"); break;
	case EXIT_ERROR: str = _T("Error"); break;
	case EXIT_MENU: str = _T("Menu"); break;  // Standard menu, not a user-defined menu.
	case EXIT_EXIT: str = _T("Exit"); break;  // ExitApp or Exit command.
	case EXIT_RELOAD: str = _T("Reload"); break;
	case EXIT_SINGLEINSTANCE: str = _T("Single"); break;
	default:  // EXIT_NONE or unknown value (unknown would be considered a bug if it ever happened).
		str = _T("");
	}
	return str;
}



BIV_DECL_R(BIV_Space_Tab)
{
	_f_return_p(aVarName[5] ? _T(" ") : _T("\t"));
}

BIV_DECL_R(BIV_AhkVersion)
{
	_f_return_p(T_AHK_VERSION);
}

BIV_DECL_R(BIV_AhkPath)
{
#ifdef AUTOHOTKEYSC
	TCHAR buf[MAX_PATH];
	size_t length;
	if (length = GetAHKInstallDir(buf))
		// Name "AutoHotkey.exe" is assumed for code size reduction and because it's not stored in the registry:
		tcslcpy(buf + length, _T("\\AutoHotkey.exe"), MAX_PATH - length); // tcslcpy() in case registry has a path that is too close to MAX_PATH to fit AutoHotkey.exe
	//else leave it blank as documented.
	_f_return(buf, length);
#else
	_f_return_p(g_script.mOurEXE);
#endif
}



BIV_DECL_R(BIV_TickCount)
{
	_f_return(GetTickCount64());
}



BIV_DECL_R(BIV_Now)
{
	SYSTEMTIME st;
	if (aVarName[5]) // A_Now[U]TC
		GetSystemTime(&st);
	else
		GetLocalTime(&st);
	SystemTimeToYYYYMMDD(_f_retval_buf, st);
	_f_return_p(_f_retval_buf);
}

BIV_DECL_R(BIV_OSVersion)
{
	_f_return_p(const_cast<LPTSTR>(g_os.Version()));
}

BIV_DECL_R(BIV_Is64bitOS)
{
	_f_return_b(IsOS64Bit());
}

BIV_DECL_R(BIV_Language)
{
	LPTSTR aBuf = _f_retval_buf;
	_stprintf(aBuf, _T("%04hX"), GetSystemDefaultUILanguage());
	_f_return_p(aBuf);
}

BIV_DECL_R(BIV_UserName_ComputerName)
{
	TCHAR buf[MAX_PATH];  // Doesn't use MAX_COMPUTERNAME_LENGTH + 1 in case longer names are allowed in the future.
	DWORD buf_size = MAX_PATH; // Below: A_Computer[N]ame (N is the 11th char, index 10, which if present at all distinguishes between the two).
	if (   !(aVarName[10] ? GetComputerName(buf, &buf_size) : GetUserName(buf, &buf_size))   )
		*buf = '\0';
	_f_return(buf);
}

BIV_DECL_R(BIV_WorkingDir)
{
	// Use GetCurrentDirectory() vs. g_WorkingDir because any in-progress FileSelect()
	// dialog is able to keep functioning even when it's quasi-thread is suspended.  The
	// dialog can thus change the current directory as seen by the active quasi-thread even
	// though g_WorkingDir hasn't been updated.  It might also be possible for the working
	// directory to change in unusual circumstances such as a network drive being lost).
	// Update: FileSelectFile changing the current directory might be OS-specific;
	// I could not reproduce it on Windows 10.
	TCHAR buf[T_MAX_PATH]; // T_MAX_PATH vs. MAX_PATH only has an effect with Windows 10 long path awareness.
	DWORD length = GetCurrentDirectory(_countof(buf), buf);
	_f_return(buf, length);
}

BIV_DECL_W(BIV_WorkingDir_Set)
{
	if (SetWorkingDir(BivRValueToString()) != OK)
		_f_throw_win32();
}

BIV_DECL_R(BIV_InitialWorkingDir)
{
	_f_return_p(g_WorkingDirOrig);
}

BIV_DECL_R(BIV_WinDir)
{
	TCHAR buf[MAX_PATH]; // MSDN (2018): The uSize parameter "should be set to MAX_PATH."
	VarSizeType length = GetSystemWindowsDirectory(buf, _countof(buf));
	_f_return(buf, length);
}

BIV_DECL_R(BIV_Temp)
{
	TCHAR buf[MAX_PATH+1]; // MSDN (2018): "The maximum possible return value is MAX_PATH+1 (261)."
	VarSizeType length = GetTempPath(_countof(buf), buf);
	if (length)
		if (buf[length - 1] == '\\') // Should always be true. MSDN: "The returned string ends with a backslash"
			buf[--length] = '\0'; // Omit the trailing backslash to improve friendliness/consistency.
	_f_return(buf, length);
}

BIV_DECL_R(BIV_ComSpec)
{
	TCHAR buf_temp[1]; // Just a fake buffer to pass to some API functions in lieu of a NULL, to avoid any chance of misbehavior. Keep the size at 1 so that API functions will always fail to copy to buf.
	auto size_required = GetEnvironmentVariable(_T("ComSpec"), buf_temp, 0);
	if (!TokenSetResult(aResultToken, nullptr, size_required)) // Avoids subtracting 1 to be conservative and to reduce code size (due to the need to otherwise check for zero and avoid subtracting 1 in that case).
		return;
	aResultToken.symbol = SYM_STRING;
	aResultToken.marker_length = GetEnvVarReliable(_T("ComSpec"), aResultToken.marker);
}

BIV_DECL_R(BIV_SpecialFolderPath)
{
	TCHAR buf[MAX_PATH]; // One caller relies on this being explicitly limited to MAX_PATH.
	// SHGetFolderPath requires a buffer size of MAX_PATH, but the function was superseded
	// by SHGetKnownFolderPath in Windows Vista, and that function returns COM-allocated
	// memory of unknown length.  However, it seems the shell still does not support long
	// paths as of 2018.
	int aFolder;
	switch (ctoupper(aVarName[2]))
	{
	case 'P': // A_[P]rogram...
		if (ctoupper(aVarName[9]) == 'S') // A_Programs(Common)
			aFolder = aVarName[10] ? CSIDL_COMMON_PROGRAMS : CSIDL_PROGRAMS;
		else // A_Program[F]iles
			aFolder = CSIDL_PROGRAM_FILES;
		break;
	case 'A': // A_AppData(Common)
		aFolder = aVarName[9] ? CSIDL_COMMON_APPDATA : CSIDL_APPDATA;
		break;
	case 'D': // A_Desktop(Common)
		aFolder = aVarName[9] ? CSIDL_COMMON_DESKTOPDIRECTORY : CSIDL_DESKTOPDIRECTORY;
		break;
	case 'S':
		if (ctoupper(aVarName[7]) == 'M') // A_Start[M]enu(Common)
			aFolder = aVarName[11] ? CSIDL_COMMON_STARTMENU : CSIDL_STARTMENU;
		else // A_Startup(Common)
			aFolder = aVarName[9] ? CSIDL_COMMON_STARTUP : CSIDL_STARTUP;
		break;
#ifdef _DEBUG
	default:
		ASSERT(!_T("DEBUG: Unhandled SpecialFolderPath variable."));
#endif
	}
	if (SHGetFolderPath(NULL, aFolder, NULL, SHGFP_TYPE_CURRENT, buf) != S_OK)
		*buf = '\0';
	_f_return(buf);
}

BIV_DECL_R(BIV_MyDocuments) // Called by multiple callers.
{
	TCHAR buf[MAX_PATH]; // SHGetFolderPath requires a buffer size of MAX_PATH.  At least one caller relies on this.
	if (SHGetFolderPath(NULL, CSIDL_MYDOCUMENTS, NULL, SHGFP_TYPE_CURRENT, buf) != S_OK)
		*buf = '\0';
	// Since it is common (such as in networked environments) to have My Documents on the root of a drive
	// (such as a mapped drive letter), remove the backslash from something like M:\ because M: is more
	// appropriate for most uses:
	VarSizeType length = (VarSizeType)strip_trailing_backslash(buf);
	_f_return(buf, length);
}



BIF_DECL(BIF_CaretGetPos)
{
	Var *varX = ParamIndexToOutputVar(0);
	Var *varY = ParamIndexToOutputVar(1);
	
	// I believe only the foreground window can have a caret position due to relationship with focused control.
	HWND target_window = GetForegroundWindow(); // Variable must be named target_window for ATTACH_THREAD_INPUT.
	if (!target_window) // No window is in the foreground, report blank coordinate.
	{
		if (varX)
			varX->Assign();
		if (varY)
			varY->Assign();
		_f_return_i(FALSE);
	}

	DWORD now_tick = GetTickCount();

	GUITHREADINFO info;
	info.cbSize = sizeof(GUITHREADINFO);
	BOOL result = GetGUIThreadInfo(GetWindowThreadProcessId(target_window, NULL), &info) // Got info okay...
		&& info.hwndCaret; // ...and there is a caret.
	if (!result)
	{
		if (varX)
			varX->Assign();
		if (varY)
			varY->Assign();
		_f_return_i(FALSE);
	}
	POINT pt;
	pt.x = info.rcCaret.left;
	pt.y = info.rcCaret.top;
	// Unconditionally convert to screen coordinates, for simplicity.
	ClientToScreen(info.hwndCaret, &pt);
	// Now convert back to whatever is expected for the current mode.
	POINT origin = {0};
	CoordToScreen(origin, COORD_MODE_CARET);
	pt.x -= origin.x;
	pt.y -= origin.y;
	
	if (varX)
		varX->Assign(pt.x);
	if (varY)
		varY->Assign(pt.y);
	_f_return_i(TRUE);
}



BIV_DECL_R(BIV_Cursor)
{
	HCURSOR current_cursor;
	CURSORINFO ci;
	ci.cbSize = sizeof(CURSORINFO);
	current_cursor = GetCursorInfo(&ci) ? ci.hCursor : NULL;
	if (!current_cursor)
	{
		#define CURSOR_UNKNOWN _T("Unknown")
		_f_return_p(CURSOR_UNKNOWN);
	}

	// Static so that it's initialized on first use (should help performance after the first time):
	static HCURSOR sCursor[] = {LoadCursor(NULL, IDC_APPSTARTING), LoadCursor(NULL, IDC_ARROW)
		, LoadCursor(NULL, IDC_CROSS), LoadCursor(NULL, IDC_HELP), LoadCursor(NULL, IDC_IBEAM)
		, LoadCursor(NULL, IDC_ICON), LoadCursor(NULL, IDC_NO), LoadCursor(NULL, IDC_SIZE)
		, LoadCursor(NULL, IDC_SIZEALL), LoadCursor(NULL, IDC_SIZENESW), LoadCursor(NULL, IDC_SIZENS)
		, LoadCursor(NULL, IDC_SIZENWSE), LoadCursor(NULL, IDC_SIZEWE), LoadCursor(NULL, IDC_UPARROW)
		, LoadCursor(NULL, IDC_WAIT)}; // If IDC_HAND were added, it would break existing scripts that rely on Unknown being synonymous with Hand.  If ever added, IDC_HAND should return NULL on Win95/NT.
	static const size_t cursor_count = _countof(sCursor);
	// The order in the below array must correspond to the order in the above array:
	static LPTSTR sCursorName[cursor_count + 1] = {_T("AppStarting"), _T("Arrow")
		, _T("Cross"), _T("Help"), _T("IBeam")
		, _T("Icon"), _T("No"), _T("Size")
		, _T("SizeAll"), _T("SizeNESW"), _T("SizeNS")  // NESW = NorthEast+SouthWest
		, _T("SizeNWSE"), _T("SizeWE"), _T("UpArrow")
		, _T("Wait"), CURSOR_UNKNOWN}; // The default item must be last.

	int i;
	for (i = 0; i < cursor_count; ++i)
		if (sCursor[i] == current_cursor)
			break;

	_f_return_p(sCursorName[i]); // If i is out-of-bounds, "Unknown" will be used.
}

BIV_DECL_R(BIV_ScreenWidth_Height)
{
	_f_return_i(GetSystemMetrics(aVarName[13] ? SM_CYSCREEN : SM_CXSCREEN));
}

BIV_DECL_R(BIV_ScriptName)
{
	_f_return_p(g_script.mScriptName ? g_script.mScriptName : g_script.mFileName);
}

BIV_DECL_W(BIV_ScriptName_Set)
{
	// For simplicity, a new buffer is allocated each time.  It is not expected to be set frequently.
	LPTSTR script_name = _tcsdup(BivRValueToString());
	if (!script_name)
		_f_throw_oom;
	free(g_script.mScriptName);
	g_script.mScriptName = script_name;
}

BIV_DECL_R(BIV_ScriptDir)
{
	_f_return_p(g_script.mFileDir);
}

BIV_DECL_R(BIV_ScriptFullPath)
{
	_f_return_p(g_script.mFileSpec);
}

BIV_DECL_R(BIV_ScriptHwnd)
{
	_f_return((UINT_PTR)g_hWnd);
}


LineNumberType Script::CurrentLine()
{
	return mCurrLine ? mCurrLine->mLineNumber : mCombinedLineNumber;
}

BIV_DECL_R(BIV_LineNumber)
{
	_f_return_i(g_script.CurrentLine());
}


LPTSTR Script::CurrentFile()
{
	return Line::sSourceFile[mCurrLine ? mCurrLine->mFileIndex : mCurrFileIndex];
}

BIV_DECL_R(BIV_LineFile)
{
	_f_return_p(g_script.CurrentFile());
}


BIV_DECL_R(BIV_LoopFileName) // Called by multiple callers.
{
	LPTSTR filename = _T("");  // Set default.
	if (g->mLoopFile)
	{
		// cAlternateFileName can be blank if the file lacks a short name, but it can also be blank
		// if the file's proper name complies with all 8.3 requirements (not just length), so use
		// cFileName whenever cAlternateFileName is empty.  GetShortPathName() also behaves this way.
		if (   ctoupper(aVarName[10]) != 'S' // It's not A_LoopFileShortName or ...
			|| !*(filename = g->mLoopFile->cAlternateFileName)   ) // ... there's no alternate name (see above).
			filename = g->mLoopFile->cFileName;
	}
	_f_return_p(filename);
}

BIV_DECL_R(BIV_LoopFileExt)
{
	LPTSTR file_ext = _T("");  // Set default.
	if (g->mLoopFile)
	{
		if (file_ext = _tcsrchr(g->mLoopFile->cFileName, '.'))
			++file_ext;
		else // Reset to empty string vs. NULL.
			file_ext = _T("");
	}
	_f_return_p(file_ext);
}

BIV_DECL_R(BIV_LoopFileDir)
{
	if (!g->mLoopFile)
		_f_return_empty;
	LoopFilesStruct &lfs = *g->mLoopFile;
	LPTSTR dir_end = lfs.file_path + lfs.dir_length; // Start of the filename.
	size_t suffix_length = dir_end - lfs.file_path_suffix; // Directory names\ added since the loop started.
	size_t total_length = lfs.orig_dir_length + suffix_length;
	if (total_length)
		--total_length; // Omit the trailing slash.
	if (!TokenSetResult(aResultToken, nullptr, total_length))
		return;
	aResultToken.symbol = SYM_STRING;
	LPTSTR buf = aResultToken.marker;
	tmemcpy(buf, lfs.orig_dir, lfs.orig_dir_length);
	tmemcpy(buf + lfs.orig_dir_length, lfs.file_path_suffix, suffix_length);
	buf[total_length] = '\0'; // This replaces the final character copied above, if any.
}

void FixLoopFilePath(LPTSTR aBuf, LPTSTR aPattern)
// Fixes aBuf to account for "." and ".." as file patterns.  These match the directory itself
// or parent directory, so for example "x\y\.." returns a directory named "x" which appears to
// be inside "y".  Without the handling here, the invalid path "x\y\x" would be returned.
// A small amount of temporary buffer space might be wasted compared to handling this in the BIV,
// but this way minimizes code size (and these cases are rare anyway).
{
	int count = 0;
	if (*aPattern == '.')
	{
		if (!aPattern[1])
			count = 1; // aBuf "x\y\y" should be "x\y" for "x\y\.".
		else if (aPattern[1] == '.' && !aPattern[2])
			count = 2; // aBuf "x\y\x" should be "x" for "x\y\..".
	}
	for ( ; count > 0; --count)
	{
		LPTSTR end = _tcsrchr(aBuf, '\\');
		if (end)
			*end = '\0';
		else if (*aBuf && aBuf[1] == ':') // aBuf "C:x" should be "C:" for "C:" or "C:.".
			aBuf[2] = '\0';
	}
}

void ReturnLoopFilePath(ResultToken &aResultToken, LPTSTR aPattern, LPTSTR aPrefix, size_t aPrefixLen, LPTSTR aSuffix, size_t aSuffixLen)
{
	if (!TokenSetResult(aResultToken, nullptr, aPrefixLen + aSuffixLen))
		return;
	aResultToken.symbol = SYM_STRING;
	LPTSTR buf = aResultToken.marker;
	tmemcpy(buf, aPrefix, aPrefixLen);
	tmemcpy(buf + aPrefixLen, aSuffix, aSuffixLen + 1); // +1 for \0.
	FixLoopFilePath(buf, aPattern);
	aResultToken.marker_length = -1; // Let caller determine actual length.
}

BIV_DECL_R(BIV_LoopFilePath)
{
	if (!g->mLoopFile)
		_f_return_empty;
	LoopFilesStruct &lfs = *g->mLoopFile;
	// Combine the original directory specified by the script with the dynamic part of file_path
	// (i.e. the sub-directory and file names appended to it since the loop started):
	ReturnLoopFilePath(aResultToken, lfs.pattern
		, lfs.orig_dir, lfs.orig_dir_length
		, lfs.file_path_suffix, lfs.file_path_length - (lfs.file_path_suffix - lfs.file_path));
}

BIV_DECL_R(BIV_LoopFileFullPath)
{
	if (!g->mLoopFile)
		_f_return_empty;
	// GetFullPathName() is done in addition to ConvertFilespecToCorrectCase() for the following reasons:
	// 1) It's currently the only easy way to get the full path of the directory in which a file resides.
	//    For example, if a script is passed a filename via command line parameter, that file could be
	//    either an absolute path or a relative path.  If relative, of course it's relative to A_WorkingDir.
	//    The problem is, the script would have to manually detect this, which would probably take several
	//    extra steps.
	// 2) A_LoopFileLongPath is mostly intended for the following cases, and in all of them it seems
	//    preferable to have the full/absolute path rather than the relative path:
	//    a) Files dragged onto a .ahk script when the drag-and-drop option has been enabled via the Installer.
	//    b) Files passed into the script via command line.
	// Currently both are done by PerformLoopFilePattern(), for performance and in case the working
	// directory changes after the Loop begins.
	LoopFilesStruct &lfs = *g->mLoopFile;
	// Combine long_dir with the dynamic part of file_path:
	ReturnLoopFilePath(aResultToken, lfs.pattern
		, lfs.long_dir, lfs.long_dir_length
		, lfs.file_path_suffix, lfs.file_path_length - (lfs.file_path_suffix - lfs.file_path));
}

BIV_DECL_R(BIV_LoopFileShortPath)
// Unlike GetLoopFileShortName(), this function returns blank when there is no short path.
// This is done so that there's a way for the script to more easily tell the difference between
// an 8.3 name not being available (due to the being disabled in the registry) and the short
// name simply being the same as the long name.  For example, if short name creation is disabled
// in the registry, A_LoopFileShortName would contain the long name instead, as documented.
// But to detect if that short name is really a long name, A_LoopFileShortPath could be checked
// and if it's blank, there is no short name available.
{
	if (!g->mLoopFile)
		_f_return_empty;
	LoopFilesStruct &lfs = *g->mLoopFile;
	// MSDN says cAlternateFileName is empty if the file does not have a long name.
	// Testing and research shows that GetShortPathName() uses the long name for a directory
	// or file if no short name exists, so there's no check for the filename's length here.
	LPTSTR name = *lfs.cAlternateFileName ? lfs.cAlternateFileName : lfs.cFileName;
	ReturnLoopFilePath(aResultToken, lfs.pattern
		, lfs.short_path, lfs.short_path_length
		, name, _tcslen(name));
}

BIV_DECL_R(BIV_LoopFileTime)
{
	LPTSTR target_buf = _f_retval_buf;
	*target_buf = '\0'; // Set default.
	if (g->mLoopFile)
	{
		FILETIME ft;
		switch(ctoupper(aVarName[14])) // A_LoopFileTime[A]ccessed
		{
		case 'M': ft = g->mLoopFile->ftLastWriteTime; break;
		case 'C': ft = g->mLoopFile->ftCreationTime; break;
		default: ft = g->mLoopFile->ftLastAccessTime;
		}
		FileTimeToYYYYMMDD(target_buf, ft, true);
	}
	_f_return_p(target_buf);
}

BIV_DECL_R(BIV_LoopFileAttrib)
{
	LPTSTR target_buf = _f_retval_buf;
	*target_buf = '\0'; // Set default.
	if (g->mLoopFile)
		FileAttribToStr(target_buf, g->mLoopFile->dwFileAttributes);
	_f_return_p(target_buf);
}

BIV_DECL_R(BIV_LoopFileSize)
{
	if (g->mLoopFile)
	{
		ULARGE_INTEGER ul;
		ul.HighPart = g->mLoopFile->nFileSizeHigh;
		ul.LowPart = g->mLoopFile->nFileSizeLow;
		int divider;
		switch (ctoupper(aVarName[14])) // A_LoopFileSize[K/M]B
		{
		case 'K': divider = 1024; break;
		case 'M': divider = 1024*1024; break;
		default:  divider = 0;
		}
		_f_return_i(divider ? ((unsigned __int64)ul.QuadPart / divider) : ul.QuadPart);
	}
	_f_return_empty;
}

BIV_DECL_R(BIV_LoopRegType)
{
	_f_return_p(g->mLoopRegItem ? Line::RegConvertValueType(g->mLoopRegItem->type) : _T(""));
}

BIV_DECL_R(BIV_LoopRegKey)
{
	LPTSTR rootkey = _T("");
	LPTSTR subkey = _T("");
	if (g->mLoopRegItem)
	{
		// Use root_key_type, not root_key (which might be a remote vs. local HKEY):
		rootkey = Line::RegConvertRootKeyType(g->mLoopRegItem->root_key_type);
		subkey = g->mLoopRegItem->subkey;
	}
	if (!TokenSetResult(aResultToken, nullptr, _tcslen(rootkey) + (*subkey != 0) + _tcslen(subkey)))
		return;
	_stprintf(aResultToken.marker, _T("%s%s%s"), rootkey, *subkey ? _T("\\") : _T(""), subkey);
	aResultToken.symbol = SYM_STRING;
}

BIV_DECL_R(BIV_LoopRegName)
{
	// This can be either the name of a subkey or the name of a value.
	_f_return_p(g->mLoopRegItem ? g->mLoopRegItem->name : _T(""));
}

BIV_DECL_R(BIV_LoopRegTimeModified)
{
	LPTSTR target_buf = _f_retval_buf;
	*target_buf = '\0'; // Set default.
	// Only subkeys (not values) have a time.
	if (g->mLoopRegItem && g->mLoopRegItem->type == REG_SUBKEY)
		FileTimeToYYYYMMDD(target_buf, g->mLoopRegItem->ftLastWriteTime, true);
	_f_return_p(target_buf);
}

BIV_DECL_R(BIV_LoopReadLine)
{
	_f_return_p(g->mLoopReadFile ? g->mLoopReadFile->mCurrentLine : _T(""));
}

BIV_DECL_R(BIV_LoopField)
{
	_f_return_p(g->mLoopField ? g->mLoopField : _T(""));
}

BIV_DECL_R(BIV_LoopIndex)
{
	_f_return_i(g->mLoopIteration);
}

BIV_DECL_W(BIV_LoopIndex_Set)
{
	g->mLoopIteration = BivRValueToInt64();
}



BIV_DECL_R(BIV_ThisFunc)
{
	LPCTSTR name;
	if (g->CurrentFunc)
		name = g->CurrentFunc->mName;
	else
		name = _T("");
	_f_return_p(const_cast<LPTSTR>(name));
}

BIV_DECL_R(BIV_ThisHotkey)
{
	_f_return_p(g_script.mThisHotkeyName);
}

BIV_DECL_R(BIV_PriorHotkey)
{
	_f_return_p(g_script.mPriorHotkeyName);
}

BIV_DECL_R(BIV_TimeSinceThisHotkey)
{
	// It must be the type of hotkey that has a label because we want the TimeSinceThisHotkey
	// value to be "in sync" with the value of ThisHotkey itself (i.e. use the same method
	// to determine which hotkey is the "this" hotkey):
	if (*g_script.mThisHotkeyName)
		// Even if GetTickCount()'s TickCount has wrapped around to zero and the timestamp hasn't,
		// DWORD subtraction still gives the right answer as long as the number of days between
		// isn't greater than about 49.  See MyGetTickCount() for explanation of %d vs. %u.
		// Update: Using 64-bit ints now, so above is obsolete:
		//sntprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - g_script.mThisHotkeyStartTime));
		_f_return_i((__int64)(GetTickCount() - g_script.mThisHotkeyStartTime));
	else
		_f_return_empty; // Cause any attempt at math to throw.
}

BIV_DECL_R(BIV_TimeSincePriorHotkey)
{
	if (*g_script.mPriorHotkeyName)
		// See MyGetTickCount() for explanation for explanation:
		//sntprintf(str, sizeof(str), "%d", (DWORD)(GetTickCount() - g_script.mPriorHotkeyStartTime));
		_f_return_i((__int64)(GetTickCount() - g_script.mPriorHotkeyStartTime));
	else
		_f_return_empty; // Cause any attempt at math to throw.
}

BIV_DECL_R(BIV_EndChar)
{
	_f_retval_buf[0] = g_script.mEndChar;
	_f_retval_buf[1] = '\0';
	_f_return_p(_f_retval_buf);
}



BIV_DECL_R(BIV_EventInfo)
// We're returning the length of the var's contents, not the size.
{
	_f_return_i(g->EventInfo);
}

BIV_DECL_W(BIV_EventInfo_Set)
{
	g->EventInfo = (EventInfoType)BivRValueToInt64();
}



BIV_DECL_R(BIV_TimeIdle)
// This is here rather than in script.h with the others because it depends on
// hotkey.h and globaldata.h, which can't be easily included in script.h due to
// mutual dependency issues.
{
	DWORD time_last_input = 0;
	switch (toupper(aVarName[10]))
	{
	case 'M': time_last_input = g_MouseHook ? g_TimeLastInputMouse : 0; break;
	case 'K': time_last_input = g_KeybdHook ? g_TimeLastInputKeyboard : 0; break;
	case 'P': time_last_input = (g_KeybdHook || g_MouseHook) ? g_TimeLastInputPhysical : 0; break;
	}
	// If the relevant hook is not active, default this to the same as the regular idle time:
	if (!time_last_input)
	{
		LASTINPUTINFO lii;
		lii.cbSize = sizeof(lii);
		if (GetLastInputInfo(&lii))
			time_last_input = lii.dwTime;
		else // This is rare; the possibility isn't even documented as of 2020-06-11.
			_f_return_empty; // Cause any attempt at math to throw.
	}
	_f_return_i(GetTickCount() - time_last_input);
}



BIF_DECL(BIF_SetBIV)
{
	static VirtualVar::Setter sBiv[] = { &BIV_DetectHiddenText_Set, &BIV_DetectHiddenWindows_Set, &BIV_FileEncoding_Set, &BIV_RegView_Set, &BIV_StoreCapsLockMode_Set, &BIV_TitleMatchMode_Set };
	auto biv = sBiv[_f_callee_id];
	_f_set_retval_p(_T(""), 0);
	biv(aResultToken, nullptr, *aParam[0]);
}
