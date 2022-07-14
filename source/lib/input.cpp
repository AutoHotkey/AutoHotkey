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

#include "globaldata.h"
#include "application.h"



///////////
// Input //
///////////


ResultType input_type::Setup(LPTSTR aOptions, LPTSTR aEndKeys, LPTSTR aMatchList, size_t aMatchList_length)
{
	ParseOptions(aOptions);
	if (!SetKeyFlags(aEndKeys))
		return FAIL;
	if (!SetMatchList(aMatchList, aMatchList_length))
		return FAIL;
	
	// For maintainability/simplicity/code size, it's allocated even if BufferLengthMax == 0.
	if (  !(Buffer = tmalloc(BufferLengthMax + 1))  )
		return MemoryError();
	*Buffer = '\0';

	return OK;
}


void InputStart(input_type &input)
{
	ASSERT(!input.InProgress());

	// Keep the object alive while it is active, even if the script discards it.
	// The corresponding Release() is done when g_input is reset by InputRelease().
	if (input.ScriptObject)
		input.ScriptObject->AddRef();
	
	// Set or update the timeout timer if needed.  The timer proc takes care to end
	// only those inputs which are due, and will reset or kill the timer as needed.
	if (input.Timeout > 0)
		input.SetTimeoutTimer();

	// It is possible for &input to already be in the list if AHK_INPUT_END is still
	// in the message queue, in which case it must be removed from its current position
	// to prevent the list from looping back on itself.
	InputUnlinkIfStopped(&input);

	input.Prev = g_input;
	input.Start();
	g_input = &input; // Signal the hook to start the input.

	Hotkey::InstallKeybdHook(); // Install the hook (if needed).
}


void input_type::ParseOptions(LPTSTR aOptions)
{
	for (LPTSTR cp = aOptions; *cp; ++cp)
	{
		switch(ctoupper(*cp))
		{
		case 'B':
			BackspaceIsUndo = false;
			break;
		case 'C':
			CaseSensitive = true;
			break;
		case 'I':
			MinSendLevel = (cp[1] <= '9' && cp[1] >= '0') ? (SendLevelType)_ttoi(cp + 1) : 1;
			break;
		case 'M':
			TranscribeModifiedKeys = true;
			break;
		case 'L':
			// Use atoi() vs. ATOI() to avoid interpreting something like 0x01C as hex
			// when in fact the C was meant to be an option letter:
			BufferLengthMax = _ttoi(cp + 1);
			if (BufferLengthMax < 0)
				BufferLengthMax = 0;
			break;
		case 'T':
			// Although ATOF() supports hex, it's been documented in the help file that hex should
			// not be used (see comment above) so if someone does it anyway, some option letters
			// might be misinterpreted:
			Timeout = (int)(ATOF(cp + 1) * 1000);
			break;
		case 'V':
			VisibleText = true;
			VisibleNonText = true;
			break;
		case '*':
			FindAnywhere = true;
			break;
		case 'E':
			// Interpret single-character keys as characters rather than converting them to VK codes.
			// This tends to work better when using multiple keyboard layouts, but changes behaviour:
			// for instance, an end char of "." cannot be triggered while holding Alt.
			EndCharMode = true;
			break;
		}
	}
}


void input_type::SetTimeoutTimer()
{
	DWORD now = GetTickCount();
	TimeoutAt = now + Timeout;
	if (!g_InputTimerExists || Timeout < int(g_InputTimeoutAt - now))
		SET_INPUT_TIMER(Timeout, TimeoutAt)
}


ResultType input_type::SetKeyFlags(LPTSTR aKeys, bool aEndKeyMode, UCHAR aFlagsRemove, UCHAR aFlagsAdd)
{
	bool vk_by_number, sc_by_number;
	vk_type vk;
	sc_type sc = 0;
	modLR_type modifiersLR;
	size_t key_text_length;
	UINT single_char_count = 0;
	TCHAR *end_pos, single_char_string[2];
	single_char_string[1] = '\0'; // Init its second character once, since the loop only changes the first char.
	
	const bool endchar_mode = aEndKeyMode && EndCharMode;
	UCHAR * const end_vk = KeyVK;
	UCHAR * const end_sc = KeySC;

	for (TCHAR *end_key = aKeys; *end_key; ++end_key) // This a modified version of the processing loop used in SendKeys().
	{
		vk = 0; // Set default.  Not strictly necessary but more maintainable.
		*single_char_string = '\0';  // Set default as "this key name is not a single-char string".

		switch (*end_key)
		{
		case '}': continue;  // Important that these be ignored.
		case '{':
		{
			if (   !(end_pos = _tcschr(end_key + 1, '}'))   )
				continue;  // Do nothing, just ignore the unclosed '{' and continue.
			if (   !(key_text_length = end_pos - end_key - 1)   )
			{
				if (end_pos[1] == '}') // The string "{}}" has been encountered, which is interpreted as a single "}".
				{
					++end_pos;
					key_text_length = 1;
				}
				else // Empty braces {} were encountered.
					continue;  // do nothing: let it proceed to the }, which will then be ignored.
			}
			if (key_text_length == 1) // A single-char key name, such as {.} or {{}.
			{
				if (endchar_mode) // Handle this single-char key name by char code, not by VK.
				{
					// Although it might be sometimes useful to treat "x" as a character and "{x}" as a key,
					// "{{}" and "{}}" can't be included without the extra braces.  {vkNN} can still be used
					// to handle the key by VK instead of by character.
					single_char_count++;
					continue; // It will be processed by another section.
				}
				*single_char_string = end_key[1]; // Only used when vk != 0.
			}

			*end_pos = '\0';  // temporarily terminate the string here.

			sc_by_number = false; // Set default.
			modifiersLR = 0;  // Init prior to below.
			// Handle the key by VK if it was given by number, such as {vk26}.
			// Otherwise, for any key name which has a VK shared by two possible SCs
			// (such as Up and NumpadUp), handle it by SC so it's identified correctly.
			if (vk = TextToVK(end_key + 1, &modifiersLR, true))
			{
				vk_by_number = ctoupper(end_key[1]) == 'V' && ctoupper(end_key[2]) == 'K';
				if (!vk_by_number && (sc = vk_to_sc(vk, true)))
				{
					sc ^= 0x100; // Convert sc to the primary scan code, which is the one named by end_key.
					vk = 0; // Handle it only by SC.
				}
			}
			else
				// No virtual key, so try to find a scan code.
				sc = TextToSC(end_key + 1, &sc_by_number);

			*end_pos = '}';  // undo the temporary termination

			end_key = end_pos;  // In prep for ++end_key at the top of the loop.
			break; // Break out of the switch() and do the vk handling beneath it (if there is a vk).
		}

		default:
			if (endchar_mode)
			{
				single_char_count++;
				continue; // It will be processed by another section.
			}
			*single_char_string = *end_key;
			modifiersLR = 0;  // Init prior to below.
			vk = TextToVK(single_char_string, &modifiersLR, true);
			vk_by_number = false;
		} // switch()

		if (vk) // A valid virtual key code was discovered above.
		{
			// Insist the shift key be down to form genuinely different symbols --
			// namely punctuation marks -- but not for alphabetic chars.
			if (*single_char_string && aEndKeyMode && !IsCharAlpha(*single_char_string)) // v1.0.46.05: Added check for "*single_char_string" so that non-single-char strings like {F9} work as end keys even when the Shift key is being held down (this fixes the behavior to be like it was in pre-v1.0.45).
			{
				// Now we know it's not alphabetic, and it's not a key whose name
				// is longer than one char such as a function key or numpad number.
				// That leaves mostly just the number keys (top row) and all
				// punctuation chars, which are the ones that we want to be
				// distinguished between shifted and unshifted:
				if (modifiersLR & (MOD_LSHIFT | MOD_RSHIFT))
					end_vk[vk] |= END_KEY_WITH_SHIFT;
				else
					end_vk[vk] |= END_KEY_WITHOUT_SHIFT;
			}
			else
			{
				end_vk[vk] = (end_vk[vk] & ~aFlagsRemove) | aFlagsAdd;
				// Apply flag removal to this key's SC as well.  This is primarily
				// to support combinations like {All} +E, {LCtrl}{RCtrl} -E.
				sc_type temp_sc;
				if (aFlagsRemove && !vk_by_number && (temp_sc = vk_to_sc(vk)))
				{
					end_sc[temp_sc] &= ~aFlagsRemove; // But apply aFlagsAdd only by VK.
					// Since aFlagsRemove implies ScriptObject != NULL and !vk_by_number
					// was also checked, that implies vk_to_sc(vk, true) was already called
					// and did not find a secondary SC.
				}
			}
		}
		if (sc || sc_by_number) // Fixed for v1.1.33.02: Allow sc000 for setting/unsetting flags for any events that lack a scan code.
		{
			end_sc[sc] = (end_sc[sc] & ~aFlagsRemove) | aFlagsAdd;
		}
	} // for()

	if (single_char_count)  // See single_char_count++ above for comments.
	{
		if (single_char_count > EndCharsMax)
		{
			// Allocate a bigger buffer.
			if (EndCharsMax) // If zero, EndChars may point to static memory.
				free(EndChars);
			if (  !(EndChars = tmalloc(single_char_count + 1))  )
				return MemoryError();
			EndCharsMax = single_char_count;
		}
		TCHAR *dst, *src;
		for (dst = EndChars, src = aKeys; *src; ++src)
		{
			switch (*src)
			{
			case '{':
				if (end_pos = _tcschr(src + 1, '}'))
				{
					if (end_pos == src + 1 && end_pos[1] == '}') // {}}
						end_pos++;
					if (end_pos == src + 2)
						*dst++ = src[1]; // Copy the single character from between the braces.
					src = end_pos; // Skip '{key'.  Loop does ++src to skip the '}'.
				}
				// Otherwise, just ignore the '{'.
			case '}':
				continue;
			}
			*dst++ = *src;
		}
		ASSERT(dst > EndChars);
		*dst = '\0';
	}
	else if (aEndKeyMode) // single_char_count is false
	{
		if (EndCharsMax)
			*EndChars = '\0';
		else
			EndChars = _T("");
	}
	return OK;
}


ResultType input_type::SetMatchList(LPTSTR aMatchList, size_t aMatchList_length)
{
	LPTSTR *realloc_temp;  // Needed since realloc returns NULL on failure but leaves original block allocated.
	MatchCount = 0;  // Set default.
	if (*aMatchList)
	{
		// If needed, create the array of pointers that points into MatchBuf to each match phrase:
		if (!match)
		{
			if (   !(match = (LPTSTR *)malloc(INPUT_ARRAY_BLOCK_SIZE * sizeof(LPTSTR)))   )
				return MemoryError();  // Short msg. since so rare.
			MatchCountMax = INPUT_ARRAY_BLOCK_SIZE;
		}
		// If needed, create or enlarge the buffer that contains all the match phrases:
		size_t space_needed = aMatchList_length + 1;  // +1 for the final zero terminator.
		if (space_needed > MatchBufSize)
		{
			MatchBufSize = (UINT)(space_needed > 4096 ? space_needed : 4096);
			if (MatchBuf) // free the old one since it's too small.
				free(MatchBuf);
			if (   !(MatchBuf = tmalloc(MatchBufSize))   )
			{
				MatchBufSize = 0;
				return MemoryError();  // Short msg. since so rare.
			}
		}
		// Copy aMatchList into the match buffer:
		LPTSTR source, dest;
		for (source = aMatchList, dest = match[MatchCount] = MatchBuf
			; *source; ++source)
		{
			if (*source != ',') // Not a comma, so just copy it over.
			{
				*dest++ = *source;
				continue;
			}
			// Otherwise: it's a comma, which becomes the terminator of the previous key phrase unless
			// it's a double comma, in which case it's considered to be part of the previous phrase
			// rather than the next.
			if (*(source + 1) == ',') // double comma
			{
				*dest++ = *source;
				++source;  // Omit the second comma of the pair, i.e. each pair becomes a single literal comma.
				continue;
			}
			// Otherwise, this is a delimiting comma.
			*dest = '\0';
			// If the previous item is blank -- which I think can only happen now if the MatchList
			// begins with an orphaned comma (since two adjacent commas resolve to one literal comma)
			// -- don't add it to the match list:
			if (*match[MatchCount])
			{
				++MatchCount;
				match[MatchCount] = ++dest;
				*dest = '\0';  // Init to prevent crash on orphaned comma such as "btw,otoh,"
			}
			if (*(source + 1)) // There is a next element.
			{
				if (MatchCount >= MatchCountMax - 1) // Rarely needed, so just realloc() to expand.
				{
					// Expand the array by one block:
					if (   !(realloc_temp = (LPTSTR *)realloc(match  // Must use a temp variable.
						, (MatchCountMax + INPUT_ARRAY_BLOCK_SIZE) * sizeof(LPTSTR)))   )
						return MemoryError();  // Short msg. since so rare.
					match = realloc_temp;
					MatchCountMax += INPUT_ARRAY_BLOCK_SIZE;
				}
			}
		} // for()
		*dest = '\0';  // Terminate the last item.
		// This check is necessary for only a single isolated case: When the match list
		// consists of nothing except a single comma.  See above comment for details:
		if (*match[MatchCount]) // i.e. omit empty strings from the match list.
			++MatchCount;
	}
	return OK;
}


LPTSTR input_type::GetEndReason(LPTSTR aKeyBuf, int aKeyBufSize)
{
	switch (Status)
	{
	case INPUT_TIMED_OUT:
		return _T("Timeout");
	case INPUT_TERMINATED_BY_MATCH:
		return _T("Match");
	case INPUT_TERMINATED_BY_ENDKEY:
	{
		LPTSTR key_name = aKeyBuf;
		if (!key_name)
			return _T("EndKey");
		if (EndingChar)
		{
			key_name[0] = EndingChar;
			key_name[1] = '\0';
		}
		else if (EndingRequiredShift)
		{
			// Since the only way a shift key can be required in our case is if it's a key whose name
			// is a single char (such as a shifted punctuation mark), use a diff. method to look up the
			// key name based on fact that the shift key was down to terminate the input.  We also know
			// that the key is an EndingVK because there's no way for the shift key to have been
			// required by a scan code based on the logic (above) that builds the end_key arrays.
			// MSDN: "Typically, ToAscii performs the translation based on the virtual-key code.
			// In some cases, however, bit 15 of the uScanCode parameter may be used to distinguish
			// between a key press and a key release. The scan code is used for translating ALT+
			// number key combinations.
			BYTE state[256] = { 0 };
			state[VK_SHIFT] |= 0x80; // Indicate that the neutral shift key is down for conversion purposes.
			Get_active_window_keybd_layout // Defines the variable active_window_keybd_layout for use below.
			int count = ToUnicodeOrAsciiEx(EndingVK, vk_to_sc(EndingVK), (PBYTE)&state // Nothing is done about ToAsciiEx's dead key side-effects here because it seems to rare to be worth it (assuming its even a problem).
				, key_name, g_MenuIsVisible ? 1 : 0, active_window_keybd_layout); // v1.0.44.03: Changed to call ToAsciiEx() so that active window's layout can be specified (see hook.cpp for details).
			key_name[count] = '\0';  // Terminate the string.
		}
		else
		{
			*key_name = '\0';
			if (EndingBySC)
				SCtoKeyName(EndingSC, key_name, aKeyBufSize, false);
			if (!*key_name)
				VKtoKeyName(EndingVK, key_name, aKeyBufSize, !EndingBySC);
			if (!*key_name)
				sntprintf(key_name, aKeyBufSize, _T("sc%03X"), EndingSC);
		}
		return _T("EndKey");
	}
	case INPUT_LIMIT_REACHED:
		return _T("Max");
	case INPUT_OFF:
		return _T("Stopped");
	default: // In progress.
		return _T("");
	}
}


void input_type::Start()
{
	ASSERT(!InProgress());
	Status = INPUT_IN_PROGRESS;
}

void input_type::EndByMatch(UINT aMatchIndex)
{
	ASSERT(InProgress());
	EndingMatchIndex = aMatchIndex;
	EndByReason(INPUT_TERMINATED_BY_MATCH);
}

void input_type::EndByKey(vk_type aVK, sc_type aSC, bool aBySC, bool aRequiredShift)
{
	ASSERT(InProgress());
	EndingVK = aVK;
	EndingSC = aSC;
	EndingBySC = aBySC;
	EndingRequiredShift = aRequiredShift;
	EndingChar = 0; // Must be zero if the above are to be used.
	EndByReason(INPUT_TERMINATED_BY_ENDKEY);
}

void input_type::EndByChar(TCHAR aChar)
{
	ASSERT(aChar && InProgress());
	EndingChar = aChar;
	// The other EndKey related fields are ignored when Char is non-zero.
	EndByReason(INPUT_TERMINATED_BY_ENDKEY);
}

void input_type::EndByReason(InputStatusType aReason)
{
	ASSERT(InProgress());
	EndingMods = g_modifiersLR_logical; // Not relevant to all end reasons, but might be useful anyway.
	Status = aReason;

	// It's done this way rather than calling InputRelease() directly...
	// ...so that we can rely on MsgSleep() to create a new thread for the OnEnd event.
	// ...because InputRelease() can't be called by the hook thread.
	// ...because some callers rely on the list not being broken by this call.
	PostMessage(g_hWnd, AHK_INPUT_END, (WPARAM)this, 0);
}


input_type **InputFindLink(input_type *aInput)
{
	if (g_input == aInput)
		return &g_input;
	else
		for (auto *input = g_input; input; input = input->Prev)
			if (input->Prev == aInput)
				return &input->Prev;
	return NULL; // aInput is not valid (faked AHK_INPUT_END message?) or not active.
}


input_type *InputUnlinkIfStopped(input_type *aInput)
{
	if (!aInput)
		return NULL;
	input_type **found = InputFindLink(aInput);
	if (!found)
		return NULL;
	// InProgress can be true if Start() is called while AHK_INPUT_END is in the queue.
	// In such cases, aInput was already moved to a new position in the list and must
	// not be removed yet.
	if (!aInput->InProgress())
	{
		*found = aInput->Prev;
		WaitHookIdle(); // Ensure any pending use of aInput by the hook is finished.
		aInput->Prev = NULL;
	}
	return aInput; // Return non-null to indicate aInput was found in the list and is therefore valid.
}


input_type *InputRelease(input_type *aInput)
{
	if (!InputUnlinkIfStopped(aInput))
		return NULL;

	if (aInput->ScriptObject)
	{
		Hotkey::MaybeUninstallHook();
		if (aInput->ScriptObject->onEnd)
			return aInput; // Return for caller to call OnEnd and Release.
		aInput->ScriptObject->Release();
		// The following is not done because this Release() is only to counteract an AddRef() in
		// InputStart().  ScriptObject != NULL indicates this input_type is actually embedded in
		// the InputObject and as such the link should never be broken until both are deleted.
		//aInput->ScriptObject = NULL;
		g_script.ExitIfNotPersistent(EXIT_EXIT); // In case this InputHook was the only thing keeping the script running.
	}
	return NULL;
}


input_type *InputFind(InputObject *object)
{
	for (auto *input = g_input; input; input = input->Prev)
		if (input->ScriptObject == object)
			return input;
	return NULL;
}
