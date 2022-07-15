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
#include <winioctl.h> // For PREVENT_MEDIA_REMOVAL and CD lock/unlock.
#include "qmath.h" // Used by Transform() [math.h incurs 2k larger code size just for ceil() & floor()]
#include "script.h"
#include "window.h" // for IF_USE_FOREGROUND_WINDOW
#include "application.h" // for MsgSleep()
#include "resources/resource.h"  // For InputBox.
#include "TextIO.h"
#include <Psapi.h> // for GetModuleBaseName.
#include "shlwapi.h" // StrCmpLogicalW

#include "script_func_impl.h"



BIF_DECL(BIF_FormatTime)
// The compressed code size of this function is about 1 KB (2 KB uncompressed), which compares
// favorably to using setlocale()+strftime(), which together are about 8 KB of compressed code
// (setlocale() seems to be needed to put the user's or system's locale into effect for strftime()).
// setlocale() weighs in at about 6.5 KB compressed (14 KB uncompressed).
{
	_f_param_string_opt(aYYYYMMDD, 0);
	_f_param_string_opt(aFormat, 1);

	#define FT_MAX_INPUT_CHARS 2000
	// Input/format length is restricted since it must be translated and expanded into a new format
	// string that uses single quotes around non-alphanumeric characters such as punctuation:
	if (_tcslen(aFormat) > FT_MAX_INPUT_CHARS)
		_f_throw_param(1);

	// Worst case expansion: .d.d.d.d. (9 chars) --> '.'d'.'d'.'d'.'d'.' (19 chars)
	// Buffer below is sized to a little more than twice as big as the largest allowed format,
	// which avoids having to constantly check for buffer overflow while translating aFormat
	// into format_buf:
	#define FT_MAX_OUTPUT_CHARS (2*FT_MAX_INPUT_CHARS + 10)
	TCHAR format_buf[FT_MAX_OUTPUT_CHARS + 1];
	TCHAR output_buf[FT_MAX_OUTPUT_CHARS + 1]; // The size of this is somewhat arbitrary, but buffer overflow is checked so it's safe.

	TCHAR yyyymmdd[256]; // Large enough to hold date/time and any options that follow it (note that D and T options can appear multiple times).
	*yyyymmdd = '\0';

	SYSTEMTIME st;
	LPTSTR options = NULL;

	if (!*aYYYYMMDD) // Use current local time by default.
		GetLocalTime(&st);
	else
	{
		tcslcpy(yyyymmdd, omit_leading_whitespace(aYYYYMMDD), _countof(yyyymmdd)); // Make a modifiable copy.
		if (*yyyymmdd < '0' || *yyyymmdd > '9') // First character isn't a digit, therefore...
		{
			// ... options are present without date (since yyyymmdd [if present] must come before options).
			options = yyyymmdd;
			GetLocalTime(&st);  // Use current local time by default.
		}
		else // Since the string starts with a digit, rules say it must be a YYYYMMDD string, possibly followed by options.
		{
			// Find first space or tab because YYYYMMDD portion might contain only the leading part of date/timestamp.
			if (options = StrChrAny(yyyymmdd, _T(" \t"))) // Find space or tab.
			{
				*options = '\0'; // Terminate yyyymmdd at the end of the YYYYMMDDHH24MISS string.
				options = omit_leading_whitespace(++options); // Point options to the right place (can be empty string).
			}
			//else leave options set to NULL to indicate that there are none.

			// Pass "false" for validation so that times can still be reported even if the year
			// is prior to 1601.  If the time and/or date is invalid, GetTimeFormat() and GetDateFormat()
			// will refuse to produce anything, which is documented behavior:
			if (!YYYYMMDDToSystemTime(yyyymmdd, st, false))
				_f_throw_param(0);
		}
	}
	
	// Set defaults.  Some can be overridden by options (if there are any options).
	LCID lcid = LOCALE_USER_DEFAULT;
	DWORD date_flags = 0, time_flags = 0;
	bool date_flags_specified = false, time_flags_specified = false, reverse_date_time = false;
	#define FT_FORMAT_NONE 0
	#define FT_FORMAT_TIME 1
	#define FT_FORMAT_DATE 2
	int format_type1 = FT_FORMAT_NONE;
	LPTSTR format2_marker = NULL; // Will hold the location of the first char of the second format (if present).
	bool do_null_format2 = false;  // Will be changed to true if a default date *and* time should be used.

	if (options) // Parse options.
	{
		LPTSTR option_end;
		TCHAR orig_char;
		for (LPTSTR next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
		{
			// Find the end of this option item:
			if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
				option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.

			// Permanently terminate in between options to help eliminate ambiguity for words contained
			// inside other words, and increase confidence in decimal and hexadecimal conversion.
			orig_char = *option_end;
			*option_end = '\0';

			++next_option;
			switch (_totupper(next_option[-1]))
			{
			case 'D':
				date_flags_specified = true;
				date_flags |= ATOU(next_option); // ATOU() for unsigned.
				break;
			case 'T':
				time_flags_specified = true;
				time_flags |= ATOU(next_option); // ATOU() for unsigned.
				break;
			case 'R':
				reverse_date_time = true;
				break;
			case 'L':
				lcid = !_tcsicmp(next_option, _T("Sys")) ? LOCALE_SYSTEM_DEFAULT : (LCID)ATOU(next_option);
				break;
			// If not one of the above, such as zero terminator or a number, just ignore it.
			}

			*option_end = orig_char; // Undo the temporary termination so that loop's omit_leading() will work.
		} // for() each item in option list
	} // Parse options.

	if (!*aFormat)
	{
		aFormat = NULL; // Tell GetDateFormat() and GetTimeFormat() to use default for the specified locale.
		if (!date_flags_specified) // No preference was given, so use long (which seems generally more useful).
			date_flags |= DATE_LONGDATE;
		if (!time_flags_specified)
			time_flags |= TIME_NOSECONDS;  // Seems more desirable/typical to default to no seconds.
		// Put the time first by default, though this is debatable (Metapad does it and I like it).
		format_type1 = reverse_date_time ? FT_FORMAT_DATE : FT_FORMAT_TIME;
		do_null_format2 = true;
	}
	else // aFormat is non-blank.
	{
		// Omit whitespace only for consideration of special keywords.  Whitespace is later kept for
		// a normal format string such as " MM/dd/yy":
		LPTSTR candidate = omit_leading_whitespace(aFormat);
		if (!_tcsicmp(candidate, _T("YWeek")))
		{
			GetISOWeekNumber(output_buf, st.wYear, GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear)), st.wDayOfWeek);
			_f_return(output_buf);
		}
		if (!_tcsicmp(candidate, _T("YDay")) || !_tcsicmp(candidate, _T("YDay0")))
		{
			int yday = GetYDay(st.wMonth, st.wDay, IS_LEAP_YEAR(st.wYear));
			// Format with leading zeroes if YDay0 was used.
			_stprintf(output_buf, candidate[4] == '0' ? _T("%03d") : _T("%d"), yday);
			_f_return(output_buf);
		}
		if (!_tcsicmp(candidate, _T("WDay")))
		{
			LPTSTR buf = _f_retval_buf;
			*buf = '1' + st.wDayOfWeek; // '1' vs '0' to convert to 1-based for compatibility with A_WDay.
			buf[1] = '\0';
			_f_return_p(buf, 1);
		}

		// Since above didn't return, check for those that require a call to GetTimeFormat/GetDateFormat
		// further below:
		if (!_tcsicmp(candidate, _T("ShortDate")))
		{
			aFormat = NULL;
			date_flags |= DATE_SHORTDATE;
			date_flags &= ~(DATE_LONGDATE | DATE_YEARMONTH); // If present, these would prevent it from working.
		}
		else if (!_tcsicmp(candidate, _T("LongDate")))
		{
			aFormat = NULL;
			date_flags |= DATE_LONGDATE;
			date_flags &= ~(DATE_SHORTDATE | DATE_YEARMONTH); // If present, these would prevent it from working.
		}
		else if (!_tcsicmp(candidate, _T("YearMonth")))
		{
			aFormat = NULL;
			date_flags |= DATE_YEARMONTH;
			date_flags &= ~(DATE_SHORTDATE | DATE_LONGDATE); // If present, these would prevent it from working.
		}
		else if (!_tcsicmp(candidate, _T("Time")))
		{
			format_type1 = FT_FORMAT_TIME;
			aFormat = NULL;
			if (!time_flags_specified)
				time_flags |= TIME_NOSECONDS;  // Seems more desirable/typical to default to no seconds.
		}
		else // Assume normal format string.
		{
			LPTSTR cp = aFormat, dp = format_buf;   // Initialize source and destination pointers.
			bool inside_their_quotes = false; // Whether we are inside a single-quoted string in the source.
			bool inside_our_quotes = false;   // Whether we are inside a single-quoted string of our own making in dest.
			for (; *cp; ++cp) // Transcribe aFormat into format_buf and also check for which comes first: date or time.
			{
				if (*cp == '\'') // Note that '''' (four consecutive quotes) is a single literal quote, which this logic handles okay.
				{
					if (inside_our_quotes)
					{
						// Upon encountering their quotes while we're still in ours, merge theirs with ours and
						// remark it as theirs.  This is done to avoid having two back-to-back quoted sections,
						// which would result in an unwanted literal single quote.  Example:
						// 'Some string'':' (the two quotes in the middle would be seen as a literal quote).
						inside_our_quotes = false;
						inside_their_quotes = true;
						continue;
					}
					if (inside_their_quotes)
					{
						// If next char needs to be quoted, don't close out this quote section because that
						// would introduce two consecutive quotes, which would be interpreted as a single
						// literal quote if its enclosed by two outer single quotes.  Instead convert this
						// quoted section over to "ours":
						if (cp[1] && !IsCharAlphaNumeric(cp[1]) && cp[1] != '\'') // Also consider single quotes to be theirs due to this example: dddd:''''y
							inside_our_quotes = true;
							// And don't do "*dp++ = *cp"
						else // there's no next-char or it's alpha-numeric, so it doesn't need to be inside quotes.
							*dp++ = *cp; // Close out their quoted section.
					}
					else // They're starting a new quoted section, so just transcribe this single quote as-is.
						*dp++ = *cp;
					inside_their_quotes = !inside_their_quotes; // Must be done after the above.
					continue;
				}
				// Otherwise, it's not a single quote.
				if (inside_their_quotes) // *cp is inside a single-quoted string, so it can be part of format/picture
					*dp++ = *cp; // Transcribe as-is.
				else
				{
					if (IsCharAlphaNumeric(*cp))
					{
						if (inside_our_quotes)
						{
							*dp++ = '\''; // Close out the previous quoted section, since this char should not be a part of it.
							inside_our_quotes = false;
						}
						if (_tcschr(_T("dMyg"), *cp)) // A format unique to Date is present.
						{
							if (!format_type1)
								format_type1 = FT_FORMAT_DATE;
							else if (format_type1 == FT_FORMAT_TIME && !format2_marker) // type2 should only be set if different than type1.
							{
								*dp++ = '\0';  // Terminate the first section and (below) indicate that there's a second.
								format2_marker = dp;  // Point it to the location in format_buf where the split should occur.
							}
						}
						else if (_tcschr(_T("hHmst"), *cp)) // A format unique to Time is present.
						{
							if (!format_type1)
								format_type1 = FT_FORMAT_TIME;
							else if (format_type1 == FT_FORMAT_DATE && !format2_marker) // type2 should only be set if different than type1.
							{
								*dp++ = '\0';  // Terminate the first section and (below) indicate that there's a second.
								format2_marker = dp;  // Point it to the location in format_buf where the split should occur.
							}
						}
						// For consistency, transcribe all AlphaNumeric chars not inside single quotes as-is
						// (numbers are transcribed in case they are ever used as part of pic/format).
						*dp++ = *cp;
					}
					else // Not alphanumeric, so enclose this and any other non-alphanumeric characters in single quotes.
					{
						if (!inside_our_quotes)
						{
							*dp++ = '\''; // Create a new quoted section of our own, since this char should be inside quotes to be understood.
							inside_our_quotes = true;
						}
						*dp++ = *cp;  // Add this character between the quotes, since it's of the right "type".
					}
				}
			} // for()
			if (inside_our_quotes)
				*dp++ = '\'';  // Close out our quotes.
			*dp = '\0'; // Final terminator.
			aFormat = format_buf; // Point it to the freshly translated format string, for use below.
		} // aFormat contains normal format/pic string.
	} // aFormat isn't blank.

	// If there are no date or time formats present, still do the transcription so that
	// any quoted strings and whatnot are resolved.  This increases runtime flexibility.
	// The below is also relied upon by "LongDate" and "ShortDate" above:
	if (!format_type1)
		format_type1 = FT_FORMAT_DATE;

	// MSDN: Time: "The function checks each of the time values to determine that it is within the
	// appropriate range of values. If any of the time values are outside the correct range, the
	// function fails, and sets the last-error to ERROR_INVALID_PARAMETER. 
	// Dates: "...year, month, day, and day of week. If the day of the week is incorrect, the
	// function uses the correct value, and returns no error. If any of the other date values
	// are outside the correct range, the function fails, and sets the last-error to ERROR_INVALID_PARAMETER.

	if (format_type1 == FT_FORMAT_DATE) // DATE comes first.
	{
		if (!GetDateFormat(lcid, date_flags, &st, aFormat, output_buf, FT_MAX_OUTPUT_CHARS))
			*output_buf = '\0';  // Ensure it's still the empty string, then try to continue to get the second half (if there is one).
	}
	else // TIME comes first.
		if (!GetTimeFormat(lcid, time_flags, &st, aFormat, output_buf, FT_MAX_OUTPUT_CHARS))
			*output_buf = '\0';  // Ensure it's still the empty string, then try to continue to get the second half (if there is one).

	if (format2_marker || do_null_format2) // There is also a second format present.
	{
		size_t output_buf_length = _tcslen(output_buf);
		LPTSTR output_buf_marker = output_buf + output_buf_length;
		LPTSTR format2;
		if (do_null_format2)
		{
			format2 = NULL;
			*output_buf_marker++ = ' '; // Provide a space between time and date.
			++output_buf_length;
		}
		else
			format2 = format2_marker;

		int buf_remaining_size = (int)(FT_MAX_OUTPUT_CHARS - output_buf_length);
		int result;

		if (format_type1 == FT_FORMAT_DATE) // DATE came first, so this one is TIME.
			result = GetTimeFormat(lcid, time_flags, &st, format2, output_buf_marker, buf_remaining_size);
		else
			result = GetDateFormat(lcid, date_flags, &st, format2, output_buf_marker, buf_remaining_size);
		if (!result)
			output_buf[output_buf_length] = '\0'; // Ensure the first part is still terminated and just return that rather than nothing.
	}

	_f_return(output_buf);
}



BIF_DECL(BIF_StrCase)
{
	size_t length;
	LPTSTR contents = ParamIndexToString(0, _f_number_buf, &length);

	// Make a modifiable copy of the string to return:
	if (!TokenSetResult(aResultToken, contents, length))
		return;
	aResultToken.symbol = SYM_STRING;
	contents = aResultToken.marker;

	if (_f_callee_id == FID_StrLower)
		CharLower(contents);
	else if (_f_callee_id == FID_StrUpper)
		CharUpper(contents);
	else // Convert to title case.
		StrToTitleCase(contents);
}



BIF_DECL(BIF_StrReplace)
{
	size_t length; // Going in to StrReplace(), it's the haystack length. Later (coming out), it's the result length. 
	_f_param_string(source, 0, &length);
	_f_param_string(oldstr, 1);
	_f_param_string_opt(newstr, 2);

	// Maintain this together with the equivalent section of BIF_InStr:
	StringCaseSenseType string_case_sense = ParamIndexToCaseSense(3);
	if (string_case_sense == SCS_INVALID 
		|| string_case_sense == SCS_INSENSITIVE_LOGICAL) // Not supported, seems more useful to throw rather than using SCS_INSENSITIVE.
		_f_throw_param(3);

	Var *output_var_count = ParamIndexToOutputVar(4); 
	UINT replacement_limit = (UINT)ParamIndexToOptionalInt64(5, UINT_MAX); 
	

	// Note: The current implementation of StrReplace() should be able to handle any conceivable inputs
	// without an empty string causing an infinite loop and without going infinite due to finding the
	// search string inside of newly-inserted replace strings (e.g. replacing all occurrences
	// of b with bcd would not keep finding b in the newly inserted bcd, infinitely).
	LPTSTR dest;
	UINT found_count = StrReplace(source, oldstr, newstr, string_case_sense, replacement_limit, -1, &dest, &length); // Length of haystack is passed to improve performance because TokenToString() can often discover it instantaneously.

	if (!dest) // Failure due to out of memory.
		_f_throw_oom; 

	if (dest != source) // StrReplace() allocated new memory rather than returning "source" to us unaltered.
	{
		// Return the newly allocated memory directly to our caller. 
		aResultToken.AcceptMem(dest, length);
	}
	else // StrReplace gave us back "source" unaltered because no replacements were needed.
	{
		_f_set_retval_p(dest, length); 
	}

	if (output_var_count)
		output_var_count->Assign((DWORD)found_count);
	_f_return_retval;
}



BIF_DECL(BIF_StrSplit)
// Array := StrSplit(String [, Delimiters, OmitChars, MaxParts])
{
	LPTSTR aInputString = ParamIndexToString(0, _f_number_buf);
	LPTSTR *aDelimiterList = NULL;
	int aDelimiterCount = 0;
	LPTSTR aOmitList = _T("");
	int splits_left = -2;

	if (aParamCount > 1)
	{
		if (auto arr = dynamic_cast<Array *>(TokenToObject(*aParam[1])))
		{
			aDelimiterCount = arr->Length();
			aDelimiterList = (LPTSTR *)_alloca(aDelimiterCount * sizeof(LPTSTR *));
			if (!arr->ToStrings(aDelimiterList, aDelimiterCount, aDelimiterCount))
				// Array contains something other than a string.
				goto throw_invalid_delimiter;
			for (int i = 0; i < aDelimiterCount; ++i)
				if (!*aDelimiterList[i])
					// Empty string in delimiter list. Although it could be treated similarly to the
					// "no delimiter" case, it's far more likely to be an error. If ever this check
					// is removed, the loop below must be changed to support "" as a delimiter.
					goto throw_invalid_delimiter;
		}
		else
		{
			aDelimiterList = (LPTSTR *)_alloca(sizeof(LPTSTR *));
			*aDelimiterList = TokenToString(*aParam[1]);
			aDelimiterCount = **aDelimiterList != '\0'; // i.e. non-empty string.
		}
		if (aParamCount > 2)
		{
			aOmitList = TokenToString(*aParam[2]);
			if (aParamCount > 3)
				splits_left = (int)TokenToInt64(*aParam[3]) - 1;
		}
	}
	
	auto output_array = Array::Create();
	if (!output_array)
		goto throw_outofmem;

	if (!*aInputString // The input variable is blank, thus there will be zero elements.
		|| splits_left == -1) // The caller specified 0 parts.
		_f_return(output_array);
	
	LPTSTR contents_of_next_element, delimiter, new_starting_pos;
	size_t element_length, delimiter_length;

	if (aDelimiterCount) // The user provided a list of delimiters, so process the input variable normally.
	{
		for (contents_of_next_element = aInputString; ; )
		{
			if (   !splits_left // Limit reached.
				|| !(delimiter = InStrAny(contents_of_next_element, aDelimiterList, aDelimiterCount, delimiter_length))   ) // No delimiter found.
				break; // This is the only way out of the loop other than critical errors.
			element_length = delimiter - contents_of_next_element;
			if (*aOmitList && element_length > 0)
			{
				contents_of_next_element = omit_leading_any(contents_of_next_element, aOmitList, element_length);
				element_length = delimiter - contents_of_next_element; // Update in case above changed it.
				if (element_length)
					element_length = omit_trailing_any(contents_of_next_element, aOmitList, delimiter - 1);
			}
			// If there are no chars to the left of the delim, or if they were all in the list of omitted
			// chars, the variable will be assigned the empty string:
			if (!output_array->Append(contents_of_next_element, element_length))
				goto outofmem;
			contents_of_next_element = delimiter + delimiter_length;  // Omit the delimiter since it's never included in contents.
			if (splits_left > 0)
				--splits_left;
		}
	}
	else
	{
		// Otherwise aDelimiterList is empty, so store each char of aInputString in its own array element.
		LPTSTR cp, dp;
		for (cp = aInputString; ; ++cp)
		{
			if (!*cp)
				_f_return(output_array); // All done.
			for (dp = aOmitList; *dp; ++dp)
				if (*cp == *dp) // This char is a member of the omitted list, thus it is not included in the output array.
					break; // (inner loop)
			if (*dp) // Omitted.
				continue;
			if (!splits_left) // Limit reached (checked only after excluding omitted chars).
				break; // This is the only way out of the loop other than critical errors.
			if (splits_left > 0)
				--splits_left;
			if (!output_array->Append(cp, 1))
				goto outofmem;
		}
		contents_of_next_element = cp;
	}
	// Since above used break rather than goto or return, either the limit was reached or there are
	// no more delimiters, so store the remainder of the string minus any characters to be omitted.
	element_length = _tcslen(contents_of_next_element);
	if (*aOmitList && element_length > 0)
	{
		new_starting_pos = omit_leading_any(contents_of_next_element, aOmitList, element_length);
		element_length -= (new_starting_pos - contents_of_next_element); // Update in case above changed it.
		contents_of_next_element = new_starting_pos;
		if (element_length)
			// If this is true, the string must contain at least one char that isn't in the list
			// of omitted chars, otherwise omit_leading_any() would have already omitted them:
			element_length = omit_trailing_any(contents_of_next_element, aOmitList
				, contents_of_next_element + element_length - 1);
	}
	// If there are no chars to the left of the delim, or if they were all in the list of omitted
	// chars, the item will be an empty string:
	if (output_array->Append(contents_of_next_element, element_length))
		_f_return(output_array); // All done.
	//else memory allocation failed, so fall through:
outofmem:
	// The fact that this section is executing means that a memory allocation failed and caused the
	// loop to break, so throw an exception.
	output_array->Release(); // Since we're not returning it.
	// Below: Using goto probably won't reduce code size (due to compiler optimizations), but keeping
	// rarely executed code at the end of the function might help performance due to cache utilization.
throw_outofmem:
	// Either goto was used or FELL THROUGH FROM ABOVE.
	_f_throw_oom;
throw_invalid_delimiter:
	_f_throw_param(1);
}



ResultType SplitPath(LPCTSTR aFileSpec, Var *output_var_name, Var *output_var_dir, Var *output_var_ext, Var *output_var_name_no_ext, Var *output_var_drive)
{
	// For URLs, "drive" is defined as the server name, e.g. http://somedomain.com
	LPCTSTR name = _T(""), name_delimiter = NULL, drive_end = NULL; // Set defaults to improve maintainability.
	LPCTSTR drive = omit_leading_whitespace(aFileSpec); // i.e. whitespace is considered for everything except the drive letter or server name, so that a pathless filename can have leading whitespace.
	LPCTSTR colon_double_slash = _tcsstr(aFileSpec, _T("://"));

	if (colon_double_slash) // This is a URL such as ftp://... or http://...
	{
		if (   !(drive_end = _tcschr(colon_double_slash + 3, '/'))   )
		{
			if (   !(drive_end = _tcschr(colon_double_slash + 3, '\\'))   ) // Try backslash so that things like file://C:\Folder\File.txt are supported.
				drive_end = colon_double_slash + _tcslen(colon_double_slash); // Set it to the position of the zero terminator instead.
				// And because there is no filename, leave name and name_delimiter set to their defaults.
			//else there is a backslash, e.g. file://C:\Folder\File.txt, so treat that backslash as the end of the drive name.
		}
		name_delimiter = drive_end; // Set default, to be possibly overridden below.
		// Above has set drive_end to one of the following:
		// 1) The slash that occurs to the right of the doubleslash in a URL.
		// 2) The backslash that occurs to the right of the doubleslash in a URL.
		// 3) The zero terminator if there is no slash or backslash to the right of the doubleslash.
		if (*drive_end) // A slash or backslash exists to the right of the server name.
		{
			if (*(drive_end + 1))
			{
				// Find the rightmost slash.  At this stage, this is known to find the correct slash.
				// In the case of a file at the root of a domain such as http://domain.com/root_file.htm,
				// the directory consists of only the domain name, e.g. http://domain.com.  This is because
				// the directory always the "drive letter" by design, since that is more often what the
				// caller wants.  A script can use StringReplace to remove the drive/server portion from
				// the directory, if desired.
				name_delimiter = _tcsrchr(aFileSpec, '/');
				if (name_delimiter == colon_double_slash + 2) // To reach this point, it must have a backslash, something like file://c:\folder\file.txt
					name_delimiter = _tcsrchr(aFileSpec, '\\'); // Will always be found.
				name = name_delimiter + 1; // This will be the empty string for something like http://domain.com/dir/
			}
			//else something like http://domain.com/, so leave name and name_delimiter set to their defaults.
		}
		//else something like http://domain.com, so leave name and name_delimiter set to their defaults.
	}
	else // It's not a URL, just a file specification such as c:\my folder\my file.txt, or \\server01\folder\file.txt
	{
		// Differences between _splitpath() and the method used here:
		// _splitpath() doesn't include drive in output_var_dir, it includes a trailing
		// backslash, it includes the . in the extension, it considers ":" to be a filename.
		// _splitpath(pathname, drive, dir, file, ext);
		//char sdrive[16], sdir[MAX_PATH], sname[MAX_PATH], sext[MAX_PATH];
		//_splitpath(aFileSpec, sdrive, sdir, sname, sext);
		//if (output_var_name_no_ext)
		//	output_var_name_no_ext->Assign(sname);
		//strcat(sname, sext);
		//if (output_var_name)
		//	output_var_name->Assign(sname);
		//if (output_var_dir)
		//	output_var_dir->Assign(sdir);
		//if (output_var_ext)
		//	output_var_ext->Assign(sext);
		//if (output_var_drive)
		//	output_var_drive->Assign(sdrive);
		//return OK;

		// Don't use _splitpath() since it supposedly doesn't handle UNC paths correctly,
		// and anyway we need more info than it provides.  Also note that it is possible
		// for a file to begin with space(s) or a dot (if created programmatically), so
		// don't trim or omit leading space unless it's known to be an absolute path.

		// Note that "C:Some File.txt" is a valid filename in some contexts, which the below
		// tries to take into account.  However, there will be no way for this command to
		// return a path that differentiates between "C:Some File.txt" and "C:\Some File.txt"
		// since the first backslash is not included with the returned path, even if it's
		// the root directory (i.e. "C:" is returned in both cases).  The "C:Filename"
		// convention is pretty rare, and anyway this trait can be detected via something like
		// IfInString, Filespec, :, IfNotInString, Filespec, :\, MsgBox Drive with no absolute path.

		// UNCs are detected with this approach so that double sets of backslashes -- which sometimes
		// occur by accident in "built filespecs" and are tolerated by the OS -- are not falsely
		// detected as UNCs.
		if (drive[0] == '\\' && drive[1] == '\\') // Relies on short-circuit evaluation order.
		{
			if (   !(drive_end = _tcschr(drive + 2, '\\'))   )
				drive_end = drive + _tcslen(drive); // Set it to the position of the zero terminator instead.
		}
		else if (*(drive + 1) == ':') // It's an absolute path.
			// Assign letter and colon for consistency with server naming convention above.
			// i.e. so that server name and drive can be used without having to worry about
			// whether it needs a colon added or not.
			drive_end = drive + 2;
		else
		{
			// It's debatable, but it seems best to return a blank drive if a aFileSpec is a relative path.
			// rather than trying to use GetFullPathName() on a potentially non-existent file/dir.
			// _splitpath() doesn't fetch the drive letter of relative paths either.  This also reports
			// a blank drive for something like file://C:\My Folder\My File.txt, which seems too rarely
			// to justify a special mode.
			drive_end = _T("");
			drive = drive_end; // This is necessary to allow Assign() to work correctly later below, since it interprets a length of zero as "use string's entire length".
		}

		if (   !(name_delimiter = _tcsrchr(aFileSpec, '\\'))   ) // No backslash.
			if (   !(name_delimiter = _tcsrchr(aFileSpec, ':'))   ) // No colon.
				name_delimiter = NULL; // Indicate that there is no directory.

		name = name_delimiter ? name_delimiter + 1 : aFileSpec; // If no delimiter, name is the entire string.
	}

	// The above has now set the following variables:
	// name: As an empty string or the actual name of the file, including extension.
	// name_delimiter: As NULL if there is no directory, otherwise, the end of the directory's name.
	// drive: As the start of the drive/server name, e.g. C:, \\Workstation01, http://domain.com, etc.
	// drive_end: As the position after the drive's last character, either a zero terminator, slash, or backslash.

	if (output_var_name && !output_var_name->Assign(name))
		return FAIL;

	if (output_var_dir)
	{
		if (!name_delimiter)
			output_var_dir->Assign(); // Shouldn't fail.
		else if (*name_delimiter == '\\' || *name_delimiter == '/')
		{
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec)))
				return FAIL;
		}
		else // *name_delimiter == ':', e.g. "C:Some File.txt".  If aFileSpec starts with just ":",
			 // the dir returned here will also start with just ":" since that's rare & illegal anyway.
			if (!output_var_dir->Assign(aFileSpec, (VarSizeType)(name_delimiter - aFileSpec + 1)))
				return FAIL;
	}

	LPCTSTR ext_dot = _tcsrchr(name, '.');
	if (output_var_ext)
	{
		// Note that the OS doesn't allow filenames to end in a period.
		if (!ext_dot)
			output_var_ext->Assign();
		else
			if (!output_var_ext->Assign(ext_dot + 1)) // Can be empty string if filename ends in just a dot.
				return FAIL;
	}

	if (output_var_name_no_ext && !output_var_name_no_ext->Assign(name, (VarSizeType)(ext_dot ? ext_dot - name : _tcslen(name))))
		return FAIL;

	if (output_var_drive && !output_var_drive->Assign(drive, (VarSizeType)(drive_end - drive)))
		return FAIL;

	return OK;
}

BIF_DECL(BIF_SplitPath)
{
	LPTSTR mem_to_free = nullptr;
	_f_param_string(aFileSpec, 0);
	Var *vars[6];
	for (int i = 1; i < _countof(vars); ++i)
		vars[i] = ParamIndexToOutputVar(i);
	if (aParam[0]->symbol == SYM_VAR) // Check for overlap of input/output vars.
	{
		vars[0] = aParam[0]->var;
		// There are cases where this could be avoided, such as by careful ordering of the assignments
		// in SplitPath(), or when there's only one output var.  Also, real paths are generally short
		// enough that stack memory could be used.  However, perhaps simple is best in this case.
		for (int i = 1; i < _countof(vars); ++i)
		{
			if (vars[i] == vars[0])
			{
				aFileSpec = mem_to_free = _tcsdup(aFileSpec);
				if (!mem_to_free)
					_f_throw_oom;
				break;
			}
		}
	}
	aResultToken.SetValue(_T(""), 0);
	if (!SplitPath(aFileSpec, vars[1], vars[2], vars[3], vars[4], vars[5]))
		aResultToken.SetExitResult(FAIL);
	free(mem_to_free);
}



int SortWithOptions(const void *a1, const void *a2)
// Decided to just have one sort function since there are so many permutations.  The performance
// will be a little bit worse, but it seems simpler to implement and maintain.
// This function's input parameters are pointers to the elements of the array.  Since those elements
// are themselves pointers, the input parameters are therefore pointers to pointers (handles).
{
	LPTSTR sort_item1 = *(LPTSTR *)a1;
	LPTSTR sort_item2 = *(LPTSTR *)a2;
	if (g_SortColumnOffset > 0)
	{
		// Adjust each string (even for numerical sort) to be the right column position,
		// or the position of its zero terminator if the column offset goes beyond its length:
		size_t length = _tcslen(sort_item1);
		sort_item1 += (size_t)g_SortColumnOffset > length ? length : g_SortColumnOffset;
		length = _tcslen(sort_item2);
		sort_item2 += (size_t)g_SortColumnOffset > length ? length : g_SortColumnOffset;
	}
	if (g_SortNumeric) // Takes precedence over g_SortCaseSensitive
	{
		// For now, assume both are numbers.  If one of them isn't, it will be sorted as a zero.
		// Thus, all non-numeric items should wind up in a sequential, unsorted group.
		// Resolve only once since parts of the ATOF() macro are inline:
		double item1_minus_2 = ATOF(sort_item1) - ATOF(sort_item2);
		if (!item1_minus_2) // Exactly equal.
			return (sort_item1 > sort_item2) ? 1 : -1; // Stable sort.
		// Otherwise, it's either greater or less than zero:
		int result = (item1_minus_2 > 0.0) ? 1 : -1;
		return g_SortReverse ? -result : result;
	}
	// Otherwise, it's a non-numeric sort.
	// v1.0.43.03: Added support the new locale-insensitive mode.
	int result = (g_SortCaseSensitive != SCS_INSENSITIVE_LOGICAL)
		? tcscmp2(sort_item1, sort_item2, g_SortCaseSensitive) // Resolve large macro only once for code size reduction.
		: StrCmpLogicalW(sort_item1, sort_item2);
	if (!result)
		result = (sort_item1 > sort_item2) ? 1 : -1; // Stable sort.
	return g_SortReverse ? -result : result;
}



int SortByNakedFilename(const void *a1, const void *a2)
// See comments in prior function for details.
{
	LPTSTR sort_item1 = *(LPTSTR *)a1;
	LPTSTR sort_item2 = *(LPTSTR *)a2;
	LPTSTR cp;
	if (cp = _tcsrchr(sort_item1, '\\'))  // Assign
		sort_item1 = cp + 1;
	if (cp = _tcsrchr(sort_item2, '\\'))  // Assign
		sort_item2 = cp + 1;
	// v1.0.43.03: Added support the new locale-insensitive mode.
	int result = (g_SortCaseSensitive != SCS_INSENSITIVE_LOGICAL)
		? tcscmp2(sort_item1, sort_item2, g_SortCaseSensitive) // Resolve large macro only once for code size reduction.
		: StrCmpLogicalW(sort_item1, sort_item2);
	if (!result)
		result = (sort_item1 > sort_item2) ? 1 : -1; // Stable sort.
	return g_SortReverse ? -result : result;
}



struct sort_rand_type
{
	LPTSTR cp; // This must be the first member of the struct, otherwise the array trickery in PerformSort will fail.
	union
	{
		// This must be the same size in bytes as the above, which is why it's maintained as a union with
		// a char* rather than a plain int.
		LPTSTR unused;
		int rand;
	};
};

int SortRandom(const void *a1, const void *a2)
// See comments in prior functions for details.
{
	return ((sort_rand_type *)a1)->rand - ((sort_rand_type *)a2)->rand;
}

int SortUDF(const void *a1, const void *a2)
// See comments in prior function for details.
{
	if (!g_SortFuncResult || g_SortFuncResult == EARLY_EXIT)
		return 0;

	// The following isn't necessary because by definition, the current thread isn't paused because it's the
	// thing that called the sort in the first place.
	//g_script.UpdateTrayIcon();

	LPTSTR aStr1 = *(LPTSTR *)a1;
	LPTSTR aStr2 = *(LPTSTR *)a2;
	ExprTokenType param[] = { aStr1, aStr2, __int64(aStr2 - aStr1) };
	__int64 i64;
	g_SortFuncResult = CallMethod(g_SortFunc, g_SortFunc, nullptr, param, _countof(param), &i64);
	// An alternative to g_SortFuncResult using 'throw' to abort qsort() produced slightly
	// smaller code, but in release builds the program crashed with code 0xC0000409 and the
	// catch() block never executed.

	int returned_int;
	if (i64)  // Maybe there's a faster/better way to do these checks. Can't simply typecast to an int because some large positives wrap into negative, maybe vice versa.
		returned_int = i64 < 0 ? -1 : 1;
	else  // The result was 0 or "".  Since returning 0 would make the results less predictable and doesn't seem to have any benefit,
		returned_int = aStr1 < aStr2 ? -1 : 1; // use the relative positions of the items as a tie-breaker.

	return returned_int;
}



BIF_DECL(BIF_Sort)
{
	// Set defaults in case of early goto:
	LPTSTR mem_to_free = NULL;
	LPTSTR *item = NULL; // The index/pointer list used for the sort.
	IObject *sort_func_orig = g_SortFunc; // Because UDFs can be interrupted by other threads -- and because UDFs can themselves call Sort with some other UDF (unlikely to be sure) -- backup & restore original g_SortFunc so that the "collapsing in reverse order" behavior will automatically ensure proper operation.
	ResultType sort_func_result_orig = g_SortFuncResult;
	g_SortFunc = NULL; // Now that original has been saved above, reset to detect whether THIS sort uses a UDF.
	g_SortFuncResult = OK;

	_f_param_string(aContents, 0);
	_f_param_string_opt(aOptions, 1);

	// Resolve options.  First set defaults for options:
	TCHAR delimiter = '\n';
	g_SortCaseSensitive = SCS_INSENSITIVE;
	g_SortNumeric = false;
	g_SortReverse = false;
	g_SortColumnOffset = 0;
	bool trailing_delimiter_indicates_trailing_blank_item = false, terminate_last_item_with_delimiter = false
		, trailing_crlf_added_temporarily = false, sort_by_naked_filename = false, sort_random = false
		, omit_dupes = false;
	LPTSTR cp;

	for (cp = aOptions; *cp; ++cp)
	{
		switch(_totupper(*cp))
		{
		case 'C':
			if (ctoupper(cp[1]) == 'L')
			{
				if (!_tcsnicmp(cp+2, _T("Logical") + 1, 6)) // CLogical.  Using "Logical" + 1 instead of "ogical" was confirmed to eliminate one string from the binary (due to string pooling).
				{
					cp += 7;
					g_SortCaseSensitive = SCS_INSENSITIVE_LOGICAL;
				}
				else
				{
					if (!_tcsnicmp(cp+2, _T("Locale") + 1, 5)) // CLocale
						cp += 6;
					else // CL
						++cp;
					g_SortCaseSensitive = SCS_INSENSITIVE_LOCALE;
				}
			}
			else if (!_tcsnicmp(cp+1, _T("Off"), 3)) // COff.  Using ctoupper() here significantly increased code size.
			{
				cp += 3;
				g_SortCaseSensitive = SCS_INSENSITIVE;
			}
			else if (cp[1] == '0') // C0
				g_SortCaseSensitive = SCS_INSENSITIVE;
			else // C  C1  COn
			{
				if (!_tcsnicmp(cp+1, _T("On"), 2)) // COn.  Using ctoupper() here significantly increased code size.
					cp += 2;
				g_SortCaseSensitive = SCS_SENSITIVE;
			}
			break;
		case 'D':
			if (!cp[1]) // Avoids out-of-bounds when the loop's own ++cp is done.
				break;
			++cp;
			if (*cp)
				delimiter = *cp;
			break;
		case 'N':
			g_SortNumeric = true;
			break;
		case 'P':
			// Use atoi() vs. ATOI() to avoid interpreting something like 0x01C as hex
			// when in fact the C was meant to be an option letter:
			g_SortColumnOffset = _ttoi(cp + 1);
			if (g_SortColumnOffset < 1)
				g_SortColumnOffset = 1;
			--g_SortColumnOffset;  // Convert to zero-based.
			break;
		case 'R':
			if (!_tcsnicmp(cp, _T("Random"), 6))
			{
				sort_random = true;
				cp += 5; // Point it to the last char so that the loop's ++cp will point to the character after it.
			}
			else
				g_SortReverse = true;
			break;
		case 'U':  // Unique.
			omit_dupes = true;
			break;
		case 'Z':
			// By setting this to true, the final item in the list, if it ends in a delimiter,
			// is considered to be followed by a blank item:
			trailing_delimiter_indicates_trailing_blank_item = true;
			break;
		case '\\':
			sort_by_naked_filename = true;
		}
	}

	if (!ParamIndexIsOmitted(2))
	{
		if (  !(g_SortFunc = ParamIndexToObject(2))  )
		{
			aResultToken.ParamError(2, aParam[2]);
			goto end;
		}
		g_SortFunc->AddRef(); // Must be done in case the parameter was SYM_VAR and that var gets reassigned.
	}
	
	if (!*aContents) // Input is empty, nothing to sort, return empty string.
	{
		aResultToken.Return(_T(""), 0);
		goto end;
	}
	
	size_t item_count;

	// Check how many delimiters are present:
	for (item_count = 1, cp = aContents; *cp; ++cp)  // Start at 1 since item_count is delimiter_count+1
		if (*cp == delimiter)
			++item_count;
	size_t aContents_length = cp - aContents;

	// If the last character in the unsorted list is a delimiter then technically the final item
	// in the list is a blank item.  However, if the options specify not to allow that, don't count that
	// blank item as an actual item (i.e. omit it from the list):
	if (!trailing_delimiter_indicates_trailing_blank_item && cp > aContents && cp[-1] == delimiter)
	{
		terminate_last_item_with_delimiter = true; // Have a later stage add the delimiter to the end of the sorted list so that the length and format of the sorted list is the same as the unsorted list.
		--item_count; // Indicate that the blank item at the end of the list isn't being counted as an item.
	}
	else // The final character isn't a delimiter or it is but there's a blank item to its right.  Either way, the following is necessary (verified correct).
	{
		if (delimiter == '\n')
		{
			LPTSTR first_delimiter = _tcschr(aContents, delimiter);
			if (first_delimiter && first_delimiter > aContents && first_delimiter[-1] == '\r')
			{
				// v1.0.47.05:
				// Here the delimiter is effectively CRLF vs. LF.  Therefore, signal a later section to append
				// an extra CRLF to simplify the code and fix bugs that existed prior to v1.0.47.05.
				// One such bug is that sorting "x`r`nx" in unique mode would previously produce two
				// x's rather than one because the first x is seen to have a `r to its right but the
				// second lacks it (due to the CRLF delimiter being simulated via LF-only).
				//
				// OLD/OBSOLETE comment from a section that was removed because it's now handled by this section:
				// Check if special handling is needed due to the following situation:
				// Delimiter is LF but the contents are lines delimited by CRLF, not just LF
				// and the original/unsorted list's last item was not terminated by an
				// "allowed delimiter".  The symptoms of this are that after the sort, the
				// last item will end in \r when it should end in no delimiter at all.
				// This happens pretty often, such as when the clipboard contains files.
				// In the future, an option letter can be added to turn off this workaround,
				// but it seems unlikely that anyone would ever want to.
				trailing_crlf_added_temporarily = true;
				terminate_last_item_with_delimiter = true; // This is done to "fool" later sections into thinking the list ends in a CRLF that doesn't have a blank item to the right, which in turn simplifies the logic and/or makes it more understandable.
			}
		}
	}

	if (item_count == 1) // 1 item is already sorted, and no dupes are possible.
	{
		aResultToken.Return(aContents, aContents_length);
		goto end;
	}

	// v1.0.47.05: It simplifies the code a lot to allocate and/or improves understandability to allocate
	// memory for trailing_crlf_added_temporarily even though technically it's done only to make room to
	// append the extra CRLF at the end.
	// v2.0: Never modify the caller's aContents, since it may be a quoted literal string or variable.
	//if (g_SortFunc || trailing_crlf_added_temporarily) // Do this here rather than earlier with the options parsing in case the function-option is present twice (unlikely, but it would be a memory leak due to strdup below).  Doing it here also avoids allocating if it isn't necessary.
	{
		// Comment is obsolete because if aContents is in a deref buffer, it has been privatized by ExpandArgs():
		// When g_SortFunc!=NULL, the copy of the string is needed because aContents may be in the deref buffer,
		// and that deref buffer is about to be overwritten by the execution of the script's UDF body.
		if (   !(mem_to_free = tmalloc(aContents_length + 3))   ) // +1 for terminator and +2 in case of trailing_crlf_added_temporarily.
		{
			aResultToken.MemoryError();
			goto end;
		}
		tmemcpy(mem_to_free, aContents, aContents_length + 1); // memcpy() usually benches a little faster than _tcscpy().
		aContents = mem_to_free;
		if (trailing_crlf_added_temporarily)
		{
			// Must be added early so that the next stage will terminate at the \n, leaving the \r as
			// the ending character in this item.
			_tcscpy(aContents + aContents_length, _T("\r\n"));
			aContents_length += 2;
		}
	}

	// Create the array of pointers that points into aContents to each delimited item.
	// Use item_count + 1 to allow space for the last (blank) item in case
	// trailing_delimiter_indicates_trailing_blank_item is false:
	int unit_size = sort_random ? 2 : 1;
	size_t item_size = unit_size * sizeof(LPTSTR);
	item = (LPTSTR *)malloc((item_count + 1) * item_size);
	if (!item)
	{
		aResultToken.MemoryError();
		goto end;
	}

	// If sort_random is in effect, the above has created an array twice the normal size.
	// This allows the random numbers to be interleaved inside the array as though it
	// were an array consisting of sort_rand_type (which it actually is when viewed that way).
	// Because of this, the array should be accessed through pointer addition rather than
	// indexing via [].

	// Scan aContents and do the following:
	// 1) Replace each delimiter with a terminator so that the individual items can be seen
	//    as real strings by the SortWithOptions() and when copying the sorted results back
	//    into the result.
	// 2) Store a marker/pointer to each item (string) in aContents so that we know where
	//    each item begins for sorting and recopying purposes.
	LPTSTR *item_curr = item; // i.e. Don't use [] indexing for the reason in the paragraph previous to above.
	for (item_count = 0, cp = *item_curr = aContents; *cp; ++cp)
	{
		if (*cp == delimiter)  // Each delimiter char becomes the terminator of the previous key phrase.
		{
			*cp = '\0';  // Terminate the item that appears before this delimiter.
			++item_count;
			if (sort_random)
			{
				auto &n = *(unsigned int *)(item_curr + 1); // i.e. the randoms are in the odd fields, the pointers in the even.
				GenRandom(&n, sizeof(n));
				n >>= 1;
				// For the above:
				// The random numbers given to SortRandom() must be within INT_MAX of each other, otherwise
				// integer overflow would cause the difference between two items to flip signs depending
				// on how far apart the numbers are, which would mean for some cases where A < B and B < C,
				// it's possible that C > A.  Without n >>= 1, the problem can be proven via the following
				// script, which shows a sharply biased distribution:
				//count := 0
				//Loop total := 10000
				//{
				//	var := Sort("1`n2`n3`n4`n5`n", "Random")
				//	if SubStr(var, 1, 1) = 5
				//		count += 1
				//}
				//MsgBox Round(count / total * 100, 2) "%"
			}
			item_curr += unit_size; // i.e. Don't use [] indexing for the reason described above.
			*item_curr = cp + 1; // Make a pointer to the next item's place in aContents.
		}
	}
	// The above reset the count to 0 and recounted it.  So now re-add the last item to the count unless it was
	// disqualified earlier. Verified correct:
	if (!terminate_last_item_with_delimiter) // i.e. either trailing_delimiter_indicates_trailing_blank_item==true OR the final character isn't a delimiter. Either way the final item needs to be added.
	{
		++item_count;
		if (sort_random) // Provide a random number for the last item.
		{
			auto &n = *(unsigned int *)(item_curr + 1); // i.e. the randoms are in the odd fields, the pointers in the even.
			GenRandom(&n, sizeof(n));
			n >>= 1;
		}
	}
	else // Since the final item is not included in the count, point item_curr to the one before the last, for use below.
		item_curr -= unit_size;

	// Now aContents has been divided up based on delimiter.  Sort the array of pointers
	// so that they indicate the correct ordering to copy aContents into output_var:
	if (g_SortFunc) // Takes precedence other sorting methods.
	{
		qsort((void *)item, item_count, item_size, SortUDF);
		if (!g_SortFuncResult || g_SortFuncResult == EARLY_EXIT)
		{
			aResultToken.SetExitResult(g_SortFuncResult);
			goto end;
		}
	}
	else if (sort_random) // Takes precedence over all remaining options.
		qsort((void *)item, item_count, item_size, SortRandom);
	else
		qsort((void *)item, item_count, item_size, sort_by_naked_filename ? SortByNakedFilename : SortWithOptions);

	// Allocate space to store the result.
	if (!TokenSetResult(aResultToken, NULL, aContents_length))
		goto end;
	aResultToken.symbol = SYM_STRING;

	// Set default in case original last item is still the last item, or if last item was omitted due to being a dupe:
	size_t i, item_count_minus_1 = item_count - 1;
	DWORD omit_dupe_count = 0;
	bool keep_this_item;
	LPTSTR source, dest;
	LPTSTR item_prev = NULL;

	// Copy the sorted result back into output_var.  Do all except the last item, since the last
	// item gets special treatment depending on the options that were specified.
	item_curr = item; // i.e. Don't use [] indexing for the reason described higher above (same applies to item += unit_size below).
	for (dest = aResultToken.marker, i = 0; i < item_count; ++i, item_curr += unit_size)
	{
		keep_this_item = true;  // Set default.
		if (omit_dupes && item_prev)
		{
			// Update to the comment below: Exact dupes will still be removed when sort_by_naked_filename
			// or g_SortColumnOffset is in effect because duplicate lines would still be adjacent to
			// each other even in these modes.  There doesn't appear to be any exceptions, even if
			// some items in the list are sorted as blanks due to being shorter than the specified 
			// g_SortColumnOffset.
			// As documented, special dupe-checking modes are not offered when sort_by_naked_filename
			// is in effect, or g_SortColumnOffset is greater than 1.  That's because the need for such
			// a thing seems too rare (and the result too strange) to justify the extra code size.
			// However, adjacent dupes are still removed when any of the above modes are in effect,
			// or when the "random" mode is in effect.  This might have some usefulness; for example,
			// if a list of songs is sorted in random order, but it contains "favorite" songs listed twice,
			// the dupe-removal feature would remove duplicate songs if they happen to be sorted
			// to lie adjacent to each other, which would be useful to prevent the same song from
			// playing twice in a row.
			if (g_SortNumeric && !g_SortColumnOffset)
				// if g_SortColumnOffset is zero, fall back to the normal dupe checking in case its
				// ever useful to anyone.  This is done because numbers in an offset column are not supported
				// since the extra code size doensn't seem justified given the rarity of the need.
				keep_this_item = (ATOF(*item_curr) != ATOF(item_prev)); // ATOF() ignores any trailing \r in CRLF mode, so no extra logic is needed for that.
			else if (g_SortCaseSensitive == SCS_INSENSITIVE_LOGICAL)
				keep_this_item = StrCmpLogicalW(*item_curr, item_prev);
			else
				keep_this_item = tcscmp2(*item_curr, item_prev, g_SortCaseSensitive); // v1.0.43.03: Added support for locale-insensitive mode.
				// Permutations of sorting case sensitive vs. eliminating duplicates based on case sensitivity:
				// 1) Sort is not case sens, but dupes are: Won't work because sort didn't necessarily put
				//    same-case dupes adjacent to each other.
				// 2) Converse: probably not reliable because there could be unrelated items in between
				//    two strings that are duplicates but weren't sorted adjacently due to their case.
				// 3) Both are case sensitive: seems okay
				// 4) Both are not case sensitive: seems okay
				//
				// In light of the above, using the g_SortCaseSensitive flag to control the behavior of
				// both sorting and dupe-removal seems best.
		}
		if (keep_this_item)
		{
			for (source = *item_curr; *source;)
				*dest++ = *source++;
			// If we're at the last item and the original list's last item had a terminating delimiter
			// and the specified options said to treat it not as a delimiter but as a final char of sorts,
			// include it after the item that is now last so that the overall layout is the same:
			if (i < item_count_minus_1 || terminate_last_item_with_delimiter)
				*dest++ = delimiter;  // Put each item's delimiter back in so that format is the same as the original.
			item_prev = *item_curr; // Since the item just processed above isn't a dupe, save this item to compare against the next item.
		}
		else // This item is a duplicate of the previous item.
		{
			++omit_dupe_count; // But don't change the value of item_prev.
			// v1.0.47.05: The following section fixes the fact that the "unique" option would sometimes leave
			// a trailing delimiter at the end of the sorted list.  For example, sorting "x|x" in unique mode
			// would formerly produce "x|":
			if (i == item_count_minus_1 && !terminate_last_item_with_delimiter) // This is the last item and its being omitted, so remove the previous item's trailing delimiter (unless a trailing delimiter is mandatory).
				--dest; // Remove the previous item's trailing delimiter there's nothing for it to delimit due to omission of this duplicate.
		}
	} // for()

	// Terminate the variable's contents.
	if (trailing_crlf_added_temporarily) // Remove the CRLF only after its presence was used above to simplify the code by reducing the number of types/cases.
	{
		dest[-2] = '\0';
		aResultToken.marker_length -= 2;
	}
	else
		*dest = '\0';

	if (omit_dupes)
	{
		if (omit_dupe_count) // Update the length to actual whenever at least one dupe was omitted.
		{
			aResultToken.marker_length = _tcslen(aResultToken.marker);
		}
	}
	//else it is not necessary to set the length here because its length hasn't
	// changed since it was originally set by the above call TokenSetResult.

end:
	if (!item)
		free(item);
	if (mem_to_free)
		free(mem_to_free);
	if (g_SortFunc)
		g_SortFunc->Release();
	g_SortFunc = sort_func_orig;
	g_SortFuncResult = sort_func_result_orig;
}



BIF_DECL(BIF_StrCompare)
{
	// In script:
	// Result := StrCompare(str1, str2 [, CaseSensitive := false])
	// str1, str2, the strings to compare, can be be pure numbers, such numbers are converted to strings before the comparison.
	// CaseSensitive, the case sensitive setting to use.
	// Result, the result of the comparison:
	//	< 0,	if str1 is less than str2.
	//	0,		if str1 is identical to str2.
	//	> 0,	if str1 is greater than str2.
	// Param 1 and 2, str1 and str2
	TCHAR str1_buf[MAX_NUMBER_SIZE];	// numeric input is converted to string.
	TCHAR str2_buf[MAX_NUMBER_SIZE];
	LPTSTR str1 = ParamIndexToString(0, str1_buf);
	LPTSTR str2 = ParamIndexToString(1, str2_buf);
	// Could return 0 here if str1 == str2, but it is probably rare to call StrCompare(str_var, str_var)
	// so for most cases that would just be an unnecessary cost, albeit low.
	// Param 3 - CaseSensitive
	StringCaseSenseType string_case_sense = ParamIndexToCaseSense(2);
	int result;
	switch (string_case_sense)		// Compare the strings according to the string_case_sense setting.
	{
	case SCS_SENSITIVE:				result = _tcscmp(str1, str2); break;
	case SCS_INSENSITIVE:			result = _tcsicmp(str1, str2); break;
	case SCS_INSENSITIVE_LOCALE:	result = lstrcmpi(str1, str2); break;
	case SCS_INSENSITIVE_LOGICAL:	result = StrCmpLogicalW(str1, str2); break;
	case SCS_INVALID:
		_f_throw_param(2);
	}
	_f_return_i(result);			// Return 
}

BIF_DECL(BIF_String)
{
	++aParam; // Skip `this`
	aResultToken.symbol = SYM_STRING;
	switch (aParam[0]->symbol)
	{
	case SYM_STRING:
		aResultToken.marker = aParam[0]->marker;
		aResultToken.marker_length = aParam[0]->marker_length;
		break;
	case SYM_VAR:
		if (aParam[0]->var->HasObject())
		{
			ObjectToString(aResultToken, *aParam[0], aParam[0]->var->Object());
			break;
		}
		aResultToken.marker = aParam[0]->var->Contents();
		aResultToken.marker_length = aParam[0]->var->CharLength();
		break;
	case SYM_INTEGER:
		aResultToken.marker = ITOA64(aParam[0]->value_int64, _f_retval_buf);
		break;
	case SYM_FLOAT:
		aResultToken.marker = _f_retval_buf;
		aResultToken.marker_length = FTOA(aParam[0]->value_double, aResultToken.marker, _f_retval_buf_size);
		break;
	case SYM_OBJECT:
		ObjectToString(aResultToken, *aParam[0], aParam[0]->object);
		break;
	// Impossible due to parameter count validation:
	//case SYM_MISSING:
	//	_f_throw_value(ERR_PARAM1_REQUIRED);
#ifdef _DEBUG
	default:
		MsgBox(_T("DEBUG: type not handled"));
		_f_return_FAIL;
#endif
	}
}



BIF_DECL(BIF_StrLen)
{
	// Caller has ensured that there's exactly one actual parameter.
	size_t length;
	ParamIndexToString(0, _f_number_buf, &length); // Allow StrLen(numeric_expr) for flexibility.
	_f_return_i(length);
}



BIF_DECL(BIF_SubStr) // Added in v1.0.46.
{
	// Set default return value in case of early return.
	_f_set_retval_p(_T(""), 0);

	// Get the first arg, which is the string used as the source of the extraction. Call it "haystack" for clarity.
	TCHAR haystack_buf[MAX_NUMBER_SIZE]; // A separate buf because _f_number_buf is sometimes used to store the result.
	INT_PTR haystack_length;
	LPTSTR haystack = ParamIndexToString(0, haystack_buf, (size_t *)&haystack_length);

	// Load-time validation has ensured that at least the first two parameters are present:
	INT_PTR starting_offset = (INT_PTR)ParamIndexToInt64(1); // The one-based starting position in haystack (if any).
	if (starting_offset > haystack_length || starting_offset == 0)
		_f_return_retval; // Yield the empty string (a default set higher above).
	if (starting_offset < 0) // Same convention as RegExMatch/Replace(): Treat negative StartingPos as a position relative to the end of the string.
	{
		starting_offset += haystack_length;
		if (starting_offset < 0)
			starting_offset = 0;
	}
	else
		--starting_offset; // Convert to zero-based.

	INT_PTR remaining_length_available = haystack_length - starting_offset;
	INT_PTR extract_length;
	if (ParamIndexIsOmitted(2)) // No length specified, so extract all the remaining length.
		extract_length = remaining_length_available;
	else
	{
		if (   !(extract_length = (INT_PTR)ParamIndexToInt64(2))   )  // It has asked to extract zero characters.
			_f_return_retval; // Yield the empty string (a default set higher above).
		if (extract_length < 0)
		{
			extract_length += remaining_length_available; // Result is the number of characters to be extracted (i.e. after omitting the number of chars specified in extract_length).
			if (extract_length < 1) // It has asked to omit all characters.
				_f_return_retval; // Yield the empty string (a default set higher above).
		}
		else // extract_length > 0
			if (extract_length > remaining_length_available)
				extract_length = remaining_length_available;
	}

	// Above has set extract_length to the exact number of characters that will actually be extracted.
	LPTSTR result = haystack + starting_offset; // This is the result except for the possible need to truncate it below.

	// No need for any copying or termination, just send back part of haystack.
	// Caller and Var::Assign() know that overlap is possible, so this seems safe.
	_f_return_p(result, extract_length);
}



BIF_DECL(BIF_InStr)
{
	// Caller has already ensured that at least two actual parameters are present.
	TCHAR needle_buf[MAX_NUMBER_SIZE];
	INT_PTR haystack_length;
	LPTSTR haystack = ParamIndexToString(0, _f_number_buf, (size_t *)&haystack_length);
	size_t needle_length;
	LPTSTR needle = ParamIndexToString(1, needle_buf, &needle_length);
	if (!needle_length) // Although arguably legitimate, this is more likely an error, so throw.
		_f_throw_param(1);

	StringCaseSenseType string_case_sense = ParamIndexToCaseSense(2);
	if (string_case_sense == SCS_INVALID
		|| string_case_sense == SCS_INSENSITIVE_LOGICAL) // Not supported, seems more useful to throw rather than using SCS_INSENSITIVE.
		_f_throw_param(2);
	// BIF_StrReplace sets string_case_sense similarly, maintain together.

	INT_PTR offset = ParamIndexToOptionalIntPtr(3, 1);
	if (!offset)
		_f_throw_param(3);
	
	int occurrence_number = ParamIndexToOptionalInt(4, 1);
	if (!occurrence_number)
		_f_throw_param(4);

	if (offset < 0)
	{
		if (ParamIndexIsOmitted(4))
			occurrence_number = -1; // Default to RTL.
		offset += haystack_length + 1; // Convert end-relative position to start-relative.
	}
	if (occurrence_number > 0)
		--offset; // Convert 1-based to 0-based.
	if (offset < 0) // To the left of the first character.
		offset = 0; // LTR: start at offset 0.  RTL: search 0 characters.
	else if (offset > haystack_length) // To the right of the last character.
		offset = haystack_length; // LTR: start at null-terminator (omit all).  RTL: search whole string.

	LPTSTR found_pos;
	if (occurrence_number < 0) // Special mode to search from the right side (RTL).
	{
		if (!ParamIndexIsOmitted(3)) // For this mode, offset is only relevant if specified by the caller.
			haystack_length = offset; // i.e. omit any characters to the right of the starting position.
		found_pos = tcsrstr(haystack, haystack_length, needle, string_case_sense, -occurrence_number);
		_f_return_i(found_pos ? (found_pos - haystack + 1) : 0);  // +1 to convert to 1-based, since 0 indicates "not found".
	}
	// Since above didn't return:
	int i;
	for (i = 1, found_pos = haystack + offset; ; ++i, found_pos += needle_length)
		if (!(found_pos = tcsstr2(found_pos, needle, string_case_sense)) || i == occurrence_number)
			break;
	_f_return_i(found_pos ? (found_pos - haystack + 1) : 0);
}



BIF_DECL(BIF_Ord)
{
	// Result will always be an integer (this simplifies scripts that work with binary zeros since an
	// empty string yields zero).
	LPTSTR cp = ParamIndexToString(0, _f_number_buf);
#ifndef UNICODE
	// This section could convert a single- or multi-byte character to a Unicode code point
	// for consistency, but since the ANSI build won't be officially supported that doesn't
	// seem necessary.
#else
	if (IS_SURROGATE_PAIR(cp[0], cp[1]))
		_f_return_i(((cp[0] - 0xd800) << 10) + (cp[1] - 0xdc00) + 0x10000);
	else
#endif
		_f_return_i((TBYTE)*cp);
}



BIF_DECL(BIF_Chr)
{
	Throw_if_Param_NaN(0);
	int param1 = ParamIndexToInt(0); // Convert to INT vs. UINT so that negatives can be detected.
	LPTSTR cp = _f_retval_buf; // If necessary, it will be moved to a persistent memory location by our caller.
	int len;
	if (param1 < 0 || param1 > UorA(0x10FFFF, UCHAR_MAX))
		_f_throw_param(0);
#ifdef UNICODE
	else if (param1 >= 0x10000)
	{
		param1 -= 0x10000;
		cp[0] = 0xd800 + ((param1 >> 10) & 0x3ff);
		cp[1] = 0xdc00 + ( param1        & 0x3ff);
		cp[2] = '\0';
		len = 2;
	}
#else
	// See #ifndef in BIF_Ord for related comment.
#endif
	else
	{
		cp[0] = param1;
		cp[1] = '\0';
		len = 1; // Even if param1 == 0.
	}
	_f_return_p(cp, len);
}



BIF_DECL(BIF_Format)
{
	if (TokenIsPureNumeric(*aParam[0]))
		_f_return_p(ParamIndexToString(0, _f_retval_buf));

	LPCTSTR fmt = ParamIndexToString(0), lit, cp, cp_end, cp_spec;
	LPTSTR target = NULL;
	int size = 0, spec_len;
	int param, last_param;
	TCHAR number_buf[MAX_NUMBER_SIZE];
	TCHAR spec[12+MAX_INTEGER_LENGTH*2];
	TCHAR custom_format;
	ExprTokenType value;
	*spec = '%';

	for (;;)
	{
		last_param = 0;

		for (lit = cp = fmt;; )
		{
			// Find next placeholder.
			for (cp_end = cp; *cp_end && *cp_end != '{'; ++cp_end);
			if (cp_end > lit)
			{
				// Handle literal text to the left of the placeholder.
				if (target)
					tmemcpy(target, lit, cp_end - lit), target += cp_end - lit;
				else
					size += int(cp_end - lit);
				lit = cp_end; // Mark this as the next literal character (to be overridden below if it's a valid placeholder).
			}
			cp = cp_end;
			if (!*cp)
				break;
			// else: Implies *cp == '{'.
			++cp;
			if ((*cp == '{' || *cp == '}') && cp[1] == '}') // {{} or {}}
			{
				if (target)
					*target++ = *cp;
				else
					++size;
				cp += 2;
				lit = cp; // Mark this as the next literal character.
				continue;
			}
			
			// Index.
			for (cp_end = cp; *cp_end >= '0' && *cp_end <= '9'; ++cp_end);
			if (cp_end > cp)
				param = ATOI(cp), cp = cp_end;
			else
				param = last_param + 1;
			if (param >= aParamCount || param < 1) // Invalid parameter index.
				continue;

			custom_format = 0; // Set default.

			// Optional format specifier.
			if (*cp == ':')
			{
				cp_spec = ++cp;
				// Skip valid format specifier options.
				for (cp = cp_spec; *cp && _tcschr(_T("-+0 #"), *cp); ++cp); // flags
				for ( ; *cp >= '0' && *cp <= '9'; ++cp); // width
				if (*cp == '.') do ++cp; while (*cp >= '0' && *cp <= '9'); // .precision
				spec_len = int(cp - cp_spec);
				// For now, size specifiers (h | l | ll | w | I | I32 | I64) are not supported.
				
				if (spec_len + 4 >= _countof(spec)) // Format specifier too long (probably invalid).
					continue;
				// Copy options, if any (+1 to leave the leading %).
				tmemcpy(spec + 1, cp_spec, spec_len);
				++spec_len; // Include the leading %.

				if (_tcschr(_T("diouxX"), *cp))
				{
					spec[spec_len++] = 'I';
					spec[spec_len++] = '6';
					spec[spec_len++] = '4';
					// Integer value; apply I64 prefix to avoid truncation.
					value.value_int64 = ParamIndexToInt64(param);
					spec[spec_len++] = *cp++;
				}
				else if (_tcschr(_T("eEfgGaA"), *cp))
				{
					value.value_double = ParamIndexToDouble(param);
					spec[spec_len++] = *cp++;
				}
				else if (_tcschr(_T("cCp"), *cp))
				{
					// Input is an integer or pointer, but I64 prefix should not be applied.
					value.value_int64 = ParamIndexToInt64(param);
					spec[spec_len++] = *cp++;
				}
				else
				{
					spec[spec_len++] = 's'; // Default to string if not specified.
					if (_tcschr(_T("ULlTt"), *cp))
						custom_format = toupper(*cp++);
					if (*cp == 's')
						++cp;
				}
			}
			else
			{
				// spec[0] contains '%'.
				spec[1] = 's';
				spec_len = 2;
			}
			if (spec[spec_len - 1] == 's')
			{
				value.marker = ParamIndexToString(param, number_buf);
			}
			spec[spec_len] = '\0';
			
			if (*cp != '}') // Syntax error.
				continue;
			++cp;
			lit = cp; // Mark this as the next literal character.

			// Now that validation is complete, set last_param for use by the next {} or {:fmt}.
			last_param = param;
			
			if (target)
			{
				int len = _stprintf(target, spec, value.value_int64);
				switch (custom_format)
				{
				case 0: break; // Might help performance to list this first.
				case 'U': CharUpper(target); break;
				case 'L': CharLower(target); break;
				case 'T': StrToTitleCase(target); break;
				}
				target += len;
			}
			else
				size += _sctprintf(spec, value.value_int64);
		}
		if (target)
		{
			// Finished second pass.
			*target = '\0';
			return;
		}
		// Finished first pass (calculating required size).
		if (!TokenSetResult(aResultToken, NULL, size))
			return;
		aResultToken.symbol = SYM_STRING;
		target = aResultToken.marker;
	}
}



BIF_DECL(BIF_Trim) // L31
{
	BuiltInFunctionID trim_type = _f_callee_id;

	LPTSTR buf = _f_retval_buf;
	size_t extract_length;
	LPTSTR str = ParamIndexToString(0, buf, &extract_length);

	LPTSTR result = str; // Prior validation has ensured at least 1 param.

	TCHAR omit_list_buf[MAX_NUMBER_SIZE]; // Support SYM_INTEGER/SYM_FLOAT even though it doesn't seem likely to happen.
	LPTSTR omit_list = ParamIndexIsOmitted(1) ? _T(" \t") : ParamIndexToString(1, omit_list_buf); // Default: space and tab.

	if (trim_type != FID_RTrim) // i.e. it's Trim() or LTrim()
	{
		result = omit_leading_any(result, omit_list, extract_length);
		extract_length -= (result - str); // Adjust for omitted characters.
	}
	if (extract_length && trim_type != FID_LTrim) // i.e. it's Trim() or RTrim();  THE LINE BELOW REQUIRES extract_length >= 1.
		extract_length = omit_trailing_any(result, omit_list, result + extract_length - 1);

	_f_return_p(result, extract_length);
}



BIF_DECL(BIF_VerCompare)
{
	_f_param_string(a, 0);
	_f_param_string(b, 1);
	_f_return(CompareVersion(a, b));
}



////////////////////
// Misc Functions //
////////////////////


BIF_DECL(BIF_Type)
{
	_f_return_p(TokenTypeString(*aParam[0]));
}

// Returns the type name of the given value.
LPTSTR TokenTypeString(ExprTokenType &aToken)
{
	switch (TypeOfToken(aToken))
	{
	case SYM_STRING: return STRING_TYPE_STRING;
	case SYM_INTEGER: return INTEGER_TYPE_STRING;
	case SYM_FLOAT: return FLOAT_TYPE_STRING;
	case SYM_OBJECT: return TokenToObject(aToken)->Type();
	default: return _T(""); // For maintainability.
	}
}



void Object::Error__New(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	LPTSTR message;
	TCHAR what_buf[MAX_NUMBER_SIZE], extra_buf[MAX_NUMBER_SIZE];
	LPCTSTR what = ParamIndexToOptionalString(1, what_buf);
	Line *line = g_script.mCurrLine;

	if (aID == M_OSError__New && (ParamIndexIsOmitted(0) || ParamIndexIsNumeric(0)))
	{
		DWORD error = ParamIndexIsOmitted(0) ? g->LastError : (DWORD)ParamIndexToInt64(0);
		SetOwnProp(_T("Number"), error);
		
		// Determine message based on error number.
		DWORD message_buf_size = _f_retval_buf_size;
		message = _f_retval_buf;
		DWORD size = (DWORD)_sntprintf(message, message_buf_size, (int)error < 0 ? _T("(0x%X) ") : _T("(%i) "), error);
		if (error) // Never show "Error: (0) The operation completed successfully."
			size += FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, NULL, error, 0, message + size, message_buf_size - size, NULL);
		if (size)
		{
			if (message[size - 1] == '\n')
				message[--size] = '\0';
			if (message[size - 1] == '\r')
				message[--size] = '\0';
		}
	}
	else
		message = ParamIndexIsOmitted(0) ? Type() : ParamIndexToString(0, _f_number_buf);

#ifndef CONFIG_DEBUGGER
	if (ParamIndexIsOmitted(1) && g->CurrentFunc)
		what = g->CurrentFunc->mName;
	SetOwnProp(_T("Stack"), _T("")); // Avoid "unknown property" in compiled scripts.
#else
	DbgStack::Entry *stack_top = g_Debugger.mStack.mTop - 1;
	if (stack_top->type == DbgStack::SE_BIF && _tcsicmp(what, stack_top->func->mName))
		--stack_top;

	if (ParamIndexIsOmitted(1)) // "What"
	{
		if (g->CurrentFunc)
			what = g->CurrentFunc->mName;
	}
	else
	{
		int offset = ParamIndexIsNumeric(1) ? ParamIndexToInt(1) : 0;
		for (auto se = stack_top; se >= g_Debugger.mStack.mBottom; --se)
		{
			if (++offset == 0 || *what && !_tcsicmp(se->Name(), what))
			{
				line = se > g_Debugger.mStack.mBottom ? se[-1].line : se->line;
				// se->line contains the line at the given offset from the top of the stack.
				// Rather than returning the name of the function or sub which contains that
				// line, return the name of the function or sub which that line called.
				// In other words, an offset of -1 gives the name of the current function and
				// the file and number of the line which it was called from.
				what = se->Name();
				stack_top = se;
				break;
			}
			if (se->type == DbgStack::SE_Thread)
				break; // Look only within the current thread.
		}
	}

	TCHAR stack_buf[2048];
	GetScriptStack(stack_buf, _countof(stack_buf), stack_top);
	SetOwnProp(_T("Stack"), stack_buf);
#endif

	LPTSTR extra = ParamIndexToOptionalString(2, extra_buf);

	SetOwnProp(_T("Message"), message);
	SetOwnProp(_T("What"), const_cast<LPTSTR>(what));
	SetOwnProp(_T("File"), Line::sSourceFile[line->mFileIndex]);
	SetOwnProp(_T("Line"), line->mLineNumber);
	SetOwnProp(_T("Extra"), extra);
}



////////////////////////////////////////////////////////
// HELPER FUNCTIONS FOR TOKENS AND BUILT-IN FUNCTIONS //
////////////////////////////////////////////////////////

BOOL ResultToBOOL(LPTSTR aResult)
{
	UINT c1 = (UINT)*aResult; // UINT vs. UCHAR might squeeze a little more performance out of it.
	if (c1 > 48)     // Any UCHAR greater than '0' can't be a space(32), tab(9), '+'(43), or '-'(45), or '.'(46)...
		return TRUE; // ...so any string starting with c1>48 can't be anything that's false (e.g. " 0", "+0", "-0", ".0", "0").
	if (!c1 // Besides performance, must be checked anyway because otherwise IsNumeric() would consider "" to be non-numeric and thus TRUE.
		|| c1 == 48 && !aResult[1]) // The string "0" seems common enough to detect explicitly, for performance.
		return FALSE;
	// IsNumeric() is called below because there are many variants of a false string:
	// e.g. "0", "0.0", "0x0", ".0", "+0", "-0", and " 0" (leading spaces/tabs).
	switch (IsNumeric(aResult, true, false, true)) // It's purely numeric and not all whitespace (and due to earlier check, it's not blank).
	{
	case PURE_INTEGER: return ATOI64(aResult) != 0; // Could call ATOF() for both integers and floats; but ATOI64() probably performs better, and a float comparison to 0.0 might be a slower than an integer comparison to 0.
	case PURE_FLOAT:   return _tstof(aResult) != 0.0; // atof() vs. ATOF() because PURE_FLOAT is never hexadecimal.
	default: // PURE_NOT_NUMERIC.
		// Even a string containing all whitespace would be considered non-numeric since it's a non-blank string
		// that isn't equal to 0.
		return TRUE;
	}
}



BOOL VarToBOOL(Var &aVar)
{
	if (!aVar.HasContents()) // Must be checked first because otherwise IsNumeric() would consider "" to be non-numeric and thus TRUE.  For performance, it also exploits the binary number cache.
		return FALSE;
	switch (aVar.IsNumeric())
	{
	case PURE_INTEGER:
		return aVar.ToInt64() != 0;
	case PURE_FLOAT:
		return aVar.ToDouble() != 0.0;
	default:
		// Even a string containing all whitespace would be considered non-numeric since it's a non-blank string
		// that isn't equal to 0.
		return TRUE;
	}
}



BOOL TokenToBOOL(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_INTEGER: // Probably the most common; e.g. both sides of "if (x>3 and x<6)" are the number 1/0.
		return aToken.value_int64 != 0; // Force it to be purely 1 or 0.
	case SYM_VAR:
		return VarToBOOL(*aToken.var);
	case SYM_FLOAT:
		return aToken.value_double != 0.0;
	case SYM_STRING:
		return ResultToBOOL(aToken.marker);
	default:
		// The only remaining valid symbol is SYM_OBJECT, which is always TRUE.
		return TRUE;
	}
}



ToggleValueType TokenToToggleValue(ExprTokenType &aToken)
{
	if (TokenIsNumeric(aToken))
	switch (TokenToInt64(aToken)) // Reserve other values for potential future use by requiring exact match.
	{
	case 1: return TOGGLED_ON;
	case 0: return TOGGLED_OFF;
	case -1: return TOGGLE;
	}
	return TOGGLE_INVALID;
}



SymbolType TokenIsNumeric(ExprTokenType &aToken)
{
	switch(aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aToken.symbol;
	case SYM_VAR: 
		return aToken.var->IsNumeric();
	default: // SYM_STRING: Callers of this function expect a "numeric" result for numeric strings.
		return IsNumeric(aToken.marker, true, false, true);
	}
}


SymbolType TokenIsPureNumeric(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aToken.symbol;
	case SYM_VAR:
		if (!aToken.var->IsUninitializedNormalVar()) // Caller doesn't want a warning, so avoid calling Contents().
			return aToken.var->IsPureNumeric();
		//else fall through:
	default:
		return PURE_NOT_NUMERIC;
	}
}


SymbolType TokenIsPureNumeric(ExprTokenType &aToken, SymbolType &aNumType)
// This function is called very frequently by ExpandExpression(), which needs to distinguish
// between numeric strings and pure numbers, but still needs to know if the string is numeric.
{
	switch(aToken.symbol)
	{
	case SYM_INTEGER:
	case SYM_FLOAT:
		return aNumType = aToken.symbol;
	case SYM_VAR:
		if (aToken.var->IsUninitializedNormalVar()) // Caller doesn't want a warning, so avoid calling Contents().
			return aNumType = PURE_NOT_NUMERIC; // i.e. empty string is non-numeric.
		if (aNumType = aToken.var->IsPureNumeric())
			return aNumType; // This var contains a pure binary number.
		// Otherwise, it might be a numeric string (i.e. impure).
		aNumType = aToken.var->IsNumeric(); // This also caches the PURE_NOT_NUMERIC result if applicable.
		return PURE_NOT_NUMERIC;
	case SYM_STRING:
		aNumType = IsNumeric(aToken.marker, true, false, true);
		return PURE_NOT_NUMERIC;
	default:
		// Only SYM_OBJECT and SYM_MISSING should be possible.
		return aNumType = PURE_NOT_NUMERIC;
	}
}


BOOL TokenIsEmptyString(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_STRING:
		return !*aToken.marker;
	case SYM_VAR:
		return !aToken.var->HasContents();
	//case SYM_MISSING: // This case is omitted because it currently should be
		// impossible for all callers except for ParamIndexIsOmittedOrEmpty(),
		// which checks for it explicitly.
		//return TRUE;
	default:
		return FALSE;
	}
}


__int64 TokenToInt64(ExprTokenType &aToken)
// Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
// Converts the contents of aToken to a 64-bit int.
{
	// Some callers, such as those that cast our return value to UINT, rely on the use of 64-bit
	// to preserve unsigned values and also wrap any signed values into the unsigned domain.
	switch (aToken.symbol)
	{
		case SYM_INTEGER:	return aToken.value_int64;
		case SYM_FLOAT:		return (__int64)aToken.value_double;
		case SYM_VAR:		return aToken.var->ToInt64();
		case SYM_STRING:	return ATOI64(aToken.marker);
	}
	// Since above didn't return, it can only be SYM_OBJECT or not an operand.
	return 0;
}



double TokenToDouble(ExprTokenType &aToken, BOOL aCheckForHex)
// Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
// Converts the contents of aToken to a double.
{
	switch (aToken.symbol)
	{
		case SYM_FLOAT:		return aToken.value_double;
		case SYM_INTEGER:	return (double)aToken.value_int64;
		case SYM_VAR:		return aToken.var->ToDouble();
		case SYM_STRING:	return aCheckForHex ? ATOF(aToken.marker) : _tstof(aToken.marker); // atof() omits the check for hexadecimal.
	}
	// Since above didn't return, it can only be SYM_OBJECT or not an operand.
	return 0;
}



LPTSTR TokenToString(ExprTokenType &aToken, LPTSTR aBuf, size_t *aLength)
// Returns "" on failure to simplify logic in callers.  Otherwise, it returns either aBuf (if aBuf was needed
// for the conversion) or the token's own string.  aBuf may be NULL, in which case the caller presumably knows
// that this token is SYM_STRING or SYM_VAR (or the caller wants "" back for anything other than those).
// If aBuf is not NULL, caller has ensured that aBuf is at least MAX_NUMBER_SIZE in size.
{
	LPTSTR result;
	switch (aToken.symbol)
	{
	case SYM_VAR: // Caller has ensured that any SYM_VAR's Type() is VAR_NORMAL.
		result = aToken.var->Contents(); // Contents() vs. mCharContents in case mCharContents needs to be updated by Contents().
		if (aLength)
			*aLength = aToken.var->Length();
		return result;
	case SYM_STRING:
		result = aToken.marker;
		if (aLength)
		{
			if (aToken.marker_length == -1)
				break; // Call _tcslen(result) below.
			*aLength = aToken.marker_length;
		}
		return result;
	case SYM_INTEGER:
		result = aBuf ? ITOA64(aToken.value_int64, aBuf) : _T("");
		break;
	case SYM_FLOAT:
		if (aBuf)
		{
			int length = FTOA(aToken.value_double, aBuf, MAX_NUMBER_SIZE);
			if (aLength)
				*aLength = length;
			return aBuf;
		}
		//else continue on to return the default at the bottom.
	//case SYM_OBJECT: // Treat objects as empty strings (or TRUE where appropriate).
	//case SYM_MISSING:
	default:
		result = _T("");
	}
	if (aLength) // Caller wants to know the string's length as well.
		*aLength = _tcslen(result);
	return result;
}



ResultType TokenToDoubleOrInt64(const ExprTokenType &aInput, ExprTokenType &aOutput)
// Converts aToken's contents to a numeric value, either int64 or double (whichever is more appropriate).
// Returns FAIL when aToken isn't an operand or is but contains a string that isn't purely numeric.
{
	LPTSTR str;
	switch (aInput.symbol)
	{
		case SYM_INTEGER:
		case SYM_FLOAT:
			aOutput.symbol = aInput.symbol;
			aOutput.value_int64 = aInput.value_int64;
			return OK;
		case SYM_VAR:
			return aInput.var->ToDoubleOrInt64(aOutput);
		case SYM_STRING:   // v1.0.40.06: Fixed to be listed explicitly so that "default" case can return failure.
			str = aInput.marker;
			break;
		//case SYM_OBJECT: // L31: Treat objects as empty strings (or TRUE where appropriate).
		//case SYM_MISSING:
		default:
			return FAIL;
	}
	// Since above didn't return, interpret "str" as a number.
	switch (aOutput.symbol = IsNumeric(str, true, false, true))
	{
	case PURE_INTEGER:
		aOutput.value_int64 = ATOI64(str);
		break;
	case PURE_FLOAT:
		aOutput.value_double = ATOF(str);
		break;
	default: // Not a pure number.
		return FAIL;
	}
	return OK; // Since above didn't return, indicate success.
}



StringCaseSenseType TokenToStringCase(ExprTokenType& aToken)
{
	// Pure integers 1 and 0 corresponds to SCS_SENSITIVE and SCS_INSENSITIVE, respectively.
	// Pure floats returns SCS_INVALID.
	// For strings see Line::ConvertStringCaseSense.
	LPTSTR str = NULL;
	__int64 int_val = 0;
	switch (aToken.symbol)
	{
	case SYM_VAR:
		
		switch (aToken.var->IsPureNumeric())
		{
		case PURE_INTEGER: int_val = aToken.var->ToInt64(); break;
		case PURE_NOT_NUMERIC: str = aToken.var->Contents(); break;
		case PURE_FLOAT: 
		default:	
			return SCS_INVALID;
		}
		break;

	case SYM_INTEGER: int_val = TokenToInt64(aToken); break;
	case SYM_FLOAT: return SCS_INVALID;
	default: str = TokenToString(aToken); break;
	}
	if (str)
		return !_tcsicmp(str, _T("Logical"))	? SCS_INSENSITIVE_LOGICAL
												: Line::ConvertStringCaseSense(str);
	return int_val == 1 ? SCS_SENSITIVE						// 1	- Sensitive
						: (int_val == 0 ? SCS_INSENSITIVE	// 0	- Insensitive
										: SCS_INVALID);		// else - invalid.
}



Var *TokenToOutputVar(ExprTokenType &aToken)
{
	if (aToken.symbol == SYM_VAR && !VARREF_IS_READ(aToken.var_usage)) // VARREF_ISSET is tolerated for use by IsSet().
		return aToken.var;
	return dynamic_cast<VarRef *>(TokenToObject(aToken));
}



IObject *TokenToObject(ExprTokenType &aToken)
// L31: Returns IObject* from SYM_OBJECT or SYM_VAR (where var->HasObject()), NULL for other tokens.
// Caller is responsible for calling AddRef() if that is appropriate.
{
	if (aToken.symbol == SYM_OBJECT)
		return aToken.object;
	
	if (aToken.symbol == SYM_VAR)
		return aToken.var->ToObject();

	return NULL;
}



ResultType ValidateFunctor(IObject *aFunc, int aParamCount, ResultToken &aResultToken, int *aUseMinParams, bool aShowError)
{
	ASSERT(aFunc);
	__int64 min_params = 0, max_params = INT_MAX;
	auto min_result = aParamCount == -1 ? INVOKE_NOT_HANDLED
		: GetObjectIntProperty(aFunc, _T("MinParams"), min_params, aResultToken, true);
	if (!min_result)
		return FAIL;
	bool has_minparams = min_result != INVOKE_NOT_HANDLED; // For readability.
	if (aUseMinParams) // CallbackCreate's signal to default to MinParams.
	{
		if (!has_minparams)
			return aShowError ? aResultToken.UnknownMemberError(ExprTokenType(aFunc), IT_GET, _T("MinParams")) : CONDITION_FALSE;
		*aUseMinParams = aParamCount = (int)min_params;
	}
	else if (has_minparams && aParamCount < (int)min_params)
		return aShowError ? aResultToken.ValueError(ERR_INVALID_FUNCTOR) : CONDITION_FALSE;
	auto max_result = (aParamCount <= 0 || has_minparams && min_params == aParamCount)
		? INVOKE_NOT_HANDLED // No need to check MaxParams in the above cases.
		: GetObjectIntProperty(aFunc, _T("MaxParams"), max_params, aResultToken, true);
	if (!max_result)
		return FAIL;
	if (max_result != INVOKE_NOT_HANDLED && aParamCount > (int)max_params)
	{
		__int64 is_variadic = 0;
		auto result = GetObjectIntProperty(aFunc, _T("IsVariadic"), is_variadic, aResultToken, true);
		if (!result)
			return FAIL;
		if (!is_variadic) // or not defined.
			return aShowError ? aResultToken.ValueError(ERR_INVALID_FUNCTOR) : CONDITION_FALSE;
	}
	// If either MinParams or MaxParams was confirmed to exist, this is likely a valid
	// function object, so skip the following check for performance.  Otherwise, catch
	// likely errors by checking that the object is callable.
	if (min_result == INVOKE_NOT_HANDLED && max_result == INVOKE_NOT_HANDLED)
		if (Object *obj = dynamic_cast<Object *>(aFunc))
			if (!obj->HasMethod(_T("Call")))
				return aShowError ? aResultToken.UnknownMemberError(ExprTokenType(aFunc), IT_CALL, _T("Call")) : CONDITION_FALSE;
		// Otherwise: COM objects can be callable via DISPID_VALUE.  There's probably
		// no way to determine whether the object supports that without invoking it.
	return OK;
}



ResultType TokenSetResult(ResultToken &aResultToken, LPCTSTR aValue, size_t aLength)
// Utility function for handling memory allocation and return to callers of built-in functions; based on BIF_SubStr.
// Returns FAIL if malloc failed, in which case our caller is responsible for returning a sensible default value.
{
	if (aLength == -1)
		aLength = _tcslen(aValue); // Caller must not pass NULL for aResult in this case.
	if (aLength <= MAX_NUMBER_LENGTH) // v1.0.46.01: Avoid malloc() for small strings.  However, this improves speed by only 10% in a test where random 25-byte strings were extracted from a 700 KB string (probably because VC++'s malloc()/free() are very fast for small allocations).
		aResultToken.marker = aResultToken.buf; // Store the address of the result for the caller.
	else
	{
		// Caller has provided a mem_to_free (initially NULL) as a means of passing back memory we allocate here.
		// So if we change "result" to be non-NULL, the caller will take over responsibility for freeing that memory.
		if (   !(aResultToken.mem_to_free = tmalloc(aLength + 1))   ) // Out of memory.
			return aResultToken.MemoryError();
		aResultToken.marker = aResultToken.mem_to_free; // Store the address of the result for the caller.
	}
	if (aValue) // Caller may pass NULL to retrieve a buffer of sufficient size.
		tmemcpy(aResultToken.marker, aValue, aLength);
	aResultToken.marker[aLength] = '\0'; // Must be done separately from the memcpy() because the memcpy() might just be taking a substring (i.e. long before result's terminator).
	aResultToken.marker_length = aLength;
	return OK;
}



// TypeOfToken: Similar result to TokenIsPureNumeric, but may return SYM_OBJECT.
SymbolType TypeOfToken(ExprTokenType &aToken)
{
	switch (aToken.symbol)
	{
	case SYM_VAR:
		switch (aToken.var->IsPureNumericOrObject())
		{
		case VAR_ATTRIB_IS_INT64: return SYM_INTEGER;
		case VAR_ATTRIB_IS_DOUBLE: return SYM_FLOAT;
		case VAR_ATTRIB_IS_OBJECT: return SYM_OBJECT;
		}
		return SYM_STRING;
	default: // Providing a default case produces smaller code on release builds as the compiler can omit the other symbol checks.
#ifdef _DEBUG
		MsgBox(_T("DEBUG: Unhandled symbol type."));
#endif
	case SYM_STRING:
	case SYM_INTEGER:
	case SYM_FLOAT:
	case SYM_OBJECT:
	case SYM_MISSING:
		return aToken.symbol;
	}
}



ResultType ResultToken::Return(LPTSTR aValue, size_t aLength)
// Copy and return a string.
{
	ASSERT(aValue);
	symbol = SYM_STRING;
	return TokenSetResult(*this, aValue, aLength);
}



int ConvertJoy(LPTSTR aBuf, int *aJoystickID, bool aAllowOnlyButtons)
// The caller TextToKey() currently relies on the fact that when aAllowOnlyButtons==true, a value
// that can fit in a sc_type (USHORT) is returned, which is true since the joystick buttons
// are very small numbers (JOYCTRL_1==12).
{
	if (aJoystickID)
		*aJoystickID = 0;  // Set default output value for the caller.
	if (!aBuf || !*aBuf) return JOYCTRL_INVALID;
	LPTSTR aBuf_orig = aBuf;
	for (; *aBuf >= '0' && *aBuf <= '9'; ++aBuf); // self-contained loop to find the first non-digit.
	if (aBuf > aBuf_orig) // The string starts with a number.
	{
		int joystick_id = ATOI(aBuf_orig) - 1;
		if (joystick_id < 0 || joystick_id >= MAX_JOYSTICKS)
			return JOYCTRL_INVALID;
		if (aJoystickID)
			*aJoystickID = joystick_id;  // Use ATOI vs. atoi even though hex isn't supported yet.
	}

	if (!_tcsnicmp(aBuf, _T("Joy"), 3))
	{
		if (IsNumeric(aBuf + 3, false, false))
		{
			int offset = ATOI(aBuf + 3);
			if (offset < 1 || offset > MAX_JOY_BUTTONS)
				return JOYCTRL_INVALID;
			return JOYCTRL_1 + offset - 1;
		}
	}
	if (aAllowOnlyButtons)
		return JOYCTRL_INVALID;

	// Otherwise:
	if (!_tcsicmp(aBuf, _T("JoyX"))) return JOYCTRL_XPOS;
	if (!_tcsicmp(aBuf, _T("JoyY"))) return JOYCTRL_YPOS;
	if (!_tcsicmp(aBuf, _T("JoyZ"))) return JOYCTRL_ZPOS;
	if (!_tcsicmp(aBuf, _T("JoyR"))) return JOYCTRL_RPOS;
	if (!_tcsicmp(aBuf, _T("JoyU"))) return JOYCTRL_UPOS;
	if (!_tcsicmp(aBuf, _T("JoyV"))) return JOYCTRL_VPOS;
	if (!_tcsicmp(aBuf, _T("JoyPOV"))) return JOYCTRL_POV;
	if (!_tcsicmp(aBuf, _T("JoyName"))) return JOYCTRL_NAME;
	if (!_tcsicmp(aBuf, _T("JoyButtons"))) return JOYCTRL_BUTTONS;
	if (!_tcsicmp(aBuf, _T("JoyAxes"))) return JOYCTRL_AXES;
	if (!_tcsicmp(aBuf, _T("JoyInfo"))) return JOYCTRL_INFO;
	return JOYCTRL_INVALID;
}



bool ScriptGetKeyState(vk_type aVK, KeyStateTypes aKeyStateType)
// Returns true if "down", false if "up".
{
    if (!aVK) // Assume "up" if indeterminate.
		return false;

	switch (aKeyStateType)
	{
	case KEYSTATE_TOGGLE: // Whether a toggleable key such as CapsLock is currently turned on.
		return IsKeyToggledOn(aVK); // This also works for non-"lock" keys, but in that case the toggle state can be out of sync with other processes/threads.
	case KEYSTATE_PHYSICAL: // Physical state of key.
		if (IsMouseVK(aVK)) // mouse button
		{
			if (g_MouseHook) // mouse hook is installed, so use it's tracking of physical state.
				return g_PhysicalKeyState[aVK] & STATE_DOWN;
			else
				return IsKeyDownAsync(aVK);
		}
		else // keyboard
		{
			if (g_KeybdHook)
			{
				// Since the hook is installed, use its value rather than that from
				// GetAsyncKeyState(), which doesn't seem to return the physical state.
				// But first, correct the hook modifier state if it needs it.  See comments
				// in GetModifierLRState() for why this is needed:
				if (KeyToModifiersLR(aVK))    // It's a modifier.
					GetModifierLRState(true); // Correct hook's physical state if needed.
				return g_PhysicalKeyState[aVK] & STATE_DOWN;
			}
			else
				return IsKeyDownAsync(aVK);
		}
	} // switch()

	// Otherwise, use the default state-type: KEYSTATE_LOGICAL
	// On XP/2K at least, a key can be physically down even if it isn't logically down,
	// which is why the below specifically calls IsKeyDown() rather than some more
	// comprehensive method such as consulting the physical key state as tracked by the hook:
	// v1.0.42.01: For backward compatibility, the following hasn't been changed to IsKeyDownAsync().
	// One example is the journal playback hook: when a window owned by the script receives
	// such a keystroke, only GetKeyState() can detect the changed state of the key, not GetAsyncKeyState().
	// A new mode can be added to KeyWait & GetKeyState if Async is ever explicitly needed.
	return IsKeyDown(aVK);
}



bool ScriptGetJoyState(JoyControls aJoy, int aJoystickID, ExprTokenType &aToken, LPTSTR aBuf)
// Caller must ensure that aToken.marker is a buffer large enough to handle the longest thing put into
// it here, which is currently jc.szPname (size=32). Caller has set aToken.symbol to be SYM_STRING.
// aToken is used for the value being returned by GetKeyState() to the script, while this function's
// bool return value is used only by KeyWait, so is false for "up" and true for "down".
// If there was a problem determining the position/state, aToken is made blank and false is returned.
{
	// Set default in case of early return.
	aToken.symbol = SYM_STRING;
	aToken.marker = aBuf;
	*aBuf = '\0'; // Blank vs. string "0" serves as an indication of failure.

	if (!aJoy) // Currently never called this way.
		return false; // And leave aToken set to blank.

	bool aJoy_is_button = IS_JOYSTICK_BUTTON(aJoy);

	JOYCAPS jc;
	if (!aJoy_is_button && aJoy != JOYCTRL_POV)
	{
		// Get the joystick's range of motion so that we can report position as a percentage.
		if (joyGetDevCaps(aJoystickID, &jc, sizeof(JOYCAPS)) != JOYERR_NOERROR)
			ZeroMemory(&jc, sizeof(jc));  // Zero it on failure, for use of the zeroes later below.
	}

	// Fetch this struct's info only if needed:
	JOYINFOEX jie;
	if (aJoy != JOYCTRL_NAME && aJoy != JOYCTRL_BUTTONS && aJoy != JOYCTRL_AXES && aJoy != JOYCTRL_INFO)
	{
		jie.dwSize = sizeof(JOYINFOEX);
		jie.dwFlags = JOY_RETURNALL;
		if (joyGetPosEx(aJoystickID, &jie) != JOYERR_NOERROR)
			return false; // And leave aToken set to blank.
		if (aJoy_is_button)
		{
			bool is_down = ((jie.dwButtons >> (aJoy - JOYCTRL_1)) & (DWORD)0x01);
			aToken.symbol = SYM_INTEGER; // Override default type.
			aToken.value_int64 = is_down; // Always 1 or 0, since it's "bool" (and if not for that, bitwise-and).
			return is_down;
		}
	}

	// Otherwise:
	UINT range;
	LPTSTR buf_ptr;
	double result_double;  // Not initialized to help catch bugs.

	switch(aJoy)
	{
	case JOYCTRL_XPOS:
		range = (jc.wXmax > jc.wXmin) ? jc.wXmax - jc.wXmin : 0;
		result_double = range ? 100 * (double)jie.dwXpos / range : jie.dwXpos;
		break;
	case JOYCTRL_YPOS:
		range = (jc.wYmax > jc.wYmin) ? jc.wYmax - jc.wYmin : 0;
		result_double = range ? 100 * (double)jie.dwYpos / range : jie.dwYpos;
		break;
	case JOYCTRL_ZPOS:
		range = (jc.wZmax > jc.wZmin) ? jc.wZmax - jc.wZmin : 0;
		result_double = range ? 100 * (double)jie.dwZpos / range : jie.dwZpos;
		break;
	case JOYCTRL_RPOS:  // Rudder or 4th axis.
		range = (jc.wRmax > jc.wRmin) ? jc.wRmax - jc.wRmin : 0;
		result_double = range ? 100 * (double)jie.dwRpos / range : jie.dwRpos;
		break;
	case JOYCTRL_UPOS:  // 5th axis.
		range = (jc.wUmax > jc.wUmin) ? jc.wUmax - jc.wUmin : 0;
		result_double = range ? 100 * (double)jie.dwUpos / range : jie.dwUpos;
		break;
	case JOYCTRL_VPOS:  // 6th axis.
		range = (jc.wVmax > jc.wVmin) ? jc.wVmax - jc.wVmin : 0;
		result_double = range ? 100 * (double)jie.dwVpos / range : jie.dwVpos;
		break;

	case JOYCTRL_POV:
		aToken.symbol = SYM_INTEGER; // Override default type.
		if (jie.dwPOV == JOY_POVCENTERED) // Need to explicitly compare against JOY_POVCENTERED because it's a WORD not a DWORD.
		{
			aToken.value_int64 = -1;
			return false;
		}
		else
		{
			aToken.value_int64 = jie.dwPOV;
			return true;
		}
		// No break since above always returns.

	case JOYCTRL_NAME:
		_tcscpy(aToken.marker, jc.szPname);
		return false; // Return value not used.

	case JOYCTRL_BUTTONS:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumButtons; // wMaxButtons is the *driver's* max supported buttons.
		return false; // Return value not used.

	case JOYCTRL_AXES:
		aToken.symbol = SYM_INTEGER; // Override default type.
		aToken.value_int64 = jc.wNumAxes; // wMaxAxes is the *driver's* max supported axes.
		return false; // Return value not used.

	case JOYCTRL_INFO:
		buf_ptr = aToken.marker;
		if (jc.wCaps & JOYCAPS_HASZ)
			*buf_ptr++ = 'Z';
		if (jc.wCaps & JOYCAPS_HASR)
			*buf_ptr++ = 'R';
		if (jc.wCaps & JOYCAPS_HASU)
			*buf_ptr++ = 'U';
		if (jc.wCaps & JOYCAPS_HASV)
			*buf_ptr++ = 'V';
		if (jc.wCaps & JOYCAPS_HASPOV)
		{
			*buf_ptr++ = 'P';
			if (jc.wCaps & JOYCAPS_POV4DIR)
				*buf_ptr++ = 'D';
			if (jc.wCaps & JOYCAPS_POVCTS)
				*buf_ptr++ = 'C';
		}
		*buf_ptr = '\0'; // Final termination.
		return false; // Return value not used.
	} // switch()

	// If above didn't return, the result should now be in result_double.
	aToken.symbol = SYM_FLOAT; // Override default type.
	aToken.value_double = result_double;
	return result_double;
}



__int64 pow_ll(__int64 base, __int64 exp)
{
	/*
	Caller must ensure exp >= 0
	Below uses 'a^b' to denote 'raising a to the power of b'.
	Computes and returns base^exp. If the mathematical result doesn't fit in __int64, the result is undefined.
	By convention, x^0 returns 1, even when x == 0,	caller should ensure base is non-zero when exp is zero to handle 0^0.
	*/
	if (exp == 0)
		return 1ll;

	// based on: https://en.wikipedia.org/wiki/Exponentiation_by_squaring (2018-11-03)
	__int64 result = 1;
	while (exp > 1)
	{
		if (exp % 2) // exp is odd
			result *= base;
		base *= base;
		exp /= 2;
	}
	return result * base;
}
