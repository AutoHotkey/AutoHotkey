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
#include <shlwapi.h> // StrCmpLogicalW



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
	free(item);
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
