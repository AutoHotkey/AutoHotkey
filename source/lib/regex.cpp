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

#define PCRE_STATIC             // For RegEx. PCRE_STATIC tells PCRE to declare its functions for normal, static
#include "lib_pcre/pcre/pcre.h" // linkage rather than as functions inside an external DLL.

#include "script_func_impl.h"



ResultType RegExCreateMatchArray(LPCTSTR haystack, pcret *re, pcret_extra *extra, int *offset, int pattern_count, int captured_pattern_count, IObject *&match_object)
{
	// For lookup performance, create a table of subpattern names indexed by subpattern number.
	LPCTSTR *subpat_name = NULL; // Set default as "no subpattern names present or available".
	bool allow_dupe_subpat_names = false; // Set default.
	LPCTSTR name_table;
	int name_count, name_entry_size;
	if (   !pcret_fullinfo(re, extra, PCRE_INFO_NAMECOUNT, &name_count) // Success. Fix for v1.0.45.01: Don't check captured_pattern_count>=0 because PCRE_ERROR_NOMATCH can still have named patterns!
		&& name_count // There's at least one named subpattern.  Relies on short-circuit boolean order.
		&& !pcret_fullinfo(re, extra, PCRE_INFO_NAMETABLE, &name_table) // Success.
		&& !pcret_fullinfo(re, extra, PCRE_INFO_NAMEENTRYSIZE, &name_entry_size)   ) // Success.
	{
		int pcre_options;
		if (!pcret_fullinfo(re, extra, PCRE_INFO_OPTIONS, &pcre_options)) // Success.
			allow_dupe_subpat_names = pcre_options & PCRE_DUPNAMES;
		// For indexing simplicity, also include an entry for the main/entire pattern at index 0 even though
		// it's never used because the entire pattern can't have a name without enclosing it in parentheses
		// (in which case it's not the entire pattern anymore, but in fact subpattern #1).
		size_t subpat_array_size = pattern_count * sizeof(LPCSTR);
		subpat_name = (LPCTSTR *)_alloca(subpat_array_size); // See other use of _alloca() above for reasons why it's used.
		ZeroMemory(subpat_name, subpat_array_size); // Set default for each index to be "no name corresponds to this subpattern number".
		for (int i = 0; i < name_count; ++i, name_table += name_entry_size)
		{
			// Below converts first two bytes of each name-table entry into the pattern number (it might be
			// possible to simplify this, but I'm not sure if big vs. little-endian will ever be a concern).
#ifdef UNICODE
			subpat_name[name_table[0]] = name_table + 1;
#else
			subpat_name[(name_table[0] << 8) + name_table[1]] = name_table + 2; // For indexing simplicity, subpat_name[0] is for the main/entire pattern though it is never actually used for that because it can't be named without being enclosed in parentheses (in which case it becomes a subpattern).
#endif
			// For simplicity and unlike PHP, IsNumeric() isn't called to forbid numeric subpattern names.
			// It seems the worst than could happen if it is numeric is that it would overlap/overwrite some of
			// the numerically-indexed elements in the output-array.  Seems pretty harmless given the rarity.
		}
	}
	//else one of the pcre_fullinfo() calls may have failed.  The PCRE docs indicate that this realistically never
	// happens unless bad inputs were given.  So due to rarity, just leave subpat_name==NULL; i.e. "no named subpatterns".

	LPTSTR mark = (extra->flags & PCRE_EXTRA_MARK) ? (LPTSTR)*extra->mark : NULL;
	return RegExMatchObject::Create(haystack, offset, subpat_name, pattern_count, captured_pattern_count, mark, match_object);
}


ResultType RegExMatchObject::Create(LPCTSTR aHaystack, int *aOffset, LPCTSTR *aPatternName
	, int aPatternCount, int aCapturedPatternCount, LPCTSTR aMark, IObject *&aNewObject)
{
	aNewObject = NULL;

	// If there was no match, seems best to not return an object:
	if (aCapturedPatternCount < 1)
		return OK;

	RegExMatchObject *m = new RegExMatchObject();
	if (!m)
		return FAIL;
	m->SetBase(sPrototype);

	if (  aMark && !(m->mMark = _tcsdup(aMark))  )
	{
		m->Release();
		return FAIL;
	}

	ASSERT(aCapturedPatternCount >= 1);
	ASSERT(aPatternCount >= aCapturedPatternCount);

	// Use aPatternCount vs aCapturedPatternCount since we want to be able to retrieve the
	// names of *all* subpatterns, even ones that weren't captured.  For instance, a loop
	// converting the object to an old-style pseudo-array would need to initialize even the
	// array items that weren't captured.
	m->mPatternCount = aPatternCount;
	
	// Allocate memory for a copy of the offset array.
	if (  !(m->mOffset = (int *)malloc(aPatternCount * 2 * sizeof(int *)))  )
	{
		m->Release();
		return FAIL;
	}
	// memcpy currently benchmarks slightly faster on x64 than copying offsets in the loop below:
	memcpy(m->mOffset, aOffset, aPatternCount * 2 * sizeof(int));

	// Do some pre-processing:
	//  - Locate the smallest portion of haystack that contains all matches.
	//  - Convert end offsets to lengths.
	int p, min_offset = INT_MAX, max_offset = -1;
	for (p = 0; p < aCapturedPatternCount; ++p)
	{
		if (m->mOffset[p*2] > -1)
		{
			// Substring is non-empty, so ensure we copy this portion of haystack.
			if (min_offset > m->mOffset[p*2])
				min_offset = m->mOffset[p*2];
			if (max_offset < m->mOffset[p*2+1])
				max_offset = m->mOffset[p*2+1];
		}
		// Convert end offset to length.
		m->mOffset[p*2+1] -= m->mOffset[p*2];
	}
	// Initialize the remainder of the offset vector (patterns which were not captured),
	// which have indeterminate values if we're called by a regex callout and therefore
	// can't be handled the same way as in the loop above.
	for ( ; p < aPatternCount; ++p)
	{
		m->mOffset[p*2] = -1;
		m->mOffset[p*2+1] = 0;
	}
	
	// Copy only the portion of aHaystack which contains matches.  This can be much faster
	// than copying the whole string for larger haystacks.  For instance, searching for "GNU"
	// in the GPL v2 (18120 chars) and producing a match object is about 5 times faster with
	// this optimization than without if caller passes us the haystack length, and about 50
	// times faster than the old code which used _tcsdup().  However, the difference is much
	// smaller for small strings.
	if (min_offset < max_offset) // There are non-empty matches.
	{
		int our_haystack_size = (max_offset - min_offset) + 1;
		if (  !(m->mHaystack = tmalloc(our_haystack_size))  )
		{
			m->Release();
			return FAIL;
		}
		tmemcpy(m->mHaystack, aHaystack + min_offset, our_haystack_size);
		m->mHaystackStart = min_offset;
	}

	// Copy subpattern names.
	if (aPatternName)
	{
		// Allocate array of pointers.
		if (  !(m->mPatternName = (LPTSTR *)malloc(aPatternCount * sizeof(LPTSTR *)))  )
		{
			m->Release(); // Also frees other things allocated above.
			return FAIL;
		}

		// Copy names and initialize array.
		m->mPatternName[0] = NULL;
		for (int p = 1; p < aPatternCount; ++p)
			if (aPatternName[p])
				// A failed allocation here seems rare and the consequences would be
				// negligible, so in that case just act as if the subpattern has no name.
				m->mPatternName[p] = _tcsdup(aPatternName[p]);
			else
				m->mPatternName[p] = NULL;
	}

	// Since above didn't return, the object has been set up successfully.
	aNewObject = m;
	return OK;
}


void RegExMatchObject::Invoke(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	switch (aID)
	{
	case M_Count: _o_return(mPatternCount - 1);
	case M_Mark: _o_return(mMark ? mMark : _T(""));
	case M___Enum: _o_return(new IndexEnumerator(this, ParamIndexToOptionalInt(0, 1)
		, static_cast<IndexEnumerator::Callback>(&RegExMatchObject::GetEnumItem)));
	}

	int p;
	if (ParamIndexIsOmitted(0))
		p = 0;
	else if (ParamIndexIsNumeric(0))
		p = ParamIndexToInt(0);
	else if (!mPatternName) // There are no named subpatterns, so param 0 is invalid.
		p = -1;
	else
	{
		auto name = ParamIndexToString(0);
		for (p = 0; p < mPatternCount; ++p)
			if (mPatternName[p] && !_tcsicmp(mPatternName[p], name))
			{
				if (mOffset[2*p] < 0)
					// This pattern wasn't matched, so check for one with a duplicate name.
					for (int i = p + 1; i < mPatternCount; ++i)
						if (mPatternName[i] && !_tcsicmp(mPatternName[i], name) // It has the same name.
							&& mOffset[2*i] >= 0) // It matched something.
						{
							// Prefer this pattern.
							p = i;
							break;
						}
				break;
			}
	}
	if (p < 0 || p >= mPatternCount)
		_o_throw_param(0); // p != 0 implies the parameter was not omitted.

	switch (aID)
	{
	case M___Get:
		if (auto arr = dynamic_cast<Array*>(ParamIndexToObject(1)))
			if (arr->Length())
				_o_throw(ERR_INVALID_USAGE);
			// Otherwise, fall through:
	// Gives the correct result even if there was no match (because length is 0):
	case M_Value: _o_return(mHaystack - mHaystackStart + mOffset[p*2], mOffset[p*2+1]);
	case M_Pos: _o_return(mOffset[2*p] + 1);
	case M_Len: _o_return(mOffset[2*p + 1]);
	case M_Name: _o_return((mPatternName && mPatternName[p]) ? mPatternName[p] : _T(""));
	}
}


struct RegExCalloutData // L14: Used by BIF_RegEx to pass necessary info to RegExCallout.
{
	pcret *re;
	LPTSTR re_text; // original NeedleRegEx
	int options_length; // used to adjust cb->pattern_position
	int pattern_count; // to save calling pcre_fullinfo unnecessarily for each callout
	pcret_extra *extra;
	ResultToken *result_token;
};

int RegExCallout(pcret_callout_block *cb)
{
	// Continuing execution on the hook thread wouldn't be safe, but there's no need to check
	// the following since cb->callout_data is non-null only when the regex is being evaluated
	// by RegExMatch/RegExReplace:
	//if (GetCurrentThreadId() != g_MainThreadID)
	//	return 0;

	if (!cb->callout_data) // Callout not coming from RegExMatch/RegExReplace.
		return 0;
	RegExCalloutData &cd = *(RegExCalloutData *)cb->callout_data;

	// Callout functions must be resolved each time, since patterns are cached, and the scope
	// may be different each time the pattern is executed.  Aside from potentially having the
	// wrong nested function, AddRef/Release would be needed in places to support closures.
	auto callout_name = *cb->user_callout ? cb->user_callout : _T("pcre_callout");
	auto callout_var = g_script.FindVar(callout_name, 0, FINDVAR_FOR_READ);
	auto callout_func = callout_var ? callout_var->ToObject() : nullptr;
	if (!callout_func)
	{
		cd.result_token->ValueError(_T("Invalid callout"), callout_name);
		return PCRE_ERROR_CALLOUT;
	}

	// Adjust offset to account for options, which are excluded from the regex passed to PCRE.
	cb->pattern_position += cd.options_length;
	
	// Save EventInfo to be restored when we return.
	EventInfoType EventInfo_saved = g->EventInfo;

	g->EventInfo = (EventInfoType) cb;
	
	/*
	callout_number:		should be available since callout number can be specified within (?C...).
	subject:			useful when behaviour might depend on text surrounding a capture.
	start_match:		as above. equivalent to return value of RegExMatch, so should be available somehow.
	
	pattern_position:	useful to debug regexes when combined with auto-callouts. otherwise not useful.
	next_item_length:	as above. combined 'next_item' instead of these two would be less useful as it cannot distinguish between multiple identical items, and would sometimes be empty.
	
	capture_top:		not sure if useful? helps to distinguish between empty capture and non-capture. could maybe use callout number to determine this instead.
	capture_last:		as above.

	current_position:	can be derived from start_match and strlen(param1), or param1 itself if P option is used.
	offset_vector:		not very useful as same information available in local variables in more convenient form.
	callout_data:		not relevant, maybe use "user data" field of (RegExCalloutData*)callout_data if implemented.
	subject_length:		not useful, use strlen(subject).
	version:			not important.
	*/

	// Since matching is still in progress, these *should* be -1.
	// For maintainability and peace of mind, save them anyway:
	int offset[] = { cb->offset_vector[0], cb->offset_vector[1] };
		
	// Temporarily set these for use by the function below:
	cb->offset_vector[0] = cb->start_match;
	cb->offset_vector[1] = cb->current_position;
	if (cd.extra->flags & PCRE_EXTRA_MARK)
		*cd.extra->mark = UorA(wchar_t *, UCHAR *) cb->mark;
	
	IObject *match_object;
	if (!RegExCreateMatchArray(cb->subject, cd.re, cd.extra, cb->offset_vector, cd.pattern_count, cb->capture_top, match_object))
	{
		cd.result_token->MemoryError();
		return PCRE_ERROR_CALLOUT; // Abort.
	}

	// Restore to former offsets (probably -1):
	cb->offset_vector[0] = offset[0];
	cb->offset_vector[1] = offset[1];
	
	// Make all string positions one-based. UPDATE: offset_vector cannot be modified, so for consistency don't do this:
	//++cb->pattern_position;
	//++cb->start_match;
	//++cb->current_position;
	
	ExprTokenType param[] =
	{
		match_object,
		cb->callout_number,
		cb->start_match + 1, // FoundPos (distinct from Match.Pos, which hasn't been set yet)
		const_cast<LPTSTR>(cb->subject), // Haystack
		cd.re_text // NeedleRegEx
	};
	__int64 number_to_return;
	auto result = CallMethod(callout_func, callout_func, nullptr, param, _countof(param), &number_to_return);
	if (result == FAIL || result == EARLY_EXIT)
	{
		number_to_return = PCRE_ERROR_CALLOUT;
		cd.result_token->SetExitResult(result);
	}
	
	g->EventInfo = EventInfo_saved;

	// Behaviour of return values is defined by PCRE.
	return (int)number_to_return;
}

pcret *get_compiled_regex(LPCTSTR aRegEx, pcret_extra *&aExtra, int *aOptionsLength, ResultToken *aResultToken)
// Returns the compiled RegEx, or NULL on failure.
// This function is called by things other than built-in functions so it should be kept general-purpose.
// Upon failure, if aResultToken!=NULL:
//   - An exception is thrown with a descriptive message on failure.
//   - *aResultToken is set up to contain an empty string.
// Upon success, the following output parameters are set based on the options that were specified:
//    aGetPositionsNotSubstrings
//    aExtra
// L14: aOptionsLength is used by callouts to adjust cb->pattern_position to be relative to beginning of actual user-specified NeedleRegEx instead of string seen by PCRE.
{	
	if (!pcret_callout)
	{	// Ensure this is initialized, even for ::RegExMatch() (to allow (?C) in window title regexes).
		pcret_callout = &RegExCallout;
	}

	// While reading from or writing to the cache, don't allow another thread entry.  This is because
	// that thread (or this one) might write to the cache while the other one is reading/writing, which
	// could cause loss of data integrity (the hook thread can enter here via #HotIf WinActive/Exist & SetTitleMatchMode RegEx).
	// Together, Enter/LeaveCriticalSection reduce performance by only 1.4% in the tightest possible script
	// loop that hits the first cache entry every time.  So that's the worst case except when there's an actual
	// collision, in which case performance suffers more because internally, EnterCriticalSection() does a
	// wait/semaphore operation, which is more costly.
	// Finally, the code size of all critical-section features together is less than 512 bytes (uncompressed),
	// so like performance, that's not a concern either.
	EnterCriticalSection(&g_CriticalRegExCache); // Request ownership of the critical section. If another thread already owns it, this thread will block until the other thread finishes.

	// SET UP THE CACHE.
	// This is a very crude cache for linear search. Of course, hashing would be better in the sense that it
	// would allow the cache to get much larger while still being fast (I believe PHP caches up to 4096 items).
	// Binary search might not be such a good idea in this case due to the time required to find the right spot
	// to insert a new cache item (however, items aren't inserted often, so it might perform quite well until
	// the cache contained thousands of RegEx's, which is unlikely to ever happen in most scripts).
	struct pcre_cache_entry
	{
		// For simplicity (and thus performance), the entire RegEx pattern including its options is cached
		// is stored in re_raw and that entire string becomes the RegEx's unique identifier for the purpose
		// of finding an entry in the cache.  Technically, this isn't optimal because some options like Study
		// and aGetPositionsNotSubstrings don't alter the nature of the compiled RegEx.  However, the CPU time
		// required to strip off some options prior to doing a cache search seems likely to offset much of the
		// cache's benefit.  So for this reason, as well as rarity and code size issues, this policy seems best.
		LPTSTR re_raw;      // The RegEx's literal string pattern such as "abc.*123".
		pcret *re_compiled; // The RegEx in compiled form.
		pcret_extra *extra; // NULL unless a study() was done (and NULL even then if study() didn't find anything).
		// int pcre_options; // Not currently needed in the cache since options are implicitly inside re_compiled.
		int options_length; // Lexikos: See aOptionsLength comment at beginning of this function.
	};

	#define PCRE_CACHE_SIZE 100 // Going too high would be counterproductive due to the slowness of linear search (and also the memory utilization of so many compiled RegEx's).
	static pcre_cache_entry sCache[PCRE_CACHE_SIZE] = {{0}};
	static int sLastInsert, sLastFound = -1; // -1 indicates "cache empty".
	int insert_pos; // v1.0.45.03: This is used to avoid updating sLastInsert until an insert actually occurs (it might not occur if a compile error occurs in the regex, or something else stops it early).

	// CHECK IF THIS REGEX IS ALREADY IN THE CACHE.
	if (sLastFound == -1) // Cache is empty, so insert this RegEx at the first position.
		insert_pos = 0;  // A section further below will change sLastFound to be 0.
	else
	{
		// Search the cache to see if it contains the caller-specified RegEx in compiled form.
		// First check if the last-found item is a match, since often it will be (such as cases
		// where a script-loop executes only one RegEx, and also for SetTitleMatchMode RegEx).
		if (!_tcscmp(aRegEx, sCache[sLastFound].re_raw)) // Match found (case sensitive).
			goto match_found; // And no need to update sLastFound because it's already set right.

		// Since above didn't find a match, search outward in both directions from the last-found match.
		// A bidirectional search is done because consecutively-called regex's tend to be adjacent to each other
		// in the array, so performance is improved on average (since most of the time when repeating a previously
		// executed regex, that regex will already be in the cache -- so optimizing the finding behavior is
		// more important than optimizing the never-found-because-not-cached behavior).
		bool go_right;
		int i, item_to_check, left, right;
		int last_populated_item = (sCache[PCRE_CACHE_SIZE-1].re_compiled) // When the array is full...
			? PCRE_CACHE_SIZE - 1  // ...all items must be checked except the one already done earlier.
			: sLastInsert;         // ...else only the items actually populated need to be checked.

		for (go_right = true, left = sLastFound, right = sLastFound, i = 0
			; i < last_populated_item  // This limits it to exactly the number of items remaining to be checked.
			; ++i, go_right = !go_right)
		{
			if (go_right) // Proceed rightward in the array.
			{
				right = (right == last_populated_item) ? 0 : right + 1; // Increment or wrap around back to the left side.
				item_to_check = right;
			}
			else // Proceed leftward.
			{
				left = (left == 0) ? last_populated_item : left - 1; // Decrement or wrap around back to the right side.
				item_to_check = left;
			}
			if (!_tcscmp(aRegEx, sCache[item_to_check].re_raw)) // Match found (case sensitive).
			{
				sLastFound = item_to_check;
				goto match_found;
			}
		}

		// Since above didn't goto, no match was found nor is one possible.  So just indicate the insert position
		// for where this RegEx will be put into the cache.
		// The following formula is for both cache-full and cache-partially-full.  When the cache is full,
		// it might not be the best possible formula; but it seems pretty good because it takes a round-robin
		// approach to overwriting/discarding old cache entries.  A discarded entry might have just been
		// used -- or even be sLastFound itself -- but on average, this approach seems pretty good because a
		// script loop that uses 50 unique RegEx's will quickly stabilize in the cache so that all 50 of them
		// stay compiled/cached until the loop ends.
		insert_pos = (sLastInsert == PCRE_CACHE_SIZE-1) ? 0 : sLastInsert + 1; // Formula works for both full and partially-full array.
	}
	// Since the above didn't goto:
	// - This RegEx isn't yet in the cache.  So compile it and put it in the cache, then return it to caller.
	// - Above is responsible for having set insert_pos to the cache position where the new RegEx will be stored.

	// The following macro is for maintainability, to enforce the definition of "default" in multiple places.
	// PCRE_NEWLINE_CRLF is the default in AutoHotkey rather than PCRE_NEWLINE_LF because *multiline* haystacks
	// that scripts will use are expected to come from:
	// 50%: FileRead: Uses `r`n by default, for performance)
	// 10%: Clipboard: Normally uses `r`n (includes files copied from Explorer, text data, etc.)
	// 20%: Download: Testing shows that it varies: e.g. microsoft.com uses `r`n, but `n is probably
	//      more common due to FTP programs automatically translating CRLF to LF when uploading to UNIX servers.
	// 20%: Other sources such as GUI edit controls: It's fairly unusual to want to use RegEx on multiline data
	//      from GUI controls, but in such case `n is much more common than `r`n.
#ifdef UNICODE
	#define AHK_PCRE_CHARSET_OPTIONS (PCRE_UTF8 | PCRE_NO_UTF8_CHECK)
#else
	#define AHK_PCRE_CHARSET_OPTIONS 0
#endif
	#define SET_DEFAULT_PCRE_OPTIONS \
	{\
		pcre_options = AHK_PCRE_CHARSET_OPTIONS;\
		do_study = false;\
	}
	#define PCRE_NEWLINE_BITS (PCRE_NEWLINE_CRLF | PCRE_NEWLINE_ANY) // Covers all bits that are used for newline options.

	// SET DEFAULT OPTIONS:
	int pcre_options;
	long long do_study;
	SET_DEFAULT_PCRE_OPTIONS

	// PARSE THE OPTIONS (if any).
	LPCTSTR pat; // When options-parsing is done, pat will point to the start of the pattern itself.
	for (pat = aRegEx;; ++pat)
	{
		switch(*pat)
		{
		case 'i': pcre_options |= PCRE_CASELESS;  break;  // Perl-compatible options.
		case 'm': pcre_options |= PCRE_MULTILINE; break;  //
		case 's': pcre_options |= PCRE_DOTALL;    break;  //
		case 'x': pcre_options |= PCRE_EXTENDED;  break;  //
		case 'A': pcre_options |= PCRE_ANCHORED;  break;      // PCRE-specific options (uppercase used by convention, even internally by PCRE itself).
		case 'D': pcre_options |= PCRE_DOLLAR_ENDONLY; break; //
		case 'J': pcre_options |= PCRE_DUPNAMES;       break; //
		case 'U': pcre_options |= PCRE_UNGREEDY;       break; //
		case 'X': pcre_options |= PCRE_EXTRA;          break; //
		case 'C': pcre_options |= PCRE_AUTO_CALLOUT;   break; // L14: PCRE_AUTO_CALLOUT causes callouts to be created with callout_number == 255 before each item in the pattern.
		case '\a':
			// Enable matching of any kind of newline, including Unicode newline characters.
			// v2: \R doesn't match Unicode newlines by default, so `a also enables that.
			pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_ANY | PCRE_BSR_UNICODE;
			break; 
		case '\n':pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_LF; break; // See below.
			// Above option: Could alternatively have called it "LF" rather than or in addition to "`n", but that
			// seems slightly less desirable due to potential overlap/conflict with future option letters,
			// plus the fact that `n should be pretty well known to AutoHotkey users, especially advanced ones
			// using RegEx.  Note: `n`r is NOT treated the same as `r`n because there's a slight chance PCRE
			// will someday support `n`r for some obscure usage (or just for symmetry/completeness).
			// The PCRE_NEWLINE_XXX options are valid for both compile() and exec(), but specifying it for exec()
			// would only serve to override the default stored inside the compiled pattern (seems rarely needed).
		case '\r':
			if (pat[1] == '\n')
			{
				++pat; // Skip over the second character so that it's not recognized as a separate option by the next iteration.
				pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_CRLF; // Set explicitly in case it was unset by an earlier option. Remember that PCRE_NEWLINE_CRLF is a bitwise combination of PCRE_NEWLINE_LF and CR.
			}
			else // For completeness, it's easy to support PCRE_NEWLINE_CR too, though nowadays I think it's quite rare (former Macintosh format).
				pcre_options = (pcre_options & ~PCRE_NEWLINE_BITS) | PCRE_NEWLINE_CR; // Do it this way because PCRE_NEWLINE_CRLF is a bitwise combination of PCRE_NEWLINE_CR and PCRE_NEWLINE_LF.
			break;

		// Other options (uppercase so that lowercase can be reserved for future/PERL options):
		case 'S':
			do_study = true;
			break;

		case ' ':  // Allow only spaces and tabs as fillers so that everything else is protected/reserved for
		case '\t': // future use (such as future PERL options).
			break;

		case ')': // This character, when not escaped, marks the normal end of the options section.  We know it's not escaped because if it had been, the loop would have stopped at the backslash before getting here.
			++pat; // Set pat to be the start of the actual RegEx pattern, and leave options set to how they were by any prior iterations above.
			goto break_both;

		default: // Namely the following:
		//case '\0': No options are present, so ignore any letters that were accidentally recognized and treat entire string as the pattern.
		//case '(' : An open parenthesis must be considered an invalid option because otherwise it would be ambiguous with a subpattern.
		//case '\\': In addition to backslash being an invalid option, it also covers "\)" as being invalid (i.e. so that it's never necessary to check for an escaped close-parenthesis).
		//case all-other-chars: All others are invalid options; so like backslash above, ignore any letters that were accidentally recognized and treat entire string as the pattern.
			SET_DEFAULT_PCRE_OPTIONS // Revert to original options in case any early letters happened to match valid options.
			pat = aRegEx; // Indicate that the entire string is the pattern (no options).
			// To distinguish between a bad option and no options at all (for better error reporting), could check if
			// within the next few chars there's an unmatched close-parenthesis (non-escaped).  If so, the user
			// intended some options but one of them was invalid.  However, that would be somewhat difficult to do
			// because both \) and [)] are valid regex patterns that contain an unmatched close-parenthesis.
			// Since I don't know for sure if there are other cases (or whether future RegEx extensions might
			// introduce more such cases), it seems best not to attempt to distinguish.  Using more than two options
			// is extremely rare anyway, so syntax errors of this type do not happen often (and the only harm done
			// is a misleading error message from PCRE rather than something like "Bad option").  In addition,
			// omitting it simplifies the code and slightly improves performance.
			goto break_both;
		} // switch(*pat)
	} // for()

break_both:
	// Reaching here means that pat has been set to the beginning of the RegEx pattern itself and all options
	// are set properly.

	LPCSTR error_msg;
	TCHAR error_buf[128];
	int error_code, error_offset;
	pcret *re_compiled;

	// COMPILE THE REGEX.
	if (   !(re_compiled = pcret_compile2(pat, pcre_options, &error_code, &error_msg, &error_offset, NULL))   )
	{
		if (aResultToken) // A non-NULL value indicates our caller is RegExMatch() or RegExReplace() in a script.
		{
			sntprintf(error_buf, _countof(error_buf), _T("Compile error %d at offset %d: %hs"), error_code
				, error_offset, error_msg);
			// Seems best to bring the error to the user's attention rather than letting it potentially
			// escape their notice.  This sort of error should be corrected immediately, not handled
			// within the script (such as by try-catch).
			aResultToken->Error(error_buf);
		}
		goto error;
	}

	if (do_study)
	{
		// Enabling JIT compilation adds about 68 KB to the final executable size, which seems to outweigh
		// the speed-up that a minority of scripts would get.  Pass the option anyway, in case it is enabled:
		aExtra = pcret_study(re_compiled, PCRE_STUDY_JIT_COMPILE, &error_msg); // aExtra is an output parameter for caller.
		// Above returns NULL on failure or inability to find anything worthwhile in its study.  NULL is exactly
		// the right value to pass to exec() to indicate "no study info".
		// The following isn't done because:
		// 1) It seems best not to abort the caller's RegEx operation just due to a study error, since the only
		//    error likely to happen (from looking at PCRE's source code) is out-of-memory.
		// 2) Reduced code size.
		//if (error_msg)
		//{
			//if (aResultToken)
			//{
			//	sntprintf(error_buf, sizeof(error_buf), "Study error: %s", error_msg);
			//	aResultToken->Error(error_buf);
			//}
			//goto error;
		//}
	}
	else // No studying desired.
		aExtra = NULL; // aExtra is an output parameter for caller.

	// ADD THE NEWLY-COMPILED REGEX TO THE CACHE.
	// An earlier stage has set insert_pos to be the desired insert-position in the cache.
	pcre_cache_entry &this_entry = sCache[insert_pos]; // For performance and convenience.
	if (this_entry.re_compiled) // An existing cache item is being overwritten, so free it's attributes.
	{
		// Free the old cache entry's attributes in preparation for overwriting them with the new one's.
		free(this_entry.re_raw);           // Free the uncompiled pattern.
		pcret_free(this_entry.re_compiled); // Free the compiled pattern.
		if (this_entry.extra)
			pcret_free_study(this_entry.extra);
	}
	//else the insert-position is an empty slot, which is usually the case because most scripts contain fewer than
	// PCRE_CACHE_SIZE unique regex's.  Nothing extra needs to be done.
	this_entry.re_raw = _tcsdup(aRegEx); // _strdup() is very tiny and basically just calls _tcslen+malloc+_tcscpy.
	this_entry.re_compiled = re_compiled;
	this_entry.extra = aExtra;
	// "this_entry.pcre_options" doesn't exist because it isn't currently needed in the cache.  This is
	// because the RE's options are implicitly stored inside re_compiled.

	// Lexikos: See aOptionsLength comment at beginning of this function.
	this_entry.options_length = (int)(pat - aRegEx);

	if (aOptionsLength) 
		*aOptionsLength = this_entry.options_length;

	sLastInsert = insert_pos; // v1.0.45.03: Must be done only *after* the insert succeeded because some things rely on sLastInsert being synonymous with the last populated item in the cache (when the cache isn't yet full).
	sLastFound = sLastInsert; // Relied upon in the case where sLastFound==-1. But it also sets things up to start the search at this item next time, because it's a bit more likely to be found here such as tight loops containing only one RegEx.
	// Remember that although sLastFound==sLastInsert in this case, it isn't always so -- namely when a previous
	// call found an existing match in the cache without having to compile and insert the item.

	LeaveCriticalSection(&g_CriticalRegExCache);
	return re_compiled; // Indicate success.

match_found: // RegEx was found in the cache at position sLastFound, so return the cached info back to the caller.
	aExtra = sCache[sLastFound].extra;
	if (aOptionsLength) // Lexikos: See aOptionsLength comment at beginning of this function.
		*aOptionsLength = sCache[sLastFound].options_length; 

	LeaveCriticalSection(&g_CriticalRegExCache);
	return sCache[sLastFound].re_compiled; // Indicate success.

error: // Since NULL is returned here, caller should ignore the contents of the output parameters.
	LeaveCriticalSection(&g_CriticalRegExCache);
	return NULL; // Indicate failure.
}



LPCTSTR RegExMatch(LPCTSTR aHaystack, LPCTSTR aNeedleRegEx)
// Returns NULL if no match.  Otherwise, returns the address where the pattern was found in aHaystack.
{
	pcret_extra *extra;
	pcret *re;

	// Compile the regex or get it from cache.
	if (   !(re = get_compiled_regex(aNeedleRegEx, extra, NULL, NULL))   ) // Compiling problem.
		return NULL; // Our callers just want there to be "no match" in this case.

	// Set up the offset array, which consists of int-pairs containing the start/end offset of each match.
	// For simplicity, use a fixed size because even if it's too small (unlikely for our types of callers),
	// PCRE will still operate properly (though it returns 0 to indicate the too-small condition).
	#define RXM_INT_COUNT 30  // Should be a multiple of 3.
	int offset[RXM_INT_COUNT];

	// Execute the regex.
	int captured_pattern_count = pcret_exec(re, extra, aHaystack, (int)_tcslen(aHaystack), 0, 0, offset, RXM_INT_COUNT);
	if (captured_pattern_count < 0) // PCRE_ERROR_NOMATCH or some kind of error.
		return NULL;

	// Otherwise, captured_pattern_count>=0 (it's 0 when offset[] was too small; but that's harmless in this case).
	return aHaystack + offset[0]; // Return the position of the entire-pattern match.
}



void RegExReplace(ResultToken &aResultToken, ExprTokenType *aParam[], int aParamCount
	, pcret *aRE, pcret_extra *aExtra, LPTSTR aHaystack, int aHaystackLength
	, int aStartingOffset, int aOffset[], int aNumberOfIntsInOffset)
{
	// If an output variable was provided for the count, resolve it early in case of early goto.
	// Fix for v1.0.47.05: In the unlikely event that output_var_count is the same script-variable as
	// as the haystack, needle, or replacement (i.e. the same memory), don't set output_var_count until
	// immediately prior to returning.  Otherwise, haystack, needle, or replacement would corrupted while
	// it's still being used here.
	Var *output_var_count = ParamIndexToOutputVar(3);
	int replacement_count = 0; // This value will be stored in output_var_count, but only at the very end due to the reason above.

	// Get the replacement text (if any) from the incoming parameters.  If it was omitted, treat it as "".
	TCHAR repl_buf[MAX_NUMBER_SIZE];
	LPTSTR replacement = _T("");
	ResultToken result_token;
	ExprTokenType matchobj_token, *params;
	IObject *callback_obj = nullptr;
	result_token.mem_to_free = nullptr;
	if (!ParamIndexIsOmitted(2))
	{
		if (callback_obj = ParamIndexToObject(2))
		{
			if (!ValidateFunctor(callback_obj, 1, aResultToken))
				return;
			params = &matchobj_token;
			matchobj_token.symbol = SYM_OBJECT;
			result_token.InitResult(repl_buf);
		}
		else
			replacement = ParamIndexToOptionalString(2, repl_buf);
	}

	// In PCRE, lengths and such are confined to ints, so there's little reason for using unsigned for anything.
	int captured_pattern_count, empty_string_is_not_a_match, match_length, ref_num
		, result_size, new_result_length, haystack_portion_length, second_iteration, substring_name_length
		, extra_offset, pcre_options;
	TCHAR *haystack_pos, *match_pos, *src, *src_orig, *closing_brace, *substring_name_pos;
	TCHAR *dest, char_after_dollar
		, substring_name[33] // In PCRE, "Names consist of up to 32 alphanumeric characters and underscores."
		, transform;

	// Caller has provided mem_to_free (initially NULL) as a means of passing back memory we allocate here.
	// So if we change "result" to be non-NULL, the caller will take over responsibility for freeing that memory.
	LPTSTR &result = aResultToken.mem_to_free; // Make an alias for convenience.
	size_t &result_length = aResultToken.marker_length; // MANDATORY FOR USERS OF MEM_TO_FREE: set marker_length to the length of the string.
	result_size = 0;   // And caller has already set "result" to be NULL.  The buffer is allocated only upon
	result_length = 0; // first use to avoid a potentially massive allocation that might be wasted and cause swapping (not to mention that we'll have better ability to estimate the correct total size after the first replacement is discovered).

	// Below uses a temp variable because realloc() returns NULL on failure but leaves original block allocated.
	// Note that if it's given a NULL pointer, realloc() does a malloc() instead.
	LPTSTR realloc_temp;
	#define REGEX_REALLOC(size) \
	{\
		result_size = size;\
		if (   !(realloc_temp = trealloc(result, result_size))   )\
			goto out_of_mem;\
		result = realloc_temp;\
	}

	// See if a replacement limit was specified.  If not, use the default (-1 means "replace all").
	int limit = -1;
	if (!ParamIndexIsOmitted(4))
	{
		Throw_if_Param_NaN(4);
		limit = ParamIndexToInt(4);
	}

	// aStartingOffset is altered further on in the loop; but for its initial value, the caller has ensured
	// that it lies within aHaystackLength.  Also, if there are no replacements yet, haystack_pos ignores
	// aStartingOffset because otherwise, when the first replacement occurs, any part of haystack that lies
	// to the left of a caller-specified aStartingOffset wouldn't get copied into the result.
	for (empty_string_is_not_a_match = 0, haystack_pos = aHaystack
		;; haystack_pos = aHaystack + aStartingOffset) // See comment above.
	{
		// Execute the expression to find the next match.
		captured_pattern_count = (limit == 0) ? PCRE_ERROR_NOMATCH // Only when limit is exactly 0 are we done replacing.  All negative values are "replace all".
			: pcret_exec(aRE, aExtra, aHaystack, aHaystackLength, aStartingOffset
				, empty_string_is_not_a_match, aOffset, aNumberOfIntsInOffset);

		if (captured_pattern_count == PCRE_ERROR_NOMATCH)
		{
			if (empty_string_is_not_a_match && aStartingOffset < aHaystackLength && limit != 0) // replacement_count>0 whenever empty_string_is_not_a_match!=0.
			{
				// This situation happens when a previous iteration found a match but it was the empty string.
				// That iteration told the pcre_exec that just occurred above to try to match something other than ""
				// at the same position.  But since we're here, it wasn't able to find such a match.  So just copy
				// the current character over literally then advance to the next character to resume normal searching.
				empty_string_is_not_a_match = 0; // Reset so that the next iteration starts off with the normal matching method.
#ifdef UNICODE
				// Need to avoid chopping a supplementary Unicode character in half.
				WCHAR c = haystack_pos[0];
				if (IS_SURROGATE_PAIR(c, haystack_pos[1])) // i.e. one supplementary character.
				{
					result[result_length++] = c;
					result[result_length++] = haystack_pos[1];
					aStartingOffset += 2; // Supplementary characters are in the range U+010000 to U+10FFFF,
					continue;
				}
#endif
				result[result_length++] = *haystack_pos; // This can't overflow because the size calculations in a previous iteration reserved 3 bytes: 1 for this character, 1 for the possible LF that follows CR, and 1 for the terminator.
				++aStartingOffset; // Advance to next candidate section of haystack.
				// v1.0.46.06: This following section was added to avoid finding a match between a CR and LF
				// when PCRE_NEWLINE_ANY mode is in effect.  The fact that this is the only change for
				// PCRE_NEWLINE_ANY relies on the belief that any pattern that matches the empty string in between
				// a CR and LF must also match the empty string that occurs right before the CRLF (even if that
				// pattern also matched a non-empty string right before the empty one in front of the CRLF).  If
				// this belief is correct, no logic similar to this is needed near the bottom of the main loop
				// because the empty string found immediately prior to this CRLF will put us into
				// empty_string_is_not_a_match mode, which will then execute this section of code (unless
				// empty_string_is_not_a_match mode actually found a match, in which case the logic here seems
				// superseded by that match?)  Even if this reasoning is not a complete solution, it might be
				// adequate if patterns that match empty strings are rare, which I believe they are.  In fact,
				// they might be so rare that arguably this could be documented as a known limitation rather than
				// having added the following section of code in the first place.
				// Examples that illustrate the effect:
				//    MsgBox % "<" . RegExReplace("`r`n", "`a).*", "xxx") . ">"
				//    MsgBox % "<" . RegExReplace("`r`n", "`am)^.*$", "xxx") . ">"
				if (*haystack_pos == '\r' && haystack_pos[1] == '\n')
				{
					// pcre_fullinfo() is a fast call, so it's called every time to simplify the code (I don't think
					// this whole "empty_string_is_not_a_match" section of code executes for most patterns anyway,
					// so performance seems less of a concern).
					if (!pcret_fullinfo(aRE, aExtra, PCRE_INFO_OPTIONS, &pcre_options) // Success.
						&& (pcre_options & PCRE_NEWLINE_ANY))
					{
						result[result_length++] = '\n'; // This can't overflow because the size calculations in a previous iteration reserved 3 bytes: 1 for this character, 1 for the possible LF that follows CR, and 1 for the terminator.
						++aStartingOffset; // Skip over this LF because it "belongs to" the CR that preceded it.
					}
				}
				continue; // i.e. we're not done yet because the "no match" above was a special one and there's still more haystack to check.
			}
			// Otherwise, there aren't any more matches.  So we're all done except for copying the last part of
			// haystack into the result (if applicable).
			if (replacement_count) // And by definition, result!=NULL due in this case to prior iterations.
			{
				if (haystack_portion_length = aHaystackLength - aStartingOffset) // This is the remaining part of haystack that needs to be copied over as-is.
				{
					new_result_length = (int)result_length + haystack_portion_length;
					if (new_result_length >= result_size)
						REGEX_REALLOC(new_result_length + 1); // This will end the loop if an alloc error occurs.
					tmemcpy(result + result_length, haystack_pos, haystack_portion_length); // memcpy() usually benches a little faster than _tcscpy().
					result_length = new_result_length; // Remember that result_length is actually an output for our caller, so even if for no other reason, it must be kept accurate for that.
				}
				result[result_length] = '\0'; // result!=NULL when replacement_count!=0.  Also, must terminate it unconditionally because other sections usually don't do it.
				// Set RegExMatch()'s return value to be "result":
				aResultToken.marker = result;  // Caller will take care of freeing result's memory.
			}
			else // No replacements were actually done, so just return the original string to avoid malloc+memcpy
				 // (in addition, returning the original might help the caller make other optimizations).
			{
				aResultToken.marker = aHaystack;
				aResultToken.marker_length = aHaystackLength;
				
				// There's no need to do the following because it should already be that way when replacement_count==0.
				//if (result)
				//	free(result);
				//result = NULL; // This tells the caller that we already freed it (i.e. from its POV, we never allocated anything).
			}
			aResultToken.symbol = SYM_STRING;
			goto set_count_and_return; // All done.
		}

		// Otherwise:
		if (captured_pattern_count < 0) // An error other than "no match". These seem very rare, so it seems best to abort rather than yielding a partially-converted result.
		{
			if (!aResultToken.Exited()) // Checked in case a callout already exited/raised an error.
			{
				ITOA(captured_pattern_count, repl_buf);
				aResultToken.Error(ERR_PCRE_EXEC, repl_buf);
			}
			goto abort; // Goto vs. break to leave replacement_count set to 0.
		}

		// Otherwise (since above didn't return or break or continue), a match has been found (i.e.
		// captured_pattern_count > 0; it should never be 0 in this case because that only happens
		// when offset[] is too small, which it isn't).
		++replacement_count;
		--limit; // It's okay if it goes below -1 because all negatives are treated as "replace all".
		match_pos = aHaystack + aOffset[0]; // This is the location in aHaystack of the entire-pattern match.
		int match_end_offset = aOffset[1];
		haystack_portion_length = (int)(match_pos - haystack_pos); // The length of the haystack section between the end of the previous match and the start of the current one.

		// Handle this replacement by making two passes through the replacement-text: The first calculates the size
		// (which avoids having to constantly check for buffer overflow with potential realloc at multiple stages).
		// The second iteration copies the replacement (along with any literal text in haystack before it) into the
		// result buffer (which was expanded if necessary by the first iteration).
		for (second_iteration = 0; second_iteration < 2; ++second_iteration) // second_iteration is used as a boolean for readability.
		{
			if (second_iteration)
			{
				// Using the required length calculated by the first iteration, expand/realloc "result" if necessary.
				if (new_result_length + 3 > result_size) // Must use +3 not +1 in case of empty_string_is_not_a_match (which needs room for up to two extra characters).
				{
					// The first expression passed to REGEX_REALLOC is the average length of each replacement so far.
					// It's more typically more accurate to pass that than the following "length of current
					// replacement":
					//    new_result_length - haystack_portion_length - (aOffset[1] - aOffset[0])
					// Above is the length difference between the current replacement text and what it's
					// replacing (it's negative when replacement is smaller than what it replaces).
					REGEX_REALLOC((int)PredictReplacementSize((new_result_length - match_end_offset) / replacement_count // See above.
						, replacement_count, limit, aHaystackLength, new_result_length+2, match_end_offset)); // +2 in case of empty_string_is_not_a_match (which needs room for up to two extra characters).  The function will also do another +1 to convert length to size (for terminator).
					// The above will return if an alloc error occurs.
				}
				//else result_size is not only large enough, but also non-zero.  Other sections rely on it always
				// being non-zero when replacement_count>0.

				// Before doing the actual replacement and its backreferences, copy over the part of haystack that
				// appears before the match.
				if (haystack_portion_length)
				{
					tmemcpy(result + result_length, haystack_pos, haystack_portion_length);
					result_length += haystack_portion_length;
				}
				dest = result + result_length; // Init dest for use by the loops further below.
			}
			else // i.e. it's the first iteration, so begin calculating the size required.
				new_result_length = (int)result_length + haystack_portion_length; // Init length to the part of haystack before the match (it must be copied over as literal text).

			// Calculate the actual replacement string through the callback function
			if (callback_obj)
			{
				if (second_iteration)
				{
					tmemcpy(dest, replacement, result_token.marker_length);
					result_length += result_token.marker_length;
					free(result_token.mem_to_free);
					result_token.mem_to_free = nullptr;
				}
				else
				{
					if (!RegExCreateMatchArray(aHaystack, aRE, aExtra, aOffset, static_cast<RegExCalloutData *>(aExtra->callout_data)->pattern_count, captured_pattern_count, matchobj_token.object))
						goto out_of_mem;
					result_token.SetValue(_T(""));
					callback_obj->Invoke(result_token, IT_CALL, nullptr, ExprTokenType{ callback_obj }, &params, 1);
					matchobj_token.object->Release();
					if (result_token.symbol == SYM_OBJECT)
					{
						auto obj = result_token.object;
						result_token.SetValue(_T(""));
						ObjectToString(result_token, ExprTokenType{ obj }, obj);
						obj->Release();
					}
					if (result_token.Exited())
					{
						aResultToken.SetExitResult(result_token.Result());
						goto abort;
					}
					replacement = TokenToString(result_token, repl_buf, &result_token.marker_length);
					new_result_length += (int)result_token.marker_length;
				}
				continue;
			}

			// DOLLAR SIGN ($) is the only method supported because it simplifies the code, improves performance,
			// and avoids the need to escape anything other than $ (which simplifies the syntax).
			for (src = replacement; ; ++src)  // For each '$' (increment to skip over the symbol just found by the inner for()).
			{
				// Find the next '$', if any.
				src_orig = src; // Init once for both loops below.
				if (second_iteration) // Mode: copy src-to-dest.
				{
					while (*src && *src != '$') // While looking for the next '$', copy over everything up until the '$'.
						*dest++ = *src++;
					result_length += (int)(src - src_orig);
				}
				else // This is the first iteration (mode: size-calculation).
				{
					for (; *src && *src != '$'; ++src); // Find the next '$', if any.
					new_result_length += (int)(src - src_orig); // '$' or '\0' was found: same expansion either way.
				}
				if (!*src)  // Reached the end of the replacement text.
					break;  // Nothing left to do, so if this is the first major iteration, begin the second.

				// Otherwise, a '$' has been found.  Check if it's a backreference and handle it.
				// But first process any special flags that are present.
				transform = '\0'; // Set default. Indicate "no transformation".
				extra_offset = 0; // Set default. Indicate that there's no need to hop over an extra character.
				if (char_after_dollar = src[1]) // This check avoids calling ctoupper on '\0', which directly or indirectly causes an assertion error in CRT.
				{
					switch(char_after_dollar = ctoupper(char_after_dollar))
					{
					case 'U':
					case 'L':
					case 'T':
						transform = char_after_dollar;
						extra_offset = 1;
						char_after_dollar = src[2]; // Ignore the transform character for the purposes of backreference recognition further below.
						break;
					//else leave things at their defaults.
					}
				}
				//else leave things at their defaults.

				ref_num = INT_MIN; // Set default to "no valid backreference".  Use INT_MIN to virtually guaranty that anything other than INT_MIN means that something like a backreference was found (even if it's invalid, such as ${-5}).
				switch (char_after_dollar)
				{
				case '{':  // Found a backreference: ${...
					substring_name_pos = src + 2 + extra_offset;
					if (closing_brace = _tcschr(substring_name_pos, '}'))
					{
						if (substring_name_length = (int)(closing_brace - substring_name_pos))
						{
							if (substring_name_length < _countof(substring_name))
							{
								tcslcpy(substring_name, substring_name_pos, substring_name_length + 1); // +1 to convert length to size, which truncates the new string at the desired position.
								if (IsNumeric(substring_name, true, false, true)) // Seems best to allow floating point such as 1.0 because it will then get truncated to an integer.  It seems to rare that anyone would want to use floats as names.
									ref_num = _ttoi(substring_name); // Uses _ttoi() vs. ATOI to avoid potential overlap with non-numeric names such as ${0x5}, which should probably be considered a name not a number?  In other words, seems best not to make some names that start with numbers "special" just because they happen to be hex numbers.
								else // For simplicity, no checking is done to ensure it consists of the "32 alphanumeric characters and underscores".  Let pcre_get_stringnumber() figure that out for us.
									ref_num = pcret_get_first_set(aRE, substring_name, aOffset); // Returns a negative on failure, which when stored in ref_num is relied upon as an indicator.
							}
							//else it's too long, so it seems best (debatable) to treat it as a unmatched/unfound name, i.e. "".
							src = closing_brace; // Set things up for the next iteration to resume at the char after "${..}"
						}
						//else it's ${}, so do nothing, which in effect will treat it all as literal text.
					}
					//else unclosed '{': for simplicity, do nothing, which in effect will treat it all as literal text.
					break;

				case '$':  // i.e. Two consecutive $ amounts to one literal $.
					++src; // Skip over the first '$', and the loop's increment will skip over the second. "extra_offset" is ignored due to rarity and silliness.  Just transcribe things like $U$ as U$ to indicate the problem.
					break; // This also sets up things properly to copy a single literal '$' into the result.

				case '\0': // i.e. a single $ was found at the end of the string.
					break; // Seems best to treat it as literal (strictly speaking the script should have escaped it).

				default:
					if (char_after_dollar >= '0' && char_after_dollar <= '9') // Treat it as a single-digit backreference. CONSEQUENTLY, $15 is really $1 followed by a literal '5'.
					{
						ref_num = char_after_dollar - '0'; // $0 is the whole pattern rather than a subpattern.
						src += 1 + extra_offset; // Set things up for the next iteration to resume at the char after $d. Consequently, $19 is seen as $1 followed by a literal 9.
					}
					//else not a digit: do nothing, which treats a $x as literal text (seems ok since like $19, $name will never be supported due to ambiguity; only ${name}).
				} // switch (char_after_dollar)

				if (ref_num == INT_MIN) // Nothing that looks like backreference is present (or the very unlikely ${-2147483648}).
				{
					if (second_iteration)
					{
						*dest++ = *src;  // src is incremented by the loop.  Copy only one character because the enclosing loop will take care of copying the rest.
						++result_length; // Update the actual length.
					}
					else
						++new_result_length; // Update the calculated length.
					// And now the enclosing loop will take care of the characters beyond src.
				}
				else // Something that looks like a backreference was found, even if it's invalid (e.g. ${-5}).
				{
					// It seems to improve convenience and flexibility to transcribe a nonexistent backreference
					// as a "" rather than literally (e.g. putting a ${1} literally into the new string).  Although
					// putting it in literally has the advantage of helping debugging, it doesn't seem to outweigh
					// the convenience of being able to specify nonexistent subpatterns. MORE IMPORANTLY a subpattern
					// might not exist per se if it hasn't been matched, such as an "or" like (abc)|(xyz), at least
					// when it's the last subpattern, in which case it should definitely be treated as "" and not
					// copied over literally.  So that would have to be checked for if this is changed.
					if (ref_num >= 0 && ref_num < captured_pattern_count) // Treat ref_num==0 as reference to the entire-pattern's match.
					{
						int ref_num0 = aOffset[ref_num*2];
						int ref_num1 = aOffset[ref_num*2 + 1];
						match_length = ref_num1 - ref_num0;
						if (match_length)
						{
							if (second_iteration)
							{
								tmemcpy(dest, aHaystack + ref_num0, match_length);
								if (transform)
								{
									dest[match_length] = '\0'; // Terminate for use below (shouldn't cause overflow because REALLOC reserved space for terminator; nor should there be any need to undo the termination afterward).
									switch(transform)
									{
									case 'U': CharUpper(dest); break;
									case 'L': CharLower(dest); break;
									case 'T': StrToTitleCase(dest); break;
									}
								}
								dest += match_length;
								result_length += match_length;
							}
							else // First iteration.
								new_result_length += match_length;
						}
					}
					//else subpattern doesn't exist (or it's invalid such as ${-5}, so treat it as blank because:
					// 1) It's boosts script flexibility and convenience (at the cost of making it hard to detect
					//    script bugs, which would be assisted by transcribing ${999} as literal text rather than "").
					// 2) It simplifies the code.
					// 3) A subpattern might not exist per se if it hasn't been matched, such as "(abc)|(xyz)"
					//    (in which case only one of them is matched).  If such a thing occurs at the end
					//    of the RegEx pattern, captured_pattern_count might not include it.  But it seems
					//    pretty clear that it should be treated as "" rather than some kind of error condition.
				}
			} // for() (for each '$')
		} // for() (a 2-iteration for-loop)

		// If we're here, a match was found.
		// Technique and comments from pcredemo.c:
		// If the previous match was NOT an empty string, we can just start the next match at the end
		// of the previous one.
		// If the previous match WAS an empty string, we can't do that, as it would lead to an
		// infinite loop. Instead, a special call of pcre_exec() is made with the PCRE_NOTEMPTY and
		// PCRE_ANCHORED flags set. The first of these tells PCRE that an empty string is not a valid match;
		// other possibilities must be tried. The second flag restricts PCRE to one match attempt at the
		// initial string position. If this match succeeds, that means there are two valid matches at the
		// SAME position: one for the empty string, and other for a non-empty string after it.  BOTH of
		// these matches are considered valid, and BOTH are eligible for replacement by RegExReplace().
		//
		// The following may be one example of this concept:
		// In the string "xy", replace the pattern "x?" by "z".  The traditional/proper answer (achieved by
		// the logic here) is "zzyz" because: 1) The first x is replaced by z; 2) The empty string before y
		// is replaced by z; 3) the logic here applies PCRE_NOTEMPTY to search again at the same position, but
		// that search doesn't find a match; so the logic higher above advances to the next character (y) and
		// continues the search; it finds the empty string at the end of haystack, which is replaced by z.
		// On the other hand, maybe there's a better example than the above that explains what would happen
		// if PCRE_NOTEMPTY actually finds a match, or what would happen if this PCRE_NOTEMPTY method weren't
		// used at all (i.e. infinite loop as mentioned in the previous paragraph).
		// 
		// If this match is "" (length 0), then by definition we just found a match in normal mode, not
		// PCRE_NOTEMPTY mode (since that mode isn't capable of finding "").  Thus, empty_string_is_not_a_match
		// is currently 0.  If "" was just found (and replaced), now try to find a second match at the same
		// position, but one that isn't "". This is done by switching to an alternate mode and doing another
		// iteration. Otherwise (the match found above wasn't "") advance to next candidate section of haystack
		// and resume searching.
		// v1.0.48.04: Fixed line below to reset to 0 (if appropriate) so that it doesn't get stuck in
		// PCRE_NOTEMPTY-mode for one extra iteration.  Otherwise there are too few replacements (4 vs. 5)
		// in examples like:
		//    RegExReplace("ABC", "Z*|A", "x")
		empty_string_is_not_a_match = (aOffset[0] == aOffset[1]) ? PCRE_NOTEMPTY|PCRE_ANCHORED : 0;
		aStartingOffset = match_end_offset; // In either case, set starting offset to the candidate for the next search.
	} // for()

	// All paths above should return (or goto some other label), so execution should never reach here except
	// through goto:
out_of_mem:
	aResultToken.MemoryError();
abort:
	if (result)
	{
		free(result);  // Since result is probably an non-terminated string (not to mention an incompletely created result), it seems best to free it here to remove it from any further consideration by the caller.
		result = NULL; // Tell caller that it was freed.
	}
	// Now fall through to below so that count is set even for out-of-memory error.
set_count_and_return:
	free(result_token.mem_to_free);
	if (output_var_count)
		output_var_count->Assign(replacement_count); // v1.0.47.05: Must be done last in case output_var_count shares the same memory with haystack, needle, or replacement.
}



BIF_DECL(BIF_RegEx)
// This function is the initial entry point for both RegExMatch() and RegExReplace().
// Caller has set aResultToken.symbol to a default of SYM_INTEGER.
{
	if (ParamIndexToObject(0))
		_f_throw_param(0, _T("String"));
	if (ParamIndexToObject(1))
		_f_throw_param(1, _T("String"));

	bool mode_is_replace = _f_callee_id == FID_RegExReplace;
	LPTSTR needle = ParamIndexToString(1, _f_number_buf); // Caller has already ensured that at least two actual parameters are present.

	pcret_extra *extra;
	pcret *re;
	int options_length;

	// COMPILE THE REGEX OR GET IT FROM CACHE.
	if (   !(re = get_compiled_regex(needle, extra, &options_length, &aResultToken))   ) // Compiling problem.
		return; // It already set aResultToken for us.

	// Since compiling succeeded, get info about other parameters.
	TCHAR haystack_buf[MAX_NUMBER_SIZE];
	size_t temp_length;
	LPTSTR haystack = ParamIndexToString(0, haystack_buf, &temp_length); // Caller has already ensured that at least two actual parameters are present.
	int haystack_length = (int)temp_length;

	int param_index = mode_is_replace ? 5 : 3;
	int starting_offset;
	if (ParamIndexIsOmitted(param_index))
		starting_offset = 0; // The one-based starting position in haystack (if any).  Convert it to zero-based.
	else
	{
		Throw_if_Param_NaN(param_index);
		starting_offset = ParamIndexToInt(param_index);
		if (starting_offset <= 0) // Same convention as SubStr(): Treat negative StartingPos as a position relative to the end of the string.
		{
			starting_offset += haystack_length;
			if (starting_offset < 0)
				starting_offset = 0;
		}
		else if (starting_offset > haystack_length)
			// Although pcre_exec() seems to work properly even without this check, its absence would allow
			// the empty string to be found beyond the length of haystack, which could lead to problems and is
			// probably more trouble than its worth (assuming it has any worth -- perhaps for a pattern that
			// looks backward from itself; but that seems too rare to support and might create code that's
			// harder to maintain, especially in RegExReplace()).
			starting_offset = haystack_length; // Due to rarity of this condition, opt for simplicity: just point it to the terminator, which is in essence an empty string (which will cause result in "no match" except when searcing for "").
		else
			--starting_offset; // Convert to zero-based.
	}

	// SET UP THE OFFSET ARRAY, which consists of int-pairs containing the start/end offset of each match.
	int pattern_count;
	pcret_fullinfo(re, extra, PCRE_INFO_CAPTURECOUNT, &pattern_count); // The number of capturing subpatterns (i.e. all except (?:xxx) I think). Failure is not checked because it seems too unlikely in this case.
	++pattern_count; // Increment to include room for the entire-pattern match.
	int number_of_ints_in_offset = pattern_count * 3; // PCRE uses 3 ints for each (sub)pattern: 2 for offsets and 1 for its internal use.
	int *offset = (int *)_alloca(number_of_ints_in_offset * sizeof(int)); // _alloca() boosts performance and seems safe because subpattern_count would usually have to be ridiculously high to cause a stack overflow.

	// The following section supports callouts (?C) and (*MARK:NAME).
	LPTSTR mark;
	RegExCalloutData callout_data;
	callout_data.re = re;
	callout_data.re_text = needle;
	callout_data.options_length = options_length;
	callout_data.pattern_count = pattern_count;
	callout_data.result_token = &aResultToken;
	if (extra)
	{	// S (study) option was specified, use existing pcre_extra struct.
		extra->flags |= PCRE_EXTRA_CALLOUT_DATA | PCRE_EXTRA_MARK;	
	}
	else
	{	// Allocate a pcre_extra struct to pass callout_data.
		extra = (pcret_extra *)_alloca(sizeof(pcret_extra));
		extra->flags = PCRE_EXTRA_CALLOUT_DATA | PCRE_EXTRA_MARK;
	}
	// extra->callout_data is used to pass callout_data to PCRE.
	extra->callout_data = &callout_data;
	// callout_data.extra is used by RegExCallout, which only receives a pointer to callout_data.
	callout_data.extra = extra;
	// extra->mark is used by PCRE to return the NAME of a (*MARK:NAME), if encountered.
	extra->mark = UorA(wchar_t **, UCHAR **) &mark;

	if (mode_is_replace) // Handle RegExReplace() completely then return.
	{
		RegExReplace(aResultToken, aParam, aParamCount, re, extra, haystack, haystack_length
			, starting_offset, offset, number_of_ints_in_offset);
		return;
	}
	// OTHERWISE, THIS IS RegExMatch() not RegExReplace().

	// EXECUTE THE REGEX.
	int captured_pattern_count = pcret_exec(re, extra, haystack, haystack_length
		, starting_offset, 0, offset, number_of_ints_in_offset);

	int match_offset = 0; // Set default for no match/error cases below.

	// SET THE RETURN VALUE BASED ON THE RESULTS OF EXECUTING THE EXPRESSION.
	if (captured_pattern_count == PCRE_ERROR_NOMATCH)
	{
		aResultToken.value_int64 = 0;
		// BUT CONTINUE ON so that the output variable (if any) is fully reset (made blank).
	}
	else if (captured_pattern_count < 0) // An error other than "no match".
	{
		if (aResultToken.Exited()) // A callout exited/raised an error.
			return;
		TCHAR err_info[MAX_INTEGER_SIZE];
		ITOA(captured_pattern_count, err_info);
		aResultToken.Error(ERR_PCRE_EXEC, err_info);
	}
	else // Match found, and captured_pattern_count >= 0 (but should never be 0 in this case because that only happens when offset[] is too small, which it isn't).
	{
		match_offset = offset[0];
		aResultToken.value_int64 = match_offset + 1; // i.e. the position of the entire-pattern match is the function's return value.
	}

	Var *output_var = ParamIndexToOutputVar(2);
	if (!output_var)
		return;

	IObject *match_object;
	if (!RegExCreateMatchArray(haystack, re, extra, offset, pattern_count, captured_pattern_count, match_object))
		aResultToken.MemoryError();
	if (match_object)
		output_var->AssignSkipAddRef(match_object);
	else // Out-of-memory or there were no captured patterns.
		output_var->Assign();
}
