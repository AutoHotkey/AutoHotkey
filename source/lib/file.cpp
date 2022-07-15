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

#include "script_func_impl.h"



// As of 2019-09-29, noinline reduces code size by over 20KB on VC++ 2019.
// Prior to merging Util_CreateDir with this, it wasn't inlined.
DECLSPEC_NOINLINE
bool Line::FileCreateDir(LPTSTR aDirSpec, LPTSTR aCanModifyDirSpec)
{
	if (!aDirSpec || !*aDirSpec)
	{
		SetLastError(ERROR_INVALID_PARAMETER);
		return false;
	}

	DWORD attr = GetFileAttributes(aDirSpec);
	if (attr != 0xFFFFFFFF)  // aDirSpec already exists.
	{
		SetLastError(ERROR_ALREADY_EXISTS);
		return (attr & FILE_ATTRIBUTE_DIRECTORY) != 0; // Indicate success if it already exists as a dir.
	}

	// If it has a backslash, make sure all its parent directories exist before we attempt
	// to create this directory:
	LPTSTR last_backslash = _tcsrchr(aDirSpec, '\\');
	if (last_backslash > aDirSpec // v1.0.48.04: Changed "last_backslash" to "last_backslash > aDirSpec" so that an aDirSpec with a leading \ (but no other backslashes), such as \dir, is supported.
		&& last_backslash[-1] != ':') // v1.1.31.00: Don't attempt FileCreateDir("C:") since that's equivalent to either "C:\" or the working directory (which already exists), or FileCreateDir("\\?\C:") since it always fails.
	{
		LPTSTR parent_dir;
		if (aCanModifyDirSpec)
		{
			parent_dir = aDirSpec; // Caller provided a modifiable aDirSpec.
			*last_backslash = '\0'; // Temporarily terminate for parent directory.
		}
		else
		{
			// v1.1.31.00: Allocate a modifiable buffer to be used by all calls (supports long paths).
			parent_dir = (LPTSTR)_alloca((last_backslash - aDirSpec + 1) * sizeof(TCHAR));
			tcslcpy(parent_dir, aDirSpec, last_backslash - aDirSpec + 1); // Omits the last backslash.
		}
		bool exists = FileCreateDir(parent_dir, parent_dir); // Recursively create all needed ancestor directories.
		if (aCanModifyDirSpec)
			*last_backslash = '\\'; // Undo temporary termination.

		// v1.0.44: Fixed ErrorLevel being set to 1 when the specified directory ends in a backslash.  In such cases,
		// two calls were made to CreateDirectory for the same folder: the first without the backslash and then with
		// it.  Since the directory already existed on the second call, ErrorLevel was wrongly set to 1 even though
		// everything succeeded.  So now, when recursion finishes creating all the ancestors of this directory
		// our own layer here does not call CreateDirectory() when there's a trailing backslash because a previous
		// layer already did:
		if (!last_backslash[1] || !exists)
			return exists;
	}

	// The above has recursively created all parent directories of aDirSpec if needed.
	// Now we can create aDirSpec.
	return CreateDirectory(aDirSpec, NULL);
}



ResultType ConvertFileOptions(ResultToken &aResultToken, LPTSTR aOptions, UINT &codepage, bool &translate_crlf_to_lf, unsigned __int64 *pmax_bytes_to_load)
{
	for (LPTSTR next, cp = aOptions; cp && *(cp = omit_leading_whitespace(cp)); cp = next)
	{
		if (*cp == '\n')
		{
			translate_crlf_to_lf = true;
			// Rather than treating "`nxxx" as invalid or ignoring "xxx", let the delimiter be
			// optional for `n.  Treating "`nxxx" and "m1024`n" and "utf-8`n" as invalid would
			// require larger code, and would produce confusing error messages because the `n
			// isn't visible; e.g. "Invalid option. Specifically: utf-8"
			next = cp + 1; 
			continue;
		}
		// \n is included below to allow "m1024`n" and "utf-8`n" (see above).
		next = StrChrAny(cp, _T(" \t\n"));

		switch (ctoupper(*cp))
		{
		case 'M':
			if (pmax_bytes_to_load) // i.e. caller is FileRead.
			{
				*pmax_bytes_to_load = ATOU64(cp + 1); // Relies upon the fact that it ceases conversion upon reaching a space or tab.
				break;
			}
			// Otherwise, fall through to treat it as invalid:
		default:
			TCHAR name[12]; // Large enough for any valid encoding.
			if (next && (next - cp) < _countof(name))
			{
				// Create a temporary null-terminated copy.
				wmemcpy(name, cp, next - cp);
				name[next - cp] = '\0';
				cp = name;
			}
			if (!_tcsicmp(cp, _T("Raw")))
			{
				codepage = -1;
			}
			else
			{
				codepage = Line::ConvertFileEncoding(cp);
				if (codepage == -1 || cisdigit(*cp)) // Require "cp" prefix in FileRead/FileAppend options.
					return aResultToken.ValueError(ERR_INVALID_OPTION, cp);
			}
			break;
		} // switch()
	} // for()
	return OK;
}

BIF_DECL(BIF_FileRead)
{
	_f_param_string(aFilespec, 0);
	_f_param_string_opt(aOptions, 1);

	const DWORD DWORD_MAX = ~0;

	// Set default options:
	bool translate_crlf_to_lf = false;
	unsigned __int64 max_bytes_to_load = ULLONG_MAX; // By default, fail if the file is too large.  See comments near bytes_to_read below.
	UINT codepage = g->Encoding;

	if (!ConvertFileOptions(aResultToken, aOptions, codepage, translate_crlf_to_lf, &max_bytes_to_load))
		return; // It already displayed the error.

	_f_set_retval_p(_T(""), 0); // Set default.

	// It seems more flexible to allow other processes to read and write the file while we're reading it.
	// For example, this allows the file to be appended to during the read operation, which could be
	// desirable, especially it's a very large log file that would take a long time to read.
	// MSDN: "To enable other processes to share the object while your process has it open, use a combination
	// of one or more of [FILE_SHARE_READ, FILE_SHARE_WRITE]."
	HANDLE hfile = CreateFile(aFilespec, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING
		, FILE_FLAG_SEQUENTIAL_SCAN, NULL); // MSDN says that FILE_FLAG_SEQUENTIAL_SCAN will often improve performance
	if (hfile == INVALID_HANDLE_VALUE)      // in cases like these (and it seems best even if max_bytes_to_load was specified).
	{
		aResultToken.SetLastErrorMaybeThrow(true);
		return;
	}

	unsigned __int64 bytes_to_read = GetFileSize64(hfile);
	if (bytes_to_read == ULLONG_MAX) // GetFileSize64() failed.
	{
		aResultToken.SetLastErrorCloseAndMaybeThrow(hfile, true);
		return;
	}
	// In addition to imposing the limit set by the *M option, the following check prevents an error
	// caused by 64 to 32-bit truncation -- that is, a file size of 0x100000001 would be truncated to
	// 0x1, allowing the command to complete even though it should fail.  UPDATE: This check was never
	// sufficient since max_bytes_to_load could exceed DWORD_MAX on x64 (prior to v1.1.16).  It's now
	// checked separately below to try to match the documented behaviour (truncating the data only to
	// the caller-specified limit).
	if (bytes_to_read > max_bytes_to_load) // This is the limit set by the caller.
		bytes_to_read = max_bytes_to_load;
	// Fixed for v1.1.16: Show an error message if the file is larger than DWORD_MAX, otherwise the
	// truncation issue described above could occur.  Reading more than DWORD_MAX could be supported
	// by calling ReadFile() in a loop, but it seems unlikely that a script will genuinely want to
	// do this AND actually be able to allocate a 4GB+ memory block (having 4GB of total free memory
	// is usually not sufficient, perhaps due to memory fragmentation).
#ifdef _WIN64
	if (bytes_to_read > DWORD_MAX)
#else
	// Reserve 2 bytes to avoid integer overflow below.  Although any amount larger than 2GB is almost
	// guaranteed to fail at the malloc stage, that might change if we ever become large address aware.
	if (bytes_to_read > DWORD_MAX - sizeof(wchar_t))
#endif
	{
		CloseHandle(hfile);
		_f_throw_oom; // Using this instead of "File too large." to reduce code size, since this condition is very rare (and malloc succeeding would be even rarer).
	}

	if (!bytes_to_read && codepage != -1) // In RAW mode, always return a zero-byte Buffer.
	{
		aResultToken.SetLastErrorCloseAndMaybeThrow(hfile, false, 0); // Indicate success (a zero-length file results in an empty string).
		return;
	}

	LPBYTE output_buf = (LPBYTE)malloc(size_t(bytes_to_read + (bytes_to_read & 1) + sizeof(wchar_t)));
	if (!output_buf)
	{
		CloseHandle(hfile);
		_f_throw_oom;
	}

	DWORD bytes_actually_read;
	BOOL result = ReadFile(hfile, output_buf, (DWORD)bytes_to_read, &bytes_actually_read, NULL);
	g->LastError = GetLastError();
	CloseHandle(hfile);

	// Upon result==success, bytes_actually_read is not checked against bytes_to_read because it
	// shouldn't be different (result should have set to failure if there was a read error).
	// If it ever is different, a partial read is considered a success since ReadFile() told us
	// that nothing bad happened.

	if (result)
	{
		if (codepage != -1) // Text mode, not "RAW" mode.
		{
			codepage &= CP_AHKCP; // Convert to plain Win32 codepage (remove CP_AHKNOBOM, which has no meaning here).
			bool has_bom;
			if ( (has_bom = (bytes_actually_read >= 2 && output_buf[0] == 0xFF && output_buf[1] == 0xFE)) // UTF-16LE BOM
					|| codepage == CP_UTF16 ) // Covers FileEncoding UTF-16 and FileEncoding UTF-16-RAW.
			{
				#ifndef UNICODE
				#error FileRead UTF-16 to ANSI string not implemented.
				#endif
				LPWSTR text = (LPWSTR)output_buf;
				DWORD length = bytes_actually_read / sizeof(WCHAR);
				if (has_bom)
				{
					// Move the data to eliminate the byte order mark.
					// Seems likely to perform better than allocating new memory and copying to it.
					--length;
					wmemmove(text, text + 1, length);
				}
				text[length] = '\0'; // Ensure text is terminated where indicated.  Two bytes were reserved for this purpose.
				aResultToken.AcceptMem(text, length);
				output_buf = NULL; // Don't free it; caller will take over.
			}
			else
			{
				LPCSTR text = (LPCSTR)output_buf;
				DWORD length = bytes_actually_read;
				if (length >= 3 && output_buf[0] == 0xEF && output_buf[1] == 0xBB && output_buf[2] == 0xBF) // UTF-8 BOM
				{
					codepage = CP_UTF8;
					length -= 3;
					text += 3;
				}
#ifndef UNICODE
				if (codepage == CP_ACP || codepage == GetACP())
				{
					// Avoid any unnecessary conversion or copying by using our malloc'd buffer directly.
					// This should be worth doing since the string must otherwise be converted to UTF-16 and back.
					output_buf[bytes_actually_read] = 0; // Ensure text is terminated where indicated.
					aResultToken.AcceptMem((LPSTR)output_buf, bytes_actually_read);
					output_buf = NULL; // Don't free it; caller will take over.
				}
				else
				#error FileRead non-ACP-ANSI to ANSI string not fully implemented.
#endif
				{
					int wlen = MultiByteToWideChar(codepage, 0, text, length, NULL, 0);
					if (wlen > 0)
					{
						if (!TokenSetResult(aResultToken, NULL, wlen))
						{
							free(output_buf);
							return;
						}
						wlen = MultiByteToWideChar(codepage, 0, text, length, aResultToken.marker, wlen);
						aResultToken.symbol = SYM_STRING;
						aResultToken.marker[wlen] = 0;
						aResultToken.marker_length = wlen;
						if (!wlen)
							result = FALSE;
					}
				}
			}
			if (output_buf) // i.e. it wasn't "claimed" above.
				free(output_buf);
			if (translate_crlf_to_lf && aResultToken.marker_length)
			{
				// Since a larger string is being replaced with a smaller, there's a good chance the 2 GB
				// address limit will not be exceeded by StrReplace even if the file is close to the
				// 1 GB limit as described above:
				StrReplace(aResultToken.marker, _T("\r\n"), _T("\n"), SCS_SENSITIVE, UINT_MAX, -1, NULL, &aResultToken.marker_length);
			}
		}
		else // codepage == -1 ("RAW" mode)
		{
			// Return the buffer to our caller.
			aResultToken.Return(BufferObject::Create(output_buf, bytes_actually_read));
		}
	}
	else
	{
		// ReadFile() failed.  Since MSDN does not document what is in the buffer at this stage, or
		// whether bytes_to_read contains a valid value, it seems best to abort the entire operation
		// rather than try to return partial file contents.  An exception will indicate the failure.
		free(output_buf);
	}

	aResultToken.SetLastErrorMaybeThrow(!result);
}



BIF_DECL(BIF_FileAppend)
{
	size_t aBuf_length;
	_f_param_string(aBuf, 0, &aBuf_length);
	_f_param_string_opt(aFilespec, 1);
	_f_param_string_opt(aOptions, 2);

	IObject *aBuf_obj = ParamIndexToObject(0); // Allow a Buffer-like object.
	if (aBuf_obj)
	{
		size_t ptr;
		GetBufferObjectPtr(aResultToken, aBuf_obj, ptr, aBuf_length);
		if (aResultToken.Exited())
			return;
		aBuf = (LPTSTR)ptr;
	}
	else
		aBuf_length *= sizeof(TCHAR); // Convert to byte count.

	_f_set_retval_p(_T(""), 0); // For all non-throw cases.

	// The below is avoided because want to allow "nothing" to be written to a file in case the
	// user is doing this to reset it's timestamp (or create an empty file).
	//if (!aBuf || !*aBuf)
	//	_f_return_retval;

	// Use the read-file loop's current item if filename was explicitly left blank (i.e. not just
	// a reference to a variable that's blank):
	LoopReadFileStruct *aCurrentReadFile = (aParamCount < 2) ? g->mLoopReadFile : NULL;
	if (aCurrentReadFile)
		aFilespec = aCurrentReadFile->mWriteFileName;
	if (!*aFilespec) // Nothing to write to.
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK);

	TextStream *ts = aCurrentReadFile ? aCurrentReadFile->mWriteFile : NULL;
	bool file_was_already_open = ts;

#ifdef CONFIG_DEBUGGER
	if (*aFilespec == '*' && !aFilespec[1] && !aBuf_obj && g_Debugger.OutputStdOut(aBuf))
	{
		// StdOut has been redirected to the debugger, and this "FileAppend" call has been
		// fully handled by the call above, so just return.
		g->LastError = 0;
		_f_return_retval;
	}
#endif

	UINT codepage;

	// Check if the file needs to be opened.  This is done here rather than at the time the
	// loop first begins so that:
	// 1) Any options/encoding specified in the first FileAppend call can take effect.
	// 2) To avoid opening the file if the file-reading loop has zero iterations (i.e. it's
	//    opened only upon first actual use to help performance and avoid changing the
	//    file-modification time when no actual text will be appended).
	if (!file_was_already_open)
	{
		codepage = aBuf_obj ? -1 : g->Encoding; // Never default to BOM if a Buffer object was passed.
		bool translate_crlf_to_lf = false;
		if (!ConvertFileOptions(aResultToken, aOptions, codepage, translate_crlf_to_lf, NULL))
			return;

		DWORD flags = TextStream::APPEND | (translate_crlf_to_lf ? TextStream::EOL_CRLF : 0);
		
		ASSERT( (~CP_AHKNOBOM) == CP_AHKCP );
		// codepage may include CP_AHKNOBOM, in which case below will not add BOM_UTFxx flag.
		if (codepage == CP_UTF8)
			flags |= TextStream::BOM_UTF8;
		else if (codepage == CP_UTF16)
			flags |= TextStream::BOM_UTF16;
		else if (codepage != -1)
			codepage &= CP_AHKCP;

		// Open the output file (if one was specified).  Unlike the input file, this is not
		// a critical error if it fails.  We want it to be non-critical so that FileAppend
		// commands in the body of the loop will throw to indicate the problem:
		ts = new TextFile; // ts was already verified NULL via !file_was_already_open.
		if ( !ts->Open(aFilespec, flags, codepage) )
		{
			aResultToken.SetLastErrorMaybeThrow(true);
			delete ts; // Must be deleted explicitly!
			return;
		}
		if (aCurrentReadFile)
			aCurrentReadFile->mWriteFile = ts;
	}
	else
		codepage = ts->GetCodePage();

	// Write to the file:
	DWORD result = 1;
	if (aBuf_length)
	{
		if (codepage == -1 || aBuf_obj) // "RAW" mode.
			result = ts->Write((LPCVOID)aBuf, (DWORD)aBuf_length);
		else
			result = ts->Write(aBuf, DWORD(aBuf_length / sizeof(TCHAR)));
	}
	//else: aBuf is empty; we've already succeeded in creating the file and have nothing further to do.
	aResultToken.SetLastErrorMaybeThrow(result == 0);

	if (!aCurrentReadFile)
		delete ts;
	// else it's the caller's responsibility, or it's caller's, to close it.
}



BOOL FileDeleteCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData)
{
	return DeleteFile(aFilename);
}

ResultType Line::FileDelete(LPTSTR aFilePattern)
{
	if (!*aFilePattern)
		return LineError(ERR_PARAM1_INVALID, FAIL_OR_OK);

	// The no-wildcard case could be handled via FilePatternApply(), but handling it this
	// way ensures deleting a non-existent path without wildcards is considered a failure:
	if (!StrChrAny(aFilePattern, _T("?*"))) // No wildcards; just a plain path/filename.
	{
		SetLastError(0); // For sanity: DeleteFile appears to set it only on failure.
		return SetLastErrorMaybeThrow(!DeleteFile(aFilePattern));
	}

	// Otherwise aFilePattern contains wildcards, so we'll search for all matches and delete them.
	return FilePatternApply(aFilePattern, FILE_LOOP_FILES_ONLY, false, FileDeleteCallback, NULL);
}



ResultType Line::FileInstall(LPTSTR aSource, LPTSTR aDest, LPTSTR aFlag)
{
	bool success;
	bool allow_overwrite = (ATOI(aFlag) == 1);
#ifndef AUTOHOTKEYSC
	if (g_script.mKind != Script::ScriptKindResource)
		success = FileInstallCopy(aSource, aDest, allow_overwrite);
	else
#endif
		success = FileInstallExtract(aSource, aDest, allow_overwrite);
	return ThrowIfTrue(!success);
}

bool Line::FileInstallExtract(LPTSTR aSource, LPTSTR aDest, bool aOverwrite)
{
	// Open the file first since it's the most likely to fail:
	HANDLE hfile = CreateFile(aDest, GENERIC_WRITE, 0, NULL, aOverwrite ? CREATE_ALWAYS : CREATE_NEW, 0, NULL);
	if (hfile == INVALID_HANDLE_VALUE)
		return false;

	// Create a temporary copy of aSource to ensure it is the correct case (upper-case).
	// Ahk2Exe converts it to upper-case before adding the resource. My testing showed that
	// using lower or mixed case in some instances prevented the resource from being found.
	// Since file paths are case-insensitive, it certainly doesn't seem harmful to do this:
	TCHAR source[T_MAX_PATH];
	size_t source_length = _tcslen(aSource);
	if (source_length >= _countof(source))
		// Probably can't happen; for simplicity, truncate it.
		source_length = _countof(source) - 1;
	tmemcpy(source, aSource, source_length + 1);
	_tcsupr(source);

	// Find and load the resource.
	HRSRC res;
	HGLOBAL res_load;
	LPVOID res_lock;
	bool success = false;
	if ( (res = FindResource(NULL, source, RT_RCDATA))
	  && (res_load = LoadResource(NULL, res))
	  && (res_lock = LockResource(res_load))  )
	{
		DWORD num_bytes_written;
		// Write the resource data to file.
		success = WriteFile(hfile, res_lock, SizeofResource(NULL, res), &num_bytes_written, NULL);
	}
	CloseHandle(hfile);
	return success;
}

#ifndef AUTOHOTKEYSC
bool Line::FileInstallCopy(LPTSTR aSource, LPTSTR aDest, bool aOverwrite)
{
	// v1.0.35.11: Must search in A_ScriptDir by default because that's where ahk2exe will search by default.
	// The old behavior was to search in A_WorkingDir, which seems pointless because ahk2exe would never
	// be able to use that value if the script changes it while running.
	TCHAR source_path[T_MAX_PATH], dest_path[T_MAX_PATH];
	GetFullPathName(aDest, _countof(dest_path), dest_path, NULL);
	// Avoid attempting the copy if both paths are the same (since it would fail with ERROR_SHARING_VIOLATION),
	// but resolve both to full paths in case mFileDir != g_WorkingDir.  There is a more thorough way to detect
	// when two *different* paths refer to the same file, but it doesn't work with different network shares, and
	// the additional complexity wouldn't be warranted.  Also, the limitations of this method are clearer.
	SetCurrentDirectory(g_script.mFileDir);
	GetFullPathName(aSource, _countof(source_path), source_path, NULL);
	SetCurrentDirectory(g_WorkingDir); // Restore to proper value.
	if (!lstrcmpi(source_path, dest_path) // Full paths are equal.
		&& !(GetFileAttributes(source_path) & FILE_ATTRIBUTE_DIRECTORY)) // Source file exists and is not a directory (otherwise, an error should be thrown).
		return true;

	return CopyFile(source_path, dest_path, !aOverwrite);
}
#endif



ResultType Line::FileCopyOrMove(LPTSTR aSource, LPTSTR aDest, bool aOverwrite)
{
	if (!*aDest) // Fix for v1.1.34.03: Previous behaviour was a Critical Error.
		return LineError(ERR_PARAM2_MUST_NOT_BE_BLANK);
	int error_count = 0;
	if (*aSource) // For backward-compatibility, empty Source is treated as "no files found".
		error_count = Util_CopyFile(aSource, aDest, aOverwrite, mActionType == ACT_FILEMOVE, g->LastError);
	return ThrowIntIfNonzero(error_count);
}



BIF_DECL(BIF_FileGetAttrib)
{
	_f_param_string_opt_def(aFilespec, 0, (g->mLoopFile ? g->mLoopFile->cFileName : _T("")));

	if (!*aFilespec)
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK);

	DWORD attr = GetFileAttributes(aFilespec);
	if (attr == 0xFFFFFFFF)  // Failure, probably because file doesn't exist.
	{
		aResultToken.SetLastErrorMaybeThrow(true);
		return;
	}

	g->LastError = 0;
	_f_return_p(FileAttribToStr(_f_retval_buf, attr));
}



BOOL FileSetAttribCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetAttribData
{
	DWORD and_mask, xor_mask;
};

ResultType Line::FileSetAttrib(LPTSTR aAttributes, LPTSTR aFilePattern
	, FileLoopModeType aOperateOnFolders, bool aDoRecurse)
// Returns the number of files and folders that could not be changed due to an error.
{
	if (!*aFilePattern)
		return LineError(ERR_PARAM2_INVALID, FAIL_OR_OK);

	// Convert the attribute string to three bit-masks: add, remove and toggle.
	FileSetAttribData attrib;
	DWORD mask;
	int op = 0;
	attrib.and_mask = 0xFFFFFFFF; // Set default: keep all bits.
	attrib.xor_mask = 0; // Set default: affect none.
	for (LPTSTR cp = aAttributes; *cp; ++cp)
	{
		switch (ctoupper(*cp))
		{
		case '+':
		case '-':
		case '^':
			op = *cp;
		case ' ':
		case '\t':
			continue;
		default:
			return LineError(ERR_PARAM1_INVALID, FAIL_OR_OK, cp);
		// Note that D (directory) and C (compressed) are currently not supported:
		case 'R': mask = FILE_ATTRIBUTE_READONLY; break;
		case 'A': mask = FILE_ATTRIBUTE_ARCHIVE; break;
		case 'S': mask = FILE_ATTRIBUTE_SYSTEM; break;
		case 'H': mask = FILE_ATTRIBUTE_HIDDEN; break;
		// N: Docs say it's valid only when used alone.  But let the API handle it if this is not so.
		case 'N': mask = FILE_ATTRIBUTE_NORMAL; break;
		case 'O': mask = FILE_ATTRIBUTE_OFFLINE; break;
		case 'T': mask = FILE_ATTRIBUTE_TEMPORARY; break;
		}
		switch (op)
		{
		case '+':
			attrib.and_mask &= ~mask; // Reset bit to 0.
			attrib.xor_mask |= mask; // Set bit to 1.
			break;
		case '-':
			attrib.and_mask &= ~mask; // Reset bit to 0.
			attrib.xor_mask &= ~mask; // Override any prior + or ^.
			break;
		case '^':
			attrib.xor_mask ^= mask; // Toggle bit.  ^= vs |= to invert any prior + or ^.
			// Leave and_mask as is, so any prior + or - will be inverted.
			break;
		default: // No +/-/^ specified, so overwrite attributes (equal and opposite to FileGetAttrib).
			attrib.and_mask = 0; // Reset all bits to 0.
			attrib.xor_mask |= mask; // Set bit to 1.  |= to accumulate if multiple attributes are present.
			break;
		}
	}
	FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetAttribCallback, &attrib);
	return OK;
}

BOOL FileSetAttribCallback(LPTSTR file_path, WIN32_FIND_DATA &current_file, void *aCallbackData)
{
	FileSetAttribData &attrib = *(FileSetAttribData *)aCallbackData;
	DWORD file_attrib = ((current_file.dwFileAttributes & attrib.and_mask) ^ attrib.xor_mask);
	if (!SetFileAttributes(file_path, file_attrib))
	{
		g->LastError = GetLastError();
		return FALSE;
	}
	return TRUE;
}



ResultType Line::FilePatternApply(LPTSTR aFilePattern, FileLoopModeType aOperateOnFolders
	, bool aDoRecurse, FilePatternCallback aCallback, void *aCallbackData)
{
	if (!*aFilePattern)
		// Caller should handle this case before calling us if an exception is to be thrown.
		return SetLastErrorMaybeThrow(true, ERROR_INVALID_PARAMETER);

	if (aOperateOnFolders == FILE_LOOP_INVALID) // In case runtime dereference of a var was an invalid value.
		aOperateOnFolders = FILE_LOOP_FILES_ONLY;  // Set default.
	g->LastError = 0; // Set default. Overridden only when a failure occurs.

	FilePatternStruct fps;

	LPTSTR last_backslash = _tcsrchr(aFilePattern, '\\');
	if (last_backslash)
		fps.dir_length = last_backslash - aFilePattern + 1; // Include the slash.
	else // Use current working directory, e.g. if user specified only *.*
		fps.dir_length = 0;
	fps.pattern_length = _tcslen(aFilePattern + fps.dir_length);
	
	// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
	// than 259, even if the pattern would match files whose names are short enough to be legal.
	if (fps.dir_length + fps.pattern_length >= _countof(fps.path)
		|| fps.pattern_length >= _countof(fps.pattern))
		return SetLastErrorMaybeThrow(true, ERROR_BUFFER_OVERFLOW);

	// Make copies in case of overwrite of deref buf during LONG_OPERATION/MsgSleep,
	// and to allow modification:
	_tcscpy(fps.path, aFilePattern); // Include the pattern initially.
	_tcscpy(fps.pattern, aFilePattern + fps.dir_length); // Just the naked filename or pattern, for use with aDoRecurse.

	if (!StrChrAny(fps.pattern, _T("?*")))
		// Since no wildcards, always operate on this single item even if it's a folder.
		aOperateOnFolders = FILE_LOOP_FILES_AND_FOLDERS;

	// Passing the parameters this way reduces code size:
	fps.aCallback = aCallback;
	fps.aCallbackData = aCallbackData;
	fps.aDoRecurse = aDoRecurse;
	fps.aOperateOnFolders = aOperateOnFolders;

	fps.failure_count = 0;

	FilePatternApply(fps);
	return ThrowIntIfNonzero(fps.failure_count); // i.e. indicate success if there were no failures.
}



void Line::FilePatternApply(FilePatternStruct &fps)
{
	size_t dir_length = fps.dir_length; // Length of this directory (saved before recursion).
	LPTSTR append_pos = fps.path + dir_length; // This is where the changing part gets appended.
	size_t space_remaining = _countof(fps.path) - dir_length - 1; // Space left in file_path for the changing part.

	LONG_OPERATION_INIT
	int failure_count = 0;
	WIN32_FIND_DATA current_file;
	HANDLE file_search = FindFirstFile(fps.path, &current_file);

	if (file_search != INVALID_HANDLE_VALUE)
	{
		do
		{
			// Since other script threads can interrupt during LONG_OPERATION_UPDATE, it's important that
			// this command not refer to sArgDeref[] and sArgVar[] anytime after an interruption becomes
			// possible. This is because an interrupting thread usually changes the values to something
			// inappropriate for this thread.
			LONG_OPERATION_UPDATE

			if (current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
			{
				if (current_file.cFileName[0] == '.' && (!current_file.cFileName[1]    // Relies on short-circuit boolean order.
					|| current_file.cFileName[1] == '.' && !current_file.cFileName[2]) //
					// Regardless of whether this folder will be recursed into, this folder
					// will not be affected when the mode is files-only:
					|| fps.aOperateOnFolders == FILE_LOOP_FILES_ONLY)
					continue; // Never operate upon or recurse into these.
			}
			else // It's a file, not a folder.
				if (fps.aOperateOnFolders == FILE_LOOP_FOLDERS_ONLY)
					continue;

			if (_tcslen(current_file.cFileName) > space_remaining)
			{
				// v1.0.45.03: Don't even try to operate upon truncated filenames in case they accidentally
				// match the name of a real/existing file.
				g->LastError = ERROR_BUFFER_OVERFLOW;
				++failure_count;
				continue;
			}
			// Otherwise, make file_path be the filespec of the file to operate upon:
			_tcscpy(append_pos, current_file.cFileName); // Above has ensured this won't overflow.
			//
			// This is the part that actually does something to the file:
			if (!fps.aCallback(fps.path, current_file, fps.aCallbackData))
				++failure_count;
			//
		} while (FindNextFile(file_search, &current_file));

		FindClose(file_search);
	} // if (file_search != INVALID_HANDLE_VALUE)

	if (fps.aDoRecurse && space_remaining > 1) // The space_remaining check ensures there's enough room to append "*", though if false, that would imply lfs.pattern is empty.
	{
		_tcscpy(append_pos, _T("*")); // Above has ensured this won't overflow.
		file_search = FindFirstFile(fps.path, &current_file);

		if (file_search != INVALID_HANDLE_VALUE)
		{
			do
			{
				LONG_OPERATION_UPDATE
				if (!(current_file.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
					|| current_file.cFileName[0] == '.' && (!current_file.cFileName[1]      // Relies on short-circuit boolean order.
						|| current_file.cFileName[1] == '.' && !current_file.cFileName[2])) //
					continue;
				size_t filename_length = _tcslen(current_file.cFileName);
				// v1.0.45.03: Skip over folders whose paths are too long to be supported by FindFirst.
				if (fps.pattern_length + filename_length >= space_remaining) // >= vs. > to reserve 1 for the backslash to be added between cFileName and naked_filename_or_pattern.
					continue; // Never recurse into these.
				// This will build the string CurrentDir+SubDir+FilePatternOrName.
				// If FilePatternOrName doesn't contain a wildcard, the recursion
				// process will attempt to operate on the originally-specified
				// single filename or folder name if it occurs anywhere else in the
				// tree, e.g. recursing C:\Temp\temp.txt would affect all occurrences
				// of temp.txt both in C:\Temp and any subdirectories it might contain:
				_stprintf(append_pos, _T("%s\\%s") // Above has ensured this won't overflow.
					, current_file.cFileName, fps.pattern);
				fps.dir_length = dir_length + filename_length + 1; // Include the slash.
				//
				// Apply the callback to files in this subdirectory:
				FilePatternApply(fps);
				//
			} while (FindNextFile(file_search, &current_file));
			FindClose(file_search);
		} // if (file_search != INVALID_HANDLE_VALUE)
	} // if (aDoRecurse)

	fps.failure_count += failure_count; // Update failure count (produces smaller code than doing ++fps.failure_count directly).
}



BIF_DECL(BIF_FileGetTime)
{
	_f_param_string_opt_def(aFilespec, 0, (g->mLoopFile ? g->mLoopFile->cFileName : _T("")));
	_f_param_string_opt(aWhichTime, 1);

	if (!*aFilespec)
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK);

	// Don't use CreateFile() & FileGetSize() size they will fail to work on a file that's in use.
	// Research indicates that this method has no disadvantages compared to the other method.
	WIN32_FIND_DATA found_file;
	HANDLE file_search = FindFirstFile(aFilespec, &found_file);
	if (file_search == INVALID_HANDLE_VALUE)
	{
		aResultToken.SetLastErrorMaybeThrow(true);
		return;
	}
	FindClose(file_search);

	FILETIME local_file_time;
	switch (ctoupper(*aWhichTime))
	{
	case 'C': // File's creation time.
		FileTimeToLocalFileTime(&found_file.ftCreationTime, &local_file_time);
		break;
	case 'A': // File's last access time.
		FileTimeToLocalFileTime(&found_file.ftLastAccessTime, &local_file_time);
		break;
	default:  // 'M', unspecified, or some other value.  Use the file's modification time.
		FileTimeToLocalFileTime(&found_file.ftLastWriteTime, &local_file_time);
	}

	g->LastError = 0;
	_f_return_p(FileTimeToYYYYMMDD(_f_retval_buf, local_file_time));
}



BOOL FileSetTimeCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetTimeData
{
	FILETIME Time;
	TCHAR WhichTime;
};

ResultType Line::FileSetTime(LPTSTR aYYYYMMDD, LPTSTR aFilePattern, TCHAR aWhichTime
	, FileLoopModeType aOperateOnFolders, bool aDoRecurse)
// Returns the number of files and folders that could not be changed due to an error.
{
	// Related to the comment at the top: Since the script subroutine that resulted in the call to
	// this function can be interrupted during our MsgSleep(), make a copy of any params that might
	// currently point directly to the deref buffer.  This is done because their contents might
	// be overwritten by the interrupting subroutine:
	TCHAR yyyymmdd[64]; // Even do this one since its value is passed recursively in calls to self.
	tcslcpy(yyyymmdd, aYYYYMMDD, _countof(yyyymmdd));

	FileSetTimeData callbackData;
	callbackData.WhichTime = aWhichTime;
	FILETIME ft;
	if (*yyyymmdd)
	{
		if (   !YYYYMMDDToFileTime(yyyymmdd, ft)  // Convert the arg into the time struct as local (non-UTC) time.
			|| !LocalFileTimeToFileTime(&ft, &callbackData.Time)   )  // Convert from local to UTC.
		{
			// Invalid parameters are the only likely cause of this condition.
			return LineError(ERR_PARAM1_INVALID, FAIL_OR_OK, aYYYYMMDD);
		}
	}
	else // User wants to use the current time (i.e. now) as the new timestamp.
		GetSystemTimeAsFileTime(&callbackData.Time);

	if (!*aFilePattern)
		return LineError(ERR_PARAM2_INVALID, FAIL_OR_OK);

	FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetTimeCallback, &callbackData);
	return OK;
}

BOOL FileSetTimeCallback(LPTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData)
{
	HANDLE hFile;
	// Open existing file.
	// FILE_FLAG_NO_BUFFERING might improve performance because all we're doing is
	// changing one of the file's attributes.  FILE_FLAG_BACKUP_SEMANTICS must be
	// used, otherwise changing the time of a directory under NT and beyond will
	// not succeed.  Win95 (not sure about Win98/Me) does not support this, but it
	// should be harmless to specify it even if the OS is Win95:
	hFile = CreateFile(aFilename, GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE
		, (LPSECURITY_ATTRIBUTES)NULL, OPEN_EXISTING
		, FILE_FLAG_NO_BUFFERING | FILE_FLAG_BACKUP_SEMANTICS, NULL);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		g->LastError = GetLastError();
		return FALSE;
	}

	BOOL success;
	FileSetTimeData &a = *(FileSetTimeData *)aCallbackData;
	switch (ctoupper(a.WhichTime))
	{
	case 'C': // File's creation time.
		success = SetFileTime(hFile, &a.Time, NULL, NULL);
		break;
	case 'A': // File's last access time.
		success = SetFileTime(hFile, NULL, &a.Time, NULL);
		break;
	default:  // 'M', unspecified, or some other value.  Use the file's modification time.
		success = SetFileTime(hFile, NULL, NULL, &a.Time);
	}
	if (!success)
		g->LastError = GetLastError();
	CloseHandle(hFile);
	return success;
}



BIF_DECL(BIF_FileGetSize)
{
	_f_param_string_opt_def(aFilespec, 0, (g->mLoopFile ? g->mLoopFile->cFileName : _T("")));
	_f_param_string_opt(aGranularity, 1);

	if (!*aFilespec)
		_f_throw_value(ERR_PARAM2_MUST_NOT_BE_BLANK); // Throw an error, since this is probably not what the user intended.
	
	BOOL got_file_size = false;
	__int64 size;

	// Try CreateFile() and GetFileSizeEx() first, since they can be more accurate. 
	// See "Why is the file size reported incorrectly for files that are still being written to?"
	// http://blogs.msdn.com/b/oldnewthing/archive/2011/12/26/10251026.aspx
	HANDLE hfile = CreateFile(aFilespec, FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
		, NULL, OPEN_EXISTING, 0, NULL);
	if (hfile != INVALID_HANDLE_VALUE)
	{
		got_file_size = GetFileSizeEx(hfile, (PLARGE_INTEGER)&size);
		CloseHandle(hfile);
	}

	if (!got_file_size)
	{
		WIN32_FIND_DATA found_file;
		HANDLE file_search = FindFirstFile(aFilespec, &found_file);
		if (file_search == INVALID_HANDLE_VALUE)
		{
			aResultToken.SetLastErrorMaybeThrow(true);
			_f_return_empty;
		}
		FindClose(file_search);
		size = ((__int64)found_file.nFileSizeHigh << 32) | found_file.nFileSizeLow;
	}

	switch(ctoupper(*aGranularity))
	{
	case 'K': // KB
		size /= 1024;
		break;
	case 'M': // MB
		size /= (1024 * 1024);
		break;
	// default: // i.e. either 'B' for bytes, or blank, or some other unknown value, so default to bytes.
		// do nothing
	}

	g->LastError = 0;
	_f_return(size);
	// The below comment is obsolete in light of the switch to 64-bit integers.  But it might
	// be good to keep for background:
	// Currently, the above is basically subject to a 2 gig limit, I believe, after which the
	// size will appear to be negative.  Beyond a 4 gig limit, the value will probably wrap around
	// to zero and start counting from there as file sizes grow beyond 4 gig (UPDATE: The size
	// is now set to -1 [the maximum DWORD when expressed as a signed int] whenever >4 gig).
	// There's not much sense in putting values larger than 2 gig into the var as a text string
	// containing a positive number because such a var could never be properly handled by anything
	// that compares to it (e.g. IfGreater) or does math on it (e.g. EnvAdd), since those operations
	// use ATOI() to convert the string.  So as a future enhancement (unless the whole program is
	// revamped to use 64bit ints or something) might add an optional param to the end to indicate
	// size should be returned in K(ilobyte) or M(egabyte).  However, this is sorta bad too since
	// adding a param can break existing scripts which use filenames containing commas (delimiters)
	// with this command.  Therefore, I think I'll just add the K/M param now.
	// Also, the above assigns an int because unsigned ints should never be stored in script
	// variables.  This is because an unsigned variable larger than INT_MAX would not be properly
	// converted by ATOI(), which is current standard method for variables to be auto-converted
	// from text back to a number whenever that is needed.
}



/////////////////////
// SetXXXLockState //
/////////////////////

ResultType Line::SetToggleState(vk_type aVK, ToggleValueType &ForceLock, LPTSTR aToggleText)
// Caller must have already validated that the args are correct.
// Always returns OK, for use as caller's return value.
{
	ToggleValueType toggle = ConvertOnOffAlways(aToggleText, NEUTRAL);
	switch (toggle)
	{
	case TOGGLED_ON:
	case TOGGLED_OFF:
		// Turning it on or off overrides any prior AlwaysOn or AlwaysOff setting.
		// Probably need to change the setting BEFORE attempting to toggle the
		// key state, otherwise the hook may prevent the state from being changed
		// if it was set to be AlwaysOn or AlwaysOff:
		ForceLock = NEUTRAL;
		ToggleKeyState(aVK, toggle);
		break;
	case ALWAYS_ON:
	case ALWAYS_OFF:
		ForceLock = (toggle == ALWAYS_ON) ? TOGGLED_ON : TOGGLED_OFF; // Must do this first.
		ToggleKeyState(aVK, ForceLock);
		// The hook is currently needed to support keeping these keys AlwaysOn or AlwaysOff, though
		// there may be better ways to do it (such as registering them as a hotkey, but
		// that may introduce quite a bit of complexity):
		Hotkey::InstallKeybdHook();
		break;
	case NEUTRAL:
		// Note: No attempt is made to detect whether the keybd hook should be deinstalled
		// because it's no longer needed due to this change.  That would require some 
		// careful thought about the impact on the status variables in the Hotkey class, etc.,
		// so it can be left for a future enhancement:
		ForceLock = NEUTRAL;
		break;
	}
	return OK;
}



////////////////////////////////
// Misc lower level functions //
////////////////////////////////


ResultType GetObjectIntProperty(IObject *aObject, LPTSTR aPropName, __int64 &aValue, ResultToken &aResultToken, bool aOptional)
{
	FuncResult result_token;
	ExprTokenType this_token = aObject;

	auto result = aObject->Invoke(result_token, IT_GET, aPropName, this_token, nullptr, 0);

	if (result_token.symbol != SYM_INTEGER)
	{
		result_token.Free();
		if (result == FAIL || result == EARLY_EXIT)
		{
			aResultToken.SetExitResult(result);
			return FAIL;
		}
		if (result != INVOKE_NOT_HANDLED) // Property exists but is not an integer.
			return aResultToken.Error(ERR_TYPE_MISMATCH, aPropName, ErrorPrototype::Type);
		//aValue = 0; // Caller should set default value for these cases.
		if (!aOptional)
			return aResultToken.UnknownMemberError(ExprTokenType(aObject), IT_GET, aPropName);
		return result; // Let caller know it wasn't found.
	}

	aValue = result_token.value_int64;
	return OK;
}

ResultType SetObjectIntProperty(IObject *aObject, LPTSTR aPropName, __int64 aValue, ResultToken &aResultToken)
{
	FuncResult result_token;
	ExprTokenType this_token = aObject, value_token = aValue, *param = &value_token;

	auto result = aObject->Invoke(result_token, IT_SET, aPropName, this_token, &param, 1);

	result_token.Free();
	if (result == FAIL || result == EARLY_EXIT)
		return aResultToken.SetExitResult(result);
	if (result == INVOKE_NOT_HANDLED)
		return aResultToken.UnknownMemberError(ExprTokenType(aObject), IT_GET, aPropName);
	return OK;
}

ResultType GetObjectPtrProperty(IObject *aObject, LPTSTR aPropName, UINT_PTR &aPtr, ResultToken &aResultToken, bool aOptional)
{
	__int64 value = NULL;
	auto result = GetObjectIntProperty(aObject, aPropName, value, aResultToken, aOptional);
	aPtr = (UINT_PTR)value;
	return result;
}



bool Line::FileIsFilteredOut(LoopFilesStruct &aCurrentFile, FileLoopModeType aFileLoopMode)
{
	if (aCurrentFile.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY) // It's a folder.
	{
		if (aFileLoopMode == FILE_LOOP_FILES_ONLY
			|| aCurrentFile.cFileName[0] == '.' && (!aCurrentFile.cFileName[1]      // Relies on short-circuit boolean order.
				|| aCurrentFile.cFileName[1] == '.' && !aCurrentFile.cFileName[2])) //
			return true; // Exclude this folder by returning true.
	}
	else // it's not a folder.
		if (aFileLoopMode == FILE_LOOP_FOLDERS_ONLY)
			return true; // Exclude this file by returning true.

	// Since file was found, also append the file's name to its directory for the caller:
	// Seems best to check length in advance because it allows a faster move/copy method further below
	// (in lieu of sntprintf(), which is probably quite a bit slower than the method here).
	size_t name_length = _tcslen(aCurrentFile.cFileName);
	if (aCurrentFile.dir_length + name_length >= _countof(aCurrentFile.file_path)) // Should be impossible with current buffer sizes.
		return true; // Exclude this file/folder.
	tmemcpy(aCurrentFile.file_path + aCurrentFile.dir_length, aCurrentFile.cFileName, name_length + 1); // +1 to include the terminator.
	aCurrentFile.file_path_length = aCurrentFile.dir_length + name_length;
	return false; // Indicate that this file is not to be filtered out.
}



Label *Line::GetJumpTarget(bool aIsDereferenced)
{
	LPTSTR target_label = aIsDereferenced ? ARG1 : RAW_ARG1;
	Label *label = g_script.FindLabel(target_label);
	if (!label)
	{
		LineError(ERR_NO_LABEL, FAIL, target_label);
		return NULL;
	}
	// If g->CurrentFunc, label is never outside the function since it would not
	// have been found by FindLabel().  So there's no need to check for that here.
	if (!aIsDereferenced)
		mRelatedLine = (Line *)label; // The script loader has ensured that label->mJumpToLine isn't NULL.
	// else don't update it, because that would permanently resolve the jump target, and we want it to stay dynamic.
	return IsJumpValid(*label);
	// Any error msg was already displayed by the above call.
}



Label *Line::IsJumpValid(Label &aTargetLabel, bool aSilent)
// Returns aTargetLabel is the jump is valid, or NULL otherwise.
{
	// aTargetLabel can be NULL if this Goto's target is the physical end of the script.
	// And such a destination is always valid, regardless of where aOrigin is.
	// UPDATE: It's no longer possible for the destination of a Goto to be
	// NULL because the script loader has ensured that the end of the script always has
	// an extra ACT_EXIT that serves as an anchor for any final labels in the script:
	//if (aTargetLabel == NULL)
	//	return OK;
	// The above check is also necessary to avoid dereferencing a NULL pointer below.

	if (!CheckValidFinallyJump(aTargetLabel.mJumpToLine, aSilent))
		return NULL;

	Line *parent_line_of_label_line;
	if (   !(parent_line_of_label_line = aTargetLabel.mJumpToLine->mParentLine)   )
		// A Goto can always jump to a point anywhere in the outermost layer
		// (i.e. outside all blocks) without restriction (except from inside a
		// function to outside, but in that case the label would not be found):
		return &aTargetLabel; // Indicate success.

	// So now we know this Goto is attempting to jump into a block somewhere.  Is that
	// block a legal place to jump?:

	for (Line *ancestor = mParentLine; ancestor != NULL; ancestor = ancestor->mParentLine)
		if (parent_line_of_label_line == ancestor)
			// Since aTargetLabel is in the same block as the Goto line itself (or a block
			// that encloses that block), it's allowed:
			return &aTargetLabel; // Indicate success.
	// This can happen if the Goto's target is at a deeper level than it, or if the target
	// is at a more shallow level but is in some block totally unrelated to it!
	// Returns FAIL by default, which is what we want because that value is zero:
	if (!aSilent)
		LineError(_T("A Goto must not jump into a block that doesn't enclose it."));
	return NULL;
}


BOOL Line::CheckValidFinallyJump(Line* jumpTarget, bool aSilent)
{
	Line* jumpParent = jumpTarget->mParentLine;
	for (Line *ancestor = mParentLine; ancestor != NULL; ancestor = ancestor->mParentLine)
	{
		if (ancestor == jumpParent)
			return TRUE; // We found the common ancestor.
		if (ancestor->mActionType == ACT_FINALLY)
		{
			if (!aSilent)
				LineError(ERR_BAD_JUMP_INSIDE_FINALLY);
			return FALSE; // The common ancestor is outside the FINALLY block and thus this jump is invalid.
		}
	}
	return TRUE; // The common ancestor is the root of the script.
}


////////////////////////
// BUILT-IN VARIABLES //
////////////////////////


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
	if (!Script::SetCoordMode(aVarName + 11, BivRValueToString())) // A_CoordMode is 11 chars.
		_f_return_FAIL;
}

BIV_DECL_R(BIV_SendMode)
{
	static LPTSTR sSendModes[] = SEND_MODES;
	_f_return_p(sSendModes[g->SendMode]);
}

BIV_DECL_W(BIV_SendMode_Set)
{
	if (!Script::SetSendMode(BivRValueToString()))
		_f_return_FAIL;
}

BIV_DECL_R(BIV_SendLevel)
{
	_f_return_i(g->SendLevel);
}

BIV_DECL_W(BIV_SendLevel_Set)
{
	if (!Script::SetSendLevel((int)BivRValueToInt64(), BivRValueToString()))
		_f_return_FAIL;
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
	// It seems more useful to return g->PeekFrequency than "On" or "Off" (ACT_CRITICAL ensures that
	// g->PeekFrequency!=0 whenever g->ThreadIsCritical==true).  Also, the word "Is" in "A_IsCritical"
	// implies a value that can be used as a boolean such as "if A_IsCritical".
	if (g->ThreadIsCritical)
		_f_return_i(g->PeekFrequency); // ACT_CRITICAL ensures that g->PeekFrequency > 0 when critical is on.
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
	if (!SetWorkingDir(BivRValueToString()))
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
		MsgBox(_T("DEBUG: Unhandled SpecialFolderPath variable."));
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



BIF_DECL(BIF_Persistent)
{
	// Need to set a return value explicitly, since the default is effectively StrPtr(""), not "".
	// Returning the old value might have some use, but if the caller doesn't want its value to change,
	// something awkward like "Persistent(isPersistent := Persistent())" is needed.  Rather than just
	// returning the current status, Persistent() makes the script persistent because that's likely to
	// be its most common use by far, and it's what users familiar with the old #Persistent may expect.
	_f_set_retval_i(g_persistent);
	g_persistent = ParamIndexToOptionalBOOL(0, true);
}



BIF_DECL(BIF_InstallHook)
{
	bool installing = ParamIndexToOptionalBOOL(0, true);
	bool use_force = ParamIndexToOptionalBOOL(1, false);
	auto which_hook = (HookType)_f_callee_id;
	// When the second parameter is true, unconditionally remove the hook.  If the first parameter is
	// also true, the hook will be reinstalled fresh.  Otherwise the hook will be left uninstalled,
	// until something happens to reinstall it, such as Hotkey::ManifestAllHotkeysHotstringsHooks().
	if (use_force)
		AddRemoveHooks(GetActiveHooks() & ~which_hook);
	Hotkey::RequireHook(which_hook, installing);
	if (!use_force || installing)
		Hotkey::ManifestAllHotkeysHotstringsHooks();
}



////////////////////
// Misc Functions //
////////////////////


void ObjectToString(ResultToken &aResultToken, ExprTokenType &aThisToken, IObject *aObject)
{
	// Something like this should be done for every TokenToString() call or
	// equivalent, but major changes are needed before that will be feasible.
	// For now, String(anytype) provides a limited workaround.
	switch (aObject->Invoke(aResultToken, IT_CALL, _T("ToString"), aThisToken, nullptr, 0))
	{
	case INVOKE_NOT_HANDLED:
		aResultToken.UnknownMemberError(aThisToken, IT_CALL, _T("ToString"));
		break;
	case FAIL:
		aResultToken.SetExitResult(FAIL);
		break;
	}
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



///////////////////////
// Interop Functions //
///////////////////////


struct NumGetParams
{
	size_t target, right_side_bound;
	size_t num_size = sizeof(DWORD_PTR);
	BOOL is_integer = TRUE, is_signed = FALSE;
};

void ConvertNumGetType(ExprTokenType &aToken, NumGetParams &op)
{
	LPTSTR type = TokenToString(aToken); // No need to pass aBuf since any numeric value would not be recognized anyway.
	if (ctoupper(*type) == 'U') // Unsigned.
	{
		++type; // Remove the first character from further consideration.
		op.is_signed = FALSE;
	}
	else
		op.is_signed = TRUE;

	switch(ctoupper(*type)) // Override "size" and aResultToken.symbol if type warrants it. Note that the above has omitted the leading "U", if present, leaving type as "Int" vs. "Uint", etc.
	{
	case 'P': // Nothing extra needed in this case.
		op.num_size = sizeof(void *), op.is_integer = TRUE;
		break;
	case 'I':
		if (_tcschr(type, '6')) // Int64. It's checked this way for performance, and to avoid access violation if string is bogus and too short such as "i64".
			op.num_size = 8, op.is_integer = TRUE;
		else
			op.num_size = 4, op.is_integer = TRUE;
		break;
	case 'S': op.num_size = 2, op.is_integer = TRUE; break; // Short.
	case 'C': op.num_size = 1, op.is_integer = TRUE; break; // Char.

	case 'D': op.num_size = 8, op.is_integer = FALSE; break; // Double.
	case 'F': op.num_size = 4, op.is_integer = FALSE; break; // Float.

	default: op.num_size = 0; break;
	}
}

void *BufferObject::sVTable = getVTable(); // Placing this here vs. in script_object.cpp improves some simple benchmarks by as much as 7%.

void GetBufferObjectPtr(ResultToken &aResultToken, IObject *obj, size_t &aPtr, size_t &aSize)
{
	if (BufferObject::IsInstanceExact(obj))
	{
		// Some primitive benchmarks showed that this was about as fast as passing
		// a pointer directly, whereas invoking the properties (below) doubled the
		// overall time taken by NumGet/NumPut.
		aPtr = (size_t)((BufferObject *)obj)->Data();
		aSize = ((BufferObject *)obj)->Size();
	}
	else
	{
		if (GetObjectPtrProperty(obj, _T("Ptr"), aPtr, aResultToken))
			GetObjectPtrProperty(obj, _T("Size"), aSize, aResultToken);
	}
}

void GetBufferObjectPtr(ResultToken &aResultToken, IObject *obj, size_t &aPtr)
// See above for comments.
{
	if (BufferObject::IsInstanceExact(obj))
		aPtr = (size_t)((BufferObject *)obj)->Data();
	else
		GetObjectPtrProperty(obj, _T("Ptr"), aPtr, aResultToken);
}

void ConvertNumGetTarget(ResultToken &aResultToken, ExprTokenType &target_token, NumGetParams &op)
{
	if (IObject *obj = TokenToObject(target_token))
	{
		GetBufferObjectPtr(aResultToken, obj, op.target, op.right_side_bound);
		if (aResultToken.Exited())
			return;
		op.right_side_bound += op.target;
	}
	else
	{
		op.target = (size_t)TokenToInt64(target_token);
		op.right_side_bound = SIZE_MAX;
	}
}


BIF_DECL(BIF_NumGet)
{
	NumGetParams op;
	ConvertNumGetTarget(aResultToken, *aParam[0], op);
	if (aResultToken.Exited())
		return;
	if (aParamCount > 2) // Offset was specified.
	{
		op.target += (ptrdiff_t)TokenToInt64(*aParam[1]);
		aParam++;
	}
	// MinParams ensures there is always one more parameter.
	ConvertNumGetType(*aParam[1], op);

	// If the target is a variable, the following check ensures that the memory to be read lies within its capacity.
	// This seems superior to an exception handler because exception handlers only catch illegal addresses,
	// not ones that are technically legal but unintentionally bugs due to being beyond a variable's capacity.
	// Moreover, __try/except is larger in code size. Another possible alternative is IsBadReadPtr()/IsBadWritePtr(),
	// but those are discouraged by MSDN.
	// The following aren't covered by the check below:
	// - Due to rarity of negative offsets, only the right-side boundary is checked, not the left.
	// - Due to rarity and to simplify things, Float/Double aren't checked.
	if (!op.num_size
		|| op.target < 65536 // Basic sanity check to catch incoming raw addresses that are zero or blank.  On Win32, the first 64KB of address space is always invalid.
		|| op.target+op.num_size > op.right_side_bound) // i.e. it's ok if target+size==right_side_bound because the last byte to be read is actually at target+size-1. In other words, the position of the last possible terminator within the variable's capacity is considered an allowable address.
	{
		_f_throw_value(ERR_PARAM_INVALID);
	}

	switch (op.num_size)
	{
	case 4: // Listed first for performance.
		if (!op.is_integer)
			aResultToken.value_double = *(float *)op.target;
		else if (op.is_signed)
			aResultToken.value_int64 = *(int *)op.target; // aResultToken.symbol defaults to SYM_INTEGER.
		else
			aResultToken.value_int64 = *(unsigned int *)op.target;
		break;
	case 8:
		if (op.is_integer)
			// Unsigned 64-bit integers aren't supported because variables/expressions can't support them.
			aResultToken.value_int64 = *(__int64 *)op.target;
		else
			aResultToken.value_double = *(double *)op.target;
		break;
	case 2:
		if (op.is_signed) // Don't use ternary because that messes up type-casting.
			aResultToken.value_int64 = *(short *)op.target;
		else
			aResultToken.value_int64 = *(unsigned short *)op.target;
		break;
	case 1:
		if (op.is_signed) // Don't use ternary because that messes up type-casting.
			aResultToken.value_int64 = *(char *)op.target;
		else
			aResultToken.value_int64 = *(unsigned char *)op.target;
		break;
	}
	if (!op.is_integer)
		aResultToken.symbol = SYM_FLOAT;
}



///////////////////////
// Interop Functions //
///////////////////////

BIF_DECL(BIF_NumPut)
{
	// Params can be any non-zero number of type-number pairs, followed by target[, offset].
	// Prior validation has ensured that there are at least three parameters.
	//   NumPut(t1, n1, t2, n2, p, o)
	//   NumPut(t1, n1, t2, n2, p)
	//   NumPut(t1, n1, p, o)
	//   NumPut(t1, n1, p)
	
	// Split target[,offset] from aParam.
	bool offset_was_specified = !(aParamCount & 1);
	aParamCount -= 1 + int(offset_was_specified);
	ExprTokenType &target_token = *aParam[aParamCount];
	
	NumGetParams op;
	ConvertNumGetTarget(aResultToken, target_token, op);
	if (aResultToken.Exited())
		return;
	if (offset_was_specified)
		op.target += (ptrdiff_t)TokenToInt64(*aParam[aParamCount + 1]);

	size_t num_end;
	for (int n_param = 1; n_param < aParamCount; n_param += 2, op.target = num_end)
	{
		ConvertNumGetType(*aParam[n_param - 1], op); // Type name.
		ExprTokenType &token_to_write = *aParam[n_param]; // Numeric value.

		num_end = op.target + op.num_size; // This is used below and also as NumPut's return value. It's the address to the right of the item to be written.

		// See comments in NumGet about the following section:
		if (!op.num_size
			|| !TokenIsNumeric(token_to_write)
			|| op.target < 65536 // Basic sanity check to catch incoming raw addresses that are zero or blank.  On Win32, the first 64KB of address space is always invalid.
			|| num_end > op.right_side_bound) // i.e. it's ok if target+size==right_side_bound because the last byte to be read is actually at target+size-1. In other words, the position of the last possible terminator within the variable's capacity is considered an allowable address.
		{
			_f_throw_value(ERR_PARAM_INVALID);
		}

		union
		{
			__int64 num_i64;
			double num_f64;
			float num_f32;
		};

		// Note that since v2.0-a083-97803aeb, TokenToInt64 supports conversion of large unsigned 64-bit
		// numbers from strings (producing a negative value, but with the right bit representation).
		if (op.is_integer)
			num_i64 = TokenToInt64(token_to_write);
		else
		{
			num_f64 = TokenToDouble(token_to_write);
			if (op.num_size == 4)
				num_f32 = (float)num_f64;
		}

		// This method benchmarked marginally faster than memcpy for the multi-param mode.
		switch (op.num_size)
		{
		case 8: *(UINT64 *)op.target = (UINT64)num_i64; break;
		case 4: *(UINT32 *)op.target = (UINT32)num_i64; break;
		case 2: *(UINT16 *)op.target = (UINT16)num_i64; break;
		case 1: *(UINT8 *)op.target = (UINT8)num_i64; break;
		}
	}
	if (target_token.symbol == SYM_VAR && !target_token.var->IsPureNumeric())
		target_token.var->Close(); // This updates various attributes of the variable.
	//else the target was an raw address.  If that address is inside some variable's contents, the above
	// attributes would already have been removed at the time the & operator was used on the variable.
	aResultToken.value_int64 = num_end; // aResultToken.symbol was set to SYM_INTEGER by our caller.
}



BIF_DECL(BIF_StrGetPut) // BIF_DECL(BIF_StrGet), BIF_DECL(BIF_StrPut)
{
	// To simplify flexible handling of parameters:
	ExprTokenType **aParam_end = aParam + aParamCount, **next_param = aParam;

	LPCVOID source_string; // This may hold an intermediate UTF-16 string in ANSI builds.
	int source_length;
	if (_f_callee_id == FID_StrPut)
	{
		// StrPut(String, Address[, Length][, Encoding])
		ExprTokenType &source_token = *aParam[0];
		source_string = (LPCVOID)TokenToString(source_token, _f_number_buf); // Safe to use _f_number_buf since StrPut won't use it for the result.
		source_length = (int)((source_token.symbol == SYM_VAR) ? source_token.var->CharLength() : _tcslen((LPCTSTR)source_string));
		++next_param; // Remove the String param from further consideration.
	}
	else
	{
		// StrGet(Address[, Length][, Encoding])
		source_string = NULL;
		source_length = 0;
	}

	aResultToken.symbol = SYM_STRING;
	aResultToken.marker = _T(""); // Set default in case of early return.

	IObject *buffer_obj;
	LPVOID 	address;
	size_t  max_bytes = SIZE_MAX;
	int 	length = -1; // actual length
	bool	length_is_max_size = false;
	UINT 	encoding = UorA(CP_UTF16, CP_ACP); // native encoding

	// Parameters are interpreted according to the following rules (highest to lowest precedence):
	// Legend:  StrPut(String[, X, Y, Z])  or  StrGet(Address[, Y, Z])
	// - If X is non-numeric, it is Encoding.  Calculates required buffer size but does nothing else.  Y and Z must be omitted.
	// - If X is numeric, it is Address.  (For StrGet, non-numeric Address is treated as an error.)
	// - If Y is numeric, it is Length.  Otherwise "Actual length" is assumed.
	// - If a parameter remains, it is Encoding.
	// Encoding may therefore only be purely numeric if Address(X) and Length(Y) are specified.

	const LPVOID FIRST_VALID_ADDRESS = (LPVOID)65536;

	if (next_param < aParam_end && TokenIsNumeric(**next_param))
	{
		address = (LPVOID)TokenToInt64(**next_param);
		++next_param;
	}
	else if (next_param < aParam_end && (buffer_obj = TokenToObject(**next_param)))
	{
		size_t ptr;
		GetBufferObjectPtr(aResultToken, buffer_obj, ptr, max_bytes);
		if (aResultToken.Exited())
			return;
		address = (LPVOID)ptr;
		++next_param;
	}
	else
	{
		if (!source_string || aParamCount > 2)
		{
			// See the "Legend" above.  Either this is StrGet and Address was invalid (it can't be omitted due
			// to prior min-param checks), or it is StrPut and there are too many parameters.
			_f_throw_value(source_string ? ERR_PARAM_INVALID : ERR_PARAM1_INVALID);  // StrPut : StrGet
		}
		// else this is the special measuring mode of StrPut, where Address and Length are omitted.
		// A length of 0 when passed to the Win API conversion functions (or the code below) means
		// "calculate the required buffer size, but don't do anything else."
		length = 0;
		address = FIRST_VALID_ADDRESS; // Skip validation below; address should never be used when length == 0.
	}

	if (next_param < aParam_end)
	{
		if (length == -1) // i.e. not StrPut(String, Encoding)
		{
			if (TokenIsNumeric(**next_param)) // Length parameter
			{
				length = (int)TokenToInt64(**next_param);
				if (!source_string) // StrGet
				{
					if (length == 0)
						return; // Get 0 chars.
					if (length < 0)
						length = -length; // Retrieve exactly this many chars, even if there are null chars.
					else
						length_is_max_size = true; // Limit to this, but stop at the first null char.
				}
				else if (length <= 0)
					_f_throw_value(ERR_INVALID_LENGTH);
				++next_param; // Let encoding be the next param, if present.
			}
			else if ((*next_param)->symbol == SYM_MISSING)
			{
				// Length was "explicitly omitted", as in StrGet(Address,, Encoding),
				// which allows Encoding to be an integer without specifying Length.
				++next_param;
			}
			// aParam now points to aParam_end or the Encoding param.
		}
		if (next_param < aParam_end)
		{
			encoding = Line::ConvertFileEncoding(**next_param);
			if (encoding == -1)
				_f_throw_value(ERR_INVALID_ENCODING);
		}
	}
	// Note: CP_AHKNOBOM is not supported; "-RAW" must be omitted.

	// Check for obvious errors to prevent an Access Violation.
	// Address can be zero for StrPut if length is also zero (see below).
	if ( address < FIRST_VALID_ADDRESS
		// Also check for overlap, in case memcpy is used instead of MultiByteToWideChar/WideCharToMultiByte.
		// (Behaviour for memcpy would be "undefined", whereas MBTWC/WCTBM would fail.)  Overlap in the
		// other direction (source_string beginning inside address..length) should not be possible.
		|| (address >= source_string && address <= ((LPTSTR)source_string + source_length))
		// The following catches StrPut(X, &Y) where Y is uninitialized or has zero capacity.
		|| (address == Var::sEmptyString && source_length) )
	{
		_f_throw_param(source_string ? 1 : 0);
	}

	if (max_bytes != SIZE_MAX)
	{
		// Target is a Buffer object with known size, so limit length accordingly.
		int max_chars = int(max_bytes >> int(encoding == CP_UTF16));
		if (length > max_chars)
			_f_throw_value(ERR_INVALID_LENGTH);
		if (source_length > max_chars)
			_f_throw_param(1);
		if (length == -1)
		{
			length = max_chars;
			length_is_max_size = true;
		}
	}

	if (source_string) // StrPut
	{
		int char_count; // Either bytes or characters, depending on the target encoding.
		aResultToken.symbol = SYM_INTEGER; // Most paths below return an integer.

		if (!source_length)
		{
			// Take a shortcut when source_string is empty, since some paths below might not handle it correctly.
			if (length) // true except when in measuring mode.
			{
				if (encoding == CP_UTF16)
					*(LPWSTR)address = '\0';
				else
					*(LPSTR)address = '\0';
			}
			aResultToken.value_int64 = encoding == CP_UTF16 ? sizeof(WCHAR) : sizeof(CHAR);
			return;
		}

		if (encoding == UorA(CP_UTF16, CP_ACP))
		{
			// No conversion required: target encoding is the same as the native encoding of this build.
			char_count = source_length + 1; // + 1 because generally a null-terminator is wanted.
			if (length)
			{
				// Check for sufficient buffer space.  Cast to UINT and compare unsigned values: if length is
				// -1 it should be interpreted as a very large unsigned value, in effect bypassing this check.
				if ((UINT)source_length <= (UINT)length)
				{
					if (source_length == length)
						// Exceptional case: caller doesn't want a null-terminator (or passed this length in error).
						--char_count;
					// Copy the string, including null-terminator if requested.
					tmemcpy((LPTSTR)address, (LPCTSTR)source_string, char_count);
				}
				else
					// For consistency with the sections below, don't truncate the string.
					_f_throw_value(ERR_INVALID_LENGTH);
			}
			//else: Caller just wants the the required buffer size (char_count), which will be returned below.
			//	Note that although this seems equivalent to StrLen(), the caller might have explicitly
			//	passed an Encoding; in that case, the result of StrLen() might be different on the
			//	opposite build (ANSI vs Unicode) as the section below would be executed instead of this one.
		}
		else
		{
			// Conversion is required. For Unicode builds, this means encoding != CP_UTF16;
#ifndef UNICODE // therefore, this section is relevant only to ANSI builds:
			if (encoding == CP_UTF16)
			{
				// See similar section below for comments.
				if (length <= 0)
				{
					char_count = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)source_string, source_length, NULL, 0) + 1;
					if (length == 0)
					{
						aResultToken.value_int64 = char_count * (1 + (encoding == CP_UTF16));
						return;
					}
					length = char_count;
				}
				char_count = MultiByteToWideChar(CP_ACP, 0, (LPCSTR)source_string, source_length, (LPWSTR)address, length);
				if (char_count && char_count < length)
					((LPWSTR)address)[char_count++] = '\0';
			}
			else // encoding != CP_UTF16
			{
				// Convert native ANSI string to UTF-16 first.
				CStringWCharFromChar wide_buf((LPCSTR)source_string, source_length, CP_ACP);				
				source_string = wide_buf.GetString();
				source_length = wide_buf.GetLength();
#endif
				// UTF-8 does not support this flag.  Although the check further below would probably
				// compensate for this, UTF-8 is probably common enough to leave this exception here.
				DWORD flags = (encoding == CP_UTF8) ? 0 : WC_NO_BEST_FIT_CHARS;
				if (length <= 0) // -1 or 0
				{
					// Determine required buffer size.
					char_count = WideCharToMultiByte(encoding, flags, (LPCWSTR)source_string, source_length, NULL, 0, NULL, NULL);
					if (!char_count) // Above has ensured source is not empty, so this must be an error.
					{
						if (GetLastError() == ERROR_INVALID_FLAGS)
						{
							// Try again without flags.  MSDN lists a number of code pages for which flags must be 0, including UTF-7 and UTF-8 (but UTF-8 is handled above).
							flags = 0; // Must be set for this call and the call further below.
							char_count = WideCharToMultiByte(encoding, flags, (LPCWSTR)source_string, source_length, NULL, 0, NULL, NULL);
						}
						if (!char_count)
							_f_throw_win32();
					}
					++char_count; // + 1 for null-terminator (source_length causes it to be excluded from char_count).
					if (length == 0) // Caller just wants the required buffer size.
					{
						aResultToken.value_int64 = char_count * (1 + (encoding == CP_UTF16));
						return;
					}
					// Assume there is sufficient buffer space and hope for the best:
					length = char_count;
				}
				// Convert to target encoding.
				char_count = WideCharToMultiByte(encoding, flags, (LPCWSTR)source_string, source_length, (LPSTR)address, length, NULL, NULL);
				// Since above did not null-terminate, check for buffer space and null-terminate if there's room.
				// It is tempting to always null-terminate (potentially replacing the last byte of data),
				// but that would exclude this function as a means to copy a string into a fixed-length array.
				if (char_count && char_count < length)
					((LPSTR)address)[char_count++] = '\0';
				// else no space to null-terminate; or conversion failed.
#ifndef UNICODE
			}
#endif
			if (!char_count)
				_f_throw_win32();
		}
		// Return the number of bytes written.
		aResultToken.value_int64 = char_count * (1 + (encoding == CP_UTF16));
	}
	else // StrGet
	{
		if (length_is_max_size) // Implies length != -1.
		{
			// Caller specified the maximum character count, not the exact length.
			// If the length includes null characters, the conversion functions below
			// would convert more than necessary and we'd still have to recalculate the
			// length.  So find the exact length up front:
			if (encoding == CP_UTF16)
				length = (int)wcsnlen((LPWSTR)address, length);
			else
				length = (int)strnlen((LPSTR)address, length);
		}
		if (encoding != UorA(CP_UTF16, CP_ACP))
		{
			// Conversion is required.
			int conv_length;
			// MS docs: "Note that, if cbMultiByte is 0, the function fails."
			if (!length)
				_f_return_empty;
#ifdef UNICODE
			// Convert multi-byte encoded string to UTF-16.
			conv_length = MultiByteToWideChar(encoding, 0, (LPCSTR)address, length, NULL, 0);
			if (!TokenSetResult(aResultToken, NULL, conv_length)) // DO NOT SUBTRACT 1, conv_length might not include a null-terminator.
				return; // Out of memory.
			conv_length = MultiByteToWideChar(encoding, 0, (LPCSTR)address, length, aResultToken.marker, conv_length);
#else
			CStringW wide_buf;
			// If the target string is not UTF-16, convert it to that first.
			if (encoding != CP_UTF16)
			{
				StringCharToWChar((LPCSTR)address, wide_buf, length, encoding);
				address = (void *)wide_buf.GetString();
				length = wide_buf.GetLength();
			}

			// Now convert UTF-16 to ACP.
			conv_length = WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, (LPCWSTR)address, length, NULL, 0, NULL, NULL);
			if (!TokenSetResult(aResultToken, NULL, conv_length)) // DO NOT SUBTRACT 1, conv_length might not include a null-terminator.
			{
				aResult = FAIL;
				return; // Out of memory.
			}
			conv_length = WideCharToMultiByte(CP_ACP, WC_NO_BEST_FIT_CHARS, (LPCWSTR)address, length, aResultToken.marker, conv_length, NULL, NULL);
#endif
			if (!conv_length) // This can only be failure, since ... (see below)
				_f_throw_win32();
			if (length == -1) // conv_length includes a null-terminator in this case.
				--conv_length;
			else
				aResultToken.marker[conv_length] = '\0'; // It wasn't terminated above.
			aResultToken.marker_length = conv_length; // Update it.
		}
		else if (length == -1)
		{
			// Return this null-terminated string, no conversion necessary.
			aResultToken.marker = (LPTSTR) address;
			aResultToken.marker_length = _tcslen(aResultToken.marker);
		}
		else
		{
			// No conversion necessary, but we might not want the whole string.
			// Copy and null-terminate the string; some callers might require it.
			TokenSetResult(aResultToken, (LPCTSTR)address, length);
		}
	}
}



BIF_DECL(BIF_StrPtr)
{
	switch (aParam[0]->symbol)
	{
	case SYM_STRING:
		_f_return((UINT_PTR)aParam[0]->marker);
	case SYM_VAR:
		if (!aParam[0]->var->IsPureNumericOrObject())
			_f_return((UINT_PTR)aParam[0]->var->Contents());
	default:
		_f_throw_type(_T("String"), *aParam[0]);
	}
}



////////////////////
// Misc Functions //
////////////////////


BIF_DECL(BIF_IsLabel)
{
	_f_return_b(g_script.FindLabel(ParamIndexToString(0, _f_number_buf)) ? 1 : 0);
}


IObject *UserFunc::CloseIfNeeded()
{
	FreeVars *fv = (mUpVarCount && mOuterFunc && sFreeVars) ? sFreeVars->ForFunc(mOuterFunc) : NULL;
	if (!fv)
	{
		AddRef();
		return this;
	}
	return new Closure(this, fv, false);
}



BIF_DECL(BIF_IsTypeish)
{
	auto variable_type = (VariableTypeType)_f_callee_id;
	bool if_condition;
	TCHAR *cp;

	// The first set of checks are for isNumber(), isInteger() and isFloat(), which permit pure numeric values.
	switch (TypeOfToken(*aParam[0]))
	{
	case SYM_INTEGER:
		switch (variable_type)
		{
		case VAR_TYPE_NUMBER:
		case VAR_TYPE_INTEGER:
			_f_return_b(true);
		case VAR_TYPE_FLOAT:
			_f_return_b(false);
		default:
			// Do not permit pure numbers for the other functions, since the results would not be intuitive.
			// For instance, isAlnum() would return false for negative values due to '-'; isXDigit() would
			// return true for positive integers even though they are always in decimal.
			goto type_mismatch;
		}
	case SYM_FLOAT:
		switch (variable_type)
		{
		case VAR_TYPE_NUMBER:
		case VAR_TYPE_FLOAT:
			_f_return_b(true);
		case VAR_TYPE_INTEGER:
			// Given that legacy "if var is float" required a decimal point in var and isFloat() is false for
			// integers which can be represented as float, it seems inappropriate for isInteger(1.0) to be true.
			// A function like isWholeNumber() could be added if that was needed.
			_f_return_b(false);
		default:
			goto type_mismatch;
		}
	case SYM_OBJECT:
		switch (variable_type)
		{
		case VAR_TYPE_NUMBER:
		case VAR_TYPE_INTEGER:
		case VAR_TYPE_FLOAT:
			_f_return_b(false);
		default:
			goto type_mismatch;
		}
	}
	// Since above did not return or goto, the value is a string.
	LPTSTR aValueStr = ParamIndexToString(0);
	auto string_case_sense = ParamIndexToCaseSense(1); // For IsAlpha, IsAlnum, IsUpper, IsLower.
	switch (string_case_sense)
	{
	case SCS_INSENSITIVE: // This case also executes when the parameter is omitted, such as for functions which don't have this parameter.
	case SCS_SENSITIVE: // Insensitive vs. sensitive doesn't mean anything for these functions, but seems fair to allow either, rather than requiring 0 or deviating from the CaseSense convention by requiring "".
	case SCS_INSENSITIVE_LOCALE: // 'Locale'
		break;
	default:
		_f_throw_param(1);
	}

	// The remainder of this function is based on the original code for ACT_IFIS, which was removed
	// in commit 3382e6e2.
	switch (variable_type)
	{
	case VAR_TYPE_NUMBER:
		if_condition = IsNumeric(aValueStr, true, false, true);
		break;
	case VAR_TYPE_INTEGER:
		if_condition = IsNumeric(aValueStr, true, false, false);  // Passes false for aAllowFloat.
		break;
	case VAR_TYPE_FLOAT:
		if_condition = (IsNumeric(aValueStr, true, false, true) == PURE_FLOAT);
		break;
	case VAR_TYPE_TIME:
	{
		SYSTEMTIME st;
		// Also insist on numeric, because even though YYYYMMDDToFileTime() will properly convert a
		// non-conformant string such as "2004.4", for future compatibility, we don't want to
		// report that such strings are valid times:
		if_condition = IsNumeric(aValueStr, false, false, false) && YYYYMMDDToSystemTime(aValueStr, st, true);
		break;
	}
	case VAR_TYPE_DIGIT:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (*cp < '0' || *cp > '9') // Avoid iswdigit; as documented, only ASCII digits 0 .. 9 are permitted.
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_XDIGIT:
		cp = aValueStr;
		if (!_tcsnicmp(cp, _T("0x"), 2)) // Allow 0x prefix.
			cp += 2;
		if_condition = true;
		for (; *cp; ++cp)
			if (!cisxdigit(*cp)) // Avoid iswxdigit; as documented, only ASCII xdigits are permitted.
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_ALNUM:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (string_case_sense == SCS_INSENSITIVE_LOCALE ? !IsCharAlphaNumeric(*cp) : !cisalnum(*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_ALPHA:
		// Like AutoIt3, the empty string is considered to be alphabetic, which is only slightly debatable.
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (string_case_sense == SCS_INSENSITIVE_LOCALE ? !IsCharAlpha(*cp) : !cisalpha(*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_UPPER:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (string_case_sense == SCS_INSENSITIVE_LOCALE ? !IsCharUpper(*cp) : !cisupper(*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_LOWER:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (string_case_sense == SCS_INSENSITIVE_LOCALE ? !IsCharLower(*cp) : !cislower(*cp))
			{
				if_condition = false;
				break;
			}
		break;
	case VAR_TYPE_SPACE:
		if_condition = true;
		for (cp = aValueStr; *cp; ++cp)
			if (!_istspace(*cp))
			{
				if_condition = false;
				break;
			}
		break;
#ifdef DEBUG
	default:
		MsgBox(_T("DEBUG: Unhandled IsXStr mode."));
#endif
	}
	_f_return_b(if_condition);

type_mismatch:
	_f_throw_type(_T("String"), *aParam[0]);
}



BIF_DECL(BIF_IsSet)
{
	Var *var = ParamIndexToOutputVar(0);
	if (!var)
		_f_throw_param(0, _T("variable reference"));
	_f_return_b(!var->IsUninitializedNormalVar());
}



////////////////////////
// Keyboard Functions //
////////////////////////


BIF_DECL(BIF_GetKeyState)
{
	TCHAR key_name_buf[MAX_NUMBER_SIZE]; // Because _f_retval_buf is used for something else below.
	LPTSTR key_name = ParamIndexToString(0, key_name_buf);
	// Keep this in sync with GetKeyJoyState().
	// See GetKeyJoyState() for more comments about the following lines.
	JoyControls joy;
	int joystick_id;
	vk_type vk = TextToVK(key_name);
	if (!vk)
	{
		if (   !(joy = (JoyControls)ConvertJoy(key_name, &joystick_id))   )
		{
			// It is neither a key name nor a joystick button/axis.
			_f_throw_param(0);
		}
		ScriptGetJoyState(joy, joystick_id, aResultToken, _f_retval_buf);
		return;
	}
	// Since above didn't return: There is a virtual key (not a joystick control).
	TCHAR mode_buf[MAX_NUMBER_SIZE];
	LPTSTR mode = ParamIndexToOptionalString(1, mode_buf);
	KeyStateTypes key_state_type;
	switch (ctoupper(*mode)) // Second parameter.
	{
	case 'T': key_state_type = KEYSTATE_TOGGLE; break; // Whether a toggleable key such as CapsLock is currently turned on.
	case 'P': key_state_type = KEYSTATE_PHYSICAL; break; // Physical state of key.
	default: key_state_type = KEYSTATE_LOGICAL;
	}
	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	aResultToken.value_int64 = ScriptGetKeyState(vk, key_state_type); // 1 for down and 0 for up.
}



BIF_DECL(BIF_GetKeyName)
{
	// Get VK and/or SC from the first parameter, which may be a key name, scXXX or vkXX.
	// Key names are allowed even for GetKeyName() for simplicity and so that it can be
	// used to normalise a key name; e.g. GetKeyName("Esc") returns "Escape".
	LPTSTR key = ParamIndexToString(0, _f_number_buf);
	vk_type vk;
	sc_type sc;
	TextToVKandSC(key, vk, sc);

	switch (_f_callee_id)
	{
	case FID_GetKeyVK:
		_f_return_i(vk ? vk : sc_to_vk(sc));
	case FID_GetKeySC:
		_f_return_i(sc ? sc : vk_to_sc(vk));
	//case FID_GetKeyName:
	default:
		_f_return_p(GetKeyName(vk, sc, _f_retval_buf, _f_retval_buf_size, _T("")));
	}
}



////////////////////
// Misc Functions //
////////////////////

BIF_DECL(BIF_VarSetStrCapacity)
// Returns: The variable's new capacity.
// Parameters:
// 1: Target variable (unquoted).
// 2: Requested capacity.
{
	Var *target_var = ParamIndexToOutputVar(0);
	// Redundant due to prior validation of OutputVars:
	//if (!target_var)
	//	_f_throw_param(0, _T("variable reference"));
	Var &var = *target_var;
	ASSERT(var.Type() == VAR_NORMAL); // Should always be true.

	if (aParamCount > 1) // Second parameter is present.
	{
		__int64 param1 = TokenToInt64(*aParam[1]);
		// Check for invalid values, in particular small negatives which end up large when converted
		// to unsigned.  Var::AssignString used to have a similar check, but integer overflow caused
		// by "* sizeof(TCHAR)" allowed some errors to go undetected.  For this same reason, we can't
		// simply rely on SetCapacity() calling malloc() and then detecting failure.
		if ((unsigned __int64)param1 > (MAXINT_PTR / sizeof(TCHAR)))
		{
			if (param1 == -1) // Adjust variable's internal length.
			{
				var.SetLengthFromContents();
				// Seems more useful to report length vs. capacity in this special case. Scripts might be able
				// to use this to boost performance.
				aResultToken.value_int64 = var.CharLength();
				return;
			}
			// x64: it's negative but not -1.
			// x86: it's either >2GB or negative and not -1.
			_f_throw_param(1);
		}
		// Since above didn't return:
		size_t new_capacity = (size_t)param1 * sizeof(TCHAR); // Chars to bytes.
		if (new_capacity)
		{
			if (!var.SetCapacity(new_capacity, true)) // This also destroys the variables contents.
			{
				aResultToken.SetExitResult(FAIL); // ScriptError() was already called.
				return;
			}
			// By design, Assign() has already set the length of the variable to reflect new_capacity.
			// This is not what is wanted in this case since it should be truly empty.
			var.ByteLength() = 0;
		}
		else // ALLOC_SIMPLE, due to its nature, will not actually be freed, which is documented.
			var.Free();
	} // if (aParamCount > 1)
	else
	{
		// RequestedCapacity was omitted, so the var is not altered; instead, the current capacity
		// is reported, which seems more intuitive/useful than having it do a Free(). In this case
		// it's an input var rather than an output var, so check if it contains a string.
		if (var.IsPureNumericOrObject())
			_f_throw_type(_T("String"), *aParam[0]);
	}

	// Caller has set aResultToken.symbol to a default of SYM_INTEGER, so no need to set it here.
	if (aResultToken.value_int64 = var.CharCapacity()) // Don't subtract 1 here in lieu doing it below (avoids underflow).
		aResultToken.value_int64 -= 1; // Omit the room for the zero terminator since script capacity is defined as length vs. size.
}



////////////////////
// File Functions //
////////////////////

BIF_DECL(BIF_FileExist)
{
	TCHAR filename_buf[MAX_NUMBER_SIZE]; // Because _f_number_buf is used for something else below.
	LPTSTR filename = ParamIndexToString(0, filename_buf);
	LPTSTR buf = _f_retval_buf; // If necessary, it will be moved to a persistent memory location by our caller.
	DWORD attr;
	if (DoesFilePatternExist(filename, &attr, _f_callee_id == FID_DirExist ? FILE_ATTRIBUTE_DIRECTORY : 0))
	{
		// Yield the attributes of the first matching file.  If not match, yield an empty string.
		// This relies upon the fact that a file's attributes are never legitimately zero, which
		// seems true but in case it ever isn't, this forces a non-empty string be used.
		// UPDATE for v1.0.44.03: Someone reported that an existing file (created by NTbackup.exe) can
		// apparently have undefined/invalid attributes (i.e. attributes that have no matching letter in
		// "RASHNDOCT").  Although this is unconfirmed, it's easy to handle that possibility here by
		// checking for a blank string.  This allows FileExist() to report boolean TRUE rather than FALSE
		// for such "mystery files":
		FileAttribToStr(buf, attr);
		if (!*buf) // See above.
		{
			// The attributes might be all 0, but more likely the file has some of the newer attributes
			// such as FILE_ATTRIBUTE_ENCRYPTED (or has undefined attributes).  So rather than storing attr as
			// a hex number (which could be zero and thus defeat FileExist's ability to detect the file), it
			// seems better to store some arbitrary letter (other than those in "RASHNDOCT") so that FileExist's
			// return value is seen as boolean "true".
			buf[0] = 'X';
			buf[1] = '\0';
		}
	}
	else // Empty string is the indicator of "not found" (seems more consistent than using an integer 0, since caller might rely on it being SYM_STRING).
		*buf = '\0';
	_f_return_p(buf);
}





///////////////////////
// Interop Functions //
///////////////////////


BIF_DECL(BIF_MenuFromHandle)
{
	auto *menu = g_script.FindMenu((HMENU)ParamIndexToInt64(0));
	if (menu)
	{
		menu->AddRef();
		_f_return(menu);
	}
	_f_return_empty;
}



//////////////////////
// Gui & GuiControl //
//////////////////////


void GuiControlType::StatusBar(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
{
	GuiType& gui = *this->gui;
	HWND control_hwnd = this->hwnd;
	LPTSTR buf = _f_number_buf;

	HICON hicon;
	switch (aID)
	{
	case FID_SB_SetText:
		_o_return(SendMessage(control_hwnd, SB_SETTEXT
			, (WPARAM)((ParamIndexIsOmitted(1) ? 0 : ParamIndexToInt64(1) - 1) // The Part# param is present.
				     | (ParamIndexIsOmitted(2) ? 0 : ParamIndexToInt64(2) << 8)) // The uType parameter is present.
			, (LPARAM)ParamIndexToString(0, buf))); // Caller has ensured that there's at least one param in this mode.

	case FID_SB_SetParts:
		LRESULT old_part_count, new_part_count;
		int edge, part[256]; // Load-time validation has ensured aParamCount is under 255, so it shouldn't overflow.
		for (edge = 0, new_part_count = 0; new_part_count < aParamCount; ++new_part_count)
		{
			edge += gui.Scale(ParamIndexToInt(new_part_count)); // For code simplicity, no check for negative (seems fairly harmless since the bar will simply show up with the wrong number of parts to indicate the problem).
			part[new_part_count] = edge;
		}
		// For code simplicity, there is currently no means to have the last part of the bar use less than
		// all of the bar's remaining width.  The desire to do so seems rare, especially since the script can
		// add an extra/unused part at the end to achieve nearly (or perhaps exactly) the same effect.
		part[new_part_count++] = -1; // Make the last part use the remaining width of the bar.

		old_part_count = SendMessage(control_hwnd, SB_GETPARTS, 0, NULL); // MSDN: "This message always returns the number of parts in the status bar [regardless of how it is called]".
		if (old_part_count > new_part_count) // Some parts are being deleted, so destroy their icons.  See other notes in GuiType::Destroy() for explanation.
			for (LRESULT i = new_part_count; i < old_part_count; ++i) // Verified correct.
				if (hicon = (HICON)SendMessage(control_hwnd, SB_GETICON, i, 0))
					DestroyIcon(hicon);

		_o_return(SendMessage(control_hwnd, SB_SETPARTS, new_part_count, (LPARAM)part)
			? (__int64)control_hwnd : 0); // Return HWND to provide an easy means for the script to get the bar's HWND.

	//case FID_SB_SetIcon:
	default:
		int unused, icon_number;
		icon_number = ParamIndexToOptionalInt(1, 1);
		if (icon_number == 0) // Must be != 0 to tell LoadPicture that "icon must be loaded, never a bitmap".
			icon_number = 1;
		if (hicon = (HICON)LoadPicture(ParamIndexToString(0, buf) // Load-time validation has ensured there is at least one parameter.
			, GetSystemMetrics(SM_CXSMICON), GetSystemMetrics(SM_CYSMICON) // Apparently the bar won't scale them for us.
			, unused, icon_number, false)) // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
		{
			WPARAM part_index = ParamIndexIsOmitted(2) ? 0 : (WPARAM)ParamIndexToInt64(2) - 1;
			HICON hicon_old = (HICON)SendMessage(control_hwnd, SB_GETICON, part_index, 0); // Get the old one before setting the new one.
			// For code simplicity, the script is responsible for destroying the hicon later, if it ever destroys
			// the window.  Though in practice, most people probably won't do this, which is usually okay (if the
			// script doesn't load too many) since they're all destroyed by the system upon program termination.
			if (SendMessage(control_hwnd, SB_SETICON, part_index, (LPARAM)hicon))
			{
				if (hicon_old)
					// Although the old icon is automatically destroyed here, the script can call SendMessage(SB_SETICON)
					// itself if it wants to work with HICONs directly (for performance reasons, etc.)
					DestroyIcon(hicon_old);
			}
			else
			{
				DestroyIcon(hicon);
				hicon = NULL;
			}
		}
		//else can't load icon, so return 0.
		_o_return((size_t)hicon); // This allows the script to destroy the HICON later when it doesn't need it (see comments above too).
	// SB_SetTipText() not implemented (though can be done via SendMessage in the script) because the conditions
	// under which tooltips are displayed don't seem like something a script would want very often:
	// This ToolTip text is displayed in two situations: 
	// When the corresponding pane in the status bar contains only an icon. 
	// When the corresponding pane in the status bar contains text that is truncated due to the size of the pane.
	// In spite of the above, SB_SETTIPTEXT doesn't actually seem to do anything, even when the text is too long
	// to fit in a narrowed part, tooltip text has been set, and the user hovers the cursor over the bar.  Maybe
	// I'm not doing it right or maybe this feature is somehow disabled under certain service packs or conditions.
	//case 'T': // SB.SetTipText()
	//	break;
	} // switch(mode)
}


void GuiControlType::LV_GetNextOrCount(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// LV.GetNext:
// Returns: The index of the found item, or 0 on failure.
// Parameters:
// 1: Starting index (one-based when it comes in).  If absent, search starts at the top.
// 2: Options string.
// 3: (FUTURE): Possible for use with LV.FindItem (though I think it can only search item text, not subitem text).
{
	HWND control_hwnd = hwnd;

	LPTSTR options;
	if (aID == FID_LV_GetCount)
	{
		options = (aParamCount > 0) ? omit_leading_whitespace(ParamIndexToString(0, _f_number_buf)) : _T("");
		if (*options)
		{
			if (ctoupper(*options) == 'S')
				_o_return(SendMessage(control_hwnd, LVM_GETSELECTEDCOUNT, 0, 0));
			else if (!_tcsnicmp(options, _T("Col"), 3)) // "Col" or "Column". Don't allow "C" by itself, so that "Checked" can be added in the future.
				_o_return(union_lv_attrib->col_count);
			else
				_o_throw_param(0);
		}
		_o_return(SendMessage(control_hwnd, LVM_GETITEMCOUNT, 0, 0));
	}
	// Since above didn't return, this is GetNext() mode.

	int index = -1;
	if (!ParamIndexIsOmitted(0))
	{
		if (!ParamIndexIsNumeric(0))
			_o_throw_param(0, _T("Number"));
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
		// For flexibility, allow index to be less than -1 to avoid first-iteration complications in script loops
		// (such as when deleting rows, which shifts the row index upward, require the search to resume at
		// the previously found index rather than the row after it).  However, reset it to -1 to ensure
		// proper return values from the API in the "find checked item" mode used below.
		if (index < -1)
			index = -1;  // Signal it to start at the top.
	}

	// For performance, decided to always find next selected item when the "C" option hasn't been specified,
	// even when the checkboxes style is in effect.  Otherwise, would have to fetch and check checkbox style
	// bit for each call, which would slow down this heavily-called function.

	options = ParamIndexToOptionalString(1, _f_number_buf);
	TCHAR first_char = ctoupper(*omit_leading_whitespace(options));
	// To retain compatibility in the future, also allow "Check(ed)" and "Focus(ed)" since any word that
	// starts with C or F is already supported.

	switch(first_char)
	{
	case '\0': // Listed first for performance.
	case 'F':
		_o_return(ListView_GetNextItem(control_hwnd, index
			, first_char ? LVNI_FOCUSED : LVNI_SELECTED) + 1); // +1 to convert to 1-based.
	case 'C': // Checkbox: Find checked items. For performance assume that the control really has checkboxes.
	  {
		int item_count = ListView_GetItemCount(control_hwnd);
		for (int i = index + 1; i < item_count; ++i) // Start at index+1 to omit the first item from the search (for consistency with the other mode above).
			if (ListView_GetCheckState(control_hwnd, i)) // Item's box is checked.
				_o_return(i + 1); // +1 to convert from zero-based to one-based.
		// Since above didn't return, no match found.
		_o_return(0);
	  }
	default:
		_o_throw_param(1);
	}
}



void GuiControlType::LV_GetText(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: Text on success.
// Throws on failure.
// Parameters:
// 1: Row index (one-based when it comes in).
// 2: Column index (one-based when it comes in).
{
	// Above sets default result in case of early return.  For code reduction, a zero is returned for all
	// the following conditions:
	// Item not found in ListView.
	// And others.

	int row_index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	if (row_index < -1) // row_index==-1 is reserved to mean "get column heading's text".
		_o_throw_param(0);
	// If parameter 2 is omitted, default to the first column (index 0):
	int col_index = ParamIndexIsOmitted(1) ? 0 : ParamIndexToInt(1) - 1; // -1 to convert to zero-based.
	if (col_index < 0)
		_o_throw_param(1);

	TCHAR buf[LV_TEXT_BUF_SIZE];

	if (row_index == -1) // Special mode to get column's text.
	{
		LVCOLUMN lvc;
		lvc.cchTextMax = LV_TEXT_BUF_SIZE - 1;  // See notes below about -1.
		lvc.pszText = buf;
		lvc.mask = LVCF_TEXT;
		if (SendMessage(hwnd, LVM_GETCOLUMN, col_index, (LPARAM)&lvc)) // Assign.
			_o_return(lvc.pszText); // See notes below about why pszText is used instead of buf (might apply to this too).
		else // On failure, it seems best to throw.
			_o_throw(ERR_FAILED);
	}
	else // Get row's indicated item or subitem text.
	{
		LVITEM lvi;
		// Subtract 1 because of that nagging doubt about size vs. length. Some MSDN examples subtract one, such as
		// TabCtrl_GetItem()'s cchTextMax:
		lvi.iItem = row_index;
		lvi.iSubItem = col_index; // Which field to fetch.  If it's zero, the item vs. subitem will be fetched.
		lvi.mask = LVIF_TEXT;
		lvi.pszText = buf;
		lvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // Note that LVM_GETITEM doesn't update this member to reflect the new length.
		// Unlike LVM_GETITEMTEXT, LVM_GETITEM indicates success or failure, which seems more useful/preferable
		// as a return value since a text length of zero would be ambiguous: could be an empty field or a failure.
		if (SendMessage(hwnd, LVM_GETITEM, 0, (LPARAM)&lvi)) // Assign
			// Must use lvi.pszText vs. buf because MSDN says: "Applications should not assume that the text will
			// necessarily be placed in the specified buffer. The control may instead change the pszText member
			// of the structure to point to the new text rather than place it in the buffer."
			_o_return(lvi.pszText);
		else // On failure, it seems best to throw.
			_o_throw(ERR_FAILED);
	}
}



void GuiControlType::LV_AddInsertModify(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: For Add(), this is the options.  For Insert/Modify, it's the row index (one-based when it comes in).
// 2: For Add(), this is the first field's text.  For Insert/Modify, it's the options.
// 3 and beyond: Additional field text.
// In Add/Insert mode, if there are no text fields present, a blank for is appended/inserted.
{
	GuiControlType &control = *this;
	auto mode = BuiltInFunctionID(aID);
	LPTSTR buf = _f_number_buf; // Resolve macro early for maintainability.

	int index;
	if (mode == FID_LV_Add) // For Add mode, use INT_MAX as a signal to append the item rather than inserting it.
	{
		index = INT_MAX;
		mode = FID_LV_Insert; // Add has now been set up to be the same as insert, so change the mode to simplify other things.
	}
	else // Insert or Modify: the target row-index is their first parameter, which load-time has ensured is present.
	{
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
		if (index < -1 || (mode != FID_LV_Modify && index < 0)) // Allow -1 to mean "all rows" when in modify mode.
			_o_throw_param(0);
		++aParam;  // Remove the first parameter from further consideration to make Insert/Modify symmetric with Add.
		--aParamCount;
	}

	LPTSTR options = ParamIndexToOptionalString(0, buf);
	bool ensure_visible = false, is_checked = false;  // Checkmark.
	int col_start_index = 0;
	LVITEM lvi;
	lvi.mask = LVIF_STATE; // LVIF_STATE: state member is valid, but only to the extent that corresponding bits are set in stateMask (the rest will be ignored).
	lvi.stateMask = 0;
	lvi.state = 0;

	// Parse list of space-delimited options:
	TCHAR *next_option, *option_end, orig_char;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
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
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		if (!_tcsnicmp(next_option, _T("Select"), 6)) // Could further allow "ed" suffix by checking for that inside, but "Selected" is getting long so it doesn't seem something many would want to use.
		{
			next_option += 6;
			// If it's Select0, invert the mode to become "no select". This allows a boolean variable
			// to be more easily applied, such as this expression: "Select" . VarContainingState
			if (*next_option && !ATOI(next_option))
				adding = !adding;
			// Another reason for not having "Select" imply "Focus" by default is that it would probably
			// reduce performance when selecting all or a large number of rows.
			// Because a row might or might not have focus, the script may wish to retain its current
			// focused state.  For this reason, "select" does not imply "focus", which allows the
			// LVIS_FOCUSED bit to be omitted from the stateMask, which in turn retains the current
			// focus-state of the row rather than disrupting it.
			lvi.stateMask |= LVIS_SELECTED;
			if (adding)
				lvi.state |= LVIS_SELECTED;
			//else removing, so the presence of LVIS_SELECTED in the stateMask above will cause it to be de-selected.
		}
		else if (!_tcsnicmp(next_option, _T("Focus"), 5))
		{
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Focus0, invert the mode to become "no focus".
				adding = !adding;
			lvi.stateMask |= LVIS_FOCUSED;
			if (adding)
				lvi.state |= LVIS_FOCUSED;
			//else removing, so the presence of LVIS_FOCUSED in the stateMask above will cause it to be de-focused.
		}
		else if (!_tcsnicmp(next_option, _T("Check"), 5))
		{
			// The rationale for not checking for an optional "ed" suffix here and incrementing next_option by 2
			// is that: 1) It would be inconsistent with the lack of support for "selected" (see reason above);
			// 2) Checkboxes in a ListView are fairly rarely used, so code size reduction might be more important.
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Check0, invert the mode to become "unchecked".
				adding = !adding;
			if (mode == FID_LV_Modify) // v1.0.46.10: Do this section only for Modify, not Add/Insert, to avoid generating an extra "unchecked" notification when a row is added/inserted with an initial state of "checked".  In other words, the script now receives only a "checked" notification, not an "unchecked+checked". Search on is_checked for more comments.
			{
				lvi.stateMask |= LVIS_STATEIMAGEMASK;
				lvi.state |= adding ? 0x2000 : 0x1000; // The #1 image is "unchecked" and the #2 is "checked".
			}
			is_checked = adding;
		}
		else if (!_tcsnicmp(next_option, _T("Col"), 3))
		{
			if (adding)
			{
				col_start_index = ATOI(next_option + 3) - 1; // The ability to start at a column other than 1 (i.e. subitem vs. item).
				if (col_start_index < 0)
					col_start_index = 0;
			}
		}
		else if (!_tcsnicmp(next_option, _T("Icon"), 4))
		{
			// Testing shows that there is no way to avoid having an item icon in report view if the
			// ListView has an associated small-icon ImageList (well, perhaps you could have it show
			// a blank square by specifying an invalid icon index, but that doesn't seem useful).
			// If LVIF_IMAGE is entirely omitted when adding and item/row, the item will take on the
			// first icon in the list.  This is probably by design because the control wants to make
			// each item look consistent by indenting its first field by a certain amount for the icon.
			if (adding)
			{
				lvi.mask |= LVIF_IMAGE;
				lvi.iImage = ATOI(next_option + 4) - 1;  // -1 to convert to zero-based.
			}
			//else removal of icon currently not supported (see comment above), so do nothing in order
			// to reserve "-Icon" in case a future way can be found to do it.
		}
		else if (!_tcsicmp(next_option, _T("Vis"))) // v1.0.44
		{
			// Since this option much more typically used with LV.Modify than LV.Add/Insert, the technique of
			// Vis%VarContainingOneOrZero% isn't supported, to reduce code size.
			ensure_visible = adding; // Ignored by modes other than LV.Modify(), since it's not really appropriate when adding a row (plus would add code complexity).
		}
		else
		{
			aResultToken.ValueError(ERR_INVALID_OPTION, next_option);
			*option_end = orig_char; // See comment below.
			return;
		}

		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	// Suppress any events raised by the changes made below:
	control.attrib |= GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS;

	// More maintainable and performs better to have a separate struct for subitems vs. items.
	LVITEM lvi_sub;
	// Ensure mask is pure to avoid giving it any excuse to fail due to the fact that
	// "You cannot set the state or lParam members for subitems."
	lvi_sub.mask = LVIF_TEXT;

	int i, j, rows_to_change;
	if (index == -1) // Modify all rows (above has ensured that this is only happens in modify-mode).
	{
		rows_to_change = ListView_GetItemCount(control.hwnd);
		lvi.iItem = 0;
		ensure_visible = false; // Not applicable when operating on all rows.
	}
	else // Modify or insert a single row.  Set it up for the loop to perform exactly one iteration.
	{
		rows_to_change = 1;
		lvi.iItem = index; // Which row to operate upon.  This can be a huge number such as 999999 if the caller wanted to append vs. insert.
	}
	lvi.iSubItem = 0;  // Always zero to operate upon the item vs. sub-item (subitems have their own LVITEM struct).
	int result = 1; // Set default from this point forward to be true/success. It will be overridden in insert mode to be the index of the new row.

	for (j = 0; j < rows_to_change; ++j, ++lvi.iItem) // ++lvi.iItem because if the loop has more than one iteration, by definition it is modifying all rows starting at 0.
	{
		if (!ParamIndexIsOmitted(1) && col_start_index == 0) // 2nd parameter: item's text (first field) is present, so include that when setting the item.
		{
			lvi.pszText = ParamIndexToString(1, buf); // Fairly low-overhead, so called every iteration for simplicity (so that buf can be used for both items and subitems).
			lvi.mask |= LVIF_TEXT;
		}
		if (mode == FID_LV_Insert) // Insert or Add.
		{
			// Note that ListView_InsertItem() will append vs. insert if the index is too large, in which case
			// it returns the items new index (which will be the last item in the list unless the control has
			// auto-sort style).
			if (   -1 == (lvi_sub.iItem = ListView_InsertItem(control.hwnd, &lvi))   )
			{
				control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
				_o_return(0); // Since item can't be inserted, no reason to try attaching any subitems to it.
			}
			// Update iItem with the actual index assigned to the item, which might be different than the
			// specified index if the control has an auto-sort style in effect.  This new iItem value
			// is used for ListView_SetCheckState() and for the attaching of any subitems to this item.
			result = lvi_sub.iItem + 1; // Convert from zero-based to 1-based.
			// For add/insert (but not modify), testing shows that checkmark must be added only after
			// the item has been inserted rather than provided in the lvi.state/stateMask fields.
			// MSDN confirms this by saying "When an item is added with [LVS_EX_CHECKBOXES],
			// it will always be set to the unchecked state [ignoring any value placed in bits
			// 12 through 15 of the state member]."
			if (is_checked)
				ListView_SetCheckState(control.hwnd, lvi_sub.iItem, TRUE); // TRUE = Check the row's checkbox.
				// Note that 95/NT4 systems that lack comctl32.dll 4.70+ distributed with MSIE 3.x
				// do not support LVS_EX_CHECKBOXES, so the above will have no effect for them.
		}
		else // Modify.
		{
			// Rather than trying to detect if anything was actually changed, this is called
			// unconditionally to simplify the code (ListView_SetItem() is probably very fast if it
			// discovers that lvi.mask==LVIF_STATE and lvi.stateMask==0).
			// By design (to help catch script bugs), a failure here does not revert to append mode.
			if (!ListView_SetItem(control.hwnd, &lvi)) // Returns TRUE/FALSE.
				result = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
			lvi_sub.iItem = lvi.iItem; // In preparation for modifying any subitems that need it.
			if (ensure_visible) // Seems best to do this one prior to "select" below.
				SendMessage(control.hwnd, LVM_ENSUREVISIBLE, lvi.iItem, FALSE); // PartialOK==FALSE is somewhat arbitrary.
		}

		// For each remaining parameter, assign its text to a subitem.
		// Testing shows that if the control has too few columns for all of the fields/parameters
		// present, the ones at the end are automatically ignored: they do not consume memory nor
		// do they significantly impact performance (at least on Windows XP).  For this reason, there
		// is no code above the for-loop above to reduce aParamCount if it's "too large" because
		// it might reduce flexibility (in case future/past OSes allow non-existent columns to be
		// populated, or in case current OSes allow the contents of recently removed columns to be modified).
		for (lvi_sub.iSubItem = (col_start_index > 1) ? col_start_index : 1 // Start at the first subitem unless we were told to start at or after the third column.
			// "i" starts at 2 (the third parameter) unless col_start_index is greater than 0, in which case
			// it starts at 1 (the second parameter) because that parameter has not yet been assigned to anything:
			, i = 2 - (col_start_index > 0)
			; i < aParamCount
			; ++i, ++lvi_sub.iSubItem)
		{
			if (aParam[i]->symbol == SYM_MISSING) // Omitted, such as LV.Modify(1,Opt,"One",,"Three").
				continue;
			lvi_sub.pszText = ParamIndexToString(i, buf); // Done every time through the outer loop since it's not high-overhead, and for code simplicity.
			if (!ListView_SetItem(control.hwnd, &lvi_sub) && mode != FID_LV_Insert) // Relies on short-circuit. Seems best to avoid loss of item's index in insert mode, since failure here should be rare.
				result = 0; // Indicate partial failure, but attempt to continue in case it failed for reason other than "row doesn't exist".
		}
	} // outer for()

	// When the control has no rows, work around the fact that LVM_SETITEMCOUNT delivers less than 20%
	// of its full benefit unless done after the first row is added (at least on XP SP1).  A non-zero
	// row_count_hint tells us that this message should be sent after the row has been inserted/appended:
	if (control.union_lv_attrib->row_count_hint > 0 && mode == FID_LV_Insert)
	{
		SendMessage(control.hwnd, LVM_SETITEMCOUNT, control.union_lv_attrib->row_count_hint, 0); // Last parameter should be 0 for LVS_OWNERDATA (verified if you look at the definition of ListView_SetItemCount macro).
		control.union_lv_attrib->row_count_hint = 0; // Reset so that it only gets set once per request.
	}

	control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
	_o_return(result);
}



void GuiControlType::LV_Delete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Row index (one-based when it comes in).
{
	if (ParamIndexIsOmitted(0))
		_o_return(SendMessage(hwnd, LVM_DELETEALLITEMS, 0, 0)); // Returns TRUE/FALSE.

	// Since above didn't return, there is a first parameter present.
	int index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	if (index > -1)
		_o_return(SendMessage(hwnd, LVM_DELETEITEM, index, 0)); // Returns TRUE/FALSE.
	else
		// Even if index==0, for safety, it seems best not to do a delete-all.
		_o_throw_param(0);
}



void GuiControlType::LV_InsertModifyDeleteCol(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: Column index (one-based when it comes in).
// 2: String of options
// 3: New text of column
// There are also some special modes when only zero or one parameter is present, see below.
{
	auto mode = BuiltInFunctionID(aID);
	LPTSTR buf = _f_number_buf; // Resolve macro early for maintainability.
	aResultToken.SetValue(0); // Set default return value.

	GuiControlType &control = *this;
	GuiType &gui = *control.gui;
	lv_attrib_type &lv_attrib = *control.union_lv_attrib;
	DWORD view_mode = mode != 'D' ? ListView_GetView(control.hwnd) : 0;

	int index;
	if (!ParamIndexIsOmitted(0))
		index = ParamIndexToInt(0) - 1; // -1 to convert to zero-based.
	else // Zero parameters.  Load-time validation has ensured that the 'D' (delete) mode cannot have zero params.
	{
		if (mode == FID_LV_ModifyCol)
		{
			if (view_mode != LV_VIEW_DETAILS)
				_o_return_retval; // Return 0 to indicate failure.
			// Otherwise:
			// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
			// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
			for (int i = 0; ; ++i) // Don't limit it to lv_attrib.col_count in case script added extra columns via direct API calls.
				if (!ListView_SetColumnWidth(control.hwnd, i, LVSCW_AUTOSIZE)) // Failure means last column has already been processed.
					break;
			_o_return(1); // Always successful, regardless of what happened in the loop above.
		}
		// Since above didn't return, mode must be 'I' (insert).
		index = lv_attrib.col_count; // When no insertion index was specified, append to the end of the list.
	}

	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == FID_LV_DeleteCol) // Delete a column.  In this mode, index parameter was made mandatory via load-time validation.
	{
		if (ListView_DeleteColumn(control.hwnd, index))  // Returns TRUE/FALSE.
		{
			// It's important to note that when the user slides columns around via drag and drop, the
			// column index as seen by the script is not changed.  This is fortunate because otherwise,
			// the lv_attrib.col array would get out of sync with the column indices.  Testing shows that
			// all of the following operations respect the original column index, regardless of where the
			// user may have moved the column physically: InsertCol, DeleteCol, ModifyCol.  Insert and Delete
			// shifts the indices of those columns that *originally* lay to the right of the affected column.
			if (lv_attrib.col_count > 0) // Avoid going negative, which would otherwise happen if script previously added columns by calling the API directly.
				--lv_attrib.col_count; // Must be done prior to the below.
			if (index < lv_attrib.col_count) // When a column other than the last was removed, adjust the array so that it stays in sync with actual columns.
				MoveMemory(lv_attrib.col+index, lv_attrib.col+index+1, sizeof(lv_col_type)*(lv_attrib.col_count-index));
			_f_set_retval_i(1);
		}
		_o_return_retval;
	}
	// Do this prior to checking if index is in bounds so that it can support columns beyond LV_MAX_COLUMNS:
	if (mode == FID_LV_ModifyCol && aParamCount < 2) // A single parameter is a special modify-mode to auto-size that column.
	{
		// v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
		// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items):
		if (view_mode == LV_VIEW_DETAILS)
			_f_set_retval_i(ListView_SetColumnWidth(control.hwnd, index, LVSCW_AUTOSIZE));
		//else leave retval set to 0.
		_o_return_retval;
	}
	if (mode == FID_LV_InsertCol)
	{
		if (lv_attrib.col_count >= LV_MAX_COLUMNS) // No room to insert or append.
			_o_return_retval;
		if (index >= lv_attrib.col_count) // For convenience, fall back to "append" when index too large.
			index = lv_attrib.col_count;
	}
	//else do nothing so that modification and deletion of columns that were added via script's
	// direct calls to the API can sort-of work (it's documented in the help file that it's not supported,
	// since col-attrib array can get out of sync with actual columns that way).

	if (index < 0 || index >= LV_MAX_COLUMNS) // For simplicity, do nothing else if index out of bounds.
		_o_return_retval; // Avoid array under/overflow below.

	// In addition to other reasons, must convert any numeric value to a string so that an isolated width is
	// recognized, e.g. LV.SetCol(1, old_width + 10):
	LPTSTR options = ParamIndexToOptionalString(1, buf);

	// It's done the following way so that when in insert-mode, if the column fails to be inserted, don't
	// have to remove the inserted array element from the lv_attrib.col array:
	lv_col_type temp_col = {0}; // Init unconditionally even though only needed for mode=='I'.
	lv_col_type &col = (mode == FID_LV_InsertCol) ? temp_col : lv_attrib.col[index]; // Done only after index has been confirmed in-bounds.

	LVCOLUMN lvc;
	lvc.mask = LVCF_FMT;
	if (mode == FID_LV_ModifyCol) // Fetch the current format so that it's possible to leave parts of it unaltered.
		ListView_GetColumn(control.hwnd, index, &lvc);
	else // Mode is "insert".
		lvc.fmt = 0;

	// Init defaults prior to parsing options:
	bool sort_now = false;
	int do_auto_size = (mode == FID_LV_InsertCol) ? LVSCW_AUTOSIZE_USEHEADER : 0;  // Default to auto-size for new columns.
	TCHAR sort_now_direction = 'A'; // Ascending.
	int new_justify = lvc.fmt & LVCFMT_JUSTIFYMASK; // Simplifies the handling of the justification bitfield.
	//lvc.iSubItem = 0; // Not necessary if the LVCF_SUBITEM mask-bit is absent.

	// Parse list of space-delimited options:
	TCHAR *next_option, *option_end, orig_char;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
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
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		// For simplicity, the value of "adding" is ignored for this and the other number/alignment options.
		if (!_tcsicmp(next_option, _T("Integer")))
		{
			// For simplicity, changing the col.type dynamically (since it's so rarely needed)
			// does not try to set up col.is_now_sorted_ascending so that the next click on the column
			// puts it into default starting order (which is ascending unless the Desc flag was originally
			// present).
			col.type = LV_COL_INTEGER;
			new_justify = LVCFMT_RIGHT;
		}
		else if (!_tcsicmp(next_option, _T("Float")))
		{
			col.type = LV_COL_FLOAT;
			new_justify = LVCFMT_RIGHT;
		}
		else if (!_tcsicmp(next_option, _T("Text"))) // Seems more approp. name than "Str" or "String"
			// Since "Text" is so general, it seems to leave existing alignment (Center/Right) as it is.
			col.type = LV_COL_TEXT;

		// The following can exist by themselves or in conjunction with the above.  They can also occur
		// *after* one of the above words so that alignment can be used to override the default for the type;
		// e.g. "Integer Left" to have left-aligned integers.
		else if (!_tcsicmp(next_option, _T("Right")))
			new_justify = adding ? LVCFMT_RIGHT : LVCFMT_LEFT;
		else if (!_tcsicmp(next_option, _T("Center")))
			new_justify = adding ? LVCFMT_CENTER : LVCFMT_LEFT;
		else if (!_tcsicmp(next_option, _T("Left"))) // Supported so that existing right/center column can be changed back to left.
			new_justify = LVCFMT_LEFT; // The value of "adding" seems inconsequential so is ignored.

		else if (!_tcsicmp(next_option, _T("Uni"))) // Unidirectional sort (clicking the column will not invert to the opposite direction).
			col.unidirectional = adding;
		else if (!_tcsicmp(next_option, _T("Desc"))) // Make descending order the default order (applies to uni and first click of col for non-uni).
			col.prefer_descending = adding; // So that the next click will toggle to the opposite direction.
		else if (!_tcsnicmp(next_option, _T("Case"), 4))
		{
			if (adding)
				col.case_sensitive = !_tcsicmp(next_option + 4, _T("Locale")) ? SCS_INSENSITIVE_LOCALE : SCS_SENSITIVE;
			else
				col.case_sensitive = SCS_INSENSITIVE;
		}
		else if (!_tcsicmp(next_option, _T("Logical"))) // v1.0.44.12: Supports StrCmpLogicalW() method of sorting.
			col.case_sensitive = SCS_INSENSITIVE_LOGICAL;

		else if (!_tcsnicmp(next_option, _T("Sort"), 4)) // This is done as an option vs. LV.SortCol/LV.Sort so that the column's options can be changed simultaneously with a "sort now" to refresh.
		{
			// Defer the sort until after all options have been parsed and applied.
			sort_now = true;
			if (!_tcsicmp(next_option + 4, _T("Desc")))
				sort_now_direction = 'D'; // Descending.
		}
		else if (!_tcsicmp(next_option, _T("NoSort"))) // Called "NoSort" so that there's a way to enable and disable the setting via +/-.
			col.sort_disabled = adding;

		else if (!_tcsnicmp(next_option, _T("Auto"), 4)) // No separate failure result is reported for this item.
			// In case the mode is "insert", defer auto-width of column until col exists.
			do_auto_size = _tcsicmp(next_option + 4, _T("Hdr")) ? LVSCW_AUTOSIZE : LVSCW_AUTOSIZE_USEHEADER;

		else if (!_tcsnicmp(next_option, _T("Icon"), 4))
		{
			next_option += 4;
			if (!_tcsicmp(next_option, _T("Right")))
			{
				if (adding)
					lvc.fmt |= LVCFMT_BITMAP_ON_RIGHT;
				else
					lvc.fmt &= ~LVCFMT_BITMAP_ON_RIGHT;
			}
			else // Assume its an icon number or the removal of the icon via -Icon.
			{
				if (adding)
				{
					lvc.mask |= LVCF_IMAGE;
					lvc.fmt |= LVCFMT_IMAGE; // Flag this column as displaying an image.
					lvc.iImage = ATOI(next_option) - 1; // -1 to convert to zero based.
				}
				else
					lvc.fmt &= ~LVCFMT_IMAGE; // Flag this column as NOT displaying an image.
			}
		}

		else // Handle things that are more general than the above, such as single letter options and pure numbers.
		{
			// Width does not have a W prefix to permit a naked expression to be used as the entirely of
			// options.  For example: LV.SetCol(1, old_width + 10)
			// v1.0.37: Fixed to allow floating point (although ATOI below will convert it to integer).
			if (IsNumeric(next_option, true, false, true)) // Above has already verified that *next_option can't be whitespace.
			{
				lvc.mask |= LVCF_WIDTH;
				int width = gui.Scale(ATOI(next_option));
				// Specifying a width when the column is initially added prevents the scrollbar from
				// updating on Windows 7 and 10 (but not XP).  As a workaround, initialise the width
				// to 0 and then resize it afterward.  do_auto_size is overloaded for this purpose
				// since it's already passed to ListView_SetColumnWidth().
				if (mode == 'I' && view_mode == LV_VIEW_DETAILS)
				{
					lvc.cx = 0; // Must be zero; if width is zero, ListView_SetColumnWidth() won't be called.
					do_auto_size = width; // If non-zero, this is passed to ListView_SetColumnWidth().
				}
				else
				{
					lvc.cx = width;
					do_auto_size = 0; // Turn off any auto-sizing that may have been put into effect (explicitly or by default).
				}
			}
			else
			{
				aResultToken.ValueError(ERR_INVALID_OPTION, next_option);
				*option_end = orig_char; // See comment below.
				return;
			}
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	// Apply any changed justification/alignment to the fmt bit field:
	lvc.fmt = (lvc.fmt & ~LVCFMT_JUSTIFYMASK) | new_justify;

	if (!ParamIndexIsOmitted(2)) // Parameter #3 (text) is present.
	{
		lvc.pszText = ParamIndexToString(2, buf);
		lvc.mask |= LVCF_TEXT;
	}

	if (mode == FID_LV_ModifyCol) // Modify vs. Insert (Delete was already returned from, higher above).
		// For code simplicity, this is called unconditionally even if nothing internal the control's column
		// needs updating.  This seems justified given how rarely columns are modified.
		_f_set_retval_i(ListView_SetColumn(control.hwnd, index, &lvc)); // Returns TRUE/FALSE.
	else // Insert
	{
		// It's important to note that when the user slides columns around via drag and drop, the
		// column index as seen by the script is not changed.  This is fortunate because otherwise,
		// the lv_attrib.col array would get out of sync with the column indices.  Testing shows that
		// all of the following operations respect the original column index, regardless of where the
		// user may have moved the column physically: InsertCol, DeleteCol, ModifyCol.  Insert and Delete
		// shifts the indices of those columns that *originally* lay to the right of the affected column.
		// Doesn't seem to do anything -- not even with respect to inserting a new first column with it's
		// unusual behavior of inheriting the previously column's contents -- so it's disabled for now.
		// Testing shows that it also does not seem to cause a new column to inherit the indicated subitem's
		// text, even when iSubItem is set to index + 1 vs. index:
		//lvc.mask |= LVCF_SUBITEM;
		//lvc.iSubItem = index;
		// Testing shows that the following serve to set the column's physical/display position in the
		// heading to iOrder without affecting the specified index.  This concept is very similar to
		// when the user drags and drops a column heading to a new position: it's index doesn't change,
		// only it's displayed position:
		//lvc.mask |= LVCF_ORDER;
		//lvc.iOrder = index + 1;
		if (   -1 == (index = ListView_InsertColumn(control.hwnd, index, &lvc))   )
			_o_return_retval; // Since column could not be inserted, return so that below, sort-now, etc. are not done.
		_f_set_retval_i(index + 1); // +1 to convert the new index to 1-based.
		if (index < lv_attrib.col_count) // Since col is not being appended to the end, make room in the array to insert this column.
			MoveMemory(lv_attrib.col+index+1, lv_attrib.col+index, sizeof(lv_col_type)*(lv_attrib.col_count-index));
			// Above: Shift columns to the right by one.
		lv_attrib.col[index] = col; // Copy temp struct's members to the correct element in the array.
		// The above is done even when index==0 because "col" may contain attributes set via the Options
		// parameter.  Therefore, for code simplicity and rarity of real-world need, no attempt is made
		// to make the following idea work:
		// When index==0, retain the existing attributes due to the unique behavior of inserting a new first
		// column: The new first column inherit's the old column's values (fields), so it seems best to also have it
		// inherit the old column's attributes.
		++lv_attrib.col_count; // New column successfully added.  Must be done only after the MoveMemory() above.
	}

	// Auto-size is done only at this late a stage, in case column was just created above.
	// Note that ListView_SetColumn() apparently does not support LVSCW_AUTOSIZE_USEHEADER for it's "cx" member.
	// do_auto_size contains the actual column width if mode == 'I' and a width was passed by the caller.
	if (do_auto_size && view_mode == LV_VIEW_DETAILS)
		ListView_SetColumnWidth(control.hwnd, index, do_auto_size); // retval was previously set to the more important result above.
	//else v1.0.36.03: Don't attempt to auto-size the columns while the view is not report-view because
	// that causes any subsequent switch to the "list" view to be corrupted (invisible icons and items).

	if (sort_now)
		GuiType::LV_Sort(control, index, false, sort_now_direction);

	_o_return_retval;
}



void GuiControlType::LV_SetImageList(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
	// Caller has ensured that there is at least one incoming parameter:
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0);
	int list_type;
	if (!ParamIndexIsOmitted(1))
		list_type = ParamIndexToInt(1);
	else // Auto-detect large vs. small icons based on the actual icon size in the image list.
	{
		int cx, cy;
		ImageList_GetIconSize(himl, &cx, &cy);
		list_type = (cx > GetSystemMetrics(SM_CXSMICON)) ? LVSIL_NORMAL : LVSIL_SMALL;
	}
	_o_return((size_t)ListView_SetImageList(hwnd, himl, list_type));
}



void GuiControlType::TV_AddModifyDelete(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// TV.Add():
// Returns the HTREEITEM of the item on success, zero on failure.
// Parameters:
//    1: Text/name of item.
//    2: Parent of item.
//    3: Options.
// TV.Modify():
// Returns the HTREEITEM of the item on success (to allow nested calls in script, zero on failure or partial failure.
// Parameters:
//    1: ID of item to modify.
//    2: Options.
//    3: New name.
// Parameters for TV.Delete():
//    1: ID of item to delete (if omitted, all items are deleted).
{
	GuiControlType &control = *this;
	auto mode = BuiltInFunctionID(aID);
	LPTSTR buf = _f_number_buf; // Resolve macro early for maintainability.

	if (mode == FID_TV_Delete)
	{
		// If param #1 is present but is zero, for safety it seems best not to do a delete-all (in case a
		// script bug is so rare that it is never caught until the script is distributed).  Another reason
		// is that a script might do something like TV.Delete(TV.GetSelection()), which would be desired
		// to fail not delete-all if there's ever any way for there to be no selection.
		_o_return(SendMessage(control.hwnd, TVM_DELETEITEM, 0
			, ParamIndexIsOmitted(0) ? NULL : (LPARAM)ParamIndexToInt64(0)));
	}

	// Since above didn't return, this is TV.Add() or TV.Modify().
	TVINSERTSTRUCT tvi; // It contains a TVITEMEX, which is okay even if MSIE pre-4.0 on Win95/NT because those OSes will simply never access the new/bottommost item in the struct.
	bool add_mode = (mode == FID_TV_Add); // For readability & maint.
	HTREEITEM retval;
	
	// Suppress any events raised by the changes made below:
	control.attrib |= GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS;

	LPTSTR options;
	if (add_mode) // TV.Add()
	{
		tvi.hParent = ParamIndexIsOmitted(1) ? NULL : (HTREEITEM)ParamIndexToInt64(1);
		tvi.hInsertAfter = TVI_LAST; // i.e. default is to insert the new item underneath the bottommost sibling.
		options = ParamIndexToOptionalString(2, buf);
		retval = 0; // Set default return value.
	}
	else // TV.Modify()
	{
		// NOTE: Must allow hitem==0 for TV.Modify, at least for the Sort option, because otherwise there would
		// be no way to sort the root-level items.
		tvi.item.hItem = (HTREEITEM)ParamIndexToInt64(0); // Load-time validation has ensured there is a first parameter for TV.Modify().
		// For modify-mode, set default return value to be "success" from this point forward.  Note that
		// in the case of sorting the root-level items, this will set it to zero, but since that almost
		// always succeeds and the script rarely cares whether it succeeds or not, adding code size for that
		// doesn't seem worth it:
		retval = tvi.item.hItem; // Set default return value.
		if (aParamCount < 2) // In one-parameter mode, simply select the item.
		{
			if (!TreeView_SelectItem(control.hwnd, tvi.item.hItem))
				retval = 0; // Override the HTREEITEM default value set above.
			control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
			_o_return((size_t)retval);
		}
		// Otherwise, there's a second parameter (even if it's 0 or "").
		options = ParamIndexToString(1, buf);
	}

	// Set defaults prior to options-parsing, to cover all omitted defaults:
	tvi.item.mask = TVIF_STATE; // TVIF_STATE: The state and stateMask members are valid (all other members are ignored).
	tvi.item.stateMask = 0; // All bits in "state" below are ignored unless the corresponding bit is present here in the mask.
	tvi.item.state = 0;
	// It seems tvi.item.cChildren is typically maintained by the control, though one exception is I_CHILDRENCALLBACK
	// and TVN_GETDISPINFO as mentioned at MSDN.

	DWORD select_flag = 0;
	bool ensure_visible = false, ensure_visible_first = false;

	// Parse list of space-delimited options:
	TCHAR *next_option, *option_end, orig_char;
	bool adding; // Whether this option is being added (+) or removed (-).

	for (next_option = options; *next_option; next_option = omit_leading_whitespace(option_end))
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
		if (   !(option_end = StrChrAny(next_option, _T(" \t")))   )  // Space or tab.
			option_end = next_option + _tcslen(next_option); // Set to position of zero terminator instead.
		if (option_end == next_option)
			continue; // i.e. the string contains a + or - with a space or tab after it, which is intentionally ignored.

		// Temporarily terminate to help eliminate ambiguity for words contained inside other words,
		// such as "Checked" inside of "CheckedGray":
		orig_char = *option_end;
		*option_end = '\0';

		if (!_tcsicmp(next_option, _T("Select"))) // Could further allow "ed" suffix by checking for that inside, but "Selected" is getting long so it doesn't seem something many would want to use.
		{
			// Selection of an item apparently needs to be done via message for the control to update itself
			// properly.  Otherwise, single-select isn't enforced via de-selecting previous item and the newly
			// selected item isn't revealed/shown.  There may be other side-effects.
			if (adding)
				select_flag = TVGN_CARET;
			//else since "de-select" is not a supported action, no need to support "-Select".
			// Furthermore, since a TreeView is by its nature has only one item selected at a time, it seems
			// unnecessary to support Select%VarContainingOneOrZero%.  This is because it seems easier for a
			// script to simply load the Tree then select the desired item afterward.
		}
		else if (!_tcsnicmp(next_option, _T("Vis"), 3))
		{
			// Since this option much more typically used with TV.Modify than TV.Add, the technique of
			// Vis%VarContainingOneOrZero% isn't supported, to reduce code size.
			next_option += 3;
			if (!_tcsicmp(next_option, _T("First"))) // VisFirst
				ensure_visible_first = adding;
			else if (!*next_option)
				ensure_visible = adding;
		}
		else if (!_tcsnicmp(next_option, _T("Bold"), 4))
		{
			next_option += 4;
			if (*next_option && !ATOI(next_option)) // If it's Bold0, invert the mode.
				adding = !adding;
			tvi.item.stateMask |= TVIS_BOLD;
			if (adding)
				tvi.item.state |= TVIS_BOLD;
			//else removing, so the fact that this TVIS flag has just been added to the stateMask above
			// but is absent from item.state should remove this attribute from the item.
		}
		else if (!_tcsnicmp(next_option, _T("Expand"), 6))
		{
			next_option += 6;
			if (*next_option && !ATOI(next_option)) // If it's Expand0, invert the mode to become "collapse".
				adding = !adding;
			if (add_mode)
			{
				if (adding)
				{
					// Don't expand via msg because it won't work: since the item is being newly added
					// now, by definition it doesn't have any children, and testing shows that sending
					// the expand message has no effect, but setting the state bit does:
					tvi.item.stateMask |= TVIS_EXPANDED;
					tvi.item.state |= TVIS_EXPANDED;
					// Since the script is deliberately expanding the item, it seems best not to send the
					// TVN_ITEMEXPANDING/-ED messages because:
					// 1) Sending TVN_ITEMEXPANDED without first sending a TVN_ITEMEXPANDING message might
					//    decrease maintainability, and possibly even produce unwanted side-effects.
					// 2) Code size and performance (avoids generating extra message traffic).
				}
				//else removing, so nothing needs to be done because "collapsed" is the default state
				// of a TV item upon creation.
			}
			else // TV.Modify(): Expand and collapse both require a message to work properly on an existing item.
				// Strangely, this generates a notification sometimes (such as the first time) but not for subsequent
				// expands/collapses of that same item.  Also, TVE_TOGGLE is not currently supported because it seems
				// like it would be too rarely used.
				if (!TreeView_Expand(control.hwnd, tvi.item.hItem, adding ? TVE_EXPAND : TVE_COLLAPSE))
					retval = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
					// It seems that despite what MSDN says, failure is returned when collapsing and item that is
					// already collapsed, but not when expanding an item that is already expanded.  For performance
					// reasons and rarity of script caring, it seems best not to try to adjust/fix this.
		}
		else if (!_tcsnicmp(next_option, _T("Check"), 5))
		{
			// The rationale for not checking for an optional "ed" suffix here and incrementing next_option by 2
			// is that: 1) It would be inconsistent with the lack of support for "selected" (see reason above);
			// 2) Checkboxes in a ListView are fairly rarely used, so code size reduction might be more important.
			next_option += 5;
			if (*next_option && !ATOI(next_option)) // If it's Check0, invert the mode to become "unchecked".
				adding = !adding;
			//else removing, so the fact that this TVIS flag has just been added to the stateMask above
			// but is absent from item.state should remove this attribute from the item.
			tvi.item.stateMask |= TVIS_STATEIMAGEMASK;  // Unlike ListViews, Tree checkmarks can be applied in the same step as creating a Tree item.
			tvi.item.state |= adding ? 0x2000 : 0x1000; // The #1 image is "unchecked" and the #2 is "checked".
		}
		else if (!_tcsnicmp(next_option, _T("Icon"), 4))
		{
			if (adding)
			{
				// To me, having a different icon for when the item is selected seems rarely used.  After all,
				// its obvious the item is selected because it's highlighted (unless it lacks a name?)  So this
				// policy makes things easier for scripts that don't want to distinguish.  If ever it is needed,
				// new options such as IconSel and IconUnsel can be added.
				tvi.item.mask |= TVIF_IMAGE|TVIF_SELECTEDIMAGE;
				tvi.item.iSelectedImage = tvi.item.iImage = ATOI(next_option + 4) - 1;  // -1 to convert to zero-based.
			}
			//else removal of icon currently not supported (see comment above), so do nothing in order
			// to reserve "-Icon" in case a future way can be found to do it.
		}
		else if (!_tcsicmp(next_option, _T("Sort")))
		{
			if (add_mode)
				tvi.hInsertAfter = TVI_SORT; // For simplicity, the value of "adding" is ignored.
			else
				// Somewhat debatable, but it seems best to report failure via the return value even though
				// failure probably only occurs when the item has no children, and the script probably
				// doesn't often care about such failures.  It does result in the loss of the HTREEITEM return
				// value, but even if that call is nested in another, the zero should produce no effect in most cases.
				if (!TreeView_SortChildren(control.hwnd, tvi.item.hItem, FALSE)) // Best default seems no-recurse, since typically this is used after a user edits merely a single item.
					retval = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
		}
		// MUST BE LISTED LAST DUE TO "ELSE IF": Options valid only for TV.Add().
		else if (add_mode && !_tcsicmp(next_option, _T("First")))
		{
			tvi.hInsertAfter = TVI_FIRST; // For simplicity, the value of "adding" is ignored.
		}
		else if (add_mode && IsNumeric(next_option, false, false, false))
		{
			tvi.hInsertAfter = (HTREEITEM)ATOI64(next_option); // ATOI64 vs. ATOU avoids need for extra casting to avoid compiler warning.
		}
		else
		{
			control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
			aResultToken.ValueError(ERR_INVALID_OPTION, next_option);
			*option_end = orig_char; // See comment below.
			return;
		}

		// If the item was not handled by the above, ignore it because it is unknown.
		*option_end = orig_char; // Undo the temporary termination because the caller needs aOptions to be unaltered.
	}

	if (add_mode) // TV.Add()
	{
		tvi.item.pszText = ParamIndexToString(0, buf);
		tvi.item.mask |= TVIF_TEXT;
		tvi.item.hItem = TreeView_InsertItem(control.hwnd, &tvi); // Update tvi.item.hItem for convenience/maint. It's for use in later sections because retval is overridden to be zero for partial failure in modify-mode.
		retval = tvi.item.hItem; // Set return value.
	}
	else // TV.Modify()
	{
		if (!ParamIndexIsOmitted(2)) // An explicit empty string is allowed, which sets it to a blank value.  By contrast, if the param is omitted, the name is left changed.
		{
			tvi.item.pszText = ParamIndexToString(2, buf); // Reuse buf now that options (above) is done with it.
			tvi.item.mask |= TVIF_TEXT;
		}
		//else name/text parameter has been omitted, so don't change the item's name.
		if (tvi.item.mask != LVIF_STATE || tvi.item.stateMask) // An item's property or one of the state bits needs changing.
			if (!TreeView_SetItem(control.hwnd, &tvi.itemex))
				retval = 0; // Indicate partial failure by overriding the HTREEITEM return value set earlier.
	}

	if (ensure_visible) // Seems best to do this one prior to "select" below.
		SendMessage(control.hwnd, TVM_ENSUREVISIBLE, 0, (LPARAM)tvi.item.hItem); // Return value is ignored in this case, since its definition seems a little weird.
	if (ensure_visible_first) // Seems best to do this one prior to "select" below.
		TreeView_Select(control.hwnd, tvi.item.hItem, TVGN_FIRSTVISIBLE); // Return value is also ignored due to rarity, code size, and because most people wouldn't care about a failure even if for some reason it failed.
	if (select_flag)
		if (!TreeView_Select(control.hwnd, tvi.item.hItem, select_flag) && !add_mode) // Relies on short-circuit boolean order.
			retval = 0; // When not in add-mode, indicate partial failure by overriding the return value set earlier (add-mode should always return the new item's ID).

	control.attrib &= ~GUI_CONTROL_ATTRIB_SUPPRESS_EVENTS; // Re-enable events.
	_o_return((size_t)retval);
}



HTREEITEM GetNextTreeItem(HWND aTreeHwnd, HTREEITEM aItem)
// Helper function for others below.
// If aItem is NULL, caller wants topmost ROOT item returned.
// Otherwise, the next child, sibling, or parent's sibling is returned in a manner that allows the caller
// to traverse every item in the tree easily.
{
	if (!aItem)
		return TreeView_GetRoot(aTreeHwnd);
	// Otherwise, do depth-first recursion.  Must be done in the following order to allow full traversal:
	// Children first.
	// Then siblings.
	// Then parent's sibling(s).
	HTREEITEM hitem;
	if (hitem = TreeView_GetChild(aTreeHwnd, aItem))
		return hitem;
	if (hitem = TreeView_GetNextSibling(aTreeHwnd, aItem))
		return hitem;
	// The last stage is trickier than the above: parent's next sibling, or if none, its parent's parent's sibling, etc.
	for (HTREEITEM hparent = aItem;;)
	{
		if (   !(hparent = TreeView_GetParent(aTreeHwnd, hparent))   ) // No parent, so this is a root-level item.
			return NULL; // There is no next item.
		// Now it's known there is a parent.  It's not necessary to check that parent's children because that
		// would have been done by a prior iteration in the script.
		if (hitem = TreeView_GetNextSibling(aTreeHwnd, hparent))
			return hitem;
		// Otherwise, parent has no sibling, but does its parent (and so on)? Continue looping to find out.
	}
}



void GuiControlType::TV_GetRelatedItem(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// TV.GetParent/Child/Selection/Next/Prev(hitem):
// The above all return the HTREEITEM (or 0 on failure).
// When TV.GetNext's second parameter is present, the search scope expands to include not just siblings,
// but also children and parents, which allows a tree to be traversed from top to bottom without the script
// having to do something fancy.
{
	GuiControlType &control = *this;
	HWND control_hwnd = control.hwnd;

	HTREEITEM hitem = (HTREEITEM)ParamIndexToOptionalIntPtr(0, NULL);

	if (ParamIndexIsOmitted(1))
	{
		WPARAM flag;
		switch (aID)
		{
		case FID_TV_GetSelection: flag = TVGN_CARET; break; // TVGN_CARET is the focused item.
		case FID_TV_GetParent: flag = TVGN_PARENT; break;
		case FID_TV_GetChild: flag = TVGN_CHILD; break;
		case FID_TV_GetPrev: flag = TVGN_PREVIOUS; break;
		case FID_TV_GetNext: flag = !hitem ? TVGN_ROOT : TVGN_NEXT; break; // TV_GetNext(no-parameters) yields very first item in Tree (TVGN_ROOT).
		// Above: It seems best to treat hitem==0 as "get root", even though it sacrifices some error detection,
		// because not doing so would be inconsistent with the fact that TV.GetNext(0, "Full") does get the root
		// (which needs to be retained to make script loops to traverse entire tree easier).
		//case FID_TV_GetCount:
		default:
			// There's a known bug mentioned at MSDN that a TreeView might report a negative count when there
			// are more than 32767 items in it (though of course HTREEITEM values are never negative since they're
			// defined as unsigned pseudo-addresses).  But apparently, that bug only applies to Visual Basic and/or
			// older OSes, because testing shows that SendMessage(TVM_GETCOUNT) returns 32800+ when there are more
			// than 32767 items in the tree, even without casting to unsigned.  So I'm not sure exactly what the
			// story is with this, so for now just casting to UINT rather than something less future-proof like WORD:
			// Older note, apparently unneeded at least on XP SP2: Cast to WORD to convert -1 through -32768 to the
			// positive counterparts.
			_o_return((UINT)SendMessage(control_hwnd, TVM_GETCOUNT, 0, 0));
		}
		// Apparently there's no direct call to get the topmost ancestor of an item, presumably because it's rarely
		// needed.  Therefore, no such mode is provide here yet (the syntax TV.GetParent(hitem, true) could be supported
		// if it's ever needed).
		_o_return(SendMessage(control_hwnd, TVM_GETNEXTITEM, flag, (LPARAM)hitem));
	}

	// Since above didn't return, this TV.GetNext's 2-parameter mode, which has an expanded scope that includes
	// not just siblings, but also children and parents.  This allows a tree to be traversed from top to bottom
	// without the script having to do something fancy.
	TCHAR first_char_upper = ctoupper(*omit_leading_whitespace(ParamIndexToString(1, _f_number_buf))); // Resolve parameter #2.
	bool search_checkmark;
	if (first_char_upper == 'C')
		search_checkmark = true;
	else if (first_char_upper == 'F')
		search_checkmark = false;
	else // Reserve other option letters/words for future use by being somewhat strict.
		_o_throw_param(1);

	// When an actual item was specified, search begins at the item *after* it.  Otherwise (when NULL):
	// It's a special mode that always considers the root node first.  Otherwise, there would be no way
	// to start the search at the very first item in the tree to find out whether it's checked or not.
	hitem = GetNextTreeItem(control_hwnd, hitem); // Handles the comment above.
	if (!search_checkmark) // Simple tree traversal, so just return the next item (if any).
		_o_return((size_t)hitem); // OK if NULL.

	// Otherwise, search for the next item having a checkmark. For performance, it seems best to assume that
	// the control has the checkbox style (the script would realistically never call it otherwise, so the
	// control's style isn't checked.
	for (; hitem; hitem = GetNextTreeItem(control_hwnd, hitem))
		if (TreeView_GetCheckState(control_hwnd, hitem) == 1) // 0 means unchecked, -1 means "no checkbox image".
			_o_return((size_t)hitem); // OK if NULL.
	// Since above didn't return, the entire tree starting at the specified item has been searched,
	// with no match found.
	_o_return(0);
}



void GuiControlType::TV_Get(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// TV.Get()
// Returns: Varies depending on param #2.
// Parameters:
//    1: HTREEITEM.
//    2: Name of attribute to get.
// TV.GetText()
// Returns: Text on success.
// Throws on failure.
// Parameters:
//    1: HTREEITEM.
{
	bool get_text = aID == FID_TV_GetText;

	HWND control_hwnd = hwnd;

	if (!get_text)
	{
		// Loadtime validation has ensured that param #1 and #2 are present for all these cases.
		HTREEITEM hitem = (HTREEITEM)ParamIndexToInt64(0);
		UINT state_mask;
		switch (ctoupper(*omit_leading_whitespace(ParamIndexToString(1, _f_number_buf))))
		{
		case 'E': state_mask = TVIS_EXPANDED; break; // Expanded
		case 'C': state_mask = TVIS_STATEIMAGEMASK; break; // Checked
		case 'B': state_mask = TVIS_BOLD; break; // Bold
		//case 'S' for "Selected" is not provided because TV.GetSelection() seems to cover that well enough.
		//case 'P' for "is item a parent?" is not provided because TV.GetChild() seems to cover that well enough.
		// (though it's possible that retrieving TVITEM's cChildren would perform a little better).
		}
		// Below seems to be need a bit-AND with state_mask to work properly, at least on XP SP2.  Otherwise,
		// extra bits are present such as 0x2002 for "expanded" when it's supposed to be either 0x00 or 0x20.
		UINT result = state_mask & (UINT)SendMessage(control_hwnd, TVM_GETITEMSTATE, (WPARAM)hitem, state_mask);
		if (state_mask == TVIS_STATEIMAGEMASK)
		{
			if (result != 0x2000) // It doesn't have a checkmark state image.
				hitem = 0;
		}
		else // For all others, anything non-zero means the flag is present.
            if (!result) // Flag not present.
				hitem = 0;
		_o_return((size_t)hitem);
	}

	TCHAR text_buf[LV_TEXT_BUF_SIZE]; // i.e. uses same size as ListView.
	TVITEM tvi;
	tvi.hItem = (HTREEITEM)ParamIndexToInt64(0);
	tvi.mask = TVIF_TEXT;
	tvi.pszText = text_buf;
	tvi.cchTextMax = LV_TEXT_BUF_SIZE - 1; // -1 because of nagging doubt about size vs. length. Some MSDN examples subtract one), such as TabCtrl_GetItem()'s cchTextMax.

	if (SendMessage(control_hwnd, TVM_GETITEM, 0, (LPARAM)&tvi))
	{
		// Must use tvi.pszText vs. text_buf because MSDN says: "Applications should not assume that the text will
		// necessarily be placed in the specified buffer. The control may instead change the pszText member
		// of the structure to point to the new text rather than place it in the buffer."
		_o_return(tvi.pszText);
	}
	else
	{
		// On failure, it seems best to throw an exception.
		_o_throw(ERR_FAILED);
	}
}



void GuiControlType::TV_SetImageList(ResultToken &aResultToken, int aID, int aFlags, ExprTokenType *aParam[], int aParamCount)
// Returns (MSDN): "handle to the image list previously associated with the control if successful; NULL otherwise."
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
// 2: Optional: Type of list.
{
	// Caller has ensured that there is at least one incoming parameter:
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0);
	int list_type;
	list_type = ParamIndexToOptionalInt(1, TVSIL_NORMAL);
	_o_return((size_t)TreeView_SetImageList(hwnd, himl, list_type));
}



BIF_DECL(BIF_IL_Create)
// Returns: Handle to the new image list, or 0 on failure.
// Parameters:
// 1: Initial image count (ImageList_Create() ignores values <=0, so no need for error checking).
// 2: Grow count (testing shows it can grow multiple times, even when this is set <=0, so it's apparently only a performance aid)
// 3: Width of each image (overloaded to mean small icon size when omitted or false, large icon size otherwise).
// 4: Future: Height of each image [if this param is present and >0, it would mean param 3 is not being used in its TRUE/FALSE mode)
// 5: Future: Flags/Color depth
{
	// The following old comment makes no sense because large icons are only used if param3 is NON-ZERO,
	// and there was never a distinction between passing zero and omitting the param:
	// So that param3 can be reserved as a future "specified width" param, to go along with "specified height"
	// after it, only when the parameter is both present and numerically zero are large icons used.  Otherwise,
	// small icons are used.
	int param3 = ParamIndexToOptionalBOOL(2, FALSE);
	_f_return_i((size_t)ImageList_Create(GetSystemMetrics(param3 ? SM_CXICON : SM_CXSMICON)
		, GetSystemMetrics(param3 ? SM_CYICON : SM_CYSMICON)
		, ILC_MASK | ILC_COLOR32  // ILC_COLOR32 or at least something higher than ILC_COLOR is necessary to support true-color icons.
		, ParamIndexToOptionalInt(0, 2)    // cInitial. 2 seems a better default than one, since it might be common to have only two icons in the list.
		, ParamIndexToOptionalInt(1, 5)));  // cGrow.  Somewhat arbitrary default.
}



BIF_DECL(BIF_IL_Destroy)
// Returns: 1 on success and 0 on failure.
// Parameters:
// 1: HIMAGELIST obtained from somewhere such as IL_Create().
{
	// Load-time validation has ensured there is at least one parameter.
	// Returns nonzero if successful, or zero otherwise, so force it to conform to TRUE/FALSE for
	// better consistency with other functions:
	_f_return_i(ImageList_Destroy((HIMAGELIST)ParamIndexToInt64(0)) ? 1 : 0);
}



BIF_DECL(BIF_IL_Add)
// Returns: the one-based index of the newly added icon, or zero on failure.
// Parameters:
// 1: HIMAGELIST: Handle of an existing ImageList.
// 2: Filename from which to load the icon or bitmap.
// 3: Icon number within the filename (or mask color for non-icon images).
// 4: The mere presence of this parameter indicates that param #3 is mask RGB-color vs. icon number.
//    This param's value should be "true" to resize the image to fit the image-list's size or false
//    to divide up the image into a series of separate images based on its width.
//    (this parameter could be overloaded to be the filename containing the mask image, or perhaps an HBITMAP
//    provided directly by the script)
// 5: Future: can be the scaling height to go along with an overload of #4 as the width.  However,
//    since all images in an image list are of the same size, the use of this would be limited to
//    only those times when the imagelist would be scaled prior to dividing it into separate images.
// The parameters above (at least #4) can be overloaded in the future calling ImageList_GetImageInfo() to determine
// whether the imagelist has a mask.
{
	HIMAGELIST himl = (HIMAGELIST)ParamIndexToInt64(0); // Load-time validation has ensured there is a first parameter.
	if (!himl)
		_f_throw_param(0);

	int param3 = ParamIndexToOptionalInt(2, 0);
	int icon_number, width = 0, height = 0; // Zero width/height causes image to be loaded at its actual width/height.
	if (!ParamIndexIsOmitted(3)) // Presence of fourth parameter switches mode to be "load a non-icon image".
	{
		icon_number = 0; // Zero means "load icon or bitmap (doesn't matter)".
		if (ParamIndexToBOOL(3)) // A value of True indicates that the image should be scaled to fit the imagelist's image size.
			ImageList_GetIconSize(himl, &width, &height); // Determine the width/height to which it should be scaled.
		//else retain defaults of zero for width/height, which loads the image at actual size, which in turn
		// lets ImageList_AddMasked() divide it up into separate images based on its width.
	}
	else
	{
		icon_number = param3; // LoadPicture() properly handles any wrong/negative value that might be here.
		ImageList_GetIconSize(himl, &width, &height); // L19: Determine the width/height of images in the image list to ensure icons are loaded at the correct size.
	}

	int image_type;
	HBITMAP hbitmap = LoadPicture(ParamIndexToString(1, _f_number_buf) // Caller has ensured there are at least two parameters.
		, width, height, image_type, icon_number, false); // Defaulting to "false" for "use GDIplus" provides more consistent appearance across multiple OSes.
	if (!hbitmap)
		_f_return_i(0);

	int index;
	if (image_type == IMAGE_BITMAP) // In this mode, param3 is always assumed to be an RGB color.
	{
		// Return the index of the new image or 0 on failure.
		index = ImageList_AddMasked(himl, hbitmap, rgb_to_bgr((int)param3)) + 1; // +1 to convert to one-based.
		DeleteObject(hbitmap);
	}
	else // ICON or CURSOR.
	{
		// Return the index of the new image or 0 on failure.
		index = ImageList_AddIcon(himl, (HICON)hbitmap) + 1; // +1 to convert to one-based.
		DestroyIcon((HICON)hbitmap); // Works on cursors too.  See notes in LoadPicture().
	}
	_f_return_i(index);
}



////////////////////
// Misc Functions //
////////////////////

BIF_DECL(BIF_LoadPicture)
{
	// h := LoadPicture(filename [, options, ByRef image_type])
	LPTSTR filename = ParamIndexToString(0, aResultToken.buf);
	LPTSTR options = ParamIndexToOptionalString(1);
	Var *image_type_var = ParamIndexToOutputVar(2);

	int width = -1;
	int height = -1;
	int icon_number = 0;
	bool use_gdi_plus = false;

	for (LPTSTR cp = options; cp; cp = StrChrAny(cp, _T(" \t")))
	{
		cp = omit_leading_whitespace(cp);
		if (tolower(*cp) == 'w')
			width = ATOI(cp + 1);
		else if (tolower(*cp) == 'h')
			height = ATOI(cp + 1);
		else if (!_tcsnicmp(cp, _T("Icon"), 4))
			icon_number = ATOI(cp + 4);
		else if (!_tcsnicmp(cp, _T("GDI+"), 4))
			// GDI+ or GDI+1 to enable, GDI+0 to disable.
			use_gdi_plus = cp[4] != '0';
	}

	if (width == -1 && height == -1)
		width = 0;

	int image_type;
	HBITMAP hbm = LoadPicture(filename, width, height, image_type, icon_number, use_gdi_plus);
	if (image_type_var)
		image_type_var->Assign(image_type);
	else if (image_type != IMAGE_BITMAP && hbm)
		// Always return a bitmap when the ImageType output var is omitted.
		hbm = IconToBitmap32((HICON)hbm, true); // Also works for cursors.
	aResultToken.value_int64 = (__int64)(UINT_PTR)hbm;
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
