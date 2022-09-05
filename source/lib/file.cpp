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
#include "application.h"
#include "TextIO.h"
#include "script_func_impl.h"
#include "abi.h"



typedef BOOL (* FilePatternCallback)(LPCTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);

struct FilePatternStruct
{
	TCHAR path[T_MAX_PATH]; // Directory and naked filename or pattern.
	TCHAR pattern[MAX_PATH]; // Naked filename or pattern.
	size_t dir_length, pattern_length;
	FilePatternCallback aCallback;
	void *aCallbackData;
	FileLoopModeType aOperateOnFolders;
	bool aDoRecurse;
	int failure_count;
};

static void FilePatternApply(FilePatternStruct &fps);

static FResult FilePatternApply(LPCTSTR aFilePattern, FileLoopModeType aOperateOnFolders
	, bool aDoRecurse, FilePatternCallback aCallback, void *aCallbackData);



static bool FileCreateDirRecursive(LPTSTR aDirSpec);

// As of 2019-09-29, noinline reduces code size by over 20KB on VC++ 2019.
// Prior to merging Util_CreateDir with this, it wasn't inlined.
DECLSPEC_NOINLINE
bool FileCreateDir(LPCTSTR aDirSpec)
{
	if (!aDirSpec || !*aDirSpec)
	{
		SetLastError(ERROR_INVALID_PARAMETER);
		return false;
	}

	// Allocate a modifiable buffer to be used by recursive calls (supports long paths).
	//auto buf = (LPTSTR)_alloca((last_backslash - aDirSpec + 1) * sizeof(TCHAR));
	TCHAR buf[T_MAX_PATH];
	auto len = _tcslen(aDirSpec);
	if (len >= T_MAX_PATH)
	{
		SetLastError(ERROR_BUFFER_OVERFLOW);
		return false;
	}
	tmemcpy(buf, aDirSpec, len + 1);

	return FileCreateDirRecursive(buf);
}

static bool FileCreateDirRecursive(LPTSTR aDirSpec)
{
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
		&& last_backslash[-1] != ':' // v1.1.31.00: Don't attempt FileCreateDir("C:") since that's equivalent to either "C:\" or the working directory (which already exists), or FileCreateDir("\\?\C:") since it always fails.
		&& last_backslash[1]) // Skip the recursive call if it's just a trailing backslash.
	{
		*last_backslash = '\0'; // Temporarily terminate for parent directory.
		auto exists = FileCreateDirRecursive(aDirSpec); // Recursively create all needed ancestor directories.
		*last_backslash = '\\'; // Undo temporary termination.
		if (!exists)
			return exists;
	}

	// The above has recursively created all parent directories of aDirSpec if needed.
	// Now we can create aDirSpec.
	return CreateDirectory(aDirSpec, NULL);
}



static FResult ConvertFileOptions(LPCTSTR aOptions, UINT &codepage, bool &translate_crlf_to_lf, unsigned __int64 *pmax_bytes_to_load)
{
	if (aOptions)
	for (LPCTSTR next, cp = aOptions; cp && *(cp = omit_leading_whitespace(cp)); cp = next)
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
					return FValueError(ERR_INVALID_OPTION, cp);
			}
			break;
		} // switch()
	} // for()
	return OK;
}



bif_impl FResult FileRead(LPCTSTR aFilespec, LPCTSTR aOptions, ResultToken &aResultToken)
{
	g->LastError = 0; // Set default for successful early return or non-Win32 errors.
	if (!*aFilespec)
		return FR_E_ARG(0); // Seems more helpful than throwing OSError(3).

	const DWORD DWORD_MAX = ~0;

	// Set default options:
	bool translate_crlf_to_lf = false;
	unsigned __int64 max_bytes_to_load = ULLONG_MAX; // By default, fail if the file is too large.  See comments near bytes_to_read below.
	UINT codepage = g->Encoding;

	auto fr = ConvertFileOptions(aOptions, codepage, translate_crlf_to_lf, &max_bytes_to_load);
	if (fr != OK)
		return fr; // It already displayed the error.

	// It seems more flexible to allow other processes to read and write the file while we're reading it.
	// For example, this allows the file to be appended to during the read operation, which could be
	// desirable, especially it's a very large log file that would take a long time to read.
	// MSDN: "To enable other processes to share the object while your process has it open, use a combination
	// of one or more of [FILE_SHARE_READ, FILE_SHARE_WRITE]."
	HANDLE hfile = CreateFile(aFilespec, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, NULL, OPEN_EXISTING
		, FILE_FLAG_SEQUENTIAL_SCAN, NULL); // MSDN says that FILE_FLAG_SEQUENTIAL_SCAN will often improve performance
	if (hfile == INVALID_HANDLE_VALUE)      // in cases like these (and it seems best even if max_bytes_to_load was specified).
	{
		g->LastError = GetLastError();
		return FR_E_WIN32(g->LastError);
	}

	unsigned __int64 bytes_to_read = GetFileSize64(hfile);
	if (bytes_to_read == ULLONG_MAX) // GetFileSize64() failed.
	{
		g->LastError = GetLastError();
		CloseHandle(hfile);
		return FR_E_WIN32(g->LastError);
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
		return FR_E_OUTOFMEM; // Using this instead of "File too large." to reduce code size, since this condition is very rare (and malloc succeeding would be even rarer).
	}

	if (!bytes_to_read && codepage != -1) // In RAW mode, return a Buffer even if the file has zero bytes.
	{
		g->LastError = GetLastError();
		CloseHandle(hfile);
		return OK; // Indicate success (a zero-length file results in an empty string).
	}

	LPBYTE output_buf = (LPBYTE)malloc(size_t(bytes_to_read + (bytes_to_read & 1) + sizeof(wchar_t)));
	if (!output_buf)
	{
		CloseHandle(hfile);
		return FR_E_OUTOFMEM;
	}

	DWORD bytes_actually_read;
	BOOL result = ReadFile(hfile, output_buf, (DWORD)bytes_to_read, &bytes_actually_read, NULL);
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
							return aResultToken.Exited() ? FR_FAIL : FR_ABORTED;
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

	if (!result)
	{
		g->LastError = GetLastError();
		return FR_E_WIN32(g->LastError);
	}
	return OK;
}



bif_impl FResult FileAppend(ExprTokenType &aValue, LPCTSTR aFilespec, LPCTSTR aOptions)
{
	g->LastError = 0; // Set default for successful early return or non-Win32 errors.

	size_t aBuf_length;
	TCHAR aBuf_buf[MAX_NUMBER_SIZE];
	LPCTSTR aBuf = TokenToString(aValue, aBuf_buf, &aBuf_length);
	
	IObject *aBuf_obj = TokenToObject(aValue); // Allow a Buffer-like object.
	if (aBuf_obj)
	{
		size_t ptr;
		FuncResult rt;
		GetBufferObjectPtr(rt, aBuf_obj, ptr, aBuf_length);
		ASSERT(rt.symbol == SYM_INTEGER && !rt.mem_to_free);
		if (rt.Exited())
			return FR_FAIL;
		aBuf = (LPTSTR)ptr;
	}
	else
		aBuf_length *= sizeof(TCHAR); // Convert to byte count.

	// The below is avoided because want to allow "nothing" to be written to a file in case the
	// user is doing this to reset it's timestamp (or create an empty file).
	//if (!aBuf || !*aBuf)
	//	return OK;

	// Use the read-file loop's current item if filename was explicitly left blank (i.e. not just
	// a reference to a variable that's blank):
	LoopReadFileStruct *aCurrentReadFile = aFilespec ? NULL : g->mLoopReadFile;
	if (aCurrentReadFile)
		aFilespec = aCurrentReadFile->mWriteFileName;
	if (!aFilespec || !*aFilespec) // Nothing to write to.
		return FR_E_ARG(1);

	TextStream *ts = aCurrentReadFile ? aCurrentReadFile->mWriteFile : NULL;
	bool file_was_already_open = ts;

#ifdef CONFIG_DEBUGGER
	if (*aFilespec == '*' && !aFilespec[1] && !aBuf_obj && g_Debugger.OutputStdOut(aBuf))
	{
		// StdOut has been redirected to the debugger, and this "FileAppend" call has been
		// fully handled by the call above, so just return.
		return OK;
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
		auto fr = ConvertFileOptions(aOptions, codepage, translate_crlf_to_lf, nullptr);
		if (fr != OK)
			return fr;

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
			g->LastError = GetLastError();
			delete ts; // Must be deleted explicitly!
			return FR_E_WIN32(g->LastError);
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
		g->LastError = GetLastError();
	}
	//else: aBuf is empty; we've already succeeded in creating the file and have nothing further to do.

	if (!aCurrentReadFile)
		delete ts;
	// else it's the caller's responsibility, or it's caller's, to close it.
	
	return result == 0 ? FR_E_WIN32(g->LastError) : OK;
}



BOOL FileDeleteCallback(LPCTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData)
{
	return DeleteFile(aFilename);
}

bif_impl FResult FileDelete(LPCTSTR aFilePattern)
{
	if (!*aFilePattern)
		return FR_E_ARG(0);

	// The no-wildcard case could be handled via FilePatternApply(), but handling it this
	// way ensures deleting a non-existent path without wildcards is considered a failure:
	if (!StrChrAny(aFilePattern, _T("?*"))) // No wildcards; just a plain path/filename.
	{
		if (!DeleteFile(aFilePattern))
		{
			g->LastError = GetLastError();
			return FR_E_WIN32(g->LastError);
		}
	}

	// Otherwise aFilePattern contains wildcards, so we'll search for all matches and delete them.
	return FilePatternApply(aFilePattern, FILE_LOOP_FILES_ONLY, false, FileDeleteCallback, NULL);
}



static bool FileInstallExtract(LPCTSTR aSource, LPCTSTR aDest, bool aOverwrite)
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
static bool FileInstallCopy(LPCTSTR aSource, LPCTSTR aDest, bool aOverwrite)
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

bif_impl FResult FileInstall(LPCTSTR aSource, LPCTSTR aDest, int *aFlag)
{
	bool success;
	bool allow_overwrite = (aFlag && *aFlag == 1);
#ifndef AUTOHOTKEYSC
	if (g_script.mKind != Script::ScriptKindResource)
		success = FileInstallCopy(aSource, aDest, allow_overwrite);
	else
#endif
		success = FileInstallExtract(aSource, aDest, allow_overwrite);
	return success ? OK : FR_E_FAILED;
}



static FResult FileCopyOrMove(LPCTSTR aSource, LPCTSTR aDest, int *aFlag, bool aMove)
{
	if (!*aSource) // v2: Empty Source is likely to be a mistake.
		return FR_E_ARG(0);
	if (!*aDest) // Fix for v1.1.34.03: Previous behaviour was a Critical Error.
		return FR_E_ARG(1);
	int error_count = Line::Util_CopyFile(aSource, aDest, aFlag && *aFlag == 1, aMove
		, g->LastError);
	return error_count ? FR_THROW_INT(error_count) : OK;
}

bif_impl FResult FileCopy(LPCTSTR aSource, LPCTSTR aDest, int *aFlag)
{
	return FileCopyOrMove(aSource, aDest, aFlag, false);
}

bif_impl FResult FileMove(LPCTSTR aSource, LPCTSTR aDest, int *aFlag)
{
	return FileCopyOrMove(aSource, aDest, aFlag, true);
}



bif_impl FResult FileGetAttrib(LPCTSTR aPath, StrRet &aRetVal)
{
	g->LastError = 0; // Set default for successful return or non-Win32 errors.

	if (!aPath && g->mLoopFile)
		aPath = g->mLoopFile->cFileName;
	if (!aPath || !*aPath)
		return FR_E_ARG(0);

	DWORD attr = GetFileAttributes(aPath);
	if (attr == 0xFFFFFFFF)  // Failure, probably because file doesn't exist.
	{
		g->LastError = GetLastError();
		return FR_E_WIN32(g->LastError);
	}

	aRetVal.SetStatic(FileAttribToStr(aRetVal.CallerBuf(), attr));
	return OK;
}



BOOL FileSetAttribCallback(LPCTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetAttribData
{
	DWORD and_mask, xor_mask;
};

bif_impl FResult FileSetAttrib(LPCTSTR aAttributes, LPCTSTR aFilePattern, LPCTSTR aMode)
{
	if (!aFilePattern && g->mLoopFile)
		aFilePattern = g->mLoopFile->file_path; // Default to the current file-loop's file.
	if (!aFilePattern || !*aFilePattern)
		return FR_E_ARG(1);
	FileLoopModeType mode = Line::ConvertLoopMode(aMode);
	if (mode == FILE_LOOP_INVALID)
		return FR_E_ARG(2);
	FileLoopModeType aOperateOnFolders = mode & ~FILE_LOOP_RECURSE;
	bool aDoRecurse = mode & FILE_LOOP_RECURSE;

	// Convert the attribute string to three bit-masks: add, remove and toggle.
	FileSetAttribData attrib;
	DWORD mask;
	int op = 0;
	attrib.and_mask = 0xFFFFFFFF; // Set default: keep all bits.
	attrib.xor_mask = 0; // Set default: affect none.
	for (auto cp = aAttributes; *cp; ++cp)
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
			return FR_E_ARG(0);
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
	return FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetAttribCallback, &attrib);
}

BOOL FileSetAttribCallback(LPCTSTR file_path, WIN32_FIND_DATA &current_file, void *aCallbackData)
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



static FResult FilePatternApply(LPCTSTR aFilePattern, FileLoopModeType aOperateOnFolders
	, bool aDoRecurse, FilePatternCallback aCallback, void *aCallbackData)
{
	if (!*aFilePattern)
	{
		g->LastError = ERROR_INVALID_PARAMETER;
		return FR_E_WIN32(g->LastError);
	}

	if (aOperateOnFolders == FILE_LOOP_INVALID) // In case runtime dereference of a var was an invalid value.
		aOperateOnFolders = FILE_LOOP_FILES_ONLY;  // Set default.
	g->LastError = 0; // Set default. Overridden only when a failure occurs.

	FilePatternStruct fps;

	auto last_backslash = _tcsrchr(aFilePattern, '\\');
	if (last_backslash)
		fps.dir_length = last_backslash - aFilePattern + 1; // Include the slash.
	else // Use current working directory, e.g. if user specified only *.*
		fps.dir_length = 0;
	fps.pattern_length = _tcslen(aFilePattern + fps.dir_length);
	
	// Testing shows that the ANSI version of FindFirstFile() will not accept a path+pattern longer
	// than 259, even if the pattern would match files whose names are short enough to be legal.
	if (fps.dir_length + fps.pattern_length >= _countof(fps.path)
		|| fps.pattern_length >= _countof(fps.pattern))
	{
		g->LastError = ERROR_BUFFER_OVERFLOW;
		return FR_E_WIN32(g->LastError);
	}

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

	return fps.failure_count ? FR_THROW_INT(fps.failure_count) : OK;
}



static void FilePatternApply(FilePatternStruct &fps)
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



bif_impl FResult FileGetTime(LPCTSTR aPath, LPCTSTR aWhichTime, StrRet &aRetVal)
{
	g->LastError = 0; // Set default for successful return or non-Win32 errors.

	if (!aPath && g->mLoopFile)
		aPath = g->mLoopFile->cFileName;
	if (!aPath || !*aPath)
		return FR_E_ARG(0);

	FILETIME *which_time;
	WIN32_FIND_DATA found_file;
	switch (aWhichTime && *aWhichTime ? ctoupper(*aWhichTime) : 'M')
	{
	case 'C': which_time = &found_file.ftCreationTime; break;
	case 'A': which_time = &found_file.ftLastAccessTime; break;
	case 'M': which_time = &found_file.ftLastWriteTime; break;
	default: return FR_E_ARG(1);
	}

	// Don't use CreateFile() & GetFileTime() since they will fail to work on a file that's in use.
	// Research indicates that this method has no disadvantages compared to the other method.
	HANDLE file_search = FindFirstFile(aPath, &found_file);
	if (file_search == INVALID_HANDLE_VALUE)
	{
		g->LastError = GetLastError();
		return FR_E_WIN32(g->LastError);
	}
	FindClose(file_search);

	FILETIME local_file_time;
	FileTimeToLocalFileTime(which_time, &local_file_time);
	aRetVal.SetStatic(FileTimeToYYYYMMDD(aRetVal.CallerBuf(), local_file_time));
	return OK;
}



BOOL FileSetTimeCallback(LPCTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData);
struct FileSetTimeData
{
	FILETIME Time;
	TCHAR WhichTime;
};

bif_impl FResult FileSetTime(LPCTSTR aYYYYMMDD, LPCTSTR aFilePattern, LPCTSTR aWhichTime, LPCTSTR aMode)
{
	if (!aFilePattern && g->mLoopFile)
		aFilePattern = g->mLoopFile->file_path; // Default to the current file-loop's file.
	if (!aFilePattern || !*aFilePattern)
		return FR_E_ARG(1);

	FileLoopModeType mode = Line::ConvertLoopMode(aMode);
	if (mode == FILE_LOOP_INVALID)
		return FR_E_ARG(3);
	FileLoopModeType aOperateOnFolders = mode & ~FILE_LOOP_RECURSE;
	bool aDoRecurse = mode & FILE_LOOP_RECURSE;
	
	FileSetTimeData callbackData;
	switch (callbackData.WhichTime = aWhichTime ? ctoupper(*aWhichTime) : 0)
	{
	case 'M': case 'C': case 'A': case '\0': break;
	default: return FR_E_ARG(2);
	}

	FILETIME ft;
	if (aYYYYMMDD && *aYYYYMMDD)
	{
		if (   !YYYYMMDDToFileTime(aYYYYMMDD, ft)  // Convert the arg into the time struct as local (non-UTC) time.
			|| !LocalFileTimeToFileTime(&ft, &callbackData.Time)   )  // Convert from local to UTC.
		{
			// Invalid parameters are the only likely cause of this condition.
			return FR_E_ARG(0);
		}
	}
	else // User wants to use the current time (i.e. now) as the new timestamp.
		GetSystemTimeAsFileTime(&callbackData.Time);

	return FilePatternApply(aFilePattern, aOperateOnFolders, aDoRecurse, FileSetTimeCallback, &callbackData);
}

BOOL FileSetTimeCallback(LPCTSTR aFilename, WIN32_FIND_DATA &aFile, void *aCallbackData)
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



bif_impl FResult FileGetSize(LPCTSTR aPath, LPCTSTR aUnits, __int64 *aRetVal)
{
	g->LastError = 0; // Set default for successful return or non-Win32 errors.

	if (!aPath && g->mLoopFile)
		aPath = g->mLoopFile->cFileName;
	if (!aPath || !*aPath)
		return FR_E_ARG(0);

	BOOL got_file_size = false;
	__int64 size;

	// Try CreateFile() and GetFileSizeEx() first, since they can be more accurate. 
	// See "Why is the file size reported incorrectly for files that are still being written to?"
	// http://blogs.msdn.com/b/oldnewthing/archive/2011/12/26/10251026.aspx
	HANDLE hfile = CreateFile(aPath, FILE_READ_ATTRIBUTES, FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
		, NULL, OPEN_EXISTING, 0, NULL);
	if (hfile != INVALID_HANDLE_VALUE)
	{
		got_file_size = GetFileSizeEx(hfile, (PLARGE_INTEGER)&size);
		CloseHandle(hfile);
	}

	if (!got_file_size)
	{
		WIN32_FIND_DATA found_file;
		HANDLE file_search = FindFirstFile(aPath, &found_file);
		if (file_search == INVALID_HANDLE_VALUE)
		{
			g->LastError = GetLastError();
			return FR_E_WIN32(g->LastError);
		}
		FindClose(file_search);
		size = ((__int64)found_file.nFileSizeHigh << 32) | found_file.nFileSizeLow;
	}

	switch (aUnits && *aUnits ? ctoupper(*aUnits) : 'B')
	{
	case 'K': size /= 1024; break; // KB
	case 'M': size /= (1024 * 1024); break; // MB
	case 'B': break; // Bytes
	default: return FR_E_ARG(1);
	}

	g->LastError = 0;
	*aRetVal = size;
	return OK;
}



static void FileOrDirExist(LPCTSTR aFilePattern, StrRet &aRetVal, DWORD aRequiredAttr)
{
	LPTSTR buf = aRetVal.CallerBuf();
	aRetVal.SetStatic(buf);
	DWORD attr;
	if (DoesFilePatternExist(aFilePattern, &attr, aRequiredAttr))
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
	else // Empty string is the indicator of "not found".
		*buf = '\0';
}

bif_impl void FileExist(LPCTSTR aFilePattern, StrRet &aRetVal)
{
	return FileOrDirExist(aFilePattern, aRetVal, 0);
}

bif_impl void DirExist(LPCTSTR aFilePattern, StrRet &aRetVal)
{
	return FileOrDirExist(aFilePattern, aRetVal, FILE_ATTRIBUTE_DIRECTORY);
}


bif_impl FResult DirCopy(LPCTSTR aSource, LPCTSTR aDest, int *aOverwrite)
{
	return Line::Util_CopyDir(aSource, aDest, aOverwrite ? *aOverwrite : FALSE, false) ? OK : FR_E_FAILED;
}


bif_impl FResult DirMove(LPCTSTR aSource, LPCTSTR aDest, LPCTSTR aFlag)
{
	if (aFlag && ctoupper(*aFlag) == 'R')
	{
		// Perform a simple rename instead, which prevents the operation from being only partially
		// complete if the source directory is in use (due to being a working dir for a currently
		// running process, or containing a file that is being written to).  In other words,
		// the operation will be "all or none":
		if (!MoveFile(aSource, aDest))
			return FR_E_WIN32;
	}
	if (aFlag && !IsNumeric(aFlag))
		return FR_E_ARG(2);
	if (!Line::Util_CopyDir(aSource, aDest, aFlag ? ATOI(aFlag) : 0, true))
		return FR_E_FAILED;
	return OK;
}


bif_impl FResult DirCreate(LPCTSTR aPath)
{
	SetLastError(0);
	auto result = FileCreateDir(aPath);
	g->LastError = GetLastError();
	return result;
}


bif_impl FResult DirDelete(LPCTSTR aPath, BOOL *aRecurse)
{
	return Line::Util_RemoveDir(aPath, aRecurse && *aRecurse) ? OK : FR_E_FAILED;
}


