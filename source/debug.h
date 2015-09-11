#pragma once
// This file debugging defines some macros like MFC does.

#ifndef TRACE
	#ifdef _DEBUG
		#ifdef _MSC_VER
			#define TRACE OutputDebugStringFormat
		#else
			#define TRACE(...) _ftprintf(stderr, __VA_ARGS__)
		#endif
	#else
		#define TRACE(...)
	#endif
#endif

/*
This part of codes map the new operator to the debug version. Although the map is contains in "crtdbg.h",
it is not really work (we will always get a wrong information show us the leaked memory blocks are allocated in "crtdbg.h").
*/
#ifdef _MSC_VER
//	#define _CRTDBG_MAP_ALLOC
	#include <crtdbg.h>
//	#ifdef _DEBUG
//		#define new new(_NORMAL_BLOCK, __FILE__, __LINE__)
//	#endif
#endif

#ifndef ASSERT
	#ifdef _ASSERTE
		#define ASSERT(expr) _ASSERTE(expr)
	#else
		#ifdef _DEBUG
			#include <assert.h>
			#define ASSERT(expr) assert(expr)
		#else
			#define ASSERT(expr)
		#endif
	#endif
#endif

#ifndef VERIFY
	#ifdef _DEBUG
		#define VERIFY(expr) ASSERT(expr)
	#else
		#define VERIFY(expr) expr
	#endif
#endif
