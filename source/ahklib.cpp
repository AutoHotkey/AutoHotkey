/*
ahklib.cpp

Original code by Steve Gray.

This software is provided 'as-is', without any express or implied
warranty. In no event will the authors be held liable for any damages
arising from the use of this software.

Permission is granted to anyone to use this software for any purpose,
including commercial applications, and to alter it and redistribute it
freely, without restriction.
*/

#include "stdafx.h"
#include "globaldata.h"


#define AHKAPI(_rettype_) extern "C" __declspec(dllexport) _rettype_ __stdcall


AHKAPI(int) Main(int argc, LPTSTR argv[])
{
	__argc = argc;
	__targv = argv;
	// _tWinMain() doesn't use lpCmdLine or nShowCmd.
	return _tWinMain(g_hInstance, NULL, nullptr, SW_SHOWNORMAL);
}


BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
	switch( fdwReason ) 
	{ 
	case DLL_PROCESS_ATTACH:
		g_hInstance = hinstDLL;
		break;
	case DLL_PROCESS_DETACH:
		if (lpvReserved != nullptr)
			break;
		// Perform any necessary cleanup.
		break;
	}
	return TRUE;
}
