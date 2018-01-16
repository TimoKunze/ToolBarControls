// TBarCtls.cpp: Implementation of DLL exports.

#include "stdafx.h"
#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif


class TBarCtlsModule :
    public CAtlDllModuleT<TBarCtlsModule>
{

public:
  #ifdef _UNICODE
		DECLARE_LIBID(LIBID_TBarCtlsLibU)
		DECLARE_REGISTRY_APPID_RESOURCEID(IDR_TBARCTLS, "{BE825E82-A8D4-468d-8F58-6E364EC4A419}")
	#else
		DECLARE_LIBID(LIBID_TBarCtlsLibA)
		DECLARE_REGISTRY_APPID_RESOURCEID(IDR_TBARCTLS, "{30AF702B-F9F8-42dd-A8D2-FF3AFA47E2C8}")
	#endif
};

TBarCtlsModule _AtlModule;


#ifdef _MANAGED
	#pragma managed(push, off)
#endif

// DLL entry point
extern "C" BOOL WINAPI DllMain(HINSTANCE /*hInstance*/, DWORD dwReason, LPVOID lpReserved)
{
	#ifdef _DEBUG
		// enable CRT memory leak detection & report
		_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF     // enable debug heap allocs & block type IDs (ie _CLIENT_BLOCK)
		    //| _CRTDBG_CHECK_CRT_DF              // check CRT allocations too
		    //| _CRTDBG_DELAY_FREE_MEM_DF       // keep freed blocks in list as _FREE_BLOCK type
		    | _CRTDBG_LEAK_CHECK_DF             // do leak report at exit (_CrtDumpMemoryLeaks)

		    // pick only one of these heap check frequencies
		    //| _CRTDBG_CHECK_ALWAYS_DF         // check heap on every alloc/free
		    //| _CRTDBG_CHECK_EVERY_16_DF
		    //| _CRTDBG_CHECK_EVERY_128_DF
		    //| _CRTDBG_CHECK_EVERY_1024_DF
		    | _CRTDBG_CHECK_DEFAULT_DF          // by default, no heap checks
		    );
		//_CrtSetBreakAlloc(209);               // break debugger on numbered allocation
		// get ID number from leak detector report of previous run
	#endif

	return _AtlModule.DllMain(dwReason, lpReserved); 
}

#ifdef _MANAGED
	#pragma managed(pop)
#endif

STDAPI DllCanUnloadNow(void)
{
	return _AtlModule.DllCanUnloadNow();
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
	return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}

STDAPI DllRegisterServer(void)
{
	return _AtlModule.DllRegisterServer();
}

STDAPI DllUnregisterServer(void)
{
	return _AtlModule.DllUnregisterServer();
}