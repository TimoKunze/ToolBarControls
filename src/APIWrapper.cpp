// APIWrapper.cpp: A wrapper class for API functions not available on all supported systems

#include "stdafx.h"
#include "APIWrapper.h"


BOOL APIWrapper::checkedSupport_CalculatePopupWindowPosition = FALSE;
BOOL APIWrapper::checkedSupport_SetWindowTheme = FALSE;
BOOL APIWrapper::checkedSupport_SHGetImageList = FALSE;
CalculatePopupWindowPositionFn* APIWrapper::pfnCalculatePopupWindowPosition = NULL;
SetWindowThemeFn* APIWrapper::pfnSetWindowTheme = NULL;
SHGetImageListFn* APIWrapper::pfnSHGetImageList = NULL;


APIWrapper::APIWrapper(void)
{
}


BOOL APIWrapper::IsSupported_CalculatePopupWindowPosition(void)
{
	if(!checkedSupport_CalculatePopupWindowPosition) {
		HMODULE hUser32 = GetModuleHandle(TEXT("user32.dll"));
		if(hUser32) {
			pfnCalculatePopupWindowPosition = reinterpret_cast<CalculatePopupWindowPositionFn*>(GetProcAddress(hUser32, "CalculatePopupWindowPosition"));
		}
		checkedSupport_CalculatePopupWindowPosition = TRUE;
	}

	return (pfnCalculatePopupWindowPosition != NULL);
}

BOOL APIWrapper::IsSupported_SetWindowTheme(void)
{
	if(!checkedSupport_SetWindowTheme) {
		HMODULE hUxTheme = GetModuleHandle(TEXT("uxtheme.dll"));
		if(hUxTheme) {
			pfnSetWindowTheme = reinterpret_cast<SetWindowThemeFn*>(GetProcAddress(hUxTheme, "SetWindowTheme"));
		}
		checkedSupport_SetWindowTheme = TRUE;
	}

	return (pfnSetWindowTheme != NULL);
}

BOOL APIWrapper::IsSupported_SHGetImageList(void)
{
	if(!checkedSupport_SHGetImageList) {
		HMODULE hShell32 = GetModuleHandle(TEXT("shell32.dll"));
		if(hShell32) {
			pfnSHGetImageList = reinterpret_cast<SHGetImageListFn*>(GetProcAddress(hShell32, "SHGetImageList"));
			if(!pfnSHGetImageList) {
				pfnSHGetImageList = reinterpret_cast<SHGetImageListFn*>(GetProcAddress(hShell32, MAKEINTRESOURCEA(727)));
			}
		}
		checkedSupport_SHGetImageList = TRUE;
	}

	return (pfnSHGetImageList != NULL);
}

HRESULT APIWrapper::CalculatePopupWindowPosition(const LPPOINT pAnchorPoint, const LPSIZE pWindowSize, UINT flags, LPRECT pExcludeRect, LPRECT pPopupWindowPosition, BOOL* pReturnValue)
{
	BOOL dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_CalculatePopupWindowPosition()) {
		*pReturnValue = pfnCalculatePopupWindowPosition(pAnchorPoint, pWindowSize, flags, pExcludeRect, pPopupWindowPosition);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SetWindowTheme(HWND hWnd, LPCWSTR pSubAppName, LPCWSTR pSubIdList, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SetWindowTheme()) {
		*pReturnValue = pfnSetWindowTheme(hWnd, pSubAppName, pSubIdList);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}

HRESULT APIWrapper::SHGetImageList(int imageList, REFIID requiredInterface, LPVOID* ppInterfaceImpl, HRESULT* pReturnValue)
{
	HRESULT dummy;
	if(!pReturnValue) {
		pReturnValue = &dummy;
	}

	if(IsSupported_SHGetImageList()) {
		*pReturnValue = pfnSHGetImageList(imageList, requiredInterface, ppInterfaceImpl);
		return S_OK;
	} else {
		return E_NOTIMPL;
	}
}