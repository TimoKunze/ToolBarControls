// VirtualToolBarButton.cpp: A wrapper for non-existing tool bar buttons.

#include "stdafx.h"
#include "VirtualToolBarButton.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualToolBarButton::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualToolBarButton, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualToolBarButton::Properties::~Properties()
{
	if(freeString && !IS_INTRESOURCE(settings.iString)) {
		free(reinterpret_cast<LPVOID>(settings.iString));
	}
	if(pOwnerTBar) {
		pOwnerTBar->Release();
	}
}

HWND VirtualToolBarButton::Properties::GetTBarHWnd(void)
{
	ATLASSUME(pOwnerTBar);

	OLE_HANDLE handle = NULL;
	pOwnerTBar->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT VirtualToolBarButton::LoadState(LPTBBUTTON pSource, int buttonIndex)
{
	ATLASSERT_POINTER(pSource, TBBUTTON);

	if(properties.freeString && !IS_INTRESOURCE(properties.settings.iString)) {
		free(reinterpret_cast<LPVOID>(properties.settings.iString));
		properties.settings.iString = NULL;
		properties.freeString = FALSE;
	}
	properties.settings = *pSource;
	if(!IS_INTRESOURCE(properties.settings.iString)) {
		// duplicate the button's text
		if(properties.settings.iString != reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACK)) {
			int bufferSize = lstrlen(reinterpret_cast<LPCTSTR>(pSource->iString));
			properties.settings.iString = reinterpret_cast<INT_PTR>(malloc((bufferSize + 1) * sizeof(TCHAR)));
			properties.freeString = TRUE;
			ATLVERIFY(SUCCEEDED(StringCchCopy(reinterpret_cast<LPTSTR>(properties.settings.iString), bufferSize + 1, reinterpret_cast<LPCTSTR>(pSource->iString))));
		}
	}
	properties.buttonIndex = buttonIndex;

	return S_OK;
}

HRESULT VirtualToolBarButton::SaveState(LPTBBUTTON pTarget, int& buttonIndex)
{
	ATLASSERT_POINTER(pTarget, TBBUTTON);

	if(!IS_INTRESOURCE(pTarget->iString)) {
		free(reinterpret_cast<LPVOID>(pTarget->iString));
		pTarget->iString = NULL;
	}
	*pTarget = properties.settings;
	if(!IS_INTRESOURCE(properties.settings.iString)) {
		// duplicate the button's text
		if(pTarget->iString != reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACK)) {
			int bufferSize = lstrlen(reinterpret_cast<LPCTSTR>(properties.settings.iString));
			pTarget->iString = reinterpret_cast<INT_PTR>(malloc((bufferSize + 1) * sizeof(TCHAR)));
			ATLASSERT(pTarget->iString);
			if(pTarget->iString) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(reinterpret_cast<LPTSTR>(pTarget->iString), bufferSize + 1, reinterpret_cast<LPCTSTR>(properties.settings.iString))));
			}
		}
	}
	buttonIndex = properties.buttonIndex;

	return S_OK;
}

void VirtualToolBarButton::SetOwner(ToolBar* pOwner)
{
	if(properties.pOwnerTBar) {
		properties.pOwnerTBar->Release();
	}
	properties.pOwnerTBar = pOwner;
	if(properties.pOwnerTBar) {
		properties.pOwnerTBar->AddRef();
	}
}


STDMETHODIMP VirtualToolBarButton::get_AutoSize(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.settings.fsStyle & BTNS_AUTOSIZE);
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_AutoSize(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != VARIANT_FALSE) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsStyle &= ~BTNS_AUTOSIZE;
	} else {
		properties.settings.fsStyle |= BTNS_AUTOSIZE;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_ButtonData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.settings.dwData;
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_ButtonData(LONG newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}

	properties.settings.dwData = newValue;
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_ButtonType(ButtonTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ButtonTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fsStyle & BTNS_SEP) {
		*pValue = btySeparator;
	} else if(properties.settings.fsStyle & BTNS_CHECK) {
		*pValue = btyCheckButton;
	} else if(properties.settings.fsStyle == BTNS_BUTTON) {
		if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand)) {
			*pValue = btyPlaceholder;
		} else {
			*pValue = btyCommandButton;
		}
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_ButtonType(ButtonTypeConstants newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}

	BOOL isPlaceholder = properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand);
	properties.settings.fsStyle &= ~(BTNS_BUTTON | BTNS_SEP | BTNS_CHECK);
	switch(newValue) {
		case btyCommandButton:
			properties.settings.fsStyle |= BTNS_BUTTON;
			break;
		case btyCheckButton:
			properties.settings.fsStyle |= BTNS_CHECK;
			break;
		case btySeparator:
			properties.settings.fsStyle |= BTNS_SEP;
			break;
		case btyPlaceholder:
			properties.settings.fsStyle |= BTNS_BUTTON;
			break;
	}
	if(isPlaceholder) {
		if(newValue != btyPlaceholder) {
			properties.pOwnerTBar->DeregisterPlaceholderButton(properties.settings.idCommand);
		}
	} else if(newValue == btyPlaceholder) {
		properties.pOwnerTBar->RegisterPlaceholderButton(properties.settings.idCommand);
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_DisplayText(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.settings.fsStyle & BTNS_SHOWTEXT);
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_DisplayText(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != VARIANT_FALSE) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsStyle &= ~BTNS_SHOWTEXT;
	} else {
		properties.settings.fsStyle |= BTNS_SHOWTEXT;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_DropDownStyle(DropDownStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DropDownStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fsStyle & BTNS_WHOLEDROPDOWN) {
		*pValue = ddstAlwaysWholeButton;
	} else if(properties.settings.fsStyle & BTNS_DROPDOWN) {
		*pValue = ddstNormal;
	} else {
		*pValue = ddstNoDropDown;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_DropDownStyle(DropDownStyleConstants newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != ddstNoDropDown) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	properties.settings.fsStyle &= ~(BTNS_WHOLEDROPDOWN | BTNS_DROPDOWN);
	switch(newValue) {
		case ddstNormal:
			properties.settings.fsStyle |= BTNS_DROPDOWN;
			break;
		case ddstAlwaysWholeButton:
			properties.settings.fsStyle |= BTNS_WHOLEDROPDOWN;
			break;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.settings.fsState & TBSTATE_ENABLED);
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_Enabled(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != VARIANT_FALSE) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsState &= ~TBSTATE_ENABLED;
	} else {
		properties.settings.fsState |= TBSTATE_ENABLED;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_FollowedByLineBreak(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.settings.fsState & TBSTATE_WRAP);
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_FollowedByLineBreak(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsState &= ~TBSTATE_WRAP;
	} else {
		properties.settings.fsState |= TBSTATE_WRAP;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fsStyle & BTNS_SEP) {
		*pValue = I_IMAGENONE;
	} else if(properties.settings.iBitmap == -1) {
		*pValue = I_IMAGECALLBACK;
	} else {
		*pValue = LOWORD(properties.settings.iBitmap);
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_IconIndex(LONG newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(newValue < -2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != -2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(!(properties.settings.fsStyle & BTNS_SEP)) {
		if(newValue == -2) {
			properties.settings.iBitmap = I_IMAGENONE;
		} else if(newValue == -1) {
			properties.settings.iBitmap = I_IMAGECALLBACK;
		} else {
			properties.settings.iBitmap = MAKELONG(newValue, HIWORD(properties.settings.iBitmap));
		}
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.settings.idCommand;
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_ID(LONG newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}

	properties.settings.idCommand = newValue;
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_ImageListIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fsStyle & BTNS_SEP) {
		*pValue = 0;
	} else if(properties.settings.iBitmap == -1) {
		*pValue = I_IMAGECALLBACK;
	} else {
		*pValue = HIWORD(properties.settings.iBitmap);
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_ImageListIndex(LONG newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(!(properties.settings.fsStyle & BTNS_SEP)) {
		if(newValue == -1) {
			properties.settings.iBitmap = I_IMAGECALLBACK;
		} else {
			properties.settings.iBitmap = MAKELONG(LOWORD(properties.settings.iBitmap), newValue);
		}
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.buttonIndex;
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_Marked(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.settings.fsState & TBSTATE_MARKED);
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_Marked(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != VARIANT_FALSE) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsState &= ~TBSTATE_MARKED;
	} else {
		properties.settings.fsState |= TBSTATE_MARKED;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_PartOfGroup(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.settings.fsStyle & BTNS_GROUP);
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_PartOfGroup(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsStyle &= ~BTNS_GROUP;
	} else {
		properties.settings.fsStyle |= BTNS_GROUP;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_Pushed(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.settings.fsState & TBSTATE_PRESSED);
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_Pushed(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != VARIANT_FALSE) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsState &= ~TBSTATE_PRESSED;
	} else {
		properties.settings.fsState |= TBSTATE_PRESSED;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_SelectionState(SelectionStateConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SelectionStateConstants);
	if(!pValue) {
		return E_POINTER;
	}

	switch(properties.settings.fsState & (TBSTATE_CHECKED | TBSTATE_INDETERMINATE)) {
		case 0:
			*pValue = ssUnchecked;
			break;
		case TBSTATE_CHECKED:
			*pValue = ssChecked;
			break;
		case TBSTATE_INDETERMINATE:
			*pValue = ssIndeterminate;
			break;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_SelectionState(SelectionStateConstants newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != ssUnchecked) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	properties.settings.fsState &= ~(TBSTATE_CHECKED | TBSTATE_INDETERMINATE);
	switch(newValue) {
		case ssChecked:
			properties.settings.fsState |= TBSTATE_CHECKED;
			break;
		case ssIndeterminate:
			properties.settings.fsState |= TBSTATE_INDETERMINATE;
			break;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_ShowingEllipsis(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.settings.fsState & TBSTATE_ELLIPSES);
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_ShowingEllipsis(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != VARIANT_FALSE) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsState &= ~TBSTATE_ELLIPSES;
	} else {
		properties.settings.fsState |= TBSTATE_ELLIPSES;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(IS_INTRESOURCE(properties.settings.iString)) {
		HWND hWndTBar = properties.GetTBarHWnd();
		ATLASSERT(IsWindow(hWndTBar));

		LPWSTR pBuffer = static_cast<LPWSTR>(HeapAlloc(GetProcessHeap(), 0, (MAX_BUTTONTEXTLENGTH + 1) * sizeof(WCHAR)));
		if(!pBuffer) {
			return E_OUTOFMEMORY;
		}
		ZeroMemory(pBuffer, (MAX_BUTTONTEXTLENGTH + 1) * sizeof(WCHAR));
		if(SendMessage(hWndTBar, TB_GETSTRING, MAKEWPARAM((MAX_BUTTONTEXTLENGTH + 1) * sizeof(WCHAR), properties.settings.iString), reinterpret_cast<LPARAM>(pBuffer)) != -1) {
			*pValue = _bstr_t(pBuffer).Detach();
		}
		HeapFree(GetProcessHeap(), 0, pBuffer);
	} else {
		*pValue = _bstr_t(reinterpret_cast<LPCTSTR>(properties.settings.iString)).Detach();
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_Text(BSTR newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue && lstrlenW(newValue) > 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.freeString && !IS_INTRESOURCE(properties.settings.iString)) {
		free(reinterpret_cast<LPVOID>(properties.settings.iString));
		properties.settings.iString = NULL;
		properties.freeString = FALSE;
	}
	if(newValue) {
		// duplicate the button's text
		COLE2T converter(newValue);
		int bufferSize = lstrlen(converter);
		properties.settings.iString = reinterpret_cast<INT_PTR>(malloc((bufferSize + 1) * sizeof(TCHAR)));
		properties.freeString = TRUE;
		ATLVERIFY(SUCCEEDED(StringCchCopy(reinterpret_cast<LPTSTR>(properties.settings.iString), bufferSize + 1, converter)));
	} else {
		properties.settings.iString = reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACK);
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_UseMnemonic(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(!(properties.settings.fsStyle & BTNS_NOPREFIX));
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_UseMnemonic(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}
	if(properties.pOwnerTBar->IsPlaceholderButton(properties.settings.idCommand) && newValue != VARIANT_FALSE) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsStyle |= BTNS_NOPREFIX;
	} else {
		properties.settings.fsStyle &= ~BTNS_NOPREFIX;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_Visible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(!(properties.settings.fsState & TBSTATE_HIDDEN));
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_Visible(VARIANT_BOOL newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}

	if(newValue == VARIANT_FALSE) {
		properties.settings.fsState |= TBSTATE_HIDDEN;
	} else {
		properties.settings.fsState &= ~TBSTATE_HIDDEN;
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::get_Width(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fsStyle & BTNS_SEP) {
		*pValue = properties.settings.iBitmap;
	} else {
		*pValue = MAKEWORD(properties.settings.bReserved[1], properties.settings.bReserved[0]);
	}
	return S_OK;
}

STDMETHODIMP VirtualToolBarButton::put_Width(OLE_XSIZE_PIXELS newValue)
{
	if(properties.readOnly) {
		// raise VB runtime error 383
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 383);
	}

	if(properties.settings.fsStyle & BTNS_SEP) {
		properties.settings.iBitmap = newValue;
	} else {
		properties.settings.bReserved[0] = HIBYTE(LOWORD(newValue));
		properties.settings.bReserved[1] = LOBYTE(LOWORD(newValue));
	}
	return S_OK;
}