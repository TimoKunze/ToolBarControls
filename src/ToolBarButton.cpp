// ToolBarButton.cpp: A wrapper for existing tool bar buttons.

#include "stdafx.h"
#include "ToolBarButton.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ToolBarButton::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IToolBarButton, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ToolBarButton::Properties::~Properties()
{
	if(pOwnerTBar) {
		pOwnerTBar->Release();
	}
}

HWND ToolBarButton::Properties::GetTBarHWnd(void)
{
	ATLASSUME(pOwnerTBar);

	OLE_HANDLE handle = NULL;
	if(partOfChevronPopup) {
		pOwnerTBar->get_hWndChevronToolBar(&handle);
	} else {
		pOwnerTBar->get_hWnd(&handle);
	}
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT ToolBarButton::SaveState(int buttonIndex, LPTBBUTTON pTarget, HWND hWndTBar/* = NULL*/)
{
	if(!hWndTBar) {
		hWndTBar = properties.GetTBarHWnd();
	}
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}
	if((buttonIndex < 0) || (buttonIndex >= static_cast<int>(SendMessage(hWndTBar, TB_BUTTONCOUNT, 0, 0)))) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pTarget, TBBUTTON);
	if(!pTarget) {
		return E_POINTER;
	}

	ZeroMemory(pTarget, sizeof(TBBUTTON));
	if(SendMessage(hWndTBar, TB_GETBUTTON, properties.buttonIndex, reinterpret_cast<LPARAM>(pTarget))) {
		return S_OK;
	}
	return E_FAIL;
}


void ToolBarButton::Attach(int buttonIndex, BOOL partOfChevronPopup)
{
	properties.buttonIndex = buttonIndex;
	properties.partOfChevronPopup = partOfChevronPopup;
}

void ToolBarButton::Detach(void)
{
	properties.buttonIndex = -1;
}

HRESULT ToolBarButton::LoadState(LPTBBUTTON pSource)
{
	ATLASSERT_POINTER(pSource, TBBUTTON);
	if(!pSource) {
		return E_POINTER;
	}

	put_AutoSize(BOOL2VARIANTBOOL(pSource->fsStyle & BTNS_AUTOSIZE));
	put_ButtonData(pSource->dwData);
	if(pSource->fsStyle & BTNS_SEP) {
		put_ButtonType(btySeparator);
	} else if(pSource->fsStyle & BTNS_CHECK) {
		put_ButtonType(btyCheckButton);
	} else if(pSource->fsStyle == BTNS_BUTTON) {
		put_ButtonType(btyCommandButton);
	}
	put_DisplayText(BOOL2VARIANTBOOL(pSource->fsStyle & BTNS_SHOWTEXT));
	if(pSource->fsStyle & BTNS_WHOLEDROPDOWN) {
		put_DropDownStyle(ddstAlwaysWholeButton);
	} else if(pSource->fsStyle & BTNS_DROPDOWN) {
		put_DropDownStyle(ddstNormal);
	} else {
		put_DropDownStyle(ddstNoDropDown);
	}
	put_Enabled(BOOL2VARIANTBOOL(pSource->fsState & TBSTATE_ENABLED));
	put_FollowedByLineBreak(BOOL2VARIANTBOOL(pSource->fsState & TBSTATE_WRAP));
	if(pSource->fsStyle & BTNS_SEP) {
		put_IconIndex(I_IMAGENONE);
		put_ImageListIndex(0);
	} else {
		put_IconIndex(LOWORD(pSource->iBitmap));
		put_ImageListIndex(HIWORD(pSource->iBitmap));
	}
	put_ID(pSource->idCommand);
	put_Marked(BOOL2VARIANTBOOL(pSource->fsState & TBSTATE_MARKED));
	put_PartOfGroup(BOOL2VARIANTBOOL(pSource->fsStyle & BTNS_GROUP));
	put_Pushed(BOOL2VARIANTBOOL(pSource->fsState & TBSTATE_PRESSED));
	switch(pSource->fsState & (TBSTATE_CHECKED | TBSTATE_INDETERMINATE)) {
		case 0:
			put_SelectionState(ssUnchecked);
			break;
		case TBSTATE_CHECKED:
			put_SelectionState(ssChecked);
			break;
		case TBSTATE_INDETERMINATE:
			put_SelectionState(ssIndeterminate);
			break;
	}
	put_ShowingEllipsis(BOOL2VARIANTBOOL(pSource->fsState & TBSTATE_ELLIPSES));
	BSTR text = NULL;
	if(IS_INTRESOURCE(pSource->iString)) {
		// TODO
	} else {
		if(pSource->iString != reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACK)) {
			text = _bstr_t(reinterpret_cast<LPCTSTR>(pSource->iString)).Detach();
		}
		put_Text(text);
	}
	SysFreeString(text);
	put_UseMnemonic(BOOL2VARIANTBOOL(!(pSource->fsStyle & BTNS_NOPREFIX)));
	put_Visible(BOOL2VARIANTBOOL(!(pSource->fsState & TBSTATE_HIDDEN)));
	if(pSource->fsStyle & BTNS_SEP) {
		put_Width(pSource->iBitmap);
	} else {
		put_Width(MAKEWORD(pSource->bReserved[1], pSource->bReserved[0]));
	}
	return S_OK;
}

HRESULT ToolBarButton::LoadState(VirtualToolBarButton* pSource)
{
	ATLASSUME(pSource);
	if(!pSource) {
		return E_POINTER;
	}

	VARIANT_BOOL b = VARIANT_FALSE;
	pSource->get_AutoSize(&b);
	put_AutoSize(b);
	b = VARIANT_FALSE;
	pSource->get_DisplayText(&b);
	put_DisplayText(b);
	b = VARIANT_FALSE;
	pSource->get_Enabled(&b);
	put_Enabled(b);
	b = VARIANT_FALSE;
	pSource->get_FollowedByLineBreak(&b);
	put_FollowedByLineBreak(b);
	b = VARIANT_FALSE;
	pSource->get_Marked(&b);
	put_Marked(b);
	b = VARIANT_FALSE;
	pSource->get_PartOfGroup(&b);
	put_PartOfGroup(b);
	b = VARIANT_FALSE;
	pSource->get_Pushed(&b);
	put_Pushed(b);
	b = VARIANT_FALSE;
	pSource->get_ShowingEllipsis(&b);
	put_ShowingEllipsis(b);
	b = VARIANT_TRUE;
	pSource->get_UseMnemonic(&b);
	put_UseMnemonic(b);
	b = VARIANT_TRUE;
	pSource->get_Visible(&b);
	put_Visible(b);

	LONG l = 0;
	pSource->get_ButtonData(&l);
	put_ButtonData(l);
	l = -2;
	pSource->get_IconIndex(&l);
	put_IconIndex(l);
	l = 0;
	pSource->get_ID(&l);
	put_ID(l);
	l = 0;
	pSource->get_ImageListIndex(&l);
	put_ImageListIndex(l);
	l = 0;
	pSource->get_Index(&l);
	put_Index(l);

	OLE_XSIZE_PIXELS cx = 0;
	pSource->get_Width(&cx);
	put_Width(cx);

	ButtonTypeConstants bty = btyCommandButton;
	pSource->get_ButtonType(&bty);
	put_ButtonType(bty);

	DropDownStyleConstants ddst = ddstNoDropDown;
	pSource->get_DropDownStyle(&ddst);
	put_DropDownStyle(ddst);

	SelectionStateConstants ss = ssUnchecked;
	pSource->get_SelectionState(&ss);
	put_SelectionState(ss);

	BSTR text = NULL;
	pSource->get_Text(&text);
	put_Text(text);
	SysFreeString(text);
	return S_OK;
}

HRESULT ToolBarButton::SaveState(LPTBBUTTON pTarget, HWND hWndTBar/* = NULL*/)
{
	return SaveState(properties.buttonIndex, pTarget, hWndTBar);
}

HRESULT ToolBarButton::SaveState(VirtualToolBarButton* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	pTarget->SetOwner(properties.pOwnerTBar);
	TBBUTTON button = {0};
	HRESULT hr = SaveState(&button);
	pTarget->LoadState(&button, properties.buttonIndex);
	// We did not allocate the string buffer, so don't free it.
	// TODO: Will this always work this way, that *we* don't allocate the string buffer?
	/*if(!IS_INTRESOURCE(button.iString)) {
		free(reinterpret_cast<LPVOID>(button.iString));
	}*/

	return hr;
}

void ToolBarButton::SetOwner(ToolBar* pOwner)
{
	if(properties.pOwnerTBar) {
		properties.pOwnerTBar->Release();
	}
	properties.pOwnerTBar = pOwner;
	if(properties.pOwnerTBar) {
		properties.pOwnerTBar->AddRef();
	}
}


STDMETHODIMP ToolBarButton::get_AutoSize(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(button.fsStyle & BTNS_AUTOSIZE);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_AutoSize(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != VARIANT_FALSE) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		if(newValue == VARIANT_FALSE) {
			button.fsStyle &= ~BTNS_AUTOSIZE;
		} else {
			button.fsStyle |= BTNS_AUTOSIZE;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_ButtonData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_LPARAM;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = button.lParam;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_ButtonData(LONG newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_LPARAM;
	button.lParam = newValue;
	if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_ButtonType(ButtonTypeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ButtonTypeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(button.fsStyle & BTNS_SEP) {
			*pValue = btySeparator;
		} else if(button.fsStyle & BTNS_CHECK) {
			*pValue = btyCheckButton;
		} else if(button.fsStyle == BTNS_BUTTON) {
			if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand)) {
				*pValue = btyPlaceholder;
			} else {
				*pValue = btyCommandButton;
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_ButtonType(ButtonTypeConstants newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE | TBIF_COMMAND;
	BOOL isPlaceholder = FALSE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		isPlaceholder = properties.pOwnerTBar->IsPlaceholderButton(button.idCommand);
	}
	button.dwMask &= ~TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		button.fsStyle &= ~(BTNS_BUTTON | BTNS_SEP | BTNS_CHECK);
		switch(newValue) {
			case btyCommandButton:
				button.fsStyle |= BTNS_BUTTON;
				break;
			case btyCheckButton:
				button.fsStyle |= BTNS_CHECK;
				break;
			case btySeparator:
				button.fsStyle |= BTNS_SEP;
				break;
			case btyPlaceholder:
				button.fsStyle |= BTNS_BUTTON;
				break;
		}
		if(isPlaceholder) {
			if(newValue != btyPlaceholder) {
				properties.pOwnerTBar->DeregisterPlaceholderButton(button.idCommand);
			}
		} else if(newValue == btyPlaceholder) {
			properties.pOwnerTBar->RegisterPlaceholderButton(button.idCommand);
		}

		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_DisplayText(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(button.fsStyle & BTNS_SHOWTEXT);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_DisplayText(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_COMMAND | TBIF_STYLE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != VARIANT_FALSE) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		int textLength = static_cast<int>(SendMessage(hWndTBar, TB_GETBUTTONTEXT, button.idCommand, NULL));
		if(textLength > -1) {
			button.dwMask &= ~TBIF_COMMAND;
			LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
			if(pBuffer) {
				// just toggling BTNS_SHOWTEXT leeds to drawing issues, so also set the text
				ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
				button.dwMask |= TBIF_TEXT;
				button.cchText = textLength;
				button.pszText = pBuffer;
			}
			if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
				if(newValue == VARIANT_FALSE) {
					button.fsStyle &= ~BTNS_SHOWTEXT;
				} else {
					button.fsStyle |= BTNS_SHOWTEXT;
				}
				if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
					hr = S_OK;
				}
			}
			if(pBuffer) {
				HeapFree(GetProcessHeap(), 0, pBuffer);
			}
		}
	}
	return hr;
}

STDMETHODIMP ToolBarButton::get_DropDownStyle(DropDownStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DropDownStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(button.fsStyle & BTNS_WHOLEDROPDOWN) {
			*pValue = ddstAlwaysWholeButton;
		} else if(button.fsStyle & BTNS_DROPDOWN) {
			*pValue = ddstNormal;
		} else {
			*pValue = ddstNoDropDown;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_DropDownStyle(DropDownStyleConstants newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != ddstNoDropDown) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		button.fsStyle &= ~(BTNS_WHOLEDROPDOWN | BTNS_DROPDOWN);
		switch(newValue) {
			case ddstNormal:
				button.fsStyle |= BTNS_DROPDOWN;
				break;
			case ddstAlwaysWholeButton:
				button.fsStyle |= BTNS_WHOLEDROPDOWN;
				break;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_DroppedDown(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.pOwnerTBar->IsDroppedDownButton(properties.buttonIndex));
	return S_OK;
}

STDMETHODIMP ToolBarButton::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STATE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(button.fsState & TBSTATE_ENABLED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_Enabled(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STATE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != VARIANT_FALSE) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		if(newValue == VARIANT_FALSE) {
			button.fsState &= ~TBSTATE_ENABLED;
		} else {
			button.fsState |= TBSTATE_ENABLED;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_FollowedByLineBreak(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STATE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(button.fsState & TBSTATE_WRAP);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_FollowedByLineBreak(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STATE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(newValue == VARIANT_FALSE) {
			button.fsState &= ~TBSTATE_WRAP;
		} else {
			button.fsState |= TBSTATE_WRAP;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_Hot(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	*pValue = BOOL2VARIANTBOOL(properties.buttonIndex == static_cast<int>(SendMessage(hWndTBar, TB_GETHOTITEM, 0, 0)));
	return S_OK;
}

STDMETHODIMP ToolBarButton::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_IMAGE | TBIF_STYLE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(button.fsStyle & BTNS_SEP) {
			*pValue = I_IMAGENONE;
		} else if(button.iImage == I_IMAGENONE) {
			*pValue = I_IMAGENONE;
		} else {
			*pValue = LOWORD(button.iImage);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_IconIndex(LONG newValue)
{
	if(newValue < -2) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_IMAGE | TBIF_STYLE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != -2) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		if(!(button.fsStyle & BTNS_SEP)) {
			if(newValue == -2) {
				button.iImage = I_IMAGENONE;
			} else if(newValue == -1) {
				button.iImage = I_IMAGECALLBACK;
			} else {
				button.iImage = MAKELONG(newValue, HIWORD(button.iImage));
			}
			if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
				return S_OK;
			}
		} else {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = button.idCommand;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_ID(LONG newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	if(SendMessage(hWndTBar, TB_SETCMDID, properties.buttonIndex, newValue)) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_ImageListIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_IMAGE | TBIF_STYLE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(button.fsStyle & BTNS_SEP) {
			*pValue = 0;
		} else if(button.iImage == I_IMAGENONE) {
			*pValue = 0;
		} else {
			*pValue = HIWORD(button.iImage);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_ImageListIndex(LONG newValue)
{
	if(newValue < -1) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_IMAGE | TBIF_STYLE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != 0) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		if(!(button.fsStyle & BTNS_SEP)) {
			if(newValue == -1) {
				button.iImage = I_IMAGECALLBACK;
			} else {
				button.iImage = MAKELONG(LOWORD(button.iImage), newValue);
			}
			if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
				return S_OK;
			}
		} else {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.buttonIndex;
	return S_OK;
}

STDMETHODIMP ToolBarButton::put_Index(LONG newValue)
{
	if(newValue == properties.buttonIndex) {
		return S_OK;
	}
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	int numberOfButtons = static_cast<int>(SendMessage(hWndTBar, TB_BUTTONCOUNT, 0, 0));
	if(newValue >= numberOfButtons) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(SendMessage(hWndTBar, TB_MOVEBUTTON, properties.buttonIndex, newValue)) {
		properties.buttonIndex = newValue;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_Marked(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STATE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(button.fsState & TBSTATE_MARKED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_Marked(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STATE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != VARIANT_FALSE) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		if(newValue == VARIANT_FALSE) {
			button.fsState &= ~TBSTATE_MARKED;
		} else {
			button.fsState |= TBSTATE_MARKED;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_PartOfChevronToolBar(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.partOfChevronPopup);
	return S_OK;
}

STDMETHODIMP ToolBarButton::get_PartOfGroup(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(button.fsStyle & BTNS_GROUP);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_PartOfGroup(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(newValue == VARIANT_FALSE) {
			button.fsStyle &= ~BTNS_GROUP;
		} else {
			button.fsStyle |= BTNS_GROUP;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_Pushed(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STATE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(button.fsState & TBSTATE_PRESSED);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_Pushed(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STATE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != VARIANT_FALSE) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		if(newValue == VARIANT_FALSE) {
			button.fsState &= ~TBSTATE_PRESSED;
		} else {
			button.fsState |= TBSTATE_PRESSED;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_SelectionState(SelectionStateConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SelectionStateConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STATE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		switch(button.fsState & (TBSTATE_CHECKED | TBSTATE_INDETERMINATE)) {
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
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_SelectionState(SelectionStateConstants newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STATE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != ssUnchecked) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		button.fsState &= ~(TBSTATE_CHECKED | TBSTATE_INDETERMINATE);
		switch(newValue) {
			case ssChecked:
				button.fsState |= TBSTATE_CHECKED;
				break;
			case ssIndeterminate:
				button.fsState |= TBSTATE_INDETERMINATE;
				break;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_ShouldBeDisplayedInChevronPopup(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.partOfChevronPopup) {
		*pValue = VARIANT_TRUE;
		return S_OK;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	BOOL isVertical = ((CWindow(hWndTBar).GetStyle() & CCS_VERT) == CCS_VERT);

	RECT availableRectangle = {0};
	GetClientRect(hWndTBar, &availableRectangle);
	CRect buttonRectangle;
	if(SendMessage(hWndTBar, TB_GETITEMRECT, properties.buttonIndex, reinterpret_cast<LPARAM>(&buttonRectangle))) {
		if(isVertical) {
			*pValue = BOOL2VARIANTBOOL(buttonRectangle.bottom > availableRectangle.bottom);
		} else {
			*pValue = BOOL2VARIANTBOOL(buttonRectangle.right > availableRectangle.right);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_ShowingEllipsis(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STATE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(button.fsState & TBSTATE_ELLIPSES);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_ShowingEllipsis(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STATE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != VARIANT_FALSE) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		if(newValue == VARIANT_FALSE) {
			button.fsState &= ~TBSTATE_ELLIPSES;
		} else {
			button.fsState |= TBSTATE_ELLIPSES;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		int textLength = static_cast<int>(SendMessage(hWndTBar, TB_GETBUTTONTEXT, button.idCommand, NULL));
		if(textLength > -1) {
			LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
			if(!pBuffer) {
				return E_OUTOFMEMORY;
			}
			ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
			if(static_cast<int>(SendMessage(hWndTBar, TB_GETBUTTONTEXT, button.idCommand, reinterpret_cast<LPARAM>(pBuffer))) > -1) {
				*pValue = _bstr_t(pBuffer).Detach();
				hr = S_OK;
			}
			HeapFree(GetProcessHeap(), 0, pBuffer);
		} else {
			*pValue = _bstr_t(_T("")).Detach();
			hr = S_OK;
		}
	}
	return hr;
}

STDMETHODIMP ToolBarButton::put_Text(BSTR newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_TEXT | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue && lstrlenW(newValue) > 0) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
	}
	COLE2T converter(newValue);
	if(newValue) {
		button.pszText = converter;
	}
	if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_UseMnemonic(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(!(button.fsStyle & BTNS_NOPREFIX));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_UseMnemonic(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STYLE | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand) && newValue != VARIANT_FALSE) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
		button.dwMask &= ~TBIF_COMMAND;
		if(newValue == VARIANT_FALSE) {
			button.fsStyle |= BTNS_NOPREFIX;
		} else {
			button.fsStyle &= ~BTNS_NOPREFIX;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_Visible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_STATE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = BOOL2VARIANTBOOL(!(button.fsState & TBSTATE_HIDDEN));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_Visible(VARIANT_BOOL newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_STATE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(newValue == VARIANT_FALSE) {
			button.fsState |= TBSTATE_HIDDEN;
		} else {
			button.fsState &= ~TBSTATE_HIDDEN;
		}
		if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::get_Width(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_SIZE;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		*pValue = button.cx;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::put_Width(OLE_XSIZE_PIXELS newValue)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_SIZE;
	button.cx = static_cast<WORD>(newValue);
	if(SendMessage(hWndTBar, TB_SETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button))) {
		return S_OK;
	}
	return E_FAIL;
}


STDMETHODIMP ToolBarButton::GetContainedWindow(HAlignmentConstants* pHorizontalAlignment/* = NULL*/, VAlignmentConstants* pVerticalAlignment/* = NULL*/, OLE_HANDLE* pHWnd/* = NULL*/)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->IsPlaceholderButton(button.idCommand)) {
			HAlignmentConstants horizontalAlignment;
			VAlignmentConstants verticalAlignment;
			*pHWnd = HandleToLong(properties.pOwnerTBar->GetPlaceholderButtonChildWindow(button.idCommand, horizontalAlignment, verticalAlignment));
			if(pHorizontalAlignment) {
				*pHorizontalAlignment = horizontalAlignment;
			}
			if(pVerticalAlignment) {
				*pVerticalAlignment = verticalAlignment;
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBarButton::GetRectangle(ButtonRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	HRESULT hr = E_FAIL;
	RECT rc = {0};
	if(rectangleType == brtButton) {
		if(SendMessage(hWndTBar, TB_GETITEMRECT, properties.buttonIndex, reinterpret_cast<LPARAM>(&rc))) {
			hr = S_OK;
		}
	} else if(rectangleType == brtDropDown) {
		if(properties.pOwnerTBar->IsComctl32Version610OrNewer()) {
			if(SendMessage(hWndTBar, TB_GETITEMDROPDOWNRECT, properties.buttonIndex, reinterpret_cast<LPARAM>(&rc))) {
				hr = S_OK;
			}
		} else {
			if(SendMessage(hWndTBar, TB_GETITEMRECT, properties.buttonIndex, reinterpret_cast<LPARAM>(&rc))) {
				rc.left = rc.right - GetSystemMetrics(SM_CYMENUCHECK);
				hr = S_OK;
			}
		}
	}
	if(SUCCEEDED(hr)) {
		if(pXLeft) {
			*pXLeft = rc.left;
		}
		if(pYTop) {
			*pYTop = rc.top;
		}
		if(pXRight) {
			*pXRight = rc.right;
		}
		if(pYBottom) {
			*pYBottom = rc.bottom;
		}
	}
	return hr;
}

STDMETHODIMP ToolBarButton::SetContainedWindow(OLE_HANDLE hWnd, HAlignmentConstants horizontalAlignment/* = halCenter*/, VAlignmentConstants verticalAlignment/* = valCenter*/)
{
	HWND hWndTBar = properties.GetTBarHWnd();
	if(!IsWindow(hWndTBar)) {
		return E_FAIL;
	}

	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_COMMAND;
	if(SendMessage(hWndTBar, TB_GETBUTTONINFO, properties.buttonIndex, reinterpret_cast<LPARAM>(&button)) == properties.buttonIndex) {
		if(properties.pOwnerTBar->SetPlaceholderButtonChildWindow(button.idCommand, static_cast<HWND>(LongToHandle(hWnd)), horizontalAlignment, verticalAlignment)) {
			return S_OK;
		}
	}
	return E_FAIL;
}