// ControlHostWindow.cpp: Creates a dialog for hosting a free-floating control.

#include "stdafx.h"
#include "ControlHostWindow.h"


HHOOK hCallWndProcHook = NULL;
volatile UINT controlHostWindowsCount = 0;


ControlHostWindow::ControlHostWindow() :
    freeFloatDialog(this, 1),
    parentWindow(this, 2)
{
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ControlHostWindow::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IControlHostWindow, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ControlHostWindow::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(properties.hostDialog));
	return S_OK;
}

STDMETHODIMP ControlHostWindow::get_hWndChild(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.hostDialog.IsWindow()) {
		*pValue = HandleToLong(FindWindowEx(properties.hostDialog, NULL, NULL, NULL));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ControlHostWindow::put_hWndChild(OLE_HANDLE newValue)
{
	if(properties.hostDialog.IsWindow()) {
		CWindow childWindow(static_cast<HWND>(LongToHandle(newValue)));
		if(childWindow.IsWindow()) {
			properties.hWndOldParent = childWindow.GetParent();

			RECT windowRectangle;
			childWindow.GetWindowRect(&windowRectangle);
			TCHAR pClassName[300];
			::GetClassName(childWindow, pClassName, 300);
			if(lstrcmpi(pClassName, TEXT("ToolBarU")) == 0 || lstrcmpi(pClassName, TEXT("ToolBarA")) == 0 || lstrcmpi(pClassName, TOOLBARCLASSNAME) == 0) {
				SIZE idealSize = {0};
				if(childWindow.SendMessage(TB_GETIDEALSIZE, FALSE, reinterpret_cast<LPARAM>(&idealSize)) && idealSize.cx > 0) {
					windowRectangle.right = windowRectangle.left + idealSize.cx;
				}
				if(childWindow.SendMessage(TB_GETIDEALSIZE, TRUE, reinterpret_cast<LPARAM>(&idealSize)) && idealSize.cy > 0) {
					windowRectangle.bottom = windowRectangle.top + idealSize.cy;
				}
			}
			AdjustWindowRectEx(&windowRectangle, properties.hostDialog.GetStyle(), (properties.hostDialog.GetMenu() != NULL), properties.hostDialog.GetExStyle());

			properties.hostDialog.SetWindowPos(NULL, &windowRectangle, SWP_NOREPOSITION | SWP_NOZORDER);
			childWindow.SetParent(properties.hostDialog);
			childWindow.GetWindowRect(&windowRectangle);
			childWindow.SetWindowPos(NULL, 0, 0, windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, SWP_NOZORDER);
		} else {
			childWindow.Attach(FindWindowEx(properties.hostDialog, NULL, NULL, NULL));
			if(childWindow.IsWindow()) {
				childWindow.SetParent(properties.hWndOldParent);
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ControlHostWindow::get_Resizeable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.hostDialog.IsWindow()) {
		properties.resizeable = !(properties.hostDialog.GetStyle() & DS_MODALFRAME);
	}

	*pValue = BOOL2VARIANTBOOL(properties.resizeable);
	return S_OK;
}

STDMETHODIMP ControlHostWindow::put_Resizeable(VARIANT_BOOL newValue)
{
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.resizeable != b) {
		properties.resizeable = b;

		if(properties.hostDialog.IsWindow()) {
			RECT windowRectangle;
			properties.hostDialog.GetClientRect(&windowRectangle);
			if(properties.resizeable) {
				properties.hostDialog.ModifyStyle(DS_MODALFRAME, WS_THICKFRAME);
			} else {
				properties.hostDialog.ModifyStyle(WS_THICKFRAME, DS_MODALFRAME);
			}
			AdjustWindowRectEx(&windowRectangle, properties.hostDialog.GetStyle(), (properties.hostDialog.GetMenu() != NULL), properties.hostDialog.GetExStyle());
			properties.hostDialog.SetWindowPos(NULL, &windowRectangle, SWP_DRAWFRAME | SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOREPOSITION | SWP_NOZORDER);
		}
	}
	return S_OK;
}

STDMETHODIMP ControlHostWindow::get_Visible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.hostDialog.IsWindow()) {
		properties.visible = properties.hostDialog.IsWindowVisible();
	}

	*pValue = BOOL2VARIANTBOOL(properties.visible);
	return S_OK;
}

STDMETHODIMP ControlHostWindow::put_Visible(VARIANT_BOOL newValue)
{
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.visible != b) {
		properties.visible = b;

		if(properties.hostDialog.IsWindow()) {
			if(properties.visible) {
				properties.hostDialog.ShowWindow(SW_SHOWNORMAL);
			} else {
				properties.hostDialog.ShowWindow(SW_HIDE);
			}
		}
	}
	return S_OK;
}

STDMETHODIMP ControlHostWindow::get_WindowTitle(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.hostDialog.IsWindow()) {
		properties.hostDialog.GetWindowText(&properties.windowTitle);
	}

	*pValue = properties.windowTitle.Copy();
	return S_OK;
}

STDMETHODIMP ControlHostWindow::put_WindowTitle(BSTR newValue)
{
	if(properties.windowTitle != newValue) {
		properties.windowTitle.AssignBSTR(newValue);

		if(properties.hostDialog.IsWindow()) {
			properties.hostDialog.SetWindowText(COLE2CT(newValue));
		}
	}
	return S_OK;
}


STDMETHODIMP ControlHostWindow::CalculateWindowSize(long clientWidth/* = 0x80000000*/, long clientHeight/* = 0x80000000*/, OLE_XSIZE_PIXELS* pWindowWidth/* = NULL*/, OLE_YSIZE_PIXELS* pWindowHeight/* = NULL*/)
{
	if(properties.hostDialog.IsWindow()) {
		if(clientWidth == 0x80000000 || clientHeight == 0x80000000) {
			RECT clientRectangle = {0};
			properties.hostDialog.GetClientRect(&clientRectangle);
			if(clientWidth == 0x80000000) {
				clientWidth = (clientRectangle.right - clientRectangle.left);
			}
			if(clientHeight == 0x80000000) {
				clientHeight = (clientRectangle.bottom - clientRectangle.top);
			}
		}
		RECT windowRectangle = {0, 0, clientWidth, clientHeight};
		AdjustWindowRectEx(&windowRectangle, properties.hostDialog.GetStyle(), properties.hostDialog.GetMenu() != NULL, properties.hostDialog.GetExStyle());
		if(pWindowWidth) {
			*pWindowWidth = windowRectangle.right - windowRectangle.left;
		}
		if(pWindowHeight) {
			*pWindowHeight = windowRectangle.bottom - windowRectangle.top;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ControlHostWindow::Create(OLE_HANDLE hWndParent)
{
	properties.hostDialog.Create(static_cast<HWND>(LongToHandle(hWndParent)));
	if(properties.hostDialog.IsWindow()) {
		if(!hCallWndProcHook) {
			hCallWndProcHook = SetWindowsHookEx(WH_CALLWNDPROC, ControlHostWindow::CallWndProc, NULL, GetCurrentThreadId());
			ATLASSERT(hCallWndProcHook);
		}
		++controlHostWindowsCount;

		if(properties.resizeable) {
			properties.hostDialog.ModifyStyle(DS_MODALFRAME, WS_THICKFRAME);
		} else {
			properties.hostDialog.ModifyStyle(WS_THICKFRAME, DS_MODALFRAME);
		}
		properties.hostDialog.SetWindowText(COLE2CT(properties.windowTitle));

		parentWindow.SubclassWindow(static_cast<HWND>(LongToHandle(hWndParent)));
		freeFloatDialog.SubclassWindow(properties.hostDialog);

		Raise_Created(HandleToLong(static_cast<HWND>(properties.hostDialog)));

		if(properties.visible) {
			properties.hostDialog.ShowWindow(SW_SHOWNORMAL);
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ControlHostWindow::Destroy(void)
{
	if(properties.hostDialog.IsWindow() && !FindWindowEx(properties.hostDialog, NULL, NULL, NULL)) {
		properties.hostDialog.DestroyWindow();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ControlHostWindow::GetClientRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(properties.hostDialog.IsWindow()) {
		RECT clientRectangle;
		properties.hostDialog.GetClientRect(&clientRectangle);
		if(pXLeft) {
			*pXLeft = clientRectangle.left;
		}
		if(pYTop) {
			*pYTop = clientRectangle.top;
		}
		if(pXRight) {
			*pXRight = clientRectangle.right;
		}
		if(pYBottom) {
			*pYBottom = clientRectangle.bottom;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ControlHostWindow::GetRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	if(properties.hostDialog.IsWindow()) {
		RECT windowRectangle;
		properties.hostDialog.GetWindowRect(&windowRectangle);
		if(pXLeft) {
			*pXLeft = windowRectangle.left;
		}
		if(pYTop) {
			*pYTop = windowRectangle.top;
		}
		if(pXRight) {
			*pXRight = windowRectangle.right;
		}
		if(pYBottom) {
			*pYBottom = windowRectangle.bottom;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ControlHostWindow::Move(long left/* = 0x80000000*/, long top/* = 0x80000000*/, long width/* = 0x80000000*/, long height/* = 0x80000000*/)
{
	if(properties.hostDialog.IsWindow()) {
		RECT windowRectangle;
		properties.hostDialog.GetWindowRect(&windowRectangle);
		RECT newWindowRectangle;
		newWindowRectangle.left = (left != 0x80000000 ? left : windowRectangle.left);
		newWindowRectangle.top = (top != 0x80000000 ? top : windowRectangle.top);
		newWindowRectangle.right = newWindowRectangle.left + (width != 0x80000000 ? width : windowRectangle.right - windowRectangle.left);
		newWindowRectangle.bottom = newWindowRectangle.top + (height != 0x80000000 ? height : windowRectangle.bottom - windowRectangle.top);
		properties.hostDialog.MoveWindow(&newWindowRectangle);
		return S_OK;
	}
	return E_FAIL;
}


LRESULT CALLBACK ControlHostWindow::CallWndProc(int code, WPARAM wParam, LPARAM lParam)
{
	LRESULT lr = 0;
	BOOL consumeMessage = FALSE;

	if(code >= 0) {
		LPCWPSTRUCT pMessage = reinterpret_cast<LPCWPSTRUCT>(lParam);
		switch(pMessage->message) {
			case WM_MOUSEACTIVATE:
			{
				HWND hWndLosingFocus = GetParent(GetFocus());
				ControlHostWindow* pCtlHostWindow = NULL;
				while(hWndLosingFocus) {
					pCtlHostWindow = reinterpret_cast<ControlHostWindow*>(SendMessage(hWndLosingFocus, GetControlHostWindowMessage(), 0, 0));
					if(pCtlHostWindow) {
						/* When activating a COM container window, COM sets the focus to the currently active contained
						   control. Since COM does not know that this control is placed in another window now, we need to
						   trick COM: We disable the free-floating dialog that hosts the control. It will be enabled again
						   after the COM container window received the focus. */
						ATLASSERT(hWndLosingFocus == pCtlHostWindow->properties.hostDialog);
						if(reinterpret_cast<HWND>(pMessage->wParam) == pCtlHostWindow->parentWindow) {
							consumeMessage = TRUE;
							lr = CallNextHookEx(hCallWndProcHook, code, wParam, lParam);
							pCtlHostWindow->properties.hostDialog.EnableWindow(FALSE);
							SetProp(pCtlHostWindow->properties.hostDialog, TEXT("EnableMe"), reinterpret_cast<HANDLE>(TRUE));
							break;
						}
						break;
					}
					hWndLosingFocus = GetParent(hWndLosingFocus);
				}
				break;
			}
		}
	}

	if(!consumeMessage) {
		lr = CallNextHookEx(hCallWndProcHook, code, wParam, lParam);
	}
	return lr;
}


LRESULT ControlHostWindow::OnHostDialogActivate(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(LOWORD(wParam) == WA_INACTIVE) {
		Raise_Deactivate(HandleToLong(reinterpret_cast<HWND>(lParam)));
	} else {
		Raise_Activate(BOOL2VARIANTBOOL(LOWORD(wParam) == WA_CLICKACTIVE), HandleToLong(reinterpret_cast<HWND>(lParam)));
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ControlHostWindow::OnHostDialogClose(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	VARIANT_BOOL cancel = VARIANT_FALSE;
	Raise_Closing(&cancel);
	if(cancel == VARIANT_FALSE) {
		properties.hostDialog.DestroyWindow();
	}
	return 0;
}

LRESULT ControlHostWindow::OnHostDialogDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(IsWindow(properties.hWndOldParent)) {
		HWND hWndChild = FindWindowEx(properties.hostDialog, NULL, NULL, NULL);
		if(hWndChild) {
			ShowWindow(hWndChild, SW_HIDE);
			SetParent(hWndChild, properties.hWndOldParent);
		}
	}

	freeFloatDialog.UnsubclassWindow();
	parentWindow.UnsubclassWindow();
	RemoveProp(properties.hostDialog, _T("EnableMe"));

	--controlHostWindowsCount;
	if(controlHostWindowsCount == 0 && hCallWndProcHook) {
		UnhookWindowsHookEx(hCallWndProcHook);
		hCallWndProcHook = NULL;
	}

	Raise_Destroyed(HandleToLong(static_cast<HWND>(properties.hostDialog)));

	wasHandled = FALSE;
	return 0;
}

LRESULT ControlHostWindow::OnHostDialogNCActivate(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam) {
		// we want the parent window to appear activated even though it isn't
		parentWindow.SendMessage(WM_NCACTIVATE, TRUE, 0);

		if(IsWindow(properties.hWndOldParent)) {
			TCHAR pClassName[300];
			GetClassName(properties.hWndOldParent, pClassName, 300);
			if(lstrcmpi(pClassName, TEXT("ReBarU")) == 0 || lstrcmpi(pClassName, TEXT("ReBarA")) == 0 || lstrcmpi(pClassName, REBARCLASSNAME) == 0) {
				// redraw band captions
				InvalidateRect(properties.hWndOldParent, NULL, TRUE);
				// redraw the bands' children
				RedrawWindow(properties.hWndOldParent, NULL, NULL, RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
			}
		}
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ControlHostWindow::OnHostDialogNCLButtonDblClk(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_TitleDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	wasHandled = FALSE;
	return 0;
}

LRESULT ControlHostWindow::OnHostDialogSysCommand(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(wParam == SC_CLOSE && lParam == 0) {
		/* TODO: We want to forward Alt+F4 only, but selecting the system menu command by keyboard ends up
		         here, too. */
		return parentWindow.SendMessage(message, wParam, lParam);
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ControlHostWindow::OnHostDialogWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);
	if(!(pDetails->flags & SWP_NOMOVE)) {
		Raise_Moved(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		CWindow childWindow(FindWindowEx(properties.hostDialog, NULL, NULL, NULL));
		if(childWindow.IsWindow()) {
			RECT clientRectangle = {0};
			properties.hostDialog.GetClientRect(&clientRectangle);
			childWindow.MoveWindow(clientRectangle.left, clientRectangle.top, clientRectangle.right - clientRectangle.left, clientRectangle.bottom - clientRectangle.top);
		}
		Raise_Resized(pDetails->cx, pDetails->cy);
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ControlHostWindow::OnHostDialogWindowPosChanging(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);
	if(!(pDetails->flags & (SWP_NOMOVE | SWP_NOSIZE))) {
		RECT windowRectangle;
		properties.hostDialog.GetWindowRect(&windowRectangle);

		if(!(pDetails->flags & SWP_NOMOVE) && (pDetails->x != windowRectangle.left || pDetails->y != windowRectangle.top)) {
			VARIANT_BOOL cancel = VARIANT_FALSE;
			Raise_Moving(reinterpret_cast<OLE_XPOS_PIXELS*>(&pDetails->x), reinterpret_cast<OLE_YPOS_PIXELS*>(&pDetails->y), &cancel);
			if(cancel != VARIANT_FALSE) {
				pDetails->flags |= SWP_NOMOVE;
			}
		}
		if(!(pDetails->flags & SWP_NOSIZE) && (pDetails->cx != windowRectangle.right - windowRectangle.left || pDetails->cy != windowRectangle.bottom - windowRectangle.top)) {
			VARIANT_BOOL cancel = VARIANT_FALSE;
			Raise_Resizing(reinterpret_cast<OLE_XSIZE_PIXELS*>(&pDetails->cx), reinterpret_cast<OLE_YSIZE_PIXELS*>(&pDetails->cy), &cancel);
			if(cancel != VARIANT_FALSE) {
				pDetails->flags |= SWP_NOSIZE;
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ControlHostWindow::OnHostDialogGetControlHostWindow(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	if(properties.hostDialog.IsWindowVisible()) {
		return reinterpret_cast<LRESULT>(this);
	}
	return NULL;
}

LRESULT ControlHostWindow::OnParentSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = parentWindow.DefWindowProc(message, wParam, lParam);
	if(GetProp(properties.hostDialog, _T("EnableMe"))) {
		RemoveProp(properties.hostDialog, _T("EnableMe"));
		properties.hostDialog.EnableWindow(TRUE);
	}
	return lr;
}


inline HRESULT ControlHostWindow::Raise_Activate(VARIANT_BOOL activatedByMouseClick, LONG hWndBeingDeactivated)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Activate(activatedByMouseClick, hWndBeingDeactivated);
	//}
	//return S_OK;
}

inline HRESULT ControlHostWindow::Raise_Closing(VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Closing(pCancel);
	//}
	//return S_OK;
}

inline HRESULT ControlHostWindow::Raise_Created(LONG hWnd)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Created(hWnd);
	//}
	//return S_OK;
}

inline HRESULT ControlHostWindow::Raise_Deactivate(LONG hWndBeingActivated)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Deactivate(hWndBeingActivated);
	//}
	//return S_OK;
}

inline HRESULT ControlHostWindow::Raise_Destroyed(LONG hWnd)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Destroyed(hWnd);
	//}
	//return S_OK;
}

inline HRESULT ControlHostWindow::Raise_Moved(OLE_XPOS_PIXELS left, OLE_YPOS_PIXELS top)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Moved(left, top);
	//}
	//return S_OK;
}

inline HRESULT ControlHostWindow::Raise_Moving(OLE_XPOS_PIXELS* pLeft, OLE_YPOS_PIXELS* pTop, VARIANT_BOOL* pPreventMove)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Moving(pLeft, pTop, pPreventMove);
	//}
	//return S_OK;
}

inline HRESULT ControlHostWindow::Raise_Resized(OLE_XSIZE_PIXELS width, OLE_YSIZE_PIXELS height)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Resized(width, height);
	//}
	//return S_OK;
}

inline HRESULT ControlHostWindow::Raise_Resizing(OLE_XSIZE_PIXELS* pWidth, OLE_YSIZE_PIXELS* pHeight, VARIANT_BOOL* pPreventResize)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_Resizing(pWidth, pHeight, pPreventResize);
	//}
	//return S_OK;
}

inline HRESULT ControlHostWindow::Raise_TitleDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TitleDblClick(button, shift, x, y);
	//}
	//return S_OK;
}