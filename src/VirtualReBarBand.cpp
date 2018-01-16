// VirtualReBarBand.cpp: A wrapper for non-existing rebar bands.

#include "stdafx.h"
#include "VirtualReBarBand.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP VirtualReBarBand::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IVirtualReBarBand, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


VirtualReBarBand::Properties::~Properties()
{
	if(settings.lpText != LPSTR_TEXTCALLBACK) {
		SECUREFREE(settings.lpText);
	}
	if(pOwnerRBar) {
		pOwnerRBar->Release();
	}
}

HWND VirtualReBarBand::Properties::GetRBarHWnd(void)
{
	ATLASSUME(pOwnerRBar);

	OLE_HANDLE handle = NULL;
	pOwnerRBar->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT VirtualReBarBand::LoadState(LPREBARBANDINFO pSource, int bandIndex)
{
	ATLASSERT_POINTER(pSource, REBARBANDINFO);

	SECUREFREE(properties.settings.lpText);
	properties.settings = *pSource;
	if(properties.settings.fMask & RBBIM_TEXT) {
		// duplicate the band's text
		if(properties.settings.lpText != LPSTR_TEXTCALLBACK) {
			properties.settings.cch = lstrlen(pSource->lpText);
			properties.settings.lpText = reinterpret_cast<LPTSTR>(malloc((properties.settings.cch + 1) * sizeof(TCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopy(properties.settings.lpText, properties.settings.cch + 1, pSource->lpText)));
		}
	}
	properties.bandIndex = bandIndex;

	return S_OK;
}

HRESULT VirtualReBarBand::SaveState(LPREBARBANDINFO pTarget, int& bandIndex)
{
	ATLASSERT_POINTER(pTarget, REBARBANDINFO);

	SECUREFREE(pTarget->lpText);
	*pTarget = properties.settings;
	if(pTarget->fMask & RBBIM_TEXT) {
		// duplicate the band's text
		if(pTarget->lpText != LPSTR_TEXTCALLBACK) {
			pTarget->lpText = reinterpret_cast<LPTSTR>(malloc((pTarget->cch + 1) * sizeof(TCHAR)));
			ATLASSERT(pTarget->lpText);
			if(pTarget->lpText) {
				ATLVERIFY(SUCCEEDED(StringCchCopy(pTarget->lpText, pTarget->cch + 1, properties.settings.lpText)));
			}
		}
	}
	bandIndex = properties.bandIndex;

	return S_OK;
}

void VirtualReBarBand::SetOwner(ReBar* pOwner)
{
	if(properties.pOwnerRBar) {
		properties.pOwnerRBar->Release();
	}
	properties.pOwnerRBar = pOwner;
	if(properties.pOwnerRBar) {
		properties.pOwnerRBar->AddRef();
	}
}


STDMETHODIMP VirtualReBarBand::get_AddMarginsAroundChild(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.fStyle & RBBS_CHILDEDGE);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	if(properties.settings.fMask & RBBIM_COLORS) {
		*pValue = (properties.settings.clrBack == CLR_DEFAULT ? static_cast<COLORREF>(SendMessage(hWndRBar, RB_GETBKCOLOR, 0, 0)) : properties.settings.clrBack);
	} else {
		*pValue = static_cast<OLE_COLOR>(-1);
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_BandData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_LPARAM) {
		*pValue = properties.settings.lParam;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_ChevronButtonObjectState(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerRBar->IsComctl32Version610OrNewer()) {
		if(properties.settings.fMask & RBBIM_CHEVRONSTATE) {
			*pValue = properties.settings.uChevronState;
		} else {
			*pValue = STATE_SYSTEM_HASPOPUP | STATE_SYSTEM_INVISIBLE | STATE_SYSTEM_UNAVAILABLE;
		}
	}
	return E_FAIL;
}

STDMETHODIMP VirtualReBarBand::get_CurrentHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_CHILDSIZE) {
		*pValue = properties.settings.cyChild;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_CurrentWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_SIZE) {
		*pValue = properties.settings.cx;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_FixedBackgroundBitmapOrigin(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.fStyle & RBBS_FIXEDBMP);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	if(properties.settings.fMask & RBBIM_COLORS) {
		*pValue = (properties.settings.clrFore == CLR_DEFAULT ? static_cast<COLORREF>(SendMessage(hWndRBar, RB_GETTEXTCOLOR, 0, 0)) : properties.settings.clrFore);
	} else {
		*pValue = static_cast<OLE_COLOR>(-1);
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_hBackgroundBitmap(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_BACKGROUND) {
		*pValue = HandleToLong(properties.settings.hbmBack);
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_hContainedWindow(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_CHILD) {
		*pValue = HandleToLong(properties.settings.hwndChild);
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_HeightChangeStepSize(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_CHILDSIZE) {
		*pValue = properties.settings.cyIntegral;
	} else {
		*pValue = 1;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_HideIfVertical(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.fStyle & RBBS_NOVERT);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_IMAGE) {
		*pValue = (properties.settings.iImage == -1 ? -2 : properties.settings.iImage);
	} else {
		*pValue = -2;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_ID) {
		*pValue = properties.settings.wID;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_IdealWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_IDEALSIZE) {
		*pValue = properties.settings.cxIdeal;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.bandIndex;
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_KeepInFirstRow(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.fStyle & RBBS_TOPALIGN);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_MaximumHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_CHILDSIZE) {
		*pValue = properties.settings.cyMaxChild;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_MinimumHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_CHILDSIZE) {
		*pValue = properties.settings.cyMinChild;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_MinimumWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_CHILDSIZE) {
		*pValue = properties.settings.cxMinChild;
	} else {
		*pValue = 0;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_NewLine(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.fStyle & RBBS_BREAK);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_Resizable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(!(properties.settings.fStyle & RBBS_FIXEDSIZE));
	} else {
		*pValue = VARIANT_TRUE;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_ShowTitle(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(!(properties.settings.fStyle & RBBS_HIDETITLE));
	} else {
		*pValue = VARIANT_TRUE;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_SizingGripVisibility(SizingGripVisibilityConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SizingGripVisibilityConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		if(properties.settings.fStyle & RBBS_NOGRIPPER) {
			*pValue = sgvNever;
		} else if(properties.settings.fStyle & RBBS_GRIPPERALWAYS) {
			*pValue = sgvAlways;
		} else {
			*pValue = sgvAutomatic;
		}
	} else {
		*pValue = sgvAutomatic;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_TEXT) {
		*pValue = _bstr_t(properties.settings.lpText).Detach();
	} else {
		*pValue = SysAllocString(L"");
	}

	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_TitleWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_HEADERSIZE) {
		*pValue = properties.settings.cxHeader;
	} else {
		*pValue = -1;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_UseChevron(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.fStyle & RBBS_USECHEVRON);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_VariableHeight(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(properties.settings.fStyle & RBBS_VARIABLEHEIGHT);
	} else {
		*pValue = VARIANT_FALSE;
	}
	return S_OK;
}

STDMETHODIMP VirtualReBarBand::get_Visible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.settings.fMask & RBBIM_STYLE) {
		*pValue = BOOL2VARIANTBOOL(!(properties.settings.fStyle & RBBS_HIDDEN));
	} else {
		*pValue = VARIANT_TRUE;
	}
	return S_OK;
}


STDMETHODIMP VirtualReBarBand::GetChevronRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(properties.pOwnerRBar->IsComctl32Version610OrNewer()) {
		if(properties.settings.fMask & RBBIM_CHEVRONLOCATION) {
			if(pXLeft) {
				*pXLeft = properties.settings.rcChevronLocation.left;
			}
			if(pYTop) {
				*pYTop = properties.settings.rcChevronLocation.top;
			}
			if(pXRight) {
				*pXRight = properties.settings.rcChevronLocation.right;
			}
			if(pYBottom) {
				*pYBottom = properties.settings.rcChevronLocation.bottom;
			}
		} else {
			if(pXLeft) {
				*pXLeft = 0;
			}
			if(pYTop) {
				*pYTop = 0;
			}
			if(pXRight) {
				*pXRight = 0;
			}
			if(pYBottom) {
				*pYBottom = 0;
			}
		}
		hr = S_OK;
	}
	return hr;
}