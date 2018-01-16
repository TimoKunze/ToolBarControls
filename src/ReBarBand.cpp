// ReBarBand.cpp: A wrapper for existing rebar bands.

#include "stdafx.h"
#include "ReBarBand.h"
#include "ClassFactory.h"


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ReBarBand::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IReBarBand, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


ReBarBand::Properties::~Properties()
{
	if(pOwnerRBar) {
		pOwnerRBar->Release();
	}
}

HWND ReBarBand::Properties::GetRBarHWnd(void)
{
	ATLASSUME(pOwnerRBar);

	OLE_HANDLE handle = NULL;
	pOwnerRBar->get_hWnd(&handle);
	return static_cast<HWND>(LongToHandle(handle));
}


HRESULT ReBarBand::SaveState(int bandIndex, LPREBARBANDINFO pTarget, HWND hWndRBar/* = NULL*/)
{
	if(!hWndRBar) {
		hWndRBar = properties.GetRBarHWnd();
	}
	ATLASSERT(IsWindow(hWndRBar));
	if((bandIndex < 0) || (bandIndex >= static_cast<int>(SendMessage(hWndRBar, RB_GETBANDCOUNT, 0, 0)))) {
		return E_INVALIDARG;
	}
	ATLASSERT_POINTER(pTarget, REBARBANDINFO);
	if(!pTarget) {
		return E_POINTER;
	}

	ZeroMemory(pTarget, sizeof(REBARBANDINFO));
	pTarget->cch = MAX_BANDTEXTLENGTH;
	pTarget->lpText = reinterpret_cast<LPTSTR>(malloc((pTarget->cch + 1) * sizeof(TCHAR)));
	pTarget->fMask = RBBIM_BACKGROUND | RBBIM_CHILD | RBBIM_CHILDSIZE | RBBIM_COLORS | RBBIM_HEADERSIZE | RBBIM_IDEALSIZE | RBBIM_ID | RBBIM_IMAGE | RBBIM_LPARAM | RBBIM_SIZE | RBBIM_TEXT | RBBIM_CHEVRONLOCATION | RBBIM_CHEVRONSTATE;

	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(pTarget))) {
		return S_OK;
	}
	return E_FAIL;
}


void ReBarBand::Attach(int bandIndex)
{
	properties.bandIndex = bandIndex;
}

void ReBarBand::Detach(void)
{
	properties.bandIndex = -1;
}

HRESULT ReBarBand::LoadState(LPREBARBANDINFO pSource)
{
	ATLASSERT_POINTER(pSource, REBARBANDINFO);
	if(!pSource) {
		return E_POINTER;
	}

	SizingGripVisibilityConstants sgv;
	if(pSource->fMask & RBBIM_STYLE) {
		put_AddMarginsAroundChild(BOOL2VARIANTBOOL(pSource->fStyle & RBBS_CHILDEDGE));
		put_FixedBackgroundBitmapOrigin(BOOL2VARIANTBOOL(pSource->fStyle & RBBS_FIXEDBMP));
		put_HideIfVertical(BOOL2VARIANTBOOL(pSource->fStyle & RBBS_NOVERT));
		put_KeepInFirstRow(BOOL2VARIANTBOOL(pSource->fStyle & RBBS_TOPALIGN));
		put_NewLine(BOOL2VARIANTBOOL(pSource->fStyle & RBBS_BREAK));
		put_Resizable(BOOL2VARIANTBOOL(!(pSource->fStyle & RBBS_FIXEDSIZE)));
		put_ShowTitle(BOOL2VARIANTBOOL(!(pSource->fStyle & RBBS_HIDETITLE)));
		put_UseChevron(BOOL2VARIANTBOOL(pSource->fStyle & RBBS_USECHEVRON));
		put_VariableHeight(BOOL2VARIANTBOOL(pSource->fStyle & RBBS_VARIABLEHEIGHT));
		put_Visible(BOOL2VARIANTBOOL(!(pSource->fStyle & RBBS_HIDDEN)));

		if(pSource->fStyle & RBBS_NOGRIPPER) {
			sgv = sgvNever;
		} else if(pSource->fStyle & RBBS_GRIPPERALWAYS) {
			sgv = sgvAlways;
		} else {
			sgv = sgvAutomatic;
		}
		put_SizingGripVisibility(sgv);
	}

	OLE_COLOR clr;
	if(pSource->fMask & RBBIM_COLORS) {
		clr = pSource->clrBack;
		put_BackColor(clr);
		clr = pSource->clrFore;
		put_ForeColor(clr);
	}

	if(pSource->fMask & RBBIM_IMAGE) {
		put_IconIndex(pSource->iImage == -1 ? -2 : pSource->iImage);
	}
	if(pSource->fMask & RBBIM_LPARAM) {
		put_BandData(pSource->lParam);
	}
	if(pSource->fMask & RBBIM_HEADERSIZE) {
		put_TitleWidth(pSource->cxHeader);
	}
	if(pSource->fMask & RBBIM_IDEALSIZE) {
		put_IdealWidth(pSource->cxIdeal);
	}
	if(pSource->fMask & RBBIM_CHILDSIZE) {
		put_MaximumHeight(pSource->cyMaxChild);
		put_MinimumHeight(pSource->cyMinChild);
		put_HeightChangeStepSize(pSource->cyIntegral);
		put_CurrentHeight(pSource->cyChild);
		put_MinimumWidth(pSource->cxMinChild);
	}
	if(pSource->fMask & RBBIM_SIZE) {
		put_CurrentWidth(pSource->cx);
	}
	if(pSource->fMask & RBBIM_BACKGROUND) {
		put_hBackgroundBitmap(HandleToLong(pSource->hbmBack));
	}
	if(pSource->fMask & RBBIM_CHILD) {
		put_hContainedWindow(HandleToLong(pSource->hwndChild));
	}

	BSTR text = NULL;
	if(pSource->fMask & RBBIM_TEXT) {
		text = _bstr_t(pSource->lpText).Detach();
		put_Text(text);
	}
	SysFreeString(text);
	return S_OK;
}

HRESULT ReBarBand::LoadState(VirtualReBarBand* pSource)
{
	ATLASSUME(pSource);
	if(!pSource) {
		return E_POINTER;
	}

	VARIANT_BOOL b = VARIANT_FALSE;
	pSource->get_AddMarginsAroundChild(&b);
	put_AddMarginsAroundChild(b);
	b = VARIANT_FALSE;
	pSource->get_FixedBackgroundBitmapOrigin(&b);
	put_FixedBackgroundBitmapOrigin(b);
	b = VARIANT_FALSE;
	pSource->get_HideIfVertical(&b);
	put_HideIfVertical(b);
	b = VARIANT_FALSE;
	pSource->get_KeepInFirstRow(&b);
	put_KeepInFirstRow(b);
	b = VARIANT_FALSE;
	pSource->get_NewLine(&b);
	put_NewLine(b);
	b = VARIANT_TRUE;
	pSource->get_Resizable(&b);
	put_Resizable(b);
	b = VARIANT_TRUE;
	pSource->get_ShowTitle(&b);
	put_ShowTitle(b);
	b = VARIANT_FALSE;
	pSource->get_UseChevron(&b);
	put_UseChevron(b);
	b = VARIANT_FALSE;
	pSource->get_VariableHeight(&b);
	put_VariableHeight(b);
	b = VARIANT_TRUE;
	pSource->get_Visible(&b);
	put_Visible(b);

	COLORREF clr = CLR_DEFAULT;
	pSource->get_BackColor(&clr);
	put_BackColor(clr);
	clr = CLR_DEFAULT;
	pSource->get_ForeColor(&clr);
	put_ForeColor(clr);

	LONG l = 0;
	pSource->get_BandData(&l);
	put_BandData(l);
	l = -2;
	pSource->get_IconIndex(&l);
	put_IconIndex(l);

	OLE_YSIZE_PIXELS cy = 0;
	pSource->get_MaximumHeight(&cy);
	put_MaximumHeight(cy);
	cy = 0;
	pSource->get_MinimumHeight(&cy);
	put_MinimumHeight(cy);
	cy = 1;
	pSource->get_HeightChangeStepSize(&cy);
	put_HeightChangeStepSize(cy);
	cy = 0;
	pSource->get_CurrentHeight(&cy);
	put_CurrentHeight(cy);

	OLE_XSIZE_PIXELS cx = 0;
	pSource->get_IdealWidth(&cx);
	put_IdealWidth(cx);
	cx = 0;
	pSource->get_MinimumWidth(&cx);
	put_MinimumWidth(cx);
	cx = -1;
	pSource->get_TitleWidth(&cx);
	put_TitleWidth(cx);
	cx = 0;
	pSource->get_CurrentWidth(&cx);
	put_CurrentWidth(cx);

	OLE_HANDLE h = NULL;
	pSource->get_hBackgroundBitmap(&h);
	put_hBackgroundBitmap(h);
	h = NULL;
	pSource->get_hContainedWindow(&h);
	put_hContainedWindow(h);

	SizingGripVisibilityConstants sgv = sgvAutomatic;
	pSource->get_SizingGripVisibility(&sgv);
	put_SizingGripVisibility(sgv);

	BSTR text = NULL;
	pSource->get_Text(&text);
	put_Text(text);
	SysFreeString(text);
	return S_OK;
}

HRESULT ReBarBand::SaveState(LPREBARBANDINFO pTarget, HWND hWndRBar/* = NULL*/)
{
	return SaveState(properties.bandIndex, pTarget, hWndRBar);
}

HRESULT ReBarBand::SaveState(VirtualReBarBand* pTarget)
{
	ATLASSUME(pTarget);
	if(!pTarget) {
		return E_POINTER;
	}

	pTarget->SetOwner(properties.pOwnerRBar);
	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	HRESULT hr = SaveState(&band);
	pTarget->LoadState(&band, properties.bandIndex);
	SECUREFREE(band.lpText);

	return hr;
}

void ReBarBand::SetOwner(ReBar* pOwner)
{
	if(properties.pOwnerRBar) {
		properties.pOwnerRBar->Release();
	}
	properties.pOwnerRBar = pOwner;
	if(properties.pOwnerRBar) {
		properties.pOwnerRBar->AddRef();
	}
}


STDMETHODIMP ReBarBand::get_AddMarginsAroundChild(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(band.fStyle & RBBS_CHILDEDGE);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_AddMarginsAroundChild(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle &= ~RBBS_CHILDEDGE;
		} else {
			band.fStyle |= RBBS_CHILDEDGE;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_COLORS;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = (band.clrBack == CLR_DEFAULT ? static_cast<COLORREF>(SendMessage(hWndRBar, RB_GETBKCOLOR, 0, 0)) : band.clrBack);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_BackColor(OLE_COLOR newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_COLORS;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		band.clrBack = (newValue == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(newValue));
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_BandData(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_LPARAM;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = band.lParam;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_BandData(LONG newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_LPARAM;
	band.lParam = newValue;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_ChevronButtonObjectState(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerRBar->IsComctl32Version610OrNewer()) {
		HWND hWndRBar = properties.GetRBarHWnd();
		ATLASSERT(IsWindow(hWndRBar));

		REBARBANDINFO band = {0};
		band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
		band.fMask = RBBIM_CHEVRONSTATE;
		if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			*pValue = band.uChevronState;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_ChevronHot(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}
	
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	if(properties.pOwnerRBar->IsComctl32Version610OrNewer()) {
		REBARBANDINFO band = {0};
		band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
		band.fMask = RBBIM_CHEVRONSTATE;
		if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			*pValue = BOOL2VARIANTBOOL(band.uChevronState & STATE_SYSTEM_HOTTRACKED);
			return S_OK;
		}
	} else {
		VARIANT_BOOL visible = VARIANT_FALSE;
		if(SUCCEEDED(get_ChevronVisible(&visible))) {
			/* NOTE: To be compatible with accessibility state (RBBIM_CHEVRONSTATE), we DON'T report the chevron
			         as being hot, if the chevron is pushed. */
			if(visible != VARIANT_FALSE && properties.bandIndex != properties.pOwnerRBar->bandChevronPushed) {
				POINT mousePosition = {0};
				GetCursorPos(&mousePosition);
				ScreenToClient(hWndRBar, &mousePosition);

				REBARBANDINFO band = {0};
				if(properties.pOwnerRBar->IsComctl32Version610OrNewer()) {
					band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
				} else {
					band.cbSize = sizeof(REBARBANDINFO);
				}
				band.fMask = RBBIM_CHEVRONLOCATION;
				if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
					*pValue = BOOL2VARIANTBOOL(PtInRect(&band.rcChevronLocation, mousePosition));
				}
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_ChevronPushed(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerRBar->IsComctl32Version610OrNewer()) {
		HWND hWndRBar = properties.GetRBarHWnd();
		ATLASSERT(IsWindow(hWndRBar));

		REBARBANDINFO band = {0};
		band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
		band.fMask = RBBIM_CHEVRONSTATE;
		if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			*pValue = BOOL2VARIANTBOOL(band.uChevronState & STATE_SYSTEM_PRESSED);
			return S_OK;
		}
	} else {
		VARIANT_BOOL visible = VARIANT_FALSE;
		if(SUCCEEDED(get_ChevronVisible(&visible))) {
			if(visible != VARIANT_FALSE) {
				*pValue = BOOL2VARIANTBOOL(properties.bandIndex == properties.pOwnerRBar->bandChevronPushed);
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_ChevronVisible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	if(properties.pOwnerRBar->IsComctl32Version610OrNewer()) {
		band.fMask = RBBIM_CHEVRONSTATE;
		if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			*pValue = BOOL2VARIANTBOOL(!(band.uChevronState & STATE_SYSTEM_INVISIBLE));
			return S_OK;
		}
	} else {
		band.fMask = RBBIM_CHILD | RBBIM_IDEALSIZE | RBBIM_STYLE;
		if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			*pValue = VARIANT_FALSE;
			if((band.fStyle & (RBBS_USECHEVRON | RBBS_FIXEDSIZE)) == RBBS_USECHEVRON) {
				RECT containedWindowRect = {0};
				GetWindowRect(band.hwndChild, &containedWindowRect);
				if(MapWindowPoints(HWND_DESKTOP, hWndRBar, reinterpret_cast<LPPOINT>(&containedWindowRect), 2)) {
					if((CWindow(hWndRBar).GetStyle() & CCS_VERT) == CCS_VERT) {
						*pValue = BOOL2VARIANTBOOL(band.cxIdeal > static_cast<UINT>(containedWindowRect.bottom - containedWindowRect.top));
					} else {
						*pValue = BOOL2VARIANTBOOL(band.cxIdeal > static_cast<UINT>(containedWindowRect.right - containedWindowRect.left));
					}
				}
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_CurrentHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = band.cyChild;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_CurrentHeight(OLE_YSIZE_PIXELS newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	band.cyMinChild = static_cast<UINT>(-1);
	band.cyMaxChild = static_cast<UINT>(-1);
	band.cyIntegral = static_cast<UINT>(-1);
	band.cxMinChild = static_cast<UINT>(-1);
	band.cyChild = newValue;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_CurrentWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_SIZE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = band.cx;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_CurrentWidth(OLE_XSIZE_PIXELS newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_SIZE;
	band.cx = newValue;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_FixedBackgroundBitmapOrigin(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(band.fStyle & RBBS_FIXEDBMP);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_FixedBackgroundBitmapOrigin(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle &= ~RBBS_FIXEDBMP;
		} else {
			band.fStyle |= RBBS_FIXEDBMP;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_COLORS;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = (band.clrFore == CLR_DEFAULT ? static_cast<COLORREF>(SendMessage(hWndRBar, RB_GETTEXTCOLOR, 0, 0)) : band.clrFore);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_ForeColor(OLE_COLOR newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_COLORS;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		band.clrFore = (newValue == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(newValue));
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_hBackgroundBitmap(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_BACKGROUND;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = HandleToLong(band.hbmBack);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_hBackgroundBitmap(OLE_HANDLE newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_BACKGROUND;
	band.hbmBack = static_cast<HBITMAP>(LongToHandle(newValue));
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_hContainedWindow(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILD;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = HandleToLong(band.hwndChild);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_hContainedWindow(OLE_HANDLE newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILD;
	band.hwndChild = static_cast<HWND>(LongToHandle(newValue));
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_HeightChangeStepSize(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = band.cyIntegral;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_HeightChangeStepSize(OLE_YSIZE_PIXELS newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	band.cyMinChild = static_cast<UINT>(-1);
	band.cyMaxChild = static_cast<UINT>(-1);
	band.cyChild = static_cast<UINT>(-1);
	band.cxMinChild = static_cast<UINT>(-1);
	band.cyIntegral = newValue;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_HideIfVertical(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(band.fStyle & RBBS_NOVERT);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_HideIfVertical(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle &= ~RBBS_NOVERT;
		} else {
			band.fStyle |= RBBS_NOVERT;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_IconIndex(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_IMAGE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = (band.iImage == -1 ? -2 : band.iImage);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_IconIndex(LONG newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_IMAGE;
	band.iImage = (newValue == -2 ? -1 : newValue);
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_ID(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(properties.pOwnerRBar) {
		if((*pValue = properties.pOwnerRBar->BandIndexToID(properties.bandIndex)) != -1) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_IdealWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_IDEALSIZE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = band.cxIdeal;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_IdealWidth(OLE_XSIZE_PIXELS newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_IDEALSIZE;
	band.cxIdeal = newValue;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_Index(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.bandIndex;
	return S_OK;
}

STDMETHODIMP ReBarBand::put_Index(LONG newValue)
{
	if(newValue == properties.bandIndex) {
		return S_OK;
	}
	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	int numberOfBands = static_cast<int>(SendMessage(hWndRBar, RB_GETBANDCOUNT, 0, 0));
	if(newValue >= numberOfBands) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(SendMessage(hWndRBar, RB_MOVEBAND, properties.bandIndex, newValue)) {
		properties.bandIndex = newValue;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_KeepInFirstRow(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(band.fStyle & RBBS_TOPALIGN);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_KeepInFirstRow(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle &= ~RBBS_TOPALIGN;
		} else {
			band.fStyle |= RBBS_TOPALIGN;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_MaximumHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = band.cyMaxChild;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_MaximumHeight(OLE_YSIZE_PIXELS newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	band.cyMinChild = static_cast<UINT>(-1);
	band.cyChild = static_cast<UINT>(-1);
	band.cyIntegral = static_cast<UINT>(-1);
	band.cxMinChild = static_cast<UINT>(-1);
	band.cyMaxChild = newValue;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_MinimumHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = band.cyMinChild;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_MinimumHeight(OLE_YSIZE_PIXELS newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	band.cyMaxChild = static_cast<UINT>(-1);
	band.cyChild = static_cast<UINT>(-1);
	band.cyIntegral = static_cast<UINT>(-1);
	band.cxMinChild = static_cast<UINT>(-1);
	band.cyMinChild = newValue;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_MinimumWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = band.cxMinChild;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_MinimumWidth(OLE_XSIZE_PIXELS newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_CHILDSIZE;
	band.cyMinChild = static_cast<UINT>(-1);
	band.cyMaxChild = static_cast<UINT>(-1);
	band.cyChild = static_cast<UINT>(-1);
	band.cyIntegral = static_cast<UINT>(-1);
	band.cxMinChild = newValue;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_NewLine(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(band.fStyle & RBBS_BREAK);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_NewLine(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle &= ~RBBS_BREAK;
		} else {
			band.fStyle |= RBBS_BREAK;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_Resizable(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(!(band.fStyle & RBBS_FIXEDSIZE));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_Resizable(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle |= RBBS_FIXEDSIZE;
		} else {
			band.fStyle &= ~RBBS_FIXEDSIZE;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_RowHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	*pValue = SendMessage(hWndRBar, RB_GETROWHEIGHT, properties.bandIndex, 0);
	return S_OK;
}

STDMETHODIMP ReBarBand::get_ShowTitle(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(!(band.fStyle & RBBS_HIDETITLE));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_ShowTitle(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle |= RBBS_HIDETITLE;
		} else {
			band.fStyle &= ~RBBS_HIDETITLE;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_SizingGripVisibility(SizingGripVisibilityConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SizingGripVisibilityConstants);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(band.fStyle & RBBS_NOGRIPPER) {
			*pValue = sgvNever;
		} else if(band.fStyle & RBBS_GRIPPERALWAYS) {
			*pValue = sgvAlways;
		} else {
			*pValue = sgvAutomatic;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_SizingGripVisibility(SizingGripVisibilityConstants newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		switch(newValue) {
			case sgvNever:
				band.fStyle &= ~RBBS_GRIPPERALWAYS;
				band.fStyle |= RBBS_NOGRIPPER;
				break;
			case sgvAutomatic:
				band.fStyle &= ~(RBBS_NOGRIPPER | RBBS_GRIPPERALWAYS);
				break;
			case sgvAlways:
				band.fStyle &= ~RBBS_NOGRIPPER;
				band.fStyle |= RBBS_GRIPPERALWAYS;
				break;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_Text(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_TEXT;
	band.cch = MAX_BANDTEXTLENGTH;
	band.lpText = reinterpret_cast<LPTSTR>(malloc((band.cch + 1) * sizeof(TCHAR)));
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = _bstr_t(band.lpText).Detach();
		SECUREFREE(band.lpText);
		return S_OK;
	}
	SECUREFREE(band.lpText);
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_Text(BSTR newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_TEXT;
	COLE2T converter(newValue);
	band.lpText = converter;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_TitleWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_HEADERSIZE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = band.cxHeader;
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_TitleWidth(OLE_XSIZE_PIXELS newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_HEADERSIZE;
	band.cxHeader = newValue;
	if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_UseChevron(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(band.fStyle & RBBS_USECHEVRON);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_UseChevron(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle &= ~RBBS_USECHEVRON;
		} else {
			band.fStyle |= RBBS_USECHEVRON;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_VariableHeight(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(band.fStyle & RBBS_VARIABLEHEIGHT);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_VariableHeight(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle &= ~RBBS_VARIABLEHEIGHT;
		} else {
			band.fStyle |= RBBS_VARIABLEHEIGHT;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::get_Visible(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		*pValue = BOOL2VARIANTBOOL(!(band.fStyle & RBBS_HIDDEN));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBarBand::put_Visible(VARIANT_BOOL newValue)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_STYLE;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(newValue == VARIANT_FALSE) {
			band.fStyle |= RBBS_HIDDEN;
		} else {
			band.fStyle &= ~RBBS_HIDDEN;
		}
		if(SendMessage(hWndRBar, RB_SETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
			return S_OK;
		}
	}
	return E_FAIL;
}


STDMETHODIMP ReBarBand::ClickChevron(LONG userData/* = 0*/)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	SendMessage(hWndRBar, RB_PUSHCHEVRON, properties.bandIndex, userData);
	return S_OK;
}

STDMETHODIMP ReBarBand::GetBorderSizes(OLE_XSIZE_PIXELS* pLeftBorder/* = NULL*/, OLE_YSIZE_PIXELS* pTopBorder/* = NULL*/, OLE_XSIZE_PIXELS* pRightBorder/* = NULL*/, OLE_YSIZE_PIXELS* pBottomBorder/* = NULL*/)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	RECT rc = {0};
	SendMessage(hWndRBar, RB_GETBANDBORDERS, properties.bandIndex, reinterpret_cast<LPARAM>(&rc));
	if(pLeftBorder) {
		*pLeftBorder = rc.left;
	}
	if(pTopBorder) {
		*pTopBorder = rc.top;
	}
	if(pRightBorder) {
		*pRightBorder = rc.right;
	}
	if(pBottomBorder) {
		*pBottomBorder = rc.bottom;
	}
	return S_OK;
}

STDMETHODIMP ReBarBand::GetChevronRectangle(OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	REBARBANDINFO band = {0};
	if(properties.pOwnerRBar->IsComctl32Version610OrNewer()) {
		band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	} else {
		band.cbSize = sizeof(REBARBANDINFO);
	}
	band.fMask = RBBIM_CHEVRONLOCATION;
	if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
		if(pXLeft) {
			*pXLeft = band.rcChevronLocation.left;
		}
		if(pYTop) {
			*pYTop = band.rcChevronLocation.top;
		}
		if(pXRight) {
			*pXRight = band.rcChevronLocation.right;
		}
		if(pYBottom) {
			*pYBottom = band.rcChevronLocation.bottom;
		}
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ReBarBand::GetRectangle(BandRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft/* = NULL*/, OLE_YPOS_PIXELS* pYTop/* = NULL*/, OLE_XPOS_PIXELS* pXRight/* = NULL*/, OLE_YPOS_PIXELS* pYBottom/* = NULL*/)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	HRESULT hr = E_FAIL;
	RECT rc = {0};
	switch(rectangleType) {
		case brtBand:
			if(SendMessage(hWndRBar, RB_GETRECT, properties.bandIndex, reinterpret_cast<LPARAM>(&rc))) {
				hr = S_OK;
			}
			break;
		case brtContainedWindow:
		{
			REBARBANDINFO band = {0};
			band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
			band.fMask = RBBIM_CHILD;
			if(SendMessage(hWndRBar, RB_GETBANDINFO, properties.bandIndex, reinterpret_cast<LPARAM>(&band))) {
				GetWindowRect(band.hwndChild, &rc);
				if(MapWindowPoints(HWND_DESKTOP, hWndRBar, reinterpret_cast<LPPOINT>(&rc), 2)) {
					hr = S_OK;
				}
			}
			break;
		}
		default:
			hr = E_INVALIDARG;
			break;
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

STDMETHODIMP ReBarBand::Hide(void)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	SendMessage(hWndRBar, RB_SHOWBAND, properties.bandIndex, FALSE);
	return S_OK;
}

STDMETHODIMP ReBarBand::Maximize(VARIANT_BOOL useIdealWidth/* = VARIANT_TRUE*/)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	SendMessage(hWndRBar, RB_MAXIMIZEBAND, properties.bandIndex, VARIANTBOOL2BOOL(useIdealWidth));
	return S_OK;
}

STDMETHODIMP ReBarBand::Minimize(void)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	SendMessage(hWndRBar, RB_MINIMIZEBAND, properties.bandIndex, 0);
	return S_OK;
}

STDMETHODIMP ReBarBand::Show(void)
{
	HWND hWndRBar = properties.GetRBarHWnd();
	ATLASSERT(IsWindow(hWndRBar));

	SendMessage(hWndRBar, RB_SHOWBAND, properties.bandIndex, TRUE);
	return S_OK;
}