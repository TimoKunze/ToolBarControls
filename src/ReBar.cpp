// ReBar.cpp: Superclasses ReBarWindow32.

#include "stdafx.h"
#include "ReBar.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


// initialize complex constants
FONTDESC ReBar::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


ReBar::ReBar()
{
	properties.font.InitializePropertyWatcher(this, DISPID_RB_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_RB_MOUSEICON);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;
	
	// initialize
	bandUnderMouse = -1;
	bandChevronPushed = -1;

	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ReBar::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IReBar, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ReBar::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 13 VT_I4 properties...
	pSize->QuadPart += 13 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 11 VT_BOOL properties...
	pSize->QuadPart += 11 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));

	// ...and 2 VT_DISPATCH properties
	pSize->QuadPart += 2 * (sizeof(VARTYPE) + sizeof(CLSID));

	// we've to query each object for its size
	CComPtr<IPersistStreamInit> pStreamInit = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.font.pFontDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	pStreamInit = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pStreamInit)))) {
			properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStreamInit, reinterpret_cast<LPVOID*>(&pStreamInit));
		}
	}
	if(pStreamInit) {
		ULARGE_INTEGER tmp = {0};
		pStreamInit->GetSizeMax(&tmp);
		pSize->QuadPart += tmp.QuadPart;
	}

	return S_OK;
}

STDMETHODIMP ReBar::Load(LPSTREAM pStream)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0;
	if(FAILED(hr = pStream->Read(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	if(signature != 0x0F0F0F0F/*4x AppID*/) {
		return E_FAIL;
	}
	LONG version = 0;
	if(FAILED(hr = pStream->Read(&version, sizeof(version), NULL))) {
		return hr;
	}
	if(version > 0x0100) {
		return E_FAIL;
	}
	LONG subSignature = 0;
	if(FAILED(hr = pStream->Read(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}
	if(subSignature != 0x01010101/*4x 0x01 (-> ReBar)*/) {
		return E_FAIL;
	}

	DWORD atlVersion;
	if(FAILED(hr = pStream->Read(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}
	if(atlVersion > _ATL_VER) {
		return E_FAIL;
	}

	if(FAILED(hr = pStream->Read(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	typedef HRESULT ReadVariantFromStreamFn(__in LPSTREAM pStream, VARTYPE expectedVarType, __inout LPVARIANT pVariant);
	ReadVariantFromStreamFn* pfnReadVariantFromStream = ReadVariantFromStream;

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL allowBandReordering = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL autoUpdateLayout = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR backColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displayBandSeparators = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displaySplitter = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL fixedBandHeight = propertyValue.boolVal;

	VARTYPE vt;
	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IFontDisp> pFont = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pFont)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR foreColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR highlightColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;

	if(FAILED(hr = pStream->Read(&vt, sizeof(VARTYPE), NULL)) || (vt != VT_DISPATCH)) {
		return hr;
	}
	CComPtr<IPictureDisp> pMouseIcon = NULL;
	if(FAILED(hr = OleLoadFromStream(pStream, IID_IDispatch, reinterpret_cast<LPVOID*>(&pMouseIcon)))) {
		if(hr != REGDB_E_CLASSNOTREG) {
			return S_OK;
		}
	}

	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MousePointerConstants mousePointer = static_cast<MousePointerConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OrientationConstants orientation = static_cast<OrientationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RegisterForOLEDragDropConstants registerForOLEDragDrop = static_cast<RegisterForOLEDragDropConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ReplaceMDIFrameMenuConstants replaceMDIFrameMenu = static_cast<ReplaceMDIFrameMenuConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RightToLeftConstants rightToLeft = static_cast<RightToLeftConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR shadowColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL toggleOnDoubleClick = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL verticalSizingGripsOnVerticalOrientation = propertyValue.boolVal;


	hr = put_AllowBandReordering(allowBandReordering);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AutoUpdateLayout(autoUpdateLayout);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackColor(backColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayBandSeparators(displayBandSeparators);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplaySplitter(displaySplitter);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FixedBandHeight(fixedBandHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ForeColor(foreColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HighlightColor(highlightColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MouseIcon(pMouseIcon);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MousePointer(mousePointer);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Orientation(orientation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ReplaceMDIFrameMenu(replaceMDIFrameMenu);
	if(FAILED(hr) && hr != MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 382)) {
		return hr;
	}
	hr = put_RightToLeft(rightToLeft);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ShadowColor(shadowColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ToggleOnDoubleClick(toggleOnDoubleClick);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_VerticalSizingGripsOnVerticalOrientation(verticalSizingGripsOnVerticalOrientation);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ReBar::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
{
	ATLASSUME(pStream);
	if(!pStream) {
		return E_POINTER;
	}

	HRESULT hr = S_OK;
	LONG signature = 0x0F0F0F0F/*4x AppID*/;
	if(FAILED(hr = pStream->Write(&signature, sizeof(signature), NULL))) {
		return hr;
	}
	LONG version = 0x0100;
	if(FAILED(hr = pStream->Write(&version, sizeof(version), NULL))) {
		return hr;
	}
	LONG subSignature = 0x01010101/*4x 0x01 (-> ReBar)*/;
	if(FAILED(hr = pStream->Write(&subSignature, sizeof(subSignature), NULL))) {
		return hr;
	}

	DWORD atlVersion = _ATL_VER;
	if(FAILED(hr = pStream->Write(&atlVersion, sizeof(atlVersion), NULL))) {
		return hr;
	}

	if(FAILED(hr = pStream->Write(&m_sizeExtent, sizeof(m_sizeExtent), NULL))) {
		return hr;
	}

	VARIANT propertyValue;
	VariantInit(&propertyValue);

	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowBandReordering);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.autoUpdateLayout);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.backColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayBandSeparators);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displaySplitter);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.fixedBandHeight);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	CComPtr<IPersistStream> pPersistStream = NULL;
	if(properties.font.pFontDisp) {
		if(FAILED(hr = properties.font.pFontDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	VARTYPE vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.foreColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.highlightColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	pPersistStream = NULL;
	if(properties.mouseIcon.pPictureDisp) {
		if(FAILED(hr = properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPersistStream, reinterpret_cast<LPVOID*>(&pPersistStream)))) {
			return hr;
		}
	}
	// store some marker
	vt = VT_DISPATCH;
	if(FAILED(hr = pStream->Write(&vt, sizeof(VARTYPE), NULL))) {
		return hr;
	}
	if(pPersistStream) {
		if(FAILED(hr = OleSaveToStream(pPersistStream, pStream))) {
			return hr;
		}
	} else {
		if(FAILED(hr = WriteClassStm(pStream, CLSID_NULL))) {
			return hr;
		}
	}

	propertyValue.lVal = properties.mousePointer;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.orientation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.registerForOLEDragDrop;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.replaceMDIFrameMenu;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.rightToLeft;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.shadowColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.toggleOnDoubleClick);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.verticalSizingGripsOnVerticalOrientation);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND ReBar::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_COOL_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<ReBar>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT ReBar::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("ReBar ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void ReBar::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP ReBar::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<ReBar>::IOleObject_SetClientSite(pClientSite);

	/* Check whether the container has an ambient font. If it does, clone it; otherwise create our own
	   font object when we hook up a client site. */
	if(!properties.font.pFontDisp) {
		FONTDESC defaultFont = properties.font.GetDefaultFont();
		CComPtr<IFontDisp> pFont;
		if(FAILED(GetAmbientFontDisp(&pFont))) {
			// use the default font
			OleCreateFontIndirect(&defaultFont, IID_IFontDisp, reinterpret_cast<LPVOID*>(&pFont));
		}
		put_Font(pFont);
	}

	return hr;
}

STDMETHODIMP ReBar::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL ReBar::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		if(pMessage->wParam == VK_TAB) {
			if(dialogCode & DLGC_WANTTAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
	}
	return CComControl<ReBar>::PreTranslateAccelerator(pMessage, hReturnValue);
}


//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP ReBar::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);

	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ReBar::DragLeave(void)
{
	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragLeave(&callDropTargetHelper);

	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->DragLeave();
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP ReBar::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition, &callDropTargetHelper);

	if(callDropTargetHelper && dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ReBar::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	BOOL callDropTargetHelper = TRUE;
	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition, &callDropTargetHelper);

	if(dragDropStatus.pDropTargetHelper) {
		if(callDropTargetHelper) {
			dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		}
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP ReBar::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
{
	switch(category) {
		case PROPCAT_Colors:
			*pName = GetResString(IDPC_COLORS).Detach();
			return S_OK;
			break;
		case PROPCAT_DragDrop:
			*pName = GetResString(IDPC_DRAGDROP).Detach();
			return S_OK;
			break;
	}
	return E_FAIL;
}

STDMETHODIMP ReBar::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_RB_APPEARANCE:
		case DISPID_RB_BORDERSTYLE:
		case DISPID_RB_CONTROLHEIGHT:
		case DISPID_RB_DISPLAYBANDSEPARATORS:
		case DISPID_RB_DISPLAYSPLITTER:
		case DISPID_RB_MOUSEICON:
		case DISPID_RB_MOUSEPOINTER:
		case DISPID_RB_ORIENTATION:
		case DISPID_RB_VERTICALSIZINGGRIPSONVERTICALORIENTATION:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_RB_ALLOWBANDREORDERING:
		case DISPID_RB_AUTOUPDATELAYOUT:
		case DISPID_RB_DISABLEDEVENTS:
		case DISPID_RB_DONTREDRAW:
		case DISPID_RB_FIXEDBANDHEIGHT:
		case DISPID_RB_HOVERTIME:
		case DISPID_RB_REPLACEMDIFRAMEMENU:
		case DISPID_RB_RIGHTTOLEFT:
		case DISPID_RB_TOGGLEONDOUBLECLICK:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_RB_BACKCOLOR:
		case DISPID_RB_FORECOLOR:
		case DISPID_RB_HIGHLIGHTCOLOR:
		case DISPID_RB_HPALETTE:
		case DISPID_RB_SHADOWCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_RB_APPID:
		case DISPID_RB_APPNAME:
		case DISPID_RB_APPSHORTNAME:
		case DISPID_RB_BUILD:
		case DISPID_RB_CHARSET:
		case DISPID_RB_HIMAGELIST:
		case DISPID_RB_HWND:
		//case DISPID_RB_HWNDTOOLTIP:
		case DISPID_RB_ISRELEASE:
		case DISPID_RB_PROGRAMMER:
		case DISPID_RB_TESTER:
		case DISPID_RB_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_RB_NATIVEDROPTARGET:
		case DISPID_RB_REGISTERFOROLEDRAGDROP:
		case DISPID_RB_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_RB_FONT:
		case DISPID_RB_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_RB_BANDS:
		case DISPID_RB_MDIFRAMEMENUBAND:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_RB_ENABLED:
			*pCategory = PROPCAT_Misc;
			return S_OK;
			break;
	}
	return E_FAIL;
}
// implementation of ICategorizeProperties
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICreditsProvider
CAtlString ReBar::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString ReBar::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString ReBar::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=DateTimeControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString ReBar::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString ReBar::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString ReBar::GetVersion(void)
{
	CAtlString ret = TEXT("Version ");
	CComBSTR buffer;
	get_Version(&buffer);
	ret += buffer;
	ret += TEXT(" (");
	get_CharSet(&buffer);
	ret += buffer;
	ret += TEXT(")\nCompilation timestamp: ");
	ret += TEXT(STRTIMESTAMP);
	ret += TEXT("\n");

	VARIANT_BOOL b;
	get_IsRelease(&b);
	if(b == VARIANT_FALSE) {
		ret += TEXT("This version is for debugging only.");
	}

	return ret;
}
// implementation of ICreditsProvider
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP ReBar::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_RB_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_RB_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_RB_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_RB_ORIENTATION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.orientation), description);
			break;
		case DISPID_RB_REGISTERFOROLEDRAGDROP:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.registerForOLEDragDrop), description);
			break;
		case DISPID_RB_REPLACEMDIFRAMEMENU:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.replaceMDIFrameMenu), description);
			break;
		case DISPID_RB_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<ReBar>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP ReBar::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_RB_BORDERSTYLE:
		case DISPID_RB_ORIENTATION:
			c = 2;
			break;
		case DISPID_RB_APPEARANCE:
		case DISPID_RB_REGISTERFOROLEDRAGDROP:
		case DISPID_RB_REPLACEMDIFRAMEMENU:
			c = 3;
			break;
		case DISPID_RB_RIGHTTOLEFT:
			c = 4;
			break;
		case DISPID_RB_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<ReBar>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_RB_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
			propertyValue = mpCustom;
		}

		CComBSTR description;
		HRESULT hr = GetDisplayStringForSetting(property, propertyValue, description);
		if(SUCCEEDED(hr)) {
			size_t bufferSize = SysStringLen(description) + 1;
			pDescriptions->pElems[iDescription] = static_cast<LPOLESTR>(CoTaskMemAlloc(bufferSize * sizeof(WCHAR)));
			ATLVERIFY(SUCCEEDED(StringCchCopyW(pDescriptions->pElems[iDescription], bufferSize, description)));
			// simply use the property value as cookie
			pCookies->pElems[iDescription] = propertyValue;
		} else {
			return DISP_E_BADINDEX;
		}
	}
	return S_OK;
}

STDMETHODIMP ReBar::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_RB_APPEARANCE:
		case DISPID_RB_BORDERSTYLE:
		case DISPID_RB_MOUSEPOINTER:
		case DISPID_RB_ORIENTATION:
		case DISPID_RB_REGISTERFOROLEDRAGDROP:
		case DISPID_RB_REPLACEMDIFRAMEMENU:
		case DISPID_RB_RIGHTTOLEFT:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<ReBar>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP ReBar::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	return IPerPropertyBrowsingImpl<ReBar>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT ReBar::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_RB_APPEARANCE:
			switch(cookie) {
				case a2D:
					description = GetResStringWithNumber(IDP_APPEARANCE2D, a2D);
					break;
				case a3D:
					description = GetResStringWithNumber(IDP_APPEARANCE3D, a3D);
					break;
				case a3DLight:
					description = GetResStringWithNumber(IDP_APPEARANCE3DLIGHT, a3DLight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RB_BORDERSTYLE:
			switch(cookie) {
				case bsNone:
					description = GetResStringWithNumber(IDP_BORDERSTYLENONE, bsNone);
					break;
				case bsFixedSingle:
					description = GetResStringWithNumber(IDP_BORDERSTYLEFIXEDSINGLE, bsFixedSingle);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RB_MOUSEPOINTER:
			switch(cookie) {
				case mpDefault:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERDEFAULT, mpDefault);
					break;
				case mpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROW, mpArrow);
					break;
				case mpCross:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCROSS, mpCross);
					break;
				case mpIBeam:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERIBEAM, mpIBeam);
					break;
				case mpIcon:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERICON, mpIcon);
					break;
				case mpSize:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZE, mpSize);
					break;
				case mpSizeNESW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENESW, mpSizeNESW);
					break;
				case mpSizeNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENS, mpSizeNS);
					break;
				case mpSizeNWSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZENWSE, mpSizeNWSE);
					break;
				case mpSizeEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEEW, mpSizeEW);
					break;
				case mpUpArrow:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERUPARROW, mpUpArrow);
					break;
				case mpHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHOURGLASS, mpHourglass);
					break;
				case mpNoDrop:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERNODROP, mpNoDrop);
					break;
				case mpArrowHourglass:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWHOURGLASS, mpArrowHourglass);
					break;
				case mpArrowQuestion:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERARROWQUESTION, mpArrowQuestion);
					break;
				case mpSizeAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSIZEALL, mpSizeAll);
					break;
				case mpHand:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERHAND, mpHand);
					break;
				case mpInsertMedia:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERINSERTMEDIA, mpInsertMedia);
					break;
				case mpScrollAll:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLALL, mpScrollAll);
					break;
				case mpScrollN:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLN, mpScrollN);
					break;
				case mpScrollNE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNE, mpScrollNE);
					break;
				case mpScrollE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLE, mpScrollE);
					break;
				case mpScrollSE:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSE, mpScrollSE);
					break;
				case mpScrollS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLS, mpScrollS);
					break;
				case mpScrollSW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLSW, mpScrollSW);
					break;
				case mpScrollW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLW, mpScrollW);
					break;
				case mpScrollNW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNW, mpScrollNW);
					break;
				case mpScrollNS:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLNS, mpScrollNS);
					break;
				case mpScrollEW:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERSCROLLEW, mpScrollEW);
					break;
				case mpCustom:
					description = GetResStringWithNumber(IDP_MOUSEPOINTERCUSTOM, mpCustom);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RB_ORIENTATION:
			switch(cookie) {
				case oHorizontal:
					description = GetResStringWithNumber(IDP_ORIENTATIONHORIZONTAL, oHorizontal);
					break;
				case oVertical:
					description = GetResStringWithNumber(IDP_ORIENTATIONVERTICAL, oVertical);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RB_REGISTERFOROLEDRAGDROP:
			switch(cookie) {
				case rfoddNoDragDrop:
					description = GetResStringWithNumber(IDP_REGISTERFOROLEDRAGDROPNONE, rfoddNoDragDrop);
					break;
				case rfoddNativeDragDrop:
					description = GetResStringWithNumber(IDP_REGISTERFOROLEDRAGDROPNATIVE, rfoddNativeDragDrop);
					break;
				case rfoddAdvancedDragDrop:
					description = GetResStringWithNumber(IDP_REGISTERFOROLEDRAGDROPADVANCED, rfoddAdvancedDragDrop);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RB_REPLACEMDIFRAMEMENU:
			switch(cookie) {
				case rmfmDontReplace:
					description = GetResStringWithNumber(IDP_REPLACEMDIFRAMEMENUDONTREPLACE, rmfmDontReplace);
					break;
				case rmfmJustRemove:
					description = GetResStringWithNumber(IDP_REPLACEMDIFRAMEMENUJUSTREMOVE, rmfmJustRemove);
					break;
				case rmfmFullReplace:
					description = GetResStringWithNumber(IDP_REPLACEMDIFRAMEMENUFULLREPLACE, rmfmFullReplace);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_RB_RIGHTTOLEFT:
			switch(cookie) {
				case 0:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTNONE, 0);
					break;
				case rtlText:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXT, rtlText);
					break;
				case rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTLAYOUT, rtlLayout);
					break;
				case rtlText | rtlLayout:
					description = GetResStringWithNumber(IDP_RIGHTTOLEFTTEXTLAYOUT, rtlText | rtlLayout);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		default:
			return DISP_E_BADINDEX;
			break;
	}

	return S_OK;
}

//////////////////////////////////////////////////////////////////////
// implementation of ISpecifyPropertyPages
STDMETHODIMP ReBar::GetPages(CAUUID* pPropertyPages)
{
	if(!pPropertyPages) {
		return E_POINTER;
	}

	pPropertyPages->cElems = 4;
	pPropertyPages->pElems = static_cast<LPGUID>(CoTaskMemAlloc(sizeof(GUID) * pPropertyPages->cElems));
	if(pPropertyPages->pElems) {
		pPropertyPages->pElems[0] = CLSID_CommonProperties;
		pPropertyPages->pElems[1] = CLSID_StockColorPage;
		pPropertyPages->pElems[2] = CLSID_StockFontPage;
		pPropertyPages->pElems[3] = CLSID_StockPicturePage;
		return S_OK;
	}
	return E_OUTOFMEMORY;
}
// implementation of ISpecifyPropertyPages
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleObject
STDMETHODIMP ReBar::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<ReBar>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP ReBar::EnumVerbs(IEnumOLEVERB** ppEnumerator)
{
	static OLEVERB oleVerbs[3] = {
	    {OLEIVERB_UIACTIVATE, L"&Edit", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	    {OLEIVERB_PROPERTIES, L"&Properties...", 0, OLEVERBATTRIB_ONCONTAINERMENU},
	    {1, L"&About...", 0, OLEVERBATTRIB_NEVERDIRTIES | OLEVERBATTRIB_ONCONTAINERMENU},
	};
	return EnumOLEVERB::CreateInstance(oleVerbs, 3, ppEnumerator);
}
// implementation of IOleObject
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IOleControl
STDMETHODIMP ReBar::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = properties.hAcceleratorTable;
	pControlInfo->cAccel = static_cast<USHORT>(properties.hAcceleratorTable ? CopyAcceleratorTable(properties.hAcceleratorTable, NULL, 0) : 0);
	pControlInfo->dwFlags = 0;
	return S_OK;
}

STDMETHODIMP ReBar::OnMnemonic(LPMSG pMessage)
{
	if(GetStyle() & WS_DISABLED) {
		return S_OK;
	}

	ATLASSERT(pMessage->message == WM_SYSKEYDOWN);

	SHORT pressedKeyCode = static_cast<SHORT>(pMessage->wParam);
	int bandToSelect = -1;

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_TEXT;
	band.cch = MAX_BANDTEXTLENGTH;
	band.lpText = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (band.cch + 1) * sizeof(TCHAR)));
	if(band.lpText) {
		int numberOfBands = SendMessage(RB_GETBANDCOUNT, 0, 0);
		for(int bandIndex = 0; bandIndex < numberOfBands; ++bandIndex) {
			SendMessage(RB_GETBANDINFO, bandIndex, reinterpret_cast<LPARAM>(&band));

			if(band.lpText) {
				for(int i = lstrlen(band.lpText); i > 0; --i) {
					if((band.lpText[i - 1] == TEXT('&')) && (band.lpText[i] != TEXT('&'))) {
						// TODO: Does this work with MFC?
						if((VkKeyScan(band.lpText[i]) == pressedKeyCode) || (VkKeyScan(static_cast<TCHAR>(tolower(band.lpText[i]))) == pressedKeyCode)) {
							bandToSelect = bandIndex;
							break;
						}
					}
				}
			}
		}
		HeapFree(GetProcessHeap(), 0, band.lpText);
	}
	if(bandToSelect != -1) {
		band.fMask = RBBIM_CHILD;
		if(SendMessage(RB_GETBANDINFO, bandToSelect, reinterpret_cast<LPARAM>(&band)) && band.hwndChild && ::IsWindow(band.hwndChild)) {
			::SetFocus(band.hwndChild);
		}
	}
	return S_OK;
}
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT ReBar::DoVerbAbout(HWND hWndParent)
{
	HRESULT hr = S_OK;
	//hr = OnPreVerbAbout();
	if(SUCCEEDED(hr))	{
		AboutDlg dlg;
		dlg.SetOwner(this);
		dlg.DoModal(hWndParent);
		hr = S_OK;
		//hr = OnPostVerbAbout();
	}
	return hr;
}

HRESULT ReBar::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_RB_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT ReBar::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP ReBar::get_AllowBandReordering(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowBandReordering = !(GetStyle() & RBS_FIXEDORDER);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowBandReordering);
	return S_OK;
}

STDMETHODIMP ReBar::put_AllowBandReordering(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_ALLOWBANDREORDERING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowBandReordering != b) {
		properties.allowBandReordering = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowBandReordering) {
				ModifyStyle(RBS_FIXEDORDER, 0);
			} else {
				ModifyStyle(0, RBS_FIXEDORDER);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_ALLOWBANDREORDERING);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_Appearance(AppearanceConstants* pValue)
{
	ATLASSERT_POINTER(pValue, AppearanceConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetExStyle() & WS_EX_CLIENTEDGE) {
			properties.appearance = a3D;
		} else if(GetExStyle() & WS_EX_STATICEDGE) {
			properties.appearance = a3DLight;
		} else {
			properties.appearance = a2D;
		}
	}

	*pValue = properties.appearance;
	return S_OK;
}

STDMETHODIMP ReBar::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_RB_APPEARANCE);
	if(properties.appearance != newValue) {
		properties.appearance = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.appearance) {
				case a2D:
					ModifyStyleEx(WS_EX_CLIENTEDGE | WS_EX_STATICEDGE, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3D:
					ModifyStyleEx(WS_EX_STATICEDGE, WS_EX_CLIENTEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case a3DLight:
					ModifyStyleEx(WS_EX_CLIENTEDGE, WS_EX_STATICEDGE, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 15;
	return S_OK;
}

STDMETHODIMP ReBar::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ToolBarControls");
	return S_OK;
}

STDMETHODIMP ReBar::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"TBarCtls");
	return S_OK;
}

STDMETHODIMP ReBar::get_AutoUpdateLayout(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.autoUpdateLayout = ((GetStyle() & RBS_AUTOSIZE) == RBS_AUTOSIZE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.autoUpdateLayout);
	return S_OK;
}

STDMETHODIMP ReBar::put_AutoUpdateLayout(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_AUTOUPDATELAYOUT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.autoUpdateLayout != b) {
		properties.autoUpdateLayout = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.autoUpdateLayout) {
				ModifyStyle(0, RBS_AUTOSIZE);
			} else {
				ModifyStyle(RBS_AUTOSIZE, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_AUTOUPDATELAYOUT);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_BackColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(RB_GETBKCOLOR, 0, 0));
		if(color == CLR_DEFAULT) {
			properties.backColor = static_cast<OLE_COLOR>(-1);
		} else if(color != OLECOLOR2COLORREF(properties.backColor)) {
			properties.backColor = color;
		}
	}

	*pValue = properties.backColor;
	return S_OK;
}

STDMETHODIMP ReBar::put_BackColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_RB_BACKCOLOR);
	if(properties.backColor != newValue) {
		properties.backColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(RB_SETBKCOLOR, 0, (properties.backColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.backColor)));
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_BACKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_Bands(IReBarBands** ppBands)
{
	ATLASSERT_POINTER(ppBands, IReBarBands*);
	if(!ppBands) {
		return E_POINTER;
	}

	ClassFactory::InitReBarBands(this, IID_IReBarBands, reinterpret_cast<LPUNKNOWN*>(ppBands));
	return S_OK;
}

STDMETHODIMP ReBar::get_BorderStyle(BorderStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BorderStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.borderStyle = ((GetStyle() & WS_BORDER) == WS_BORDER ? bsFixedSingle : bsNone);
	}

	*pValue = properties.borderStyle;
	return S_OK;
}

STDMETHODIMP ReBar::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_RB_BORDERSTYLE);
	if(properties.borderStyle != newValue) {
		properties.borderStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.borderStyle) {
				case bsNone:
					ModifyStyle(WS_BORDER, 0, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
				case bsFixedSingle:
					ModifyStyle(0, WS_BORDER, SWP_DRAWFRAME | SWP_FRAMECHANGED);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP ReBar::get_CharSet(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef UNICODE
		*pValue = SysAllocString(L"Unicode");
	#else
		*pValue = SysAllocString(L"ANSI");
	#endif
	return S_OK;
}

STDMETHODIMP ReBar::get_ControlHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = SendMessage(RB_GETBARHEIGHT, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBar::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP ReBar::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_RB_DISABLEDEVENTS);
	if(properties.disabledEvents != newValue) {
		if((static_cast<long>(properties.disabledEvents) & deMouseEvents) != (static_cast<long>(newValue) & deMouseEvents)) {
			if(IsWindow()) {
				if(static_cast<long>(newValue) & deMouseEvents) {
					// nothing to do
				} else {
					TRACKMOUSEEVENT trackingOptions = {0};
					trackingOptions.cbSize = sizeof(trackingOptions);
					trackingOptions.hwndTrack = *this;
					trackingOptions.dwFlags = TME_HOVER | TME_LEAVE | TME_CANCEL;
					TrackMouseEvent(&trackingOptions);
					bandUnderMouse = -1;
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RB_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_DisplayBandSeparators(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.displayBandSeparators = ((GetStyle() & RBS_BANDBORDERS) == RBS_BANDBORDERS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayBandSeparators);
	return S_OK;
}

STDMETHODIMP ReBar::put_DisplayBandSeparators(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_DISPLAYBANDSEPARATORS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayBandSeparators != b) {
		properties.displayBandSeparators = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.displayBandSeparators) {
				ModifyStyle(0, RBS_BANDBORDERS);
			} else {
				ModifyStyle(RBS_BANDBORDERS, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_DISPLAYBANDSEPARATORS);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_DisplaySplitter(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && IsComctl32Version610OrNewer()) {
		properties.displaySplitter = ((SendMessage(RB_GETEXTENDEDSTYLE, 0, 0) & RBS_EX_SPLITTER) == RBS_EX_SPLITTER);
	}

	*pValue = BOOL2VARIANTBOOL(properties.displaySplitter);
	return S_OK;
}

STDMETHODIMP ReBar::put_DisplaySplitter(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_DISPLAYSPLITTER);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displaySplitter != b) {
		properties.displaySplitter = b;
		SetDirty(TRUE);

		if(IsWindow() && IsComctl32Version610OrNewer()) {
			if(properties.displaySplitter) {
				SendMessage(RB_SETEXTENDEDSTYLE, RBS_EX_SPLITTER, RBS_EX_SPLITTER);
			} else {
				SendMessage(RB_SETEXTENDEDSTYLE, RBS_EX_SPLITTER, 0);
			}
		}
		FireOnChanged(DISPID_RB_DISPLAYSPLITTER);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP ReBar::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_RB_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_Enabled(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.enabled = !(GetStyle() & WS_DISABLED);
	}

	*pValue = BOOL2VARIANTBOOL(properties.enabled);
	return S_OK;
}

STDMETHODIMP ReBar::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_ENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.enabled != b) {
		properties.enabled = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			EnableWindow(properties.enabled);
			FireViewChange();
		}

		if(!properties.enabled) {
			IOleInPlaceObject_UIDeactivate();
		}

		FireOnChanged(DISPID_RB_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_FixedBandHeight(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.fixedBandHeight = !(GetStyle() & RBS_VARHEIGHT);
	}

	*pValue = BOOL2VARIANTBOOL(properties.fixedBandHeight);
	return S_OK;
}

STDMETHODIMP ReBar::put_FixedBandHeight(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_FIXEDBANDHEIGHT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.fixedBandHeight != b) {
		properties.fixedBandHeight = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.fixedBandHeight) {
				ModifyStyle(RBS_VARHEIGHT, 0);
			} else {
				ModifyStyle(0, RBS_VARHEIGHT);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_FIXEDBANDHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_Font(IFontDisp** ppFont)
{
	ATLASSERT_POINTER(ppFont, IFontDisp*);
	if(!ppFont) {
		return E_POINTER;
	}

	if(*ppFont) {
		(*ppFont)->Release();
		*ppFont = NULL;
	}
	if(properties.font.pFontDisp) {
		properties.font.pFontDisp->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(ppFont));
	}
	return S_OK;
}

STDMETHODIMP ReBar::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_RB_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			CComQIPtr<IFont, &IID_IFont> pFont(pNewFont);
			if(pFont) {
				CComPtr<IFont> pClonedFont = NULL;
				pFont->Clone(&pClonedFont);
				if(pClonedFont) {
					pClonedFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
				}
			}
		}
		properties.font.StartWatching();
	}
	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RB_FONT);
	return S_OK;
}

STDMETHODIMP ReBar::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_RB_FONT);
	if(properties.font.pFontDisp != pNewFont) {
		properties.font.StopWatching();
		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(pNewFont) {
			pNewFont->QueryInterface(IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
		}
		properties.font.StartWatching();
	} else if(pNewFont) {
		pNewFont->AddRef();
	}

	if(!properties.useSystemFont) {
		ApplyFont();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RB_FONT);
	return S_OK;
}

STDMETHODIMP ReBar::get_ForeColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(RB_GETTEXTCOLOR, 0, 0));
		if(color == CLR_DEFAULT) {
			properties.foreColor = static_cast<OLE_COLOR>(-1);
		} else if(color != OLECOLOR2COLORREF(properties.foreColor)) {
			properties.foreColor = color;
		}
	}

	*pValue = properties.foreColor;
	return S_OK;
}

STDMETHODIMP ReBar::put_ForeColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_RB_FORECOLOR);
	if(properties.foreColor != newValue) {
		properties.foreColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(RB_SETTEXTCOLOR, 0, (properties.foreColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.foreColor)));
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_FORECOLOR);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_HighlightColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
		if(SendMessage(RB_GETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme))) {
			if(colorScheme.clrBtnHighlight == CLR_DEFAULT) {
				properties.highlightColor = static_cast<OLE_COLOR>(-1);
			} else if(colorScheme.clrBtnHighlight != OLECOLOR2COLORREF(properties.highlightColor)) {
				properties.highlightColor = colorScheme.clrBtnHighlight;
			}
		}
	}

	*pValue = properties.highlightColor;
	return S_OK;
}

STDMETHODIMP ReBar::put_HighlightColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_RB_HIGHLIGHTCOLOR);
	if(properties.highlightColor != newValue) {
		properties.highlightColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
			if(SendMessage(RB_GETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme))) {
				colorScheme.clrBtnHighlight = (properties.highlightColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.highlightColor));
				SendMessage(RB_SETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme));
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_RB_HIGHLIGHTCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = NULL;
	switch(imageList) {
		case ilBands:
			if(IsWindow()) {
				REBARINFO barInfo = {sizeof(REBARINFO), RBIM_IMAGELIST, NULL};
				if(SendMessage(RB_GETBARINFO, 0, reinterpret_cast<LPARAM>(&barInfo))) {
					*pValue = HandleToLong(barInfo.himl);
				} else {
					return E_FAIL;
				}
			}
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}
	return S_OK;
}

STDMETHODIMP ReBar::put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_RB_HIMAGELIST);
	BOOL fireViewChange = TRUE;
	switch(imageList) {
		case ilBands:
			if(IsWindow()) {
				REBARINFO barInfo = {sizeof(REBARINFO), RBIM_IMAGELIST, static_cast<HIMAGELIST>(LongToHandle(newValue))};
				if(!SendMessage(RB_SETBARINFO, 0, reinterpret_cast<LPARAM>(&barInfo))) {
					return E_FAIL;
				}
			}
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	FireOnChanged(DISPID_RB_HIMAGELIST);
	if(fireViewChange) {
		FireViewChange();
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP ReBar::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_RB_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RB_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_hPalette(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HPALETTE>(SendMessage(RB_GETPALETTE, 0, 0)));
	}
	return S_OK;
}

STDMETHODIMP ReBar::put_hPalette(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_RB_HPALETTE);
	if(IsWindow()) {
		SendMessage(RB_SETPALETTE, 0, reinterpret_cast<LPARAM>(LongToHandle(newValue)));
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RB_HPALETTE);
	return S_OK;
}

STDMETHODIMP ReBar::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP ReBar::get_hWndNotificationReceiver(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		HWND h = reinterpret_cast<HWND>(SendMessage(RB_SETPARENT, NULL, 0));
		SendMessage(RB_SETPARENT, reinterpret_cast<WPARAM>(h), 0);
		*pValue = HandleToLong(h);
	}
	return S_OK;
}

STDMETHODIMP ReBar::put_hWndNotificationReceiver(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_RB_HWNDNOTIFICATIONRECEIVER);
	if(IsWindow()) {
		SendMessage(RB_SETPARENT, reinterpret_cast<WPARAM>(LongToHandle(newValue)), 0);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RB_HWNDNOTIFICATIONRECEIVER);
	return S_OK;
}

/*STDMETHODIMP ReBar::get_hWndToolTip(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(RB_GETTOOLTIPS, 0, 0)));
	}
	return S_OK;
}

STDMETHODIMP ReBar::put_hWndToolTip(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_RB_HWNDTOOLTIP);
	if(IsWindow()) {
		SendMessage(RB_SETTOOLTIPS, reinterpret_cast<WPARAM>(LongToHandle(newValue)), 0);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RB_HWNDTOOLTIP);
	return S_OK;
}*/

STDMETHODIMP ReBar::get_IsRelease(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	#ifdef NDEBUG
		*pValue = VARIANT_TRUE;
	#else
		*pValue = VARIANT_FALSE;
	#endif
	return S_OK;
}

STDMETHODIMP ReBar::get_MDIFrameMenuBand(IReBarBand** ppMenuBand)
{
	ATLASSERT_POINTER(ppMenuBand, IReBarBand*);
	if(!ppMenuBand) {
		return E_POINTER;
	}

	*ppMenuBand = NULL;
	if(properties.replaceMDIFrameMenu == rmfmFullReplace && menuBandWindow.IsWindow()) {
		// the band's index can change, so use the child window handle (which actually can change, too :P) to find the band
		int bandIndex = FindBandByChildWindow(menuBandWindow);
		if(bandIndex >= 0) {
			ClassFactory::InitReBarBand(bandIndex, this, IID_IReBarBand, reinterpret_cast<LPUNKNOWN*>(ppMenuBand), FALSE);
		}
	}

	return S_OK;
}

STDMETHODIMP ReBar::putref_MDIFrameMenuBand(IReBarBand* pNewMenuBand)
{
	PUTPROPPROLOG(DISPID_RB_MDIFRAMEMENUBAND);

	HRESULT hr = E_FAIL;

	HWND hWndChild = NULL;
	if(pNewMenuBand) {
		OLE_HANDLE h = NULL;
		pNewMenuBand->get_hContainedWindow(&h);
		hWndChild = static_cast<HWND>(LongToHandle(h));
	}
	if(hWndChild == menuBandWindow) {
		return S_OK;
	}

	if(flags.customizerOfMdiMessageHook) {
		// force size changes for the old and new child window and redraw everything
		CWindowEx2 oldMenuBandChildWindow;
		CWindowEx2 newMenuBandChildWindow;
		if(menuBandWindow.IsWindow()) {
			oldMenuBandChildWindow.Attach(menuBandWindow);
			ATLVERIFY(RemoveWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this)));
			menuBandWindow.Detach();
		}
		if(::IsWindow(hWndChild)) {
			newMenuBandChildWindow.Attach(hWndChild);
		}
		if(oldMenuBandChildWindow.IsWindow()) {
			RECT windowRectangle = {0};
			oldMenuBandChildWindow.GetWindowRect(&windowRectangle);
			::MapWindowPoints(HWND_DESKTOP, *this, reinterpret_cast<LPPOINT>(&windowRectangle), 2);
			oldMenuBandChildWindow.InternalSetRedraw(FALSE);
			oldMenuBandChildWindow.SetWindowPos(NULL, 0, 0, 1, 1, SWP_NOZORDER | SWP_NOMOVE);
			oldMenuBandChildWindow.SetWindowPos(NULL, &windowRectangle, SWP_NOZORDER | SWP_NOMOVE);
			oldMenuBandChildWindow.InternalSetRedraw(TRUE);
		}
		if(newMenuBandChildWindow.IsWindow()) {
			menuBandWindow.Attach(newMenuBandChildWindow);
			ATLVERIFY(SetWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this), NULL));
			UpdateMDINonClientAreaSizes();
			RECT windowRectangle = {0};
			newMenuBandChildWindow.GetWindowRect(&windowRectangle);
			::MapWindowPoints(HWND_DESKTOP, *this, reinterpret_cast<LPPOINT>(&windowRectangle), 2);
			newMenuBandChildWindow.InternalSetRedraw(FALSE);
			newMenuBandChildWindow.SetWindowPos(NULL, 0, 0, 1, 1, SWP_NOZORDER | SWP_NOMOVE);
			newMenuBandChildWindow.SetWindowPos(NULL, &windowRectangle, SWP_NOZORDER | SWP_NOMOVE);
			newMenuBandChildWindow.InternalSetRedraw(TRUE);
			newMenuBandChildWindow.RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW);
		}
		hr = S_OK;
	}
	FireOnChanged(DISPID_RB_MDIFRAMEMENUBAND);
	return hr;
}

STDMETHODIMP ReBar::get_MouseIcon(IPictureDisp** ppMouseIcon)
{
	ATLASSERT_POINTER(ppMouseIcon, IPictureDisp*);
	if(!ppMouseIcon) {
		return E_POINTER;
	}

	if(*ppMouseIcon) {
		(*ppMouseIcon)->Release();
		*ppMouseIcon = NULL;
	}
	if(properties.mouseIcon.pPictureDisp) {
		properties.mouseIcon.pPictureDisp->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(ppMouseIcon));
	}
	return S_OK;
}

STDMETHODIMP ReBar::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_RB_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			// clone the picture by storing it into a stream
			CComQIPtr<IPersistStream, &IID_IPersistStream> pPersistStream(pNewMouseIcon);
			if(pPersistStream) {
				ULARGE_INTEGER pictureSize = {0};
				pPersistStream->GetSizeMax(&pictureSize);
				HGLOBAL hGlobalMem = GlobalAlloc(GHND, pictureSize.LowPart);
				if(hGlobalMem) {
					CComPtr<IStream> pStream = NULL;
					CreateStreamOnHGlobal(hGlobalMem, TRUE, &pStream);
					if(pStream) {
						if(SUCCEEDED(pPersistStream->Save(pStream, FALSE))) {
							LARGE_INTEGER startPosition = {0};
							pStream->Seek(startPosition, STREAM_SEEK_SET, NULL);
							OleLoadPicture(pStream, startPosition.LowPart, FALSE, IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
						}
						pStream.Release();
					}
					GlobalFree(hGlobalMem);
				}
			}
		}
		properties.mouseIcon.StartWatching();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ReBar::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_RB_MOUSEICON);
	if(properties.mouseIcon.pPictureDisp != pNewMouseIcon) {
		properties.mouseIcon.StopWatching();
		if(properties.mouseIcon.pPictureDisp) {
			properties.mouseIcon.pPictureDisp->Release();
			properties.mouseIcon.pPictureDisp = NULL;
		}
		if(pNewMouseIcon) {
			pNewMouseIcon->QueryInterface(IID_IPictureDisp, reinterpret_cast<LPVOID*>(&properties.mouseIcon.pPictureDisp));
		}
		properties.mouseIcon.StartWatching();
	} else if(pNewMouseIcon) {
		pNewMouseIcon->AddRef();
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_RB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ReBar::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP ReBar::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_RB_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RB_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_NativeDropTarget(LPUNKNOWN* ppValue)
{
	ATLASSERT_POINTER(ppValue, LPUNKNOWN);
	if(!ppValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SendMessage(RB_GETDROPTARGET, 0, reinterpret_cast<LPARAM>(ppValue));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBar::get_NumberOfRows(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = SendMessage(RB_GETROWCOUNT, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBar::get_Orientation(OrientationConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OrientationConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.orientation = ((GetStyle() & CCS_VERT) == CCS_VERT ? oVertical : oHorizontal);
	}

	*pValue = properties.orientation;
	return S_OK;
}

STDMETHODIMP ReBar::put_Orientation(OrientationConstants newValue)
{
	PUTPROPPROLOG(DISPID_RB_ORIENTATION);
	if(properties.orientation != newValue) {
		properties.orientation = newValue;
		SetDirty(TRUE);

		RECT windowRect = m_rcPos;
		SIZEL himetric = {m_sizeExtent.cx, m_sizeExtent.cy};
		SIZEL pixels = {0};
		AtlHiMetricToPixel(&himetric, &pixels);
		windowRect.right = windowRect.left + pixels.cy;
		windowRect.bottom = windowRect.top + pixels.cx;
		if(m_spInPlaceSite) {
			m_spInPlaceSite->OnPosRectChange(&windowRect);
		}

		if(IsWindow()) {
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_RB_ORIENTATION);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ReBar::get_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RegisterForOLEDragDropConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if((GetStyle() & RBS_REGISTERDROP) == RBS_REGISTERDROP) {
			properties.registerForOLEDragDrop = rfoddNativeDragDrop;
		}
	}

	*pValue = properties.registerForOLEDragDrop;
	return S_OK;
}

STDMETHODIMP ReBar::put_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants newValue)
{
	PUTPROPPROLOG(DISPID_RB_REGISTERFOROLEDRAGDROP);
	if(properties.registerForOLEDragDrop != newValue) {
		properties.registerForOLEDragDrop = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			ModifyStyle(RBS_REGISTERDROP, 0);
			RevokeDragDrop(*this);
			switch(properties.registerForOLEDragDrop) {
				case rfoddNativeDragDrop:
					ModifyStyle(0, RBS_REGISTERDROP);
					break;
				case rfoddAdvancedDragDrop: {
					ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
					break;
				}
			}
		}
		FireOnChanged(DISPID_RB_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_ReplaceMDIFrameMenu(ReplaceMDIFrameMenuConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ReplaceMDIFrameMenuConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.replaceMDIFrameMenu;
	return S_OK;
}

STDMETHODIMP ReBar::put_ReplaceMDIFrameMenu(ReplaceMDIFrameMenuConstants newValue)
{
	PUTPROPPROLOG(DISPID_RB_REPLACEMDIFRAMEMENU);

	if(!IsInDesignMode() && mdiClient.IsWindow()) {
		// Set not supported at runtime - raise VB runtime error 382
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 382);
	}
	if(properties.replaceMDIFrameMenu != newValue) {
		properties.replaceMDIFrameMenu = newValue;
		SetDirty(TRUE);

		FireOnChanged(DISPID_RB_REPLACEMDIFRAMEMENU);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_RightToLeft(RightToLeftConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RightToLeftConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.rightToLeft = static_cast<RightToLeftConstants>(0);
		DWORD style = GetExStyle();
		if(style & WS_EX_LAYOUTRTL) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlLayout);
		}
		if(style & WS_EX_RTLREADING) {
			properties.rightToLeft = static_cast<RightToLeftConstants>(static_cast<long>(properties.rightToLeft) | rtlText);
		}
	}

	*pValue = properties.rightToLeft;
	return S_OK;
}

STDMETHODIMP ReBar::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_RB_RIGHTTOLEFT);
	if(properties.rightToLeft != newValue) {
		properties.rightToLeft = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.rightToLeft & rtlLayout) {
				ModifyStyleEx(0, WS_EX_LAYOUTRTL);
			} else {
				ModifyStyleEx(WS_EX_LAYOUTRTL, 0);
			}
			if(properties.rightToLeft & rtlText) {
				ModifyStyleEx(0, WS_EX_RTLREADING);
			} else {
				ModifyStyleEx(WS_EX_RTLREADING, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_ShadowColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
		if(SendMessage(RB_GETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme))) {
			if(colorScheme.clrBtnShadow == CLR_DEFAULT) {
				properties.shadowColor = static_cast<OLE_COLOR>(-1);
			} else if(colorScheme.clrBtnShadow != OLECOLOR2COLORREF(properties.shadowColor)) {
				properties.shadowColor = colorScheme.clrBtnShadow;
			}
		}
	}

	*pValue = properties.shadowColor;
	return S_OK;
}

STDMETHODIMP ReBar::put_ShadowColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_RB_SHADOWCOLOR);
	if(properties.shadowColor != newValue) {
		properties.shadowColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
			if(SendMessage(RB_GETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme))) {
				colorScheme.clrBtnShadow = (properties.shadowColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.shadowColor));
				SendMessage(RB_SETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme));
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_RB_SHADOWCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP ReBar::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_RB_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ReBar::get_ToggleOnDoubleClick(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.toggleOnDoubleClick = ((GetStyle() & RBS_DBLCLKTOGGLE) == RBS_DBLCLKTOGGLE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.toggleOnDoubleClick);
	return S_OK;
}

STDMETHODIMP ReBar::put_ToggleOnDoubleClick(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_TOGGLEONDOUBLECLICK);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.toggleOnDoubleClick != b) {
		properties.toggleOnDoubleClick = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.toggleOnDoubleClick) {
				ModifyStyle(0, RBS_DBLCLKTOGGLE);
			} else {
				ModifyStyle(RBS_DBLCLKTOGGLE, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_TOGGLEONDOUBLECLICK);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP ReBar::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_RB_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP ReBar::get_Version(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	TCHAR pBuffer[50];
	ATLVERIFY(SUCCEEDED(StringCbPrintf(pBuffer, 50 * sizeof(TCHAR), TEXT("%i.%i.%i.%i"), VERSION_MAJOR, VERSION_MINOR, VERSION_REVISION1, VERSION_BUILD)));
	*pValue = CComBSTR(pBuffer);
	return S_OK;
}

STDMETHODIMP ReBar::get_VerticalSizingGripsOnVerticalOrientation(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.verticalSizingGripsOnVerticalOrientation = ((GetStyle() & RBS_VERTICALGRIPPER) == RBS_VERTICALGRIPPER);
	}

	*pValue = BOOL2VARIANTBOOL(properties.verticalSizingGripsOnVerticalOrientation);
	return S_OK;
}

STDMETHODIMP ReBar::put_VerticalSizingGripsOnVerticalOrientation(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_RB_VERTICALSIZINGGRIPSONVERTICALORIENTATION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.verticalSizingGripsOnVerticalOrientation != b) {
		properties.verticalSizingGripsOnVerticalOrientation = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.verticalSizingGripsOnVerticalOrientation) {
				ModifyStyle(0, RBS_VERTICALGRIPPER);
			} else {
				ModifyStyle(RBS_VERTICALGRIPPER, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_RB_VERTICALSIZINGGRIPSONVERTICALORIENTATION);
	}
	return S_OK;
}

STDMETHODIMP ReBar::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP ReBar::BeginDragBand(IReBarBand* pBandToDrag, LONG xMousePosition/* = -1*/, LONG yMousePosition/* = -1*/)
{
	ATLASSERT_POINTER(pBandToDrag, IReBarBand);
	if(!pBandToDrag) {
		return E_INVALIDARG;
	}

	if(IsWindow()) {
		DWORD pos;
		if(xMousePosition == -1 || yMousePosition == -1) {
			pos = static_cast<DWORD>(-1);
		} else if(xMousePosition == -2 || yMousePosition == -2) {
			pos = static_cast<DWORD>(-2);
		} else {
			pos = static_cast<DWORD>(MAKELONG(xMousePosition, yMousePosition));
		}
		LONG bandIndex = 0;
		if(SUCCEEDED(pBandToDrag->get_Index(&bandIndex))) {
			SendMessage(RB_BEGINDRAG, bandIndex, pos);
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ReBar::DragMoveBand(LONG xMousePosition/* = -1*/, LONG yMousePosition/* = -1*/)
{
	if(IsWindow()) {
		DWORD pos;
		if(xMousePosition == -1 || yMousePosition == -1) {
			pos = static_cast<DWORD>(-1);
		} else {
			pos = static_cast<DWORD>(MAKELONG(xMousePosition, yMousePosition));
		}
		SendMessage(RB_BEGINDRAG, 0, pos);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBar::EndDragBand(void)
{
	if(IsWindow()) {
		SendMessage(RB_ENDDRAG, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBar::FinishOLEDragDrop(void)
{
	if(dragDropStatus.pDropTargetHelper && dragDropStatus.drop_pDataObject) {
		dragDropStatus.pDropTargetHelper->Drop(dragDropStatus.drop_pDataObject, &dragDropStatus.drop_mousePosition, dragDropStatus.drop_effect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
		return S_OK;
	}
	// Can't perform requested operation - raise VB runtime error 17
	return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 17);
}

STDMETHODIMP ReBar::GetMargins(OLE_XSIZE_PIXELS* pLeftMargin/* = NULL*/, OLE_YSIZE_PIXELS* pTopMargin/* = NULL*/, OLE_XSIZE_PIXELS* pRightMargin/* = NULL*/, OLE_YSIZE_PIXELS* pBottomMargin/* = NULL*/)
{
	HRESULT hr = E_FAIL;
	if(RunTimeHelper::IsCommCtrl6()) {
		MARGINS margins = {0};
		SendMessage(RB_GETBANDMARGINS, 0, reinterpret_cast<LPARAM>(&margins));
		if(pLeftMargin) {
			*pLeftMargin = margins.cxLeftWidth;
		}
		if(pTopMargin) {
			*pTopMargin = margins.cyTopHeight;
		}
		if(pRightMargin) {
			*pRightMargin = margins.cxRightWidth;
		}
		if(pBottomMargin) {
			*pBottomMargin = margins.cyBottomHeight;
		}
		hr = S_OK;
	}
	return hr;
}

STDMETHODIMP ReBar::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IReBarBand** ppHitBand)
{
	ATLASSERT_POINTER(ppHitBand, IReBarBand*);
	if(!ppHitBand) {
		return E_POINTER;
	}

	if(IsWindow()) {
		UINT hitTestFlags = static_cast<UINT>(*pHitTestDetails);
		int bandIndex = HitTest(x, y, &hitTestFlags/*, TRUE*/);

		if(pHitTestDetails) {
			*pHitTestDetails = static_cast<HitTestConstants>(hitTestFlags);
		}
		ClassFactory::InitReBarBand(bandIndex, this, IID_IReBarBand, reinterpret_cast<LPUNKNOWN*>(ppHitBand));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ReBar::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// open the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_READ | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// read settings
		if(Load(pStream) == S_OK) {
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP ReBar::Refresh(void)
{
	if(IsWindow()) {
		Invalidate();
		RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE);
	}
	return S_OK;
}

STDMETHODIMP ReBar::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
{
	ATLASSERT_POINTER(pSucceeded, VARIANT_BOOL);
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	// create the file
	COLE2T converter(file);
	LPTSTR pFilePath = converter;
	CComPtr<IStream> pStream = NULL;
	HRESULT hr = SHCreateStreamOnFile(pFilePath, STGM_CREATE | STGM_WRITE | STGM_SHARE_DENY_WRITE, &pStream);
	if(SUCCEEDED(hr) && pStream) {
		// write settings
		if(Save(pStream, FALSE) == S_OK) {
			if(FAILED(pStream->Commit(STGC_DEFAULT))) {
				return S_OK;
			}
			*pSucceeded = VARIANT_TRUE;
		}
	}
	return S_OK;
}

STDMETHODIMP ReBar::SizeToRectangle(OLE_XPOS_PIXELS xLeft, OLE_YPOS_PIXELS yTop, OLE_XPOS_PIXELS xRight, OLE_YPOS_PIXELS yBottom, VARIANT_BOOL* pLayoutChanged)
{
	ATLASSERT_POINTER(pLayoutChanged, VARIANT_BOOL);
	if(!pLayoutChanged) {
		return E_POINTER;
	}
	*pLayoutChanged = VARIANT_FALSE;

	if(IsWindow()) {
		RECT rc = {xLeft, yTop, xRight, yBottom};
		*pLayoutChanged = BOOL2VARIANTBOOL(SendMessage(RB_SIZETORECT, 0, reinterpret_cast<LPARAM>(&rc)));
		return S_OK;
	}
	return E_FAIL;
}


LRESULT ReBar::OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_ContextMenu(button, shift, mousePosition.x, mousePosition.y);
	return 0;
}

LRESULT ReBar::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();

		Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));
	}
	return lr;
}

LRESULT ReBar::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT ReBar::OnForwardMsg(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = 0;
	if(flags.customizerOfMdiMessageHook) {
		LPMSG pMessage = reinterpret_cast<LPMSG>(lParam);
		ProcessWindowMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam, lr, 4);
	}
	return lr;
}

LRESULT ReBar::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<ReBar>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	return lr;
}

LRESULT ReBar::OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_DblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ReBar::OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	/* HACK: Fixes the Validate event of VB6. RebarWindow32 doesn't seem to get the focus on WM_MOUSEACTIVATE
	 * by default. If we call SetFocus on WM_MOUSEACTIVATE, the Validate event is not fired for some reason.
	 * Therefore we call SetFocus here.
	 */
	SetFocus();
	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		mouseStatus.StoreClickCandidate(1/*MouseButtonConstants.vbLeftButton*/);
	}

	return 0;
}

LRESULT ReBar::OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ReBar::OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ReBar::OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	/* HACK: Fixes the Validate event of VB6. RebarWindow32 doesn't seem to get the focus on WM_MOUSEACTIVATE
	 * by default. If we call SetFocus on WM_MOUSEACTIVATE, the Validate event is not fired for some reason.
	 * Therefore we call SetFocus here.
	 */
	SetFocus();
	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		mouseStatus.StoreClickCandidate(4/*MouseButtonConstants.vbMiddleButton*/);
	}

	return 0;
}

LRESULT ReBar::OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ReBar::OnMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = 0;
	VARIANT_BOOL handled = VARIANT_FALSE;
	if(!(properties.disabledEvents & deRawMenuMessage)) {
		LONG result = lr;
		Raise_RawMenuMessage(message, wParam, lParam, &result, &handled);
		lr = result;
	}
	wasHandled = VARIANTBOOL2BOOL(handled);
	return lr;
}

LRESULT ReBar::OnMenuSelect(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = 1;
	VARIANT_BOOL handled = VARIANT_FALSE;
	if(!(properties.disabledEvents & deRawMenuMessage)) {
		LONG result = lr;
		Raise_RawMenuMessage(message, wParam, lParam, &result, &handled);
		lr = result;
	}

	Raise_SelectedMenuItem(LOWORD(wParam), static_cast<MenuItemStateConstants>(HIWORD(wParam)), HandleToLong(reinterpret_cast<HMENU>(lParam)));

	wasHandled = VARIANTBOOL2BOOL(handled);
	return lr;
}

LRESULT ReBar::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(m_bInPlaceActive && !m_bUIActive) {
		flags.uiActivationPending = TRUE;
	} else {
		wasHandled = FALSE;
	}
	return MA_ACTIVATE;
}

LRESULT ReBar::OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseHover(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ReBar::OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_MouseLeave(button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);
	}

	return 0;
}

LRESULT ReBar::OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ReBar::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ReBar::OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ReBar::OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	/* HACK: Fixes the Validate event of VB6. RebarWindow32 doesn't seem to get the focus on WM_MOUSEACTIVATE
	 * by default. If we call SetFocus on WM_MOUSEACTIVATE, the Validate event is not fired for some reason.
	 * Therefore we call SetFocus here.
	 */
	SetFocus();
	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		mouseStatus.StoreClickCandidate(2/*MouseButtonConstants.vbRightButton*/);
	}

	return 0;
}

LRESULT ReBar::OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ReBar::OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	CRect clientArea;
	GetClientRect(&clientArea);
	ClientToScreen(&clientArea);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == *this) {
			setCursor = TRUE;
		}
	}

	if(setCursor) {
		if(properties.mousePointer == mpCustom) {
			if(properties.mouseIcon.pPictureDisp) {
				CComQIPtr<IPicture, &IID_IPicture> pPicture(properties.mouseIcon.pPictureDisp);
				if(pPicture) {
					OLE_HANDLE h = NULL;
					pPicture->get_Handle(&h);
					hCursor = static_cast<HCURSOR>(LongToHandle(h));
				}
			}
		} else {
			hCursor = MousePointerConst2hCursor(properties.mousePointer);
		}

		if(hCursor) {
			SetCursor(hCursor);
			return TRUE;
		}
	}

	wasHandled = FALSE;
	return FALSE;
}

LRESULT ReBar::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<ReBar>::OnSetFocus(message, wParam, lParam, wasHandled);
	if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
		flags.uiActivationPending = FALSE;

		// now execute what usually would have been done on WM_MOUSEACTIVATE
		BOOL dummy = TRUE;
		CComControl<ReBar>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
	}
	return lr;
}

LRESULT ReBar::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_RB_FONT) == S_FALSE) {
		return 0;
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(!properties.font.dontGetFontObject) {
		// this message wasn't sent by ourselves, so we have to get the new font.pFontDisp object
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(reinterpret_cast<HFONT>(wParam));
		properties.font.owningFontResource = FALSE;
		properties.useSystemFont = FALSE;
		properties.font.StopWatching();

		if(properties.font.pFontDisp) {
			properties.font.pFontDisp->Release();
			properties.font.pFontDisp = NULL;
		}
		if(!properties.font.currentFont.IsNull()) {
			LOGFONT logFont = {0};
			int bytes = properties.font.currentFont.GetLogFont(&logFont);
			if(bytes) {
				FONTDESC font = {0};
				CT2OLE converter(logFont.lfFaceName);

				HDC hDC = GetDC();
				if(hDC) {
					LONG fontHeight = logFont.lfHeight;
					if(fontHeight < 0) {
						fontHeight = -fontHeight;
					}

					int pixelsPerInch = GetDeviceCaps(hDC, LOGPIXELSY);
					ReleaseDC(hDC);
					font.cySize.Lo = fontHeight * 720000 / pixelsPerInch;
					font.cySize.Hi = 0;

					font.lpstrName = converter;
					font.sWeight = static_cast<SHORT>(logFont.lfWeight);
					font.sCharset = logFont.lfCharSet;
					font.fItalic = logFont.lfItalic;
					font.fUnderline = logFont.lfUnderline;
					font.fStrikethrough = logFont.lfStrikeOut;
				}
				font.cbSizeofstruct = sizeof(FONTDESC);

				OleCreateFontIndirect(&font, IID_IFontDisp, reinterpret_cast<LPVOID*>(&properties.font.pFontDisp));
			}
		}
		properties.font.StartWatching();

		SetDirty(TRUE);
		FireOnChanged(DISPID_RB_FONT);
	}

	return lr;
}

LRESULT ReBar::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == 71216) {
		// the message was sent by ourselves
		lParam = 0;
		if(wParam) {
			// We're gonna activate redrawing - does the client allow this?
			if(properties.dontRedraw) {
				// no, so eat this message
				return 0;
			}
		}
	} else {
		// TODO: Should we really do this?
		properties.dontRedraw = !wParam;
	}

	return DefWindowProc(message, wParam, lParam);
}

LRESULT ReBar::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETICONTITLELOGFONT) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	} else if(wParam == SPI_SETNONCLIENTMETRICS) {
		if(menuBandWindow.IsWindow()) {
			UpdateMDINonClientAreaSizes();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ReBar::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	// Microsoft couldn't make it more difficult to detect whether we should use themes or not...
	flags.usingThemes = FALSE;
	if(CTheme::IsThemingSupported() && RunTimeHelper::IsCommCtrl6()) {
		HMODULE hThemeDLL = LoadLibrary(TEXT("uxtheme.dll"));
		if(hThemeDLL) {
			typedef BOOL WINAPI IsAppThemedFn();
			IsAppThemedFn* pfnIsAppThemed = static_cast<IsAppThemedFn*>(GetProcAddress(hThemeDLL, "IsAppThemed"));
			if(pfnIsAppThemed()) {
				flags.usingThemes = TRUE;
			}
			FreeLibrary(hThemeDLL);
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ReBar::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	switch(wParam) {
		case timers.ID_REDRAW:
			if(IsWindowVisible()) {
				KillTimer(timers.ID_REDRAW);
				SetRedraw(!properties.dontRedraw);
			} else {
				// wait... (this fixes visibility problems if another control displays a nag screen)
			}
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ReBar::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	LPWINDOWPOS pDetails = reinterpret_cast<LPWINDOWPOS>(lParam);

	CRect windowRectangle = m_rcPos;

	/* Ugly hack: We depend on this message being sent without SWP_NOMOVE at least once, but this requirement
	              not always will be fulfilled. Fortunately pDetails seems to contain correct x and y values
	              even if SWP_NOMOVE is set.
	 */
	if(!(pDetails->flags & SWP_NOMOVE) || (windowRectangle.IsRectNull()/* && pDetails->x != 0 && pDetails->y != 0*/)) {
		windowRectangle.MoveToXY(pDetails->x, pDetails->y);
	}
	if(!(pDetails->flags & SWP_NOSIZE)) {
		windowRectangle.right = windowRectangle.left + pDetails->cx;
		windowRectangle.bottom = windowRectangle.top + pDetails->cy;
	}

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos)) {
			m_spInPlaceSite->OnPosRectChange(&windowRectangle);
		}
		if(!(pDetails->flags & SWP_NOSIZE)) {
			/* Problem: When the control is resized, m_rcPos already contains the new rectangle, even before the
			 *          message is sent without SWP_NOSIZE. Therefore raise the event even if the rectangles are
			 *          equal. Raising the event too often is better than raising it too few.
			 */
			Raise_ResizedControlWindow();
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ReBar::OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_XDblClick(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ReBar::OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	/* HACK: Fixes the Validate event of VB6. RebarWindow32 doesn't seem to get the focus on WM_MOUSEACTIVATE
	 * by default. If we call SetFocus on WM_MOUSEACTIVATE, the Validate event is not fired for some reason.
	 * Therefore we call SetFocus here.
	 */
	SetFocus();
	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		Raise_MouseDown(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		SHORT button = 0;
		switch(GET_XBUTTON_WPARAM(wParam)) {
			case XBUTTON1:
				button = embXButton1;
				break;
			case XBUTTON2:
				button = embXButton2;
				break;
		}
		mouseStatus.StoreClickCandidate(button);
	}

	return 0;
}

LRESULT ReBar::OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(GET_KEYSTATE_WPARAM(wParam), button, shift);
	switch(GET_XBUTTON_WPARAM(wParam)) {
		case XBUTTON1:
			button = embXButton1;
			break;
		case XBUTTON2:
			button = embXButton2;
			break;
	}
	Raise_MouseUp(button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ReBar::OnGetRebar(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = NULL;
	LONG controlType = 0;
	if(wParam) {
		// if the ToolBar control has sent us this message, wParam will be 0x02
		controlType = *reinterpret_cast<PLONG>(wParam);
	}

	if(controlType == 0 && menuBandWindow.IsWindow()) {
		lr = menuBandWindow.SendMessage(message, wParam, lParam);
	}
	// If we don't need this message, don't prevent others from getting asked whether they want this message.
	if(!lr && flags.customizerOfMdiMessageHook && IsWindowVisible()) {
		if(wParam) {
			*reinterpret_cast<PLONG>(wParam) = 0x01;
		}
		if(lParam) {
			*reinterpret_cast<LPBOOL>(lParam) = flags.customizerOfMdiMessageHook;
		}
		HWND hWnd = *this;
		lr = reinterpret_cast<LRESULT>(hWnd);
	}
	return lr;
}

LRESULT ReBar::OnAutoPopup(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	if(IsWindowVisible() && menuBandWindow.IsWindow()) {
		BOOL rightToLeft = (menuBandWindow.GetExStyle() & WS_EX_LAYOUTRTL);

		RECT windowRectangle = {0};
		menuBandWindow.GetWindowRect(&windowRectangle);

		RECT iconRectangle = {0};
		CalcMDIChildIconRectangle(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, &iconRectangle, rightToLeft);

		CMenuHandle menu = ::GetSystemMenu(mdiStatus.hWndChildMaximized, FALSE);
		UINT menuID = 0;
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_BeforeDisplayMDIChildSystemMenu(HandleToLong(mdiStatus.hWndChildMaximized), HandleToLong(static_cast<HMENU>(menu)), rightToLeft ? windowRectangle.right : windowRectangle.left, windowRectangle.bottom, &cancel);
		if(cancel == VARIANT_FALSE) {
			menuID = static_cast<UINT>(menu.TrackPopupMenu(TPM_LEFTBUTTON | TPM_VERTICAL | TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_VERPOSANIMATION, rightToLeft ? windowRectangle.right : windowRectangle.left, windowRectangle.bottom, *this));
		}

		// eat next message if click is on the same button
		OffsetRect(&iconRectangle, windowRectangle.left, windowRectangle.top);
		MSG msg = {0};
		if(PeekMessage(&msg, menuBandWindow, WM_NCLBUTTONDOWN, WM_NCLBUTTONDOWN, PM_NOREMOVE) && PtInRect(&iconRectangle, msg.pt)) {
			PeekMessage(&msg, menuBandWindow, WM_NCLBUTTONDOWN, WM_NCLBUTTONDOWN, PM_REMOVE);
		}

		if(menuID != 0) {
			::SendMessage(mdiStatus.hWndChildMaximized, WM_SYSCOMMAND, menuID, 0);
		}
		Raise_CleanupMDIChildSystemMenu(HandleToLong(mdiStatus.hWndChildMaximized), HandleToLong(static_cast<HMENU>(menu)));
	}
	return 0;
}

LRESULT ReBar::OnReflectedNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case NM_CUSTOMDRAW:
			if(reinterpret_cast<LPNMHDR>(lParam)->hwndFrom == *this) {
				return OnCustomDrawNotification(message, wParam, lParam, wasHandled);
			}
		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ReBar::OnReflectedNotifyFormat(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam == NF_QUERY) {
		#ifdef UNICODE
			return NFR_UNICODE;
		#else
			return NFR_ANSI;
		#endif
	}
	return 0;
}

LRESULT ReBar::OnDeleteBand(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IReBarBand> pRBBandItf = NULL;
	CComObject<ReBarBand>* pRBBandObj = NULL;
	if(!(properties.disabledEvents & deBandDeletionEvents)) {
		CComObject<ReBarBand>::CreateInstance(&pRBBandObj);
		pRBBandObj->AddRef();
		pRBBandObj->SetOwner(this);
		pRBBandObj->Attach(wParam);
		pRBBandObj->QueryInterface(IID_IReBarBand, reinterpret_cast<LPVOID*>(&pRBBandItf));
		pRBBandObj->Release();
		Raise_RemovingBand(pRBBandItf, &cancel);
	}

	if(cancel == VARIANT_FALSE) {
		CComPtr<IVirtualReBarBand> pVRBBandItf = NULL;
		if(!(properties.disabledEvents & deBandDeletionEvents)) {
			CComObject<VirtualReBarBand>* pVRBBandObj = NULL;
			CComObject<VirtualReBarBand>::CreateInstance(&pVRBBandObj);
			pVRBBandObj->AddRef();
			pVRBBandObj->SetOwner(this);
			if(pRBBandObj) {
				pRBBandObj->SaveState(pVRBBandObj);
			}

			pVRBBandObj->QueryInterface(IID_IVirtualReBarBand, reinterpret_cast<LPVOID*>(&pVRBBandItf));
			pVRBBandObj->Release();
		}

		if(!(properties.disabledEvents & deMouseEvents)) {
			if(static_cast<int>(wParam) == bandUnderMouse) {
				// we're removing the band below the mouse cursor
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);

				UINT hitTestDetails = 0;
				HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				Raise_BandMouseLeave(pRBBandItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				bandUnderMouse = -1;
			}
			if(static_cast<int>(wParam) == bandChevronPushed) {
				bandChevronPushed = -1;
			}
		}

		// finally pass the message to the rebar
		lr = DefWindowProc(message, wParam, lParam);
		if(lr) {
			RebuildAcceleratorTable();

			if(!(properties.disabledEvents & deBandDeletionEvents)) {
				Raise_RemovedBand(pVRBBandItf);
			}
		}

		if(!(properties.disabledEvents & deMouseEvents)) {
			if(lr) {
				// maybe we have a new band below the mouse cursor now
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);

				UINT hitTestDetails = 0;
				int newBandUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
				if(newBandUnderMouse != bandUnderMouse) {
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					if(bandUnderMouse != -1) {
						pRBBandItf = ClassFactory::InitReBarBand(bandUnderMouse, this);
						Raise_BandMouseLeave(pRBBandItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}

					bandUnderMouse = newBandUnderMouse;

					if(bandUnderMouse != -1) {
						pRBBandItf = ClassFactory::InitReBarBand(bandUnderMouse, this);
						Raise_BandMouseEnter(pRBBandItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}

	return lr;
}

LRESULT ReBar::OnGetBandInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPREBARBANDINFO pBandData = reinterpret_cast<LPREBARBANDINFO>(lParam);
	#ifdef DEBUG
		if(message == RB_GETBANDINFOA) {
			ATLASSERT_POINTER(pBandData, REBARBANDINFOA);
		} else {
			ATLASSERT_POINTER(pBandData, REBARBANDINFOW);
		}
	#endif
	if(!pBandData || !(pBandData->fMask & RBBIM_CHEVRONLOCATION) || pBandData->cbSize <= REBARBANDINFO_V6_SIZE || IsComctl32Version610OrNewer()) {
		return DefWindowProc(message, wParam, lParam);
	}

	LRESULT lr = TRUE;

	REBARBANDINFO band = {0};
	CopyMemory(&band, pBandData, min(pBandData->cbSize, sizeof(REBARBANDINFO)));
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask &= ~RBBIM_CHEVRONLOCATION;
	if(band.fMask) {
		lr = DefWindowProc(message, wParam, reinterpret_cast<LPARAM>(&band));
		UINT size = pBandData->cbSize;
		CopyMemory(pBandData, &band, min(pBandData->cbSize, sizeof(REBARBANDINFO)));
		pBandData->cbSize = size;
		pBandData->fMask |= RBBIM_CHEVRONLOCATION;
	}

	if(lr) {
		BOOL chevronVisible = FALSE;
		ZeroMemory(&band, RunTimeHelper::SizeOf_REBARBANDINFO());
		band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
		band.fMask = RBBIM_CHILD | RBBIM_IDEALSIZE | RBBIM_STYLE;
		if(SendMessage(RB_GETBANDINFO, wParam, reinterpret_cast<LPARAM>(&band))) {
			if((band.fStyle & (RBBS_USECHEVRON | RBBS_FIXEDSIZE)) == RBBS_USECHEVRON) {
				RECT containedWindowRect = {0};
				::GetWindowRect(band.hwndChild, &containedWindowRect);
				if(::MapWindowPoints(HWND_DESKTOP, *this, reinterpret_cast<LPPOINT>(&containedWindowRect), 2)) {
					if((GetStyle() & CCS_VERT) == CCS_VERT) {
						chevronVisible = (band.cxIdeal > static_cast<UINT>(containedWindowRect.bottom - containedWindowRect.top));
					} else {
						chevronVisible = (band.cxIdeal > static_cast<UINT>(containedWindowRect.right - containedWindowRect.left));
					}
				}
			}
		}

		if(chevronVisible) {
			// emulate RBBIM_CHEVRONLOCATION by hit-testing
			RECT bandRectangle = {0};
			if(SendMessage(RB_GETRECT, wParam, reinterpret_cast<LPARAM>(&bandRectangle))) {
				RBHITTESTINFO hitTestInfo = {0};
				if((GetStyle() & CCS_VERT) == CCS_VERT) {
					int tmp = bandRectangle.left;
					bandRectangle.left = bandRectangle.top;
					bandRectangle.top = tmp;
					tmp = bandRectangle.right;
					bandRectangle.right = bandRectangle.bottom;
					bandRectangle.bottom = tmp;
					BOOL foundBottom = FALSE;
					hitTestInfo.pt.x = (bandRectangle.left + bandRectangle.right) >> 1;
					for(hitTestInfo.pt.y = bandRectangle.bottom; hitTestInfo.pt.y >= bandRectangle.top; hitTestInfo.pt.y--) {
						SendMessage(RB_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
						if(hitTestInfo.flags & RBHT_CHEVRON) {
							if(!foundBottom) {
								pBandData->rcChevronLocation.bottom = hitTestInfo.pt.y + 1;
								foundBottom = TRUE;
							}
						} else {
							if(foundBottom) {
								// NOTE: Only if we add 1 here, we calculate exactly the same rectangle as comctl32.dll 6.10.
								pBandData->rcChevronLocation.top = hitTestInfo.pt.y + 1;
								break;
							}
						}
					}
					hitTestInfo.pt.y = (pBandData->rcChevronLocation.top + pBandData->rcChevronLocation.bottom) >> 1;
					for(hitTestInfo.pt.x = bandRectangle.left; hitTestInfo.pt.x <= bandRectangle.right; hitTestInfo.pt.x++) {
						SendMessage(RB_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
						if(hitTestInfo.flags & RBHT_CHEVRON) {
							pBandData->rcChevronLocation.left = hitTestInfo.pt.x;
							break;
						}
					}
					for(hitTestInfo.pt.x = bandRectangle.right; hitTestInfo.pt.x >= bandRectangle.left; hitTestInfo.pt.x--) {
						SendMessage(RB_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
						if(hitTestInfo.flags & RBHT_CHEVRON) {
							pBandData->rcChevronLocation.right = hitTestInfo.pt.x + 1;
							break;
						}
					}
				} else {
					BOOL foundRight = FALSE;
					hitTestInfo.pt.y = (bandRectangle.top + bandRectangle.bottom) >> 1;
					for(hitTestInfo.pt.x = bandRectangle.right; hitTestInfo.pt.x >= bandRectangle.left; hitTestInfo.pt.x--) {
						SendMessage(RB_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
						if(hitTestInfo.flags & RBHT_CHEVRON) {
							if(!foundRight) {
								pBandData->rcChevronLocation.right = hitTestInfo.pt.x + 1;
								foundRight = TRUE;
							}
						} else {
							if(foundRight) {
								// NOTE: Only if we add 1 here, we calculate exactly the same rectangle as comctl32.dll 6.10.
								pBandData->rcChevronLocation.left = hitTestInfo.pt.x + 1;
								break;
							}
						}
					}
					hitTestInfo.pt.x = (pBandData->rcChevronLocation.left + pBandData->rcChevronLocation.right) >> 1;
					for(hitTestInfo.pt.y = bandRectangle.top; hitTestInfo.pt.y <= bandRectangle.bottom; hitTestInfo.pt.y++) {
						SendMessage(RB_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
						if(hitTestInfo.flags & RBHT_CHEVRON) {
							pBandData->rcChevronLocation.top = hitTestInfo.pt.y;
							break;
						}
					}
					for(hitTestInfo.pt.y = bandRectangle.bottom; hitTestInfo.pt.y >= bandRectangle.top; hitTestInfo.pt.y--) {
						SendMessage(RB_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo));
						if(hitTestInfo.flags & RBHT_CHEVRON) {
							pBandData->rcChevronLocation.bottom = hitTestInfo.pt.y + 1;
							break;
						}
					}
				}
			} else {
				lr = FALSE;
			}
		}
	}
	return lr;
}

LRESULT ReBar::OnInsertBand(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT succeeded = FALSE;

	if(!(properties.disabledEvents & deBandInsertionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;

		CComObject<VirtualReBarBand>* pVRBarBandObj = NULL;
		CComPtr<IVirtualReBarBand> pVRBarBandItf = NULL;
		CComObject<VirtualReBarBand>::CreateInstance(&pVRBarBandObj);
		pVRBarBandObj->AddRef();
		pVRBarBandObj->SetOwner(this);

		#ifdef UNICODE
			BOOL mustConvert = (message == RB_INSERTBANDA);
		#else
			BOOL mustConvert = (message == RB_INSERTBANDW);
		#endif
		if(mustConvert) {
			#ifdef UNICODE
				LPREBARBANDINFOA pBandData = reinterpret_cast<LPREBARBANDINFOA>(lParam);
				REBARBANDINFOW convertedBandData = {0};
				convertedBandData.cbSize = sizeof(REBARBANDINFOW);
				if(!RunTimeHelper::IsVista() || !RunTimeHelper::IsCommCtrl6()) {
					convertedBandData.cbSize = REBARBANDINFOW_V6_SIZE;
				}
				LPSTR p = NULL;
				if(pBandData->lpText != LPSTR_TEXTCALLBACKA) {
					p = pBandData->lpText;
				}
				CA2W converter(p);
				if(pBandData->lpText == LPSTR_TEXTCALLBACKA) {
					convertedBandData.lpText = LPSTR_TEXTCALLBACKW;
				} else {
					convertedBandData.lpText = converter;
				}
			#else
				LPREBARBANDINFOW pBandData = reinterpret_cast<LPREBARBANDINFOW>(lParam);
				REBARBANDINFOA convertedBandData = {0};
				convertedBandData.cbSize = sizeof(REBARBANDINFOA);
				if(!RunTimeHelper::IsVista() || !RunTimeHelper::IsCommCtrl6()) {
					convertedBandData.cbSize = REBARBANDINFOA_V6_SIZE;
				}
				LPWSTR p = NULL;
				if(pBandData->lpText != LPSTR_TEXTCALLBACKW) {
					p = pBandData->lpText;
				}
				CW2A converter(p);
				if(pBandData->lpText == LPSTR_TEXTCALLBACKW) {
					convertedBandData.lpText = LPSTR_TEXTCALLBACKA;
				} else {
					convertedBandData.lpText = converter;
				}
			#endif
			convertedBandData.cch = pBandData->cch;
			convertedBandData.fMask = pBandData->fMask;
			convertedBandData.fStyle = pBandData->fStyle;
			convertedBandData.clrFore = pBandData->clrFore;
			convertedBandData.clrBack = pBandData->clrBack;
			convertedBandData.iImage = pBandData->iImage;
			convertedBandData.hwndChild = pBandData->hwndChild;
			convertedBandData.cxMinChild = pBandData->cxMinChild;
			convertedBandData.cyMinChild = pBandData->cyMinChild;
			convertedBandData.cx = pBandData->cx;
			convertedBandData.hbmBack = pBandData->hbmBack;
			convertedBandData.wID = pBandData->wID;
			convertedBandData.cyChild = pBandData->cyChild;
			convertedBandData.cyMaxChild = pBandData->cyMaxChild;
			convertedBandData.cyIntegral = pBandData->cyIntegral;
			convertedBandData.cxIdeal = pBandData->cxIdeal;
			convertedBandData.lParam = pBandData->lParam;
			convertedBandData.cxHeader = pBandData->cxHeader;
			// the REBARBANDINFO struct may end here, so be more careful with the remaining members
			if(convertedBandData.fMask & RBBIM_CHEVRONLOCATION) {
				convertedBandData.rcChevronLocation = pBandData->rcChevronLocation;
			}
			if(convertedBandData.fMask & RBBIM_CHEVRONSTATE) {
				convertedBandData.uChevronState = pBandData->uChevronState;
			}
			pVRBarBandObj->LoadState(&convertedBandData, (wParam == -1 ? static_cast<int>(SendMessage(RB_GETBANDCOUNT, 0, 0)) : wParam));
		} else {
			pVRBarBandObj->LoadState(reinterpret_cast<LPREBARBANDINFO>(lParam), (wParam == -1 ? static_cast<int>(SendMessage(RB_GETBANDCOUNT, 0, 0)) : wParam));
		}

		pVRBarBandObj->QueryInterface(IID_IVirtualReBarBand, reinterpret_cast<LPVOID*>(&pVRBarBandItf));
		pVRBarBandObj->Release();

		Raise_InsertingBand(pVRBarBandItf, &cancel);
		pVRBarBandObj = NULL;

		if(cancel == VARIANT_FALSE) {
			// finally pass the message to the rebar
			LPREBARBANDINFO pDetails = reinterpret_cast<LPREBARBANDINFO>(lParam);
			if(pDetails->fMask & RBBIM_SETID) {
				pDetails->fMask &= ~RBBIM_SETID;
				pDetails->wID = GetNewBandID();
				pDetails->fMask |= RBBIM_ID;
			}
			BOOL subclassedChild = FALSE;
			if(pDetails->fMask & RBBIM_CHILD) {
				if(flags.customizerOfMdiMessageHook && (wParam == -1 ? static_cast<int>(SendMessage(RB_GETBANDCOUNT, 0, 0)) : wParam) == 0 && !menuBandWindow.IsWindow()) {
					// by default use the first band as the menu band
					if(::IsWindow(pDetails->hwndChild)) {
						menuBandWindow.Attach(pDetails->hwndChild);
						subclassedChild = SetWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this), NULL);
						UpdateMDINonClientAreaSizes();
					}
				}
			}

			succeeded = DefWindowProc(message, wParam, lParam);
			if(succeeded) {
				RebuildAcceleratorTable();
				int insertedBandIndex = IDToBandIndex(pDetails->wID);
				CComPtr<IReBarBand> pRBarBandItf = ClassFactory::InitReBarBand(insertedBandIndex, this);
				if(pRBarBandItf) {
					Raise_InsertedBand(pRBarBandItf);
				}
			} else if(subclassedChild) {
				ATLVERIFY(RemoveWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this)));
				menuBandWindow.Detach();
			}
		}
	} else {
		// finally pass the message to the rebar
		LPREBARBANDINFO pDetails = reinterpret_cast<LPREBARBANDINFO>(lParam);
		if(pDetails->fMask & RBBIM_SETID) {
			pDetails->fMask &= ~RBBIM_SETID;
			pDetails->wID = GetNewBandID();
			pDetails->fMask |= RBBIM_ID;
		}
		BOOL subclassedChild = FALSE;
		if(pDetails->fMask & RBBIM_CHILD) {
			if(flags.customizerOfMdiMessageHook && (wParam == -1 ? static_cast<int>(SendMessage(RB_GETBANDCOUNT, 0, 0)) : wParam) == 0 && !menuBandWindow.IsWindow()) {
				// by default use the first band as the menu band
				if(::IsWindow(pDetails->hwndChild)) {
					menuBandWindow.Attach(pDetails->hwndChild);
					subclassedChild = SetWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this), NULL);
					UpdateMDINonClientAreaSizes();
				}
			}
		}

		succeeded = DefWindowProc(message, wParam, lParam);
		if(!succeeded && subclassedChild) {
			ATLVERIFY(RemoveWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this)));
			menuBandWindow.Detach();
		}
		RebuildAcceleratorTable();
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		if(succeeded) {
			// maybe we have a new band below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);

			UINT hitTestDetails = 0;
			int newBandUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
			if(newBandUnderMouse != bandUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(bandUnderMouse != -1) {
					CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandUnderMouse, this);
					Raise_BandMouseLeave(pRBBand, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				bandUnderMouse = newBandUnderMouse;

				if(bandUnderMouse != -1) {
					CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandUnderMouse, this);
					Raise_BandMouseEnter(pRBBand, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	return succeeded;
}

LRESULT ReBar::OnSetBandInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPREBARBANDINFO pBandData = reinterpret_cast<LPREBARBANDINFO>(lParam);
	#ifdef DEBUG
		if(message == RB_SETBANDINFOA) {
			ATLASSERT_POINTER(pBandData, REBARBANDINFOA);
		} else {
			ATLASSERT_POINTER(pBandData, REBARBANDINFOW);
		}
	#endif
	if(!pBandData) {
		return DefWindowProc(message, wParam, lParam);
	}

	// let ReBarWindow32 handle the message
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		if(pBandData->fMask & RBBIM_CHILD) {
			if(flags.customizerOfMdiMessageHook && wParam == 0 && !menuBandWindow.IsWindow()) {
				// by default use the first band as the menu band
				if(::IsWindow(pBandData->hwndChild)) {
					menuBandWindow.Attach(pBandData->hwndChild);
					ATLVERIFY(SetWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this), NULL));
					UpdateMDINonClientAreaSizes();
				}
			}
		}
		if(pBandData->fMask & RBBIM_STYLE) {
			if(pBandData->fStyle & RBBS_USECHEVRON) {
				/* If the band is visible, we need to convince the native rebar control to recalculate the layout
				   of this band. Otherwise the band's contained window might be positioned wrongly. This usually
				   happens in MDI frame windows. */
				REBARBANDINFO bandInfo = {0};
				bandInfo.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
				bandInfo.fMask = RBBIM_STYLE;
				SendMessage(RB_GETBANDINFO, wParam, reinterpret_cast<LPARAM>(&bandInfo));
				if(!(bandInfo.fStyle & RBBS_HIDDEN)) {
					SendMessage(RB_SHOWBAND, wParam, TRUE);
				}
			}
		}
		if(pBandData->fMask & RBBIM_TEXT) {
			RebuildAcceleratorTable();
		}
	}
	return lr;
}


LRESULT ReBar::OnMDIClientMDISetMenu(UINT message, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	switch(properties.replaceMDIFrameMenu) {
		case rmfmJustRemove:
		case rmfmFullReplace:
			// remove the menu
			HMENU hOldMenu = mdiFrame.GetMenu();
			DefSubclassProc(mdiClient, message, NULL, lParam);
			return reinterpret_cast<LRESULT>(hOldMenu);
	}
	wasHandled = FALSE;
	return 0;
}


LRESULT ReBar::OnMDIFrameActivate(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	mdiStatus.mdiFrameIsActive = (LOWORD(wParam) != WA_INACTIVE);
	if(menuBandWindow.IsWindow()) {
		// if the child window is transparent, redrawing the child window is not enough
		//menuBandWindow.RedrawWindow(NULL, NULL, RDW_INVALIDATE | RDW_FRAME | RDW_UPDATENOW);
		RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_INVALIDATE | RDW_UPDATENOW);
	}
	wasHandled = FALSE;
	return 0;
}


LRESULT ReBar::OnMenuBandChildCaptureChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(mdiStatus.activeMDIChildMaximized) {
		if(mdiStatus.pressedButton != -1) {
			ATLASSERT(mdiStatus.pressedButton == mdiStatus.lastPressedButton);
			mdiStatus.pressedButton = -1;
			RECT windowRectangle = {0};
			menuBandWindow.GetWindowRect(&windowRectangle);
			RECT buttonRectangles[3] = {0};
			CalcMDIChildButtonRectangles(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, buttonRectangles);

			CWindowDC dc(menuBandWindow);
			CMemoryDC memoryDC(dc, buttonRectangles[mdiStatus.lastPressedButton]);
			DrawMDIChildButton(memoryDC, buttonRectangles, mdiStatus.lastPressedButton);
		}
		mdiStatus.lastPressedButton = -1;
	} else {
		wasHandled = FALSE;
	}
	return 0;
}

LRESULT ReBar::OnMenuBandChildLButtonUp(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	if(!mdiStatus.activeMDIChildMaximized || GetCapture() != menuBandWindow || mdiStatus.lastPressedButton == -1) {
		wasHandled = FALSE;
		return 1;
	}

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	menuBandWindow.ClientToScreen(&pt);
	RECT windowRectangle = {0};
	menuBandWindow.GetWindowRect(&windowRectangle);
	pt.x -= windowRectangle.left;
	pt.y -= windowRectangle.top;

	int lastPressedButton = mdiStatus.lastPressedButton;
	ReleaseCapture();

	RECT buttonRectangles[3] = {0};
	CalcMDIChildButtonRectangles(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, buttonRectangles, menuBandWindow.GetExStyle() & WS_EX_LAYOUTRTL);
	if(PtInRect(&buttonRectangles[lastPressedButton], pt)) {
		switch(lastPressedButton) {
			case 0:     // close
				::SendMessage(mdiStatus.hWndChildMaximized, WM_SYSCOMMAND, SC_CLOSE, 0);
				break;
			case 1:     // restore
				::SendMessage(mdiStatus.hWndChildMaximized, WM_SYSCOMMAND, SC_RESTORE, 0);
				break;
			case 2:     // minimize
				::SendMessage(mdiStatus.hWndChildMaximized, WM_SYSCOMMAND, SC_MINIMIZE, 0);
				break;
			default:
				break;
		}
	}
	return 0;
}

LRESULT ReBar::OnMenuBandChildMouseMove(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	if(!mdiStatus.activeMDIChildMaximized || GetCapture() != menuBandWindow || mdiStatus.lastPressedButton == -1) {
		wasHandled = FALSE;
		return 1;
	}

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	menuBandWindow.ClientToScreen(&pt);
	RECT windowRectangle = {0};
	menuBandWindow.GetWindowRect(&windowRectangle);
	pt.x -= windowRectangle.left;
	pt.y -= windowRectangle.top;

	RECT buttonRectangles[3] = {0};
	CalcMDIChildButtonRectangles(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, buttonRectangles, menuBandWindow.GetExStyle() & WS_EX_LAYOUTRTL);
	int oldPressedButton = mdiStatus.pressedButton;
	mdiStatus.pressedButton = PtInRect(&buttonRectangles[mdiStatus.lastPressedButton], pt) ? mdiStatus.lastPressedButton : -1;
	if(oldPressedButton != mdiStatus.pressedButton) {
		CWindowDC dc(menuBandWindow);
		CalcMDIChildButtonRectangles(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, buttonRectangles);
		CMemoryDC memoryDC(dc, buttonRectangles[mdiStatus.pressedButton != -1 ? mdiStatus.pressedButton : oldPressedButton]);
		DrawMDIChildButton(memoryDC, buttonRectangles, mdiStatus.pressedButton != -1 ? mdiStatus.pressedButton : oldPressedButton);
	}
	return 0;
}

LRESULT ReBar::OnMenuBandChildNCCalcSize(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefSubclassProc(menuBandWindow, message, wParam, lParam);

	if(mdiStatus.activeMDIChildMaximized && wParam) {
		LPNCCALCSIZE_PARAMS pDetails = reinterpret_cast<LPNCCALCSIZE_PARAMS>(lParam);
		if(menuBandWindow.GetExStyle() & WS_EX_LAYOUTRTL) {
			pDetails->rgrc[0].left += mdiStatus.mdiChildWindowButtonAreaWidth;
			pDetails->rgrc[0].right -= mdiStatus.mdiChildWindowIconWidth;
		} else {
			pDetails->rgrc[0].left += mdiStatus.mdiChildWindowIconWidth;
			pDetails->rgrc[0].right -= mdiStatus.mdiChildWindowButtonAreaWidth;
		}
	}
	return lr;
}

LRESULT ReBar::OnMenuBandChildNCHitTest(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefSubclassProc(menuBandWindow, message, wParam, lParam);

	if(mdiStatus.activeMDIChildMaximized) {
		BOOL checkForButtons = FALSE;
		RECT windowRectangle = {0};
		menuBandWindow.GetWindowRect(&windowRectangle);
		POINT pt = {GET_X_LPARAM(lParam) - windowRectangle.left, GET_Y_LPARAM(lParam) - windowRectangle.top};
		if(menuBandWindow.GetExStyle() & WS_EX_LAYOUTRTL) {
			if(pt.x < mdiStatus.mdiChildWindowButtonAreaWidth) {
				lr = HTBORDER;
			} else if(pt.x > windowRectangle.right - windowRectangle.left - mdiStatus.mdiChildWindowIconWidth) {
				lr = HTBORDER;
				checkForButtons = TRUE;
			}
		} else {
			if(pt.x < mdiStatus.mdiChildWindowIconWidth) {
				lr = HTBORDER;
			} else if(pt.x > windowRectangle.right - windowRectangle.left - mdiStatus.mdiChildWindowButtonAreaWidth) {
				lr = HTBORDER;
				checkForButtons = TRUE;
			}
		}
		
		if(checkForButtons) {
			// TODO: Tooltip is displayed for the close button only.
			/*RECT buttonRectangles[3] = {0};
			CalcMDIChildButtonRectangles(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, buttonRectangles, menuBandWindow.GetExStyle() & WS_EX_LAYOUTRTL);
			
			if(PtInRect(&buttonRectangles[0], pt)) {
				lr = HTCLOSE;
			} else if(PtInRect(&buttonRectangles[1], pt)) {
				lr = HTMAXBUTTON;
			} else if(PtInRect(&buttonRectangles[2], pt)) {
				lr = HTMINBUTTON;
			}*/
		}
	}
	return lr;
}

LRESULT ReBar::OnMenuBandChildLButtonDblClk(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	if(!mdiStatus.activeMDIChildMaximized || mdiStatus.lastPressedButton != -1) {
		wasHandled = FALSE;
		return 1;
	}
	BOOL rightToLeft = (menuBandWindow.GetExStyle() & WS_EX_LAYOUTRTL);

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	RECT windowRectangle = {0};
	menuBandWindow.GetWindowRect(&windowRectangle);
	pt.x -= windowRectangle.left;
	pt.y -= windowRectangle.top;

	RECT iconRectangle = {0};
	CalcMDIChildIconRectangle(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, &iconRectangle, rightToLeft);

	if(PtInRect(&iconRectangle, pt)) {
		CMenuHandle menu = ::GetSystemMenu(mdiStatus.hWndChildMaximized, FALSE);
		UINT defaultCommandID = menu.GetMenuDefaultItem();
		if(defaultCommandID == static_cast<UINT>(-1)) {
			defaultCommandID = SC_CLOSE;
		}
		::SendMessage(mdiStatus.hWndChildMaximized, WM_SYSCOMMAND, defaultCommandID, 0);
	}
	return 0;
}

LRESULT ReBar::OnMenuBandChildNCLButtonDown(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	if(!mdiStatus.activeMDIChildMaximized) {
		wasHandled = FALSE;
		return 1;
	}
	BOOL rightToLeft = (menuBandWindow.GetExStyle() & WS_EX_LAYOUTRTL);

	POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	RECT windowRectangle = {0};
	menuBandWindow.GetWindowRect(&windowRectangle);
	pt.x -= windowRectangle.left;
	pt.y -= windowRectangle.top;

	RECT iconRectangle = {0};
	CalcMDIChildIconRectangle(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, &iconRectangle, rightToLeft);
	RECT buttonRectangles[3] = {0};
	CalcMDIChildButtonRectangles(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, buttonRectangles, rightToLeft);

	if(PtInRect(&iconRectangle, pt)) {
		CMenuHandle menu = ::GetSystemMenu(mdiStatus.hWndChildMaximized, FALSE);
		UINT menuID = 0;
		VARIANT_BOOL cancel = VARIANT_FALSE;
		Raise_BeforeDisplayMDIChildSystemMenu(HandleToLong(mdiStatus.hWndChildMaximized), HandleToLong(static_cast<HMENU>(menu)), rightToLeft ? windowRectangle.right : windowRectangle.left, windowRectangle.bottom, &cancel);
		if(cancel == VARIANT_FALSE) {
			menuID = static_cast<UINT>(menu.TrackPopupMenu(TPM_LEFTBUTTON | TPM_VERTICAL | TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RETURNCMD | TPM_VERPOSANIMATION, rightToLeft ? windowRectangle.right : windowRectangle.left, windowRectangle.bottom, *this));
		}

		// eat next message if click is on the same button
		OffsetRect(&iconRectangle, windowRectangle.left, windowRectangle.top);
		MSG msg = {0};
		if(PeekMessage(&msg, menuBandWindow, WM_NCLBUTTONDOWN, WM_NCLBUTTONDOWN, PM_NOREMOVE) && PtInRect(&iconRectangle, msg.pt)) {
			PeekMessage(&msg, menuBandWindow, WM_NCLBUTTONDOWN, WM_NCLBUTTONDOWN, PM_REMOVE);
		}

		if(menuID != 0) {
			::SendMessage(mdiStatus.hWndChildMaximized, WM_SYSCOMMAND, menuID, 0);
		}
		Raise_CleanupMDIChildSystemMenu(HandleToLong(mdiStatus.hWndChildMaximized), HandleToLong(static_cast<HMENU>(menu)));
	} else if(PtInRect(&buttonRectangles[0], pt)) {
		mdiStatus.pressedButton = 0;
		mdiStatus.lastPressedButton = mdiStatus.pressedButton;
	} else if(PtInRect(&buttonRectangles[1], pt)) {
		mdiStatus.pressedButton = 1;
		mdiStatus.lastPressedButton = mdiStatus.pressedButton;
	} else if(PtInRect(&buttonRectangles[2], pt)) {
		mdiStatus.pressedButton = 2;
		mdiStatus.lastPressedButton = mdiStatus.pressedButton;
	} else {
		wasHandled = FALSE;
	}

	// draw the button state if it was pressed
	if(mdiStatus.pressedButton != -1) {
		menuBandWindow.SetCapture();
		CWindowDC dc(menuBandWindow);
		CalcMDIChildButtonRectangles(windowRectangle.right - windowRectangle.left, windowRectangle.bottom - windowRectangle.top, buttonRectangles);
		CMemoryDC memoryDC(dc, buttonRectangles[mdiStatus.pressedButton]);
		DrawMDIChildButton(memoryDC, buttonRectangles, mdiStatus.pressedButton);
	}
	return 0;
}

LRESULT ReBar::OnMenuBandChildNCPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefSubclassProc(menuBandWindow, message, wParam, lParam);
	
	if(!mdiStatus.activeMDIChildMaximized) {
		return lr;
	}

	ATLASSERT(mdiStatus.hWndChildMaximized && mdiStatus.hIconChildMaximized);

	// get DC and window rectangle
	RECT rc;
	menuBandWindow.GetWindowRect(&rc);
	int availableWidth = rc.right - rc.left;
	int availableHeight = rc.bottom - rc.top;
	CWindowDC dc(menuBandWindow);

	// paint left side nonclient background and draw icon
	SetRect(&rc, 0, 0, mdiStatus.mdiChildWindowIconWidth, availableHeight);
	{
		CMemoryDC memoryDC(dc, rc);
		memoryDC.FillRect(&rc, COLOR_3DFACE);
		if(flags.usingThemes) {
			DrawThemeParentBackground(menuBandWindow, memoryDC, &rc);
		}

		RECT iconRectangle = {0};
		CalcMDIChildIconRectangle(availableWidth, availableHeight, &iconRectangle);
		memoryDC.DrawIconEx(iconRectangle.left, iconRectangle.top, mdiStatus.hIconChildMaximized, mdiStatus.smallIconSize.cx, mdiStatus.smallIconSize.cy);
	}
	// paint right side nonclient background
	SetRect(&rc, availableWidth - mdiStatus.mdiChildWindowButtonAreaWidth, 0, availableWidth, availableHeight);
	{
		CMemoryDC memoryDC(dc, rc);
		if(flags.usingThemes) {
			memoryDC.FillRect(&rc, COLOR_3DFACE);

			POINT pt = {0};
			memoryDC.GetViewportOrg(&pt);
			memoryDC.SetViewportOrg(pt.x + mdiStatus.mdiChildWindowIconWidth, pt.y);
			OffsetRect(&rc, -mdiStatus.mdiChildWindowIconWidth, 0);
			DrawThemeParentBackground(menuBandWindow, memoryDC, &rc);
			memoryDC.SetViewportOrg(pt);
			OffsetRect(&rc, mdiStatus.mdiChildWindowIconWidth, 0);
		} else {
			memoryDC.FillRect(&rc, COLOR_3DFACE);
		}

		// draw buttons
		RECT buttonRectangles[3] = {0};
		CalcMDIChildButtonRectangles(availableWidth, availableHeight, buttonRectangles);
		DrawMDIChildButton(memoryDC, buttonRectangles, -1);
	}

	return lr;
}

LRESULT ReBar::OnMenuBandChildSize(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefSubclassProc(menuBandWindow, message, wParam, lParam);
	AdjustMDINonClientButtonSize(GET_Y_LPARAM(lParam));
	return lr;
}


LRESULT ReBar::OnAllHookMessages(UINT message, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	if(message != WM_MDIGETACTIVE && message != WM_MDISETMENU && properties.replaceMDIFrameMenu == rmfmFullReplace) {
		BOOL activeChildIsMaximized = FALSE;
		HWND hWndChild = reinterpret_cast<HWND>(mdiClient.SendMessage(WM_MDIGETACTIVE, 0, reinterpret_cast<LPARAM>(&activeChildIsMaximized)));
		BOOL wasMaximized = mdiStatus.activeMDIChildMaximized;
		mdiStatus.activeMDIChildMaximized = (hWndChild && activeChildIsMaximized);
		HICON hIconOld = mdiStatus.hIconChildMaximized;

		if(mdiStatus.activeMDIChildMaximized) {
			if(mdiStatus.hWndChildMaximized != hWndChild) {
				mdiStatus.hWndChildMaximized = hWndChild;
				CWindow wnd(hWndChild);
				//mdiStatus.hIconChildMaximized = wnd.GetIcon(FALSE);
				mdiStatus.hIconChildMaximized = reinterpret_cast<HICON>(wnd.SendMessage(WM_GETICON, ICON_SMALL2, 0));
				if(!mdiStatus.hIconChildMaximized) {
					mdiStatus.hIconChildMaximized = wnd.GetIcon(TRUE);
					if(!mdiStatus.hIconChildMaximized) {
						// no icon set with WM_SETICON, get the class one
						// need conditional code because types don't match in winuser.h
						#ifdef _WIN64
							mdiStatus.hIconChildMaximized = reinterpret_cast<HICON>(::GetClassLongPtr(wnd, GCLP_HICONSM));
						#else
							mdiStatus.hIconChildMaximized = reinterpret_cast<HICON>(LongToHandle(::GetClassLongPtr(wnd, GCLP_HICONSM)));
						#endif
					}
				}
			}
		} else {
			mdiStatus.hWndChildMaximized = NULL;
			mdiStatus.hIconChildMaximized = NULL;
		}

		if(wasMaximized == !mdiStatus.activeMDIChildMaximized) {
			int bandIndex = FindBandByChildWindow(menuBandWindow);
			if(bandIndex >= 0) {
				REBARBANDINFO bandInfo = {0};
				bandInfo.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
				bandInfo.fMask = RBBIM_CHILDSIZE | RBBIM_IDEALSIZE | RBBIM_STYLE;
				SendMessage(RB_GETBANDINFO, bandIndex, reinterpret_cast<LPARAM>(&bandInfo));
				if(bandInfo.fStyle & RBBS_USECHEVRON) {
					int cxDiff = (mdiStatus.activeMDIChildMaximized ? 1 : -1) * (mdiStatus.mdiChildWindowIconWidth + mdiStatus.mdiChildWindowButtonAreaWidth);
					bandInfo.fMask = RBBIM_CHILDSIZE | RBBIM_IDEALSIZE;
					bandInfo.cxMinChild += cxDiff;
					bandInfo.cxIdeal += cxDiff;
					SendMessage(RB_SETBANDINFO, bandIndex, reinterpret_cast<LPARAM>(&bandInfo));
				}
			}
			// TK: WTL does not support switching between MDI child window system menu and the normal menu.
			//menuBandWindow.SendMessage(GetMDIChildWindowStateChangedMessage(), mdiStatus.activeMDIChildMaximized, 0);
		}

		if(wasMaximized == !mdiStatus.activeMDIChildMaximized || hIconOld != mdiStatus.hIconChildMaximized) {
			// force size change and redraw everything
			if(menuBandWindow.IsWindow()) {
				CWindowEx2 childWindow(menuBandWindow);
				RECT windowRectangle = {0};
				childWindow.GetWindowRect(&windowRectangle);
				::MapWindowPoints(HWND_DESKTOP, *this, reinterpret_cast<LPPOINT>(&windowRectangle), 2);
				childWindow.InternalSetRedraw(FALSE);
				childWindow.SetWindowPos(NULL, 0, 0, 1, 1, SWP_NOZORDER | SWP_NOMOVE);
				childWindow.SetWindowPos(NULL, &windowRectangle, SWP_NOZORDER | SWP_NOMOVE);
				childWindow.InternalSetRedraw(TRUE);
				childWindow.RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW);
			}
		}
	}
	return 0;
}


LRESULT ReBar::OnNCHitTestNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(properties.disabledEvents & deMouseEvents) {
		return 0;
	}

	LPNMMOUSE pDetails = reinterpret_cast<LPNMMOUSE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMMOUSE);
	if(!pDetails) {
		return 0;
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	UINT hitTestDetails = 0;
	#ifdef DEBUG
		ATLASSERT(static_cast<DWORD_PTR>(HitTest(pDetails->pt.x, pDetails->pt.y, &hitTestDetails)) == pDetails->dwItemSpec);
	#else
		HitTest(pDetails->pt.x, pDetails->pt.y, &hitTestDetails);
	#endif

	CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(pDetails->dwItemSpec, this);
	LONG ret = 0;
	Raise_NonClientHitTest(pRBBand, button, shift, pDetails->pt.x, pDetails->pt.y, static_cast<HitTestConstants>(hitTestDetails), &ret);
	return ret;
}

LRESULT ReBar::OnReleasedCaptureNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	Raise_ReleasedMouseCapture();
	return 0;
}

LRESULT ReBar::OnAutoBreakNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMREBARAUTOBREAK pDetails = reinterpret_cast<LPNMREBARAUTOBREAK>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMREBARAUTOBREAK);
	if(!pDetails) {
		return 0;
	}
	ATLASSERT(pDetails->uMsg == 1);

	CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(pDetails->uBand, this);
	VARIANT_BOOL doAutoBreak = BOOL2VARIANTBOOL(pDetails->fAutoBreak);
	Raise_AutoBreakingBand(pRBBand, &doAutoBreak);
	pDetails->fAutoBreak = VARIANTBOOL2BOOL(doAutoBreak);
	//pDetails->uMsg = 0;
	return 0;
}

LRESULT ReBar::OnAutoSizeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMRBAUTOSIZE pDetails = reinterpret_cast<LPNMRBAUTOSIZE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMRBAUTOSIZE);
	if(!pDetails) {
		return 0;
	}

	Raise_AutoSized(reinterpret_cast<RECTANGLE*>(&pDetails->rcTarget), reinterpret_cast<RECTANGLE*>(&pDetails->rcActual), BOOL2VARIANTBOOL(pDetails->fChanged));
	return 0;
}

LRESULT ReBar::OnBeginDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMREBAR pDetails = reinterpret_cast<LPNMREBAR>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMREBAR);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(pDetails->uBand, this);
	if(button & 1/*MouseButtonConstants.vbLeftButton*/) {
		Raise_BandBeginDrag(pRBBand, button, shift, mousePosition.x, mousePosition.y, hitTestDetails, &cancel);
	} else {
		Raise_BandBeginRDrag(pRBBand, button, shift, mousePosition.x, mousePosition.y, hitTestDetails, &cancel);
	}
	return VARIANTBOOL2BOOL(cancel);
}

LRESULT ReBar::OnChevronPushedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMREBARCHEVRON pDetails = reinterpret_cast<LPNMREBARCHEVRON>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMREBARCHEVRON);
	if(!pDetails) {
		return 0;
	}
	bandChevronPushed = pDetails->uBand;

	VARIANT_BOOL doDefault = VARIANT_TRUE;
	CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(pDetails->uBand, this);
	Raise_ChevronClick(pRBBand, reinterpret_cast<RECTANGLE*>(&pDetails->rc), pDetails->lParamNM, &doDefault);
	if(doDefault != VARIANT_FALSE) {
		REBARBANDINFO band = {0};
		band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
		band.fMask = RBBIM_CHILD;
		if(SendMessage(RB_GETBANDINFO, pDetails->uBand, reinterpret_cast<LPARAM>(&band))) {
			POINT pt = {pDetails->rc.right, pDetails->rc.top};
			ClientToScreen(&pt);
			::SendMessage(band.hwndChild, GetDisplayChevronPopupMessage(), 0, reinterpret_cast<LPARAM>(&pt));
		}
	}
	bandChevronPushed = -1;
	return 0;
}

LRESULT ReBar::OnChildSizeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMREBARCHILDSIZE pDetails = reinterpret_cast<LPNMREBARCHILDSIZE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMREBARCHILDSIZE);
	if(!pDetails) {
		return 0;
	}

	CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(pDetails->uBand, this);
	Raise_ResizingContainedWindow(pRBBand, reinterpret_cast<RECTANGLE*>(&pDetails->rcBand), reinterpret_cast<RECTANGLE*>(&pDetails->rcChild));
	return 0;
}

LRESULT ReBar::OnDeletingBandNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMREBAR pDetails = reinterpret_cast<LPNMREBAR>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMREBAR);
	if(!pDetails) {
		return 0;
	}

	if(!(properties.disabledEvents & deFreeBandData)) {
		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(pDetails->uBand, this);
		Raise_FreeBandData(pRBBand);
	}
	return 0;
}

LRESULT ReBar::OnEndDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMREBAR pDetails = reinterpret_cast<LPNMREBAR>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMREBAR);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	ScreenToClient(&mousePosition);
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);

	CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(pDetails->uBand, this);
	Raise_BandEndDrag(pRBBand, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
	return 0;
}

LRESULT ReBar::OnGetObjectNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMOBJECTNOTIFY pDetails = reinterpret_cast<LPNMOBJECTNOTIFY>(pNotificationDetails);
	ATLASSERT(*(pDetails->piid) == IID_IDropTarget);
	pDetails->hResult = QueryInterface(*(pDetails->piid), &pDetails->pObject);
	return 0;
}

LRESULT ReBar::OnHeightChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	Raise_HeightChanged();
	return 0;
}

LRESULT ReBar::OnLayoutChangedNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	RebuildAcceleratorTable();
	Raise_LayoutChanged();
	return 0;
}

LRESULT ReBar::OnMinMaxNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	if(!pNotificationDetails) {
		return 0;
	}

	PINT pBandIndex = NULL;
	// TODO: This doesn't work on Windows 7, but should work on Windows 2000.
	/*LPBYTE p = reinterpret_cast<LPBYTE>(pNotificationDetails) + sizeof(NMHDR);
	HFONT hFont = GetFont();
	do {
		if(RtlEqualMemory(p, &hFont, sizeof(hFont))) {
			pBandIndex = reinterpret_cast<PINT>(p + sizeof(HFONT) + 2 * sizeof(UINT));
			break;
		}
	} while(p++ < reinterpret_cast<LPBYTE>(pNotificationDetails) + 100);*/

	CComPtr<IReBarBand> pRBBand;
	if(pBandIndex) {
		pRBBand = ClassFactory::InitReBarBand(*pBandIndex, this);
	}
	VARIANT_BOOL cancel = VARIANT_FALSE;
	Raise_TogglingBand(pRBBand, &cancel);
	return VARIANTBOOL2BOOL(cancel);
}

LRESULT ReBar::OnSplitterDragNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	Raise_DraggingSplitter();
	return 0;
}


LRESULT ReBar::OnCustomDrawNotification(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPNMCUSTOMDRAW pDetails = reinterpret_cast<LPNMCUSTOMDRAW>(lParam);
	CustomDrawReturnValuesConstants returnValue = static_cast<CustomDrawReturnValuesConstants>(this->DefWindowProc(message, wParam, lParam));
	ATLASSERT(pDetails->uItemState == 0);
	ATLASSERT((pDetails->dwDrawStage & CDDS_SUBITEM) == 0);

	if(!(properties.disabledEvents & deCustomDraw)) {
		CComPtr<IReBarBand> pRBBand = NULL;
		if(pDetails->dwDrawStage & (CDDS_ITEM | CDDS_SUBITEM)) {
			int bandIndex = IDToBandIndex(pDetails->dwItemSpec);
			if(bandIndex >= 0) {
				if(pDetails->dwDrawStage & CDDS_ITEM) {
					pRBBand = ClassFactory::InitReBarBand(bandIndex, this, FALSE);
				}
			} else {
				return static_cast<LRESULT>(returnValue);
			}
		}
		Raise_CustomDraw(pRBBand, static_cast<CustomDrawStageConstants>(pDetails->dwDrawStage), static_cast<CustomDrawItemStateConstants>(pDetails->uItemState), HandleToLong(pDetails->hdc), reinterpret_cast<RECTANGLE*>(&pDetails->rc), &returnValue);
	}
	return static_cast<LRESULT>(returnValue);
}


void ReBar::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// use the system font
			/*NONCLIENTMETRICS nonClientMetrics = {0};
			nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
			SystemParametersInfo(SPI_GETNONCLIENTMETRICS, nonClientMetrics.cbSize, &nonClientMetrics, 0);
			properties.font.currentFont.CreateFontIndirect(&nonClientMetrics.lfCaptionFont);
			properties.font.owningFontResource = TRUE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));*/
			SendMessage(WM_SETFONT, 0, MAKELPARAM(TRUE, 0));
		} else {
			/* The whole font object or at least some of its attributes were changed. 'font.pFontDisp' is
			   still valid, so simply update our font. */
			if(properties.font.pFontDisp) {
				CComQIPtr<IFont, &IID_IFont> pFont(properties.font.pFontDisp);
				if(pFont) {
					HFONT hFont = NULL;
					pFont->get_hFont(&hFont);
					properties.font.currentFont.Attach(hFont);
					properties.font.owningFontResource = FALSE;

					SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
				} else {
					SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
				}
			} else {
				SendMessage(WM_SETFONT, NULL, MAKELPARAM(TRUE, 0));
			}
			Invalidate();
		}
	}
	properties.font.dontGetFontObject = FALSE;
	FireViewChange();
}


inline HRESULT ReBar::Raise_AutoBreakingBand(IReBarBand* pBand, VARIANT_BOOL* pDoAutoBreak)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_AutoBreakingBand(pBand, pDoAutoBreak);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_AutoSized(RECTANGLE* pTargetRectangle, RECTANGLE* pActualRectangle, VARIANT_BOOL changedBandHeightOrStyle)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_AutoSized(pTargetRectangle, pActualRectangle, changedBandHeightOrStyle);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_BandBeginDrag(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails, VARIANT_BOOL* pCancelDrag)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BandBeginDrag(pBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails), pCancelDrag);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_BandBeginRDrag(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails, VARIANT_BOOL* pCancelDrag)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BandBeginRDrag(pBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails), pCancelDrag);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_BandEndDrag(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BandEndDrag(pBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_BandMouseEnter(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_BandMouseEnter(pBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT ReBar::Raise_BandMouseLeave(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_BandMouseLeave(pBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT ReBar::Raise_BeforeDisplayMDIChildSystemMenu(LONG hActiveMDIChild, LONG hMenu, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pCancelMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeforeDisplayMDIChildSystemMenu(hActiveMDIChild, hMenu, x, y, pCancelMenu);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_ChevronClick(IReBarBand* pBand, RECTANGLE* pChevronRectangle, LONG userData, VARIANT_BOOL* pDoDefault)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ChevronClick(pBand, pChevronRectangle, userData, pDoDefault);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_CleanupMDIChildSystemMenu(LONG hActiveMDIChild, LONG hMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CleanupMDIChildSystemMenu(hActiveMDIChild, hMenu);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedBand = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(mouseStatus.lastClickedBand, this);
		return Fire_Click(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			/*if(properties.processContextMenuKeys) {
				// well?!
			} else {*/
				return S_OK;
			//}
		}

		UINT hitTestDetails = 0;
		int bandIndex = HitTest(x, y, &hitTestDetails);

		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
		return Fire_ContextMenu(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_CustomDraw(IReBarBand* pBand, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants bandState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CustomDraw(pBand, drawStage, bandState, hDC, pDrawingRectangle, pFurtherProcessing);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int bandIndex = HitTest(x, y, &hitTestDetails);
	if(bandIndex != mouseStatus.lastClickedBand) {
		bandIndex = -1;
	}
	mouseStatus.lastClickedBand = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
		return Fire_DblClick(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_DestroyedControlWindow(LONG hWnd)
{
	KillTimer(timers.ID_REDRAW);
	if(properties.registerForOLEDragDrop == rfoddAdvancedDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	if(mdiClient.IsWindow()) {
		ATLVERIFY(RemoveWindowSubclass(mdiClient, ReBar::MDIClientSubclass, reinterpret_cast<UINT_PTR>(this)));
		mdiClient.Detach();
	}
	if(mdiFrame.IsWindow()) {
		ATLVERIFY(RemoveWindowSubclass(mdiFrame, ReBar::MDIFrameSubclass, reinterpret_cast<UINT_PTR>(this)));
		mdiFrame.Detach();
	}
	if(menuBandWindow.IsWindow()) {
		ATLVERIFY(RemoveWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this)));
		menuBandWindow.Detach();
	}
	if(flags.customizerOfMdiMessageHook) {
		RemoveMessageHookClient();
		flags.customizerOfMdiMessageHook = FALSE;
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_DraggingSplitter(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DraggingSplitter();
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_FreeBandData(IReBarBand* pBand)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeBandData(pBand);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_HeightChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HeightChanged();
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_InsertedBand(IReBarBand* pBand)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedBand(pBand);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_InsertingBand(IVirtualReBarBand* pBand, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingBand(pBand, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_LayoutChanged(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_LayoutChanged();
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedBand = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(mouseStatus.lastClickedBand, this);
		return Fire_MClick(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int bandIndex = HitTest(x, y, &hitTestDetails);
	if(bandIndex != mouseStatus.lastClickedBand) {
		bandIndex = -1;
	}
	mouseStatus.lastClickedBand = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
		return Fire_MDblClick(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	if(!mouseStatus.hoveredControl) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = *this;
		trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		Raise_MouseHover(button, shift, x, y);
	}
	mouseStatus.StoreClickCandidate(button);

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		int bandIndex = HitTest(x, y, &hitTestDetails);

		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
		return Fire_MouseDown(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = *this;
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	UINT hitTestDetails = 0;
	int bandIndex = HitTest(x, y, &hitTestDetails);
	bandUnderMouse = bandIndex;

	CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		Fire_MouseEnter(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	if(pRBBand) {
		Raise_BandMouseEnter(pRBBand, button, shift, x, y, hitTestDetails);
	}
	return hr;
}

inline HRESULT ReBar::Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.hoveredControl) {
		mouseStatus.HoverControl();

		//if(m_nFreezeEvents == 0) {
			UINT hitTestDetails = 0;
			int bandIndex = HitTest(x, y, &hitTestDetails);

			CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
			return Fire_MouseHover(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		//}
	}
	return S_OK;
}

inline HRESULT ReBar::Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int bandIndex = HitTest(x, y, &hitTestDetails);

	CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandUnderMouse, this);
	if(pRBBand) {
		Raise_BandMouseLeave(pRBBand, button, shift, x, y, hitTestDetails);
	}
	bandUnderMouse = -1;

	mouseStatus.LeaveControl();

	//if(m_nFreezeEvents == 0) {
		pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
		return Fire_MouseLeave(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	UINT hitTestDetails = 0;
	int bandIndex = HitTest(x, y, &hitTestDetails);

	CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
	// Do we move over another band than before?
	if(bandIndex != bandUnderMouse) {
		CComPtr<IReBarBand> pPrevRBBand = ClassFactory::InitReBarBand(bandUnderMouse, this);
		if(pPrevRBBand) {
			Raise_BandMouseLeave(pPrevRBBand, button, shift, x, y, hitTestDetails);
		}
		bandUnderMouse = bandIndex;
		if(pRBBand) {
			Raise_BandMouseEnter(pRBBand, button, shift, x, y, hitTestDetails);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea;
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != *this) {
				hasLeftControl = TRUE;
			}
		} else {
			hasLeftControl = TRUE;
		}

		if(!hasLeftControl) {
			// we don't have left the control, so raise the click event
			switch(button) {
				case 1/*MouseButtonConstants.vbLeftButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_Click(button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		//if(m_nFreezeEvents == 0) {
			UINT hitTestDetails = 0;
			int bandIndex = HitTest(x, y, &hitTestDetails);
			CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
			return Fire_MouseUp(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		//}
		//return S_OK;
	} else {
		return S_OK;
	}
}

inline HRESULT ReBar::Raise_NonClientHitTest(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, LONG* pReturnValue)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_NonClientHitTest(pBand, button, shift, x, y, hitTestDetails, pReturnValue);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition, BOOL* /*pCallDropTargetHelper*/)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails/*, TRUE*/);
	CComPtr<IReBarBand> pDropTarget = ClassFactory::InitReBarBand(dropTarget, this);

	if(pData) {
		/* Actually we wouldn't need the next line, because the data object passed to this method should
		   always be the same as the data object that was passed to Raise_OLEDragEnter. */
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragDrop(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT ReBar::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition, BOOL* /*pCallDropTargetHelper*/)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	IReBarBand* pDropTarget = NULL;
	ClassFactory::InitReBarBand(dropTarget, this, IID_IReBarBand, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}
	return hr;
}

inline HRESULT ReBar::Raise_OLEDragLeave(BOOL* /*pCallDropTargetHelper*/)
{
	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails);
	CComPtr<IReBarBand> pDropTarget = ClassFactory::InitReBarBand(dropTarget, this);

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragLeave(dragDropStatus.pActiveDataObject, pDropTarget, button, shift, dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	dragDropStatus.pActiveDataObject = NULL;
	dragDropStatus.OLEDragLeaveOrDrop();
	Invalidate();

	return hr;
}

inline HRESULT ReBar::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition, BOOL* /*pCallDropTargetHelper*/)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails);
	IReBarBand* pDropTarget = NULL;
	ClassFactory::InitReBarBand(dropTarget, this, IID_IReBarBand, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}
	return hr;
}

inline HRESULT ReBar::Raise_RawMenuMessage(LONG message, LONG wParam, LONG lParam, LONG* pResult, VARIANT_BOOL* pHandledEvent)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RawMenuMessage(message, wParam, lParam, pResult, pHandledEvent);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedBand = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(mouseStatus.lastClickedBand, this);
		return Fire_RClick(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int bandIndex = HitTest(x, y, &hitTestDetails);
	if(bandIndex != mouseStatus.lastClickedBand) {
		bandIndex = -1;
	}
	mouseStatus.lastClickedBand = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
		return Fire_RDblClick(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_RecreatedControlWindow(LONG hWnd)
{
	// configure the control
	SendConfigurationMessages();

	if(properties.registerForOLEDragDrop == rfoddAdvancedDragDrop) {
		ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
	}

	if(properties.dontRedraw) {
		SetTimer(timers.ID_REDRAW, timers.INT_REDRAW);
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_RecreatedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_ReleasedMouseCapture(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ReleasedMouseCapture();
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_RemovedBand(IVirtualReBarBand* pBand)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedBand(pBand);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_RemovingBand(IReBarBand* pBand, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingBand(pBand, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_ResizingContainedWindow(IReBarBand* pBand, RECTANGLE* pBandClientRectangle, RECTANGLE* pContainedWindowRectangle)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizingContainedWindow(pBand, pBandClientRectangle, pContainedWindowRectangle);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_SelectedMenuItem(LONG commandIDOrSubMenuIndex, MenuItemStateConstants menuItemState, LONG hMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectedMenuItem(commandIDOrSubMenuIndex, menuItemState, hMenu);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_TogglingBand(IReBarBand* pBand, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_TogglingBand(pBand, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedBand = HitTest(x, y, &hitTestDetails);

	//if(m_nFreezeEvents == 0) {
		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(mouseStatus.lastClickedBand, this);
		return Fire_XClick(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ReBar::Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int bandIndex = HitTest(x, y, &hitTestDetails);
	if(bandIndex != mouseStatus.lastClickedBand) {
		bandIndex = -1;
	}
	mouseStatus.lastClickedBand = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IReBarBand> pRBBand = ClassFactory::InitReBarBand(bandIndex, this);
		return Fire_XDblClick(pRBBand, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}


void ReBar::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

DWORD ReBar::GetExStyleBits(void)
{
	DWORD extendedStyle = WS_EX_LEFT | WS_EX_LTRREADING;
	switch(properties.appearance) {
		case a3D:
			extendedStyle |= WS_EX_CLIENTEDGE;
			break;
		case a3DLight:
			extendedStyle |= WS_EX_STATICEDGE;
			break;
	}
	if(properties.rightToLeft & rtlLayout) {
		extendedStyle |= WS_EX_LAYOUTRTL;
	}
	if(properties.rightToLeft & rtlText) {
		extendedStyle |= WS_EX_RTLREADING;
	}
	return extendedStyle;
}

DWORD ReBar::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	// Set CCS_NODIVIDER, otherwise we will get problems in MDI windows, if the MDI frame is resized.
	style |= CCS_NODIVIDER;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}
	if(!(properties.allowBandReordering)) {
		style |= RBS_FIXEDORDER;
	}
	if(properties.autoUpdateLayout) {
		style |= RBS_AUTOSIZE;
	}
	if(properties.displayBandSeparators) {
		style |= RBS_BANDBORDERS;
	}
	if(!properties.fixedBandHeight) {
		style |= RBS_VARHEIGHT;
	}
	if(properties.registerForOLEDragDrop == rfoddNativeDragDrop) {
		style |= RBS_REGISTERDROP;
	}
	if(properties.toggleOnDoubleClick) {
		style |= RBS_DBLCLKTOGGLE;
	}
	if(properties.verticalSizingGripsOnVerticalOrientation) {
		style |= RBS_VERTICALGRIPPER;
	}

	if(m_spClientSite) {
		IOleControlSite* pControlSite = NULL;
		m_spClientSite->QueryInterface(IID_IOleControlSite, reinterpret_cast<LPVOID*>(&pControlSite));
		if(pControlSite) {
			IDispatch* pExtendedControl = NULL;
			pControlSite->GetExtendedControl(&pExtendedControl);
			if(pExtendedControl) {
				LPWSTR pPropertyName = OLESTR("Align");
				DISPID dispidAlign = 0;
				if(pExtendedControl->GetIDsOfNames(IID_NULL, &pPropertyName, 1, LOCALE_SYSTEM_DEFAULT, &dispidAlign) == S_OK) {
					VARIANT v;
					VariantInit(&v);
					DISPPARAMS params = {NULL, NULL, 0, 0};
					if(SUCCEEDED(pExtendedControl->Invoke(dispidAlign, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &v, NULL, NULL))) {
						if(SUCCEEDED(VariantChangeType(&v, &v, 0, VT_I4))) {
							switch(v.lVal) {
								case 1/*vbAlignTop*/:
									style |= CCS_NOPARENTALIGN | CCS_NOMOVEY/*CCS_TOP*/;
									properties.orientation = oHorizontal;
									FireOnChanged(DISPID_RB_ORIENTATION);
									break;
								case 2/*vbAlignBottom*/:
									style |= CCS_NOPARENTALIGN | CCS_NOMOVEY/*CCS_BOTTOM*/;
									properties.orientation = oHorizontal;
									FireOnChanged(DISPID_RB_ORIENTATION);
									break;
								case 3/*vbAlignLeft*/:
									style |= CCS_NOPARENTALIGN | CCS_NOMOVEY | CCS_VERT/*CCS_LEFT*/;
									properties.orientation = oVertical;
									FireOnChanged(DISPID_RB_ORIENTATION);
									break;
								case 4/*vbAlignRight*/:
									style |= CCS_NOPARENTALIGN | CCS_NOMOVEY | CCS_VERT/*CCS_RIGHT*/;
									properties.orientation = oVertical;
									FireOnChanged(DISPID_RB_ORIENTATION);
									break;
								default:
									style |= CCS_NOPARENTALIGN | CCS_NOMOVEY;
									switch(properties.orientation) {
										case oVertical:
											style |= CCS_VERT;
											break;
									}
									break;
							}
						}
					}
					VariantClear(&v);
				}
				pExtendedControl->Release();
			}
			pControlSite->Release();
		}
	}
	return style;
}

void ReBar::SendConfigurationMessages(void)
{
	if(!IsInDesignMode()) {
		switch(properties.replaceMDIFrameMenu) {
			case rmfmJustRemove:
			case rmfmFullReplace:
			{
				HWND hWndMDIClient = FindWindowEx(GetParent(), NULL, _T("MDIClient"), NULL);
				if(mdiClient != hWndMDIClient) {
					if(mdiClient.IsWindow()) {
						ATLVERIFY(RemoveWindowSubclass(mdiClient, ReBar::MDIClientSubclass, reinterpret_cast<UINT_PTR>(this)));
						mdiClient.Detach();
					}
					if(mdiFrame.IsWindow()) {
						ATLVERIFY(RemoveWindowSubclass(mdiFrame, ReBar::MDIFrameSubclass, reinterpret_cast<UINT_PTR>(this)));
						mdiFrame.Detach();
					}
					if(menuBandWindow.IsWindow()) {
						ATLVERIFY(RemoveWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this)));
						menuBandWindow.Detach();
					}

					if(hWndMDIClient) {
						mdiClient.Attach(hWndMDIClient);
						ATLVERIFY(SetWindowSubclass(mdiClient, ReBar::MDIClientSubclass, reinterpret_cast<UINT_PTR>(this), NULL));

						if(properties.replaceMDIFrameMenu == rmfmFullReplace) {
							mdiFrame.Attach(mdiClient.GetParent());
							ATLVERIFY(SetWindowSubclass(mdiFrame, ReBar::MDIFrameSubclass, reinterpret_cast<UINT_PTR>(this), NULL));
							if(!flags.customizerOfMdiMessageHook) {
								flags.customizerOfMdiMessageHook = TRUE;
								AddMessageHookClient();
							}

							// by default use the first band as the menu band
							REBARBANDINFO band = {0};
							band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
							band.fMask = RBBIM_CHILD;
							if(SendMessage(RB_GETBANDINFO, 0, reinterpret_cast<LPARAM>(&band)) && ::IsWindow(band.hwndChild)) {
								menuBandWindow.Attach(band.hwndChild);
								ATLVERIFY(SetWindowSubclass(menuBandWindow, ReBar::MenuBandWindowSubclass, reinterpret_cast<UINT_PTR>(this), NULL));
								UpdateMDINonClientAreaSizes();
							}
						}
					}
				}
				break;
			}
		}
	}

	if(properties.backColor != static_cast<OLE_COLOR>(-1)) {
		SendMessage(RB_SETBKCOLOR, 0, OLECOLOR2COLORREF(properties.backColor));
	}
	if(properties.foreColor != static_cast<OLE_COLOR>(-1)) {
		SendMessage(RB_SETTEXTCOLOR, 0, OLECOLOR2COLORREF(properties.foreColor));
	}
	if(properties.highlightColor != static_cast<OLE_COLOR>(-1) || properties.shadowColor != static_cast<OLE_COLOR>(-1)) {
		COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
		colorScheme.clrBtnHighlight = (properties.highlightColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.highlightColor));
		colorScheme.clrBtnShadow = (properties.shadowColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.shadowColor));
		SendMessage(RB_SETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme));
	}
	if(IsComctl32Version610OrNewer()) {
		DWORD extendedStyle = 0;
		if(properties.displaySplitter) {
			extendedStyle |= RBS_EX_SPLITTER;
		}
		SendMessage(RB_SETEXTENDEDSTYLE, 0, extendedStyle);
	}
	ApplyFont();

	if(IsInDesignMode()) {
		// insert a dummy band
		REBARBANDINFO insertionData = {0};
		insertionData.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
		insertionData.fMask = RBBIM_STYLE | RBBIM_TEXT | RBBIM_CHILDSIZE;
		insertionData.fStyle = RBBS_GRIPPERALWAYS;
		insertionData.cyMinChild = 22;
		insertionData.lpText = TEXT("Dummy Band 1");
		SendMessage(RB_INSERTBAND, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(&insertionData));
		insertionData.fStyle |= RBBS_BREAK;
		insertionData.lpText = TEXT("Dummy Band 2");
		SendMessage(RB_INSERTBAND, static_cast<WPARAM>(-1), reinterpret_cast<LPARAM>(&insertionData));
	}
}

HCURSOR ReBar::MousePointerConst2hCursor(MousePointerConstants mousePointer)
{
	WORD flag = 0;
	switch(mousePointer) {
		case mpArrow:
			flag = OCR_NORMAL;
			break;
		case mpCross:
			flag = OCR_CROSS;
			break;
		case mpIBeam:
			flag = OCR_IBEAM;
			break;
		case mpIcon:
			flag = OCR_ICOCUR;
			break;
		case mpSize:
			flag = OCR_SIZEALL;     // OCR_SIZE is obsolete
			break;
		case mpSizeNESW:
			flag = OCR_SIZENESW;
			break;
		case mpSizeNS:
			flag = OCR_SIZENS;
			break;
		case mpSizeNWSE:
			flag = OCR_SIZENWSE;
			break;
		case mpSizeEW:
			flag = OCR_SIZEWE;
			break;
		case mpUpArrow:
			flag = OCR_UP;
			break;
		case mpHourglass:
			flag = OCR_WAIT;
			break;
		case mpNoDrop:
			flag = OCR_NO;
			break;
		case mpArrowHourglass:
			flag = OCR_APPSTARTING;
			break;
		case mpArrowQuestion:
			flag = 32651;
			break;
		case mpSizeAll:
			flag = OCR_SIZEALL;
			break;
		case mpHand:
			flag = OCR_HAND;
			break;
		case mpInsertMedia:
			flag = 32663;
			break;
		case mpScrollAll:
			flag = 32654;
			break;
		case mpScrollN:
			flag = 32655;
			break;
		case mpScrollNE:
			flag = 32660;
			break;
		case mpScrollE:
			flag = 32658;
			break;
		case mpScrollSE:
			flag = 32662;
			break;
		case mpScrollS:
			flag = 32656;
			break;
		case mpScrollSW:
			flag = 32661;
			break;
		case mpScrollW:
			flag = 32657;
			break;
		case mpScrollNW:
			flag = 32659;
			break;
		case mpScrollNS:
			flag = 32652;
			break;
		case mpScrollEW:
			flag = 32653;
			break;
		default:
			return NULL;
	}

	return static_cast<HCURSOR>(LoadImage(0, MAKEINTRESOURCE(flag), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED));
}

int ReBar::IDToBandIndex(LONG ID)
{
	return static_cast<int>(SendMessage(RB_IDTOINDEX, ID, 0));
}

LONG ReBar::BandIndexToID(int bandIndex)
{
	ATLASSERT(IsWindow());

	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_ID;
	if(SendMessage(RB_GETBANDINFO, bandIndex, reinterpret_cast<LPARAM>(&band))) {
		return band.wID;
	}
	return -1;
}

int ReBar::FindBandByChildWindow(HWND hWndChild)
{
	REBARBANDINFO bandInfo = {0};
	bandInfo.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	bandInfo.fMask = RBBIM_CHILD;
	int numberOfBands = static_cast<int>(SendMessage(RB_GETBANDCOUNT, 0, 0));
	for(int bandIndex = 0; bandIndex < numberOfBands; bandIndex++) {
		SendMessage(RB_GETBANDINFO, bandIndex, reinterpret_cast<LPARAM>(&bandInfo));
		if(bandInfo.hwndChild == hWndChild) {
			return bandIndex;
		}
	}
	return -1;
}


LRESULT CALLBACK ReBar::MDIClientSubclass(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/)
{
	LRESULT lr = 0;
	ReBar* pThis = reinterpret_cast<ReBar*>(idSubclass);
	if(pThis) {
		if(!pThis->ProcessWindowMessage(hWnd, message, wParam, lParam, lr, 1)) {
			lr = DefSubclassProc(hWnd, message, wParam, lParam);
		}
	} else {
		lr = DefSubclassProc(hWnd, message, wParam, lParam);
	}
	return lr;
}

LRESULT CALLBACK ReBar::MDIFrameSubclass(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/)
{
	LRESULT lr = 0;
	ReBar* pThis = reinterpret_cast<ReBar*>(idSubclass);
	if(pThis) {
		if(!pThis->ProcessWindowMessage(hWnd, message, wParam, lParam, lr, 2)) {
			lr = DefSubclassProc(hWnd, message, wParam, lParam);
		}
	} else {
		lr = DefSubclassProc(hWnd, message, wParam, lParam);
	}
	return lr;
}

LRESULT CALLBACK ReBar::MenuBandWindowSubclass(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/)
{
	LRESULT lr = 0;
	ReBar* pThis = reinterpret_cast<ReBar*>(idSubclass);
	if(pThis) {
		if(!pThis->ProcessWindowMessage(hWnd, message, wParam, lParam, lr, 3)) {
			lr = DefSubclassProc(hWnd, message, wParam, lParam);
		}
	} else {
		lr = DefSubclassProc(hWnd, message, wParam, lParam);
	}
	return lr;
}

void ReBar::UpdateMDINonClientAreaSizes(void)
{
	mdiStatus.smallIconSize.cx = GetSystemMetrics(SM_CXSMICON);
	mdiStatus.smallIconSize.cy = GetSystemMetrics(SM_CYSMICON);
	mdiStatus.mdiChildWindowIconWidth = mdiStatus.smallIconSize.cx + 1;

	SIZE fallbackSize = {GetSystemMetrics(SM_CXMENUSIZE), GetSystemMetrics(SM_CYMENUSIZE)};
	CTheme themingEngine;
	if(flags.usingThemes) {
		themingEngine.OpenThemeData(*this, VSCLASS_WINDOW);
	}
	if(!themingEngine.IsThemeNull()) {
		CWindowDC dc(*this);
		int buttonWidth = 0;
		int buttonHeight = 0;
		if(SUCCEEDED(themingEngine.GetThemeMetric(dc, WP_MDICLOSEBUTTON, CBS_NORMAL, TMT_WIDTH, &buttonWidth)) || SUCCEEDED(themingEngine.GetThemeMetric(dc, WP_MDICLOSEBUTTON, CBS_NORMAL, TMT_HEIGHT, &buttonHeight))) {
			mdiStatus.buttonSize[0].cx = buttonWidth;
			mdiStatus.buttonSize[0].cy = buttonHeight;
		} else {
			mdiStatus.buttonSize[0] = fallbackSize;
		}
		if(SUCCEEDED(themingEngine.GetThemeMetric(dc, WP_MDIRESTOREBUTTON, CBS_NORMAL, TMT_WIDTH, &buttonWidth)) || SUCCEEDED(themingEngine.GetThemeMetric(dc, WP_MDIRESTOREBUTTON, CBS_NORMAL, TMT_HEIGHT, &buttonHeight))) {
			mdiStatus.buttonSize[1].cx = buttonWidth;
			mdiStatus.buttonSize[1].cy = buttonHeight;
		} else {
			mdiStatus.buttonSize[1] = fallbackSize;
		}
		if(SUCCEEDED(themingEngine.GetThemeMetric(dc, WP_MDIMINBUTTON, CBS_NORMAL, TMT_WIDTH, &buttonWidth)) || SUCCEEDED(themingEngine.GetThemeMetric(dc, WP_MDIMINBUTTON, CBS_NORMAL, TMT_HEIGHT, &buttonHeight))) {
			mdiStatus.buttonSize[2].cx = buttonWidth;
			mdiStatus.buttonSize[2].cy = buttonHeight;
		} else {
			mdiStatus.buttonSize[2] = fallbackSize;
		}
		mdiStatus.mdiChildWindowButtonAreaWidth = mdiStatus.buttonSize[0].cx + mdiStatus.buttonSize[1].cx + mdiStatus.buttonSize[2].cx + 1;
	} else {
		mdiStatus.buttonSize[0] = fallbackSize;
		mdiStatus.buttonSize[1] = fallbackSize;
		mdiStatus.buttonSize[2] = fallbackSize;
		mdiStatus.mdiChildWindowButtonAreaWidth = mdiStatus.buttonSize[0].cx + mdiStatus.buttonSize[1].cx + 2 + mdiStatus.buttonSize[2].cx;
	}

	if(menuBandWindow.IsWindow()) {
		RECT clientRectangle = {0};
		menuBandWindow.GetClientRect(&clientRectangle);
		AdjustMDINonClientButtonSize(clientRectangle.bottom - clientRectangle.top);
	}
}

void ReBar::AdjustMDINonClientButtonSize(int height)
{
	if(height > 2) {
		for(int i = 0; i < 3; ++i) {
			if(mdiStatus.buttonSize[i].cy > height - 2) {
				mdiStatus.buttonSize[i].cx = (mdiStatus.buttonSize[i].cx * (height - 2)) / mdiStatus.buttonSize[i].cy;
				mdiStatus.buttonSize[i].cy = height - 2;
			}
		}
		CTheme themingEngine;
		if(flags.usingThemes) {
			themingEngine.OpenThemeData(*this, VSCLASS_WINDOW);
		}
		if(!themingEngine.IsThemeNull()) {
			mdiStatus.mdiChildWindowButtonAreaWidth = mdiStatus.buttonSize[0].cx + mdiStatus.buttonSize[1].cx + mdiStatus.buttonSize[2].cx + 1;
		} else {
			mdiStatus.mdiChildWindowButtonAreaWidth = mdiStatus.buttonSize[0].cx + mdiStatus.buttonSize[1].cx + 2 + mdiStatus.buttonSize[2].cx;
		}
	}
}

void ReBar::CalcMDIChildIconRectangle(int availableWidth, int availableHeight, LPRECT pIconRectangle, BOOL invertX/* = FALSE*/) const
{
	int top = max((availableHeight - mdiStatus.smallIconSize.cy) >> 1, 0);
	if(invertX) {
		SetRect(pIconRectangle, availableWidth - mdiStatus.smallIconSize.cx, top, availableWidth, top + mdiStatus.smallIconSize.cy);
	} else {
		SetRect(pIconRectangle, 0, top, mdiStatus.smallIconSize.cx, top + mdiStatus.smallIconSize.cy);
	}
}

void ReBar::CalcMDIChildButtonRectangles(int availableWidth, int availableHeight, RECT buttonRectangles[3], BOOL invertX/* = FALSE*/) const
{
	for(int i = 0; i < 3; ++i) {
		buttonRectangles[i].top = max((availableHeight - mdiStatus.buttonSize[i].cy) >> 1, 0);
		buttonRectangles[i].bottom = buttonRectangles[i].top + mdiStatus.buttonSize[i].cy;
	}
	if(invertX) {
		CTheme themingEngine;
		if(flags.usingThemes) {
			themingEngine.OpenThemeData(*this, VSCLASS_WINDOW);
		}
		if(!themingEngine.IsThemeNull()) {
			buttonRectangles[0].left = 1;
			buttonRectangles[0].right = buttonRectangles[0].left + mdiStatus.buttonSize[0].cx;
			buttonRectangles[1].left = buttonRectangles[0].right;
			buttonRectangles[1].right = buttonRectangles[1].left + mdiStatus.buttonSize[1].cx;
			buttonRectangles[2].left = buttonRectangles[1].right;
			buttonRectangles[2].right = buttonRectangles[2].left + mdiStatus.buttonSize[2].cx;
		} else {
			buttonRectangles[0].left = 0;
			buttonRectangles[0].right = buttonRectangles[0].left + mdiStatus.buttonSize[0].cx;
			buttonRectangles[1].left = buttonRectangles[0].right + 2;
			buttonRectangles[1].right = buttonRectangles[1].left + mdiStatus.buttonSize[1].cx;
			buttonRectangles[2].left = buttonRectangles[1].right;
			buttonRectangles[2].right = buttonRectangles[2].left + mdiStatus.buttonSize[2].cx;
		}
	} else {
		CTheme themingEngine;
		if(flags.usingThemes) {
			themingEngine.OpenThemeData(*this, VSCLASS_WINDOW);
		}
		if(!themingEngine.IsThemeNull()) {
			buttonRectangles[0].right = availableWidth - 1;
			buttonRectangles[0].left = buttonRectangles[0].right - mdiStatus.buttonSize[0].cx;
			buttonRectangles[1].right = buttonRectangles[0].left;
			buttonRectangles[1].left = buttonRectangles[1].right - mdiStatus.buttonSize[1].cx;
			buttonRectangles[2].right = buttonRectangles[1].left;
			buttonRectangles[2].left = buttonRectangles[2].right - mdiStatus.buttonSize[2].cx;
		} else {
			buttonRectangles[0].right = availableWidth;
			buttonRectangles[0].left = buttonRectangles[0].right - mdiStatus.buttonSize[0].cx;
			buttonRectangles[1].right = buttonRectangles[0].left - 2;
			buttonRectangles[1].left = buttonRectangles[1].right - mdiStatus.buttonSize[1].cx;
			buttonRectangles[2].right = buttonRectangles[1].left;
			buttonRectangles[2].left = buttonRectangles[2].right - mdiStatus.buttonSize[2].cx;
		}
	}
}

void ReBar::DrawMDIChildButton(CMemoryDC& memoryDC, RECT buttonRectangles[3], int button)
{
	CTheme themingEngine;
	if(flags.usingThemes) {
		themingEngine.OpenThemeData(*this, VSCLASS_WINDOW);
	}
	if(!themingEngine.IsThemeNull()) {
		if(button == -1 || button == 0) {
			themingEngine.DrawThemeBackground(memoryDC, WP_MDICLOSEBUTTON, mdiStatus.mdiFrameIsActive ? (mdiStatus.pressedButton == 0 ? CBS_PUSHED : CBS_NORMAL) : CBS_DISABLED, &buttonRectangles[0], NULL);
		}
		if(button == -1 || button == 1) {
			themingEngine.DrawThemeBackground(memoryDC, WP_MDIRESTOREBUTTON, mdiStatus.mdiFrameIsActive ? (mdiStatus.pressedButton == 1 ? RBS_PUSHED : RBS_NORMAL) : RBS_DISABLED, &buttonRectangles[1], NULL);
		}
		if(button == -1 || button == 2) {
			themingEngine.DrawThemeBackground(memoryDC, WP_MDIMINBUTTON, mdiStatus.mdiFrameIsActive ? (mdiStatus.pressedButton == 2 ? MINBS_PUSHED : MINBS_NORMAL) : MINBS_DISABLED, &buttonRectangles[2], NULL);
		}
	} else {
		if(button == -1 || button == 0) {
			memoryDC.DrawFrameControl(&buttonRectangles[0], DFC_CAPTION, DFCS_CAPTIONCLOSE | (mdiStatus.pressedButton == 0 ? DFCS_PUSHED : 0));
		}
		if(button == -1 || button == 1) {
			memoryDC.DrawFrameControl(&buttonRectangles[1], DFC_CAPTION, DFCS_CAPTIONRESTORE | (mdiStatus.pressedButton == 1 ? DFCS_PUSHED : 0));
		}
		if(button == -1 || button == 2) {
			memoryDC.DrawFrameControl(&buttonRectangles[2], DFC_CAPTION, DFCS_CAPTIONMIN | (mdiStatus.pressedButton == 2 ? DFCS_PUSHED : 0));
		}
	}
}


UINT ReBar::GetNewBandID(void)
{
	static UINT nextID = 0;

	return ++nextID;
}

int ReBar::HitTest(LONG x, LONG y, UINT* pFlags/*, BOOL ignoreBoundingBoxDefinition = FALSE*/)
{
	ATLASSERT(IsWindow());

	UINT hitTestFlags = 0;
	if(pFlags) {
		hitTestFlags = *pFlags;
	}
	RBHITTESTINFO hitTestInfo = {{x, y}, hitTestFlags, 0};
	int bandIndex = -1;

	POINT pt = {x, y};
	ClientToScreen(&pt);
	CRect rc;
	GetWindowRect(&rc);
	if(rc.PtInRect(pt)) {
		bandIndex = static_cast<int>(SendMessage(RB_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo)));
		ATLASSERT(bandIndex == hitTestInfo.iBand);
	} else {
		hitTestInfo.flags = static_cast<HitTestConstants>(0);
		if(pt.x < rc.left) {
			hitTestInfo.flags = static_cast<HitTestConstants>(hitTestInfo.flags | htToLeft);
		} else if(pt.x >= rc.right) {
			hitTestInfo.flags = static_cast<HitTestConstants>(hitTestInfo.flags | htToRight);
		}
		if(pt.y < rc.top) {
			hitTestInfo.flags = static_cast<HitTestConstants>(hitTestInfo.flags | htAbove);
		} else if(pt.y >= rc.bottom) {
			hitTestInfo.flags = static_cast<HitTestConstants>(hitTestInfo.flags | htBelow);
		}
	}

	hitTestFlags = hitTestInfo.flags;
	if(pFlags) {
		*pFlags = hitTestFlags;
	}
	/*TODO: if(!ignoreBoundingBoxDefinition && ((properties.bandBoundingBoxDefinition & hitTestFlags) != hitTestFlags)) {
		bandIndex = -1;
	}*/
	return bandIndex;
}

BOOL ReBar::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

void ReBar::RebuildAcceleratorTable(void)
{
	if(properties.hAcceleratorTable) {
		DestroyAcceleratorTable(properties.hAcceleratorTable);
		properties.hAcceleratorTable = NULL;
	}

	// create a new accelerator table
	REBARBANDINFO band = {0};
	band.cbSize = RunTimeHelper::SizeOf_REBARBANDINFO();
	band.fMask = RBBIM_TEXT;
	band.cch = MAX_BANDTEXTLENGTH;
	band.lpText = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (band.cch + 1) * sizeof(TCHAR)));
	if(band.lpText) {
		int numberOfBands = SendMessage(RB_GETBANDCOUNT, 0, 0);
		TCHAR* pAcceleratorChars = static_cast<TCHAR*>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, numberOfBands * sizeof(TCHAR)));
		if(pAcceleratorChars) {
			int numberOfBandsWithAccelerator = 0;
			for(int bandIndex = 0; bandIndex < numberOfBands; ++bandIndex) {
				SendMessage(RB_GETBANDINFO, bandIndex, reinterpret_cast<LPARAM>(&band));

				pAcceleratorChars[bandIndex] = TEXT('\0');
				if(band.lpText) {
					for(int i = lstrlen(band.lpText) - 1; i > 0; --i) {
						if((band.lpText[i - 1] == TEXT('&')) && (band.lpText[i] != TEXT('&'))) {
							++numberOfBandsWithAccelerator;
							pAcceleratorChars[bandIndex] = band.lpText[i];
							break;
						}
					}
				}
			}

			if(numberOfBandsWithAccelerator > 0) {
				LPACCEL pAccelerators = static_cast<LPACCEL>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, (numberOfBandsWithAccelerator * 4) * sizeof(ACCEL)));
				if(pAccelerators) {
					int i = 0;
					for(int bandIndex = 0; bandIndex < numberOfBands; ++bandIndex) {
						if(pAcceleratorChars[bandIndex] != TEXT('\0')) {
							pAccelerators[i * 4].cmd = static_cast<WORD>(bandIndex);
							pAccelerators[i * 4].fVirt = FALT;
							pAccelerators[i * 4].key = static_cast<WORD>(tolower(pAcceleratorChars[bandIndex]));

							pAccelerators[i * 4 + 1].cmd = static_cast<WORD>(bandIndex);
							pAccelerators[i * 4 + 1].fVirt = 0;
							pAccelerators[i * 4 + 1].key = static_cast<WORD>(tolower(pAcceleratorChars[bandIndex]));

							pAccelerators[i * 4 + 2].cmd = static_cast<WORD>(bandIndex);
							pAccelerators[i * 4 + 2].fVirt = FVIRTKEY | FALT;
							pAccelerators[i * 4 + 2].key = LOBYTE(VkKeyScan(pAcceleratorChars[bandIndex]));

							pAccelerators[i * 4 + 3].cmd = static_cast<WORD>(bandIndex);
							pAccelerators[i * 4 + 3].fVirt = FVIRTKEY | FALT | FSHIFT;
							pAccelerators[i * 4 + 3].key = LOBYTE(VkKeyScan(pAcceleratorChars[bandIndex]));

							++i;
						}
					}
					properties.hAcceleratorTable = CreateAcceleratorTable(pAccelerators, numberOfBandsWithAccelerator * 4);
					HeapFree(GetProcessHeap(), 0, pAccelerators);
				}
			}

			HeapFree(GetProcessHeap(), 0, pAcceleratorChars);
		}
		HeapFree(GetProcessHeap(), 0, band.lpText);
	}

	// report the new accelerator table to the container
	CComQIPtr<IOleControlSite, &IID_IOleControlSite> pSite(m_spClientSite);
	if(pSite) {
		ATLVERIFY(SUCCEEDED(pSite->OnControlInfoChanged()));
	}
}


BOOL ReBar::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}