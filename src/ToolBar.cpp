// ToolBar.cpp: Superclasses ToolbarWindow32.

#include "stdafx.h"
#include "ToolBar.h"
#include "ClassFactory.h"

#pragma comment(lib, "comctl32.lib")


__declspec(selectany) HHOOK ToolBar::hCBTHook = NULL;
__declspec(selectany) ToolBar* ToolBar::pCurrentToolbar = NULL;


// initialize complex constants
FONTDESC ToolBar::Properties::FontProperty::defaultFont = {
    sizeof(FONTDESC),
    OLESTR("MS Sans Serif"),
    120000,
    FW_NORMAL,
    ANSI_CHARSET,
    FALSE,
    FALSE,
    FALSE
};


ToolBar::ToolBar()
{
	SIZEL size = {100, 24};
	AtlPixelToHiMetric(&size, &m_sizeExtent);

	properties.font.InitializePropertyWatcher(this, DISPID_TB_FONT);
	properties.mouseIcon.InitializePropertyWatcher(this, DISPID_TB_MOUSEICON);

	// always create a window, even if the container supports windowless controls
	m_bWindowOnly = TRUE;

	// initialize
	buttonUnderMouse = -1;
	droppedDownButton = -1;

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

	if(RunTimeHelper::IsVista()) {
		CoCreateInstance(CLSID_WICImagingFactory, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pWICImagingFactory));
		ATLASSUME(pWICImagingFactory);
	}
}


//////////////////////////////////////////////////////////////////////
// implementation of ISupportErrorInfo
STDMETHODIMP ToolBar::InterfaceSupportsErrorInfo(REFIID interfaceToCheck)
{
	if(InlineIsEqualGUID(IID_IToolBar, interfaceToCheck)) {
		return S_OK;
	}
	return S_FALSE;
}
// implementation of ISupportErrorInfo
//////////////////////////////////////////////////////////////////////


STDMETHODIMP ToolBar::Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog)
{
	HRESULT hr = IPersistPropertyBagImpl<ToolBar>::Load(pPropertyBag, pErrorLog);
	/*if(SUCCEEDED(hr)) {
	}*/
	return hr;
}

STDMETHODIMP ToolBar::Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties)
{
	HRESULT hr = IPersistPropertyBagImpl<ToolBar>::Save(pPropertyBag, clearDirtyFlag, saveAllProperties);
	/*if(SUCCEEDED(hr)) {
	}*/
	return hr;
}

STDMETHODIMP ToolBar::GetSizeMax(ULARGE_INTEGER* pSize)
{
	ATLASSERT_POINTER(pSize, ULARGE_INTEGER);
	if(!pSize) {
		return E_POINTER;
	}

	pSize->LowPart = 0;
	pSize->HighPart = 0;
	pSize->QuadPart = sizeof(LONG/*signature*/) + sizeof(LONG/*version*/) + sizeof(LONG/*subSignature*/) + sizeof(DWORD/*atlVersion*/) + sizeof(m_sizeExtent);

	// we've 32 VT_I4 properties...
	pSize->QuadPart += 32 * (sizeof(VARTYPE) + sizeof(LONG));
	// ...and 17 VT_BOOL properties...
	pSize->QuadPart += 17 * (sizeof(VARTYPE) + sizeof(VARIANT_BOOL));

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

STDMETHODIMP ToolBar::Load(LPSTREAM pStream)
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
	if(subSignature != 0x02020202/*4x 0x02 (-> ToolBar)*/) {
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
	VARIANT_BOOL allowCustomization = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL alwaysDisplayButtonText = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL anchorHighlighting = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	AppearanceConstants appearance = static_cast<AppearanceConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BackStyleConstants backStyle = static_cast<BackStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	BorderStyleConstants borderStyle = static_cast<BorderStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS buttonHeight = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ButtonStyleConstants buttonStyle = static_cast<ButtonStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	ButtonTextPositionConstants buttonTextPosition = static_cast<ButtonTextPositionConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS buttonWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DisabledEventsConstants disabledEvents = static_cast<DisabledEventsConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displayMenuDivider = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL displayPartiallyClippedButtons = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL dontRedraw = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG dragClickTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	DragDropCustomizationModifierKeyConstants dragDropCustomizationModifierKey = static_cast<DragDropCustomizationModifierKeyConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS dropDownGap = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL enabled = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS firstButtonIndentation = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL focusOnClick = propertyValue.boolVal;

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
	OLE_COLOR highlightColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS horizontalButtonPadding = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS horizontalButtonSpacing = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS horizontalIconCaptionGap = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	HAlignmentConstants horizontalTextAlignment = static_cast<HAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG hoverTime = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_COLOR insertMarkColor = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS maximumButtonWidth = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	LONG maximumTextRows = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	MenuBarThemeConstants menuBarTheme = static_cast<MenuBarThemeConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL menuMode = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_XSIZE_PIXELS minimumButtonWidth = propertyValue.lVal;

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
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL multiColumn = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	NormalDropDownButtonStyleConstants normalDropDownButtonStyle = static_cast<NormalDropDownButtonStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLEDragImageStyleConstants oleDragImageStyle = static_cast<OLEDragImageStyleConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OrientationConstants orientation = static_cast<OrientationConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL processContextMenuKeys = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL raiseCustomDrawEventOnEraseBackground = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	RegisterForOLEDragDropConstants registerForOLEDragDrop = static_cast<RegisterForOLEDragDropConstants>(propertyValue.lVal);
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
	VARIANT_BOOL showToolTips = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL supportOLEDragImages = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useMnemonics = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL useSystemFont = propertyValue.boolVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS verticalButtonPadding = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	OLE_YSIZE_PIXELS verticalButtonSpacing = propertyValue.lVal;
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_I4, &propertyValue))) {
		return hr;
	}
	VAlignmentConstants verticalTextAlignment = static_cast<VAlignmentConstants>(propertyValue.lVal);
	if(FAILED(hr = pfnReadVariantFromStream(pStream, VT_BOOL, &propertyValue))) {
		return hr;
	}
	VARIANT_BOOL wrapButtons = propertyValue.boolVal;


	hr = put_AllowCustomization(allowCustomization);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AlwaysDisplayButtonText(alwaysDisplayButtonText);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_AnchorHighlighting(anchorHighlighting);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Appearance(appearance);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BackStyle(backStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_BorderStyle(borderStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ButtonHeight(buttonHeight);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ButtonStyle(buttonStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ButtonTextPosition(buttonTextPosition);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ButtonWidth(buttonWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisabledEvents(disabledEvents);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayMenuDivider(displayMenuDivider);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DisplayPartiallyClippedButtons(displayPartiallyClippedButtons);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DontRedraw(dontRedraw);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragClickTime(dragClickTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DragDropCustomizationModifierKey(dragDropCustomizationModifierKey);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_DropDownGap(dropDownGap);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Enabled(enabled);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FirstButtonIndentation(firstButtonIndentation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_FocusOnClick(focusOnClick);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Font(pFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HighlightColor(highlightColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HorizontalButtonPadding(horizontalButtonPadding);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HorizontalButtonSpacing(horizontalButtonSpacing);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HorizontalIconCaptionGap(horizontalIconCaptionGap);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HorizontalTextAlignment(horizontalTextAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_HoverTime(hoverTime);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_InsertMarkColor(insertMarkColor);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaximumButtonWidth(maximumButtonWidth);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MaximumTextRows(maximumTextRows);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MenuBarTheme(menuBarTheme);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MenuMode(menuMode);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_MinimumButtonWidth(minimumButtonWidth);
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
	hr = put_MultiColumn(multiColumn);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_NormalDropDownButtonStyle(normalDropDownButtonStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_OLEDragImageStyle(oleDragImageStyle);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_Orientation(orientation);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_ProcessContextMenuKeys(processContextMenuKeys);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RaiseCustomDrawEventOnEraseBackground(raiseCustomDrawEventOnEraseBackground);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_RegisterForOLEDragDrop(registerForOLEDragDrop);
	if(FAILED(hr)) {
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
	hr = put_ShowToolTips(showToolTips);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_SupportOLEDragImages(supportOLEDragImages);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseMnemonics(useMnemonics);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_UseSystemFont(useSystemFont);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_VerticalButtonPadding(verticalButtonPadding);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_VerticalButtonSpacing(verticalButtonSpacing);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_VerticalTextAlignment(verticalTextAlignment);
	if(FAILED(hr)) {
		return hr;
	}
	hr = put_WrapButtons(wrapButtons);
	if(FAILED(hr)) {
		return hr;
	}

	SetDirty(FALSE);
	return S_OK;
}

STDMETHODIMP ToolBar::Save(LPSTREAM pStream, BOOL clearDirtyFlag)
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
	LONG subSignature = 0x02020202/*4x 0x02 (-> ToolBar)*/;
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
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.allowCustomization);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.alwaysDisplayButtonText);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.anchorHighlighting);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.appearance;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.backStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.borderStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.buttonHeight;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.buttonStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.buttonTextPosition;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.buttonWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.disabledEvents;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayMenuDivider);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.displayPartiallyClippedButtons);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.dontRedraw);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.dragClickTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.dragDropCustomizationModifierKey;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.dropDownGap;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.enabled);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.firstButtonIndentation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.focusOnClick);
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
	propertyValue.lVal = properties.highlightColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.horizontalButtonPadding;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.horizontalButtonSpacing;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.horizontalIconCaptionGap;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.horizontalTextAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.hoverTime;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.insertMarkColor;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maximumButtonWidth;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.maximumTextRows;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.menuBarTheme;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.menuMode);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.minimumButtonWidth;
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
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.multiColumn);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.normalDropDownButtonStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.oleDragImageStyle;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.orientation;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.raiseCustomDrawEventOnEraseBackground);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.registerForOLEDragDrop;
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
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.showToolTips);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useMnemonics);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.useSystemFont);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_I4;
	propertyValue.lVal = properties.verticalButtonPadding;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.verticalButtonSpacing;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.lVal = properties.verticalTextAlignment;
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}
	propertyValue.vt = VT_BOOL;
	propertyValue.boolVal = BOOL2VARIANTBOOL(properties.wrapButtons);
	if(FAILED(hr = WriteVariantToStream(pStream, &propertyValue))) {
		return hr;
	}

	if(clearDirtyFlag) {
		SetDirty(FALSE);
	}
	return S_OK;
}


HWND ToolBar::Create(HWND hWndParent, ATL::_U_RECT rect/* = NULL*/, LPCTSTR szWindowName/* = NULL*/, DWORD dwStyle/* = 0*/, DWORD dwExStyle/* = 0*/, ATL::_U_MENUorID MenuOrID/* = 0U*/, LPVOID lpCreateParam/* = NULL*/)
{
	INITCOMMONCONTROLSEX data = {0};
	data.dwSize = sizeof(data);
	data.dwICC = ICC_BAR_CLASSES;
	InitCommonControlsEx(&data);

	dwStyle = GetStyleBits();
	dwExStyle = GetExStyleBits();
	return CComControl<ToolBar>::Create(hWndParent, rect, szWindowName, dwStyle, dwExStyle, MenuOrID, lpCreateParam);
}

HRESULT ToolBar::OnDraw(ATL_DRAWINFO& drawInfo)
{
	if(IsInDesignMode()) {
		CAtlString text = TEXT("ToolBar ");
		CComBSTR buffer;
		get_Version(&buffer);
		text += buffer;
		SetTextAlign(drawInfo.hdcDraw, TA_CENTER | TA_BASELINE);
		TextOut(drawInfo.hdcDraw, drawInfo.prcBounds->left + (drawInfo.prcBounds->right - drawInfo.prcBounds->left) / 2, drawInfo.prcBounds->top + (drawInfo.prcBounds->bottom - drawInfo.prcBounds->top) / 2, text, text.GetLength());
	}

	return S_OK;
}

void ToolBar::OnFinalMessage(HWND /*hWnd*/)
{
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	Release();
}

STDMETHODIMP ToolBar::IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite)
{
	HRESULT hr = CComControl<ToolBar>::IOleObject_SetClientSite(pClientSite);

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

STDMETHODIMP ToolBar::OnDocWindowActivate(BOOL /*fActivate*/)
{
	return S_OK;
}

BOOL ToolBar::PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue)
{
	if((pMessage->message >= WM_KEYFIRST) && (pMessage->message <= WM_KEYLAST)) {
		/* TODO: This code seems right, but if we activate it, keys like VK_LEFT and VK_RIGHT are handled
		 *       twice.
		 *       The native control handles WM_KEYDOWN, WM_KEYUP and WM_CHAR, but only if the control has
		 *       the keyboard focus. The handled keys are: VK_LEFT, VK_RIGHT, VK_DOWN, VK_UP, VK_ESCAPE,
		 *       VK_SPACE and VK_RETURN.
		 */
		/*if(SendMessage(TB_TRANSLATEACCELERATOR, 0, reinterpret_cast<LPARAM>(pMessage))) {
			hReturnValue = S_FALSE;
			return TRUE;
		}*/
		LRESULT dialogCode = SendMessage(pMessage->hwnd, WM_GETDLGCODE, 0, 0);
		if(pMessage->wParam == VK_TAB) {
			if(dialogCode & DLGC_WANTTAB) {
				hReturnValue = S_FALSE;
				return TRUE;
			}
		}
		switch (pMessage->message) {
			case WM_KEYUP:
			case WM_KEYDOWN:
        switch (pMessage->wParam) {
					case VK_RIGHT:
					case VK_LEFT:
					case VK_UP:
					case VK_DOWN:
					case VK_ESCAPE:
					case VK_SPACE:
					case VK_RETURN:
						hReturnValue = S_FALSE;
						return TRUE;
				}
				break;
			case WM_CHAR:
				switch (pMessage->wParam) {
					case VK_ESCAPE:
					case VK_SPACE:
					case VK_RETURN:
						hReturnValue = S_FALSE;
						return TRUE;
				}
				break;
		}
	}
	return CComControl<ToolBar>::PreTranslateAccelerator(pMessage, hReturnValue);
}


LRESULT CALLBACK ToolBar::CBTProc(int code, WPARAM wParam, LPARAM lParam)
{
	LRESULT lr = 0;
	BOOL consumeMessage = FALSE;

	if(code == HCBT_CREATEWND) {
		TCHAR pBuffer[7] = {0};

		HWND hWndMenu = reinterpret_cast<HWND>(wParam);
		GetClassName(hWndMenu, pBuffer, 7);
		if(!lstrcmp(TEXT("#32768"), pBuffer)) {
			pCurrentToolbar->menuModeState.menuWindowStack.Push(hWndMenu);
		}
	}
	if(code == HCBT_DESTROYWND) {
		TCHAR pBuffer[7] = {0};

		HWND hWndMenu = reinterpret_cast<HWND>(wParam);
		GetClassName(hWndMenu, pBuffer, 7);
		if(!lstrcmp(TEXT("#32768"), pBuffer)) {
			ATLASSERT(hWndMenu == pCurrentToolbar->menuModeState.menuWindowStack.GetCurrent());
			pCurrentToolbar->menuModeState.menuWindowStack.Pop();
		}
	}

	if(!consumeMessage) {
		lr = CallNextHookEx(hCBTHook, code, wParam, lParam);
	}
	return lr;
}


HIMAGELIST ToolBar::CreateLegacyDragImage(int buttonIndex, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle)
{
	/********************************************************************************************************
	 * Known problems:                                                                                      *
	 * - No multi-line support for button texts                                                             *
	 ********************************************************************************************************/

	// retrieve button details
	TCHAR pButtonTextBuffer[MAX_BUTTONTEXTLENGTH + 1];
	ZeroMemory(pButtonTextBuffer, (MAX_BUTTONTEXTLENGTH + 1) * sizeof(TCHAR));
	int iconToDraw = I_IMAGENONE;
	int imageListIndex = I_IMAGENONE;
	HIMAGELIST hSourceImageList = NULL;
	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	button.dwMask = TBIF_BYINDEX | TBIF_COMMAND | TBIF_IMAGE | TBIF_STATE | TBIF_STYLE;
	if(SendMessage(TB_GETBUTTONINFO, buttonIndex, reinterpret_cast<LPARAM>(&button)) == buttonIndex) {
		if(!(button.fsStyle & BTNS_SEP)) {
			if(button.iImage != I_IMAGENONE) {
				iconToDraw = LOWORD(button.iImage);
				imageListIndex = HIWORD(button.iImage);
			}
			SendMessage(TB_GETBUTTONTEXT, button.idCommand, reinterpret_cast<LPARAM>(pButtonTextBuffer));
		}
	}
	BOOL buttonIsEnabled = button.fsState & TBSTATE_ENABLED;
	BOOL buttonIsHot = (buttonIndex == static_cast<int>(SendMessage(TB_GETHOTITEM, 0, 0)));
	BOOL buttonIsPressed = button.fsState & TBSTATE_PRESSED;
	BOOL hasSpecialDisabledImageList = FALSE;

	if(imageListIndex != I_IMAGENONE) {
		if(!buttonIsEnabled) {
			hSourceImageList = reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETDISABLEDIMAGELIST, imageListIndex, 0));
			hasSpecialDisabledImageList = (hSourceImageList != NULL);
		} else if(buttonIsPressed) {
			hSourceImageList = reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETPRESSEDIMAGELIST, imageListIndex, 0));
		} else if(buttonIsHot) {
			hSourceImageList = reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETHOTIMAGELIST, imageListIndex, 0));
		}
		if(!hSourceImageList) {
			hSourceImageList = reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETIMAGELIST, imageListIndex, 0));
		}
	}
	SIZE imageSize = {0};
	if(hSourceImageList) {
		ImageList_GetIconSize(hSourceImageList, reinterpret_cast<PINT>(&imageSize.cx), reinterpret_cast<PINT>(&imageSize.cy));
	}

	// retrieve window details
	BOOL layoutRTL = ((GetExStyle() & WS_EX_LAYOUTRTL) == WS_EX_LAYOUTRTL);
	// TODO: Multiline
	DWORD textDrawStyle = properties.drawTextFlags;
	if(GetExStyle() & WS_EX_RTLREADING) {
		textDrawStyle |= DT_RTLREADING;
	}
	if(button.fsState & TBSTATE_ELLIPSES) {
		textDrawStyle |= DT_END_ELLIPSIS;
	} else {
		textDrawStyle &= ~DT_END_ELLIPSIS;
	}
	if(button.fsStyle & BTNS_NOPREFIX) {
		textDrawStyle |= DT_NOPREFIX;
	} else {
		textDrawStyle &= ~DT_NOPREFIX;
	}
	BOOL textOnRight = ((GetStyle() & TBSTYLE_LIST) == TBSTYLE_LIST);
	BOOL displayButtonText = !textOnRight || !(SendMessage(TB_GETEXTENDEDSTYLE, 0, 0) & TBSTYLE_EX_MIXEDBUTTONS);
	if(!displayButtonText) {
		displayButtonText = button.fsStyle & BTNS_SHOWTEXT;
	}
	if(displayButtonText) {
		displayButtonText = (lstrlen(pButtonTextBuffer) > 0);
	}

	// create the DCs we'll draw into
	HDC hCompatibleDC = GetDC();
	CDC memoryDC;
	memoryDC.CreateCompatibleDC(hCompatibleDC);
	CDC maskMemoryDC;
	maskMemoryDC.CreateCompatibleDC(hCompatibleDC);

	// calculate the bounding rectangle
	CRect buttonBoundingRect;
	CRect labelBoundingRect;
	CRect iconBoundingRect;

	SendMessage(TB_GETITEMRECT, buttonIndex, reinterpret_cast<LPARAM>(&buttonBoundingRect));
	if(pBoundingRectangle) {
		*pBoundingRectangle = buttonBoundingRect;
	}
	labelBoundingRect = buttonBoundingRect;
	if(hSourceImageList) {
		if(textOnRight) {
			iconBoundingRect.top = labelBoundingRect.top + ((buttonBoundingRect.bottom - buttonBoundingRect.top - imageSize.cy) >> 1);
			iconBoundingRect.bottom = iconBoundingRect.top + imageSize.cy;
			iconBoundingRect.left = labelBoundingRect.left + (GetSystemMetrics(SM_CXEDGE) << 1);
			iconBoundingRect.right = iconBoundingRect.left + imageSize.cx;
		} else {
			if(displayButtonText) {
				iconBoundingRect.top = labelBoundingRect.top + GetSystemMetrics(SM_CYEDGE) + 1;
			} else {
				iconBoundingRect.top = labelBoundingRect.top + ((buttonBoundingRect.bottom - buttonBoundingRect.top - imageSize.cy) >> 1);
			}
			iconBoundingRect.bottom = iconBoundingRect.top + imageSize.cy;
			iconBoundingRect.left = labelBoundingRect.left + ((buttonBoundingRect.right - buttonBoundingRect.left - imageSize.cx) >> 1);
			iconBoundingRect.right = iconBoundingRect.left + imageSize.cx;
		}
	}
	if(displayButtonText) {
		if(textOnRight) {
			if(iconToDraw != I_IMAGENONE) {
				labelBoundingRect.left = iconBoundingRect.right + (properties.horizontalIconCaptionGap == -1 ? GetSystemMetrics(SM_CXEDGE) : properties.horizontalIconCaptionGap - GetSystemMetrics(SM_CXEDGE)) + 1;
			}
		} else {
			if(iconToDraw == I_IMAGENONE) {
				labelBoundingRect.top += GetSystemMetrics(SM_CYEDGE) + 1 + GetSystemMetrics(SM_CYSMICON) + 1;
			} else {
				labelBoundingRect.top = iconBoundingRect.bottom + 1;
			}
		}
	}

	// calculate drag image size and upper-left corner
	SIZE dragImageSize = {0};
	if(pUpperLeftPoint) {
		pUpperLeftPoint->x = buttonBoundingRect.left;
		pUpperLeftPoint->y = buttonBoundingRect.top;
	}
	dragImageSize.cx = buttonBoundingRect.Width();
	dragImageSize.cy = buttonBoundingRect.Height();

	// offset RECTs
	SIZE offset = {0};
	offset.cx = buttonBoundingRect.left;
	offset.cy = buttonBoundingRect.top;
	labelBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	iconBoundingRect.OffsetRect(-offset.cx, -offset.cy);
	buttonBoundingRect.OffsetRect(-offset.cx, -offset.cy);

	// setup the DCs we'll draw into
	memoryDC.SetBkColor(GetSysColor(COLOR_WINDOW));
	memoryDC.SetTextColor(GetSysColor(COLOR_WINDOWTEXT));
	memoryDC.SetBkMode(TRANSPARENT);

	// create drag image bitmap
	/* NOTE: We prefer creating 32bpp drag images, because this improves performance of
	         ToolBarButtonContainer::CreateDragImage(). */
	BOOL doAlphaChannelProcessing = RunTimeHelper::IsCommCtrl6();
	BITMAPINFO bitmapInfo = {0};
	if(doAlphaChannelProcessing) {
		bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
		bitmapInfo.bmiHeader.biWidth = dragImageSize.cx;
		bitmapInfo.bmiHeader.biHeight = -dragImageSize.cy;
		bitmapInfo.bmiHeader.biPlanes = 1;
		bitmapInfo.bmiHeader.biBitCount = 32;
		bitmapInfo.bmiHeader.biCompression = BI_RGB;
	}
	CBitmap dragImage;
	LPRGBQUAD pDragImageBits = NULL;
	if(doAlphaChannelProcessing) {
		dragImage.CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
	} else {
		dragImage.CreateCompatibleBitmap(hCompatibleDC, dragImageSize.cx, dragImageSize.cy);
	}
	HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(dragImage);
	CBitmap dragImageMask;
	dragImageMask.CreateBitmap(dragImageSize.cx, dragImageSize.cy, 1, 1, NULL);
	HBITMAP hPreviousBitmapMask = maskMemoryDC.SelectBitmap(dragImageMask);

	// initialize the bitmap
	RECT rc = buttonBoundingRect;
	if(doAlphaChannelProcessing && pDragImageBits) {
		// we need a transparent background
		LPRGBQUAD pPixel = pDragImageBits;
		for(int y = 0; y < dragImageSize.cy; ++y) {
			for(int x = 0; x < dragImageSize.cx; ++x, ++pPixel) {
				pPixel->rgbRed = 0xFF;
				pPixel->rgbGreen = 0xFF;
				pPixel->rgbBlue = 0xFF;
				pPixel->rgbReserved = 0x00;
			}
		}
	} else {
		memoryDC.FillRect(&rc, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));
	}
	maskMemoryDC.FillRect(&rc, static_cast<HBRUSH>(GetStockObject(WHITE_BRUSH)));

	CFontHandle font = reinterpret_cast<HFONT>(SendMessage(WM_GETFONT, 0, 0));
	HFONT hPreviousFont = NULL;
	if(!font.IsNull()) {
		hPreviousFont = memoryDC.SelectFont(font);
	}

	// draw the icon
	if(hSourceImageList) {
		ImageList_DrawEx(hSourceImageList, iconToDraw, memoryDC, iconBoundingRect.left, iconBoundingRect.top, imageSize.cx, imageSize.cy, CLR_NONE, CLR_NONE, (!buttonIsEnabled && !hasSpecialDisabledImageList ? ILD_SELECTED : ILD_NORMAL));
		ImageList_Draw(hSourceImageList, iconToDraw, maskMemoryDC, iconBoundingRect.left, iconBoundingRect.top, ILD_MASK);
	}

	if(displayButtonText) {
		// draw the text
		memoryDC.DrawText(pButtonTextBuffer, lstrlen(pButtonTextBuffer), &labelBoundingRect, textDrawStyle);
		COLORREF bkColor = memoryDC.GetBkColor();
		for(int y = labelBoundingRect.top; y <= labelBoundingRect.bottom; ++y) {
			for(int x = labelBoundingRect.left; x <= labelBoundingRect.right; ++x) {
				if(memoryDC.GetPixel(x, y) != bkColor) {
					maskMemoryDC.SetPixelV(x, y, 0x00000000);
				}
			}
		}
	}

	if(doAlphaChannelProcessing && pDragImageBits) {
		// correct the alpha channel
		LPRGBQUAD pPixel = pDragImageBits;
		POINT pt;
		for(pt.y = 0; pt.y < dragImageSize.cy; ++pt.y) {
			for(pt.x = 0; pt.x < dragImageSize.cx; ++pt.x, ++pPixel) {
				if(layoutRTL) {
					// we're working on raw data, so we've to handle WS_EX_LAYOUTRTL ourselves
					POINT pt2 = pt;
					pt2.x = dragImageSize.cx - pt.x - 1;
					if(maskMemoryDC.GetPixel(pt2.x, pt2.y) == 0x00000000) {
						if(pPixel->rgbReserved == 0x00) {
							pPixel->rgbReserved = 0xFF;
						}
					}
				} else {
					// layout is left to right
					if(maskMemoryDC.GetPixel(pt.x, pt.y) == 0x00000000) {
						if(pPixel->rgbReserved == 0x00) {
							pPixel->rgbReserved = 0xFF;
						}
					}
				}
			}
		}
	}

	memoryDC.SelectBitmap(hPreviousBitmap);
	maskMemoryDC.SelectBitmap(hPreviousBitmapMask);
	if(hPreviousFont) {
		memoryDC.SelectFont(hPreviousFont);
	}

	// create the imagelist
	HIMAGELIST hDragImageList = ImageList_Create(dragImageSize.cx, dragImageSize.cy, (RunTimeHelper::IsCommCtrl6() ? ILC_COLOR32 : ILC_COLOR24) | ILC_MASK, 1, 0);
	ImageList_SetBkColor(hDragImageList, CLR_NONE);
	ImageList_Add(hDragImageList, dragImage, dragImageMask);

	ReleaseDC(hCompatibleDC);

	return hDragImageList;
}

BOOL ToolBar::CreateLegacyOLEDragImage(IToolBarButtonContainer* pButtons, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pButtons);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	// use a normal legacy drag image as base
	OLE_HANDLE h = NULL;
	OLE_XPOS_PIXELS xUpperLeft = 0;
	OLE_YPOS_PIXELS yUpperLeft = 0;
	pButtons->CreateDragImage(&xUpperLeft, &yUpperLeft, &h);
	if(h) {
		HIMAGELIST hImageList = static_cast<HIMAGELIST>(LongToHandle(h));

		// retrieve the drag image's size
		int bitmapHeight;
		int bitmapWidth;
		ImageList_GetIconSize(hImageList, &bitmapWidth, &bitmapHeight);
		pDragImage->sizeDragImage.cx = bitmapWidth;
		pDragImage->sizeDragImage.cy = bitmapHeight;

		CDC memoryDC;
		memoryDC.CreateCompatibleDC();
		pDragImage->hbmpDragImage = NULL;

		if(RunTimeHelper::IsCommCtrl6()) {
			// handle alpha channel
			IImageList* pImgLst = NULL;
			HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
			if(hMod) {
				typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
				HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
				if(pfnHIMAGELIST_QueryInterface) {
					pfnHIMAGELIST_QueryInterface(hImageList, IID_IImageList, reinterpret_cast<LPVOID*>(&pImgLst));
				}
				FreeLibrary(hMod);
			}
			if(!pImgLst) {
				pImgLst = reinterpret_cast<IImageList*>(hImageList);
				pImgLst->AddRef();
			}
			ATLASSUME(pImgLst);

			DWORD flags = 0;
			pImgLst->GetItemFlags(0, &flags);
			if(flags & ILIF_ALPHA) {
				// the drag image makes use of the alpha channel
				IMAGEINFO imageInfo = {0};
				ImageList_GetImageInfo(hImageList, 0, &imageInfo);

				// fetch raw data
				BITMAPINFO bitmapInfo = {0};
				bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
				bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
				bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
				bitmapInfo.bmiHeader.biPlanes = 1;
				bitmapInfo.bmiHeader.biBitCount = 32;
				bitmapInfo.bmiHeader.biCompression = BI_RGB;
				LPRGBQUAD pSourceBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * sizeof(RGBQUAD)));
				GetDIBits(memoryDC, imageInfo.hbmImage, 0, pDragImage->sizeDragImage.cy, pSourceBits, &bitmapInfo, DIB_RGB_COLORS);
				// create target bitmap
				LPRGBQUAD pDragImageBits = NULL;
				pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);
				ATLASSERT(pDragImageBits);
				pDragImage->crColorKey = 0xFFFFFFFF;

				if(pDragImageBits) {
					// transfer raw data
					CopyMemory(pDragImageBits, pSourceBits, pDragImage->sizeDragImage.cx * pDragImage->sizeDragImage.cy * 4);
				}

				// clean up
				HeapFree(GetProcessHeap(), 0, pSourceBits);
				DeleteObject(imageInfo.hbmImage);
				DeleteObject(imageInfo.hbmMask);
			}
			pImgLst->Release();
		}

		if(!pDragImage->hbmpDragImage) {
			// fallback mode
			memoryDC.SetBkMode(TRANSPARENT);

			// create target bitmap
			HDC hCompatibleDC = ::GetDC(HWND_DESKTOP);
			pDragImage->hbmpDragImage = CreateCompatibleBitmap(hCompatibleDC, bitmapWidth, bitmapHeight);
			::ReleaseDC(HWND_DESKTOP, hCompatibleDC);
			HBITMAP hPreviousBitmap = memoryDC.SelectBitmap(pDragImage->hbmpDragImage);

			// draw target bitmap
			pDragImage->crColorKey = RGB(0xF4, 0x00, 0x00);
			CBrush backroundBrush;
			backroundBrush.CreateSolidBrush(pDragImage->crColorKey);
			memoryDC.FillRect(CRect(0, 0, bitmapWidth, bitmapHeight), backroundBrush);
			ImageList_Draw(hImageList, 0, memoryDC, 0, 0, ILD_NORMAL);

			// clean up
			memoryDC.SelectBitmap(hPreviousBitmap);
		}

		ImageList_Destroy(hImageList);

		if(pDragImage->hbmpDragImage) {
			// retrieve the offset
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);
			if(GetExStyle() & WS_EX_LAYOUTRTL) {
				pDragImage->ptOffset.x = xUpperLeft + pDragImage->sizeDragImage.cx - mousePosition.x;
			} else {
				pDragImage->ptOffset.x = mousePosition.x - xUpperLeft;
			}
			pDragImage->ptOffset.y = mousePosition.y - yUpperLeft;

			succeeded = TRUE;
		}
	}

	return succeeded;
}

BOOL ToolBar::CreateVistaOLEDragImage(IToolBarButtonContainer* pButtons, LPSHDRAGIMAGE pDragImage)
{
	ATLASSUME(pButtons);
	ATLASSERT_POINTER(pDragImage, SHDRAGIMAGE);

	BOOL succeeded = FALSE;

	CTheme themingEngine;
	themingEngine.OpenThemeData(NULL, VSCLASS_DRAGDROP);
	if(themingEngine.IsThemeNull()) {
		// FIXME: What should we do here?
		ATLASSERT(FALSE && "Current theme does not define the \"DragDrop\" class.");
	} else {
		// retrieve the drag image's size
		CDC memoryDC;
		memoryDC.CreateCompatibleDC();

		themingEngine.GetThemePartSize(memoryDC, DD_IMAGEBG, 1, NULL, TS_TRUE, &pDragImage->sizeDragImage);
		MARGINS margins = {0};
		themingEngine.GetThemeMargins(memoryDC, DD_IMAGEBG, 1, TMT_CONTENTMARGINS, NULL, &margins);
		pDragImage->sizeDragImage.cx -= margins.cxLeftWidth + margins.cxRightWidth;
		pDragImage->sizeDragImage.cy -= margins.cyTopHeight + margins.cyBottomHeight;
	}

	ATLASSERT(pDragImage->sizeDragImage.cx > 0);
	ATLASSERT(pDragImage->sizeDragImage.cy > 0);

	// collect image lists
	BOOL hasHighResImages = FALSE;
	BOOL hasAnyImageList = FALSE;
	#ifdef USE_STL
		std::vector<HIMAGELIST> hSourceImageLists;
	#else
		CAtlArray<HIMAGELIST> hSourceImageLists;
	#endif
	CComPtr<IUnknown> pUnknownEnum = NULL;
	pButtons->get__NewEnum(&pUnknownEnum);
	CComQIPtr<IEnumVARIANT> pEnum = pUnknownEnum;
	ATLASSUME(pEnum);
	if(pEnum) {
		VARIANT v;
		int i = 0;
		CComPtr<IToolBarButton> pButton = NULL;
		while(pEnum->Next(1, &v, NULL) == S_OK) {
			if(v.vt == VT_DISPATCH) {
				v.pdispVal->QueryInterface(IID_IToolBarButton, reinterpret_cast<LPVOID*>(&pButton));
				ATLASSUME(pButton);
				if(pButton) {
					LONG imageListIndex = 0;
					pButton->get_ImageListIndex(&imageListIndex);
					#ifdef USE_STL
						std::unordered_map<LONG, HIMAGELIST>::iterator iter = properties.hHighResImageLists.find(imageListIndex);
						if(iter != properties.hHighResImageLists.end()) {
							hSourceImageLists.push_back(iter->second);
							hasAnyImageList = hasAnyImageList || (iter->second != NULL);
							hasHighResImages = TRUE;
						} else {
							HIMAGELIST hImgLst = reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETIMAGELIST, imageListIndex, 0));
							hSourceImageLists.push_back(hImgLst);
							hasAnyImageList = hasAnyImageList || (hImgLst != NULL);
						}
					#else
						CAtlMap<LONG, HIMAGELIST>::CPair* pEntry = properties.hHighResImageLists.Lookup(imageListIndex);
						if(pEntry) {
							hSourceImageLists.Add(pEntry->m_value);
							hasAnyImageList = hasAnyImageList || (pEntry->m_value != NULL);
							hasHighResImages = TRUE;
						} else {
							HIMAGELIST hImgLst = reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETIMAGELIST, imageListIndex, 0));
							hSourceImageLists.Add(hImgLst);
							hasAnyImageList = hasAnyImageList || (hImgLst != NULL);
						}
					#endif
					++i;
				}
				pButton = NULL;
			}
			VariantClear(&v);
			if(i >= (hasHighResImages ? 5 : 10)) {
				break;
			}
		}
	}

	// create target bitmap
	BITMAPINFO bitmapInfo = {0};
	bitmapInfo.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
	bitmapInfo.bmiHeader.biWidth = pDragImage->sizeDragImage.cx;
	bitmapInfo.bmiHeader.biHeight = -pDragImage->sizeDragImage.cy;
	bitmapInfo.bmiHeader.biPlanes = 1;
	bitmapInfo.bmiHeader.biBitCount = 32;
	bitmapInfo.bmiHeader.biCompression = BI_RGB;
	LPRGBQUAD pDragImageBits = NULL;
	pDragImage->hbmpDragImage = CreateDIBSection(NULL, &bitmapInfo, DIB_RGB_COLORS, reinterpret_cast<LPVOID*>(&pDragImageBits), NULL, 0);

	if(!hasAnyImageList) {
		// report success, although we've an empty drag image
		return TRUE;
	}

	LONG numberOfButtons = 0;
	pButtons->Count(&numberOfButtons);
	ATLASSERT(numberOfButtons > 0);
	// don't display more than 5 (10) thumbnails
	numberOfButtons = min(numberOfButtons, (hasHighResImages ? 5 : 10));
	SIZE thumbnailSize;
	thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfButtons - 1);
	if(thumbnailSize.cy < 8) {
		// don't get smaller than 8x8 thumbnails
		numberOfButtons = (pDragImage->sizeDragImage.cy - 8) / 3 + 1;
		thumbnailSize.cy = pDragImage->sizeDragImage.cy - 3 * (numberOfButtons - 1);
	}
	thumbnailSize.cx = thumbnailSize.cy;
	int thumbnailBufferSize = thumbnailSize.cx * thumbnailSize.cy * sizeof(RGBQUAD);
	LPRGBQUAD pThumbnailBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, thumbnailBufferSize));
	ATLASSERT(pThumbnailBits);
	if(pThumbnailBits) {
		// iterate over the dragged buttons
		VARIANT v;
		int i = 0;
		CComPtr<IToolBarButton> pButton = NULL;
		while(pEnum->Next(1, &v, NULL) == S_OK) {
			if(v.vt == VT_DISPATCH) {
				v.pdispVal->QueryInterface(IID_IToolBarButton, reinterpret_cast<LPVOID*>(&pButton));
				ATLASSUME(pButton);
				if(pButton) {
					HIMAGELIST hSourceImageList = hSourceImageLists[i];
					IImageList2* pImgLst = NULL;
					HMODULE hMod = LoadLibrary(TEXT("comctl32.dll"));
					if(hMod) {
						typedef HRESULT WINAPI HIMAGELIST_QueryInterfaceFn(HIMAGELIST, REFIID, LPVOID*);
						HIMAGELIST_QueryInterfaceFn* pfnHIMAGELIST_QueryInterface = reinterpret_cast<HIMAGELIST_QueryInterfaceFn*>(GetProcAddress(hMod, "HIMAGELIST_QueryInterface"));
						if(pfnHIMAGELIST_QueryInterface) {
							pfnHIMAGELIST_QueryInterface(hSourceImageList, IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
						}
						FreeLibrary(hMod);
					}
					if(!pImgLst) {
						IImageList* p = reinterpret_cast<IImageList*>(hSourceImageList);
						p->QueryInterface(IID_IImageList2, reinterpret_cast<LPVOID*>(&pImgLst));
					}
					ATLASSUME(pImgLst);

					// get the button's icon
					LONG icon = 0;
					pButton->get_IconIndex(&icon);

					if(pImgLst) {
						pImgLst->ForceImagePresent(icon, ILFIP_ALWAYS);
						HICON hIcon = NULL;
						pImgLst->GetIcon(icon, ILD_TRANSPARENT, &hIcon);
						ATLASSERT(hIcon);
						if(hIcon) {
							// finally create the thumbnail
							ZeroMemory(pThumbnailBits, thumbnailBufferSize);
							HRESULT hr = CreateThumbnail(hIcon, thumbnailSize, pThumbnailBits, TRUE);
							DestroyIcon(hIcon);
							if(FAILED(hr)) {
								pButton = NULL;
								VariantClear(&v);
								break;
							}

							// add the thumbail to the drag image keeping the alpha channel intact
							if(i == 0) {
								LPRGBQUAD pDragImagePixel = pDragImageBits;
								LPRGBQUAD pThumbnailPixel = pThumbnailBits;
								for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx, pThumbnailPixel += thumbnailSize.cx) {
									CopyMemory(pDragImagePixel, pThumbnailPixel, thumbnailSize.cx * sizeof(RGBQUAD));
								}
							} else {
								LPRGBQUAD pDragImagePixel = pDragImageBits;
								LPRGBQUAD pThumbnailPixel = pThumbnailBits;
								pDragImagePixel += 3 * i * pDragImage->sizeDragImage.cx;
								for(int scanline = 0; scanline < thumbnailSize.cy; ++scanline, pDragImagePixel += pDragImage->sizeDragImage.cx) {
									LPRGBQUAD p = pDragImagePixel + 2 * i;
									for(int x = 0; x < thumbnailSize.cx; ++x, ++p, ++pThumbnailPixel) {
										// merge the pixels
										p->rgbRed = pThumbnailPixel->rgbRed * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbRed / 0xFF;
										p->rgbGreen = pThumbnailPixel->rgbGreen * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbGreen / 0xFF;
										p->rgbBlue = pThumbnailPixel->rgbBlue * pThumbnailPixel->rgbReserved / 0xFF + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbBlue / 0xFF;
										p->rgbReserved = pThumbnailPixel->rgbReserved + (0xFF - pThumbnailPixel->rgbReserved) * p->rgbReserved / 0xFF;
									}
								}
							}
						}
						pImgLst->Release();
					}

					++i;
					pButton = NULL;
					if(i == numberOfButtons) {
						VariantClear(&v);
						break;
					}
				}
			}
			VariantClear(&v);
		}
		HeapFree(GetProcessHeap(), 0, pThumbnailBits);
		succeeded = TRUE;
	}
	return succeeded;
}

//////////////////////////////////////////////////////////////////////
// implementation of IDropTarget
STDMETHODIMP ToolBar::DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	if(properties.supportOLEDragImages && !dragDropStatus.pDropTargetHelper) {
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&dragDropStatus.pDropTargetHelper));
	}

	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragEnter(pDataObject, pEffect, keyState, mousePosition);

	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
		if(dragDropStatus.useItemCountLabelHack) {
			dragDropStatus.pDropTargetHelper->DragLeave();
			dragDropStatus.pDropTargetHelper->DragEnter(*this, pDataObject, &buffer, *pEffect);
			dragDropStatus.useItemCountLabelHack = FALSE;
		}
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::DragLeave(void)
{
	Raise_OLEDragLeave();
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragLeave();
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	return S_OK;
}

STDMETHODIMP ToolBar::DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	CComQIPtr<IDataObject> pDataObject = dragDropStatus.pActiveDataObject;
	DROPDESCRIPTION oldDropDescription;
	ZeroMemory(&oldDropDescription, sizeof(DROPDESCRIPTION));
	IDataObject_GetDropDescription(pDataObject, oldDropDescription);

	POINT buffer = {mousePosition.x, mousePosition.y};
	Raise_OLEDragMouseMove(pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->DragOver(&buffer, *pEffect);
	}

	DROPDESCRIPTION newDropDescription;
	ZeroMemory(&newDropDescription, sizeof(DROPDESCRIPTION));
	if(SUCCEEDED(IDataObject_GetDropDescription(pDataObject, newDropDescription)) && (newDropDescription.type > DROPIMAGE_NONE || memcmp(&oldDropDescription, &newDropDescription, sizeof(DROPDESCRIPTION)))) {
		InvalidateDragWindow(pDataObject);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect)
{
	// NOTE: pDataObject can be NULL

	POINT buffer = {mousePosition.x, mousePosition.y};
	dragDropStatus.drop_pDataObject = pDataObject;
	dragDropStatus.drop_mousePosition = buffer;
	dragDropStatus.drop_effect = *pEffect;

	Raise_OLEDragDrop(pDataObject, pEffect, keyState, mousePosition);
	if(dragDropStatus.pDropTargetHelper) {
		dragDropStatus.pDropTargetHelper->Drop(pDataObject, &buffer, *pEffect);
		dragDropStatus.pDropTargetHelper->Release();
		dragDropStatus.pDropTargetHelper = NULL;
	}
	dragDropStatus.drop_pDataObject = NULL;
	return S_OK;
}
// implementation of IDropTarget
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSource
STDMETHODIMP ToolBar::GiveFeedback(DWORD effect)
{
	VARIANT_BOOL useDefaultCursors = VARIANT_TRUE;
	//if(flags.usingThemes && RunTimeHelper::IsVista()) {
		ATLASSUME(dragDropStatus.pSourceDataObject);

		BOOL isShowingLayered = FALSE;
		FORMATETC format = {0};
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("IsShowingLayered")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		STGMEDIUM medium = {0};
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				isShowingLayered = *static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}
		BOOL useDropDescriptionHack = FALSE;
		format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
		format.dwAspect = DVASPECT_CONTENT;
		format.lindex = -1;
		format.tymed = TYMED_HGLOBAL;
		if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
			if(medium.hGlobal) {
				useDropDescriptionHack = *static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				GlobalUnlock(medium.hGlobal);
			}
			ReleaseStgMedium(&medium);
		}

		if(isShowingLayered && properties.oleDragImageStyle != odistClassic) {
			SetCursor(static_cast<HCURSOR>(LoadImage(NULL, MAKEINTRESOURCE(OCR_NORMAL), IMAGE_CURSOR, 0, 0, LR_DEFAULTCOLOR | LR_DEFAULTSIZE | LR_SHARED)));
			useDefaultCursors = VARIANT_FALSE;
		}
		if(useDropDescriptionHack) {
			// this will make drop descriptions work
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("DragWindow")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			if(SUCCEEDED(dragDropStatus.pSourceDataObject->GetData(&format, &medium))) {
				if(medium.hGlobal) {
					// WM_USER + 1 (with wParam = 0 and lParam = 0) hides the drag image
					#define WM_SETDROPEFFECT				WM_USER + 2     // (wParam = DCID_*, lParam = 0)
					#define DDWM_UPDATEWINDOW				WM_USER + 3     // (wParam = 0, lParam = 0)
					typedef enum DROPEFFECTS
					{
						DCID_NULL = 0,
						DCID_NO = 1,
						DCID_MOVE = 2,
						DCID_COPY = 3,
						DCID_LINK = 4,
						DCID_MAX = 5
					} DROPEFFECTS;

					HWND hWndDragWindow = *static_cast<HWND*>(GlobalLock(medium.hGlobal));
					GlobalUnlock(medium.hGlobal);

					DROPEFFECTS dropEffect = DCID_NULL;
					switch(effect) {
						case DROPEFFECT_NONE:
							dropEffect = DCID_NO;
							break;
						case DROPEFFECT_COPY:
							dropEffect = DCID_COPY;
							break;
						case DROPEFFECT_MOVE:
							dropEffect = DCID_MOVE;
							break;
						case DROPEFFECT_LINK:
							dropEffect = DCID_LINK;
							break;
					}
					if(::IsWindow(hWndDragWindow)) {
						::PostMessage(hWndDragWindow, WM_SETDROPEFFECT, dropEffect, 0);
					}
				}
				ReleaseStgMedium(&medium);
			}
		}
	//}

	Raise_OLEGiveFeedback(effect, &useDefaultCursors);
	return (useDefaultCursors == VARIANT_FALSE ? S_OK : DRAGDROP_S_USEDEFAULTCURSORS);
}

STDMETHODIMP ToolBar::QueryContinueDrag(BOOL pressedEscape, DWORD keyState)
{
	HRESULT actionToContinueWith = S_OK;
	if(pressedEscape) {
		actionToContinueWith = DRAGDROP_S_CANCEL;
	} else if(!(keyState & dragDropStatus.draggingMouseButton)) {
		actionToContinueWith = DRAGDROP_S_DROP;
	}
	Raise_OLEQueryContinueDrag(pressedEscape, keyState, &actionToContinueWith);
	return actionToContinueWith;
}
// implementation of IDropSource
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IDropSourceNotify
STDMETHODIMP ToolBar::DragEnterTarget(HWND hWndTarget)
{
	Raise_OLEDragEnterPotentialTarget(HandleToLong(hWndTarget));
	return S_OK;
}

STDMETHODIMP ToolBar::DragLeaveTarget(void)
{
	Raise_OLEDragLeavePotentialTarget();
	return S_OK;
}
// implementation of IDropSourceNotify
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of ICategorizeProperties
STDMETHODIMP ToolBar::GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName)
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

STDMETHODIMP ToolBar::MapPropertyToCategory(DISPID property, PROPCAT* pCategory)
{
	if(!pCategory) {
		return E_POINTER;
	}

	switch(property) {
		case DISPID_TB_ANCHORHIGHLIGHTING:
		case DISPID_TB_APPEARANCE:
		case DISPID_TB_BACKSTYLE:
		case DISPID_TB_BORDERSTYLE:
		case DISPID_TB_BUTTONHEIGHT:
		case DISPID_TB_BUTTONROWCOUNT:
		case DISPID_TB_BUTTONSTYLE:
		case DISPID_TB_BUTTONTEXTPOSITION:
		case DISPID_TB_BUTTONWIDTH:
		case DISPID_TB_DISPLAYMENUDIVIDER:
		case DISPID_TB_DROPDOWNGAP:
		case DISPID_TB_FIRSTBUTTONINDENTATION:
		case DISPID_TB_HORIZONTALBUTTONPADDING:
		case DISPID_TB_HORIZONTALBUTTONSPACING:
		case DISPID_TB_HORIZONTALICONCAPTIONGAP:
		case DISPID_TB_HORIZONTALTEXTALIGNMENT:
		case DISPID_TB_IDEALHEIGHT:
		case DISPID_TB_IDEALWIDTH:
		case DISPID_TB_MAXIMUMBUTTONWIDTH:
		case DISPID_TB_MAXIMUMHEIGHT:
		case DISPID_TB_MAXIMUMTEXTROWS:
		case DISPID_TB_MAXIMUMWIDTH:
		case DISPID_TB_MENUBARTHEME:
		case DISPID_TB_MENUMODE:
		case DISPID_TB_MINIMUMBUTTONWIDTH:
		case DISPID_TB_MOUSEICON:
		case DISPID_TB_MOUSEPOINTER:
		case DISPID_TB_MULTICOLUMN:
		case DISPID_TB_NORMALDROPDOWNBUTTONSTYLE:
		case DISPID_TB_ORIENTATION:
		case DISPID_TB_SUGGESTEDICONSIZE:
		case DISPID_TB_USEMNEMONICS:
		case DISPID_TB_VERTICALBUTTONPADDING:
		case DISPID_TB_VERTICALBUTTONSPACING:
		case DISPID_TB_VERTICALTEXTALIGNMENT:
			*pCategory = PROPCAT_Appearance;
			return S_OK;
			break;
		case DISPID_TB_ALLOWCUSTOMIZATION:
		case DISPID_TB_ALWAYSDISPLAYBUTTONTEXT:
		case DISPID_TB_DISABLEDEVENTS:
		case DISPID_TB_DISPLAYPARTIALLYCLIPPEDBUTTONS:
		case DISPID_TB_DONTREDRAW:
		case DISPID_TB_DRAGDROPCUSTOMIZATIONMODIFIERKEY:
		case DISPID_TB_FOCUSONCLICK:
		case DISPID_TB_HOVERTIME:
		case DISPID_TB_PROCESSCONTEXTMENUKEYS:
		case DISPID_TB_RAISECUSTOMDRAWEVENTONERASEBACKGROUND:
		case DISPID_TB_RIGHTTOLEFT:
		case DISPID_TB_SHOWTOOLTIPS:
		case DISPID_TB_WRAPBUTTONS:
			*pCategory = PROPCAT_Behavior;
			return S_OK;
			break;
		case DISPID_TB_HIGHLIGHTCOLOR:
		case DISPID_TB_INSERTMARKCOLOR:
		case DISPID_TB_SHADOWCOLOR:
			*pCategory = PROPCAT_Colors;
			return S_OK;
			break;
		case DISPID_TB_APPID:
		case DISPID_TB_APPNAME:
		case DISPID_TB_APPSHORTNAME:
		case DISPID_TB_BUILD:
		case DISPID_TB_CHARSET:
		case DISPID_TB_HCHEVRONMENU:
		case DISPID_TB_HIMAGELIST:
		case DISPID_TB_HWND:
		case DISPID_TB_HWNDCHEVRONTOOLBAR:
		case DISPID_TB_HWNDTOOLTIP:
		case DISPID_TB_IMAGELISTCOUNT:
		case DISPID_TB_ISRELEASE:
		case DISPID_TB_PROGRAMMER:
		case DISPID_TB_TESTER:
		case DISPID_TB_VERSION:
			*pCategory = PROPCAT_Data;
			return S_OK;
			break;
		case DISPID_TB_DRAGCLICKTIME:
		case DISPID_TB_DRAGGEDBUTTONS:
		case DISPID_TB_NATIVEDROPTARGET:
		case DISPID_TB_OLEDRAGIMAGESTYLE:
		case DISPID_TB_REGISTERFOROLEDRAGDROP:
		case DISPID_TB_SHOWDRAGIMAGE:
		case DISPID_TB_SUPPORTOLEDRAGIMAGES:
			*pCategory = PROPCAT_DragDrop;
			return S_OK;
			break;
		case DISPID_TB_FONT:
		case DISPID_TB_USESYSTEMFONT:
			*pCategory = PROPCAT_Font;
			return S_OK;
			break;
		case DISPID_TB_BUTTONS:
		case DISPID_TB_HOTBUTTON:
			*pCategory = PROPCAT_List;
			return S_OK;
			break;
		case DISPID_TB_COMMANDENABLED:
		case DISPID_TB_ENABLED:
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
CAtlString ToolBar::GetAuthors(void)
{
	CComBSTR buffer;
	get_Programmer(&buffer);
	return CAtlString(buffer);
}

CAtlString ToolBar::GetHomepage(void)
{
	return TEXT("https://www.TimoSoft-Software.de");
}

CAtlString ToolBar::GetPaypalLink(void)
{
	return TEXT("https://www.paypal.com/xclick/business=TKunze71216%40gmx.de&item_name=ToolBarControls&no_shipping=1&tax=0&currency_code=EUR");
}

CAtlString ToolBar::GetSpecialThanks(void)
{
	return TEXT("Geoff Chappell, Wine Headquarters");
}

CAtlString ToolBar::GetThanks(void)
{
	CAtlString ret = TEXT("Google, various newsgroups and mailing lists, many websites,\n");
	ret += TEXT("Heaven Shall Burn, Arch Enemy, Machine Head, Trivium, Deadlock, Draconian, Soulfly, Delain, Lacuna Coil, Ensiferum, Epica, Nightwish, Guns N' Roses and many other musicians");
	return ret;
}

CAtlString ToolBar::GetVersion(void)
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
// implementation of IKeyboardHookHandler
LRESULT ToolBar::HandleKeyboardMessage(int /*code*/, WPARAM wParam, LPARAM lParam, BOOL& consumeMessage)
{
	LRESULT lr = FALSE;

	HWND hWnd = *this;
	while(hWnd) {
		if(!::IsWindowEnabled(hWnd) || !::IsWindowVisible(hWnd)) {
			return TRUE;
		}
		hWnd = ::GetParent(hWnd);
	}

	if(!(HIWORD(lParam) & KF_UP) && properties.pAccelerators && properties.numberOfAccelerators > 0 && !flags.executingHotKeyCommand) {
		BYTE modifiers = FVIRTKEY;
		if(GetAsyncKeyState(VK_SHIFT) & 0x8000) {
			modifiers |= FSHIFT;
		}
		if(GetAsyncKeyState(VK_CONTROL) & 0x8000) {
			modifiers |= FCONTROL;
		}
		if(HIWORD(lParam) & KF_ALTDOWN) {
			modifiers |= FALT;
		}
		for(UINT i = 0; i < properties.numberOfAccelerators; i++) {
			if(properties.pAccelerators[i].fVirt == modifiers && properties.pAccelerators[i].key == static_cast<WORD>(wParam)) {
				flags.executingHotKeyCommand = TRUE;
				BOOL commandIsDisabled = FALSE;
				VARIANT_BOOL forward = VARIANT_TRUE;
				Raise_ExecuteCommand(properties.pAccelerators[i].cmd, NULL, coHotkey, &forward, &commandIsDisabled);
				if(forward != VARIANT_FALSE && IsWindow()) {
					CWindow parent = GetTopLevelParent();
					if(parent.IsWindow()) {
						parent.SendMessage(WM_COMMAND, MAKEWPARAM(properties.pAccelerators[i].cmd, 1), 0);
					}
				}
				flags.executingHotKeyCommand = FALSE;
				consumeMessage = !commandIsDisabled;
				return TRUE;
			}
		}
	}

	if(HIWORD(lParam) & KF_ALTDOWN) {
		if(!(HIWORD(lParam) & KF_UP) && !flags.recursiveHookCall) {
			// ALT key is being pressed
			CWindow parentWindow = GetTopLevelParent();
			if(parentWindow.IsWindow()) {
				// we do this to close any chevron popup windows
				parentWindow.SendMessage(WM_CANCELMODE, 0, 0);
			}
			ShowKeyboardCues(TRUE);
			// TODO: This won't work for things like ALT+SHIFT+?. Idea: Use TranslateMessage and do the whole stuff in OnChar.
			USHORT acceleratorChar = LOWORD(MapVirtualKey(wParam, MAPVK_VK_TO_CHAR));
			if(acceleratorChar > 0) {
				UINT buttonID = static_cast<UINT>(-1);
				if(SendMessage(TB_MAPACCELERATOR, acceleratorChar, reinterpret_cast<LPARAM>(&buttonID))) {
					// we found a matching button
					/* NOTE: Comctl32 has a strange feature. If a button does not include a mnemonic, it uses the
					         first character of this button as accelerator. So do a sanity-check whether the returned
					         button really has this acceleratorChar. */
					// NOTE: We could do this always on TB_MAPACCELERATOR, but this probably would break menu mode.
					BOOL realHit = FALSE;
					int textLength = static_cast<int>(SendMessage(TB_GETBUTTONTEXT, buttonID, NULL));
					if(textLength > -1) {
						LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
						if(pBuffer) {
							ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
							if(static_cast<int>(SendMessage(TB_GETBUTTONTEXT, buttonID, reinterpret_cast<LPARAM>(pBuffer))) > -1) {
								for(int i = lstrlen(pBuffer) - 1; i > 0; --i) {
									if((pBuffer[i - 1] == TEXT('&')) && !ChrCmpI(pBuffer[i], acceleratorChar)) {
										realHit = TRUE;
										break;
									}
								}
							}
							HeapFree(GetProcessHeap(), 0, pBuffer);
						}
					}

					if(realHit) {
						TBBUTTONINFO button = {0};
						button.cbSize = sizeof(button);
						button.dwMask = TBIF_STATE | TBIF_STYLE;
						int buttonIndex = SendMessage(TB_GETBUTTONINFO, buttonID, reinterpret_cast<LPARAM>(&button));
						if(buttonIndex >= -1) {
							if((button.fsState & TBSTATE_ENABLED) == TBSTATE_ENABLED && !(button.fsStyle & BTNS_SEP)) {
								consumeMessage = TRUE;
								lr = TRUE;

								BOOL isDropDownButton = (button.fsStyle & BTNS_WHOLEDROPDOWN) == BTNS_WHOLEDROPDOWN;
								isDropDownButton = isDropDownButton || ((button.fsStyle & BTNS_DROPDOWN) == BTNS_DROPDOWN && !(SendMessage(TB_GETEXTENDEDSTYLE, 0, 0) & TBSTYLE_EX_DRAWDDARROWS));
								if(isDropDownButton) {
									BOOL needsChevron = FALSE;
									RECT availableRectangle = {0};
									GetClientRect(&availableRectangle);
									RECT buttonRectangle = {0};
									if(SendMessage(TB_GETITEMRECT, buttonIndex, reinterpret_cast<LPARAM>(&buttonRectangle))) {
										if(buttonRectangle.right > availableRectangle.right) {
											needsChevron = TRUE;
										}
									}
									if(needsChevron) {
										// assume we are in a rebar
										HWND hWndReBar = GetParent();
										REBARBANDINFO bandInfo = {RunTimeHelper::SizeOf_REBARBANDINFO(), RBBIM_CHILD | RBBIM_STYLE};
										int bandCount = ::SendMessage(hWndReBar, RB_GETBANDCOUNT, 0, 0);
										for(int i = 0; i < bandCount; i++) {
											if(::SendMessage(hWndReBar, RB_GETBANDINFO, i, reinterpret_cast<LPARAM>(&bandInfo)) && bandInfo.hwndChild == *this) {
												if(bandInfo.fStyle & RBBS_USECHEVRON) {
													::PostMessage(hWndReBar, RB_PUSHCHEVRON, i, 0);
													properties.acceleratorToSendToChevronMenu = acceleratorChar;
													SetTimer(timers.ID_PRESELECTCHEVRONMENUITEM, timers.INT_PRESELECTCHEVRONMENUITEM);
												}
												break;
											}
										}
									} else {
										NMTOOLBAR data = {0};
										data.hdr.code = TBN_DROPDOWN;
										data.hdr.hwndFrom = *this;
										data.hdr.idFrom = GetDlgCtrlID();
										data.iItem = buttonID;
										SendMessage(TB_GETITEMRECT, buttonIndex, reinterpret_cast<LPARAM>(&data.rcButton));
										/* Usually the client app will display a drop-down menu. Further accelerator keys should go
											 to this menu, so guard against recursive calls. */
										flags.recursiveHookCall = TRUE;
										GetParent().SendMessage(WM_NOTIFY, reinterpret_cast<WPARAM>(static_cast<HWND>(*this)), reinterpret_cast<LPARAM>(&data));
										flags.recursiveHookCall = FALSE;
									}
								} else {
									GetParent().PostMessage(WM_COMMAND, MAKEWPARAM(buttonID, BN_CLICKED), reinterpret_cast<LPARAM>(static_cast<HWND>(*this)));
								}
							}
						}
					}
				}
			}
		}
	}
	return lr;
}
// implementation of IKeyboardHookHandler
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IMouseHookHandler
LRESULT ToolBar::HandleMouseMessage(int /*code*/, WPARAM wParam, LPARAM lParam, BOOL& consumeMessage)
{
	BOOL dismiss = FALSE;
	if(wParam == WM_LBUTTONDOWN || wParam == WM_RBUTTONDOWN || wParam == WM_MBUTTONDOWN || wParam == WM_NCLBUTTONDOWN || wParam == WM_NCRBUTTONDOWN || wParam == WM_NCMBUTTONDOWN) {
		LPMOUSEHOOKSTRUCT pDetails = reinterpret_cast<LPMOUSEHOOKSTRUCT>(lParam);
		ATLASSERT_POINTER(pDetails, MOUSEHOOKSTRUCT);

		HWND hWnd = WindowFromPoint(pDetails->pt);
		dismiss = TRUE;
		if(hWnd == chevronPopupMenuWindow || ::GetParent(hWnd) == chevronPopupMenuWindow) {
			dismiss = FALSE;
		} else {
			TCHAR pClassName[201];
			::GetClassName(hWnd, pClassName, 250);
			if(lstrcmpi(pClassName, TEXT("#32768")) == 0) {
				dismiss = FALSE;
			}
		}
	}

	if(dismiss) {
		CWindow topLevelParent = GetTopLevelParent();
		if(topLevelParent.IsWindow()) {
			topLevelParent.SendMessage(WM_CANCELMODE, 0, 0);
		}
		consumeMessage = TRUE;
		return TRUE;
	}

	return FALSE;
}
// implementation of IMouseHookHandler
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// implementation of IPerPropertyBrowsing
STDMETHODIMP ToolBar::GetDisplayString(DISPID property, BSTR* pDescription)
{
	if(!pDescription) {
		return E_POINTER;
	}

	CComBSTR description;
	HRESULT hr = S_OK;
	switch(property) {
		case DISPID_TB_APPEARANCE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.appearance), description);
			break;
		case DISPID_TB_BACKSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.backStyle), description);
			break;
		case DISPID_TB_BORDERSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.borderStyle), description);
			break;
		case DISPID_TB_BUTTONSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.buttonStyle), description);
			break;
		case DISPID_TB_BUTTONTEXTPOSITION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.buttonTextPosition), description);
			break;
		case DISPID_TB_DRAGDROPCUSTOMIZATIONMODIFIERKEY:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.dragDropCustomizationModifierKey), description);
			break;
		case DISPID_TB_HORIZONTALTEXTALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.horizontalTextAlignment), description);
			break;
		case DISPID_TB_MENUBARTHEME:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.menuBarTheme), description);
			break;
		case DISPID_TB_MOUSEPOINTER:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.mousePointer), description);
			break;
		case DISPID_TB_NORMALDROPDOWNBUTTONSTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.normalDropDownButtonStyle), description);
			break;
		case DISPID_TB_OLEDRAGIMAGESTYLE:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.oleDragImageStyle), description);
			break;
		case DISPID_TB_ORIENTATION:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.orientation), description);
			break;
		case DISPID_TB_REGISTERFOROLEDRAGDROP:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.registerForOLEDragDrop), description);
			break;
		case DISPID_TB_RIGHTTOLEFT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.rightToLeft), description);
			break;
		case DISPID_TB_VERTICALTEXTALIGNMENT:
			hr = GetDisplayStringForSetting(property, static_cast<DWORD>(properties.verticalTextAlignment), description);
			break;
		default:
			return IPerPropertyBrowsingImpl<ToolBar>::GetDisplayString(property, pDescription);
			break;
	}
	if(SUCCEEDED(hr)) {
		*pDescription = description.Detach();
	}

	return *pDescription ? S_OK : E_OUTOFMEMORY;
}

STDMETHODIMP ToolBar::GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies)
{
	if(!pDescriptions || !pCookies) {
		return E_POINTER;
	}

	int c = 0;
	switch(property) {
		case DISPID_TB_BACKSTYLE:
		case DISPID_TB_BORDERSTYLE:
		case DISPID_TB_BUTTONSTYLE:
		case DISPID_TB_BUTTONTEXTPOSITION:
		case DISPID_TB_DRAGDROPCUSTOMIZATIONMODIFIERKEY:
		case DISPID_TB_MENUBARTHEME:
		case DISPID_TB_NORMALDROPDOWNBUTTONSTYLE:
		case DISPID_TB_OLEDRAGIMAGESTYLE:
		case DISPID_TB_ORIENTATION:
			c = 2;
			break;
		case DISPID_TB_APPEARANCE:
		case DISPID_TB_HORIZONTALTEXTALIGNMENT:
		case DISPID_TB_REGISTERFOROLEDRAGDROP:
		case DISPID_TB_VERTICALTEXTALIGNMENT:
			c = 3;
			break;
		case DISPID_TB_RIGHTTOLEFT:
			c = 4;
			break;
		case DISPID_TB_MOUSEPOINTER:
			c = 30;
			break;
		default:
			return IPerPropertyBrowsingImpl<ToolBar>::GetPredefinedStrings(property, pDescriptions, pCookies);
			break;
	}
	pDescriptions->cElems = c;
	pCookies->cElems = c;
	pDescriptions->pElems = static_cast<LPOLESTR*>(CoTaskMemAlloc(pDescriptions->cElems * sizeof(LPOLESTR)));
	pCookies->pElems = static_cast<LPDWORD>(CoTaskMemAlloc(pCookies->cElems * sizeof(DWORD)));

	for(UINT iDescription = 0; iDescription < pDescriptions->cElems; ++iDescription) {
		UINT propertyValue = iDescription;
		if((property == DISPID_TB_MOUSEPOINTER) && (iDescription == pDescriptions->cElems - 1)) {
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

STDMETHODIMP ToolBar::GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue)
{
	switch(property) {
		case DISPID_TB_APPEARANCE:
		case DISPID_TB_BACKSTYLE:
		case DISPID_TB_BORDERSTYLE:
		case DISPID_TB_BUTTONSTYLE:
		case DISPID_TB_BUTTONTEXTPOSITION:
		case DISPID_TB_DRAGDROPCUSTOMIZATIONMODIFIERKEY:
		case DISPID_TB_HORIZONTALTEXTALIGNMENT:
		case DISPID_TB_MENUBARTHEME:
		case DISPID_TB_MOUSEPOINTER:
		case DISPID_TB_NORMALDROPDOWNBUTTONSTYLE:
		case DISPID_TB_OLEDRAGIMAGESTYLE:
		case DISPID_TB_ORIENTATION:
		case DISPID_TB_REGISTERFOROLEDRAGDROP:
		case DISPID_TB_RIGHTTOLEFT:
		case DISPID_TB_VERTICALTEXTALIGNMENT:
			VariantInit(pPropertyValue);
			pPropertyValue->vt = VT_I4;
			// we used the property value itself as cookie
			pPropertyValue->lVal = cookie;
			break;
		default:
			return IPerPropertyBrowsingImpl<ToolBar>::GetPredefinedValue(property, cookie, pPropertyValue);
			break;
	}
	return S_OK;
}

STDMETHODIMP ToolBar::MapPropertyToPage(DISPID property, CLSID* pPropertyPage)
{
	return IPerPropertyBrowsingImpl<ToolBar>::MapPropertyToPage(property, pPropertyPage);
}
// implementation of IPerPropertyBrowsing
//////////////////////////////////////////////////////////////////////

HRESULT ToolBar::GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description)
{
	switch(property) {
		case DISPID_TB_APPEARANCE:
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
		case DISPID_TB_BACKSTYLE:
			switch(cookie) {
				case bksTransparent:
					description = GetResStringWithNumber(IDP_BACKSTYLETRANSPARENT, bksTransparent);
					break;
				case bksOpaque:
					description = GetResStringWithNumber(IDP_BACKSTYLEOPAQUE, bksOpaque);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TB_BORDERSTYLE:
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
		case DISPID_TB_BUTTONSTYLE:
			switch(cookie) {
				case bst3D:
					description = GetResStringWithNumber(IDP_BUTTONSTYLE3D, bst3D);
					break;
				case bstFlat:
					description = GetResStringWithNumber(IDP_BUTTONSTYLEFLAT, bstFlat);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TB_BUTTONTEXTPOSITION:
			switch(cookie) {
				case btpBelowIcon:
					description = GetResStringWithNumber(IDP_BUTTONTEXTPOSITIONBELOW, btpBelowIcon);
					break;
				case btpRightToIcon:
					description = GetResStringWithNumber(IDP_BUTTONTEXTPOSITIONRIGHT, btpRightToIcon);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TB_DRAGDROPCUSTOMIZATIONMODIFIERKEY:
			switch(cookie) {
				case ddcmkShift:
					description = GetResStringWithNumber(IDP_DRAGDROPCUSTOMIZATIONMODIFIERKEYSHIFT, ddcmkShift);
					break;
				case ddcmkAlt:
					description = GetResStringWithNumber(IDP_DRAGDROPCUSTOMIZATIONMODIFIERKEYALT, ddcmkAlt);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TB_HORIZONTALTEXTALIGNMENT:
			switch(cookie) {
				case halLeft:
					description = GetResStringWithNumber(IDP_HORIZONTALTEXTALIGNMENTLEFT, halLeft);
					break;
				case halCenter:
					description = GetResStringWithNumber(IDP_HORIZONTALTEXTALIGNMENTCENTER, halCenter);
					break;
				case halRight:
					description = GetResStringWithNumber(IDP_HORIZONTALTEXTALIGNMENTRIGHT, halRight);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TB_MENUBARTHEME:
			switch(cookie) {
				case mbtNativeToolbar:
					description = GetResStringWithNumber(IDP_MENUBARTHEMENATIVETOOLBAR, mbtNativeToolbar);
					break;
				case mbtNativeMenuBar:
					description = GetResStringWithNumber(IDP_MENUBARTHEMENATIVEMENUBAR, mbtNativeMenuBar);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TB_MOUSEPOINTER:
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
		case DISPID_TB_NORMALDROPDOWNBUTTONSTYLE:
			switch(cookie) {
				case nddbsSplitButton:
					description = GetResStringWithNumber(IDP_NORMALDROPDOWNBUTTONSTYLESPLITBUTTON, nddbsSplitButton);
					break;
				case nddbsWithoutArrow:
					description = GetResStringWithNumber(IDP_NORMALDROPDOWNBUTTONSTYLEWITHOUTARROW, nddbsWithoutArrow);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TB_OLEDRAGIMAGESTYLE:
			switch(cookie) {
				case odistClassic:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLECLASSIC, odistClassic);
					break;
				case odistAeroIfAvailable:
					description = GetResStringWithNumber(IDP_OLEDRAGIMAGESTYLEAERO, odistAeroIfAvailable);
					break;
				default:
					return DISP_E_BADINDEX;
					break;
			}
			break;
		case DISPID_TB_ORIENTATION:
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
		case DISPID_TB_REGISTERFOROLEDRAGDROP:
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
		case DISPID_TB_RIGHTTOLEFT:
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
		case DISPID_TB_VERTICALTEXTALIGNMENT:
			switch(cookie) {
				case valTop:
					description = GetResStringWithNumber(IDP_VERTICALTEXTALIGNMENTTOP, valTop);
					break;
				case valCenter:
					description = GetResStringWithNumber(IDP_VERTICALTEXTALIGNMENTCENTER, valCenter);
					break;
				case valBottom:
					description = GetResStringWithNumber(IDP_VERTICALTEXTALIGNMENTBOTTOM, valBottom);
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
STDMETHODIMP ToolBar::GetPages(CAUUID* pPropertyPages)
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
STDMETHODIMP ToolBar::DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle)
{
	switch(verbID) {
		case 1:     // About...
			return DoVerbAbout(hWndParent);
			break;
		default:
			return IOleObjectImpl<ToolBar>::DoVerb(verbID, pMessage, pActiveSite, reserved, hWndParent, pBoundingRectangle);
			break;
	}
}

STDMETHODIMP ToolBar::EnumVerbs(IEnumOLEVERB** ppEnumerator)
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
STDMETHODIMP ToolBar::GetControlInfo(LPCONTROLINFO pControlInfo)
{
	ATLASSERT_POINTER(pControlInfo, CONTROLINFO);
	if(!pControlInfo) {
		return E_POINTER;
	}

	// our control can have an accelerator
	// see HandleKeyboardMessage
	pControlInfo->cb = sizeof(CONTROLINFO);
	pControlInfo->hAccel = NULL;//properties.hAcceleratorTable;
	pControlInfo->cAccel = 0;//static_cast<USHORT>(properties.hAcceleratorTable ? CopyAcceleratorTable(properties.hAcceleratorTable, NULL, 0) : 0);
	pControlInfo->dwFlags = 0;
	return S_OK;
}

/*STDMETHODIMP ToolBar::OnMnemonic(LPMSG pMessage)
{
	if(GetStyle() & WS_DISABLED) {
		return S_OK;
	}

	ATLASSERT(pMessage->message == WM_SYSKEYDOWN);

	SHORT pressedKeyCode = static_cast<SHORT>(pMessage->wParam);
	int buttonToSelect = -1;

	int numberOfButtons = SendMessage(TB_BUTTONCOUNT, 0, 0);
	TBBUTTONINFO button = {0};
	button.cbSize = sizeof(button);
	// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
	button.dwMask = TBIF_BYINDEX | TBIF_COMMAND | TBIF_STYLE | TBIF_STATE;
	for(int buttonIndex = 0; buttonIndex < numberOfButtons; ++buttonIndex) {
		if(SendMessage(TB_GETBUTTONINFO, buttonIndex, reinterpret_cast<LPARAM>(&button)) == buttonIndex) {
			if((button.fsStyle & (BTNS_SHOWTEX | BTNS_NOPREFIX | BTNS_SEP)) == BTNS_SHOWTEXT) {
				if(button.fsState & (TBSTATE_HIDDEN | TBSTATE_ENABLED) == TBSTATE_ENABLED) {
					int textLength = static_cast<int>(SendMessage(TB_GETBUTTONTEXT, button.idCommand, NULL));
					if(textLength > -1) {
						LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
						if(pBuffer) {
							ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
							if(static_cast<int>(SendMessage(TB_GETBUTTONTEXT, button.idCommand, reinterpret_cast<LPARAM>(pBuffer))) > -1) {
								for(int i = lstrlen(pBuffer); i > 0; --i) {
									if((pBuffer[i - 1] == TEXT('&')) && (pBuffer[i] != TEXT('&'))) {
										// TODO: Does this work with MFC?
										if((VkKeyScan(pBuffer[i]) == pressedKeyCode) || (VkKeyScan(static_cast<TCHAR>(tolower(pBuffer[i]))) == pressedKeyCode)) {
											buttonToSelect = buttonIndex;
											break;
										}
									}
								}
							}
							HeapFree(GetProcessHeap(), 0, pBuffer);
						}
					}
				}
			}
		}
	}
	if(buttonToSelect != -1) {
		// TODO
	}
	return S_OK;
}*/
// implementation of IOleControl
//////////////////////////////////////////////////////////////////////

HRESULT ToolBar::DoVerbAbout(HWND hWndParent)
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

HRESULT ToolBar::OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/)
{
	switch(propertyObject) {
		case DISPID_TB_FONT:
			if(!properties.useSystemFont) {
				ApplyFont();
			}
			break;
	}
	return S_OK;
}

HRESULT ToolBar::OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/)
{
	return S_OK;
}


STDMETHODIMP ToolBar::get_AllowCustomization(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.allowCustomization = ((GetStyle() & CCS_ADJUSTABLE) == CCS_ADJUSTABLE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.allowCustomization);
	return S_OK;
}

STDMETHODIMP ToolBar::put_AllowCustomization(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_ALLOWCUSTOMIZATION);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.allowCustomization != b) {
		properties.allowCustomization = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.allowCustomization) {
				ModifyStyle(0, CCS_ADJUSTABLE);
			} else {
				ModifyStyle(CCS_ADJUSTABLE, 0);
			}
		}
		FireOnChanged(DISPID_TB_ALLOWCUSTOMIZATION);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_AlwaysDisplayButtonText(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.alwaysDisplayButtonText = !(SendMessage(TB_GETEXTENDEDSTYLE, 0, 0) & TBSTYLE_EX_MIXEDBUTTONS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.alwaysDisplayButtonText);
	return S_OK;
}

STDMETHODIMP ToolBar::put_AlwaysDisplayButtonText(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_ALWAYSDISPLAYBUTTONTEXT);
	if(properties.menuMode && newValue == VARIANT_FALSE) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.alwaysDisplayButtonText != b) {
		properties.alwaysDisplayButtonText = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TB_SETEXTENDEDSTYLE, TBSTYLE_EX_MIXEDBUTTONS, (properties.alwaysDisplayButtonText ? 0 : TBSTYLE_EX_MIXEDBUTTONS));
		}
		FireOnChanged(DISPID_TB_ALWAYSDISPLAYBUTTONTEXT);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_AnchorHighlighting(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.anchorHighlighting = SendMessage(TB_GETANCHORHIGHLIGHT, 0, 0);
	}

	*pValue = BOOL2VARIANTBOOL(properties.anchorHighlighting);
	return S_OK;
}

STDMETHODIMP ToolBar::put_AnchorHighlighting(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_ANCHORHIGHLIGHTING);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.anchorHighlighting != b) {
		properties.anchorHighlighting = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TB_SETANCHORHIGHLIGHT, properties.anchorHighlighting, 0);
		}
		FireOnChanged(DISPID_TB_ANCHORHIGHLIGHTING);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_Appearance(AppearanceConstants* pValue)
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

STDMETHODIMP ToolBar::put_Appearance(AppearanceConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_APPEARANCE);
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
		FireOnChanged(DISPID_TB_APPEARANCE);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_AppID(SHORT* pValue)
{
	ATLASSERT_POINTER(pValue, SHORT);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = 15;
	return S_OK;
}

STDMETHODIMP ToolBar::get_AppName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"ToolBarControls");
	return S_OK;
}

STDMETHODIMP ToolBar::get_AppShortName(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"TBarCtls");
	return S_OK;
}

STDMETHODIMP ToolBar::get_BackStyle(BackStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, BackStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.backStyle = ((GetStyle() & TBSTYLE_TRANSPARENT) == TBSTYLE_TRANSPARENT ? bksTransparent : bksOpaque);
	}

	*pValue = properties.backStyle;
	return S_OK;
}

STDMETHODIMP ToolBar::put_BackStyle(BackStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_BACKSTYLE);
	if(properties.backStyle != newValue) {
		properties.backStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.backStyle) {
				case bksTransparent:
					ModifyStyle(0, TBSTYLE_TRANSPARENT);
					break;
				case bksOpaque:
					ModifyStyle(TBSTYLE_TRANSPARENT, 0);
					break;
			}
			flags.applyBackgroundHack = FALSE;
			if(GetStyle() & TBSTYLE_TRANSPARENT) {
				if(!GetWindowTheme(*this)) {
					TCHAR pClassName[300];
					GetClassName(GetParent(), pClassName, 300);
					if(lstrcmpi(pClassName, TEXT("ReBarU")) == 0 || lstrcmpi(pClassName, TEXT("ReBarA")) == 0 || lstrcmpi(pClassName, REBARCLASSNAME) == 0) {
						flags.applyBackgroundHack = TRUE;
					}
				}
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_BACKSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_BorderStyle(BorderStyleConstants* pValue)
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

STDMETHODIMP ToolBar::put_BorderStyle(BorderStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_BORDERSTYLE);
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
		FireOnChanged(DISPID_TB_BORDERSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_Build(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = VERSION_BUILD;
	return S_OK;
}

STDMETHODIMP ToolBar::get_ButtonHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.buttonHeight = HIWORD(SendMessage(TB_GETBUTTONSIZE, 0, 0));
	}

	*pValue = properties.buttonHeight;
	return S_OK;
}

STDMETHODIMP ToolBar::put_ButtonHeight(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_BUTTONHEIGHT);

	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.buttonHeight != newValue) {
		properties.buttonHeight = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TB_SETBUTTONSIZE, 0, MAKELPARAM(LOWORD(SendMessage(TB_GETBUTTONSIZE, 0, 0)), properties.buttonHeight));
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_BUTTONHEIGHT);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_ButtonRowCount(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = SendMessage(TB_GETROWS, 0, 0);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_Buttons(IToolBarButtons** ppButtons)
{
	ATLASSERT_POINTER(ppButtons, IToolBarButtons*);
	if(!ppButtons) {
		return E_POINTER;
	}

	ClassFactory::InitToolBarButtons(this, IID_IToolBarButtons, reinterpret_cast<LPUNKNOWN*>(ppButtons));
	return S_OK;
}

STDMETHODIMP ToolBar::get_ButtonStyle(ButtonStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ButtonStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & TBSTYLE_FLAT) {
			properties.buttonStyle = bstFlat;
		} else {
			properties.buttonStyle = bst3D;
		}
	}

	*pValue = properties.buttonStyle;
	return S_OK;
}

STDMETHODIMP ToolBar::put_ButtonStyle(ButtonStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_BUTTONSTYLE);
	if(properties.buttonStyle != newValue) {
		properties.buttonStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.buttonStyle) {
				case bst3D:
					ModifyStyle(TBSTYLE_FLAT, 0);
					break;
				case bstFlat:
					ModifyStyle(0, TBSTYLE_FLAT);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_BUTTONSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_ButtonTextPosition(ButtonTextPositionConstants* pValue)
{
	ATLASSERT_POINTER(pValue, ButtonTextPositionConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & TBSTYLE_LIST) {
			properties.buttonTextPosition = btpRightToIcon;
		} else {
			properties.buttonTextPosition = btpBelowIcon;
		}
	}

	*pValue = properties.buttonTextPosition;
	return S_OK;
}

STDMETHODIMP ToolBar::put_ButtonTextPosition(ButtonTextPositionConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_BUTTONTEXTPOSITION);
	if(properties.menuMode && newValue != btpRightToIcon) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.buttonTextPosition != newValue) {
		properties.buttonTextPosition = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.buttonTextPosition) {
				case btpBelowIcon:
					properties.defaultDrawTextFlags &= ~(DT_LEFT | DT_VCENTER | DT_SINGLELINE);
					properties.defaultDrawTextFlags |= DT_CENTER | DT_TOP;
					ModifyStyle(TBSTYLE_LIST, 0);
					// reset to default alignment
					put_HorizontalTextAlignment(halCenter);
					put_VerticalTextAlignment(valTop);
					break;
				case btpRightToIcon:
					properties.defaultDrawTextFlags &= ~(DT_CENTER | DT_TOP);
					properties.defaultDrawTextFlags |= DT_LEFT | DT_VCENTER | DT_SINGLELINE;
					ModifyStyle(0, TBSTYLE_LIST);
					// reset to default alignment
					put_HorizontalTextAlignment(halLeft);
					put_VerticalTextAlignment(valCenter);
					break;
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_BUTTONTEXTPOSITION);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_ButtonWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.buttonWidth = LOWORD(SendMessage(TB_GETBUTTONSIZE, 0, 0));
	}

	*pValue = properties.buttonWidth;
	return S_OK;
}

STDMETHODIMP ToolBar::put_ButtonWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_BUTTONWIDTH);

	if(newValue < 0) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.buttonWidth != newValue) {
		properties.buttonWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TB_SETBUTTONSIZE, 0, MAKELPARAM(properties.buttonWidth, HIWORD(SendMessage(TB_GETBUTTONSIZE, 0, 0))));
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_BUTTONWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_CharSet(BSTR* pValue)
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

STDMETHODIMP ToolBar::get_CommandEnabled(LONG commandID, VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode()) {
		*pValue = VARIANT_TRUE;
		#ifdef USE_STL
			for(std::vector<LONG>::iterator it = properties.disabledCommands.begin(); it != properties.disabledCommands.end(); it++) {
				if(*it == commandID) {
					*pValue = VARIANT_FALSE;
					break;
				}
			}
		#else
			for(size_t i = 0; i < properties.disabledCommands.GetCount(); ++i) {
				if(properties.disabledCommands[i] == commandID) {
					*pValue = VARIANT_FALSE;
					break;
				}
			}
		#endif
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::put_CommandEnabled(LONG commandID, VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_COMMANDENABLED);
	UINT b = VARIANTBOOL2BOOL(newValue);

	if(!IsInDesignMode()) {
		BOOL addEntry = !b;
		#ifdef USE_STL
			for(std::vector<LONG>::iterator it = properties.disabledCommands.begin(); it != properties.disabledCommands.end(); it++) {
				if(*it == commandID) {
					if(b) {
						// remove this entry
						properties.disabledCommands.erase(it);
					} else {
						// already disabled
						addEntry = FALSE;
					}
					break;
				}
			}
			if(addEntry) {
				properties.disabledCommands.push_back(commandID);
			}
		#else
			for(size_t i = 0; i < properties.disabledCommands.GetCount(); ++i) {
				if(properties.disabledCommands[i] == commandID) {
					if(b) {
						// remove this entry
						properties.disabledCommands.RemoveAt(i);
					} else {
						// already disabled
						addEntry = FALSE;
					}
					break;
				}
			}
			if(addEntry) {
				properties.disabledCommands.Add(commandID);
			}
		#endif
		SendMessage(TB_ENABLEBUTTON, commandID, MAKELPARAM(b, 0));
		FireViewChange();
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::get_DisabledEvents(DisabledEventsConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DisabledEventsConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.disabledEvents;
	return S_OK;
}

STDMETHODIMP ToolBar::put_DisabledEvents(DisabledEventsConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_DISABLEDEVENTS);
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
					buttonUnderMouse = -1;
				}
			}
		}

		properties.disabledEvents = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_DISABLEDEVENTS);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_DisplayMenuDivider(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.displayMenuDivider = !(GetStyle() & CCS_NODIVIDER);
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayMenuDivider);
	return S_OK;
}

STDMETHODIMP ToolBar::put_DisplayMenuDivider(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_DISPLAYMENUDIVIDER);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayMenuDivider != b) {
		properties.displayMenuDivider = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.displayMenuDivider) {
				ModifyStyle(CCS_NODIVIDER, 0);
			} else {
				//ModifyStyle(0, CCS_NODIVIDER);
				RecreateControlWindow();
			}
		}
		FireOnChanged(DISPID_TB_DISPLAYMENUDIVIDER);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_DisplayPartiallyClippedButtons(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.displayPartiallyClippedButtons = !(SendMessage(TB_GETEXTENDEDSTYLE, 0, 0) & TBSTYLE_EX_HIDECLIPPEDBUTTONS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.displayPartiallyClippedButtons);
	return S_OK;
}

STDMETHODIMP ToolBar::put_DisplayPartiallyClippedButtons(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_DISPLAYPARTIALLYCLIPPEDBUTTONS);

	if(newValue == VARIANT_FALSE) {
		OrientationConstants orientation = oHorizontal;
		if(SUCCEEDED(get_Orientation(&orientation)) && orientation == oVertical) {
			// invalid value - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
		}
	}
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.displayPartiallyClippedButtons != b) {
		properties.displayPartiallyClippedButtons = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TB_SETEXTENDEDSTYLE, TBSTYLE_EX_HIDECLIPPEDBUTTONS, (properties.displayPartiallyClippedButtons ? 0 : TBSTYLE_EX_HIDECLIPPEDBUTTONS));
		}
		FireOnChanged(DISPID_TB_DISPLAYPARTIALLYCLIPPEDBUTTONS);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_DontRedraw(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.dontRedraw);
	return S_OK;
}

STDMETHODIMP ToolBar::put_DontRedraw(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_DONTREDRAW);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.dontRedraw != b) {
		properties.dontRedraw = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetRedraw(!b);
		}
		FireOnChanged(DISPID_TB_DONTREDRAW);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_DragClickTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dragClickTime;
	return S_OK;
}

STDMETHODIMP ToolBar::put_DragClickTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TB_DRAGCLICKTIME);

	if((newValue < -1) || (newValue > 60000)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}
	if(properties.dragClickTime != newValue) {
		properties.dragClickTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_DRAGCLICKTIME);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_DragDropCustomizationModifierKey(DragDropCustomizationModifierKeyConstants* pValue)
{
	ATLASSERT_POINTER(pValue, DragDropCustomizationModifierKeyConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(GetStyle() & TBSTYLE_ALTDRAG) {
			properties.dragDropCustomizationModifierKey = ddcmkAlt;
		} else {
			properties.dragDropCustomizationModifierKey = ddcmkShift;
		}
	}

	*pValue = properties.dragDropCustomizationModifierKey;
	return S_OK;
}

STDMETHODIMP ToolBar::put_DragDropCustomizationModifierKey(DragDropCustomizationModifierKeyConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_DRAGDROPCUSTOMIZATIONMODIFIERKEY);
	if(properties.dragDropCustomizationModifierKey != newValue) {
		properties.dragDropCustomizationModifierKey = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			switch(properties.dragDropCustomizationModifierKey) {
				case ddcmkShift:
					ModifyStyle(TBSTYLE_ALTDRAG, 0);
					break;
				case ddcmkAlt:
					ModifyStyle(0,TBSTYLE_ALTDRAG);
					break;
			}
		}
		FireOnChanged(DISPID_TB_DRAGDROPCUSTOMIZATIONMODIFIERKEY);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_DraggedButtons(IToolBarButtonContainer** ppButtons)
{
	ATLASSERT_POINTER(ppButtons, IToolBarButtonContainer*);
	if(!ppButtons) {
		return E_POINTER;
	}

	*ppButtons = NULL;
	if(dragDropStatus.pDraggedButtons) {
		return dragDropStatus.pDraggedButtons->Clone(ppButtons);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_DropDownGap(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.dropDownGap;
	return S_OK;
}

STDMETHODIMP ToolBar::put_DropDownGap(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_DROPDOWNGAP);
	if(properties.dropDownGap != newValue) {
		properties.dropDownGap = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.dropDownGap == -1) {
				SendMessage(TB_SETDROPDOWNGAP, GetSystemMetrics(SM_CXEDGE) * 2, 0);
			} else {
				SendMessage(TB_SETDROPDOWNGAP, properties.dropDownGap, 0);
			}
			// this will cause a proper recalculation of the button rectangles
			properties.alwaysDisplayButtonText = !(SendMessage(TB_GETEXTENDEDSTYLE, 0, 0) & TBSTYLE_EX_MIXEDBUTTONS);
			if(properties.alwaysDisplayButtonText) {
				SendMessage(TB_SETEXTENDEDSTYLE, TBSTYLE_EX_MIXEDBUTTONS, (properties.alwaysDisplayButtonText ? TBSTYLE_EX_MIXEDBUTTONS : 0));
				SendMessage(TB_SETEXTENDEDSTYLE, TBSTYLE_EX_MIXEDBUTTONS, (properties.alwaysDisplayButtonText ? 0 : TBSTYLE_EX_MIXEDBUTTONS));
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_DROPDOWNGAP);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_Enabled(VARIANT_BOOL* pValue)
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

STDMETHODIMP ToolBar::put_Enabled(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_ENABLED);
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

		FireOnChanged(DISPID_TB_ENABLED);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_FirstButtonIndentation(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.firstButtonIndentation;
	return S_OK;
}

STDMETHODIMP ToolBar::put_FirstButtonIndentation(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_FIRSTBUTTONINDENTATION);
	if(properties.firstButtonIndentation != newValue) {
		properties.firstButtonIndentation = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TB_SETINDENT, properties.firstButtonIndentation, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_FIRSTBUTTONINDENTATION);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_FocusOnClick(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.focusOnClick);
	return S_OK;
}

STDMETHODIMP ToolBar::put_FocusOnClick(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_FOCUSONCLICK);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.focusOnClick != b) {
		properties.focusOnClick = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_FOCUSONCLICK);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_Font(IFontDisp** ppFont)
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

STDMETHODIMP ToolBar::put_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_TB_FONT);
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
	FireOnChanged(DISPID_TB_FONT);
	return S_OK;
}

STDMETHODIMP ToolBar::putref_Font(IFontDisp* pNewFont)
{
	PUTPROPPROLOG(DISPID_TB_FONT);
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
	FireOnChanged(DISPID_TB_FONT);
	return S_OK;
}

STDMETHODIMP ToolBar::get_hChevronMenu(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(properties.hChevronPopupMenu);
	return S_OK;
}

STDMETHODIMP ToolBar::get_HighlightColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
		if(SendMessage(TB_GETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme))) {
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

STDMETHODIMP ToolBar::put_HighlightColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TB_HIGHLIGHTCOLOR);
	if(properties.highlightColor != newValue) {
		properties.highlightColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
			if(SendMessage(TB_GETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme))) {
				colorScheme.clrBtnHighlight = (properties.highlightColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.highlightColor));
				SendMessage(TB_SETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme));
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_TB_HIGHLIGHTCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_hImageList(ImageListConstants imageList, LONG imageListIndex/* = 0*/, OLE_HANDLE* pValue/* = NULL*/)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = NULL;
	switch(imageList) {
		case ilNormalButtons:
			if(IsWindow()) {
				*pValue = HandleToLong(reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETIMAGELIST, imageListIndex, 0)));
				return S_OK;
			}
			break;
		case ilHotButtons:
			if(IsWindow()) {
				*pValue = HandleToLong(reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETHOTIMAGELIST, imageListIndex, 0)));
				return S_OK;
			}
			break;
		case ilPushedButtons:
			if(IsWindow()) {
				*pValue = HandleToLong(reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETPRESSEDIMAGELIST, imageListIndex, 0)));
				return S_OK;
			}
			break;
		case ilDisabledButtons:
			if(IsWindow()) {
				*pValue = HandleToLong(reinterpret_cast<HIMAGELIST>(SendMessage(TB_GETDISABLEDIMAGELIST, imageListIndex, 0)));
				return S_OK;
			}
			break;
		case ilHighResolution:
			{
				#ifdef USE_STL
					std::unordered_map<LONG, HIMAGELIST>::iterator iter = properties.hHighResImageLists.find(imageListIndex);
					if(iter != properties.hHighResImageLists.end()) {
						*pValue = HandleToLong(iter->second);
					} else {
						*pValue = NULL;
					}
				#else
					CAtlMap<LONG, HIMAGELIST>::CPair* pEntry = properties.hHighResImageLists.Lookup(imageListIndex);
					if(pEntry) {
						*pValue = HandleToLong(pEntry->m_value);
					} else {
						*pValue = NULL;
					}
				#endif
			}
			return S_OK;
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}
	return S_OK;
}

STDMETHODIMP ToolBar::put_hImageList(ImageListConstants imageList, LONG imageListIndex/* = 0*/, OLE_HANDLE newValue/* = NULL*/)
{
	PUTPROPPROLOG(DISPID_TB_HIMAGELIST);
	BOOL fireViewChange = TRUE;
	switch(imageList) {
		case ilNormalButtons:
			if(IsWindow()) {
				SendMessage(TB_SETIMAGELIST, imageListIndex, reinterpret_cast<LPARAM>(LongToHandle(newValue)));
				return S_OK;
			}
			break;
		case ilHotButtons:
			if(IsWindow()) {
				SendMessage(TB_SETHOTIMAGELIST, imageListIndex, reinterpret_cast<LPARAM>(LongToHandle(newValue)));
				return S_OK;
			}
			break;
		case ilPushedButtons:
			if(IsWindow()) {
				SendMessage(TB_SETPRESSEDIMAGELIST, imageListIndex, reinterpret_cast<LPARAM>(LongToHandle(newValue)));
				return S_OK;
			}
			break;
		case ilDisabledButtons:
			if(IsWindow()) {
				SendMessage(TB_SETDISABLEDIMAGELIST, imageListIndex, reinterpret_cast<LPARAM>(LongToHandle(newValue)));
				return S_OK;
			}
			break;
		case ilHighResolution:
			properties.hHighResImageLists[imageListIndex] = reinterpret_cast<HIMAGELIST>(LongToHandle(newValue));
			return S_OK;
			break;
		default:
			// invalid arg - raise VB runtime error 380
			return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
			break;
	}

	FireOnChanged(DISPID_TB_HIMAGELIST);
	if(fireViewChange) {
		FireViewChange();
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_HorizontalButtonPadding(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		/*TBMETRICS metrics = {0};
		metrics.cbSize = sizeof(metrics);
		metrics.dwMask = TBMF_PAD;
		SendMessage(TB_GETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
		properties.horizontalButtonPadding = metrics.cxPad;*/
		properties.horizontalButtonPadding = GET_X_LPARAM(SendMessage(TB_GETPADDING, 0, 0));
	}

	*pValue = properties.horizontalButtonPadding;
	return S_OK;
}

STDMETHODIMP ToolBar::put_HorizontalButtonPadding(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_HORIZONTALBUTTONPADDING);
	if(properties.horizontalButtonPadding != newValue) {
		properties.horizontalButtonPadding = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			/*TBMETRICS metrics = {0};
			metrics.cbSize = sizeof(metrics);
			metrics.dwMask = TBMF_PAD;
			SendMessage(TB_GETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
			metrics.cxPad = properties.horizontalButtonPadding;
			SendMessage(TB_SETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));*/
			SHORT horizontalPadding = static_cast<SHORT>(properties.horizontalButtonPadding);
			if(horizontalPadding == -1) {
				horizontalPadding = GET_X_LPARAM(properties.defaultButtonPadding);
			}
			SendMessage(TB_SETPADDING, 0, MAKELPARAM(horizontalPadding, -1));
			// this will cause a proper recalculation of the button rectangles
			SendMessage(TB_SETINDENT, properties.firstButtonIndentation, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_HORIZONTALBUTTONPADDING);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_HorizontalButtonSpacing(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		TBMETRICS metrics = {0};
		metrics.cbSize = sizeof(metrics);
		metrics.dwMask = TBMF_BUTTONSPACING;
		SendMessage(TB_GETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
		properties.horizontalButtonSpacing = metrics.cxButtonSpacing;
	}

	*pValue = properties.horizontalButtonSpacing;
	return S_OK;
}

STDMETHODIMP ToolBar::put_HorizontalButtonSpacing(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_HORIZONTALBUTTONSPACING);
	if(properties.horizontalButtonSpacing != newValue) {
		properties.horizontalButtonSpacing = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			TBMETRICS metrics = {0};
			metrics.cbSize = sizeof(metrics);
			metrics.dwMask = TBMF_BUTTONSPACING;
			SendMessage(TB_GETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
			metrics.cxButtonSpacing = properties.horizontalButtonSpacing;
			SendMessage(TB_SETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
			// this will cause a proper recalculation of the button rectangles
			SendMessage(TB_SETINDENT, properties.firstButtonIndentation, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_HORIZONTALBUTTONSPACING);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_HorizontalIconCaptionGap(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.horizontalIconCaptionGap;
	return S_OK;
}

STDMETHODIMP ToolBar::put_HorizontalIconCaptionGap(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_HORIZONTALICONCAPTIONGAP);
	if(properties.horizontalIconCaptionGap != newValue) {
		properties.horizontalIconCaptionGap = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.horizontalIconCaptionGap == -1) {
				SendMessage(TB_SETLISTGAP, GetSystemMetrics(SM_CXEDGE) * 2, 0);
			} else {
				SendMessage(TB_SETLISTGAP, properties.horizontalIconCaptionGap, 0);
			}
			// this will cause a proper recalculation of the button rectangles
			SendMessage(TB_SETINDENT, properties.firstButtonIndentation, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_HORIZONTALICONCAPTIONGAP);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_HorizontalTextAlignment(HAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, HAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch((properties.drawTextFlags & properties.drawTextFlagsMask) & (DT_LEFT | DT_CENTER | DT_RIGHT)) {
			case DT_CENTER:
				properties.horizontalTextAlignment = halCenter;
				break;
			case DT_RIGHT:
				properties.horizontalTextAlignment = halRight;
				break;
			default:
				properties.horizontalTextAlignment = halLeft;
				break;
		}
	}

	*pValue = properties.horizontalTextAlignment;
	return S_OK;
}

STDMETHODIMP ToolBar::put_HorizontalTextAlignment(HAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_HORIZONTALTEXTALIGNMENT);
	if(properties.horizontalTextAlignment != newValue) {
		properties.horizontalTextAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetDrawTextFlags();
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_HORIZONTALTEXTALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_HotButton(IToolBarButton** ppHotButton)
{
	ATLASSERT_POINTER(ppHotButton, IToolBarButton*);
	if(!ppHotButton) {
		return E_POINTER;
	}

	if(IsWindow()) {
		int hotButton = static_cast<int>(SendMessage(TB_GETHOTITEM, 0, 0));
		ClassFactory::InitToolBarButton(hotButton, FALSE, this, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppHotButton));
	}
	return S_OK;
}

STDMETHODIMP ToolBar::putref_HotButton(IToolBarButton* pNewHotButton)
{
	PUTPROPPROLOG(DISPID_TB_HOTBUTTON);

	int newHotButton = -1;
	if(pNewHotButton) {
		LONG l = -1;
		pNewHotButton->get_Index(&l);
		newHotButton = l;
		// TODO: Shouldn't we AddRef' pNewHotButton?
	}

	if(IsWindow()) {
		SendMessage(TB_SETHOTITEM, newHotButton, 0);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TB_HOTBUTTON);
	return S_OK;
}

STDMETHODIMP ToolBar::get_HoverTime(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.hoverTime;
	return S_OK;
}

STDMETHODIMP ToolBar::put_HoverTime(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TB_HOVERTIME);
	if((newValue < 0) && (newValue != -1)) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.hoverTime != newValue) {
		properties.hoverTime = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_HOVERTIME);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_hWnd(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(*this));
	return S_OK;
}

STDMETHODIMP ToolBar::get_hWndChevronToolBar(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = HandleToLong(static_cast<HWND>(chevronPopupToolbar));
	return S_OK;
}

STDMETHODIMP ToolBar::get_hWndToolTip(OLE_HANDLE* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_HANDLE);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = HandleToLong(reinterpret_cast<HWND>(SendMessage(TB_GETTOOLTIPS, 0, 0)));
	}
	return S_OK;
}

STDMETHODIMP ToolBar::put_hWndToolTip(OLE_HANDLE newValue)
{
	PUTPROPPROLOG(DISPID_TB_HWNDTOOLTIP);
	if(IsWindow()) {
		SendMessage(TB_SETTOOLTIPS, reinterpret_cast<WPARAM>(LongToHandle(newValue)), 0);
	}

	SetDirty(TRUE);
	FireOnChanged(DISPID_TB_HWNDTOOLTIP);
	return S_OK;
}

STDMETHODIMP ToolBar::get_IdealHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SIZE idealSize = {0};
		if(SendMessage(TB_GETIDEALSIZE, TRUE, reinterpret_cast<LPARAM>(&idealSize))) {
			*pValue = idealSize.cy;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::get_IdealWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SIZE idealSize = {0};
		if(SendMessage(TB_GETIDEALSIZE, FALSE, reinterpret_cast<LPARAM>(&idealSize))) {
			*pValue = idealSize.cx;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::get_ImageListCount(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		*pValue = SendMessage(TB_GETIMAGELISTCOUNT, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::get_InsertMarkColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORREF color = static_cast<COLORREF>(SendMessage(TB_GETINSERTMARKCOLOR, 0, 0));
		if(color == CLR_NONE) {
			properties.insertMarkColor = 0;
		} else if(color != OLECOLOR2COLORREF(properties.insertMarkColor)) {
			properties.insertMarkColor = color;
		}
	}

	*pValue = properties.insertMarkColor;
	return S_OK;
}

STDMETHODIMP ToolBar::put_InsertMarkColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TB_INSERTMARKCOLOR);
	if(properties.insertMarkColor != newValue) {
		properties.insertMarkColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TB_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_INSERTMARKCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_IsRelease(VARIANT_BOOL* pValue)
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

STDMETHODIMP ToolBar::get_MaximumButtonWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.maximumButtonWidth;
	return S_OK;
}

STDMETHODIMP ToolBar::put_MaximumButtonWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_MAXIMUMBUTTONWIDTH);
	if(properties.maximumButtonWidth != newValue) {
		properties.maximumButtonWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TB_SETBUTTONWIDTH, 0, MAKELPARAM(properties.minimumButtonWidth, properties.maximumButtonWidth));
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_MAXIMUMBUTTONWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_MaximumHeight(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SIZE maximumSize = {0};
		if(SendMessage(TB_GETMAXSIZE, 0, reinterpret_cast<LPARAM>(&maximumSize))) {
			*pValue = maximumSize.cy;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::get_MaximumTextRows(LONG* pValue)
{
	ATLASSERT_POINTER(pValue, LONG);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.maximumTextRows = SendMessage(TB_GETTEXTROWS, 0, 0);
	}

	*pValue = properties.maximumTextRows;
	return S_OK;
}

STDMETHODIMP ToolBar::put_MaximumTextRows(LONG newValue)
{
	PUTPROPPROLOG(DISPID_TB_MAXIMUMTEXTROWS);
	if(properties.maximumTextRows != newValue) {
		properties.maximumTextRows = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.maximumTextRows > 1) {
				properties.defaultDrawTextFlags &= ~DT_SINGLELINE;
				properties.defaultDrawTextFlags |= DT_WORDBREAK | DT_EDITCONTROL;
			} else {
				properties.defaultDrawTextFlags &= ~(DT_WORDBREAK | DT_EDITCONTROL);
				properties.defaultDrawTextFlags |= DT_SINGLELINE;
			}
			SendMessage(TB_SETMAXTEXTROWS, properties.maximumTextRows, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_MAXIMUMTEXTROWS);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_MaximumWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SIZE maximumSize = {0};
		if(SendMessage(TB_GETMAXSIZE, 0, reinterpret_cast<LPARAM>(&maximumSize))) {
			*pValue = maximumSize.cx;
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::get_MenuBarTheme(MenuBarThemeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MenuBarThemeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  properties.menuBarTheme;
	return S_OK;
}

STDMETHODIMP ToolBar::put_MenuBarTheme(MenuBarThemeConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_MENUBARTHEME);

	if(properties.menuBarTheme != newValue) {
		properties.menuBarTheme = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_MENUBARTHEME);

		if(properties.menuMode) {
			FireViewChange();
		}
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_MenuMode(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.menuMode);
	return S_OK;
}

STDMETHODIMP ToolBar::put_MenuMode(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_MENUMODE);
	if(!IsInDesignMode() && IsWindow()) {
		// Set not supported at runtime - raise VB runtime error 382
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 382);
	}

	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.menuMode != b) {
		properties.menuMode = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_MENUMODE);

		if(properties.menuMode) {
			put_AlwaysDisplayButtonText(VARIANT_TRUE);
			put_ButtonTextPosition(btpRightToIcon);
			put_HorizontalTextAlignment(halCenter);
			put_NormalDropDownButtonStyle(nddbsWithoutArrow);
			put_VerticalTextAlignment(valCenter);
		}
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_MinimumButtonWidth(OLE_XSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_XSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.minimumButtonWidth;
	return S_OK;
}

STDMETHODIMP ToolBar::put_MinimumButtonWidth(OLE_XSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_MINIMUMBUTTONWIDTH);
	if(properties.minimumButtonWidth != newValue) {
		properties.minimumButtonWidth = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SendMessage(TB_SETBUTTONWIDTH, 0, MAKELPARAM(properties.minimumButtonWidth, properties.maximumButtonWidth));
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_MINIMUMBUTTONWIDTH);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_MouseIcon(IPictureDisp** ppMouseIcon)
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

STDMETHODIMP ToolBar::put_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TB_MOUSEICON);
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
	FireOnChanged(DISPID_TB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ToolBar::putref_MouseIcon(IPictureDisp* pNewMouseIcon)
{
	PUTPROPPROLOG(DISPID_TB_MOUSEICON);
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
	FireOnChanged(DISPID_TB_MOUSEICON);
	return S_OK;
}

STDMETHODIMP ToolBar::get_MousePointer(MousePointerConstants* pValue)
{
	ATLASSERT_POINTER(pValue, MousePointerConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.mousePointer;
	return S_OK;
}

STDMETHODIMP ToolBar::put_MousePointer(MousePointerConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_MOUSEPOINTER);
	if(properties.mousePointer != newValue) {
		properties.mousePointer = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_MOUSEPOINTER);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_MultiColumn(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.multiColumn = !(SendMessage(TB_GETEXTENDEDSTYLE, 0, 0) & TBSTYLE_EX_MULTICOLUMN);
	}

	*pValue =  BOOL2VARIANTBOOL(properties.multiColumn);
	return S_OK;
}

STDMETHODIMP ToolBar::put_MultiColumn(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_MULTICOLUMN);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.multiColumn != b) {
		properties.multiColumn = b;
		SetDirty(TRUE);

		HWND hWnd = *this;
		if(properties.multiColumn) {
			put_Orientation(oVertical);
		}
		if(IsWindow()) {
			if(hWnd == *this) {
				RecreateControlWindow();
			}
		}
		FireOnChanged(DISPID_TB_MULTICOLUMN);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_NativeDropTarget(LPUNKNOWN* ppValue)
{
	ATLASSERT_POINTER(ppValue, LPUNKNOWN);
	if(!ppValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		SendMessage(TB_GETOBJECT, reinterpret_cast<WPARAM>(&IID_IDropTarget), reinterpret_cast<LPARAM>(ppValue));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::get_NormalDropDownButtonStyle(NormalDropDownButtonStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, NormalDropDownButtonStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if(SendMessage(TB_GETEXTENDEDSTYLE, 0, 0) & TBSTYLE_EX_DRAWDDARROWS) {
			properties.normalDropDownButtonStyle = nddbsSplitButton;
		} else {
			properties.normalDropDownButtonStyle = nddbsWithoutArrow;
		}
	}

	*pValue = properties.normalDropDownButtonStyle;
	return S_OK;
}

STDMETHODIMP ToolBar::put_NormalDropDownButtonStyle(NormalDropDownButtonStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_NORMALDROPDOWNBUTTONSTYLE);
	if(properties.menuMode && newValue != nddbsWithoutArrow) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.normalDropDownButtonStyle != newValue) {
		properties.normalDropDownButtonStyle = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			/*switch(properties.normalDropDownButtonStyle) {
				case nddbsSplitButton:
					SendMessage(TB_SETEXTENDEDSTYLE, TBSTYLE_EX_DRAWDDARROWS, TBSTYLE_EX_DRAWDDARROWS);
					break;
				case nddbsWithoutArrow:
					SendMessage(TB_SETEXTENDEDSTYLE, TBSTYLE_EX_DRAWDDARROWS, 0);
					break;
			}*/
			RecreateControlWindow();
		}
		FireOnChanged(DISPID_TB_NORMALDROPDOWNBUTTONSTYLE);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue)
{
	ATLASSERT_POINTER(pValue, OLEDragImageStyleConstants);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = properties.oleDragImageStyle;
	return S_OK;
}

STDMETHODIMP ToolBar::put_OLEDragImageStyle(OLEDragImageStyleConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_OLEDRAGIMAGESTYLE);
	if(properties.oleDragImageStyle != newValue) {
		properties.oleDragImageStyle = newValue;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_OLEDRAGIMAGESTYLE);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_Orientation(OrientationConstants* pValue)
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

STDMETHODIMP ToolBar::put_Orientation(OrientationConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_ORIENTATION);
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

		HWND hWnd = *this;
		if(properties.orientation == oHorizontal) {
			put_MultiColumn(VARIANT_FALSE);
		}
		if(IsWindow()) {
			if(hWnd == *this) {
				RecreateControlWindow();
			}
		}
		FireOnChanged(DISPID_TB_ORIENTATION);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_ProcessContextMenuKeys(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.processContextMenuKeys);
	return S_OK;
}

STDMETHODIMP ToolBar::put_ProcessContextMenuKeys(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_PROCESSCONTEXTMENUKEYS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.processContextMenuKeys != b) {
		properties.processContextMenuKeys = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_PROCESSCONTEXTMENUKEYS);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_Programmer(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ToolBar::get_RaiseCustomDrawEventOnEraseBackground(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.raiseCustomDrawEventOnEraseBackground = ((GetStyle() & TBSTYLE_CUSTOMERASE) == TBSTYLE_CUSTOMERASE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.raiseCustomDrawEventOnEraseBackground);
	return S_OK;
}

STDMETHODIMP ToolBar::put_RaiseCustomDrawEventOnEraseBackground(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_RAISECUSTOMDRAWEVENTONERASEBACKGROUND);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.raiseCustomDrawEventOnEraseBackground != b) {
		properties.raiseCustomDrawEventOnEraseBackground = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.raiseCustomDrawEventOnEraseBackground || flags.applyBackgroundHack) {
				ModifyStyle(0, TBSTYLE_CUSTOMERASE);
			} else {
				ModifyStyle(TBSTYLE_CUSTOMERASE, 0);
			}
		}
		FireOnChanged(DISPID_TB_RAISECUSTOMDRAWEVENTONERASEBACKGROUND);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants* pValue)
{
	ATLASSERT_POINTER(pValue, RegisterForOLEDragDropConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		if((GetStyle() & TBSTYLE_REGISTERDROP) == TBSTYLE_REGISTERDROP) {
			properties.registerForOLEDragDrop = rfoddNativeDragDrop;
		}
	}

	*pValue = properties.registerForOLEDragDrop;
	return S_OK;
}

STDMETHODIMP ToolBar::put_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_REGISTERFOROLEDRAGDROP);
	if(properties.registerForOLEDragDrop != newValue) {
		properties.registerForOLEDragDrop = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			ModifyStyle(TBSTYLE_REGISTERDROP, 0);
			RevokeDragDrop(*this);
			switch(properties.registerForOLEDragDrop) {
				case rfoddNativeDragDrop:
					ModifyStyle(0, TBSTYLE_REGISTERDROP);
					break;
				case rfoddAdvancedDragDrop: {
					ATLVERIFY(RegisterDragDrop(*this, static_cast<IDropTarget*>(this)) == S_OK);
					break;
				}
			}
		}
		FireOnChanged(DISPID_TB_REGISTERFOROLEDRAGDROP);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_RightToLeft(RightToLeftConstants* pValue)
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

STDMETHODIMP ToolBar::put_RightToLeft(RightToLeftConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_RIGHTTOLEFT);
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
		FireOnChanged(DISPID_TB_RIGHTTOLEFT);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_ShadowColor(OLE_COLOR* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_COLOR);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
		if(SendMessage(TB_GETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme))) {
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

STDMETHODIMP ToolBar::put_ShadowColor(OLE_COLOR newValue)
{
	PUTPROPPROLOG(DISPID_TB_SHADOWCOLOR);
	if(properties.shadowColor != newValue) {
		properties.shadowColor = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
			if(SendMessage(TB_GETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme))) {
				colorScheme.clrBtnShadow = (properties.shadowColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.shadowColor));
				SendMessage(TB_SETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme));
				FireViewChange();
			}
		}
		FireOnChanged(DISPID_TB_SHADOWCOLOR);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_ShowDragImage(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!dragDropStatus.IsDragging()) {
		return E_FAIL;
	}

	*pValue =  BOOL2VARIANTBOOL(dragDropStatus.IsDragImageVisible());
	return S_OK;
}

STDMETHODIMP ToolBar::put_ShowDragImage(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_SHOWDRAGIMAGE);

	if(!dragDropStatus.hDragImageList && !dragDropStatus.pDropTargetHelper) {
		return E_FAIL;
	}

	if(newValue == VARIANT_FALSE) {
		dragDropStatus.HideDragImage(FALSE);
	} else {
		dragDropStatus.ShowDragImage(FALSE);
	}

	FireOnChanged(DISPID_TB_SHOWDRAGIMAGE);
	return S_OK;
}

STDMETHODIMP ToolBar::get_ShowToolTips(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.showToolTips = ((GetStyle() & TBSTYLE_TOOLTIPS) == TBSTYLE_TOOLTIPS);
	}

	*pValue = BOOL2VARIANTBOOL(properties.showToolTips);
	return S_OK;
}

STDMETHODIMP ToolBar::put_ShowToolTips(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_SHOWTOOLTIPS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.showToolTips != b) {
		properties.showToolTips = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.showToolTips) {
				ModifyStyle(0, TBSTYLE_TOOLTIPS);
			} else {
				ModifyStyle(TBSTYLE_TOOLTIPS, 0);
			}
		}
		FireOnChanged(DISPID_TB_SHOWTOOLTIPS);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_SuggestedIconSize(SuggestedIconSizeConstants* pValue)
{
	ATLASSERT_POINTER(pValue, SuggestedIconSizeConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(IsWindow()) {
		DWORD flags = SendMessage(TB_GETBITMAPFLAGS, 0, 0);
		if(flags & TBBF_LARGE) {
			*pValue = sisLarge;
		} else {
			*pValue = sisSmall;
		}
		return S_OK;
	}

	return E_FAIL;
}

STDMETHODIMP ToolBar::get_SupportOLEDragImages(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue =  BOOL2VARIANTBOOL(properties.supportOLEDragImages);
	return S_OK;
}

STDMETHODIMP ToolBar::put_SupportOLEDragImages(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_SUPPORTOLEDRAGIMAGES);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.supportOLEDragImages != b) {
		properties.supportOLEDragImages = b;
		SetDirty(TRUE);
		FireOnChanged(DISPID_TB_SUPPORTOLEDRAGIMAGES);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_Tester(BSTR* pValue)
{
	ATLASSERT_POINTER(pValue, BSTR);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = SysAllocString(L"Timo \"TimoSoft\" Kunze");
	return S_OK;
}

STDMETHODIMP ToolBar::get_UseMnemonics(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.useMnemonics = !((properties.drawTextFlags & properties.drawTextFlagsMask) & DT_NOPREFIX);
	}

	*pValue = BOOL2VARIANTBOOL(properties.useMnemonics);
	return S_OK;
}

STDMETHODIMP ToolBar::put_UseMnemonics(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_USEMNEMONICS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useMnemonics != b) {
		properties.useMnemonics = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetDrawTextFlags();
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_USEMNEMONICS);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_UseSystemFont(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	*pValue = BOOL2VARIANTBOOL(properties.useSystemFont);
	return S_OK;
}

STDMETHODIMP ToolBar::put_UseSystemFont(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_USESYSTEMFONT);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.useSystemFont != b) {
		properties.useSystemFont = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			ApplyFont();
		}
		FireOnChanged(DISPID_TB_USESYSTEMFONT);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_Version(BSTR* pValue)
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

STDMETHODIMP ToolBar::get_VerticalButtonPadding(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		/*TBMETRICS metrics = {0};
		metrics.cbSize = sizeof(metrics);
		metrics.dwMask = TBMF_PAD;
		SendMessage(TB_GETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
		properties.verticalButtonPadding = metrics.cyPad;*/
		properties.verticalButtonPadding = GET_Y_LPARAM(SendMessage(TB_GETPADDING, 0, 0));
	}

	*pValue = properties.verticalButtonPadding;
	return S_OK;
}

STDMETHODIMP ToolBar::put_VerticalButtonPadding(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_VERTICALBUTTONPADDING);
	if(properties.verticalButtonPadding != newValue) {
		properties.verticalButtonPadding = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			/*TBMETRICS metrics = {0};
			metrics.cbSize = sizeof(metrics);
			metrics.dwMask = TBMF_PAD;
			SendMessage(TB_GETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
			metrics.cyPad = properties.verticalButtonPadding;
			SendMessage(TB_SETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));*/
			SHORT verticalPadding = static_cast<SHORT>(properties.verticalButtonPadding);
			if(verticalPadding == -1) {
				verticalPadding = GET_Y_LPARAM(properties.defaultButtonPadding);
			}
			SendMessage(TB_SETPADDING, 0, MAKELPARAM(-1, verticalPadding));
			// this will cause a proper recalculation of the button rectangles
			SendMessage(TB_SETINDENT, properties.firstButtonIndentation, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_VERTICALBUTTONPADDING);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_VerticalButtonSpacing(OLE_YSIZE_PIXELS* pValue)
{
	ATLASSERT_POINTER(pValue, OLE_YSIZE_PIXELS);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow() && RunTimeHelper::IsCommCtrl6()) {
		TBMETRICS metrics = {0};
		metrics.cbSize = sizeof(metrics);
		metrics.dwMask = TBMF_BUTTONSPACING;
		SendMessage(TB_GETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
		properties.verticalButtonSpacing = metrics.cyButtonSpacing;
	}

	*pValue = properties.verticalButtonSpacing;
	return S_OK;
}

STDMETHODIMP ToolBar::put_VerticalButtonSpacing(OLE_YSIZE_PIXELS newValue)
{
	PUTPROPPROLOG(DISPID_TB_VERTICALBUTTONSPACING);
	if(properties.verticalButtonSpacing != newValue) {
		properties.verticalButtonSpacing = newValue;
		SetDirty(TRUE);

		if(IsWindow() && RunTimeHelper::IsCommCtrl6()) {
			TBMETRICS metrics = {0};
			metrics.cbSize = sizeof(metrics);
			metrics.dwMask = TBMF_BUTTONSPACING;
			SendMessage(TB_GETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
			metrics.cyButtonSpacing = properties.verticalButtonSpacing;
			SendMessage(TB_SETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
			// this will cause a proper recalculation of the button rectangles
			SendMessage(TB_SETINDENT, properties.firstButtonIndentation, 0);
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_VERTICALBUTTONSPACING);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_VerticalTextAlignment(VAlignmentConstants* pValue)
{
	ATLASSERT_POINTER(pValue, VAlignmentConstants);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		switch((properties.drawTextFlags & properties.drawTextFlagsMask) & (DT_TOP | DT_VCENTER | DT_BOTTOM)) {
			case DT_VCENTER:
				properties.verticalTextAlignment = valCenter;
				break;
			case DT_BOTTOM:
				properties.verticalTextAlignment = valBottom;
				break;
			default:
				properties.verticalTextAlignment = valTop;
				break;
		}
	}

	*pValue = properties.verticalTextAlignment;
	return S_OK;
}

STDMETHODIMP ToolBar::put_VerticalTextAlignment(VAlignmentConstants newValue)
{
	PUTPROPPROLOG(DISPID_TB_VERTICALTEXTALIGNMENT);
	if(properties.menuMode && newValue != valCenter) {
		// invalid value - raise VB runtime error 380
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 380);
	}

	if(properties.verticalTextAlignment != newValue) {
		properties.verticalTextAlignment = newValue;
		SetDirty(TRUE);

		if(IsWindow()) {
			SetDrawTextFlags();
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_VERTICALTEXTALIGNMENT);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::get_WrapButtons(VARIANT_BOOL* pValue)
{
	ATLASSERT_POINTER(pValue, VARIANT_BOOL);
	if(!pValue) {
		return E_POINTER;
	}

	if(!IsInDesignMode() && IsWindow()) {
		properties.wrapButtons = ((GetStyle() & TBSTYLE_WRAPABLE) == TBSTYLE_WRAPABLE);
	}

	*pValue = BOOL2VARIANTBOOL(properties.wrapButtons);
	return S_OK;
}

STDMETHODIMP ToolBar::put_WrapButtons(VARIANT_BOOL newValue)
{
	PUTPROPPROLOG(DISPID_TB_WRAPBUTTONS);
	UINT b = VARIANTBOOL2BOOL(newValue);
	if(properties.wrapButtons != b) {
		properties.wrapButtons = b;
		SetDirty(TRUE);

		if(IsWindow()) {
			if(properties.wrapButtons) {
				ModifyStyle(0, TBSTYLE_WRAPABLE);
			} else {
				ModifyStyle(TBSTYLE_WRAPABLE, 0);
			}
			FireViewChange();
		}
		FireOnChanged(DISPID_TB_WRAPBUTTONS);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::About(void)
{
	AboutDlg dlg;
	dlg.SetOwner(this);
	dlg.DoModal();
	return S_OK;
}

STDMETHODIMP ToolBar::AutoSize(void)
{
	if(IsWindow()) {
		SendMessage(TB_AUTOSIZE, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::CountAcceleratorOccurrences(SHORT accelerator, LONG* pCount)
{
	if(IsWindow()) {
		SendMessage(TB_HASACCELERATOR, accelerator, reinterpret_cast<LPARAM>(pCount));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::CreateButtonContainer(VARIANT buttons/* = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR)*/, IToolBarButtonContainer** ppContainer/* = NULL*/)
{
	ATLASSERT_POINTER(ppContainer, IToolBarButtonContainer*);
	if(!ppContainer) {
		return E_POINTER;
	}

	*ppContainer = NULL;
	CComObject<ToolBarButtonContainer>* pTBButtonContainerObj = NULL;
	CComObject<ToolBarButtonContainer>::CreateInstance(&pTBButtonContainerObj);
	pTBButtonContainerObj->AddRef();

	// clone all settings
	pTBButtonContainerObj->SetOwner(this);

	pTBButtonContainerObj->QueryInterface(__uuidof(IToolBarButtonContainer), reinterpret_cast<LPVOID*>(ppContainer));
	pTBButtonContainerObj->Release();

	if(*ppContainer) {
		(*ppContainer)->Add(buttons);
		RegisterButtonContainer(static_cast<IButtonContainer*>(pTBButtonContainerObj));
	}
	return S_OK;
}

STDMETHODIMP ToolBar::Customize(void)
{
	if(IsWindow()) {
		SendMessage(TB_CUSTOMIZE, 0, 0);
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::DisplayChevronPopupWindow(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!IsWindow()) {
		return E_FAIL;
	}

	BOOL isVertical = ((GetStyle() & CCS_VERT) == CCS_VERT);

	// find first invisible button
	RECT availableRectangle = {0};
	GetClientRect(&availableRectangle);
	int buttonIndex = 0;
	int buttonCount = SendMessage(TB_BUTTONCOUNT, 0, 0);
	CRect buttonRectangle;
	for(; buttonIndex < buttonCount; buttonIndex++) {
		if(SendMessage(TB_GETITEMRECT, buttonIndex, reinterpret_cast<LPARAM>(&buttonRectangle))) {
			if(isVertical) {
				if(buttonRectangle.bottom > availableRectangle.bottom) {
					break;
				}
			} else {
				if(buttonRectangle.right > availableRectangle.right) {
					break;
				}
			}
		}
	}
	if(buttonIndex < buttonCount) {
		if(properties.menuMode) {
			properties.hChevronPopupMenu = CreatePopupMenu();
			if(properties.hChevronPopupMenu) {
				TBBUTTON button = {0};
				TBBUTTONINFO buttonInfo = {0};
				buttonInfo.cbSize = sizeof(TBBUTTONINFO);
				buttonInfo.dwMask = TBIF_BYINDEX | TBIF_IMAGE | TBIF_STYLE;
				#ifndef UNICODE
					buttonInfo.dwMask |= TBIF_TEXT;
					TCHAR pBuffer[1000];
					buttonInfo.pszText = pBuffer;
					buttonInfo.cchText = 1000;
				#endif
				MENUITEMINFO menuItem = {0};
				menuItem.cbSize = sizeof(menuItem);
				menuItem.fMask = MIIM_FTYPE | MIIM_ID | MIIM_STATE | MIIM_STRING | MIIM_SUBMENU;
				for(; buttonIndex < buttonCount; ++buttonIndex) {
					ZeroMemory(&button, sizeof(TBBUTTON));
					if(SendMessage(TB_GETBUTTON, buttonIndex, reinterpret_cast<LPARAM>(&button))) {
						menuItem.fState = (button.fsState & TBSTATE_ENABLED ? MFS_ENABLED : MFS_DISABLED);
						if(SendMessage(TB_GETBUTTONINFO, buttonIndex, reinterpret_cast<LPARAM>(&buttonInfo)) == buttonIndex) {
							if(buttonInfo.fsStyle & BTNS_SEP) {
								menuItem.fType = MFT_SEPARATOR;
							} else {
								#ifndef UNICODE
									button.iString = reinterpret_cast<INT_PTR>(buttonInfo.pszText);
								#endif
								menuItem.fType = MFT_STRING;
								menuItem.dwTypeData = reinterpret_cast<LPTSTR>(button.iString);
								CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonIndex, FALSE, this);
								LONG h = NULL;
								Raise_ButtonGetDropDownMenu(pTBarButton, &h);
								menuItem.hSubMenu = static_cast<HMENU>(LongToHandle(h));
							}
						}
						menuItem.wID = button.idCommand;
						InsertMenuItem(properties.hChevronPopupMenu, MAXUINT, TRUE, &menuItem);
					}
				}
				TPMPARAMS params = {0};
				params.cbSize = sizeof(params);
				CWindow reBar = GetParent();
				REBARBANDINFO bandInfo = {RunTimeHelper::SizeOf_REBARBANDINFO(), RBBIM_CHILD | RBBIM_STYLE | RBBIM_CHEVRONLOCATION};
				int bandCount = reBar.SendMessage(RB_GETBANDCOUNT, 0, 0);
				for(int i = 0; i < bandCount; i++) {
					if(reBar.SendMessage(RB_GETBANDINFO, i, reinterpret_cast<LPARAM>(&bandInfo)) && bandInfo.hwndChild == *this) {
						if(bandInfo.fStyle & RBBS_USECHEVRON) {
							params.rcExclude = bandInfo.rcChevronLocation;
							reBar.ClientToScreen(&params.rcExclude);
						}
						break;
					}
				}

				UINT commandID = 0;
				VARIANT_BOOL cancel = VARIANT_FALSE;
				Raise_BeforeDisplayChevronPopup(HandleToLong(properties.hChevronPopupMenu), x, y, VARIANT_TRUE, &cancel, reinterpret_cast<LONG*>(&commandID));
				if(cancel == VARIANT_FALSE) {
					//commandID = static_cast<UINT>(TrackPopupMenuEx(properties.hChevronPopupMenu, TPM_TOPALIGN | TPM_LEFTALIGN | TPM_RETURNCMD, x, y, *this, &params));
					commandID = static_cast<UINT>(DoTrackPopupMenu(properties.hChevronPopupMenu, TPM_LEFTBUTTON | TPM_LEFTALIGN | TPM_TOPALIGN | TPM_RETURNCMD, x, y, &params));
				}
				KillTimer(timers.ID_PRESELECTCHEVRONMENUITEM);
				properties.acceleratorToSendToChevronMenu = 0;
				Raise_DestroyingChevronPopup(HandleToLong(properties.hChevronPopupMenu), VARIANT_TRUE);
				/* DestroyMenu will also destroy any sub-menus. For some reason setting the sub-menus to NULL doesn't help, but removing the
				   menu items before calling DestroyMenu does prevent the sub-menus from being destroyed. */
				while(GetMenuItemCount(properties.hChevronPopupMenu) > 0) {
					RemoveMenu(properties.hChevronPopupMenu, 0, MF_BYPOSITION);
				}
				DestroyMenu(properties.hChevronPopupMenu);
				properties.hChevronPopupMenu = NULL;

				if(commandID > 0) {
					PostMessage(WM_COMMAND, MAKEWPARAM(commandID, 0), 0);
				}
				return S_OK;
			}
		} else {
			// create the tool bar
			DWORD extendedStyle = WS_EX_TOOLWINDOW;
			extendedStyle |= (GetExStyle() & (WS_EX_RIGHT | WS_EX_RTLREADING | WS_EX_LAYOUTRTL));
			DWORD style = WS_CHILD | WS_VISIBLE | CCS_NOPARENTALIGN | CCS_NORESIZE | CCS_NODIVIDER | TBSTYLE_WRAPABLE;
			style |= (GetStyle() & (TBSTYLE_FLAT | TBSTYLE_REGISTERDROP | TBSTYLE_LIST | TBSTYLE_TOOLTIPS));
			chevronPopupToolbar.Create(TOOLBARCLASSNAME, this, 6, *this, 0, NULL, style, extendedStyle);
			if(chevronPopupToolbar.IsWindow()) {
				// setup the popup tool bar
				chevronPopupToolbar.SendMessage(TB_SETPARENT, reinterpret_cast<WPARAM>(static_cast<HWND>(*this)), 0);
				extendedStyle = 0;
				if(RunTimeHelper::IsCommCtrl6()) {
					// TBSTYLE_EX_DOUBLEBUFFER eliminates flicker
					extendedStyle |= TBSTYLE_EX_DOUBLEBUFFER;
				}
				if(!(style & TBSTYLE_LIST) && !properties.alwaysDisplayButtonText) {
					extendedStyle |= TBSTYLE_EX_MIXEDBUTTONS;
				}
				switch(properties.normalDropDownButtonStyle) {
					case nddbsSplitButton:
						extendedStyle |= TBSTYLE_EX_DRAWDDARROWS;
						break;
				}
				chevronPopupToolbar.SendMessage(TB_SETEXTENDEDSTYLE, 0, extendedStyle);
				chevronPopupToolbar.SendMessage(TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);
				if(properties.highlightColor != static_cast<OLE_COLOR>(-1) || properties.shadowColor != static_cast<OLE_COLOR>(-1)) {
					COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
					colorScheme.clrBtnHighlight = (properties.highlightColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.highlightColor));
					colorScheme.clrBtnShadow = (properties.shadowColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.shadowColor));
					chevronPopupToolbar.SendMessage(TB_SETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme));
				}
				chevronPopupToolbar.SendMessage(TB_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));
				chevronPopupToolbar.SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(GetFont()), MAKELPARAM(TRUE, 0));

				if(CTheme::IsThemingSupported() && APIWrapper::IsSupported_SetWindowTheme() && RunTimeHelper::IsCommCtrl6()) {
					// use the same theme
					WCHAR pSubAppNameBuffer[300] = {0};
					LPWSTR pSubAppName = pSubAppNameBuffer;
					WCHAR pSubIDListBuffer[300] = {0};
					LPWSTR pSubIDList = pSubIDListBuffer;
					ATOM valueSubAppName = reinterpret_cast<ATOM>(GetPropW(*this, L"#43281"));
					if(valueSubAppName) {
						GetAtomNameW(valueSubAppName, pSubAppNameBuffer, 300);
						if(lstrlenW(pSubAppNameBuffer) == 1 && pSubAppNameBuffer[0] == L'$') {
							pSubAppNameBuffer[0] = L'\0';
						}
					} else {
						pSubAppName = NULL;
					}
					ATOM valueSubIDList = reinterpret_cast<ATOM>(GetPropW(*this, L"#43280"));
					if(valueSubIDList) {
						GetAtomNameW(valueSubIDList, pSubIDListBuffer, 300);
						if(lstrlenW(pSubIDListBuffer) == 1 && pSubIDListBuffer[0] == L'$') {
							pSubIDListBuffer[0] = L'\0';
						}
					} else {
						pSubIDList = NULL;
					}
					APIWrapper::SetWindowTheme(chevronPopupToolbar, pSubAppName, pSubIDList, NULL);
				}

				// transfer the invisible buttons to the popup tool bar
				int requiredWidth = 0;
				int requiredHeight = 0;
				int popupButtonCount = 0;
				int imageListIndex = -1;
				TBBUTTON button = {0};
				TBBUTTONINFO buttonInfo = {0};
				buttonInfo.cbSize = sizeof(TBBUTTONINFO);
				buttonInfo.dwMask = TBIF_BYINDEX | TBIF_IMAGE | TBIF_STYLE;
				#ifndef UNICODE
					buttonInfo.dwMask |= TBIF_TEXT;
					TCHAR pBuffer[1000];
					buttonInfo.pszText = pBuffer;
					buttonInfo.cchText = 1000;
				#endif
				for(; buttonIndex < buttonCount; ++buttonIndex) {
					ZeroMemory(&button, sizeof(TBBUTTON));
					if(SendMessage(TB_GETBUTTON, buttonIndex, reinterpret_cast<LPARAM>(&button))) {
						/* NOTE: For now we don't add separators, because they are displayed strangely. A vertical
						 *       separator is drawn, followed by a horizontal separator. */
						if(SendMessage(TB_GETBUTTONINFO, buttonIndex, reinterpret_cast<LPARAM>(&buttonInfo)) == buttonIndex && !(buttonInfo.fsStyle & BTNS_SEP) && !IsPlaceholderButton(button.idCommand)) {
							#ifndef UNICODE
								button.iString = reinterpret_cast<INT_PTR>(buttonInfo.pszText);
							#endif
							button.fsStyle &= ~BTNS_AUTOSIZE;
							/*HWND hContainedWindow = NULL;
							HAlignmentConstants horizontalAlignment = halCenter;
							VAlignmentConstants verticalAlignment = valCenter;
							#ifdef USE_STL
								std::unordered_map<LONG, LPPLACEHOLDERBUTTON>::iterator iter = placeholderButtons.find(button.idCommand);
								if(iter != placeholderButtons.end()) {
									if(iter->second) {
										horizontalAlignment = iter->second->horizontalChildWindowAlignment;
										verticalAlignment = iter->second->verticalChildWindowAlignment;
										hContainedWindow = iter->second->hContainedWindow;
									}
								}
							#else
								CAtlMap<LONG, LPPLACEHOLDERBUTTON>::CPair* pEntry = placeholderButtons.Lookup(button.idCommand);
								if(pEntry) {
									if(pEntry->m_value) {
										horizontalAlignment = pEntry->m_value->horizontalChildWindowAlignment;
										verticalAlignment = pEntry->m_value->verticalChildWindowAlignment;
										hContainedWindow = pEntry->m_value->hContainedWindow;
									}
								}
							#endif*/
							if(chevronPopupToolbar.SendMessage(TB_ADDBUTTONS, 1, reinterpret_cast<LPARAM>(&button))) {
								if(!(buttonInfo.fsStyle & BTNS_SEP)) {
									if(HIWORD(buttonInfo.iImage) != imageListIndex) {
										imageListIndex = HIWORD(buttonInfo.iImage);
										// transfer image lists
										chevronPopupToolbar.SendMessage(TB_SETIMAGELIST, imageListIndex, SendMessage(TB_GETIMAGELIST, imageListIndex, 0));
										chevronPopupToolbar.SendMessage(TB_SETHOTIMAGELIST, imageListIndex, SendMessage(TB_GETHOTIMAGELIST, imageListIndex, 0));
										chevronPopupToolbar.SendMessage(TB_SETPRESSEDIMAGELIST, imageListIndex, SendMessage(TB_GETPRESSEDIMAGELIST, imageListIndex, 0));
										chevronPopupToolbar.SendMessage(TB_SETDISABLEDIMAGELIST, imageListIndex, SendMessage(TB_GETDISABLEDIMAGELIST, imageListIndex, 0));
									}
								}
								// update required size
								if(chevronPopupToolbar.SendMessage(TB_GETITEMRECT, popupButtonCount, reinterpret_cast<LPARAM>(&buttonRectangle))) {
									if(buttonInfo.fsStyle & BTNS_SEP) {
										if(buttonRectangle.right - buttonRectangle.left < buttonRectangle.bottom - buttonRectangle.top) {
											// rotate the rectangle
											int tmp = buttonRectangle.bottom;
											buttonRectangle.bottom = buttonRectangle.right;
											buttonRectangle.right = tmp;
											tmp = buttonRectangle.top;
											buttonRectangle.top = buttonRectangle.left;
											buttonRectangle.left = tmp;
										}
									/*} else if(hContainedWindow) {
										::SetParent(hContainedWindow, chevronPopupToolbar);

										CRect windowRectangle;
										::GetWindowRect(hContainedWindow, &windowRectangle);
										CRect targetRectangle;
										switch(horizontalAlignment) {
											case halLeft:
												targetRectangle.left = buttonRectangle.left;
												break;
											case halCenter:
												targetRectangle.left = buttonRectangle.left + ((buttonRectangle.Width() - windowRectangle.Width()) >> 1);
												break;
											case halRight:
												targetRectangle.left = buttonRectangle.right - windowRectangle.Width();
												break;
										}
										targetRectangle.right = targetRectangle.left + windowRectangle.Width();
										switch(verticalAlignment) {
											case valTop:
												targetRectangle.top = buttonRectangle.top;
												break;
											case valCenter:
												targetRectangle.top = buttonRectangle.top + ((buttonRectangle.Height() - windowRectangle.Height()) >> 1);
												break;
											case valBottom:
												targetRectangle.top = buttonRectangle.bottom - windowRectangle.Height();
												break;
										}
										targetRectangle.bottom = targetRectangle.top + windowRectangle.Height();
										::MoveWindow(hContainedWindow, targetRectangle.left, targetRectangle.top, targetRectangle.Width(), targetRectangle.Height(), TRUE);*/
									}
									requiredWidth = max(requiredWidth, buttonRectangle.right - buttonRectangle.left);
									requiredHeight += buttonRectangle.bottom - buttonRectangle.top;
								}
								popupButtonCount++;
							}
						}
					}
				}

				if(parentWindowChevronPopupMenu.IsWindow()) {
					ATLVERIFY(RemoveWindowSubclass(parentWindowChevronPopupMenu, ToolBar::ParentWindowSubclass_ChevronPopupMenu, reinterpret_cast<UINT_PTR>(this)));
					parentWindowChevronPopupMenu.Detach();
				}
				parentWindowChevronPopupMenu.Attach(GetTopLevelParent());
				if(parentWindowChevronPopupMenu.IsWindow()) {
					ATLVERIFY(SetWindowSubclass(parentWindowChevronPopupMenu, ToolBar::ParentWindowSubclass_ChevronPopupMenu, reinterpret_cast<UINT_PTR>(this), NULL));
					// create a menu window that will hold the tool bar
					RECT popupRectangle = {0, 0, requiredWidth + 4, requiredHeight + 4};
					chevronPopupMenuWindow.Create(MAKEINTATOM(0x8000), parentWindowChevronPopupMenu, &popupRectangle, NULL, WS_POPUP | WS_CLIPSIBLINGS, WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_CONTROLPARENT | WS_EX_NOACTIVATE);
					if(chevronPopupMenuWindow.IsWindow()) {
						// make the tool bar a child of the menu window
						chevronPopupToolbar.SetParent(chevronPopupMenuWindow);

						RECT popupWindowRectangle = {x, y, x + requiredWidth + 4, y + requiredHeight + 4};
						LPRECT pExcludeRect = NULL;
						// assume we are in a rebar
						CWindow reBar = GetParent();
						REBARBANDINFO bandInfo = {RunTimeHelper::SizeOf_REBARBANDINFO(), RBBIM_CHILD | RBBIM_STYLE | RBBIM_CHEVRONLOCATION};
						int bandCount = reBar.SendMessage(RB_GETBANDCOUNT, 0, 0);
						for(int i = 0; i < bandCount; i++) {
							if(reBar.SendMessage(RB_GETBANDINFO, i, reinterpret_cast<LPARAM>(&bandInfo)) && bandInfo.hwndChild == *this) {
								if(bandInfo.fStyle & RBBS_USECHEVRON) {
									pExcludeRect = &bandInfo.rcChevronLocation;
									reBar.ClientToScreen(pExcludeRect);
								}
								break;
							}
						}
						POINT anchor = {x, y};
						SIZE windowSize = {requiredWidth + 4, requiredHeight + 4};
						if(APIWrapper::IsSupported_CalculatePopupWindowPosition()) {
							APIWrapper::CalculatePopupWindowPosition(&anchor, &windowSize, TPM_LEFTALIGN | TPM_TOPALIGN, pExcludeRect, &popupWindowRectangle, NULL);
						} else {
							// emulate CalculatePopupWindowPosition
							HMONITOR hMonitor = MonitorFromPoint(anchor, MONITOR_DEFAULTTONULL);
							if(!hMonitor) {
								hMonitor = MonitorFromWindow(*this, MONITOR_DEFAULTTONEAREST);
							}
							if(hMonitor) {
								MONITORINFO monitor = {0};
								monitor.cbSize = sizeof(monitor);
								if(GetMonitorInfo(hMonitor, &monitor)) {
									if(popupWindowRectangle.top < monitor.rcWork.top) {
										popupWindowRectangle.top = monitor.rcWork.top;
										popupWindowRectangle.bottom = popupWindowRectangle.top + windowSize.cy;
									}
									if(popupWindowRectangle.top > monitor.rcWork.bottom - windowSize.cy) {
										// bottom-align with monitor.rcWork.bottom
										popupWindowRectangle.bottom = monitor.rcWork.bottom;
										popupWindowRectangle.top = popupWindowRectangle.bottom - windowSize.cy;
									}
									if(popupWindowRectangle.left < monitor.rcWork.left) {
										popupWindowRectangle.left = monitor.rcWork.left;
										popupWindowRectangle.right = popupWindowRectangle.left + windowSize.cx;
									}
									anchor.x = popupWindowRectangle.left;
									anchor.y = popupWindowRectangle.top;

									if(pExcludeRect) {
										RECT intersection = {0};
										if(IntersectRect(&intersection, pExcludeRect, &popupWindowRectangle)) {
											// move to the right of the exclude rectangle
											popupWindowRectangle.left = pExcludeRect->right;
											popupWindowRectangle.right = popupWindowRectangle.left + windowSize.cx;
										}
									}

									if(popupWindowRectangle.left > monitor.rcWork.right - windowSize.cx) {
										// flip to right-alignment
										popupWindowRectangle.right = popupWindowRectangle.left;
										popupWindowRectangle.left = popupWindowRectangle.right - windowSize.cx;
										if(pExcludeRect) {
											RECT intersection = {0};
											if(IntersectRect(&intersection, pExcludeRect, &popupWindowRectangle)) {
												// move to the left of the exclude rectangle
												popupWindowRectangle.right = pExcludeRect->left;
												popupWindowRectangle.left = popupWindowRectangle.right - windowSize.cx;
											}
										}
									}
								}
							}
						}

						UINT commandID = 0;
						VARIANT_BOOL cancel = VARIANT_FALSE;
						Raise_BeforeDisplayChevronPopup(HandleToLong(static_cast<HWND>(chevronPopupToolbar)), popupWindowRectangle.left, popupWindowRectangle.top, VARIANT_FALSE, &cancel, reinterpret_cast<LONG*>(&commandID));
						if(cancel == VARIANT_FALSE) {
							// display the menu window with the tool bar
							parentWindowChevronPopupMenu.SendMessage(WM_ENTERMENULOOP, TRUE, 0);
							chevronPopupMenuWindow.SetWindowPos(NULL, &popupWindowRectangle, SWP_SHOWWINDOW | SWP_FRAMECHANGED | SWP_NOOWNERZORDER | SWP_NOZORDER | SWP_NOACTIVATE);
							// align the tool bar
							CRect clientRectangle;
							chevronPopupMenuWindow.GetClientRect(&clientRectangle);
							clientRectangle.DeflateRect(2, 2);
							chevronPopupToolbar.MoveWindow(&clientRectangle);
							flags.chevronPopupVisible = TRUE;

							// setup mouse hook
							MouseHookProvider::InstallHook(this);

							// pump messages to keep the chevron pushed
							MSG msg;
							BOOL ret;
							while(flags.chevronPopupVisible && (ret = GetMessage(&msg, NULL, 0, 0)) != 0) {
								if(ret == -1) {
									// some error occured
									flags.chevronPopupVisible = FALSE;
								} else {
									TranslateMessage(&msg);
									DispatchMessage(&msg);
								}
							}
						}
						KillTimer(timers.ID_PRESELECTCHEVRONMENUITEM);
						properties.acceleratorToSendToChevronMenu = 0;

						if(cancel == VARIANT_FALSE) {
							// remove mouse hook
							MouseHookProvider::RemoveHook(this);
						}

						Raise_DestroyingChevronPopup(HandleToLong(static_cast<HWND>(chevronPopupToolbar)), VARIANT_TRUE);
						if(chevronPopupToolbar.IsWindow()) {
							chevronPopupToolbar.DestroyWindow();
						}
						if(chevronPopupMenuWindow.IsWindow()) {
							chevronPopupMenuWindow.DestroyWindow();
						}

						if(cancel != VARIANT_FALSE && commandID > 0) {
							PostMessage(WM_COMMAND, MAKEWPARAM(commandID, 0), 0);
						}
						return S_OK;
					} else {
						chevronPopupToolbar.DestroyWindow();
					}
				} else {
					chevronPopupToolbar.DestroyWindow();
				}
			}
		}
	} else {
		// all buttons are visible
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::FinishOLEDragDrop(void)
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

STDMETHODIMP ToolBar::GetClosestInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, IToolBarButton** ppToolButton, VARIANT_BOOL* pIsOverButton/* = NULL*/)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppToolButton, IToolBarButton*);
	if(!ppToolButton) {
		return E_POINTER;
	}

	if(IsWindow()) {
		POINT pt = {x, y};
		TBINSERTMARK insertMark = {0};
		if(SendMessage(TB_INSERTMARKHITTEST, reinterpret_cast<WPARAM>(&pt), reinterpret_cast<LPARAM>(&insertMark))) {
			// we hit the edge of a button
			ATLASSERT(!(insertMark.dwFlags & TBIMHT_BACKGROUND));
			if(insertMark.iButton == -1) {
				*ppToolButton = NULL;
				*pRelativePosition = impNowhere;
				if(pIsOverButton) {
					*pIsOverButton = VARIANT_FALSE;
				}
			} else {
				if(insertMark.dwFlags & TBIMHT_AFTER) {
					*pRelativePosition = impAfter;
				} else {
					*pRelativePosition = impBefore;
				}
				if(pIsOverButton) {
					*pIsOverButton = VARIANT_TRUE;
				}
				ClassFactory::InitToolBarButton(insertMark.iButton, FALSE, this, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppToolButton), FALSE);
			}
			return S_OK;
		} else {
			/* We are either over a button (then dwFlags is 0) or over the tool bar background (then dwFlags is
			   TBIMHT_BACKGROUND | TBIMHT_AFTER or just TBIMHT_BACKGROUND). If we are over a button, then iButton
			   identifies the button. If we are over the background, then iButton identifies the first or last
			   button in the row. */
			if(insertMark.iButton == -1) {
				// probably won't happen
				*ppToolButton = NULL;
				*pRelativePosition = impNowhere;
				if(pIsOverButton) {
					*pIsOverButton = VARIANT_FALSE;
				}
			} else {
				if(insertMark.dwFlags & TBIMHT_BACKGROUND) {
					// over the background
					if(insertMark.dwFlags & TBIMHT_AFTER) {
						*pRelativePosition = impAfter;
					} else {
						*pRelativePosition = impBefore;
					}
					if(pIsOverButton) {
						*pIsOverButton = VARIANT_FALSE;
					}
				} else {
					/* over a button - We don't want a insertion mark to be displayed, but we want the client app to
					   be able to recognize this case. */
					*pRelativePosition = impNowhere;
					if(pIsOverButton) {
						*pIsOverButton = VARIANT_TRUE;
					}
				}
				ClassFactory::InitToolBarButton(insertMark.iButton, FALSE, this, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppToolButton), FALSE);
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::GetCommandForHotkey(ModifierKeysConstants modifierKeys, SHORT acceleratorKeyCode, LONG* pCommandID)
{
	ATLASSERT_POINTER(pCommandID, LONG);
	if(!pCommandID) {
		return E_POINTER;
	}

	ACCEL newAccelerator = {0};
	newAccelerator.fVirt = FVIRTKEY;
	if(modifierKeys & mkShift) {
		newAccelerator.fVirt |= FSHIFT;
	}
	if(modifierKeys & mkCtrl) {
		newAccelerator.fVirt |= FCONTROL;
	}
	if(modifierKeys & mkAlt) {
		newAccelerator.fVirt |= FALT;
	}
	newAccelerator.key = static_cast<WORD>(acceleratorKeyCode);
	
	ATLASSERT(properties.numberOfAccelerators == 0 || properties.pAccelerators);
	if(properties.numberOfAccelerators > 0 && !properties.pAccelerators) {
		properties.numberOfAccelerators = 0;
	}

	*pCommandID = -1;
	if(properties.pAccelerators) {
		for(UINT i = 0; i < properties.numberOfAccelerators; i++) {
			if(properties.pAccelerators[i].fVirt == newAccelerator.fVirt && properties.pAccelerators[i].key == newAccelerator.key) {
				// found the entry
				*pCommandID = properties.pAccelerators[i].cmd;
				break;
			}
		}
	}
	return S_OK;
}

STDMETHODIMP ToolBar::GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, IToolBarButton** ppToolButton)
{
	ATLASSERT_POINTER(pRelativePosition, InsertMarkPositionConstants);
	if(!pRelativePosition) {
		return E_POINTER;
	}
	ATLASSERT_POINTER(ppToolButton, IToolBarButton*);
	if(!ppToolButton) {
		return E_POINTER;
	}

	if(IsWindow()) {
		TBINSERTMARK insertMark = {0};
		if(SendMessage(TB_GETINSERTMARK, 0, reinterpret_cast<LPARAM>(&insertMark))) {
			if(insertMark.iButton == -1) {
				*ppToolButton = NULL;
				*pRelativePosition = impNowhere;
			} else {
				if(insertMark.dwFlags & TBIMHT_AFTER) {
					*pRelativePosition = impAfter;
				} else {
					*pRelativePosition = impBefore;
				}
				ClassFactory::InitToolBarButton(insertMark.iButton, FALSE, this, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppToolButton), FALSE);
			}
			return S_OK;
		}
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IToolBarButton** ppHitButton)
{
	ATLASSERT_POINTER(ppHitButton, IToolBarButton*);
	if(!ppHitButton) {
		return E_POINTER;
	}

	if(IsWindow()) {
		UINT flags = static_cast<UINT>(*pHitTestDetails);
		int buttonIndex = HitTest(x, y, &flags, *this/*, TRUE*/);

		if(pHitTestDetails) {
			*pHitTestDetails = static_cast<HitTestConstants>(flags);
		}
		ClassFactory::InitToolBarButton(buttonIndex, FALSE, this, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppHitButton));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::HitTestChevronToolBar(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IToolBarButton** ppHitButton)
{
	ATLASSERT_POINTER(ppHitButton, IToolBarButton*);
	if(!ppHitButton) {
		return E_POINTER;
	}

	if(chevronPopupToolbar.IsWindow()) {
		UINT flags = static_cast<UINT>(*pHitTestDetails);
		int buttonIndex = HitTest(x, y, &flags, chevronPopupToolbar/*, TRUE*/);

		if(pHitTestDetails) {
			*pHitTestDetails = static_cast<HitTestConstants>(flags);
		}
		ClassFactory::InitToolBarButton(buttonIndex, TRUE, this, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppHitButton));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::LoadDefaultImages(SystemImageListTypeConstants imageListType, LONG* pImageCount)
{
	ATLASSERT_POINTER(pImageCount, LONG);
	if(!pImageCount) {
		return E_POINTER;
	}
	*pImageCount = 0;

	if(IsWindow()) {
		if(imageListType >= 0) {
			*pImageCount = SendMessage(TB_LOADIMAGES, imageListType, reinterpret_cast<LPARAM>(HINST_COMMCTRL));
		} else {
			HIMAGELIST hImageList = NULL;
			switch(imageListType) {
				case siltShellIcons16x16:
					if(APIWrapper::IsSupported_SHGetImageList()) {
						APIWrapper::SHGetImageList(SHIL_SMALL, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
					} else {
						SHFILEINFO details = {0};
						hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES));
					}
					break;
				case siltShellIcons32x32:
					if(APIWrapper::IsSupported_SHGetImageList()) {
						APIWrapper::SHGetImageList(SHIL_LARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
					} else {
						SHFILEINFO details = {0};
						hImageList = reinterpret_cast<HIMAGELIST>(SHGetFileInfo(TEXT("txt"), FILE_ATTRIBUTE_NORMAL, &details, sizeof(SHFILEINFO), SHGFI_SYSICONINDEX | SHGFI_LARGEICON | SHGFI_USEFILEATTRIBUTES));
					}
					break;
				case siltShellIcons48x48:
					if(APIWrapper::IsSupported_SHGetImageList()) {
						APIWrapper::SHGetImageList(SHIL_EXTRALARGE, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
					}
					break;
				case siltShellIcons256x256:
					if(APIWrapper::IsSupported_SHGetImageList()) {
						APIWrapper::SHGetImageList(SHIL_JUMBO, IID_IImageList, reinterpret_cast<LPVOID*>(&hImageList), NULL);
					}
					break;
			}
			if(hImageList) {
				SendMessage(TB_SETIMAGELIST, 0, reinterpret_cast<LPARAM>(hImageList));
				*pImageCount = ImageList_GetImageCount(hImageList);
			}
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP ToolBar::LoadToolBarStateFromRegistry(BSTR subKeyName, BSTR valueName, OLE_HANDLE hParentKey/* = 0x80000001*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	if(IsWindow()) {
		TBSAVEPARAMS params = {0};
		params.hkr = reinterpret_cast<HKEY>(HandleToLong(hParentKey));
		COLE2T converter1(subKeyName);
		params.pszSubKey = converter1;
		COLE2T converter2(valueName);
		params.pszValueName = converter2;
		*pSucceeded = BOOL2VARIANTBOOL(SendMessage(TB_SAVERESTORE, FALSE, reinterpret_cast<LPARAM>(&params)));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::OLEDrag(LONG* pDataObject/* = NULL*/, OLEDropEffectConstants supportedEffects/* = odeCopyOrMove*/, OLE_HANDLE hWndToAskForDragImage/* = -1*/, IToolBarButtonContainer* pDraggedButtons/* = NULL*/, LONG itemCountToDisplay/* = -1*/, OLEDropEffectConstants* pPerformedEffects/* = NULL*/)
{
	if(supportedEffects == odeNone) {
		// don't waste time
		return S_OK;
	}
	if(hWndToAskForDragImage == -1) {
		ATLASSUME(pDraggedButtons);
		if(!pDraggedButtons) {
			return E_INVALIDARG;
		}
	}
	ATLASSERT_POINTER(pDataObject, LONG);
	LONG dummy = NULL;
	if(!pDataObject) {
		pDataObject = &dummy;
	}
	ATLASSERT_POINTER(pPerformedEffects, OLEDropEffectConstants);
	OLEDropEffectConstants performedEffect = odeNone;
	if(!pPerformedEffects) {
		pPerformedEffects = &performedEffect;
	}

	HWND hWnd = NULL;
	if(hWndToAskForDragImage == -1) {
		hWnd = *this;
	} else {
		hWnd = static_cast<HWND>(LongToHandle(hWndToAskForDragImage));
	}

	CComPtr<IOLEDataObject> pOLEDataObject = NULL;
	CComPtr<IDataObject> pDataObjectToUse = NULL;
	if(*pDataObject) {
		pDataObjectToUse = *reinterpret_cast<LPDATAOBJECT*>(pDataObject);

		CComObject<TargetOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<TargetOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->Attach(pDataObjectToUse);
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->Release();
	} else {
		CComObject<SourceOLEDataObject>* pOLEDataObjectObj = NULL;
		CComObject<SourceOLEDataObject>::CreateInstance(&pOLEDataObjectObj);
		pOLEDataObjectObj->AddRef();
		pOLEDataObjectObj->SetOwner(this);
		if(itemCountToDisplay == -1) {
			if(pDraggedButtons) {
				if(flags.usingThemes && RunTimeHelper::IsVista()) {
					pDraggedButtons->Count(&pOLEDataObjectObj->properties.numberOfItemsToDisplay);
				}
			}
		} else if(itemCountToDisplay >= 0) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				pOLEDataObjectObj->properties.numberOfItemsToDisplay = itemCountToDisplay;
			}
		}
		pOLEDataObjectObj->QueryInterface(IID_IOLEDataObject, reinterpret_cast<LPVOID*>(&pOLEDataObject));
		pOLEDataObjectObj->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&pDataObjectToUse));
		pOLEDataObjectObj->Release();
	}
	ATLASSUME(pDataObjectToUse);
	pDataObjectToUse->QueryInterface(IID_IDataObject, reinterpret_cast<LPVOID*>(&dragDropStatus.pSourceDataObject));
	CComQIPtr<IDropSource, &IID_IDropSource> pDragSource(this);

	if(pDraggedButtons) {
		pDraggedButtons->Clone(&dragDropStatus.pDraggedButtons);
	}
	POINT mousePosition = {0};
	GetCursorPos(&mousePosition);
	ScreenToClient(&mousePosition);

	if(properties.supportOLEDragImages) {
		IDragSourceHelper* pDragSourceHelper = NULL;
		CoCreateInstance(CLSID_DragDropHelper, NULL, CLSCTX_ALL, IID_PPV_ARGS(&pDragSourceHelper));
		if(pDragSourceHelper) {
			if(flags.usingThemes && RunTimeHelper::IsVista()) {
				IDragSourceHelper2* pDragSourceHelper2 = NULL;
				pDragSourceHelper->QueryInterface(IID_IDragSourceHelper2, reinterpret_cast<LPVOID*>(&pDragSourceHelper2));
				if(pDragSourceHelper2) {
					pDragSourceHelper2->SetFlags(DSH_ALLOWDROPDESCRIPTIONTEXT);
					// this was the only place we actually use IDragSourceHelper2
					pDragSourceHelper->Release();
					pDragSourceHelper = static_cast<IDragSourceHelper*>(pDragSourceHelper2);
				}
			}

			HRESULT hr = pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
			if(FAILED(hr)) {
				/* This happens if full window dragging is deactivated. Actually, InitializeFromWindow() contains a
				   fallback mechanism for this case. This mechanism retrieves the passed window's class name and
				   builds the drag image using TVM_CREATEDRAGIMAGE if it's SysTreeView32, LVM_CREATEDRAGIMAGE if
				   it's SysListView32 and so on. Our class name is ExplorerListView[U|A], so we're doomed.
				   So how can we have drag images anyway? Well, we use a very ugly hack: We temporarily activate
				   full window dragging. */
				BOOL fullWindowDragging;
				SystemParametersInfo(SPI_GETDRAGFULLWINDOWS, 0, &fullWindowDragging, 0);
				if(!fullWindowDragging) {
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, TRUE, NULL, 0);
					pDragSourceHelper->InitializeFromWindow(hWnd, &mousePosition, pDataObjectToUse);
					SystemParametersInfo(SPI_SETDRAGFULLWINDOWS, FALSE, NULL, 0);
				}
			}

			if(pDragSourceHelper) {
				pDragSourceHelper->Release();
			}
		}
	}

	if(IsLeftMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_LBUTTON;
	} else if(IsRightMouseButtonDown()) {
		dragDropStatus.draggingMouseButton = MK_RBUTTON;
	}
	if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
		dragDropStatus.useItemCountLabelHack = TRUE;
	}

	if(pOLEDataObject) {
		Raise_OLEStartDrag(pOLEDataObject);
	}
	HRESULT hr = DoDragDrop(pDataObjectToUse, pDragSource, supportedEffects, reinterpret_cast<LPDWORD>(pPerformedEffects));
	if((hr == DRAGDROP_S_DROP) && pOLEDataObject) {
		Raise_OLECompleteDrag(pOLEDataObject, *pPerformedEffects);
	}

	dragDropStatus.Reset();
	return S_OK;
}

STDMETHODIMP ToolBar::Refresh(void)
{
	if(IsWindow()) {
		Invalidate();
		RedrawWindow(NULL, NULL, RDW_FRAME | RDW_INVALIDATE | RDW_UPDATENOW | RDW_ERASE);
	}
	return S_OK;
}

STDMETHODIMP ToolBar::RegisterHotkey(ModifierKeysConstants modifierKeys, SHORT acceleratorKeyCode, LONG commandID)
{
	ACCEL newAccelerator = {0};
	newAccelerator.cmd = static_cast<WORD>(commandID);
	newAccelerator.fVirt = FVIRTKEY;
	if(modifierKeys & mkShift) {
		newAccelerator.fVirt |= FSHIFT;
	}
	if(modifierKeys & mkCtrl) {
		newAccelerator.fVirt |= FCONTROL;
	}
	if(modifierKeys & mkAlt) {
		newAccelerator.fVirt |= FALT;
	}
	newAccelerator.key = static_cast<WORD>(acceleratorKeyCode);
	
	ATLASSERT(properties.numberOfAccelerators == 0 || properties.pAccelerators);
	if(properties.numberOfAccelerators > 0 && !properties.pAccelerators) {
		properties.numberOfAccelerators = 0;
	}

	if(properties.pAccelerators) {
		for(UINT i = 0; i < properties.numberOfAccelerators; i++) {
			if(properties.pAccelerators[i].fVirt == newAccelerator.fVirt && properties.pAccelerators[i].key == newAccelerator.key) {
				// this entry already exists
				properties.pAccelerators[i].cmd = newAccelerator.cmd;
				return S_OK;
			}
		}
	}

	LPACCEL pNewAccelerators = static_cast<LPACCEL>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, (properties.numberOfAccelerators + 1) * sizeof(ACCEL)));
	if(properties.pAccelerators) {
		CopyMemory(pNewAccelerators, properties.pAccelerators, properties.numberOfAccelerators * sizeof(ACCEL));
	}
	// append the new accelerator
	CopyMemory(pNewAccelerators + properties.numberOfAccelerators, &newAccelerator, sizeof(ACCEL));
	LPACCEL pOldAccelerators = properties.pAccelerators;
	properties.pAccelerators = pNewAccelerators;
	properties.numberOfAccelerators++;
	if(pOldAccelerators) {
		HeapFree(GetProcessHeap(), 0, pOldAccelerators);
		pOldAccelerators = NULL;
	}

	/*if(properties.hAcceleratorTable) {
		DestroyAcceleratorTable(properties.hAcceleratorTable);
	}
	properties.hAcceleratorTable = CreateAcceleratorTable(properties.pAccelerators, properties.numberOfAccelerators);
	if(!properties.hAcceleratorTable) {
		HeapFree(GetProcessHeap(), 0, properties.pAccelerators);
		properties.pAccelerators = NULL;
		properties.numberOfAccelerators = 0;
	}
	return (properties.hAcceleratorTable ? S_OK : E_FAIL);*/
	return S_OK;
}

STDMETHODIMP ToolBar::SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded)
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

STDMETHODIMP ToolBar::SaveToolBarStateToRegistry(BSTR subKeyName, BSTR valueName, OLE_HANDLE hParentKey/* = 0x80000001*/, VARIANT_BOOL* pSucceeded/* = NULL*/)
{
	if(!pSucceeded) {
		return E_POINTER;
	}
	*pSucceeded = VARIANT_FALSE;

	if(IsWindow()) {
		TBSAVEPARAMS params = {0};
		params.hkr = reinterpret_cast<HKEY>(HandleToLong(hParentKey));
		COLE2T converter1(subKeyName);
		params.pszSubKey = converter1;
		COLE2T converter2(valueName);
		params.pszValueName = converter2;
		*pSucceeded = BOOL2VARIANTBOOL(SendMessage(TB_SAVERESTORE, TRUE, reinterpret_cast<LPARAM>(&params)));
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::SetButtonRowCount(LONG rowCount, VARIANT_BOOL allowMoreRows/* = VARIANT_TRUE*/, OLE_XSIZE_PIXELS* pControlWidth/* = NULL*/, OLE_YSIZE_PIXELS* pControlHeight/* = NULL*/)
{
	if(rowCount < 1) {
		// invalid procedure call or argument - raise VB runtime error 5
		return MAKE_HRESULT(1, FACILITY_WINDOWS | FACILITY_DISPATCH, 5);
	}

	if(IsWindow()) {
		RECT clientRectangle = {0};
		SendMessage(TB_SETROWS, MAKEWPARAM(rowCount, VARIANTBOOL2BOOL(allowMoreRows)), reinterpret_cast<LPARAM>(&clientRectangle));
		if(pControlWidth) {
			*pControlWidth = clientRectangle.right - clientRectangle.left;
		}
		if(pControlHeight) {
			*pControlHeight = clientRectangle.bottom - clientRectangle.top;
		}
		return S_OK;
	}
	return E_FAIL;
}

STDMETHODIMP ToolBar::SetHotButton(IToolBarButton* pNewHotButton, HotButtonChangingCausedByConstants hotButtonChangeReason/* = hbccbOther*/, HotButtonChangingAdditionalInfoConstants additionalInfo/* = static_cast<HotButtonChangingAdditionalInfoConstants>(0)*/, IToolBarButton** ppPreviousHotButton/* = NULL*/)
{
	ATLASSERT_POINTER(ppPreviousHotButton, IToolBarButton*);
	if(!ppPreviousHotButton) {
		return E_POINTER;
	}

	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		int buttonIndex = -1;
		if(pNewHotButton) {
			LONG l = -1;
			pNewHotButton->get_Index(&l);
			buttonIndex = l;
		}

		LPARAM flags = (hotButtonChangeReason & (HICF_OTHER | HICF_MOUSE | HICF_ARROWKEYS | HICF_ACCELERATOR));
		flags |= (additionalInfo & (HICF_DUPACCEL | HICF_ENTERING | HICF_LEAVING | HICF_RESELECT | HICF_LMOUSE | HICF_TOGGLEDROPDOWN));
		buttonIndex = SendMessage(TB_SETHOTITEM2, buttonIndex, flags);
		ClassFactory::InitToolBarButton(buttonIndex, FALSE, this, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(ppPreviousHotButton));
		hr = S_OK;
	}

	return hr;
}

STDMETHODIMP ToolBar::SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, IToolBarButton* pToolButton)
{
	int buttonIndex = -1;
	if(pToolButton) {
		LONG l = -1;
		pToolButton->get_Index(&l);
		buttonIndex = l;
	}

	HRESULT hr = E_FAIL;
	if(IsWindow()) {
		TBINSERTMARK insertMark = {0};
		insertMark.iButton = buttonIndex;
		switch(relativePosition) {
			case impNowhere:
				insertMark.iButton = -1;
				break;
			case impAfter:
				insertMark.dwFlags |= TBIMHT_AFTER;
				break;
			case impDontChange:
				return S_OK;
				break;
		}

		SendMessage(TB_SETINSERTMARK, 0, reinterpret_cast<LPARAM>(&insertMark));
		hr = S_OK;
	}

	return hr;
}

STDMETHODIMP ToolBar::UnregisterHotkey(ModifierKeysConstants modifierKeys, SHORT acceleratorKeyCode)
{
	ACCEL newAccelerator = {0};
	newAccelerator.fVirt = FVIRTKEY;
	if(modifierKeys & mkShift) {
		newAccelerator.fVirt |= FSHIFT;
	}
	if(modifierKeys & mkCtrl) {
		newAccelerator.fVirt |= FCONTROL;
	}
	if(modifierKeys & mkAlt) {
		newAccelerator.fVirt |= FALT;
	}
	newAccelerator.key = static_cast<WORD>(acceleratorKeyCode);
	
	ATLASSERT(properties.numberOfAccelerators == 0 || properties.pAccelerators);
	if(properties.numberOfAccelerators > 0 && !properties.pAccelerators) {
		properties.numberOfAccelerators = 0;
	}

	if(properties.pAccelerators) {
		for(UINT i = 0; i < properties.numberOfAccelerators; i++) {
			if(properties.pAccelerators[i].fVirt == newAccelerator.fVirt && properties.pAccelerators[i].key == newAccelerator.key) {
				// found the entry - remove it
				if(properties.numberOfAccelerators > 1) {
					LPACCEL pNewAccelerators = static_cast<LPACCEL>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, (properties.numberOfAccelerators - 1) * sizeof(ACCEL)));
					if(properties.pAccelerators) {
						if(i > 0) {
							CopyMemory(pNewAccelerators, properties.pAccelerators, i * sizeof(ACCEL));
						}
						if(i < properties.numberOfAccelerators - 1) {
							CopyMemory(pNewAccelerators + i, properties.pAccelerators + i + 1, (properties.numberOfAccelerators - i - 1) * sizeof(ACCEL));
						}
						LPACCEL pOldAccelerators = properties.pAccelerators;
						properties.pAccelerators = pNewAccelerators;
						properties.numberOfAccelerators--;
						if(pOldAccelerators) {
							HeapFree(GetProcessHeap(), 0, pOldAccelerators);
							pOldAccelerators = NULL;
						}
					}
				} else {
					properties.numberOfAccelerators = 0;
					HeapFree(GetProcessHeap(), 0, properties.pAccelerators);
					properties.pAccelerators = NULL;
				}
				break;
			}
		}
	}
	return S_OK;
}


LRESULT ToolBar::OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyAscii = static_cast<SHORT>(wParam);
		if(SUCCEEDED(Raise_KeyPress(&keyAscii))) {
			// the client may have changed the key code (actually it can be changed to 0 only)
			wParam = keyAscii;
			if(wParam == 0) {
				wasHandled = TRUE;
			}
		}
	}
	if(properties.menuMode) {
		if(wParam == VK_SPACE) {
			wasHandled = TRUE;
			return 0;
		}
		if(!parentWindowMenuMode.IsWindowEnabled() || ::GetFocus() != *this) {
			return 0;
		}

		// Handle mnemonic press when we have focus
		int buttonIndex = 0;
		if(wParam != VK_RETURN && !SendMessage(TB_MAPACCELERATOR, static_cast<WPARAM>(LOWORD(wParam)), reinterpret_cast<LPARAM>(&buttonIndex))) {
			if(LOWORD(wParam) != TEXT('/')) {
				MessageBeep(MB_OK);
			}
		} else {
			RECT clientRectangle = {0};
			GetClientRect(&clientRectangle);
			RECT buttonRectangle = {0};
			SendMessage(TB_GETITEMRECT, buttonIndex, reinterpret_cast<LPARAM>(&buttonRectangle));
			TBBUTTON button = {0};
			SendMessage(TB_GETBUTTON, buttonIndex, reinterpret_cast<LPARAM>(&button));
			if((button.fsState & TBSTATE_ENABLED) && !(button.fsState & TBSTATE_HIDDEN) && buttonRectangle.right <= clientRectangle.right) {
				PostMessage(WM_KEYDOWN, VK_DOWN, 0);
				if(wParam != VK_RETURN) {
					SendMessage(TB_SETHOTITEM, buttonIndex, 0);
				}
			} else {
				MessageBeep(MB_OK);
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT ToolBar::OnCommand(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(flags.chevronPopupVisible && reinterpret_cast<HWND>(lParam) == chevronPopupToolbar && parentWindowChevronPopupMenu.IsWindow()) {
		CWindow parentWindow = parentWindowChevronPopupMenu;
		parentWindow.SendMessage(WM_CANCELMODE, 0, 0);
		parentWindow.SendMessage(WM_EXITMENULOOP, 0, 0);
		parentWindow = GetParent();
		if(parentWindow.IsWindow()) {
			parentWindow.PostMessage(message, wParam, reinterpret_cast<LPARAM>(static_cast<HWND>(*this)));
		}
	} else if(HIWORD(wParam) == 0) {
		VARIANT_BOOL forward = VARIANT_TRUE;
		Raise_ExecuteCommand(LOWORD(wParam), NULL, coMenu, &forward, NULL);
		if(forward != VARIANT_FALSE && IsWindow()) {
			CWindow parent = GetTopLevelParent();
			if(parent.IsWindow()) {
				parent.SendMessage(message, wParam, lParam);
			}
		}
	} else {
		wasHandled = FALSE;
	}
	return 0;
}

LRESULT ToolBar::OnContextMenu(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	POINT mousePosition = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
	if((mousePosition.x != -1) && (mousePosition.y != -1)) {
		ScreenToClient(&mousePosition);
	}

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	Raise_ContextMenu(index, button, shift, mousePosition.x, mousePosition.y);
	return 0;
}

LRESULT ToolBar::OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);

	if(*this) {
		// this will keep the object alive if the client destroys the control window in an event handler
		AddRef();
		GetSystemSettings();

		Raise_RecreatedControlWindow(HandleToLong(static_cast<HWND>(*this)));
	}
	return lr;
}

LRESULT ToolBar::OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	KeyboardHookProvider::RemoveHook(this);

	Raise_DestroyedControlWindow(HandleToLong(static_cast<HWND>(*this)));

	wasHandled = FALSE;
	return 0;
}

LRESULT ToolBar::OnForwardMsg(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = 0;
	HWND hWndParent = GetParent();
	if(hWndParent) {
		LONG controlType = 0x02;
		BOOL wantsForwardMsg = FALSE;
		HWND hWndControl = reinterpret_cast<HWND>(::SendMessage(hWndParent, GetBarMessage(), reinterpret_cast<WPARAM>(&controlType), reinterpret_cast<LPARAM>(&wantsForwardMsg)));
		if(hWndControl && wantsForwardMsg && controlType == 0x01) {
			// our parent is a rebar control which also wants to process this message
			lr = ::SendMessage(hWndControl, message, wParam, lParam);
		}
	}
	if(properties.menuMode) {
		LPMSG pMessage = reinterpret_cast<LPMSG>(lParam);
		menuModeState.hWndForwardedMessage = pMessage->hwnd;
		if(pMessage->message >= WM_MOUSEFIRST && pMessage->message <= WM_MOUSELAST) {
			menuModeState.keyboardInput = FALSE;
		} else if(pMessage->message >= WM_KEYFIRST && pMessage->message <= WM_KEYLAST) {
			menuModeState.keyboardInput = TRUE;
		}
		ProcessWindowMessage(pMessage->hwnd, pMessage->message, pMessage->wParam, pMessage->lParam, lr, 4);
	}
	return lr;
}

LRESULT ToolBar::OnInitMenuPopup(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = 0;
	VARIANT_BOOL handled = VARIANT_FALSE;
	if(reinterpret_cast<HMENU>(wParam) != properties.hChevronPopupMenu) {
		if(!(properties.disabledEvents & deRawMenuMessage)) {
			LONG result = lr;
			Raise_RawMenuMessage(message, wParam, lParam, &result, &handled);
			lr = result;
		}

		LONG menuItemIndex = LOWORD(lParam);
		if(menuItemIndex == 0 && !menuModeState.selectedMenuItemHasSubMenu) {
			menuItemIndex = SendMessage(TB_GETHOTITEM, 0, 0);
		}
		Raise_OpenPopupMenu(HandleToLong(reinterpret_cast<HMENU>(wParam)), menuItemIndex, BOOL2VARIANTBOOL(HIWORD(lParam)));
	}

	wasHandled = VARIANTBOOL2BOOL(handled);
	return lr;
}

LRESULT ToolBar::OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyDown(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	if(properties.menuMode) {
		// simulate Alt+Space for the parent
		if(wParam == VK_SPACE) {
			parentWindowMenuMode.PostMessage(WM_SYSKEYDOWN, wParam, lParam | (1 << 29));
			return 0;
		} else if(wParam == VK_LEFT || wParam == VK_RIGHT) {
			DWORD styleEx = GetExStyle();
			WPARAM keyCodeNext = (styleEx & WS_EX_LAYOUTRTL) ? VK_LEFT : VK_RIGHT;

			if(!menuModeState.menuIsActive) {
				int hotButton = SendMessage(TB_GETHOTITEM, 0, 0);
				int nextButton = (wParam == keyCodeNext ? GetNextMenuItem(hotButton) : GetPreviousMenuItem(hotButton));
				if(nextButton == -2) {
					SendMessage(TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
					if(DisplayChevronMenu()) {
						return 0;
					}
				}
			}
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ToolBar::OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(!(properties.disabledEvents & deKeyboardEvents)) {
		SHORT keyCode = static_cast<SHORT>(wParam);
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		if(SUCCEEDED(Raise_KeyUp(&keyCode, shift))) {
			// the client may have changed the key code
			wParam = keyCode;
			if(wParam == 0) {
				return 0;
			}
		}
	}
	if(properties.menuMode) {
		if(wParam != VK_SPACE) {
			return DefWindowProc(message, wParam, lParam);
		} else {
			return 0;
		}
	}
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ToolBar::OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = CComControl<ToolBar>::OnKillFocus(message, wParam, lParam, wasHandled);
	flags.uiActivationPending = FALSE;
	if(properties.menuMode) {
		if(menuModeState.useContextSensitiveKeyboardCues && menuModeState.displayKeyboardCues) {
			ShowKeyboardCues(FALSE);
		}
	}
	return lr;
}

LRESULT ToolBar::OnLButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_DblClick(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ToolBar::OnLButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.focusOnClick || flags.allowSetFocus) {
		/* HACK: Fixes the Validate event of VB6. ToolbarWindow32 doesn't seem to get the focus on
		 * WM_MOUSEACTIVATE by default. If we call SetFocus on WM_MOUSEACTIVATE, the Validate event is not fired
		 * for some reason. Therefore we call SetFocus here.
		 */
		SetFocus();
	}
	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 1/*MouseButtonConstants.vbLeftButton*/;
		Raise_MouseDown(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		mouseStatus.StoreClickCandidate(1/*MouseButtonConstants.vbLeftButton*/);
	}

	return 0;
}

LRESULT ToolBar::OnLButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 1/*MouseButtonConstants.vbLeftButton*/;
	Raise_MouseUp(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ToolBar::OnMButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MDblClick(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ToolBar::OnMButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.focusOnClick || flags.allowSetFocus) {
		/* HACK: Fixes the Validate event of VB6. ToolbarWindow32 doesn't seem to get the focus on
		 * WM_MOUSEACTIVATE by default. If we call SetFocus on WM_MOUSEACTIVATE, the Validate event is not fired
		 * for some reason. Therefore we call SetFocus here.
		 */
		SetFocus();
	}
	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 4/*MouseButtonConstants.vbMiddleButton*/;
		Raise_MouseDown(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		mouseStatus.StoreClickCandidate(4/*MouseButtonConstants.vbMiddleButton*/);
	}

	return 0;
}

LRESULT ToolBar::OnMButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 4/*MouseButtonConstants.vbMiddleButton*/;
	Raise_MouseUp(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ToolBar::OnMenuChar(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = TRUE;

	LRESULT lr;
	if(menuModeState.menuIsActive && LOWORD(wParam) != 0x0D) {
		lr = 0;
	} else {
		lr = MAKELRESULT(1, 1);
	}

	if(menuModeState.menuIsActive && HIWORD(wParam) == MF_POPUP) {
		VARIANT_BOOL handled = VARIANT_FALSE;
		if(!(properties.disabledEvents & deRawMenuMessage)) {
			LONG result = lr;
			Raise_RawMenuMessage(message, wParam, lParam, &result, &handled);
			lr = result;
		}
		int returnCode = MNC_EXECUTE;
		WORD mnemonic = 0;
		BOOL found = FALSE;
		if(handled == VARIANT_FALSE) {
			// convert character to lower/uppercase and possibly Unicode, using current keyboard layout
			TCHAR pressedChar = static_cast<TCHAR>(LOWORD(wParam));
			CMenuHandle menu = reinterpret_cast<HMENU>(lParam);
			int menuItemCount = menu.GetMenuItemCount();
			BOOL ret = FALSE;
			TCHAR pMenuItemText[1024];
			for(int i = 0; i < menuItemCount; i++) {
				CMenuItemInfo menuItemInfo;
				menuItemInfo.cch = 1024;
				menuItemInfo.fMask = MIIM_CHECKMARKS | MIIM_DATA | MIIM_ID | MIIM_STATE | MIIM_SUBMENU | MIIM_TYPE;
				menuItemInfo.dwTypeData = pMenuItemText;
				ret = menu.GetMenuItemInfo(i, TRUE, &menuItemInfo);
				if(!ret || (menuItemInfo.fType & MFT_SEPARATOR)) {
					continue;
				}
				if(menuItemInfo.fType & (MFT_OWNERDRAW | MFT_BITMAP)) {
					// TODO: Ask the client app for the menu item text.
				}
				LPTSTR p = pMenuItemText;
				if(p) {
					while(*p && *p != TEXT('&')) {
						p = CharNext(p);
					}
					if(p && *p) {
						if(CharLower(reinterpret_cast<LPTSTR>(ULongToPtr(MAKELONG(*(++p), 0)))) == CharLower(reinterpret_cast<LPTSTR>(ULongToPtr(MAKELONG(pressedChar, 0))))) {
							if(!found) {
								mnemonic = static_cast<WORD>(i);
								found = TRUE;
							} else {
								returnCode = MNC_SELECT;
								break;
							}
						}
					}
				}
			}
		} else {
			found = TRUE;
			mnemonic = LOWORD(lr);
			returnCode = HIWORD(lr);
		}

		if(found) {
			if(returnCode == MNC_EXECUTE) {
				PostMessage(TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
				GiveFocusBack();
			}
			wasHandled = TRUE;
			lr = MAKELRESULT(mnemonic, returnCode);
		}
	} else if(!menuModeState.menuIsActive) {
		int buttonIndex = 0;
		if(!SendMessage(TB_MAPACCELERATOR, static_cast<WPARAM>(LOWORD(wParam)), reinterpret_cast<LPARAM>(&buttonIndex))) {
			wasHandled = FALSE;
			PostMessage(TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
			GiveFocusBack();

			// check if we should display chevron menu
			if(static_cast<TCHAR>(LOWORD(wParam)) == TEXT('/')) {
				if(DisplayChevronMenu()) {
					wasHandled = TRUE;
				}
			}
		} else if(parentWindowMenuMode.IsWindowEnabled()) {
			RECT clientRectangle = {0};
			GetClientRect(&clientRectangle);
			RECT buttonRectangle = {0};
			SendMessage(TB_GETITEMRECT, buttonIndex, reinterpret_cast<LPARAM>(&buttonRectangle));
			TBBUTTON button = {0};
			SendMessage(TB_GETBUTTON, buttonIndex, reinterpret_cast<LPARAM>(&button));
			if((button.fsState & TBSTATE_ENABLED) && !(button.fsState & TBSTATE_HIDDEN) && buttonRectangle.right <= clientRectangle.right) {
				if(menuModeState.useContextSensitiveKeyboardCues && !menuModeState.displayKeyboardCues) {
					menuModeState.allowDisplayingKeyboardCues = TRUE;
					ShowKeyboardCues(TRUE);
				}
				TakeFocus();
				PostMessage(WM_KEYDOWN, VK_DOWN, 0);
				SendMessage(TB_SETHOTITEM, buttonIndex, 0);
			} else {
				MessageBeep(0);
			}
		}
	}
	return lr;
}

LRESULT ToolBar::OnMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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

LRESULT ToolBar::OnMenuSelect(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	LRESULT lr = 1;
	VARIANT_BOOL handled = VARIANT_FALSE;
	if(reinterpret_cast<HMENU>(lParam) != properties.hChevronPopupMenu || !IsMenu(properties.hChevronPopupMenu)) {
		if(!(properties.disabledEvents & deRawMenuMessage)) {
			LONG result = lr;
			Raise_RawMenuMessage(message, wParam, lParam, &result, &handled);
			lr = result;
		}

		// NOTE: This is what WTL does (properties.menuMode corresponds to m_bAttachedMenu, which means "We've been attached to an existing menu")
		/*if(!properties.menuMode) {
			menuModeState.selectedMenuItemHasSubMenu = lParam && reinterpret_cast<HMENU>(lParam) != m_hMenu && (HIWORD(wParam) & MF_POPUP);
			if(parentWindowMenuMode.IsWindow()) {
				parentWindowMenuMode.SendMessage(message, wParam, lParam);
			}
			wasHandled = VARIANTBOOL2BOOL(handled);
			return lr;
		}*/
		// The above does not make much sense. To make keyboard navigation work for items with sub-menus we do something different.
		if(properties.menuMode) {
			menuModeState.selectedMenuItemHasSubMenu = lParam && (HIWORD(wParam) & MF_POPUP);
		}

		Raise_SelectedMenuItem(LOWORD(wParam), static_cast<MenuItemStateConstants>(HIWORD(wParam)), HandleToLong(reinterpret_cast<HMENU>(lParam)));

		// check if a menu is closing, do a cleanup
		if(HIWORD(wParam) == 0xFFFF && !lParam) {
			// menu closing
			ATLASSERT(menuModeState.menuWindowStack.GetSize() == 0);
		}
	}
	wasHandled = VARIANTBOOL2BOOL(handled);
	return lr;
}

LRESULT ToolBar::OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(properties.focusOnClick || flags.allowSetFocus) {
		if(m_bInPlaceActive && !m_bUIActive) {
			flags.uiActivationPending = TRUE;
		} else {
			wasHandled = FALSE;
		}
		return MA_ACTIVATE;
	}
	return MA_NOACTIVATE;
}

LRESULT ToolBar::OnMouseHover(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseHover(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ToolBar::OnMouseLeave(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(-1, button, shift);
		Raise_MouseLeave(index, button, shift, mouseStatus.lastPosition.x, mouseStatus.lastPosition.y);
	}

	return 0;
}

LRESULT ToolBar::OnMouseMove(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		Raise_MouseMove(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else if(!mouseStatus.enteredControl) {
		mouseStatus.EnterControl();
	}

	return 0;
}

LRESULT ToolBar::OnNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case TTN_GETDISPINFOA:
			return OnToolTipGetDispInfoNotificationA(FALSE, message, wParam, lParam, wasHandled);
			break;

		case TTN_GETDISPINFOW:
			return OnToolTipGetDispInfoNotificationW(FALSE, message, wParam, lParam, wasHandled);
			break;

		default:
			if(flags.chevronPopupVisible && parentWindowChevronPopupMenu.IsWindow()) {
				LPNMHDR pNotificationDetails = reinterpret_cast<LPNMHDR>(lParam);
				if(pNotificationDetails->hwndFrom == chevronPopupToolbar) {
					return chevronPopupToolbar.SendMessage(OCM_NOTIFY, wParam, lParam);
				}
				break;
			}
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ToolBar::OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	return DefWindowProc(message, wParam, lParam);
}

LRESULT ToolBar::OnRButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(!(properties.disabledEvents & deClickEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_RDblClick(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ToolBar::OnRButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.focusOnClick || flags.allowSetFocus) {
		/* HACK: Fixes the Validate event of VB6. ToolbarWindow32 doesn't seem to get the focus on
		 * WM_MOUSEACTIVATE by default. If we call SetFocus on WM_MOUSEACTIVATE, the Validate event is not fired
		 * for some reason. Therefore we call SetFocus here.
		 */
		SetFocus();
	}
	if(!(properties.disabledEvents & deMouseEvents)) {
		SHORT button = 0;
		SHORT shift = 0;
		WPARAM2BUTTONANDSHIFT(wParam, button, shift);
		button = 2/*MouseButtonConstants.vbRightButton*/;
		Raise_MouseDown(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	} else {
		if(!mouseStatus.enteredControl) {
			mouseStatus.EnterControl();
		}
		mouseStatus.StoreClickCandidate(2/*MouseButtonConstants.vbRightButton*/);
	}

	return 0;
}

LRESULT ToolBar::OnRButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(wParam, button, shift);
	button = 2/*MouseButtonConstants.vbRightButton*/;
	Raise_MouseUp(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ToolBar::OnSetCursor(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	HCURSOR hCursor = NULL;
	BOOL setCursor = FALSE;

	// Are we really over the control?
	CRect clientArea;
	if(index == 6) {
		chevronPopupToolbar.GetClientRect(&clientArea);
		chevronPopupToolbar.ClientToScreen(&clientArea);
	} else {
		GetClientRect(&clientArea);
		ClientToScreen(&clientArea);
	}
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	if(clientArea.PtInRect(mousePosition)) {
		// maybe the control is overlapped by a foreign window
		if(WindowFromPoint(mousePosition) == (index == 6 ? chevronPopupToolbar : static_cast<HWND>(*this))) {
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

LRESULT ToolBar::OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	if(properties.focusOnClick || flags.allowSetFocus || (GetAsyncKeyState(VK_TAB) & 0x8000)) {
		LRESULT lr = CComControl<ToolBar>::OnSetFocus(message, wParam, lParam, wasHandled);
		if(m_bInPlaceActive && !m_bUIActive && flags.uiActivationPending) {
			flags.uiActivationPending = FALSE;

			// now execute what usually would have been done on WM_MOUSEACTIVATE
			BOOL dummy = TRUE;
			CComControl<ToolBar>::OnMouseActivate(WM_MOUSEACTIVATE, 0, 0, dummy);
		}
		return lr;
	}
	return 0;
}

LRESULT ToolBar::OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(FireOnRequestEdit(DISPID_TB_FONT) == S_FALSE) {
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
		FireOnChanged(DISPID_TB_FONT);
	}

	return lr;
}

LRESULT ToolBar::OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
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

LRESULT ToolBar::OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(wParam == SPI_SETICONTITLELOGFONT) {
		if(properties.useSystemFont) {
			ApplyFont();
			//Invalidate();
		}
	} else if(wParam == SPI_SETKEYBOARDCUES) {
		GetSystemSettings();
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ToolBar::OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
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

	flags.applyBackgroundHack = FALSE;
	if(GetStyle() & TBSTYLE_TRANSPARENT) {
		if(!GetWindowTheme(*this)) {
			TCHAR pClassName[300];
			GetClassName(GetParent(), pClassName, 300);
			if(lstrcmpi(pClassName, TEXT("ReBarU")) != 0 && lstrcmpi(pClassName, TEXT("ReBarA")) != 0 && lstrcmpi(pClassName, REBARCLASSNAME) != 0) {
				flags.applyBackgroundHack = TRUE;
				ModifyStyle(0, TBSTYLE_CUSTOMERASE);
			}
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ToolBar::OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
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

		case timers.ID_DRAGCLICK:
			KillTimer(timers.ID_DRAGCLICK);
			if((dragDropStatus.lastDropTarget != -1) && (dragDropStatus.lastDropTarget != dragDropStatus.lastClickedButton)) {
				TBBUTTONINFO button = {0};
				button.cbSize = sizeof(button);
				button.dwMask = TBIF_BYINDEX | TBIF_COMMAND | TBIF_STATE | TBIF_STYLE;
				if(SendMessage(TB_GETBUTTONINFO, dragDropStatus.lastDropTarget, reinterpret_cast<LPARAM>(&button)) == dragDropStatus.lastDropTarget) {
					dragDropStatus.lastClickedButton = dragDropStatus.lastDropTarget;
					if((button.fsState & TBSTATE_ENABLED) == TBSTATE_ENABLED && !(button.fsStyle & BTNS_SEP)) {
						BOOL isDropDownButton = (button.fsStyle & BTNS_WHOLEDROPDOWN) == BTNS_WHOLEDROPDOWN;
						isDropDownButton = isDropDownButton || ((button.fsStyle & BTNS_DROPDOWN) == BTNS_DROPDOWN && !(SendMessage(TB_GETEXTENDEDSTYLE, 0, 0) & TBSTYLE_EX_DRAWDDARROWS));
						if(isDropDownButton) {
							NMTOOLBAR data = {0};
							data.hdr.code = TBN_DROPDOWN;
							data.hdr.hwndFrom = *this;
							data.hdr.idFrom = GetDlgCtrlID();
							data.iItem = dragDropStatus.lastDropTarget;
							SendMessage(TB_GETITEMRECT, dragDropStatus.lastDropTarget, reinterpret_cast<LPARAM>(&data.rcButton));
							GetParent().SendMessage(WM_NOTIFY, reinterpret_cast<WPARAM>(static_cast<HWND>(*this)), reinterpret_cast<LPARAM>(&data));
						} else {
							GetParent().PostMessage(WM_COMMAND, MAKEWPARAM(button.idCommand, BN_CLICKED), reinterpret_cast<LPARAM>(static_cast<HWND>(*this)));
						}
					}
				}
			}
			break;

		case timers.ID_PRESELECTCHEVRONMENUITEM:
			KillTimer(timers.ID_PRESELECTCHEVRONMENUITEM);
			if(properties.acceleratorToSendToChevronMenu > 0) {
				// ugly hack, but Windows XP Explorer seems to do something similar
				if(IsMenu(properties.hChevronPopupMenu)) {
					HWND hWndMenu = NULL;
					do {
						hWndMenu = FindWindowEx(NULL, hWndMenu, MAKEINTATOM(0x8000), NULL);
						if(hWndMenu) {
							if(properties.hChevronPopupMenu == reinterpret_cast<HMENU>(::SendMessage(hWndMenu, MN_GETHMENU, 0, 0))) {
								::PostMessage(hWndMenu, WM_CHAR, properties.acceleratorToSendToChevronMenu, 0);
								break;
							}
						}
					} while(hWndMenu);
				} else if(chevronPopupToolbar.IsWindow()) {
					// not sure this really works
					chevronPopupToolbar.PostMessage(WM_CHAR, properties.acceleratorToSendToChevronMenu, 0);
				}
				properties.acceleratorToSendToChevronMenu = 0;
			}
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ToolBar::OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
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

	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE) || (pDetails->flags & SWP_FRAMECHANGED)) {
		TCHAR pBuffer[255];
		ZeroMemory(pBuffer, 255 * sizeof(TCHAR));
		HWND hWndParent = GetParent();
		GetClassName(hWndParent, pBuffer, 255);
		BOOL parentIsPager = (!lstrcmp(TEXT("PagerU"), pBuffer) || !lstrcmp(TEXT("PagerA"), pBuffer) || !lstrcmp(TEXT("SysPager"), pBuffer));
		if(parentIsPager) {
			/* If we sit on a pager, and our OLE size becomes larger than the pager's client area, the pager's
				* scroll buttons will stop working if they're entered from the wrong side.
				* Therefore we have to setup some clipping.
				*/
			CRect pagerClientRectangle;
			::GetWindowRect(hWndParent, &pagerClientRectangle);
			::MapWindowPoints(HWND_DESKTOP, hWndParent, reinterpret_cast<LPPOINT>(&pagerClientRectangle), 2);
			SetObjectRects(&windowRectangle, &pagerClientRectangle);
		}
	}
	if(!(pDetails->flags & SWP_NOMOVE) || !(pDetails->flags & SWP_NOSIZE)) {
		ATLASSUME(m_spInPlaceSite);
		if(m_spInPlaceSite && !windowRectangle.EqualRect(&m_rcPos) && !flags.ignoreResize) {
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
	if(!(pDetails->flags & SWP_NOSIZE)) {
		if(SendMessage(TB_GETEXTENDEDSTYLE, 0, 0) & TBSTYLE_EX_MULTICOLUMN) {
			SIZE boundingSize = windowRectangle.Size();
			SendMessage(TB_SETBOUNDINGSIZE, 0, reinterpret_cast<LPARAM>(&boundingSize));
		}
	}

	wasHandled = FALSE;
	return 0;
}

LRESULT ToolBar::OnXButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
		Raise_XDblClick(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
	}

	return 0;
}

LRESULT ToolBar::OnXButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	wasHandled = FALSE;

	if(properties.focusOnClick || flags.allowSetFocus) {
		/* HACK: Fixes the Validate event of VB6. ToolbarWindow32 doesn't seem to get the focus on
		 * WM_MOUSEACTIVATE by default. If we call SetFocus on WM_MOUSEACTIVATE, the Validate event is not fired
		 * for some reason. Therefore we call SetFocus here.
		 */
		SetFocus();
	}
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
		Raise_MouseDown(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
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

LRESULT ToolBar::OnXButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
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
	Raise_MouseUp(index, button, shift, GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));

	return 0;
}

LRESULT ToolBar::OnReflectedNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case NM_CUSTOMDRAW:
			if(reinterpret_cast<LPNMHDR>(lParam)->hwndFrom == *this) {
				return OnCustomDrawNotification(0, message, wParam, lParam, wasHandled);
			} else if(reinterpret_cast<LPNMHDR>(lParam)->hwndFrom == chevronPopupToolbar) {
				return OnCustomDrawNotification(6, message, wParam, lParam, wasHandled);
			}
			break;
		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}

LRESULT ToolBar::OnReflectedNotifyFormat(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
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

LRESULT ToolBar::OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled)
{
	BOOL succeeded = FALSE;
	BOOL useVistaDragImage = FALSE;
	if(dragDropStatus.pDraggedButtons) {
		if(flags.usingThemes && properties.oleDragImageStyle == odistAeroIfAvailable && RunTimeHelper::IsVista()) {
			succeeded = CreateVistaOLEDragImage(dragDropStatus.pDraggedButtons, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
			useVistaDragImage = succeeded;
		}
		if(!succeeded) {
			// use a legacy drag image as fallback
			succeeded = CreateLegacyOLEDragImage(dragDropStatus.pDraggedButtons, reinterpret_cast<LPSHDRAGIMAGE>(lParam));
		}

		if(succeeded && RunTimeHelper::IsVista()) {
			FORMATETC format = {0};
			format.cfFormat = static_cast<CLIPFORMAT>(RegisterClipboardFormat(TEXT("UsingDefaultDragImage")));
			format.dwAspect = DVASPECT_CONTENT;
			format.lindex = -1;
			format.tymed = TYMED_HGLOBAL;
			STGMEDIUM medium = {0};
			medium.tymed = TYMED_HGLOBAL;
			medium.hGlobal = GlobalAlloc(GPTR, sizeof(BOOL));
			if(medium.hGlobal) {
				LPBOOL pUseVistaDragImage = static_cast<LPBOOL>(GlobalLock(medium.hGlobal));
				*pUseVistaDragImage = useVistaDragImage;
				GlobalUnlock(medium.hGlobal);

				dragDropStatus.pSourceDataObject->SetData(&format, &medium, TRUE);
			}
		}
	}

	wasHandled = succeeded;
	// TODO: Why do we have to return FALSE to have the set offset not ignored if a Vista drag image is used?
	return succeeded && !useVistaDragImage;
}

LRESULT ToolBar::OnGetToolbar(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	// If we don't need this message, don't prevent others from getting asked whether they want this message.
	if(IsWindowVisible() && properties.menuMode) {
		if(wParam) {
			*reinterpret_cast<PLONG>(wParam) = 0x02;
		}
		if(lParam) {
			*reinterpret_cast<LPBOOL>(lParam) = properties.menuMode;
		}
		HWND hWnd = *this;
		return reinterpret_cast<LRESULT>(hWnd);
	}
	return NULL;
}

LRESULT ToolBar::OnAutoPopup(LONG index, UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	if(properties.menuMode) {
		droppedDownButton = wParam;
		CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(droppedDownButton, index == 6, this);
		HMENU hPopupMenu = NULL;
		LONG h = NULL;
		Raise_ButtonGetDropDownMenu(pTBarButton, &h);
		hPopupMenu = static_cast<HMENU>(LongToHandle(h));

		RECT tmpRectangle = {0};
		if(index == 6) {
			chevronPopupToolbar.SendMessage(TB_GETITEMRECT, droppedDownButton, reinterpret_cast<LPARAM>(&tmpRectangle));
		} else {
			SendMessage(TB_GETITEMRECT, droppedDownButton, reinterpret_cast<LPARAM>(&tmpRectangle));
		}
		DropDownReturnValuesConstants furtherProcessing = ddrvDoDefault;
		Raise_DropDown(pTBarButton, reinterpret_cast<RECTANGLE*>(&tmpRectangle), &furtherProcessing);

		if(furtherProcessing == ddrvDoDefault) {
			DoPopupMenu(wParam, hPopupMenu, FALSE);
		}
		droppedDownButton = -1;
	}
	return 0;
}

LRESULT ToolBar::OnMDIChildWindowStateChanged(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	if(properties.menuMode) {
		menuModeState.mdiChildIsMaximized = wParam;
	}
	return 0;
}

LRESULT ToolBar::OnDisplayChevronPopup(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/)
{
	if(lParam) {
		POINT pt = *reinterpret_cast<LPPOINT>(lParam);
		DisplayChevronPopupWindow(pt.x, pt.y);
	}
	return 0;
}

LRESULT ToolBar::OnDeferredSetButtonText(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/)
{
	int buttonIndex = IDToButtonIndex(wParam);
	CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonIndex, FALSE, this, FALSE);
	if(pTBarButton) {
		LONG iconIndex = 0;
		LONG imageListIndex = 0;
		CComBSTR buttonText = L"";
		LONG maxButtonTextLength = 0;

		RequestedInfoConstants requestedInfo = riButtonText;
		maxButtonTextLength = 500;

		VARIANT_BOOL dontAskAgain = VARIANT_FALSE;
		Raise_ButtonGetDisplayInfo(pTBarButton, requestedInfo, &iconIndex, &imageListIndex, maxButtonTextLength, &buttonText, &dontAskAgain);
		pTBarButton->put_Text(buttonText);
	}
	return 0;
}

LRESULT ToolBar::OnAddButtons(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT succeeded = FALSE;

	if(!(properties.disabledEvents & deButtonInsertionEvents) && wParam > 0) {
		LPTBBUTTON pButtons = new TBBUTTON[wParam];
		ZeroMemory(pButtons, wParam * sizeof(TBBUTTON));
		LPTBBUTTON pNextButton = pButtons;
		int realNumberOfButtons = 0;

		#ifdef UNICODE
			BOOL mustConvert = (message == TB_ADDBUTTONSA);
		#else
			BOOL mustConvert = (message == TB_ADDBUTTONSW);
		#endif
		for(WPARAM i = 0; i < wParam; i++) {
			VARIANT_BOOL cancel = VARIANT_FALSE;

			CComObject<VirtualToolBarButton>* pVTBarButtonObj = NULL;
			CComPtr<IVirtualToolBarButton> pVTBarButtonItf = NULL;
			CComObject<VirtualToolBarButton>::CreateInstance(&pVTBarButtonObj);
			pVTBarButtonObj->AddRef();
			pVTBarButtonObj->SetOwner(this);

			if(mustConvert) {
				LPTBBUTTON pButtonData = reinterpret_cast<LPTBBUTTON>(lParam) + i;
				TBBUTTON convertedButtonData = {0};
				#ifdef UNICODE
					LPSTR p = NULL;
					if(pButtonData->iString != reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKA)) {
						p = reinterpret_cast<LPSTR>(pButtonData->iString);
					}
					CA2W converter(p);
					if(pButtonData->iString == reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKA)) {
						convertedButtonData.iString = reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKW);
					} else {
						convertedButtonData.iString = reinterpret_cast<INT_PTR>(static_cast<LPWSTR>(converter));
					}
				#else
					LPWSTR p = NULL;
					if(pButtonData->iString != reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKW)) {
						p = reinterpret_cast<LPWSTR>(pButtonData->iString);
					}
					CW2A converter(p);
					if(pButtonData->iString == reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKW)) {
						convertedButtonData.iString = reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKA);
					} else {
						convertedButtonData.iString = reinterpret_cast<INT_PTR>(static_cast<LPSTR>(converter));
					}
				#endif
				convertedButtonData.iBitmap = pButtonData->iBitmap;
				convertedButtonData.idCommand = pButtonData->idCommand;
				convertedButtonData.fsState = pButtonData->fsState;
				convertedButtonData.fsStyle = pButtonData->fsStyle;
				convertedButtonData.dwData = pButtonData->dwData;
				CopyMemory(convertedButtonData.bReserved, pButtonData->bReserved, ARRAYSIZE(convertedButtonData.bReserved) * sizeof(BYTE));
				pVTBarButtonObj->LoadState(&convertedButtonData, wParam);
			} else {
				pVTBarButtonObj->LoadState(reinterpret_cast<LPTBBUTTON>(lParam) + i, wParam);
			}

			pVTBarButtonObj->QueryInterface(IID_IVirtualToolBarButton, reinterpret_cast<LPVOID*>(&pVTBarButtonItf));
			pVTBarButtonObj->Release();

			Raise_InsertingButton(pVTBarButtonItf, &cancel);
			pVTBarButtonObj = NULL;

			if(cancel == VARIANT_FALSE) {
				// we will pass this button to the tool bar
				*pNextButton = reinterpret_cast<LPTBBUTTON>(lParam)[i];
				pNextButton++;
				realNumberOfButtons++;
			}
		}

		// finally pass the message to the tool bar
		if(realNumberOfButtons > 0) {
			succeeded = DefWindowProc(message, realNumberOfButtons, reinterpret_cast<LPARAM>(pButtons));
			if(succeeded) {
				for(int i = 0; i < realNumberOfButtons; i++) {
					int insertedButtonIndex = IDToButtonIndex(pButtons[i].idCommand);
					CComPtr<IToolBarButton> pTBarButtonItf = ClassFactory::InitToolBarButton(insertedButtonIndex, FALSE, this);
					if(pTBarButtonItf) {
						Raise_InsertedButton(pTBarButtonItf);
					}
				}
			}
		}
		delete[] pButtons;
	} else {
		// finally pass the message to the tool bar
		succeeded = DefWindowProc(message, wParam, lParam);
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		if(succeeded) {
			// maybe we have a new button below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);

			UINT hitTestDetails = 0;
			int newButtonUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, *this);
			if(newButtonUnderMouse != buttonUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(buttonUnderMouse != -1) {
					CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonUnderMouse, FALSE, this);
					Raise_ButtonMouseLeave(pTBarButton, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				buttonUnderMouse = newButtonUnderMouse;

				if(buttonUnderMouse != -1) {
					CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonUnderMouse, FALSE, this);
					Raise_ButtonMouseEnter(pTBarButton, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	return succeeded;
}

LRESULT ToolBar::OnDeleteButton(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = FALSE;

	VARIANT_BOOL cancel = VARIANT_FALSE;
	CComPtr<IToolBarButton> pTBarButtonItf = NULL;
	CComObject<ToolBarButton>* pTBarButtonObj = NULL;
	if(!(properties.disabledEvents & deButtonDeletionEvents)) {
		CComObject<ToolBarButton>::CreateInstance(&pTBarButtonObj);
		pTBarButtonObj->AddRef();
		pTBarButtonObj->SetOwner(this);
		pTBarButtonObj->Attach(wParam, FALSE);
		pTBarButtonObj->QueryInterface(IID_IToolBarButton, reinterpret_cast<LPVOID*>(&pTBarButtonItf));
		pTBarButtonObj->Release();
		Raise_RemovingButton(pTBarButtonItf, &cancel);
	}

	if(cancel == VARIANT_FALSE) {
		CComPtr<IVirtualToolBarButton> pVTBarButtonItf = NULL;
		if(!(properties.disabledEvents & deButtonDeletionEvents)) {
			CComObject<VirtualToolBarButton>* pVTBarButtonObj = NULL;
			CComObject<VirtualToolBarButton>::CreateInstance(&pVTBarButtonObj);
			pVTBarButtonObj->AddRef();
			pVTBarButtonObj->SetOwner(this);
			if(pTBarButtonObj) {
				pTBarButtonObj->SaveState(pVTBarButtonObj);
			}

			pVTBarButtonObj->QueryInterface(IID_IVirtualToolBarButton, reinterpret_cast<LPVOID*>(&pVTBarButtonItf));
			pVTBarButtonObj->Release();
		}

		if(!(properties.disabledEvents & deMouseEvents)) {
			if(static_cast<int>(wParam) == buttonUnderMouse) {
				// we're removing the button below the mouse cursor
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);

				UINT hitTestDetails = 0;
				HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, *this);
				Raise_ButtonMouseLeave(pTBarButtonItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				buttonUnderMouse = -1;
			}
			if(static_cast<int>(wParam) == droppedDownButton) {
				droppedDownButton = -1;
			}
		}

		// finally pass the message to the tool bar
		LONG buttonIDBeingRemoved = ButtonIndexToID(wParam);
		flags.buttonIsRemovedByCustomizationDialog = FALSE;
		lr = DefWindowProc(message, wParam, lParam);
		flags.buttonIsRemovedByCustomizationDialog = flags.isBeingCustomized;
		if(lr) {
			RemoveButtonFromButtonContainers(buttonIDBeingRemoved);
			//RebuildAcceleratorTable();

			if(!(properties.disabledEvents & deButtonDeletionEvents)) {
				Raise_RemovedButton(pVTBarButtonItf);
			}
			DeregisterPlaceholderButton(buttonIDBeingRemoved);
		}

		if(!(properties.disabledEvents & deMouseEvents)) {
			if(lr) {
				// maybe we have a new button below the mouse cursor now
				DWORD position = GetMessagePos();
				POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
				ScreenToClient(&mousePosition);

				UINT hitTestDetails = 0;
				int newButtonUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, *this);
				if(newButtonUnderMouse != buttonUnderMouse) {
					SHORT button = 0;
					SHORT shift = 0;
					WPARAM2BUTTONANDSHIFT(-1, button, shift);
					if(buttonUnderMouse != -1) {
						pTBarButtonItf = ClassFactory::InitToolBarButton(buttonUnderMouse, FALSE, this);
						Raise_ButtonMouseLeave(pTBarButtonItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}

					buttonUnderMouse = newButtonUnderMouse;

					if(buttonUnderMouse != -1) {
						pTBarButtonItf = ClassFactory::InitToolBarButton(buttonUnderMouse, FALSE, this);
						Raise_ButtonMouseEnter(pTBarButtonItf, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
					}
				}
			}
		}
	}

	return lr;
}

LRESULT ToolBar::OnInsertButton(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT succeeded = FALSE;

	if(!(properties.disabledEvents & deButtonInsertionEvents)) {
		VARIANT_BOOL cancel = VARIANT_FALSE;

		CComObject<VirtualToolBarButton>* pVTBarButtonObj = NULL;
		CComPtr<IVirtualToolBarButton> pVTBarButtonItf = NULL;
		CComObject<VirtualToolBarButton>::CreateInstance(&pVTBarButtonObj);
		pVTBarButtonObj->AddRef();
		pVTBarButtonObj->SetOwner(this);

		#ifdef UNICODE
			BOOL mustConvert = (message == TB_INSERTBUTTONA);
		#else
			BOOL mustConvert = (message == TB_INSERTBUTTONW);
		#endif
		if(mustConvert) {
			LPTBBUTTON pButtonData = reinterpret_cast<LPTBBUTTON>(lParam);
			TBBUTTON convertedButtonData = {0};
			#ifdef UNICODE
				LPSTR p = NULL;
				if(pButtonData->iString != reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKA)) {
					p = reinterpret_cast<LPSTR>(pButtonData->iString);
				}
				CA2W converter(p);
				if(pButtonData->iString == reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKA)) {
					convertedButtonData.iString = reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKW);
				} else {
					convertedButtonData.iString = reinterpret_cast<INT_PTR>(static_cast<LPWSTR>(converter));
				}
			#else
				LPWSTR p = NULL;
				if(pButtonData->iString != reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKW)) {
					p = reinterpret_cast<LPWSTR>(pButtonData->iString);
				}
				CW2A converter(p);
				if(pButtonData->iString == reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKW)) {
					convertedButtonData.iString = reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKA);
				} else {
					convertedButtonData.iString = reinterpret_cast<INT_PTR>(static_cast<LPSTR>(converter));
				}
			#endif
			convertedButtonData.iBitmap = pButtonData->iBitmap;
			convertedButtonData.idCommand = pButtonData->idCommand;
			convertedButtonData.fsState = pButtonData->fsState;
			convertedButtonData.fsStyle = pButtonData->fsStyle;
			convertedButtonData.dwData = pButtonData->dwData;
			CopyMemory(convertedButtonData.bReserved, pButtonData->bReserved, ARRAYSIZE(convertedButtonData.bReserved) * sizeof(BYTE));
			pVTBarButtonObj->LoadState(&convertedButtonData, wParam);
		} else {
			pVTBarButtonObj->LoadState(reinterpret_cast<LPTBBUTTON>(lParam), wParam);
		}

		pVTBarButtonObj->QueryInterface(IID_IVirtualToolBarButton, reinterpret_cast<LPVOID*>(&pVTBarButtonItf));
		pVTBarButtonObj->Release();

		Raise_InsertingButton(pVTBarButtonItf, &cancel);
		pVTBarButtonObj = NULL;

		if(cancel == VARIANT_FALSE) {
			// finally pass the message to the tool bar
			succeeded = DefWindowProc(message, wParam, lParam);
			if(succeeded) {
				//RebuildAcceleratorTable();
				int insertedButtonIndex = IDToButtonIndex(reinterpret_cast<LPTBBUTTON>(lParam)->idCommand);
				CComPtr<IToolBarButton> pTBarButtonItf = ClassFactory::InitToolBarButton(insertedButtonIndex, FALSE, this);
				if(pTBarButtonItf) {
					Raise_InsertedButton(pTBarButtonItf);
				}
			}
		}
	} else {
		// finally pass the message to the tool bar
		succeeded = DefWindowProc(message, wParam, lParam);
		if(succeeded) {
			//RebuildAcceleratorTable();
		}
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		if(succeeded) {
			// maybe we have a new button below the mouse cursor now
			DWORD position = GetMessagePos();
			POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
			ScreenToClient(&mousePosition);

			UINT hitTestDetails = 0;
			int newButtonUnderMouse = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, *this);
			if(newButtonUnderMouse != buttonUnderMouse) {
				SHORT button = 0;
				SHORT shift = 0;
				WPARAM2BUTTONANDSHIFT(-1, button, shift);
				if(buttonUnderMouse != -1) {
					CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonUnderMouse, FALSE, this);
					Raise_ButtonMouseLeave(pTBarButton, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}

				buttonUnderMouse = newButtonUnderMouse;

				if(buttonUnderMouse != -1) {
					CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonUnderMouse, FALSE, this);
					Raise_ButtonMouseEnter(pTBarButton, button, shift, mousePosition.x, mousePosition.y, hitTestDetails);
				}
			}
		}
	}

	return succeeded;
}

LRESULT ToolBar::OnSetButtonInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPTBBUTTONINFO pButton = reinterpret_cast<LPTBBUTTONINFO>(lParam);
	/*BYTE oldState = 0;
	BYTE newState = pButton->fsState & (TBSTATE_CHECKED | TBSTATE_INDETERMINATE);
	BOOL getState = (pButton->dwMask & TBIF_STATE);
	BOOL changeSelectionState = FALSE;*/

	#ifdef USE_STL
		BOOL update = (placeholderButtons.size() > 0);
	#else
		BOOL update = (placeholderButtons.GetCount() > 0);
	#endif
	LONG oldID = wParam;
	LONG newID = -1;
	if(update) {
		// okay, we have placeholders, so we need to update the map
		if(pButton) {
			update = (pButton->dwMask & TBIF_COMMAND);
			if(update) {
				if(pButton->dwMask & TBIF_BYINDEX) {
					TBBUTTONINFO button = {0};
					button.cbSize = sizeof(button);
					button.dwMask = TBIF_BYINDEX | TBIF_COMMAND;
					/*if(getState) {
						button.dwMask |= TBIF_STATE;
					}*/
					if(static_cast<WPARAM>(SendMessage(TB_GETBUTTONINFO, wParam, reinterpret_cast<LPARAM>(&button))) == wParam) {
						oldID = button.idCommand;
						/*if(getState) {
							oldState = button.fsState & (TBSTATE_CHECKED | TBSTATE_INDETERMINATE);
							getState = FALSE;
							changeSelectionState = (oldState != newState);
						}*/
					}
				}
				newID = pButton->idCommand;
			}
		} else {
			update = FALSE;
		}
	}
	/*if(getState) {
		TBBUTTONINFO button = {0};
		button.cbSize = sizeof(button);
		button.dwMask = TBIF_BYINDEX | TBIF_STATE;
		if(static_cast<WPARAM>(SendMessage(TB_GETBUTTONINFO, wParam, reinterpret_cast<LPARAM>(&button))) == wParam) {
			oldState = button.fsState & (TBSTATE_CHECKED | TBSTATE_INDETERMINATE);
			changeSelectionState = (oldState != newState);
		}
	}*/

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr && update && oldID != -1 && oldID != newID) {
		// there was an ID change, so update the map of placeholders
		LPPLACEHOLDERBUTTON pDetails = NULL;
		#ifdef USE_STL
			std::unordered_map<LONG, LPPLACEHOLDERBUTTON>::iterator iter = placeholderButtons.find(oldID);
			if(iter != placeholderButtons.end()) {
				pDetails = iter->second;
				placeholderButtons.erase(iter);
			}
		#else
			CAtlMap<LONG, LPPLACEHOLDERBUTTON>::CPair* pEntry = placeholderButtons.Lookup(oldID);
			if(pEntry) {
				pDetails = pEntry->m_value;
				placeholderButtons.RemoveKey(oldID);
			}
		#endif
		placeholderButtons[newID] = pDetails;
	}
	/*if(lr && changeSelectionState) {
		int buttonIndex;
		if(pButton->dwMask & TBIF_BYINDEX) {
			buttonIndex = wParam;
		} else {
			buttonIndex = IDToButtonIndex(wParam);
		}
		CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonIndex, this);
		Raise_ButtonSelectionStateChanged(pTBarButton, static_cast<SelectionStateConstants>(oldState & (TBSTATE_CHECKED | TBSTATE_INDETERMINATE)), static_cast<SelectionStateConstants>(newState & (TBSTATE_CHECKED | TBSTATE_INDETERMINATE)));
	}*/
	return lr;
}

LRESULT ToolBar::OnSetButtonWidth(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		properties.minimumButtonWidth = LOWORD(lParam);
		properties.maximumButtonWidth = HIWORD(lParam);
	}
	return lr;
}

LRESULT ToolBar::OnSetCmdId(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	#ifdef USE_STL
		BOOL update = (placeholderButtons.size() > 0);
	#else
		BOOL update = (placeholderButtons.GetCount() > 0);
	#endif
	LONG oldID = -1;
	if(update) {
		// okay, we have placeholders, so we need to update the map
		TBBUTTONINFO button = {0};
		button.cbSize = sizeof(button);
		// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
		button.dwMask = TBIF_BYINDEX | TBIF_COMMAND;
		if(static_cast<WPARAM>(SendMessage(TB_GETBUTTONINFO, wParam, reinterpret_cast<LPARAM>(&button))) == wParam) {
			oldID = button.idCommand;
		}
	}

	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr && update && oldID != -1 && oldID != lParam) {
		// there was an ID change, so update the map of placeholders
		LPPLACEHOLDERBUTTON pDetails = NULL;
		#ifdef USE_STL
			std::unordered_map<LONG, LPPLACEHOLDERBUTTON>::iterator iter = placeholderButtons.find(oldID);
			if(iter != placeholderButtons.end()) {
				pDetails = iter->second;
				placeholderButtons.erase(iter);
			}
		#else
			CAtlMap<LONG, LPPLACEHOLDERBUTTON>::CPair* pEntry = placeholderButtons.Lookup(oldID);
			if(pEntry) {
				pDetails = pEntry->m_value;
				placeholderButtons.RemoveKey(oldID);
			}
		#endif
		placeholderButtons[lParam] = pDetails;
	}
	return lr;
}

LRESULT ToolBar::OnSetDrawTextFlags(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	properties.drawTextFlagsMask = wParam;
	properties.drawTextFlags = (properties.defaultDrawTextFlags & ~wParam) | (lParam & wParam);
	return lr;
}

LRESULT ToolBar::OnSetDropDownGap(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		if(!IsInDesignMode() || !(properties.dropDownGap == -1 && wParam == static_cast<WPARAM>(GetSystemMetrics(SM_CXEDGE) * 2))) {
			properties.dropDownGap = wParam;
		}
	}
	return lr;
}

LRESULT ToolBar::OnSetIndent(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		properties.firstButtonIndentation = wParam;
	}
	return lr;
}

LRESULT ToolBar::OnSetListGap(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefWindowProc(message, wParam, lParam);
	if(lr) {
		if(!IsInDesignMode() || !(properties.horizontalIconCaptionGap == -1 && wParam == static_cast<WPARAM>(GetSystemMetrics(SM_CXEDGE) * 2))) {
			properties.horizontalIconCaptionGap = wParam;
		}
	}
	return lr;
}

LRESULT ToolBar::OnParentActivate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LRESULT lr = DefSubclassProc(parentWindowMenuMode, message, wParam, lParam);
	Invalidate();
	UpdateWindow();
	return lr;
}

LRESULT ToolBar::OnParentSysCommand(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if((menuModeState.pressedSysKey == VK_MENU || (menuModeState.pressedSysKey == VK_F10 && !(GetKeyState(VK_SHIFT) & 0x80)) || menuModeState.pressedSysKey == VK_SPACE) && wParam == SC_KEYMENU) {
		if(GetFocus() == *this) {
			GiveFocusBack();     // exit menu "loop"
			PostMessage(TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
		} else if(menuModeState.pressedSysKey != VK_SPACE && !menuModeState.skipSysCommandMessage) {
			if(menuModeState.useContextSensitiveKeyboardCues && !menuModeState.displayKeyboardCues && menuModeState.allowDisplayingKeyboardCues) {
				ShowKeyboardCues(TRUE);
			}
			TakeFocus();     // enter menu "loop"
			wasHandled = TRUE;
		} else if(menuModeState.pressedSysKey != VK_SPACE) {
			wasHandled = TRUE;
		}
	}
	menuModeState.skipSysCommandMessage = FALSE;
	return 0;
}

LRESULT ToolBar::OnHookMouseMove(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};

	wasHandled = FALSE;
	if(menuModeState.menuIsActive) {
		if(::WindowFromPoint(mousePosition) == *this) {
			ScreenToClient(&mousePosition);
			int hitButton = HitTest(mousePosition.x, mousePosition.y, NULL, *this);

			if((mousePosition.x != menuModeState.lastMouseMovePosition.x || mousePosition.y != menuModeState.lastMouseMovePosition.y) && hitButton >= 0 && hitButton < SendMessage(TB_BUTTONCOUNT, 0, 0) && hitButton != menuModeState.activePopupMenuButtonIndex && menuModeState.activePopupMenuButtonIndex != -1) {
				BOOL shouldBeDisplayedInChevron = FALSE;
				RECT availableRectangle = {0};
				GetClientRect(&availableRectangle);
				CRect buttonRectangle;
				if(SendMessage(TB_GETITEMRECT, hitButton, reinterpret_cast<LPARAM>(&buttonRectangle))) {
					if(GetStyle() & CCS_VERT) {
						shouldBeDisplayedInChevron = (buttonRectangle.bottom > availableRectangle.bottom);
					} else {
						shouldBeDisplayedInChevron = (buttonRectangle.right > availableRectangle.right);
					}
					if(!shouldBeDisplayedInChevron) {
						TBBUTTON button = {0};
						SendMessage(TB_GETBUTTON, hitButton, reinterpret_cast<LPARAM>(&button));
						if(button.fsState & TBSTATE_ENABLED) {
							menuModeState.nextPopupMenuButtonIndex = hitButton | 0xFFFF0000;
							HWND hWndMenu = menuModeState.menuWindowStack.GetCurrent();
							ATLASSERT(hWndMenu);

							// this one is needed to close a menu if mouse button was down
							::PostMessage(hWndMenu, WM_LBUTTONUP, 0, MAKELPARAM(mousePosition.x, mousePosition.y));
							// this one closes a popup menu
							::PostMessage(hWndMenu, WM_KEYDOWN, VK_ESCAPE, 0);

							wasHandled = TRUE;
						}
					}
				}
			}
		}
	} else {
		ScreenToClient(&mousePosition);
	}
	menuModeState.lastMouseMovePosition = mousePosition;
	return 0;
}

LRESULT ToolBar::OnHookSysKeyDown(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(wParam == VK_MENU && menuModeState.parentWindowIsActive && menuModeState.useContextSensitiveKeyboardCues && !menuModeState.displayKeyboardCues && menuModeState.allowDisplayingKeyboardCues) {
		ShowKeyboardCues(TRUE);
	}

	if(wParam != VK_SPACE && !menuModeState.menuIsActive && GetFocus() == *this) {
		menuModeState.allowDisplayingKeyboardCues = FALSE;
		PostMessage(TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
		GiveFocusBack();
		menuModeState.skipSysCommandMessage = TRUE;
	} else {
		if(wParam == VK_SPACE && menuModeState.useContextSensitiveKeyboardCues && menuModeState.displayKeyboardCues) {
			menuModeState.allowDisplayingKeyboardCues = TRUE;
			ShowKeyboardCues(FALSE);
		}
		menuModeState.pressedSysKey = static_cast<UINT>(wParam);
	}
	return 0;
}

LRESULT ToolBar::OnHookSysKeyUp(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	if(!menuModeState.allowDisplayingKeyboardCues) {
		menuModeState.allowDisplayingKeyboardCues = TRUE;
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ToolBar::OnHookSysChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	if(!menuModeState.menuIsActive && menuModeState.hWndForwardedMessage != *this && wParam != VK_SPACE) {
		wasHandled = TRUE;
	}
	return 0;
}

LRESULT ToolBar::OnHookKeyDown(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;

	int stackSize = menuModeState.menuWindowStack.GetSize();
	if(wParam == VK_ESCAPE && stackSize <= 1) {
		if(menuModeState.menuIsActive && !menuModeState.displayingPopupMenu) {
			int hotButton = SendMessage(TB_GETHOTITEM, 0, 0);
			if(hotButton == -1) {
				hotButton = menuModeState.activePopupMenuButtonIndex;
			}
			if(hotButton == -1) {
				hotButton = 0;
			}
			SendMessage(TB_SETHOTITEM, hotButton, 0);
			wasHandled = TRUE;
			TakeFocus();
			menuModeState.escapePressed = TRUE;     // To keep focus
			menuModeState.skipPostingKeyDown = FALSE;
		} else if(GetFocus() == *this && parentWindowMenuMode.IsWindow()) {
			SendMessage(TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
			GiveFocusBack();
			wasHandled = TRUE;
		}
	} else if(wParam == VK_RETURN || wParam == VK_UP || wParam == VK_DOWN) {
		if(!menuModeState.menuIsActive && GetFocus() == *this && parentWindowMenuMode.IsWindow()) {
			int hotButton = SendMessage(TB_GETHOTITEM, 0, 0);
			if(hotButton != -1) {
				if(wParam != VK_RETURN) {
					if(!menuModeState.skipPostingKeyDown) {
						PostMessage(WM_KEYDOWN, VK_DOWN, 0);
						menuModeState.skipPostingKeyDown = TRUE;
					} else {
						menuModeState.skipPostingKeyDown = FALSE;
					}
				}
			}
		}
		if(wParam == VK_RETURN && menuModeState.menuIsActive) {
			PostMessage(TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
			menuModeState.nextPopupMenuButtonIndex = -1;
			// TK: WTL does not do this. WTL does not support switching between MDI child window system menu and the normal menu.
			if(menuModeState.activePopupMenuButtonIndex == -3) {
				menuModeState.activePopupMenuButtonIndex = -1;
			}
			GiveFocusBack();
		}
	} else if(wParam == VK_LEFT || wParam == VK_RIGHT) {
		DWORD styleEx = GetExStyle();
		WPARAM keyCodeNext = (styleEx & WS_EX_LAYOUTRTL) ? VK_LEFT : VK_RIGHT;
		WPARAM keyCodePrev = (styleEx & WS_EX_LAYOUTRTL) ? VK_RIGHT : VK_LEFT;

		if(menuModeState.menuIsActive && !menuModeState.displayingPopupMenu && !(wParam == keyCodeNext && menuModeState.selectedMenuItemHasSubMenu)) {
			BOOL closeMenu = FALSE;
			stackSize = menuModeState.menuWindowStack.GetSize();
			if(wParam == keyCodePrev && stackSize == 1)	{
				menuModeState.nextPopupMenuButtonIndex = GetPreviousMenuItem(menuModeState.activePopupMenuButtonIndex);
				if(menuModeState.nextPopupMenuButtonIndex != -1) {
					closeMenu = TRUE;
				}
			} else if(wParam == keyCodeNext) {
				menuModeState.nextPopupMenuButtonIndex = GetNextMenuItem(menuModeState.activePopupMenuButtonIndex);
				if(menuModeState.nextPopupMenuButtonIndex != -1) {
					closeMenu = TRUE;
				}
			}
			HWND hWndMenu = menuModeState.menuWindowStack.GetCurrent();
			ATLASSERT(hWndMenu);

			// Close the popup menu
			if(closeMenu) {
				::PostMessage(hWndMenu, WM_KEYDOWN, VK_ESCAPE, 0);
				if(wParam == keyCodeNext) {
					int menuIndex = menuModeState.menuWindowStack.GetSize() - 1;
					while(menuIndex >= 0) {
						hWndMenu = menuModeState.menuWindowStack[menuIndex];
						if(hWndMenu) {
							::PostMessage(hWndMenu, WM_KEYDOWN, VK_ESCAPE, 0);
						}
						menuIndex--;
					}
				}
				// TK: WTL does not do this. WTL does not support switching between MDI child window system menu and the normal menu.
				/*if(menuModeState.nextPopupMenuButtonIndex == -3) {
					menuModeState.activePopupMenuButtonIndex = -3;
					menuModeState.nextPopupMenuButtonIndex = -1;
					SendMessage(TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
					CWindow parentWindow = GetParent();
					if(parentWindow.IsWindow()) {
						parentWindow.PostMessage(GetAutoPopupMessage(), 0, 0);
					}
				} else*/ if(menuModeState.nextPopupMenuButtonIndex == -2) {
					menuModeState.nextPopupMenuButtonIndex = -1;
					DisplayChevronMenu();
				}
				wasHandled = TRUE;
			}
		// TK: WTL does not do this. WTL does not support switching between MDI child window system menu and the normal menu.
		} else if(menuModeState.activePopupMenuButtonIndex == -3 && !(wParam == keyCodeNext && menuModeState.selectedMenuItemHasSubMenu)) {
			BOOL closeMenu = FALSE;
			if(wParam == keyCodePrev)	{
				menuModeState.nextPopupMenuButtonIndex = GetPreviousMenuItem(menuModeState.activePopupMenuButtonIndex);
				if(menuModeState.nextPopupMenuButtonIndex != -1) {
					closeMenu = TRUE;
				}
			} else if(wParam == keyCodeNext) {
				menuModeState.nextPopupMenuButtonIndex = GetNextMenuItem(menuModeState.activePopupMenuButtonIndex);
				if(menuModeState.nextPopupMenuButtonIndex != -1) {
					closeMenu = TRUE;
				}
			}
			// Close the popup menu
			if(closeMenu) {
				HWND hWndMenu = FindWindowEx(NULL, NULL, MAKEINTATOM(0x8000), NULL);
				if(hWndMenu) {
					::PostMessage(hWndMenu, WM_KEYDOWN, VK_ESCAPE, 0);
				}
				if(menuModeState.nextPopupMenuButtonIndex == -2) {
					menuModeState.nextPopupMenuButtonIndex = -1;
					DisplayChevronMenu();
				} else if(menuModeState.nextPopupMenuButtonIndex >= 0) {
					PostMessage(GetAutoPopupMessage(), menuModeState.nextPopupMenuButtonIndex & 0xFFFF);
					if(!(menuModeState.nextPopupMenuButtonIndex & 0xFFFF0000) && !menuModeState.selectedMenuItemHasSubMenu) {
						PostMessage(WM_KEYDOWN, VK_DOWN, 0);
					}
					menuModeState.nextPopupMenuButtonIndex = -1;
				}
				wasHandled = TRUE;
			}
		}
	}
	return 0;
}

LRESULT ToolBar::OnHookNextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = FALSE;
	return 1;
}

LRESULT ToolBar::OnHookChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled)
{
	wasHandled = (wParam == VK_ESCAPE);
	return 0;
}


LRESULT ToolBar::OnChevronPopupMenuDismiss(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	droppedDownButton = -1;
	if(chevronPopupToolbar.IsWindow()) {
		chevronPopupToolbar.ShowWindow(SW_HIDE);
		chevronPopupToolbar.SetParent(*this);
		/* Don't destroy the window here! For instance we might end up here from a TBN_DROPDOWN notification.
			 Destroying the tool bar while it waits for the return of a callback will lead to strange behavior
			 like heap corruptions that are very difficult to track. */
		//chevronPopupToolbar.DestroyWindow();
	}
	if(chevronPopupMenuWindow.IsWindow()) {
		chevronPopupMenuWindow.DestroyWindow();
	}
	flags.chevronPopupVisible = FALSE;

	LRESULT lr = 0;
	if(parentWindowChevronPopupMenu.IsWindow()) {
		lr = DefSubclassProc(parentWindowChevronPopupMenu, message, wParam, lParam);
		ATLVERIFY(RemoveWindowSubclass(parentWindowChevronPopupMenu, ToolBar::ParentWindowSubclass_ChevronPopupMenu, reinterpret_cast<UINT_PTR>(this)));
		parentWindowChevronPopupMenu.Detach();
	}
	return lr;
}


LRESULT ToolBar::OnChevronPopupToolbarNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled)
{
	switch(reinterpret_cast<LPNMHDR>(lParam)->code) {
		case TTN_GETDISPINFOA:
			return OnToolTipGetDispInfoNotificationA(TRUE, message, wParam, lParam, wasHandled);
			break;

		case TTN_GETDISPINFOW:
			return OnToolTipGetDispInfoNotificationW(TRUE, message, wParam, lParam, wasHandled);
			break;

		default:
			wasHandled = FALSE;
			break;
	}
	return 0;
}


LRESULT ToolBar::OnBeginAdjustNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	VARIANT_BOOL dontRestoreFromRegistry = VARIANT_FALSE;
	Raise_BeginCustomization(BOOL2VARIANTBOOL(flags.isBeingRestored), &dontRestoreFromRegistry);
	LRESULT lr = VARIANTBOOL2BOOL(flags.isBeingRestored && dontRestoreFromRegistry);
	if(lr) {
		// TBN_ENDADJUST won't be sent if we return TRUE
		flags.isBeingRestored = FALSE;
	}
	return lr;
}

LRESULT ToolBar::OnBeginDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTOOLBAR pDetails = reinterpret_cast<LPNMTOOLBAR>(pNotificationDetails);
	//ATLASSERT_POINTER(pDetails, NMTOOLBAR);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);
	DWORD position = GetMessagePos();
	POINT mousePosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
	UINT hitTestDetails = 0;
	HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, *this);

	/* NOTE: TBN_BEGINDRAG is sent on mouse-down, which is not what we want. But calling DragDetect here has
	 *       unwanted side effects (important messages are not sent at all, others are sent twice).
	 */
	CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(IDToButtonIndex(pDetails->iItem), FALSE, this);
	if(button & 2/*MouseButtonConstants.vbRightButton*/) {
		Raise_ButtonBeginRDrag(pTBarButton, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
	} else {
		Raise_ButtonBeginDrag(pTBarButton, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails));
	}

	return 0;
}

LRESULT ToolBar::OnCustHelpNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	Raise_DisplayCustomizationHelp();
	return 0;
}

LRESULT ToolBar::OnDeletingButtonNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTOOLBAR pDetails = reinterpret_cast<LPNMTOOLBAR>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTOOLBAR);
	if(!pDetails) {
		return 0;
	}

	CComPtr<IToolBarButton> pTBarButton = NULL;
	if(flags.buttonIsRemovedByCustomizationDialog) {
		pTBarButton = ClassFactory::InitToolBarButton(IDToButtonIndex(pDetails->iItem), FALSE, this);
		Raise_CustomizeDialogRemovingButton(pTBarButton);
	}
	if(!(properties.disabledEvents & deFreeButtonData)) {
		if(!pTBarButton) {
			pTBarButton = ClassFactory::InitToolBarButton(IDToButtonIndex(pDetails->iItem), FALSE, this);
		}
		Raise_FreeButtonData(pTBarButton);
	}
	return 0;
}

LRESULT ToolBar::OnDropDownNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTOOLBAR pDetails = reinterpret_cast<LPNMTOOLBAR>(pNotificationDetails);
	int buttonIndex = IDToButtonIndex(pDetails->iItem, (index == 6 ? chevronPopupToolbar : NULL));
	CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonIndex, index == 6, this);
	droppedDownButton = buttonIndex;

	HMENU hPopupMenu = NULL;
	if(properties.menuMode) {
		LONG h = NULL;
		Raise_ButtonGetDropDownMenu(pTBarButton, &h);
		hPopupMenu = static_cast<HMENU>(LongToHandle(h));
	}

	RECT tmpRectangle = pDetails->rcButton;
	::MapWindowPoints(pDetails->hdr.hwndFrom, *this, reinterpret_cast<LPPOINT>(&tmpRectangle), 2);

	DropDownReturnValuesConstants furtherProcessing = ddrvDoDefault;
	Raise_DropDown(pTBarButton, reinterpret_cast<RECTANGLE*>(&tmpRectangle), &furtherProcessing);

	if(properties.menuMode && furtherProcessing == ddrvDoDefault && IsMenu(hPopupMenu)) {
		CWindow topLevelParent = GetTopLevelParent();
		if(topLevelParent.IsWindow()) {
			topLevelParent.SendMessage(WM_CANCELMODE, 0, 0);
		}
		if(properties.focusOnClick || flags.allowSetFocus) {
			if(GetFocus() != *this) {
				TakeFocus();
			}
		}
		menuModeState.displayingPopupMenu = FALSE;
		menuModeState.escapePressed = FALSE;
		DoPopupMenu(buttonIndex, hPopupMenu, TRUE);
	} else if(flags.chevronPopupVisible && parentWindowChevronPopupMenu.IsWindow()) {
		// this is important if we have a drop-down button in the chevron popup window
		parentWindowChevronPopupMenu.SendMessage(WM_CANCELMODE, 0, 0);
	}
	droppedDownButton = -1;

	return furtherProcessing;
}

LRESULT ToolBar::OnDupAcceleratorNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTBDUPACCELERATOR pDetails = reinterpret_cast<LPNMTBDUPACCELERATOR>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTBDUPACCELERATOR);
	if(!pDetails) {
		return FALSE;
	}

	VARIANT_BOOL handledEvent = VARIANT_FALSE;
	if(!(properties.disabledEvents & deAcceleratorEvents)) {
		VARIANT_BOOL isDuplicate = VARIANT_FALSE;
		Raise_IsDuplicateAccelerator(static_cast<SHORT>(pDetails->ch), &isDuplicate, &handledEvent);
		if(handledEvent != VARIANT_FALSE) {
			pDetails->fDup = VARIANTBOOL2BOOL(isDuplicate);
		}
	}
	return VARIANTBOOL2BOOL(handledEvent);
}

LRESULT ToolBar::OnEndAdjustNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	Raise_EndCustomization();
	return 0;
}

LRESULT ToolBar::OnGetButtonInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTOOLBAR pDetails = reinterpret_cast<LPNMTOOLBAR>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TBN_GETBUTTONINFOA) {
			ATLASSERT_POINTER(pDetails, NMTOOLBARA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTOOLBARW);
		}
	#endif

	CComObject<VirtualToolBarButton>* pVTBarButtonObj = NULL;
	CComPtr<IVirtualToolBarButton> pVTBarButtonItf = NULL;
	CComObject<VirtualToolBarButton>::CreateInstance(&pVTBarButtonObj);
	pVTBarButtonObj->AddRef();
	pVTBarButtonObj->SetOwner(this);
	pVTBarButtonObj->properties.readOnly = FALSE;

	pDetails->tbButton.dwData = 0;
	pDetails->tbButton.iBitmap = 0;
	pDetails->tbButton.idCommand = 0;
	pDetails->tbButton.fsState = TBSTATE_ENABLED;
	pDetails->tbButton.fsStyle = BTNS_AUTOSIZE;
	pDetails->tbButton.iString = NULL;
	pVTBarButtonObj->LoadState(&pDetails->tbButton, pDetails->iItem);

	pVTBarButtonObj->QueryInterface(IID_IVirtualToolBarButton, reinterpret_cast<LPVOID*>(&pVTBarButtonItf));

	VARIANT_BOOL noMoreButtons = VARIANT_TRUE;
	Raise_GetAvailableButton(pVTBarButtonItf, &noMoreButtons);
	if(noMoreButtons == VARIANT_FALSE) {
		TBBUTTON button = {0};
		pVTBarButtonObj->SaveState(&button, pDetails->iItem);

		CopyMemory(pDetails->tbButton.bReserved, button.bReserved, ARRAYSIZE(pDetails->tbButton.bReserved));
		pDetails->tbButton.dwData = button.dwData;
		pDetails->tbButton.fsState = button.fsState;
		pDetails->tbButton.fsStyle = button.fsStyle;
		if(flags.isBeingRestored/* && !SendMessage(TB_GETIMAGELIST, 0, 0)*/) {
			// the native tool bar deletes any buttons that don't have a bitmap
			pDetails->tbButton.iBitmap = max(0, button.iBitmap);
		} else {
			pDetails->tbButton.iBitmap = button.iBitmap;
		}
		pDetails->tbButton.idCommand = button.idCommand;
		if(pNotificationDetails->code == TBN_GETBUTTONINFOA) {
			if(button.iString == reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKA) || button.iString == reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKW)) {
				reinterpret_cast<LPNMTOOLBARA>(pDetails)->tbButton.iString = reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKA);
			} else if(!IS_INTRESOURCE(button.iString)) {
				/* NOTE: pDetails->tbButton.iString seems to be always Unicode! Unfortunatly the hack that we
				         in revision 193 doesn't really work, because each button gets the same string pointer and
				         this leaves the control in a unstable state. Therefore we use set the button text later.
				         The client app can prevent this situation by using the unicode version of the control and
				         making the parent window accept unicode notifications (as demonstrated by the samples). */
				PostMessage(GetDeferredSetButtonTextMessage(), button.idCommand, 0);
				free(reinterpret_cast<LPVOID>(button.iString));
				button.iString = NULL;
			} else {
				reinterpret_cast<LPNMTOOLBARA>(pDetails)->tbButton.iString = button.iString;
			}
		} else {
			if(button.iString == reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKA) || button.iString == reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKW)) {
				pDetails->tbButton.iString = reinterpret_cast<INT_PTR>(LPSTR_TEXTCALLBACKW);
			} else if(!IS_INTRESOURCE(button.iString)) {
				#ifdef UNICODE
					//ATLVERIFY(SUCCEEDED(StringCchCopy(reinterpret_cast<LPWSTR>(pDetails->tbButton.iString), pDetails->cchText, reinterpret_cast<LPCWSTR>(button.iString))));
					ATLVERIFY(SUCCEEDED(StringCchCopy(pDetails->pszText, pDetails->cchText, reinterpret_cast<LPCWSTR>(button.iString))));
				#else
					CA2W converter(reinterpret_cast<LPCSTR>(button.iString));
					//ATLVERIFY(SUCCEEDED(StringCchCopyW(reinterpret_cast<LPWSTR>(pDetails->tbButton.iString), pDetails->cchText, converter)));
					ATLVERIFY(SUCCEEDED(StringCchCopyW(reinterpret_cast<LPNMTOOLBARW>(pDetails)->pszText, reinterpret_cast<LPNMTOOLBARW>(pDetails)->cchText, converter)));
				#endif
				// NOTE: pDetails->tbButton.iString seems to be always Unicode!
				reinterpret_cast<LPNMTOOLBARW>(pDetails)->tbButton.iString = reinterpret_cast<INT_PTR>(reinterpret_cast<LPNMTOOLBARW>(pDetails)->pszText);
				free(reinterpret_cast<LPVOID>(button.iString));
				button.iString = NULL;
			} else {
				reinterpret_cast<LPNMTOOLBARW>(pDetails)->tbButton.iString = button.iString;
			}
		}
	}
	pVTBarButtonObj->Release();
	pVTBarButtonObj = NULL;

	return !VARIANTBOOL2BOOL(noMoreButtons);
}

LRESULT ToolBar::OnGetDispInfoNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTBDISPINFO pDetails = reinterpret_cast<LPNMTBDISPINFO>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TBN_GETDISPINFOA) {
			ATLASSERT_POINTER(pDetails, NMTBDISPINFOA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTBDISPINFOW);
		}
	#endif

	int buttonIndex = IDToButtonIndex(pDetails->idCommand, (index == 6 ? chevronPopupToolbar : NULL));
	CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonIndex, index == 6, this, FALSE);
	LONG iconIndex = 0;
	LONG imageListIndex = 0;
	CComBSTR buttonText = L"";
	LONG maxButtonTextLength = 0;

	RequestedInfoConstants requestedInfo = static_cast<RequestedInfoConstants>(0);
	if(pDetails->dwMask & TBNF_IMAGE) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riIconIndex);
	}
	if(pDetails->dwMask & TBNF_TEXT) {
		requestedInfo = static_cast<RequestedInfoConstants>(static_cast<long>(requestedInfo) | riButtonText);
		// exclude the terminating 0
		maxButtonTextLength = pDetails->cchText - 1;
	}

	VARIANT_BOOL dontAskAgain = VARIANT_FALSE;
	Raise_ButtonGetDisplayInfo(pTBarButton, requestedInfo, &iconIndex, &imageListIndex, maxButtonTextLength, &buttonText, &dontAskAgain);

	if(pDetails->dwMask & TBNF_IMAGE) {
		pDetails->iImage = MAKELONG(iconIndex, imageListIndex);
	}

	if(pDetails->dwMask & TBNF_TEXT) {
		// don't forget the terminating 0
		int bufferSize = buttonText.Length() + 1;
		if(bufferSize > pDetails->cchText) {
			bufferSize = pDetails->cchText;
		}
		if(pNotificationDetails->code == TBN_GETDISPINFOA) {
			lstrcpynA(reinterpret_cast<LPNMTBDISPINFOA>(pDetails)->pszText, CW2A(buttonText), bufferSize);
		} else {
			lstrcpynW(reinterpret_cast<LPNMTBDISPINFOW>(pDetails)->pszText, buttonText, bufferSize);
		}
	}

	if(dontAskAgain == VARIANT_FALSE) {
		pDetails->dwMask &= ~TBNF_DI_SETITEM;
	} else {
		pDetails->dwMask |= TBNF_DI_SETITEM;
	}
	return 0;
}

LRESULT ToolBar::OnGetInfoTipNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTBGETINFOTIP pDetails = reinterpret_cast<LPNMTBGETINFOTIP>(pNotificationDetails);
	#ifdef DEBUG
		if(pNotificationDetails->code == TBN_GETINFOTIPA) {
			ATLASSERT_POINTER(pDetails, NMTBGETINFOTIPA);
		} else {
			ATLASSERT_POINTER(pDetails, NMTBGETINFOTIPW);
		}
	#endif

	int buttonIndex = IDToButtonIndex(pDetails->iItem, (index == 6 ? chevronPopupToolbar : NULL));
	CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(buttonIndex, index == 6, this, FALSE);
	CComBSTR text = L"";
	if(pTBarButton) {
		if(pNotificationDetails->code == TBN_GETINFOTIPA) {
			if(lstrlenA(reinterpret_cast<LPNMTBGETINFOTIPA>(pDetails)->pszText) == 0) {
				pTBarButton->get_Text(&text);
			} else {
				text = reinterpret_cast<LPNMTBGETINFOTIPA>(pDetails)->pszText;
			}
		} else {
			if(lstrlenW(reinterpret_cast<LPNMTBGETINFOTIPW>(pDetails)->pszText) == 0) {
				pTBarButton->get_Text(&text);
			} else {
				text = reinterpret_cast<LPNMTBGETINFOTIPW>(pDetails)->pszText;
			}
		}
	} else {
		if(pNotificationDetails->code == TBN_GETINFOTIPA) {
			text = reinterpret_cast<LPNMTBGETINFOTIPA>(pDetails)->pszText;
		} else {
			text = reinterpret_cast<LPNMTBGETINFOTIPW>(pDetails)->pszText;
		}
	}
	// exclude the terminating 0
	LONG maxTextLength = pDetails->cchTextMax - 1;
	VARIANT_BOOL abortToolTip = VARIANT_FALSE;
	Raise_ButtonGetInfoTipText(pTBarButton, maxTextLength, &text, &abortToolTip);
	flags.cancelToolTip = VARIANTBOOL2BOOL(abortToolTip);
	
	// don't forget the terminating 0
	int bufferSize = text.Length() + 1;
	if(bufferSize > pDetails->cchTextMax) {
		bufferSize = pDetails->cchTextMax;
	}
	if(pNotificationDetails->code == TBN_GETINFOTIPA) {
		int requiredSize = WideCharToMultiByte(CP_ACP, 0, OLE2CW(text), bufferSize, "", 0, "", NULL);
		if(requiredSize > 0) {
			LPSTR pBuffer = static_cast<LPSTR>(HeapAlloc(GetProcessHeap(), 0, (requiredSize + 1) * sizeof(CHAR)));
			if(pBuffer) {
				WideCharToMultiByte(CP_ACP, 0, OLE2CW(text), bufferSize, pBuffer, requiredSize, "", NULL);
				/* On Vista and newer, many info tips contain Unicode characters that cannot be converted.
				 * We have told WideCharToMultiByte to replace them with \0. Now remove those \0.
				 */
				LPSTR pDst = reinterpret_cast<LPNMTBGETINFOTIPA>(pDetails)->pszText;
				ZeroMemory(pDst, pDetails->cchTextMax);
				int j = 0;
				for(int i = 0; i < min(bufferSize, requiredSize); i++) {
					if(pBuffer[i] != '\0') {
						pDst[j++] = pBuffer[i];
					}
				}
				HeapFree(GetProcessHeap(), 0, pBuffer);
			}
		}
	} else {
		lstrcpynW(reinterpret_cast<LPNMTBGETINFOTIPW>(pDetails)->pszText, OLE2CW(text), bufferSize);
	}

	return 0;
}

LRESULT ToolBar::OnGetObjectNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMOBJECTNOTIFY pDetails = reinterpret_cast<LPNMOBJECTNOTIFY>(pNotificationDetails);
	ATLASSERT(*(pDetails->piid) == IID_IDropTarget);
	pDetails->hResult = QueryInterface(*(pDetails->piid), &pDetails->pObject);
	return 0;
}

LRESULT ToolBar::OnHotItemChangeNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled)
{
	LPNMTBHOTITEM pDetails = reinterpret_cast<LPNMTBHOTITEM>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTBHOTITEM);
	if(!pDetails) {
		return FALSE;
	}

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!(properties.disabledEvents & deHotButtonChangeEvents)) {
		CComPtr<IToolBarButton> pPreviousHotButton = NULL;
		if(!(pDetails->dwFlags & HICF_ENTERING)) {
			pPreviousHotButton = ClassFactory::InitToolBarButton(IDToButtonIndex(pDetails->idOld, (index == 6 ? chevronPopupToolbar : NULL)), index == 6, this);
		}
		CComPtr<IToolBarButton> pNewHotButton = NULL;
		if(!(pDetails->dwFlags & HICF_LEAVING)) {
			pNewHotButton = ClassFactory::InitToolBarButton(IDToButtonIndex(pDetails->idNew, (index == 6 ? chevronPopupToolbar : NULL)), index == 6, this);
		}
		Raise_HotButtonChanging(pPreviousHotButton, pNewHotButton, static_cast<HotButtonChangingCausedByConstants>(pDetails->dwFlags & (HICF_OTHER | HICF_MOUSE | HICF_ARROWKEYS | HICF_ACCELERATOR)), static_cast<HotButtonChangingAdditionalInfoConstants>(pDetails->dwFlags & (HICF_DUPACCEL | HICF_ENTERING | HICF_LEAVING | HICF_RESELECT | HICF_LMOUSE | HICF_TOGGLEDROPDOWN)), &cancel);
	}

	if(cancel != VARIANT_FALSE) {
		return TRUE;
	}
	wasHandled = FALSE;
	return FALSE;
}

LRESULT ToolBar::OnInitCustomizeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTBCUSTOMIZEDLG pDetails = reinterpret_cast<LPNMTBCUSTOMIZEDLG>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTBCUSTOMIZEDLG);
	if(!pDetails) {
		return 0;
	}

	VARIANT_BOOL displayHelpButton = VARIANT_FALSE;
	Raise_InitializeCustomizationDialog(HandleToLong(pDetails->hDlg), &displayHelpButton);
	return (VARIANTBOOL2BOOL(displayHelpButton) ? 0 : TBNRF_HIDEHELP);
}

LRESULT ToolBar::OnMapAcceleratorNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMCHAR pDetails = reinterpret_cast<LPNMCHAR>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMCHAR);
	if(!pDetails) {
		return FALSE;
	}

	VARIANT_BOOL handledEvent = VARIANT_FALSE;
	if(!(properties.disabledEvents & deAcceleratorEvents)) {
		CComPtr<IToolBarButton> pStartingPointOfSearch = ClassFactory::InitToolBarButton(pDetails->dwItemPrev, index == 6, this);
		IToolBarButton* pButton = NULL;
		Raise_MapAccelerator(static_cast<SHORT>(pDetails->ch), pStartingPointOfSearch, VARIANT_FALSE, &pButton, &handledEvent);
		if(handledEvent != VARIANT_FALSE) {
			// return the button that was set by the event handler
			LONG buttonIndex = -1;
			if(pButton) {
				pButton->get_Index(&buttonIndex);
			}
			pDetails->dwItemNext = buttonIndex;
		}
		if(pButton) {
			pButton->Release();
		}
	}
	return VARIANTBOOL2BOOL(handledEvent);
}

LRESULT ToolBar::OnQueryDeleteNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTOOLBAR pDetails = reinterpret_cast<LPNMTOOLBAR>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTOOLBAR);
	if(!pDetails) {
		return FALSE;
	}

	CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(pDetails->iItem, FALSE, this);
	VARIANT_BOOL allowRemoval = VARIANT_TRUE;
	Raise_QueryRemoveButton(pTBarButton, &allowRemoval);
	return VARIANTBOOL2BOOL(allowRemoval);
}

LRESULT ToolBar::OnQueryInsertNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTOOLBAR pDetails = reinterpret_cast<LPNMTOOLBAR>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTOOLBAR);
	if(!pDetails) {
		return FALSE;
	}

	CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(pDetails->iItem, FALSE, this);
	VARIANT_BOOL allowInsertionToLeft = VARIANT_TRUE;
	Raise_QueryInsertButton(pTBarButton, &allowInsertionToLeft);
	return VARIANTBOOL2BOOL(allowInsertionToLeft);
}

LRESULT ToolBar::OnResetNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTBCUSTOMIZEDLG pDetails = reinterpret_cast<LPNMTBCUSTOMIZEDLG>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTBCUSTOMIZEDLG);
	if(!pDetails) {
		return 0;
	}

	VARIANT_BOOL endCustomization = VARIANT_FALSE;
	Raise_ResetCustomizations(HandleToLong(pDetails->hDlg), &endCustomization);
	return (VARIANTBOOL2BOOL(endCustomization) ? TBNRF_ENDCUSTOMIZE : 0);
}

LRESULT ToolBar::OnRestoreNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTBRESTORE pDetails = reinterpret_cast<LPNMTBRESTORE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTBRESTORE);
	if(!pDetails) {
		return FALSE;
	}

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(pDetails->iItem == -1) {
		// this is the initial notification
		Raise_InitializeToolBarStateRegistryRestorage(reinterpret_cast<LONG*>(&pDetails->cButtons), pDetails->cbData, pDetails->cbBytesPerRecord, reinterpret_cast<LONG>(pDetails->pData), reinterpret_cast<LONG*>(&pDetails->pCurrent), &cancel);
	} else {
		ATLASSERT(pDetails->iItem >= 0);

		CComObject<VirtualToolBarButton>* pVTBarButtonObj = NULL;
		CComPtr<IVirtualToolBarButton> pVTBarButtonItf = NULL;
		CComObject<VirtualToolBarButton>::CreateInstance(&pVTBarButtonObj);
		pVTBarButtonObj->AddRef();
		pVTBarButtonObj->properties.readOnly = FALSE;
		pVTBarButtonObj->SetOwner(this);
		pDetails->tbButton.iString = NULL;
		pVTBarButtonObj->LoadState(&pDetails->tbButton, pDetails->iItem);
		pVTBarButtonObj->QueryInterface(IID_IVirtualToolBarButton, reinterpret_cast<LPVOID*>(&pVTBarButtonItf));

		Raise_RestoreButtonFromRegistryStream(pVTBarButtonItf, pDetails->cButtons, pDetails->cbData, pDetails->cbBytesPerRecord, reinterpret_cast<LONG>(pDetails->pData), reinterpret_cast<LONG*>(&pDetails->pCurrent));

		pVTBarButtonObj->SaveState(&pDetails->tbButton, pDetails->iItem);
		pVTBarButtonObj->Release();
	}

	return VARIANTBOOL2BOOL(cancel);
}

LRESULT ToolBar::OnSaveNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTBSAVE pDetails = reinterpret_cast<LPNMTBSAVE>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTBSAVE);
	if(!pDetails) {
		return 0;
	}

	if(pDetails->iItem == -1) {
		// this is the initial notification
		Raise_InitializeToolBarStateRegistryStorage(reinterpret_cast<LONG*>(&pDetails->cButtons), reinterpret_cast<LONG*>(&pDetails->cbData), reinterpret_cast<LONG*>(&pDetails->pData), reinterpret_cast<LONG*>(&pDetails->pCurrent));
	} else {
		ATLASSERT(pDetails->iItem >= 0);

		CComObject<VirtualToolBarButton>* pVTBarButtonObj = NULL;
		CComPtr<IVirtualToolBarButton> pVTBarButtonItf = NULL;
		CComObject<VirtualToolBarButton>::CreateInstance(&pVTBarButtonObj);
		pVTBarButtonObj->AddRef();
		pVTBarButtonObj->SetOwner(this);
		pVTBarButtonObj->LoadState(&pDetails->tbButton, pDetails->iItem);
		pVTBarButtonObj->QueryInterface(IID_IVirtualToolBarButton, reinterpret_cast<LPVOID*>(&pVTBarButtonItf));
		pVTBarButtonObj->Release();

		Raise_SaveButtonToRegistryStream(pVTBarButtonItf, pDetails->cbData, reinterpret_cast<LONG>(pDetails->pData), reinterpret_cast<LONG*>(&pDetails->pCurrent));
	}

	return 0;
}

LRESULT ToolBar::OnToolBarChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/)
{
	Raise_CustomizedControl();
	return 0;
}

LRESULT ToolBar::OnWrapAcceleratorNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTBWRAPACCELERATOR pDetails = reinterpret_cast<LPNMTBWRAPACCELERATOR>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTBWRAPACCELERATOR);
	if(!pDetails) {
		return FALSE;
	}

	VARIANT_BOOL handledEvent = VARIANT_FALSE;
	if(!(properties.disabledEvents & deAcceleratorEvents)) {
		IToolBarButton* pButton = NULL;
		Raise_MapAccelerator(static_cast<SHORT>(pDetails->ch), NULL, VARIANT_TRUE, &pButton, &handledEvent);
		if(handledEvent != VARIANT_FALSE) {
			// return the button that was set by the event handler
			LONG buttonIndex = -1;
			if(pButton) {
				pButton->get_Index(&buttonIndex);
			}
			pDetails->iButton = buttonIndex;
		}
		if(pButton) {
			pButton->Release();
		}
	}
	return VARIANTBOOL2BOOL(handledEvent);
}

LRESULT ToolBar::OnWrapHotItemNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/)
{
	LPNMTBWRAPHOTITEM pDetails = reinterpret_cast<LPNMTBWRAPHOTITEM>(pNotificationDetails);
	ATLASSERT_POINTER(pDetails, NMTBWRAPHOTITEM);
	if(!pDetails) {
		return FALSE;
	}

	VARIANT_BOOL cancel = VARIANT_FALSE;
	if(!(properties.disabledEvents & deHotButtonChangeEvents)) {
		CComPtr<IToolBarButton> pPreviousHotButton = ClassFactory::InitToolBarButton(IDToButtonIndex(pDetails->iStart, (index == 6 ? chevronPopupToolbar : NULL)), index == 6, this);
		Raise_HotButtonChangeWrapping(pPreviousHotButton, static_cast<WrappingDirectionConstants>(pDetails->iDir), static_cast<HotButtonChangingCausedByConstants>(pDetails->nReason & (HICF_OTHER | HICF_MOUSE | HICF_ARROWKEYS | HICF_ACCELERATOR)), &cancel);
	}
	return VARIANTBOOL2BOOL(cancel);
}


LRESULT ToolBar::OnReflectedClicked(LONG index, WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& wasHandled)
{
	CComPtr<IToolBarButton> pTBarButton = ClassFactory::InitToolBarButton(IDToButtonIndex(controlID, (index == 6 ? chevronPopupToolbar : NULL)), index == 6, this);
	VARIANT_BOOL forward = VARIANT_FALSE;
	Raise_ExecuteCommand(controlID, pTBarButton, coButton, &forward, NULL);
	if(forward != VARIANT_FALSE && IsWindow()) {
		CWindow parent = GetParent();
		if(parent.IsWindow()) {
			CWindow topLevelParent = GetTopLevelParent();
			if(topLevelParent.IsWindow() && static_cast<HWND>(topLevelParent) != static_cast<HWND>(parent)) {
				topLevelParent.SendMessage(WM_COMMAND, MAKEWPARAM(controlID, BN_CLICKED), reinterpret_cast<LPARAM>(static_cast<HWND>(*this)));
			}
		}
	}
	wasHandled = FALSE;
	return 0;
}

LRESULT ToolBar::OnCustomDrawNotification(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	LPNMTBCUSTOMDRAW pDetails = reinterpret_cast<LPNMTBCUSTOMDRAW>(lParam);
	CustomDrawReturnValuesConstants returnValue;
	if(index == 6) {
		returnValue = static_cast<CustomDrawReturnValuesConstants>(chevronPopupToolbar.DefWindowProc(message, wParam, lParam));
	} else {
		returnValue = static_cast<CustomDrawReturnValuesConstants>(this->DefWindowProc(message, wParam, lParam));
	}

	int buttonIndex = -1;
	if(properties.drawTextFlags & (DT_CENTER | DT_RIGHT)) {
		CWindow wnd(index == 6 ? chevronPopupToolbar : static_cast<HWND>(*this));
		if(wnd.GetStyle() & TBSTYLE_LIST) {
			if(pDetails->nmcd.dwDrawStage & (CDDS_ITEM | CDDS_SUBITEM)) {
				buttonIndex = IDToButtonIndex(static_cast<LONG>(pDetails->nmcd.dwItemSpec), (index == 6 ? chevronPopupToolbar : NULL));
			}
			if(RunTimeHelper::IsCommCtrl6()) {
				// Comctl32.dll 6.x does not align the text properly if the button contains text only and the control is in list mode
				if(pDetails->nmcd.dwDrawStage == CDDS_PREPAINT) {
					returnValue = static_cast<CustomDrawReturnValuesConstants>(CDRF_NOTIFYITEMDRAW);
				} else if(pDetails->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
					int iconIndex = 0;
					TBBUTTONINFO button = {0};
					button.cbSize = sizeof(button);
					button.dwMask = TBIF_BYINDEX | TBIF_IMAGE | TBIF_STYLE;
					if(::SendMessage((index == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)), TB_GETBUTTONINFO, buttonIndex, reinterpret_cast<LPARAM>(&button)) == buttonIndex) {
						if(button.fsStyle & BTNS_SEP) {
							iconIndex = I_IMAGENONE;
						} else if(button.iImage == I_IMAGENONE) {
							iconIndex = I_IMAGENONE;
						} else {
							iconIndex = LOWORD(button.iImage);
						}
					}
					if(iconIndex == I_IMAGENONE) {
						pDetails->rcText.right = pDetails->nmcd.rc.right - pDetails->nmcd.rc.left;
						pDetails->rcText.left = 0;
					}
				}
			}
		}
	}

	BOOL drawInactive = FALSE;
	if(properties.menuMode) {
		if(properties.menuBarTheme == mbtNativeMenuBar) {
			if(!RunTimeHelper::IsVista()) {
				// Windows XP or older
				if(pDetails->nmcd.dwDrawStage == CDDS_PREPAINT) {
					returnValue = static_cast<CustomDrawReturnValuesConstants>(CDRF_NOTIFYITEMDRAW);
				} else if(pDetails->nmcd.dwDrawStage == CDDS_ITEMPREPAINT) {
					if(parentWindowMenuMode.IsWindow()) {
						HWND hWndForeground = GetForegroundWindow();
						drawInactive = (hWndForeground != parentWindowMenuMode && !::IsChild(parentWindowMenuMode, hWndForeground));
					}
					if((pDetails->nmcd.uItemState & (CDIS_HOT | CDIS_SELECTED)) || drawInactive) {
						pDetails->nHLStringBkMode = TRANSPARENT;
					}
					/* This would be enough if we could live with non-pressed items having black text.
					colorHotButtonBackground = GetSysColor(COLOR_MENUHILIGHT);
					pDetails->nHLStringBkMode = TRANSPARENT;
					returnValue = static_cast<CustomDrawReturnValuesConstants>(TBCDRF_HILITEHOTTRACK);*/
				}
			}
		}
	}

	if(!(properties.disabledEvents & deCustomDraw)) {
		CComPtr<IToolBarButton> pTBarButton = NULL;
		if(pDetails->nmcd.dwDrawStage & (CDDS_ITEM | CDDS_SUBITEM)) {
			if(buttonIndex == -1) {
				buttonIndex = IDToButtonIndex(static_cast<LONG>(pDetails->nmcd.dwItemSpec), (index == 6 ? chevronPopupToolbar : NULL));
			}
			if(pDetails->nmcd.dwDrawStage & CDDS_ITEM) {
				pTBarButton = ClassFactory::InitToolBarButton(buttonIndex, index == 6, this, FALSE);
			}
		}
		OLE_COLOR colorNormalText = pDetails->clrText;
		OLE_COLOR colorNormalButtonBackground = pDetails->clrBtnFace;
		OLE_COLOR colorHotText = pDetails->clrTextHighlight;
		OLE_COLOR colorHotButtonBackground = pDetails->clrHighlightHotTrack;
		OLE_COLOR colorMarkedTextBackground = pDetails->clrMark;
		OLE_COLOR colorMarkedButtonBackground = pDetails->clrBtnHighlight;

		CustomDrawStageConstants drawStage = static_cast<CustomDrawStageConstants>(pDetails->nmcd.dwDrawStage | (index == 6 ? cdsChevronPopupToolbar : 0));
		Raise_CustomDraw(pTBarButton, &colorNormalText, &colorNormalButtonBackground, reinterpret_cast<StringBackgroundModeConstants*>(&pDetails->nStringBkMode), &colorHotText, &colorHotButtonBackground, &colorMarkedTextBackground, &colorMarkedButtonBackground, reinterpret_cast<StringBackgroundModeConstants*>(&pDetails->nHLStringBkMode), drawStage, static_cast<CustomDrawItemStateConstants>(pDetails->nmcd.uItemState), HandleToLong(pDetails->nmcd.hdc), reinterpret_cast<RECTANGLE*>(&pDetails->nmcd.rc), reinterpret_cast<RECTANGLE*>(&pDetails->rcText), reinterpret_cast<OLE_XSIZE_PIXELS*>(&pDetails->iListGap), &returnValue);

		pDetails->clrText = colorNormalText;
		pDetails->clrBtnFace = colorNormalButtonBackground;
		pDetails->clrTextHighlight = colorHotText;
		pDetails->clrHighlightHotTrack = colorHotButtonBackground;
		pDetails->clrMark = colorMarkedTextBackground;
		pDetails->clrBtnHighlight = colorMarkedButtonBackground;
	}
	
	if(pDetails->nmcd.dwDrawStage == CDDS_PREERASE && !(returnValue & CDRF_SKIPDEFAULT) && flags.applyBackgroundHack) {
		/* A transparent tool bar with classic theme won't draw its background (of course), which
		   makes transparent tool bars a bit tricky. */
		CWindow wnd(index == 6 ? chevronPopupToolbar : static_cast<HWND>(*this));
		/* TODO: Maybe we should use the drawing rectangle on comctl32.dll 6.x?! On older versions,
		 *       the drawing rectangle seems to be invalid, so we should use the client rectangle then.
		 */
		RECT clientRectangle = {0};
		wnd.GetClientRect(&clientRectangle);
		FillRect(pDetails->nmcd.hdc, &clientRectangle, GetSysColorBrush(COLOR_3DFACE));
	}

	if(properties.menuMode) {
		if(properties.menuBarTheme == mbtNativeMenuBar) {
			if(!RunTimeHelper::IsVista()) {
				// Windows XP or older
				if(pDetails->nmcd.dwDrawStage == CDDS_ITEMPREPAINT && returnValue == CDRF_DODEFAULT) {
					if((pDetails->nmcd.uItemState & (CDIS_HOT | CDIS_SELECTED)) || drawInactive) {
						returnValue = static_cast<CustomDrawReturnValuesConstants>(CDRF_SKIPDEFAULT);

						if(pDetails->nmcd.uItemState & (CDIS_HOT | CDIS_SELECTED)) {
							FillRect(pDetails->nmcd.hdc, &pDetails->nmcd.rc, GetSysColorBrush(COLOR_MENUHILIGHT));
							FrameRect(pDetails->nmcd.hdc, &pDetails->nmcd.rc, GetSysColorBrush(COLOR_HIGHLIGHT));
						}

						TBBUTTONINFO button = {0};
						button.cbSize = sizeof(button);
						// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
						button.dwMask = TBIF_BYINDEX | TBIF_COMMAND;
						if(::SendMessage((index == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)), TB_GETBUTTONINFO, buttonIndex, reinterpret_cast<LPARAM>(&button)) == buttonIndex) {
							int textLength = static_cast<int>(::SendMessage((index == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)), TB_GETBUTTONTEXT, button.idCommand, NULL));
							if(textLength > -1) {
								LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
								if(pBuffer) {
									ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
									if(static_cast<int>(::SendMessage((index == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)), TB_GETBUTTONTEXT, button.idCommand, reinterpret_cast<LPARAM>(pBuffer))) > -1) {
										int oldBkMode = SetBkMode(pDetails->nmcd.hdc, pDetails->nHLStringBkMode);
										HGDIOBJ hOldFont = SelectObject(pDetails->nmcd.hdc, properties.font.currentFont);
										COLORREF oldTextColor = SetTextColor(pDetails->nmcd.hdc, (pDetails->nmcd.uItemState & (CDIS_HOT | CDIS_SELECTED)) ? pDetails->clrTextHighlight : GetSysColor(COLOR_GRAYTEXT));

										DrawText(pDetails->nmcd.hdc, pBuffer, textLength, &pDetails->nmcd.rc, DT_SINGLELINE | (properties.drawTextFlags & properties.drawTextFlagsMask));

										SetTextColor(pDetails->nmcd.hdc, oldTextColor);
										SelectObject(pDetails->nmcd.hdc, hOldFont);
										SetBkMode(pDetails->nmcd.hdc, oldBkMode);
									}
									HeapFree(GetProcessHeap(), 0, pBuffer);
								}
							}
						}
					}
				}
			}
		}
	}

	return static_cast<LRESULT>(returnValue);
}

LRESULT ToolBar::OnToolTipGetDispInfoNotificationA(BOOL chevronPopup, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	NMTTDISPINFOA backup = *reinterpret_cast<LPNMTTDISPINFOA>(lParam);
	LRESULT lr;
	if(chevronPopup) {
		lr = chevronPopupToolbar.DefWindowProc(message, wParam, lParam);
	} else {
		lr = DefWindowProc(message, wParam, lParam);
	}
	if(flags.cancelToolTip) {
		flags.cancelToolTip = FALSE;
		/* TODO: Find a better way than copying the old values back. Don't we produce a memleak here?
		         Who frees the string buffer (reinterpret_cast<LPNMTTDISPINFO>(lParam)->lpszText) that the
		         tool bar set up? */
		*reinterpret_cast<LPNMTTDISPINFOA>(lParam) = backup;
		lr = 0;     // actually the return value is ignored
	}

	return lr;
}

LRESULT ToolBar::OnToolTipGetDispInfoNotificationW(BOOL chevronPopup, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/)
{
	NMTTDISPINFOW backup = *reinterpret_cast<LPNMTTDISPINFOW>(lParam);
	LRESULT lr;
	if(chevronPopup) {
		lr = chevronPopupToolbar.DefWindowProc(message, wParam, lParam);
	} else {
		lr = DefWindowProc(message, wParam, lParam);
	}
	if(flags.cancelToolTip) {
		flags.cancelToolTip = FALSE;
		/* TODO: Find a better way than copying the old values back. Don't we produce a memleak here?
		         Who frees the string buffer (reinterpret_cast<LPNMTTDISPINFO>(lParam)->lpszText) that the
		         tool bar set up? */
		*reinterpret_cast<LPNMTTDISPINFOW>(lParam) = backup;
		lr = 0;     // actually the return value is ignored
	}

	return lr;
}


void ToolBar::ApplyFont(void)
{
	properties.font.dontGetFontObject = TRUE;
	if(IsWindow()) {
		if(!properties.font.owningFontResource) {
			properties.font.currentFont.Detach();
		}
		properties.font.currentFont.Attach(NULL);

		if(properties.useSystemFont) {
			// retrieve the system font
			if(properties.menuMode) {
				NONCLIENTMETRICS nonClientMetrics = {0};
				nonClientMetrics.cbSize = RunTimeHelper::SizeOf_NONCLIENTMETRICS();
				SystemParametersInfo(SPI_GETNONCLIENTMETRICS, nonClientMetrics.cbSize, &nonClientMetrics, 0);
				properties.font.currentFont.CreateFontIndirect(&nonClientMetrics.lfMenuFont);
			} else {
				LOGFONT iconFont = {0};
				SystemParametersInfo(SPI_GETICONTITLELOGFONT, sizeof(LOGFONT), &iconFont, 0);
				properties.font.currentFont.CreateFontIndirect(&iconFont);
			}
			properties.font.owningFontResource = TRUE;

			// apply the font
			SendMessage(WM_SETFONT, reinterpret_cast<WPARAM>(static_cast<HFONT>(properties.font.currentFont)), MAKELPARAM(TRUE, 0));
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

void ToolBar::SetDrawTextFlags(void)
{
	UINT flags = properties.drawTextFlags;
	switch(properties.horizontalTextAlignment) {
		case halLeft:
			flags &= ~(DT_CENTER | DT_RIGHT);
			flags |= DT_LEFT;
			break;
		case halCenter:
			flags &= ~(DT_LEFT | DT_RIGHT);
			flags |= DT_CENTER;
			break;
		case halRight:
			flags &= ~(DT_LEFT | DT_CENTER);
			flags |= DT_RIGHT;
			break;
	}
	if(properties.useMnemonics) {
		flags &= ~DT_NOPREFIX;
	} else {
		flags |= DT_NOPREFIX;
	}
	if(properties.menuMode) {
		if(menuModeState.displayKeyboardCues) {
			flags &= ~DT_HIDEPREFIX;
		} else {
			flags |= DT_HIDEPREFIX;
		}
	}
	switch(properties.verticalTextAlignment) {
		case valTop:
			flags &= ~(DT_VCENTER | DT_BOTTOM);
			flags |= DT_TOP;
			break;
		case valCenter:
			flags &= ~(DT_TOP | DT_BOTTOM);
			flags |= DT_VCENTER;
			break;
		case valBottom:
			flags &= ~(DT_TOP | DT_VCENTER);
			flags |= DT_BOTTOM;
			break;
	}
	UINT mask = DT_LEFT | DT_CENTER | DT_RIGHT | DT_NOPREFIX | DT_TOP | DT_VCENTER | DT_BOTTOM;
	if(properties.menuMode) {
		mask |= DT_HIDEPREFIX;
	}
	SendMessage(TB_SETDRAWTEXTFLAGS, mask, flags);
}


inline HRESULT ToolBar::Raise_BeforeDisplayChevronPopup(LONG hPopup, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL isMenu, VARIANT_BOOL* pCancelPopup, LONG* pCommandToExecute)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_BeforeDisplayChevronPopup(hPopup, x, y, isMenu, pCancelPopup, pCommandToExecute);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_BeginCustomization(VARIANT_BOOL restoringFromRegistry, VARIANT_BOOL* pDontRestoreFromRegistry)
{
	flags.isBeingCustomized = TRUE;
	flags.buttonIsRemovedByCustomizationDialog = TRUE;
	//if(m_nFreezeEvents == 0) {
		return Fire_BeginCustomization(restoringFromRegistry, pDontRestoreFromRegistry);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_ButtonBeginDrag(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ButtonBeginDrag(pToolButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_ButtonBeginRDrag(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ButtonBeginRDrag(pToolButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_ButtonGetDisplayInfo(IToolBarButton* pToolButton, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pImageListIndex, LONG maxButtonTextLength, BSTR* pButtonText, VARIANT_BOOL* pDontAskAgain)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ButtonGetDisplayInfo(pToolButton, requestedInfo, pIconIndex, pImageListIndex, maxButtonTextLength, pButtonText, pDontAskAgain);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_ButtonGetDropDownMenu(IToolBarButton* pToolButton, LONG* phMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ButtonGetDropDownMenu(pToolButton, phMenu);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_ButtonGetInfoTipText(IToolBarButton* pToolButton, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ButtonGetInfoTipText(pToolButton, maxInfoTipLength, pInfoTipText, pAbortToolTip);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_ButtonMouseEnter(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_ButtonMouseEnter(pToolButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT ToolBar::Raise_ButtonMouseLeave(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails)
{
	if(/*(m_nFreezeEvents == 0) && */mouseStatus.enteredControl) {
		return Fire_ButtonMouseLeave(pToolButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	}
	return S_OK;
}

inline HRESULT ToolBar::Raise_ButtonSelectionStateChanged(IToolBarButton* pToolButton, SelectionStateConstants previousSelectionState, SelectionStateConstants newSelectionState)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ButtonSelectionStateChanged(pToolButton, previousSelectionState, newSelectionState);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_Click(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedButton = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));

	//if(m_nFreezeEvents == 0) {
		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(mouseStatus.lastClickedButton, messageMapIndex == 6, this);
		return Fire_Click(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_ContextMenu(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	//if(m_nFreezeEvents == 0) {
		if((x == -1) && (y == -1)) {
			// the event was caused by the keyboard
			if(properties.processContextMenuKeys) {
				HWND hWndToUse = (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this));
				int hotButton = static_cast<int>(::SendMessage(hWndToUse, TB_GETHOTITEM, 0, 0));
				CRect buttonRectangle;
				if(::SendMessage(hWndToUse, TB_GETITEMRECT, hotButton, reinterpret_cast<LPARAM>(&buttonRectangle))) {
					CPoint centerPoint = buttonRectangle.CenterPoint();
					x = centerPoint.x;
					y = centerPoint.y;
				}
			} else {
				return S_OK;
			}
		}

		UINT hitTestDetails = 0;
		int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));

		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
		return Fire_ContextMenu(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_CustomDraw(IToolBarButton* pToolButton, OLE_COLOR* pNormalTextColor, OLE_COLOR* pNormalButtonBackColor, StringBackgroundModeConstants* pNormalBackgroundMode, OLE_COLOR* pHotTextColor, OLE_COLOR* pHotButtonBackColor, OLE_COLOR* pMarkedTextBackColor, OLE_COLOR* pMarkedButtonBackColor, StringBackgroundModeConstants* pMarkedBackgroundMode, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants buttonState, LONG hDC, RECTANGLE* pDrawingRectangle, RECTANGLE* pTextRectangle, OLE_XSIZE_PIXELS* pHorizontalIconCaptionGap, CustomDrawReturnValuesConstants* pFurtherProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CustomDraw(pToolButton, pNormalTextColor, pNormalButtonBackColor, pNormalBackgroundMode, pHotTextColor, pHotButtonBackColor, pMarkedTextBackColor, pMarkedButtonBackColor, pMarkedBackgroundMode, drawStage, buttonState, hDC, pDrawingRectangle, pTextRectangle, pHorizontalIconCaptionGap, pFurtherProcessing);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_CustomizedControl(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CustomizedControl();
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_CustomizeDialogRemovingButton(IToolBarButton* pToolButton)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_CustomizeDialogRemovingButton(pToolButton);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_DblClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));
	if(buttonIndex != mouseStatus.lastClickedButton) {
		buttonIndex = -1;
	}
	mouseStatus.lastClickedButton = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
		return Fire_DblClick(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_DestroyedControlWindow(LONG hWnd)
{
	KillTimer(timers.ID_REDRAW);
	KillTimer(timers.ID_PRESELECTCHEVRONMENUITEM);
	RemoveButtonFromButtonContainers(-1);
	DeregisterPlaceholderButton(-1);
	if(properties.registerForOLEDragDrop == rfoddAdvancedDragDrop) {
		ATLVERIFY(RevokeDragDrop(*this) == S_OK);
	}

	if(parentWindowChevronPopupMenu.IsWindow()) {
		ATLVERIFY(RemoveWindowSubclass(parentWindowChevronPopupMenu, ToolBar::ParentWindowSubclass_ChevronPopupMenu, reinterpret_cast<UINT_PTR>(this)));
		parentWindowChevronPopupMenu.Detach();
	}
	if(properties.menuMode) {
		if(parentWindowMenuMode.IsWindow()) {
			ATLVERIFY(RemoveWindowSubclass(parentWindowMenuMode, ToolBar::ParentWindowSubclass_MenuMode, reinterpret_cast<UINT_PTR>(this)));
			parentWindowMenuMode.Detach();
		}
		if(menuModeState.attachedToMessageHook) {
			RemoveMessageHookClient();
			menuModeState.attachedToMessageHook = FALSE;
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyedControlWindow(hWnd);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_DestroyingChevronPopup(LONG hPopup, VARIANT_BOOL isMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DestroyingChevronPopup(hPopup, isMenu);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_DisplayCustomizationHelp(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DisplayCustomizationHelp();
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_DropDown(IToolBarButton* pToolButton, RECTANGLE* pButtonRectangle, DropDownReturnValuesConstants* pFurtherProcessing)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_DropDown(pToolButton, pButtonRectangle, pFurtherProcessing);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_EndCustomization(void)
{
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		hr = Fire_EndCustomization();
	//}
	flags.isBeingCustomized = FALSE;
	flags.buttonIsRemovedByCustomizationDialog = FALSE;
	flags.isBeingRestored = FALSE;
	return hr;
}

inline HRESULT ToolBar::Raise_ExecuteCommand(LONG commandID, IToolBarButton* pToolButton, CommandOriginConstants commandOrigin, VARIANT_BOOL* pForwardMessage, LPBOOL pCommandIsDisabled)
{
	#ifdef USE_STL
		for(std::vector<LONG>::iterator it = properties.disabledCommands.begin(); it != properties.disabledCommands.end(); it++) {
			if(*it == commandID) {
				if(pCommandIsDisabled) {
					*pCommandIsDisabled = TRUE;
				}
				*pForwardMessage = VARIANT_FALSE;
				return S_OK;
			}
		}
	#else
		for(size_t i = 0; i < properties.disabledCommands.GetCount(); ++i) {
			if(properties.disabledCommands[i] == commandID) {
				if(pCommandIsDisabled) {
					*pCommandIsDisabled = TRUE;
				}
				*pForwardMessage = VARIANT_FALSE;
				return S_OK;
			}
		}
	#endif
	//if(m_nFreezeEvents == 0) {
		return Fire_ExecuteCommand(commandID, pToolButton, commandOrigin, pForwardMessage);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_FreeButtonData(IToolBarButton* pToolButton)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_FreeButtonData(pToolButton);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_GetAvailableButton(IVirtualToolBarButton* pToolButton, VARIANT_BOOL* pNoMoreButtons)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_GetAvailableButton(pToolButton, pNoMoreButtons);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_HotButtonChangeWrapping(IToolBarButton* pPreviousHotButton, WrappingDirectionConstants wrappingDirection, HotButtonChangingCausedByConstants causedBy, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HotButtonChangeWrapping(pPreviousHotButton, wrappingDirection, causedBy, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_HotButtonChanging(IToolBarButton* pPreviousHotButton, IToolBarButton* pNewHotButton, HotButtonChangingCausedByConstants causedBy, HotButtonChangingAdditionalInfoConstants additionalInfo, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_HotButtonChanging(pPreviousHotButton, pNewHotButton, causedBy, additionalInfo, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_InitializeCustomizationDialog(LONG hCustomizationDialog, VARIANT_BOOL* pDisplayHelpButton)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InitializeCustomizationDialog(hCustomizationDialog, pDisplayHelpButton);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_InitializeToolBarStateRegistryRestorage(LONG* pNumberOfButtonsToLoad, LONG totalDataSize, LONG perButtonDataSize, LONG pDataStream, LONG* ppStartOfNextDataBlock, VARIANT_BOOL* pCancelLoading)
{
	flags.isBeingRestored = TRUE;
	//if(m_nFreezeEvents == 0) {
		return Fire_InitializeToolBarStateRegistryRestorage(pNumberOfButtonsToLoad, totalDataSize, perButtonDataSize, pDataStream, ppStartOfNextDataBlock, pCancelLoading);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_InitializeToolBarStateRegistryStorage(LONG* pNumberOfButtonsToSave, LONG* pTotalDataSize, LONG* ppDataStream, LONG* ppStartOfUnusedSpace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InitializeToolBarStateRegistryStorage(pNumberOfButtonsToSave, pTotalDataSize, ppDataStream, ppStartOfUnusedSpace);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_InsertedButton(IToolBarButton* pToolButton)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertedButton(pToolButton);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_InsertingButton(IVirtualToolBarButton* pToolButton, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_InsertingButton(pToolButton, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_IsDuplicateAccelerator(SHORT accelerator, VARIANT_BOOL* pIsDuplicate, VARIANT_BOOL* pHandledEvent)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_IsDuplicateAccelerator(accelerator, pIsDuplicate, pHandledEvent);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_KeyDown(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyDown(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_KeyPress(SHORT* pKeyAscii)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyPress(pKeyAscii);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_KeyUp(SHORT* pKeyCode, SHORT shift)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_KeyUp(pKeyCode, shift);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_MapAccelerator(SHORT accelerator, IToolBarButton* pStartingPointOfSearch, VARIANT_BOOL resumingSearchWithFirstButton, IToolBarButton** ppMatchingButton, VARIANT_BOOL* pHandledEvent)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_MapAccelerator(accelerator, pStartingPointOfSearch, resumingSearchWithFirstButton, ppMatchingButton, pHandledEvent);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_MClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedButton = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));

	//if(m_nFreezeEvents == 0) {
		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(mouseStatus.lastClickedButton, messageMapIndex == 6, this);
		return Fire_MClick(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_MDblClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));
	if(buttonIndex != mouseStatus.lastClickedButton) {
		buttonIndex = -1;
	}
	mouseStatus.lastClickedButton = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
		return Fire_MDblClick(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_MouseDown(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(messageMapIndex, button, shift, x, y);
	}
	if(!mouseStatus.hoveredControl) {
		TRACKMOUSEEVENT trackingOptions = {0};
		trackingOptions.cbSize = sizeof(trackingOptions);
		trackingOptions.hwndTrack = (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this));
		trackingOptions.dwFlags = TME_HOVER | TME_CANCEL;
		TrackMouseEvent(&trackingOptions);

		Raise_MouseHover(messageMapIndex, button, shift, x, y);
	}
	mouseStatus.StoreClickCandidate(button);

	//if(m_nFreezeEvents == 0) {
		UINT hitTestDetails = 0;
		int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));

		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
		return Fire_MouseDown(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_MouseEnter(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	TRACKMOUSEEVENT trackingOptions = {0};
	trackingOptions.cbSize = sizeof(trackingOptions);
	trackingOptions.hwndTrack = (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this));
	trackingOptions.dwHoverTime = (properties.hoverTime == -1 ? HOVER_DEFAULT : properties.hoverTime);
	trackingOptions.dwFlags = TME_HOVER | TME_LEAVE;
	TrackMouseEvent(&trackingOptions);

	mouseStatus.EnterControl();

	UINT hitTestDetails = 0;
	int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));
	buttonUnderMouse = buttonIndex;

	CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		Fire_MouseEnter(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	if(pTBButton) {
		Raise_ButtonMouseEnter(pTBButton, button, shift, x, y, hitTestDetails);
	}
	return hr;
}

inline HRESULT ToolBar::Raise_MouseHover(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.hoveredControl) {
		mouseStatus.HoverControl();

		//if(m_nFreezeEvents == 0) {
			UINT hitTestDetails = 0;
			int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));

			CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
			return Fire_MouseHover(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		//}
	}
	return S_OK;
}

inline HRESULT ToolBar::Raise_MouseLeave(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));

	CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonUnderMouse, messageMapIndex == 6, this);
	if(pTBButton) {
		Raise_ButtonMouseLeave(pTBButton, button, shift, x, y, hitTestDetails);
	}
	buttonUnderMouse = -1;

	mouseStatus.LeaveControl();

	//if(m_nFreezeEvents == 0) {
		pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
		return Fire_MouseLeave(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_MouseMove(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(!mouseStatus.enteredControl) {
		Raise_MouseEnter(messageMapIndex, button, shift, x, y);
	}
	mouseStatus.lastPosition.x = x;
	mouseStatus.lastPosition.y = y;

	UINT hitTestDetails = 0;
	int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));

	CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
	// Do we move over another button than before?
	if(buttonIndex != buttonUnderMouse) {
		CComPtr<IToolBarButton> pPrevTBButton = ClassFactory::InitToolBarButton(buttonUnderMouse, messageMapIndex == 6, this);
		if(pPrevTBButton) {
			Raise_ButtonMouseLeave(pPrevTBButton, button, shift, x, y, hitTestDetails);
		}
		buttonUnderMouse = buttonIndex;
		if(pTBButton) {
			Raise_ButtonMouseEnter(pTBButton, button, shift, x, y, hitTestDetails);
		}
	}

	//if(m_nFreezeEvents == 0) {
		return Fire_MouseMove(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_MouseUp(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	if(mouseStatus.IsClickCandidate(button)) {
		/* Watch for clicks.
		   Are we still within the control's client area? */
		BOOL hasLeftControl = FALSE;
		DWORD position = GetMessagePos();
		POINT cursorPosition = {GET_X_LPARAM(position), GET_Y_LPARAM(position)};
		RECT clientArea;
		if(messageMapIndex == 6) {
			chevronPopupToolbar.GetClientRect(&clientArea);
			chevronPopupToolbar.ClientToScreen(&clientArea);
		} else {
			GetClientRect(&clientArea);
			ClientToScreen(&clientArea);
		}
		if(PtInRect(&clientArea, cursorPosition)) {
			// maybe the control is overlapped by a foreign window
			if(WindowFromPoint(cursorPosition) != (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this))) {
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
						Raise_Click(messageMapIndex, button, shift, x, y);
					}
					break;
				case 2/*MouseButtonConstants.vbRightButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_RClick(messageMapIndex, button, shift, x, y);
					}
					break;
				case 4/*MouseButtonConstants.vbMiddleButton*/:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_MClick(messageMapIndex, button, shift, x, y);
					}
					break;
				case embXButton1:
				case embXButton2:
					if(!(properties.disabledEvents & deClickEvents)) {
						Raise_XClick(messageMapIndex, button, shift, x, y);
					}
					break;
			}
		}

		mouseStatus.RemoveClickCandidate(button);
	}

	if(!(properties.disabledEvents & deMouseEvents)) {
		//if(m_nFreezeEvents == 0) {
			UINT hitTestDetails = 0;
			int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));
			CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
			return Fire_MouseUp(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
		//}
	} else {
		return S_OK;
	}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLECompleteDrag(pData, performedEffect);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	KillTimer(timers.ID_DRAGCLICK);

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, *this/*, TRUE*/);
	CComPtr<IToolBarButton> pDropTarget = ClassFactory::InitToolBarButton(dropTarget, FALSE, this);

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

inline HRESULT ToolBar::Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	// NOTE: pData can be NULL

	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	dragDropStatus.OLEDragEnter();

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, *this);
	IToolBarButton* pDropTarget = NULL;
	ClassFactory::InitToolBarButton(dropTarget, FALSE, this, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	VARIANT_BOOL autoClickButton = VARIANT_FALSE;
	/*if(dropTarget == -1 || !(hitTestDetails & htButton)) {
		autoClickButton = VARIANT_FALSE;
	}*/

	if(pData) {
		dragDropStatus.pActiveDataObject = ClassFactory::InitOLEDataObject(pData);
	} else {
		dragDropStatus.pActiveDataObject = NULL;
	}
	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragEnter(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &autoClickButton);
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	if(autoClickButton == VARIANT_FALSE) {
		// cancel auto-clicking
		KillTimer(timers.ID_DRAGCLICK);
	} else {
		if(dropTarget != dragDropStatus.lastDropTarget) {
			// cancel auto-clicking of previous target
			KillTimer(timers.ID_DRAGCLICK);
			if(properties.dragClickTime != 0) {
				// start timered auto-activation of new target
				SetTimer(timers.ID_DRAGCLICK, (properties.dragClickTime == -1 ? GetDoubleClickTime() : properties.dragClickTime));
			}
		}
	}
	dragDropStatus.lastDropTarget = dropTarget;

	return hr;
}

inline HRESULT ToolBar::Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragEnterPotentialTarget(hWndPotentialTarget);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_OLEDragLeave(void)
{
	KillTimer(timers.ID_DRAGCLICK);

	SHORT button = 0;
	SHORT shift = 0;
	WPARAM2BUTTONANDSHIFT(-1, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(dragDropStatus.lastMousePosition.x, dragDropStatus.lastMousePosition.y, &hitTestDetails, *this);
	CComPtr<IToolBarButton> pDropTarget = ClassFactory::InitToolBarButton(dropTarget, FALSE, this);

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

inline HRESULT ToolBar::Raise_OLEDragLeavePotentialTarget(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEDragLeavePotentialTarget();
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition)
{
	ScreenToClient(reinterpret_cast<LPPOINT>(&mousePosition));
	dragDropStatus.lastMousePosition = mousePosition;
	SHORT button = 0;
	SHORT shift = 0;
	OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);

	UINT hitTestDetails = 0;
	int dropTarget = HitTest(mousePosition.x, mousePosition.y, &hitTestDetails, *this);
	IToolBarButton* pDropTarget = NULL;
	ClassFactory::InitToolBarButton(dropTarget, FALSE, this, IID_IToolBarButton, reinterpret_cast<LPUNKNOWN*>(&pDropTarget));

	VARIANT_BOOL autoClickButton = VARIANT_FALSE;
	/*if(dropTarget == -1 || !(hitTestDetails & htButton)) {
		autoClickButton = VARIANT_FALSE;
	}*/

	HRESULT hr = S_OK;
	//if(m_nFreezeEvents == 0) {
		if(dragDropStatus.pActiveDataObject) {
			hr = Fire_OLEDragMouseMove(dragDropStatus.pActiveDataObject, reinterpret_cast<OLEDropEffectConstants*>(pEffect), &pDropTarget, button, shift, mousePosition.x, mousePosition.y, static_cast<HitTestConstants>(hitTestDetails), &autoClickButton);
		}
	//}

	if(pDropTarget) {
		// we're using a raw pointer
		pDropTarget->Release();
	}

	if(autoClickButton == VARIANT_FALSE) {
		// cancel auto-clicking
		KillTimer(timers.ID_DRAGCLICK);
	} else {
		if(dropTarget != dragDropStatus.lastDropTarget) {
			// cancel auto-clicking of previous target
			KillTimer(timers.ID_DRAGCLICK);
			if(properties.dragClickTime != 0) {
				// start timered auto-activation of new target
				SetTimer(timers.ID_DRAGCLICK, (properties.dragClickTime == -1 ? GetDoubleClickTime() : properties.dragClickTime));
			}
		}
	}
	dragDropStatus.lastDropTarget = dropTarget;

	return hr;
}

inline HRESULT ToolBar::Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEGiveFeedback(static_cast<OLEDropEffectConstants>(effect), pUseDefaultCursors);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith)
{
	//if(m_nFreezeEvents == 0) {
		SHORT button = 0;
		SHORT shift = 0;
		OLEKEYSTATE2BUTTONANDSHIFT(keyState, button, shift);
		return Fire_OLEQueryContinueDrag(BOOL2VARIANTBOOL(pressedEscape), button, shift, reinterpret_cast<OLEActionToContinueWithConstants*>(pActionToContinueWith));
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT ToolBar::Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEReceivedNewData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

/* We can't make this one inline, because it's called from SourceOLEDataObject only, so the compiler
   would try to integrate it into SourceOLEDataObject, which of course won't work. */
HRESULT ToolBar::Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLESetData(pData, formatID, index, dataOrViewAspect);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_OLEStartDrag(IOLEDataObject* pData)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OLEStartDrag(pData);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_OpenPopupMenu(LONG hMenu, LONG parentMenuItemIndex, VARIANT_BOOL isSystemMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_OpenPopupMenu(hMenu, parentMenuItemIndex, isSystemMenu);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_QueryInsertButton(IToolBarButton* pToolButton, VARIANT_BOOL* pAllowInsertionToLeft)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_QueryInsertButton(pToolButton, pAllowInsertionToLeft);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_QueryRemoveButton(IToolBarButton* pToolButton, VARIANT_BOOL* pAllowRemoval)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_QueryRemoveButton(pToolButton, pAllowRemoval);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_RawMenuMessage(LONG message, LONG wParam, LONG lParam, LONG* pResult, VARIANT_BOOL* pHandledEvent)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RawMenuMessage(message, wParam, lParam, pResult, pHandledEvent);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_RClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedButton = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));

	//if(m_nFreezeEvents == 0) {
		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(mouseStatus.lastClickedButton, messageMapIndex == 6, this);
		return Fire_RClick(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_RDblClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));
	if(buttonIndex != mouseStatus.lastClickedButton) {
		buttonIndex = -1;
	}
	mouseStatus.lastClickedButton = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
		return Fire_RDblClick(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_RecreatedControlWindow(LONG hWnd)
{
	flags.applyBackgroundHack = FALSE;
	if(GetStyle() & TBSTYLE_TRANSPARENT) {
		if(!GetWindowTheme(*this)) {
			TCHAR pClassName[300];
			GetClassName(GetParent(), pClassName, 300);
			if(lstrcmpi(pClassName, TEXT("ReBarU")) != 0 && lstrcmpi(pClassName, TEXT("ReBarA")) != 0 && lstrcmpi(pClassName, REBARCLASSNAME) != 0) {
				flags.applyBackgroundHack = TRUE;
				ModifyStyle(0, TBSTYLE_CUSTOMERASE);
			}
		}
	}

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

inline HRESULT ToolBar::Raise_RemovedButton(IVirtualToolBarButton* pToolButton)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovedButton(pToolButton);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_RemovingButton(IToolBarButton* pToolButton, VARIANT_BOOL* pCancel)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RemovingButton(pToolButton, pCancel);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_ResetCustomizations(LONG hCustomizationDialog, VARIANT_BOOL* pEndCustomization)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResetCustomizations(hCustomizationDialog, pEndCustomization);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_ResizedControlWindow(void)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_ResizedControlWindow();
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_RestoreButtonFromRegistryStream(IVirtualToolBarButton* pToolButton, LONG numberOfButtonsToLoad, LONG totalDataSize, LONG perButtonDataSize, LONG pDataStream, LONG* ppStartOfNextDataBlock)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_RestoreButtonFromRegistryStream(pToolButton, numberOfButtonsToLoad, totalDataSize, perButtonDataSize, pDataStream, ppStartOfNextDataBlock);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_SaveButtonToRegistryStream(IVirtualToolBarButton* pToolButton, LONG totalDataSize, LONG pDataStream, LONG* ppStartOfUnusedSpace)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SaveButtonToRegistryStream(pToolButton, totalDataSize, pDataStream, ppStartOfUnusedSpace);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_SelectedMenuItem(LONG commandIDOrSubMenuIndex, MenuItemStateConstants menuItemState, LONG hMenu)
{
	//if(m_nFreezeEvents == 0) {
		return Fire_SelectedMenuItem(commandIDOrSubMenuIndex, menuItemState, hMenu);
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_XClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	mouseStatus.lastClickedButton = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));

	//if(m_nFreezeEvents == 0) {
		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(mouseStatus.lastClickedButton, messageMapIndex == 6, this);
		return Fire_XClick(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}

inline HRESULT ToolBar::Raise_XDblClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
{
	UINT hitTestDetails = 0;
	int buttonIndex = HitTest(x, y, &hitTestDetails, (messageMapIndex == 6 ? chevronPopupToolbar : static_cast<HWND>(*this)));
	if(buttonIndex != mouseStatus.lastClickedButton) {
		buttonIndex = -1;
	}
	mouseStatus.lastClickedButton = -1;

	//if(m_nFreezeEvents == 0) {
		CComPtr<IToolBarButton> pTBButton = ClassFactory::InitToolBarButton(buttonIndex, messageMapIndex == 6, this);
		return Fire_XDblClick(pTBButton, button, shift, x, y, static_cast<HitTestConstants>(hitTestDetails));
	//}
	//return S_OK;
}


void ToolBar::RecreateControlWindow(void)
{
	if(m_bInPlaceActive) {
		BOOL isUIActive = m_bUIActive;
		InPlaceDeactivate();
		ATLASSERT(m_hWnd == NULL);
		InPlaceActivate((isUIActive ? OLEIVERB_UIACTIVATE : OLEIVERB_INPLACEACTIVATE));
	}
}

DWORD ToolBar::GetExStyleBits(void)
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

DWORD ToolBar::GetStyleBits(void)
{
	DWORD style = WS_CHILDWINDOW | WS_CLIPCHILDREN | WS_CLIPSIBLINGS | WS_VISIBLE;
	/* Comctl32.dll has a bug (at least version 5.x - not sure about newer versions):
	 * It takes the tool bar's window rectangle relative to the parent window's client area, and calls
	 * SetWindowPos with the width parameter set to rect.right instead of rect.right - rect.left.
	 * Therefore specify CCS_NORESIZE.
	 */
	style |= CCS_NORESIZE;
	switch(properties.borderStyle) {
		case bsFixedSingle:
			style |= WS_BORDER;
			break;
	}
	if(!properties.enabled) {
		style |= WS_DISABLED;
	}

	if(properties.allowCustomization) {
		style |= CCS_ADJUSTABLE;
	}
	if(!properties.displayMenuDivider) {
		style |= CCS_NODIVIDER;
	}
	switch(properties.backStyle) {
		case bksTransparent:
			style |= TBSTYLE_TRANSPARENT;
			break;
	}
	switch(properties.buttonStyle) {
		case bstFlat:
			style |= TBSTYLE_FLAT;
			break;
	}
	switch(properties.buttonTextPosition) {
		case btpRightToIcon:
			style |= TBSTYLE_LIST;
			properties.defaultDrawTextFlags &= ~DT_CENTER;
			properties.defaultDrawTextFlags |= DT_LEFT | DT_VCENTER | DT_SINGLELINE;
			break;
		default:
			properties.defaultDrawTextFlags &= ~(DT_LEFT | DT_VCENTER | DT_SINGLELINE);
			properties.defaultDrawTextFlags |= DT_CENTER;
			break;
	}
	switch(properties.dragDropCustomizationModifierKey) {
		case ddcmkAlt:
			style |= TBSTYLE_ALTDRAG;
			break;
	}
	if(properties.raiseCustomDrawEventOnEraseBackground || flags.applyBackgroundHack) {
		style |= TBSTYLE_CUSTOMERASE;
	}
	if(properties.registerForOLEDragDrop == rfoddNativeDragDrop) {
		style |= TBSTYLE_REGISTERDROP;
	}
	if(properties.showToolTips) {
		style |= TBSTYLE_TOOLTIPS;
	}
	if(properties.wrapButtons) {
		style |= TBSTYLE_WRAPABLE;
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
									FireOnChanged(DISPID_TB_ORIENTATION);
									break;
								case 2/*vbAlignBottom*/:
									style |= CCS_NOPARENTALIGN | CCS_NOMOVEY/*CCS_BOTTOM*/;
									properties.orientation = oHorizontal;
									FireOnChanged(DISPID_TB_ORIENTATION);
									break;
								case 3/*vbAlignLeft*/:
									style |= CCS_NOPARENTALIGN | CCS_NOMOVEY | CCS_VERT/*CCS_LEFT*/;
									properties.orientation = oVertical;
									FireOnChanged(DISPID_TB_ORIENTATION);
									properties.displayPartiallyClippedButtons = TRUE;
									FireOnChanged(DISPID_TB_DISPLAYPARTIALLYCLIPPEDBUTTONS);
									break;
								case 4/*vbAlignRight*/:
									style |= CCS_NOPARENTALIGN | CCS_NOMOVEY | CCS_VERT/*CCS_RIGHT*/;
									properties.orientation = oVertical;
									FireOnChanged(DISPID_TB_ORIENTATION);
									properties.displayPartiallyClippedButtons = TRUE;
									FireOnChanged(DISPID_TB_DISPLAYPARTIALLYCLIPPEDBUTTONS);
									break;
								default:
									// NOTE: Don't specify CCS_NOMOVEY here! It will right-align the control if run outside VB6 IDE.
									style |= CCS_NOPARENTALIGN;
									switch(properties.orientation) {
										case oVertical:
											style |= CCS_VERT;
											properties.displayPartiallyClippedButtons = TRUE;
											FireOnChanged(DISPID_TB_DISPLAYPARTIALLYCLIPPEDBUTTONS);
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

	if(IsInDesignMode()) {
		style &= ~(TBSTYLE_FLAT | TBSTYLE_TRANSPARENT);
	}
	return style;
}

void ToolBar::SendConfigurationMessages(void)
{
	if(!IsInDesignMode()) {
		if(properties.menuMode) {
			if(parentWindowMenuMode.IsWindow()) {
				ATLVERIFY(RemoveWindowSubclass(parentWindowMenuMode, ToolBar::ParentWindowSubclass_MenuMode, reinterpret_cast<UINT_PTR>(this)));
				parentWindowMenuMode.Detach();
			}
			parentWindowMenuMode.Attach(GetTopLevelParent());
			ATLVERIFY(SetWindowSubclass(parentWindowMenuMode, ToolBar::ParentWindowSubclass_MenuMode, reinterpret_cast<UINT_PTR>(this), NULL));

			AddMessageHookClient();
			menuModeState.attachedToMessageHook = TRUE;
		}
		ATLVERIFY(KeyboardHookProvider::InstallHook(this));
	}

	DWORD extendedStyle = 0;
	if(RunTimeHelper::IsCommCtrl6()) {
		// TBSTYLE_EX_DOUBLEBUFFER eliminates flicker
		extendedStyle |= TBSTYLE_EX_DOUBLEBUFFER;
	}
	if(!properties.alwaysDisplayButtonText) {
		extendedStyle |= TBSTYLE_EX_MIXEDBUTTONS;
	}
	if(!properties.displayPartiallyClippedButtons && !IsInDesignMode()) {
		extendedStyle |= TBSTYLE_EX_HIDECLIPPEDBUTTONS;
	}
	switch(properties.normalDropDownButtonStyle) {
		case nddbsSplitButton:
			extendedStyle |= TBSTYLE_EX_DRAWDDARROWS;
			break;
	}
	switch(properties.orientation) {
		case oVertical:
			flags.ignoreResize = TRUE;
			extendedStyle |= TBSTYLE_EX_VERTICAL;
			extendedStyle &= ~TBSTYLE_EX_HIDECLIPPEDBUTTONS;
			break;
	}
	if(properties.multiColumn) {
		extendedStyle |= TBSTYLE_EX_MULTICOLUMN;
	}
	SendMessage(TB_SETEXTENDEDSTYLE, 0, extendedStyle);
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
								case 2/*vbAlignBottom*/:
								case 3/*vbAlignLeft*/:
								case 4/*vbAlignRight*/:
									break;
								default:
									switch(properties.orientation) {
										case oVertical:
											ModifyStyle(CCS_TOP | CCS_BOTTOM | CCS_NOMOVEY, CCS_NOPARENTALIGN, SWP_DRAWFRAME | SWP_FRAMECHANGED);
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
	flags.ignoreResize = FALSE;

	SendMessage(TB_SETANCHORHIGHLIGHT, properties.anchorHighlighting, 0);
	properties.defaultButtonPadding = SendMessage(TB_GETPADDING, 0, 0);
	if(RunTimeHelper::IsCommCtrl6()) {
		// should be done before TB_SETINDENT, because TB_SETINDENT will properly update the button rectangles
		TBMETRICS metrics = {0};
		metrics.cbSize = sizeof(metrics);
		metrics.dwMask = /*TBMF_PAD |*/ TBMF_BUTTONSPACING;
		/*metrics.cxPad = properties.horizontalButtonPadding;
		metrics.cyPad = properties.verticalButtonPadding;*/
		metrics.cxButtonSpacing = properties.horizontalButtonSpacing;
		metrics.cyButtonSpacing = properties.verticalButtonSpacing;
		SendMessage(TB_SETMETRICS, 0, reinterpret_cast<LPARAM>(&metrics));
	}
	if(properties.horizontalButtonPadding != -1 || properties.verticalButtonPadding != -1) {
		SendMessage(TB_SETPADDING, 0, MAKELPARAM(properties.horizontalButtonPadding, properties.verticalButtonPadding));
	}
	if(properties.horizontalIconCaptionGap != -1) {
		SendMessage(TB_SETLISTGAP, properties.horizontalIconCaptionGap, 0);
	}
	if(properties.dropDownGap != -1) {
		SendMessage(TB_SETDROPDOWNGAP, properties.dropDownGap, 0);
	}
	SendMessage(TB_SETINDENT, properties.firstButtonIndentation, 0);
	if(properties.maximumTextRows > 1) {
		properties.defaultDrawTextFlags &= ~DT_SINGLELINE;
		properties.defaultDrawTextFlags |= DT_WORDBREAK | DT_EDITCONTROL;
	} else {
		properties.defaultDrawTextFlags &= ~(DT_WORDBREAK | DT_EDITCONTROL);
		properties.defaultDrawTextFlags |= DT_SINGLELINE;
	}
	SendMessage(TB_SETMAXTEXTROWS, properties.maximumTextRows, 0);
	SendMessage(TB_BUTTONSTRUCTSIZE, sizeof(TBBUTTON), 0);

	if(properties.highlightColor != static_cast<OLE_COLOR>(-1) || properties.shadowColor != static_cast<OLE_COLOR>(-1)) {
		COLORSCHEME colorScheme = {sizeof(COLORSCHEME), 0, 0};
		colorScheme.clrBtnHighlight = (properties.highlightColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.highlightColor));
		colorScheme.clrBtnShadow = (properties.shadowColor == static_cast<OLE_COLOR>(-1) ? CLR_DEFAULT : OLECOLOR2COLORREF(properties.shadowColor));
		SendMessage(TB_SETCOLORSCHEME, 0, reinterpret_cast<LPARAM>(&colorScheme));
	}
	SendMessage(TB_SETINSERTMARKCOLOR, 0, OLECOLOR2COLORREF(properties.insertMarkColor));

	properties.drawTextFlags = properties.defaultDrawTextFlags;
	SetDrawTextFlags();

	ApplyFont();

	SendMessage(TB_SETBUTTONWIDTH, 0, MAKELPARAM(properties.minimumButtonWidth, properties.maximumButtonWidth));
	if(properties.buttonHeight != 0 || properties.buttonWidth != 0) {
		SendMessage(TB_SETBUTTONSIZE, 0, MAKELPARAM(properties.buttonWidth, properties.buttonHeight));
	}

	// To verify: Doing this here could make the background drawing hack unnecessary.
	switch(properties.backStyle) {
		case bksTransparent:
			ModifyStyle(0, TBSTYLE_TRANSPARENT);
			break;
		case bksOpaque:
			ModifyStyle(TBSTYLE_TRANSPARENT, 0);
			break;
	}

	if(IsInDesignMode()) {
		// for some reason the TBSTYLE_FLAT style gets reverted
		switch(properties.buttonStyle) {
			case bst3D:
				ModifyStyle(TBSTYLE_FLAT, 0);
				break;
			case bstFlat:
				ModifyStyle(0, TBSTYLE_FLAT);
				break;
		}
		// insert some dummy buttons
		PostMessage(TB_LOADIMAGES, IDB_STD_SMALL_COLOR, reinterpret_cast<LPARAM>(HINST_COMMCTRL));
		TBBUTTON buttons[4];
		ZeroMemory(buttons, 4 * sizeof(TBBUTTON));
		buttons[0].idCommand = 1;
		//buttons[0].iString = reinterpret_cast<INT_PTR>(TEXT("Button 1"));
		buttons[0].fsState = TBSTATE_ENABLED;
		buttons[0].fsStyle = BTNS_AUTOSIZE | BTNS_BUTTON;
		buttons[0].iBitmap = STD_FILENEW;
		buttons[1].idCommand = 2;
		//buttons[1].iString = reinterpret_cast<INT_PTR>(TEXT("Button 2"));
		buttons[1].fsState = TBSTATE_ENABLED;
		buttons[1].fsStyle = BTNS_AUTOSIZE | BTNS_BUTTON;
		buttons[1].iBitmap = STD_FILEOPEN;
		buttons[2].idCommand = 3;
		buttons[2].fsStyle = BTNS_SEP;
		buttons[3].idCommand = 4;
		//buttons[3].iString = reinterpret_cast<INT_PTR>(TEXT("Button 3"));
		buttons[3].fsState = TBSTATE_ENABLED;
		buttons[3].fsStyle = BTNS_AUTOSIZE | BTNS_BUTTON;
		buttons[3].iBitmap = STD_FILESAVE;
		SendMessage(TB_ADDBUTTONS, 4, reinterpret_cast<LPARAM>(buttons));
	}
}

HCURSOR ToolBar::MousePointerConst2hCursor(MousePointerConstants mousePointer)
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

int ToolBar::IDToButtonIndex(LONG ID, HWND hWndToUse/* = NULL*/)
{
	if(hWndToUse) {
		return static_cast<int>(::SendMessage(hWndToUse, TB_COMMANDTOINDEX, ID, 0));
	} else {
		return static_cast<int>(SendMessage(TB_COMMANDTOINDEX, ID, 0));
	}
}

LONG ToolBar::ButtonIndexToID(int buttonIndex)
{
	ATLASSERT(IsWindow());

	TBBUTTON button = {0};
	if(SendMessage(TB_GETBUTTON, buttonIndex, reinterpret_cast<LPARAM>(&button))) {
		return button.idCommand;
	}
	return -1;
}


LRESULT CALLBACK ToolBar::ParentWindowSubclass_MenuMode(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/)
{
	LRESULT lr = 0;
	ToolBar* pThis = reinterpret_cast<ToolBar*>(idSubclass);
	if(pThis) {
		if(!pThis->ProcessWindowMessage(hWnd, message, wParam, lParam, lr, 1)) {
			lr = DefSubclassProc(hWnd, message, wParam, lParam);
		}
	} else {
		lr = DefSubclassProc(hWnd, message, wParam, lParam);
	}
	return lr;
}

LRESULT CALLBACK ToolBar::ParentWindowSubclass_ChevronPopupMenu(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/)
{
	LRESULT lr = 0;
	ToolBar* pThis = reinterpret_cast<ToolBar*>(idSubclass);
	if(pThis) {
		if(!pThis->ProcessWindowMessage(hWnd, message, wParam, lParam, lr, 5)) {
			lr = DefSubclassProc(hWnd, message, wParam, lParam);
		}
	} else {
		lr = DefSubclassProc(hWnd, message, wParam, lParam);
	}
	return lr;
}

void ToolBar::GetSystemSettings(void)
{
	BOOL alwaysDisplayKeyboardCues = TRUE;
	BOOL ret = SystemParametersInfo(SPI_GETKEYBOARDCUES, 0, &alwaysDisplayKeyboardCues, 0);
	menuModeState.useContextSensitiveKeyboardCues = (ret && !alwaysDisplayKeyboardCues);
	menuModeState.allowDisplayingKeyboardCues = TRUE;
	ShowKeyboardCues(!menuModeState.useContextSensitiveKeyboardCues);
}

void ToolBar::TakeFocus(void)
{
	if(!menuModeState.hFocusedWindow) {
		menuModeState.hFocusedWindow = GetFocus();
	}
	flags.allowSetFocus = TRUE;
	SetFocus();
	flags.allowSetFocus = FALSE;
}

void ToolBar::GiveFocusBack(void)
{
	if(menuModeState.parentWindowIsActive) {
		if(::IsWindow(menuModeState.hFocusedWindow) && ::IsWindowVisible(menuModeState.hFocusedWindow)) {
			::SetFocus(menuModeState.hFocusedWindow);
		}
	}
	menuModeState.hFocusedWindow = NULL;
	SendMessage(TB_SETANCHORHIGHLIGHT, FALSE, 0);
	if(menuModeState.useContextSensitiveKeyboardCues && menuModeState.displayKeyboardCues) {
		ShowKeyboardCues(FALSE);
	}
	menuModeState.skipPostingKeyDown = FALSE;
}

void ToolBar::ShowKeyboardCues(BOOL show)
{
	menuModeState.displayKeyboardCues = show;
	SetDrawTextFlags();
	CWindow parent = GetTopLevelParent();
	if(parent.IsWindow()) {
		/*TCHAR pClassName[200] = {0};
		GetClassName(parent, pClassName, 200);*/
		if(show) {
			parent.SendMessage(WM_CHANGEUISTATE, MAKEWPARAM(UIS_CLEAR, UISF_HIDEACCEL | UISF_HIDEFOCUS), 0);
		} else {
			parent.SendMessage(WM_CHANGEUISTATE, MAKEWPARAM(UIS_SET, UISF_HIDEACCEL | UISF_HIDEFOCUS), 0);
		}
	}
	Invalidate();
	UpdateWindow();
}

int ToolBar::GetPreviousMenuItem(int buttonIndex)
{
	if(buttonIndex == -1) {
		return -1;
	}
	RECT clientRectangle;
	GetClientRect(&clientRectangle);
	int nextButton;
	for(nextButton = buttonIndex - 1; nextButton != buttonIndex; nextButton--) {
		if(nextButton < 0) {
			if(menuModeState.mdiChildIsMaximized && buttonIndex != -3) {
				nextButton = -3;     // MDI child system menu
				break;
			} else {
				nextButton = SendMessage(TB_BUTTONCOUNT, 0, 0) - 1;
			}
		}
		TBBUTTON button = {0};
		SendMessage(TB_GETBUTTON, nextButton, reinterpret_cast<LPARAM>(&button));
		RECT buttonRectangle;
		SendMessage(TB_GETITEMRECT, nextButton, reinterpret_cast<LPARAM>(&buttonRectangle));
		if(buttonRectangle.right > clientRectangle.right) {
			nextButton = -2;     // chevron
			break;
		}
		if((button.fsState & TBSTATE_ENABLED) && !(button.fsState & TBSTATE_HIDDEN)) {
			break;
		}
	}
	return (nextButton != buttonIndex) ? nextButton : -1;
}

int ToolBar::GetNextMenuItem(int buttonIndex)
{
	if(buttonIndex == -1) {
		return -1;
	}
	RECT clientRectangle;
	GetClientRect(&clientRectangle);
	int nextButton;
	int buttonCount = SendMessage(TB_BUTTONCOUNT, 0, 0);
	for(nextButton = buttonIndex + 1; nextButton != buttonIndex; nextButton++) {
		if(nextButton >= buttonCount) {
			if(menuModeState.mdiChildIsMaximized && buttonIndex != -3) {
				nextButton = -3;     // MDI child system menu
				break;
			} else {
				nextButton = 0;
			}
		}
		TBBUTTON button = {0};
		SendMessage(TB_GETBUTTON, nextButton, reinterpret_cast<LPARAM>(&button));
		RECT buttonRectangle;
		SendMessage(TB_GETITEMRECT, nextButton, reinterpret_cast<LPARAM>(&buttonRectangle));
		if(buttonRectangle.right > clientRectangle.right) {
			nextButton = -2;     // chevron
			break;
		}
		if((button.fsState & TBSTATE_ENABLED) && !(button.fsState & TBSTATE_HIDDEN)) {
			break;
		}
	}
	return (nextButton != buttonIndex) ? nextButton : -1;
}

BOOL ToolBar::DisplayChevronMenu(void)
{
	// assume we are in a rebar
	HWND hWndReBar = GetParent();
	REBARBANDINFO bandInfo = {RunTimeHelper::SizeOf_REBARBANDINFO(), RBBIM_CHILD | RBBIM_STYLE};
	int bandCount = ::SendMessage(hWndReBar, RB_GETBANDCOUNT, 0, 0);
	BOOL ret = FALSE;
	for(int i = 0; i < bandCount; i++) {
		if(::SendMessage(hWndReBar, RB_GETBANDINFO, i, reinterpret_cast<LPARAM>(&bandInfo)) && bandInfo.hwndChild == *this) {
			if(bandInfo.fStyle & RBBS_USECHEVRON) {
				::PostMessage(hWndReBar, RB_PUSHCHEVRON, i, 0);
				PostMessage(WM_KEYDOWN, VK_DOWN, 0);
				ret = TRUE;
			}
			break;
		}
	}
	return ret;
}

void ToolBar::DoPopupMenu(int buttonIndex, HMENU hPopupMenu, BOOL animate)
{
	if(!IsMenu(hPopupMenu)) {
		return;
	}
	LONG buttonID = ButtonIndexToID(buttonIndex);

	RECT buttonRectangle = {0};
	SendMessage(TB_GETITEMRECT, buttonIndex, reinterpret_cast<LPARAM>(&buttonRectangle));
	POINT menuPosition = {buttonRectangle.left, buttonRectangle.bottom};
	MapWindowPoints(NULL, &menuPosition, 1);
	MapWindowPoints(NULL, &buttonRectangle);
	TPMPARAMS tpmexParams = {0};
	tpmexParams.cbSize = sizeof(TPMPARAMS);
	tpmexParams.rcExclude = buttonRectangle;

	menuModeState.activePopupMenuButtonIndex = buttonIndex;

	SendMessage(TB_PRESSBUTTON, buttonID, MAKELPARAM(TRUE, 0));
	SendMessage(TB_SETHOTITEM, buttonIndex, 0);
	
	DoTrackPopupMenu(hPopupMenu, TPM_LEFTBUTTON | TPM_VERTICAL | TPM_LEFTALIGN | TPM_TOPALIGN | (animate ? TPM_VERPOSANIMATION : TPM_NOANIMATION), menuPosition.x, menuPosition.y, &tpmexParams);
	SendMessage(TB_PRESSBUTTON, buttonID, MAKELPARAM(FALSE, 0));
	if(GetFocus() != *this) {
		SendMessage(TB_SETHOTITEM, static_cast<WPARAM>(-1), 0);
	}

	// TK: WTL does not do this. WTL does not support switching between MDI child window system menu and the normal menu.
	if(menuModeState.activePopupMenuButtonIndex != -3) {
		menuModeState.activePopupMenuButtonIndex = -1;
	}

	// eat next message if click is on the same button
	MSG msg = {0};
	if(PeekMessage(&msg, *this, WM_LBUTTONDOWN, WM_LBUTTONDOWN, PM_NOREMOVE) && PtInRect(&buttonRectangle, msg.pt)) {
		PeekMessage(&msg, *this, WM_LBUTTONDOWN, WM_LBUTTONDOWN, PM_REMOVE);
	}

	// check if another popup menu should be displayed
	if(menuModeState.nextPopupMenuButtonIndex != -1) {
		PostMessage(GetAutoPopupMessage(), menuModeState.nextPopupMenuButtonIndex & 0xFFFF);
		if(!(menuModeState.nextPopupMenuButtonIndex & 0xFFFF0000) && !menuModeState.selectedMenuItemHasSubMenu) {
			PostMessage(WM_KEYDOWN, VK_DOWN, 0);
		}
		menuModeState.nextPopupMenuButtonIndex = -1;
	} else if(menuModeState.activePopupMenuButtonIndex != -3) {     // TK: WTL does not do this. WTL does not support switching between MDI child window system menu and the normal menu.
		menuModeState.displayingPopupMenu = FALSE;
		// if user didn't hit escape, give focus back
		if(!menuModeState.escapePressed) {
			if(menuModeState.useContextSensitiveKeyboardCues && menuModeState.displayKeyboardCues) {
				menuModeState.allowDisplayingKeyboardCues = FALSE;
			}
			GiveFocusBack();
		} else {
			SendMessage(TB_SETHOTITEM, buttonIndex, 0);
			SendMessage(TB_SETANCHORHIGHLIGHT, TRUE, 0);
		}
	}
}

BOOL ToolBar::DoTrackPopupMenu(HMENU hPopupMenu, UINT flags, int x, int y, LPTPMPARAMS pParams/* = NULL*/)
{
	ATLASSERT(hPopupMenu);

	CWindowCreateCriticalSectionLock lock;
	if(FAILED(lock.Lock())) {
		ATLASSERT(FALSE && "Unable to lock critical section in ToolBar::DoTrackPopupMenu.");
		return FALSE;
	}
	pCurrentToolbar = this;

	hCBTHook = SetWindowsHookEx(WH_CBT, CBTProc, ModuleHelper::GetModuleInstance(), GetCurrentThreadId());
	ATLASSERT(hCBTHook);

	menuModeState.selectedMenuItemHasSubMenu = FALSE;
	menuModeState.menuIsActive = TRUE;
	BOOL ret = TrackPopupMenuEx(hPopupMenu, flags, x, y, *this, pParams);
	menuModeState.menuIsActive = FALSE;

	UnhookWindowsHookEx(hCBTHook);
	hCBTHook = NULL;
	pCurrentToolbar = NULL;
	lock.Unlock();

	ATLASSERT(menuModeState.menuWindowStack.GetSize() == 0);

	UpdateWindow();
	CWindow topLevelWindow = GetTopLevelParent();
	if(topLevelWindow.IsWindow()) {
		topLevelWindow.UpdateWindow();
	}

	return ret;
}


int ToolBar::HitTest(LONG x, LONG y, UINT* pFlags, HWND hWndToUse/* = NULL*//*, BOOL ignoreBoundingBoxDefinition = FALSE*/)
{
	if(!hWndToUse) {
		hWndToUse = *this;
	}
	ATLASSERT(::IsWindow(hWndToUse));

	UINT flags = 0;
	if(pFlags) {
		flags = *pFlags;
	}
	POINT hitTestInfo = {x, y};
	UINT hitTestInfoFlags = 0;
	int buttonIndex = static_cast<int>(::SendMessage(hWndToUse, TB_HITTEST, 0, reinterpret_cast<LPARAM>(&hitTestInfo)));
	if(buttonIndex < 0) {
		TBBUTTONINFO button = {0};
		button.cbSize = sizeof(button);
		button.dwMask = TBIF_BYINDEX | TBIF_STYLE;
		::SendMessage(hWndToUse, TB_GETBUTTONINFO, abs(buttonIndex), reinterpret_cast<LPARAM>(&button));
		if(button.fsStyle & BTNS_SEP) {
			hitTestInfoFlags = static_cast<HitTestConstants>(htSplitter);
		} else {
			hitTestInfoFlags = static_cast<HitTestConstants>(htNotInButton);
		}
		buttonIndex = -buttonIndex;
	} else {
		hitTestInfoFlags = static_cast<HitTestConstants>(htButton);
	}

	POINT pt = {x, y};
	::ClientToScreen(hWndToUse, &pt);
	CRect rc;
	::GetWindowRect(hWndToUse, &rc);
	if(!rc.PtInRect(pt)) {
		if(pt.x < rc.left) {
			hitTestInfoFlags = static_cast<HitTestConstants>(hitTestInfoFlags | htToLeft);
		} else if(pt.x >= rc.right) {
			hitTestInfoFlags = static_cast<HitTestConstants>(hitTestInfoFlags | htToRight);
		}
		if(pt.y < rc.top) {
			hitTestInfoFlags = static_cast<HitTestConstants>(hitTestInfoFlags | htAbove);
		} else if(pt.y >= rc.bottom) {
			hitTestInfoFlags = static_cast<HitTestConstants>(hitTestInfoFlags | htBelow);
		}
	}

	flags = hitTestInfoFlags;
	if(pFlags) {
		*pFlags = flags;
	}
	/*TODO: if(!ignoreBoundingBoxDefinition && ((properties.buttonBoundingBoxDefinition & flags) != flags)) {
		buttonIndex = -1;
	}*/
	return buttonIndex;
}

BOOL ToolBar::IsInDesignMode(void)
{
	BOOL b = TRUE;
	GetAmbientUserMode(b);
	return !b;
}

/*void ToolBar::RebuildAcceleratorTable(void)
{
	// NOTE: This method should also be called on TB_ENABLEBUTTON, TB_HIDEBUTTON, TB_SETBUTTONINFO, TB_SETSTATE and TB_SETSTYLE

	if(properties.hAcceleratorTable) {
		DestroyAcceleratorTable(properties.hAcceleratorTable);
		properties.hAcceleratorTable = NULL;
	}

	// create a new accelerator table
	int numberOfButtons = SendMessage(TB_BUTTONCOUNT, 0, 0);
	TCHAR* pAcceleratorChars = static_cast<TCHAR*>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, numberOfButtons * sizeof(TCHAR)));
	if(pAcceleratorChars) {
		int numberOfButtonsWithAccelerator = 0;
		TBBUTTONINFO button = {0};
		button.cbSize = sizeof(button);
		// NOTE: TBIF_BYINDEX requires comctl32.dll 5.80.
		button.dwMask = TBIF_BYINDEX | TBIF_COMMAND | TBIF_STYLE | TBIF_STATE;
		for(int buttonIndex = 0; buttonIndex < numberOfButtons; ++buttonIndex) {
			pAcceleratorChars[buttonIndex] = TEXT('\0');
			if(SendMessage(TB_GETBUTTONINFO, buttonIndex, reinterpret_cast<LPARAM>(&button)) == buttonIndex) {
				if((button.fsStyle & (BTNS_SHOWTEX | BTNS_NOPREFIX | BTNS_SEP)) == BTNS_SHOWTEXT) {
					if(button.fsState & (TBSTATE_HIDDEN | TBSTATE_ENABLED) == TBSTATE_ENABLED) {
						int textLength = static_cast<int>(SendMessage(TB_GETBUTTONTEXT, button.idCommand, NULL));
						if(textLength > -1) {
							LPTSTR pBuffer = static_cast<LPTSTR>(HeapAlloc(GetProcessHeap(), 0, (textLength + 1) * sizeof(TCHAR)));
							if(pBuffer) {
								ZeroMemory(pBuffer, (textLength + 1) * sizeof(TCHAR));
								if(static_cast<int>(SendMessage(TB_GETBUTTONTEXT, button.idCommand, reinterpret_cast<LPARAM>(pBuffer))) > -1) {
									for(int i = lstrlen(pBuffer) - 1; i > 0; --i) {
										if((pBuffer[i - 1] == TEXT('&')) && (pBuffer[i] != TEXT('&'))) {
											++numberOfButtonsWithAccelerator;
											pAcceleratorChars[buttonIndex] = pBuffer[i];
											break;
										}
									}
								}
								HeapFree(GetProcessHeap(), 0, pBuffer);
							}
						}
					}
				}
			}
		}

		if(numberOfButtonsWithAccelerator > 0) {
			LPACCEL pAccelerators = static_cast<LPACCEL>(HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, (numberOfButtonsWithAccelerator * 4) * sizeof(ACCEL)));
			if(pAccelerators) {
				int i = 0;
				for(int buttonIndex = 0; buttonIndex < numberOfButtons; ++buttonIndex) {
					if(pAcceleratorChars[buttonIndex] != TEXT('\0')) {
						pAccelerators[i * 4].cmd = static_cast<WORD>(buttonIndex);
						pAccelerators[i * 4].fVirt = FALT;
						pAccelerators[i * 4].key = static_cast<WORD>(tolower(pAcceleratorChars[buttonIndex]));

						pAccelerators[i * 4 + 1].cmd = static_cast<WORD>(buttonIndex);
						pAccelerators[i * 4 + 1].fVirt = 0;
						pAccelerators[i * 4 + 1].key = static_cast<WORD>(tolower(pAcceleratorChars[buttonIndex]));

						pAccelerators[i * 4 + 2].cmd = static_cast<WORD>(buttonIndex);
						pAccelerators[i * 4 + 2].fVirt = FVIRTKEY | FALT;
						pAccelerators[i * 4 + 2].key = LOBYTE(VkKeyScan(pAcceleratorChars[buttonIndex]));

						pAccelerators[i * 4 + 3].cmd = static_cast<WORD>(buttonIndex);
						pAccelerators[i * 4 + 3].fVirt = FVIRTKEY | FALT | FSHIFT;
						pAccelerators[i * 4 + 3].key = LOBYTE(VkKeyScan(pAcceleratorChars[buttonIndex]));

						++i;
					}
				}
				properties.hAcceleratorTable = CreateAcceleratorTable(pAccelerators, numberOfButtonsWithAccelerator * 4);
				HeapFree(GetProcessHeap(), 0, pAccelerators);
			}
		}
		HeapFree(GetProcessHeap(), 0, pAcceleratorChars);
	}

	// report the new accelerator table to the container
	CComQIPtr<IOleControlSite, &IID_IOleControlSite> pSite(m_spClientSite);
	if(pSite) {
		ATLVERIFY(SUCCEEDED(pSite->OnControlInfoChanged()));
	}
}*/

void ToolBar::RegisterButtonContainer(IButtonContainer* pContainer)
{
	ATLASSUME(pContainer);
	#ifdef _DEBUG
		#ifdef USE_STL
			std::unordered_map<DWORD, IButtonContainer*>::iterator iter = buttonContainers.find(pContainer->GetID());
			ATLASSERT(iter == buttonContainers.end());
		#else
			CAtlMap<DWORD, IButtonContainer*>::CPair* pEntry = buttonContainers.Lookup(pContainer->GetID());
			ATLASSERT(!pEntry);
		#endif
	#endif
	buttonContainers[pContainer->GetID()] = pContainer;
}

void ToolBar::DeregisterButtonContainer(DWORD containerID)
{
	#ifdef USE_STL
		std::unordered_map<DWORD, IButtonContainer*>::iterator iter = buttonContainers.find(containerID);
		ATLASSERT(iter != buttonContainers.end());
		if(iter != buttonContainers.end()) {
			buttonContainers.erase(iter);
		}
	#else
		buttonContainers.RemoveKey(containerID);
	#endif
}

void ToolBar::RemoveButtonFromButtonContainers(LONG buttonID)
{
	#ifdef USE_STL
		for(std::unordered_map<DWORD, IButtonContainer*>::const_iterator iter = buttonContainers.begin(); iter != buttonContainers.end(); ++iter) {
			iter->second->RemovedButton(buttonID);
		}
		if(buttonID == -1) {
			properties.disabledCommands.clear();
		} else {
			for(std::vector<LONG>::iterator it = properties.disabledCommands.begin(); it != properties.disabledCommands.end(); it++) {
				if(*it == buttonID) {
					// remove this entry
					properties.disabledCommands.erase(it);
					break;
				}
			}
		}
	#else
		POSITION p = buttonContainers.GetStartPosition();
		while(p) {
			buttonContainers.GetValueAt(p)->RemovedButton(buttonID);
			buttonContainers.GetNextValue(p);
		}
		if(buttonID == -1) {
			properties.disabledCommands.RemoveAll();
		} else {
			for(size_t i = 0; i < properties.disabledCommands.GetCount(); ++i) {
				if(properties.disabledCommands[i] == buttonID) {
					// remove this entry
					properties.disabledCommands.RemoveAt(i);
					break;
				}
			}
		}
	#endif
}

void ToolBar::RegisterPlaceholderButton(LONG buttonID, HWND hWnd/* = NULL*/)
{
	LPPLACEHOLDERBUTTON pDetails = NULL;
	if(::IsWindow(hWnd)) {
		pDetails = new PLACEHOLDERBUTTON();
		pDetails->hContainedWindow = hWnd;
		pDetails->hOriginalParentWindow = ::GetParent(hWnd);

		::SetParent(hWnd, *this);
		int buttonIndex = IDToButtonIndex(buttonID);
		if(buttonIndex != -1) {
			CRect buttonRectangle;
			if(SendMessage(TB_GETITEMRECT, buttonIndex, reinterpret_cast<LPARAM>(&buttonRectangle))) {
				CRect windowRectangle;
				::GetWindowRect(hWnd, &windowRectangle);
				CRect targetRectangle;
				switch(pDetails->horizontalChildWindowAlignment) {
					case halLeft:
						targetRectangle.left = buttonRectangle.left;
						break;
					case halCenter:
						targetRectangle.left = buttonRectangle.left + ((buttonRectangle.Width() - windowRectangle.Width()) >> 1);
						break;
					case halRight:
						targetRectangle.left = buttonRectangle.right - windowRectangle.Width();
						break;
				}
				targetRectangle.right = targetRectangle.left + windowRectangle.Width();
				switch(pDetails->verticalChildWindowAlignment) {
					case valTop:
						targetRectangle.top = buttonRectangle.top;
						break;
					case valCenter:
						targetRectangle.top = buttonRectangle.top + ((buttonRectangle.Height() - windowRectangle.Height()) >> 1);
						break;
					case valBottom:
						targetRectangle.top = buttonRectangle.bottom - windowRectangle.Height();
						break;
				}
				targetRectangle.bottom = targetRectangle.top + windowRectangle.Height();
				::MoveWindow(hWnd, targetRectangle.left, targetRectangle.top, targetRectangle.Width(), targetRectangle.Height(), TRUE);
			}
		}
	}
	#ifdef _DEBUG
		#ifdef USE_STL
			std::unordered_map<LONG, LPPLACEHOLDERBUTTON>::iterator iter = placeholderButtons.find(buttonID);
			ATLASSERT(iter == placeholderButtons.end());
		#else
			CAtlMap<LONG, LPPLACEHOLDERBUTTON>::CPair* pEntry = placeholderButtons.Lookup(buttonID);
			ATLASSERT(!pEntry);
		#endif
	#endif
	placeholderButtons[buttonID] = pDetails;
}

void ToolBar::DeregisterPlaceholderButton(LONG buttonID)
{
	#ifdef USE_STL
		if(buttonID == -1) {
			for(std::unordered_map<LONG, LPPLACEHOLDERBUTTON>::iterator iter = placeholderButtons.begin(); iter != placeholderButtons.end(); ++iter) {
				if(iter->second) {
					delete iter->second;
				}
			}
			placeholderButtons.clear();
		} else {
			std::unordered_map<LONG, LPPLACEHOLDERBUTTON>::iterator iter = placeholderButtons.find(buttonID);
			ATLASSERT(iter != placeholderButtons.end());
			if(iter != placeholderButtons.end()) {
				if(iter->second) {
					delete iter->second;
				}
				placeholderButtons.erase(iter);
			}
		}
	#else
		if(buttonID == -1) {
			POSITION p = placeholderButtons.GetStartPosition();
			while(p) {
				LPPLACEHOLDERBUTTON pDetails = placeholderButtons.GetValueAt(p);
				if(pDetails) {
					delete pDetails;
				}
				placeholderButtons.GetNextValue(p);
			}
			placeholderButtons.RemoveAll();
		} else {
			CAtlMap<LONG, LPPLACEHOLDERBUTTON>::CPair* pEntry = placeholderButtons.Lookup(buttonID);
			ATLASSERT(!pEntry);
			if(pEntry && pEntry->m_value) {
				delete pEntry->m_value;
			}
			placeholderButtons.RemoveKey(buttonID);
		}
	#endif
}

BOOL ToolBar::IsPlaceholderButton(LONG buttonID)
{
	#ifdef USE_STL
		return (placeholderButtons.find(buttonID) != placeholderButtons.end());
	#else
		return (placeholderButtons.Lookup(buttonID) != NULL);
	#endif
}

HWND ToolBar::GetPlaceholderButtonChildWindow(LONG buttonID, HAlignmentConstants& horizontalAlignment, VAlignmentConstants& verticalAlignment)
{
	#ifdef USE_STL
		std::unordered_map<LONG, LPPLACEHOLDERBUTTON>::iterator iter = placeholderButtons.find(buttonID);
		if(iter != placeholderButtons.end()) {
			if(iter->second) {
				horizontalAlignment = iter->second->horizontalChildWindowAlignment;
				verticalAlignment = iter->second->verticalChildWindowAlignment;
				return iter->second->hContainedWindow;
			}
		}
	#else
		CAtlMap<LONG, LPPLACEHOLDERBUTTON>::CPair* pEntry = placeholderButtons.Lookup(buttonID);
		if(pEntry) {
			if(pEntry->m_value) {
				horizontalAlignment = pEntry->m_value->horizontalChildWindowAlignment;
				verticalAlignment = pEntry->m_value->verticalChildWindowAlignment;
				return pEntry->m_value->hContainedWindow;
			}
		}
	#endif
	return NULL;
}

BOOL ToolBar::SetPlaceholderButtonChildWindow(LONG buttonID, HWND hWnd, HAlignmentConstants horizontalAlignment, VAlignmentConstants verticalAlignment)
{
	BOOL ret = FALSE;
	LPPLACEHOLDERBUTTON pOldDetails = NULL;
	LPPLACEHOLDERBUTTON pNewDetails = NULL;
	if(hWnd) {
		pNewDetails = new PLACEHOLDERBUTTON();
		pNewDetails->hContainedWindow = hWnd;
		pNewDetails->hOriginalParentWindow = ::GetParent(hWnd);
		pNewDetails->horizontalChildWindowAlignment = horizontalAlignment;
		pNewDetails->verticalChildWindowAlignment = verticalAlignment;
	}
	#ifdef USE_STL
		std::unordered_map<LONG, LPPLACEHOLDERBUTTON>::iterator iter = placeholderButtons.find(buttonID);
		if(iter != placeholderButtons.end()) {
			pOldDetails = iter->second;
			placeholderButtons[buttonID] = pNewDetails;
			ret = TRUE;
		}
	#else
		CAtlMap<LONG, LPPLACEHOLDERBUTTON>::CPair* pEntry = placeholderButtons.Lookup(buttonID);
		if(pEntry) {
			pOldDetails = pEntry->m_value;
			placeholderButtons[buttonID] = pNewDetails;
			ret = TRUE;
		}
	#endif

	if(ret) {
		if(pOldDetails) {
			if(pOldDetails->hContainedWindow != hWnd) {
				// restore original parent window
				if(::IsWindow(pOldDetails->hOriginalParentWindow)) {
					::SetParent(pOldDetails->hContainedWindow, pOldDetails->hOriginalParentWindow);
				}
			}
			delete pOldDetails;
		}

		if(::IsWindow(hWnd)) {
			::SetParent(hWnd, *this);
			int buttonIndex = IDToButtonIndex(buttonID);
			if(buttonIndex != -1) {
				CRect buttonRectangle;
				if(SendMessage(TB_GETITEMRECT, buttonIndex, reinterpret_cast<LPARAM>(&buttonRectangle))) {
					CRect windowRectangle;
					::GetWindowRect(hWnd, &windowRectangle);
					CRect targetRectangle;
					switch(pNewDetails->horizontalChildWindowAlignment) {
						case halLeft:
							targetRectangle.left = buttonRectangle.left;
							break;
						case halCenter:
							targetRectangle.left = buttonRectangle.left + ((buttonRectangle.Width() - windowRectangle.Width()) >> 1);
							break;
						case halRight:
							targetRectangle.left = buttonRectangle.right - windowRectangle.Width();
							break;
					}
					targetRectangle.right = targetRectangle.left + windowRectangle.Width();
					switch(pNewDetails->verticalChildWindowAlignment) {
						case valTop:
							targetRectangle.top = buttonRectangle.top;
							break;
						case valCenter:
							targetRectangle.top = buttonRectangle.top + ((buttonRectangle.Height() - windowRectangle.Height()) >> 1);
							break;
						case valBottom:
							targetRectangle.top = buttonRectangle.bottom - windowRectangle.Height();
							break;
					}
					targetRectangle.bottom = targetRectangle.top + windowRectangle.Height();
					::MoveWindow(hWnd, targetRectangle.left, targetRectangle.top, targetRectangle.Width(), targetRectangle.Height(), TRUE);
				}
			}
		}
	} else if(pNewDetails) {
		delete pNewDetails;
	}
	return ret;
}

BOOL ToolBar::IsDroppedDownButton(int buttonIndex)
{
	return (droppedDownButton == buttonIndex);
}

BOOL ToolBar::IsLeftMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	}
}

BOOL ToolBar::IsRightMouseButtonDown(void)
{
	if(GetSystemMetrics(SM_SWAPBUTTON)) {
		return (GetAsyncKeyState(VK_LBUTTON) & 0x8000);
	} else {
		return (GetAsyncKeyState(VK_RBUTTON) & 0x8000);
	}
}


HRESULT ToolBar::CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing)
{
	if(!hIcon || !pBits || !pWICImagingFactory) {
		return E_FAIL;
	}

	ICONINFO iconInfo;
	GetIconInfo(hIcon, &iconInfo);
	ATLASSERT(iconInfo.hbmColor);
	BITMAP bitmapInfo = {0};
	if(iconInfo.hbmColor) {
		GetObject(iconInfo.hbmColor, sizeof(BITMAP), &bitmapInfo);
	} else if(iconInfo.hbmMask) {
		GetObject(iconInfo.hbmMask, sizeof(BITMAP), &bitmapInfo);
	}
	bitmapInfo.bmHeight = abs(bitmapInfo.bmHeight);
	BOOL needsFrame = ((bitmapInfo.bmWidth < size.cx) || (bitmapInfo.bmHeight < size.cy));
	if(iconInfo.hbmColor) {
		DeleteObject(iconInfo.hbmColor);
	}
	if(iconInfo.hbmMask) {
		DeleteObject(iconInfo.hbmMask);
	}

	HRESULT hr = E_FAIL;

	CComPtr<IWICBitmapScaler> pWICBitmapScaler = NULL;
	if(!needsFrame) {
		hr = pWICImagingFactory->CreateBitmapScaler(&pWICBitmapScaler);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapScaler);
	}
	if(needsFrame || SUCCEEDED(hr)) {
		CComPtr<IWICBitmap> pWICBitmapSource = NULL;
		hr = pWICImagingFactory->CreateBitmapFromHICON(hIcon, &pWICBitmapSource);
		ATLASSERT(SUCCEEDED(hr));
		ATLASSUME(pWICBitmapSource);
		if(SUCCEEDED(hr)) {
			if(!needsFrame) {
				hr = pWICBitmapScaler->Initialize(pWICBitmapSource, size.cx, size.cy, WICBitmapInterpolationModeFant);
			}
			if(SUCCEEDED(hr)) {
				WICRect rc = {0};
				if(needsFrame) {
					rc.Height = bitmapInfo.bmHeight;
					rc.Width = bitmapInfo.bmWidth;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					LPRGBQUAD pIconBits = static_cast<LPRGBQUAD>(HeapAlloc(GetProcessHeap(), 0, rc.Width * rc.Height * sizeof(RGBQUAD)));
					hr = pWICBitmapSource->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pIconBits));
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						// center the icon
						int xIconStart = (size.cx - bitmapInfo.bmWidth) / 2;
						int yIconStart = (size.cy - bitmapInfo.bmHeight) / 2;
						LPRGBQUAD pIconPixel = pIconBits;
						LPRGBQUAD pPixel = pBits;
						pPixel += yIconStart * size.cx;
						for(int y = yIconStart; y < yIconStart + bitmapInfo.bmHeight; ++y, pPixel += size.cx, pIconPixel += bitmapInfo.bmWidth) {
							CopyMemory(pPixel + xIconStart, pIconPixel, bitmapInfo.bmWidth * sizeof(RGBQUAD));
						}
						HeapFree(GetProcessHeap(), 0, pIconBits);

						rc.Height = size.cy;
						rc.Width = size.cx;

						// TODO: now draw a frame around it
					}
				} else {
					rc.Height = size.cy;
					rc.Width = size.cx;
					UINT stride = rc.Width * sizeof(RGBQUAD);
					hr = pWICBitmapScaler->CopyPixels(&rc, stride, rc.Height * stride, reinterpret_cast<LPBYTE>(pBits));
					ATLASSERT(SUCCEEDED(hr));

					if(SUCCEEDED(hr) && doAlphaChannelPostProcessing) {
						for(int i = 0; i < rc.Width * rc.Height; ++i, ++pBits) {
							if(pBits->rgbReserved == 0x00) {
								ZeroMemory(pBits, sizeof(RGBQUAD));
							}
						}
					}
				}
			} else {
				ATLASSERT(FALSE && "Bitmap scaler failed");
			}
		}
	}
	return hr;
}


BOOL ToolBar::IsComctl32Version610OrNewer(void)
{
	DWORD major = 0;
	DWORD minor = 0;
	HRESULT hr = ATL::AtlGetCommCtrlVersion(&major, &minor);
	if(SUCCEEDED(hr)) {
		return (((major == 6) && (minor >= 10)) || (major > 6));
	}
	return FALSE;
}