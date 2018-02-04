//////////////////////////////////////////////////////////////////////
/// \class ToolBar
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c ToolbarWindow32</em>
///
/// This class superclasses \c ToolbarWindow32 and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa TBarCtlsLibU::IToolBar, ToolBarButton, ToolBarButtons, ToolBarButtonContainer,
///       VirtualToolBarButton
/// \else
///   \sa TBarCtlsLibA::IToolBar, ToolBarButton, ToolBarButtons, ToolBarButtonContainer,
///       VirtualToolBarButton
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "CWindowEx2.h"
#include "_IToolBarEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "IKeyboardHookHandler.h"
#include "IMouseHookHandler.h"
#include "KeyboardHookProvider.h"
#include "MouseHookProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "ToolBarButton.h"
#include "VirtualToolBarButton.h"
#include "ToolBarButtonContainer.h"
#include "ToolBarButtons.h"
#include "CommonProperties.h"
#include "TargetOLEDataObject.h"
#include "SourceOLEDataObject.h"


class ATL_NO_VTABLE ToolBar : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef _UNICODE
    	public IDispatchImpl<IToolBar, &IID_IToolBar, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IToolBar, &IID_IToolBar, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<ToolBar>,
    public IOleControlImpl<ToolBar>,
    public IOleObjectImpl<ToolBar>,
    public IOleInPlaceActiveObjectImpl<ToolBar>,
    public IViewObjectExImpl<ToolBar>,
    public IOleInPlaceObjectWindowlessImpl<ToolBar>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ToolBar>,
    public Proxy_IToolBarEvents<ToolBar>,
    public IPersistStorageImpl<ToolBar>,
    public IPersistPropertyBagImpl<ToolBar>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<ToolBar>,
    #ifdef _UNICODE
    	public IProvideClassInfo2Impl<&CLSID_ToolBar, &__uuidof(_IToolBarEvents), &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_ToolBar, &__uuidof(_IToolBarEvents), &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<ToolBar>,
    public CComCoClass<ToolBar, &CLSID_ToolBar>,
    public CComControl<ToolBar>,
    public IPerPropertyBrowsingImpl<ToolBar>,
    public IDropTarget,
    public IDropSource,
    public IDropSourceNotify,
    public ICategorizeProperties,
    public ICreditsProvider,
    public IKeyboardHookHandler,
    public IMouseHookHandler
{
	friend class ToolBarButton;
	friend class ToolBarButtonContainer;
	friend class ToolBarButtons;
	friend class VirtualToolBarButton;
	friend class SourceOLEDataObject;

public:
	/// \brief <em>The top-level parent window that we subclass for the \c MenuMode property</em>
	///
	/// \sa get_MenuMode
	CContainedWindow parentWindowMenuMode;
	/// \brief <em>The top-level parent window that we subclass for the \c DisplayChevronPopupWindow method</em>
	///
	/// \sa DisplayChevronPopupWindow, chevronPopupMenuWindow, chevronPopupToolbar
	CContainedWindow parentWindowChevronPopupMenu;
	/// \brief <em>The menu window that we create for the \c DisplayChevronPopupWindow method</em>
	///
	/// \sa DisplayChevronPopupWindow, chevronPopupToolbar, parentWindowChevronPopupMenu
	CWindow chevronPopupMenuWindow;
	/// \brief <em>The tool bar window that we create for the \c DisplayChevronPopupWindow method</em>
	///
	/// \sa DisplayChevronPopupWindow, chevronPopupMenuWindow, parentWindowChevronPopupMenu
	CContainedWindowT<CToolBarCtrl, CControlWinTraits> chevronPopupToolbar;
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	ToolBar();

	/// \brief <em>The handle of the CBT hook</em>
	static HHOOK hCBTHook;
	/// \brief <em>The instance of the \c ToolBar class that is currently displaying a popup menu</em>
	static ToolBar* pCurrentToolbar;

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_CANTLINKINSIDE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_TOOLBAR)

		#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("ToolBarU"), TOOLBARCLASSNAMEW)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("ToolBarA"), TOOLBARCLASSNAMEA)
		#endif

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(ToolBar)
			COM_INTERFACE_ENTRY(IToolBar)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(IViewObjectEx)
			COM_INTERFACE_ENTRY(IViewObject2)
			COM_INTERFACE_ENTRY(IViewObject)
			COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceObject)
			COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
			COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
			COM_INTERFACE_ENTRY(IOleControl)
			COM_INTERFACE_ENTRY(IOleObject)
			COM_INTERFACE_ENTRY(IPersistStreamInit)
			COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IPersistPropertyBag)
			COM_INTERFACE_ENTRY(IQuickActivate)
			COM_INTERFACE_ENTRY(IPersistStorage)
			COM_INTERFACE_ENTRY(IProvideClassInfo)
			COM_INTERFACE_ENTRY(IProvideClassInfo2)
			COM_INTERFACE_ENTRY_IID(IID_ICategorizeProperties, ICategorizeProperties)
			COM_INTERFACE_ENTRY(ISpecifyPropertyPages)
			COM_INTERFACE_ENTRY(IPerPropertyBrowsing)
			COM_INTERFACE_ENTRY(IDropTarget)
			COM_INTERFACE_ENTRY(IDropSource)
			COM_INTERFACE_ENTRY(IDropSourceNotify)
		END_COM_MAP()

		BEGIN_PROP_MAP(ToolBar)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AllowCustomization", DISPID_TB_ALLOWCUSTOMIZATION, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AlwaysDisplayButtonText", DISPID_TB_ALWAYSDISPLAYBUTTONTEXT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("AnchorHighlighting", DISPID_TB_ANCHORHIGHLIGHTING, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_TB_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BackStyle", DISPID_TB_BACKSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_TB_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ButtonHeight", DISPID_TB_BUTTONHEIGHT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ButtonStyle", DISPID_TB_BUTTONSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ButtonTextPosition", DISPID_TB_BUTTONTEXTPOSITION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ButtonWidth", DISPID_TB_BUTTONWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_TB_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisplayMenuDivider", DISPID_TB_DISPLAYMENUDIVIDER, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DisplayPartiallyClippedButtons", DISPID_TB_DISPLAYPARTIALLYCLIPPEDBUTTONS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_TB_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DragClickTime", DISPID_TB_DRAGCLICKTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DragDropCustomizationModifierKey", DISPID_TB_DRAGDROPCUSTOMIZATIONMODIFIERKEY, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DropDownGap", DISPID_TB_DROPDOWNGAP, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Enabled", DISPID_TB_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("FirstButtonIndentation", DISPID_TB_FIRSTBUTTONINDENTATION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("FocusOnClick", DISPID_TB_FOCUSONCLICK, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_TB_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("HighlightColor", DISPID_TB_HIGHLIGHTCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("HorizontalButtonPadding", DISPID_TB_HORIZONTALBUTTONPADDING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HorizontalButtonSpacing", DISPID_TB_HORIZONTALBUTTONSPACING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HorizontalIconCaptionGap", DISPID_TB_HORIZONTALICONCAPTIONGAP, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HorizontalTextAlignment", DISPID_TB_HORIZONTALTEXTALIGNMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("HoverTime", DISPID_TB_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("InsertMarkColor", DISPID_TB_INSERTMARKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("MaximumButtonWidth", DISPID_TB_MAXIMUMBUTTONWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MaximumTextRows", DISPID_TB_MAXIMUMTEXTROWS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MenuBarTheme", DISPID_TB_MENUBARTHEME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MenuMode", DISPID_TB_MENUMODE, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("MinimumButtonWidth", DISPID_TB_MINIMUMBUTTONWIDTH, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_TB_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_TB_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MultiColumn", DISPID_TB_MULTICOLUMN, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("NormalDropDownButtonStyle", DISPID_TB_NORMALDROPDOWNBUTTONSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("OLEDragImageStyle", DISPID_TB_OLEDRAGIMAGESTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Orientation", DISPID_TB_ORIENTATION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ProcessContextMenuKeys", DISPID_TB_PROCESSCONTEXTMENUKEYS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RaiseCustomDrawEventOnEraseBackground", DISPID_TB_RAISECUSTOMDRAWEVENTONERASEBACKGROUND, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_TB_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_TB_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ShadowColor", DISPID_TB_SHADOWCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("ShowToolTips", DISPID_TB_SHOWTOOLTIPS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_TB_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseMnemonics", DISPID_TB_USEMNEMONICS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_TB_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("VerticalButtonPadding", DISPID_TB_VERTICALBUTTONPADDING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("VerticalButtonSpacing", DISPID_TB_VERTICALBUTTONSPACING, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("VerticalTextAlignment", DISPID_TB_VERTICALTEXTALIGNMENT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("WrapButtons", DISPID_TB_WRAPBUTTONS, CLSID_NULL, VT_BOOL)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(ToolBar)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IToolBarEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(ToolBar)
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			MESSAGE_HANDLER(WM_COMMAND, OnCommand)
			INDEXED_MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_CREATE, OnCreate)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_DRAWITEM, OnMenuMessage)
			MESSAGE_HANDLER(WM_FORWARDMSG, OnForwardMsg)
			MESSAGE_HANDLER(WM_INITMENUPOPUP, OnInitMenuPopup)
			MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MEASUREITEM, OnMenuMessage)
			MESSAGE_HANDLER(WM_MENUCHAR, OnMenuChar)
			MESSAGE_HANDLER(WM_MENUSELECT, OnMenuSelect)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			INDEXED_MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			MESSAGE_HANDLER(WM_NOTIFY, OnNotify)
			MESSAGE_HANDLER(WM_PAINT, OnPaint)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

			MESSAGE_HANDLER(OCM_NOTIFY, OnReflectedNotify)
			MESSAGE_HANDLER(OCM__BASE + WM_NOTIFYFORMAT, OnReflectedNotifyFormat)

			MESSAGE_HANDLER(GetDragImageMessage(), OnGetDragImage)
			MESSAGE_HANDLER(GetBarMessage(), OnGetToolbar)
			INDEXED_MESSAGE_HANDLER(GetAutoPopupMessage(), OnAutoPopup)
			MESSAGE_HANDLER(GetMDIChildWindowStateChangedMessage(), OnMDIChildWindowStateChanged)
			MESSAGE_HANDLER(GetDisplayChevronPopupMessage(), OnDisplayChevronPopup)
			MESSAGE_HANDLER(GetDeferredSetButtonTextMessage(), OnDeferredSetButtonText)

			MESSAGE_HANDLER(TB_ADDBUTTONSA, OnAddButtons)
			MESSAGE_HANDLER(TB_ADDBUTTONSW, OnAddButtons)
			MESSAGE_HANDLER(TB_DELETEBUTTON, OnDeleteButton)
			MESSAGE_HANDLER(TB_INSERTBUTTONA, OnInsertButton)
			MESSAGE_HANDLER(TB_INSERTBUTTONW, OnInsertButton)
			MESSAGE_HANDLER(TB_SETBUTTONINFOA, OnSetButtonInfo)
			MESSAGE_HANDLER(TB_SETBUTTONINFOW, OnSetButtonInfo)
			MESSAGE_HANDLER(TB_SETBUTTONWIDTH, OnSetButtonWidth)
			MESSAGE_HANDLER(TB_SETCMDID, OnSetCmdId)
			MESSAGE_HANDLER(TB_SETDRAWTEXTFLAGS, OnSetDrawTextFlags)
			MESSAGE_HANDLER(TB_SETLISTGAP, OnSetDropDownGap)
			MESSAGE_HANDLER(TB_SETINDENT, OnSetIndent)
			MESSAGE_HANDLER(TB_SETLISTGAP, OnSetListGap)

			REFLECTED_NOTIFY_CODE_HANDLER(TBN_BEGINADJUST, OnBeginAdjustNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_BEGINDRAG, OnBeginDragNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_CUSTHELP, OnCustHelpNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_DELETINGBUTTON, OnDeletingButtonNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_DROPDOWN, OnDropDownNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_DUPACCELERATOR, OnDupAcceleratorNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_ENDADJUST, OnEndAdjustNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETBUTTONINFOA, OnGetButtonInfoNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETBUTTONINFOW, OnGetButtonInfoNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETDISPINFOA, OnGetDispInfoNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETDISPINFOW, OnGetDispInfoNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETINFOTIPA, OnGetInfoTipNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETINFOTIPW, OnGetInfoTipNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETOBJECT, OnGetObjectNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_HOTITEMCHANGE, OnHotItemChangeNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_INITCUSTOMIZE, OnInitCustomizeNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_MAPACCELERATOR, OnMapAcceleratorNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_QUERYDELETE, OnQueryDeleteNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_QUERYINSERT, OnQueryInsertNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_RESET, OnResetNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_RESTORE, OnRestoreNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_SAVE, OnSaveNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_TOOLBARCHANGE, OnToolBarChangeNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_WRAPACCELERATOR, OnWrapAcceleratorNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_WRAPHOTITEM, OnWrapHotItemNotification)

			INDEXED_REFLECTED_COMMAND_CODE_HANDLER(BN_CLICKED, OnReflectedClicked)

			CHAIN_MSG_MAP(CComControl<ToolBar>)

			ALT_MSG_MAP(1)     // parentWindowMenuMode
			MESSAGE_HANDLER(WM_ACTIVATE, OnParentActivate)
			MESSAGE_HANDLER(WM_MENUCHAR, OnMenuChar)
			MESSAGE_HANDLER(WM_SYSCOMMAND, OnParentSysCommand)
			MESSAGE_HANDLER(GetBarMessage(), OnGetToolbar)

			ALT_MSG_MAP(4)     // message hook
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnHookMouseMove)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnHookSysKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnHookSysKeyUp)
			MESSAGE_HANDLER(WM_SYSCHAR, OnHookSysChar)
			MESSAGE_HANDLER(WM_KEYDOWN, OnHookKeyDown)
			MESSAGE_HANDLER(WM_NEXTMENU, OnHookNextMenu)
			MESSAGE_HANDLER(WM_CHAR, OnHookChar)

			ALT_MSG_MAP(5)     // parentWindowChevronPopupMenu
			MESSAGE_HANDLER(WM_ACTIVATEAPP, OnChevronPopupMenuDismiss)
			MESSAGE_HANDLER(WM_CANCELMODE, OnChevronPopupMenuDismiss)

			ALT_MSG_MAP(6)     // chevronPopupToolbar
			MESSAGE_HANDLER(WM_CHAR, OnChar)
			//MESSAGE_HANDLER(WM_COMMAND, OnCommand)				// those messages will end up in message map 1, because we specify *this as owner window
			INDEXED_MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			//MESSAGE_HANDLER(WM_CREATE, OnCreate)
			//MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_DRAWITEM, OnMenuMessage)
			//MESSAGE_HANDLER(WM_FORWARDMSG, OnForwardMsg)
			MESSAGE_HANDLER(WM_INITMENUPOPUP, OnInitMenuPopup)
			MESSAGE_HANDLER(WM_KEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_KEYUP, OnKeyUp)
			//MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MEASUREITEM, OnMenuMessage)
			MESSAGE_HANDLER(WM_MENUCHAR, OnMenuChar)
			MESSAGE_HANDLER(WM_MENUSELECT, OnMenuSelect)
			//MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			INDEXED_MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			INDEXED_MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			MESSAGE_HANDLER(WM_NOTIFY, OnChevronPopupToolbarNotify)
			//MESSAGE_HANDLER(WM_PAINT, OnPaint)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			INDEXED_MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			//MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			//MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			//MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			//MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			MESSAGE_HANDLER(WM_SYSKEYDOWN, OnKeyDown)
			MESSAGE_HANDLER(WM_SYSKEYUP, OnKeyUp)
			//MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			//MESSAGE_HANDLER(WM_TIMER, OnTimer)
			//MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			INDEXED_MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)

			MESSAGE_HANDLER(OCM_NOTIFY, OnReflectedNotify)
			//MESSAGE_HANDLER(OCM__BASE + WM_NOTIFYFORMAT, OnReflectedNotifyFormat)

			//MESSAGE_HANDLER(GetDragImageMessage(), OnGetDragImage)
			//MESSAGE_HANDLER(GetBarMessage(), OnGetToolbar)
			INDEXED_MESSAGE_HANDLER(GetAutoPopupMessage(), OnAutoPopup)
			//MESSAGE_HANDLER(GetMDIChildWindowStateChangedMessage(), OnMDIChildWindowStateChanged)
			//MESSAGE_HANDLER(GetDisplayChevronPopupMessage(), OnDisplayChevronPopup)

			//MESSAGE_HANDLER(TB_ADDBUTTONSA, OnAddButtons)
			//MESSAGE_HANDLER(TB_ADDBUTTONSW, OnAddButtons)
			//MESSAGE_HANDLER(TB_DELETEBUTTON, OnDeleteButton)
			//MESSAGE_HANDLER(TB_INSERTBUTTONA, OnInsertButton)
			//MESSAGE_HANDLER(TB_INSERTBUTTONW, OnInsertButton)
			//MESSAGE_HANDLER(TB_SETBUTTONINFOA, OnSetButtonInfo)
			//MESSAGE_HANDLER(TB_SETBUTTONINFOW, OnSetButtonInfo)
			//MESSAGE_HANDLER(TB_SETBUTTONWIDTH, OnSetButtonWidth)
			//MESSAGE_HANDLER(TB_SETCMDID, OnSetCmdId)
			//MESSAGE_HANDLER(TB_SETDRAWTEXTFLAGS, OnSetDrawTextFlags)
			//MESSAGE_HANDLER(TB_SETLISTGAP, OnSetDropDownGap)
			//MESSAGE_HANDLER(TB_SETINDENT, OnSetIndent)
			//MESSAGE_HANDLER(TB_SETLISTGAP, OnSetListGap)

			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_BEGINADJUST, OnBeginAdjustNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_BEGINDRAG, OnBeginDragNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_CUSTHELP, OnCustHelpNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_DELETINGBUTTON, OnDeletingButtonNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_DROPDOWN, OnDropDownNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_DUPACCELERATOR, OnDupAcceleratorNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_ENDADJUST, OnEndAdjustNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETBUTTONINFOA, OnGetButtonInfoNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETBUTTONINFOW, OnGetButtonInfoNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETDISPINFOA, OnGetDispInfoNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETDISPINFOW, OnGetDispInfoNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETINFOTIPA, OnGetInfoTipNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETINFOTIPW, OnGetInfoTipNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_GETOBJECT, OnGetObjectNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_HOTITEMCHANGE, OnHotItemChangeNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_INITCUSTOMIZE, OnInitCustomizeNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_MAPACCELERATOR, OnMapAcceleratorNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_QUERYDELETE, OnQueryDeleteNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_QUERYINSERT, OnQueryInsertNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_RESET, OnResetNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_RESTORE, OnRestoreNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_SAVE, OnSaveNotification)
			//REFLECTED_NOTIFY_CODE_HANDLER(TBN_TOOLBARCHANGE, OnToolBarChangeNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(TBN_WRAPACCELERATOR, OnWrapAcceleratorNotification)
			INDEXED_REFLECTED_NOTIFY_CODE_HANDLER(TBN_WRAPHOTITEM, OnWrapHotItemNotification)

			INDEXED_REFLECTED_COMMAND_CODE_HANDLER(BN_CLICKED, OnReflectedClicked)
		END_MSG_MAP()
	#endif

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISupportErrorInfo
	///
	//@{
	/// \brief <em>Retrieves whether an interface supports the \c IErrorInfo interface</em>
	///
	/// \param[in] interfaceToCheck The IID of the interface to check.
	///
	/// \return \c S_OK if the interface identified by \c interfaceToCheck supports \c IErrorInfo;
	///         otherwise \c S_FALSE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221233.aspx">IErrorInfo</a>
	virtual HRESULT STDMETHODCALLTYPE InterfaceSupportsErrorInfo(REFIID interfaceToCheck);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of persistance
	///
	//@{
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Load to make the control persistent</em>
	///
	/// We want to persist an indexed property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Load and read directly from the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] pErrorLog The caller's \c IErrorLog implementation which will receive any errors
	///            that occur during property loading.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768206.aspx">IPersistPropertyBag::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPPROPERTYBAG pPropertyBag, LPERRORLOG pErrorLog);
	/// \brief <em>Overrides \c IPersistPropertyBagImpl::Save to make the control persistent</em>
	///
	/// We want to persist an indexed property. This can't be done by just using ATL's persistence
	/// macros. So we override \c IPersistPropertyBagImpl::Save and write directly into the property bag.
	///
	/// \param[in] pPropertyBag The \c IPropertyBag implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	/// \param[in] saveAllProperties Flag indicating whether the control should save all its properties
	///            (\c TRUE) or only those that have changed from the default value (\c FALSE).
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768207.aspx">IPersistPropertyBag::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa768196.aspx">IPropertyBag</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPPROPERTYBAG pPropertyBag, BOOL clearDirtyFlag, BOOL saveAllProperties);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::GetSizeMax to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we communicate directly with the stream. This requires we override
	/// \c IPersistStreamInitImpl::GetSizeMax.
	///
	/// \param[in] pSize The maximum number of bytes that persistence of the control's properties will
	///            consume.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, Save,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687287.aspx">IPersistStreamInit::GetSizeMax</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	virtual HRESULT STDMETHODCALLTYPE GetSizeMax(ULARGE_INTEGER* pSize);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Load to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Load and read directly from
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Save, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680730.aspx">IPersistStreamInit::Load</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Load(LPSTREAM pStream);
	/// \brief <em>Overrides \c IPersistStreamInitImpl::Save to make object properties persistent</em>
	///
	/// Object properties can't be persisted through \c IPersistStreamInitImpl by just using ATL's
	/// persistence macros. So we override \c IPersistStreamInitImpl::Save and write directly into
	/// the stream.
	///
	/// \param[in] pStream The \c IStream implementation which stores the control's properties.
	/// \param[in] clearDirtyFlag Flag indicating whether the control should clear its dirty flag after
	///            saving. If \c TRUE, the flag is cleared, otherwise not. A value of \c FALSE allows
	///            the caller to do a "Save Copy As" operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Load, GetSizeMax,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694439.aspx">IPersistStreamInit::Save</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms682273.aspx">IPersistStreamInit</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/aa380034.aspx">IStream</a>
	virtual HRESULT STDMETHODCALLTYPE Save(LPSTREAM pStream, BOOL clearDirtyFlag);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IToolBar
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AllowCustomization property</em>
	///
	/// Retrieves whether the control can be customized via comctl32's built-in customization features.
	/// If set to \c VARIANT_TRUE, the built-in customization features are enabled; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The built-in customization features include a dialog for button management, that can be
	///          opened by double-clicking the control, and the capability to manage buttons using
	///          drag'n'drop.
	///
	/// \sa put_AllowCustomization, get_DragDropCustomizationModifierKey, Customize
	virtual HRESULT STDMETHODCALLTYPE get_AllowCustomization(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowCustomization property</em>
	///
	/// Sets whether the control can be customized via comctl32's built-in customization features.
	/// If set to \c VARIANT_TRUE, the built-in customization features are enabled; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The built-in customization features include a dialog for button management, that can be
	///          opened by double-clicking the control, and the capability to manage buttons using
	///          drag'n'drop.
	///
	/// \sa get_AllowCustomization, put_DragDropCustomizationModifierKey, Customize
	virtual HRESULT STDMETHODCALLTYPE put_AllowCustomization(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AlwaysDisplayButtonText property</em>
	///
	/// Retrieves whether the visibility of button's texts can be set on a per-button basis. If set to
	/// \c VARIANT_TRUE, the \c IToolBarButton::DisplayText property is ignored and button texts are always
	/// displayed; otherwise the visibility of the text depends on the setting of the
	/// \c IToolBarButton::DisplayText property.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c ButtonTextPosition is not set to \c btpRightToIcon.
	///
	/// \sa put_AlwaysDisplayButtonText, get_ButtonTextPosition, ToolBarButton::get_Text,
	///     ToolBarButton::get_DisplayText
	virtual HRESULT STDMETHODCALLTYPE get_AlwaysDisplayButtonText(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AlwaysDisplayButtonText property</em>
	///
	/// Sets whether the visibility of button's texts can be set on a per-button basis. If set to
	/// \c VARIANT_TRUE, the \c IToolBarButton::DisplayText property is ignored and button texts are always
	/// displayed; otherwise the visibility of the text depends on the setting of the
	/// \c IToolBarButton::DisplayText property.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c ButtonTextPosition is not set to \c btpRightToIcon.
	///
	/// \sa get_AlwaysDisplayButtonText, put_ButtonTextPosition, ToolBarButton::put_Text,
	///     ToolBarButton::put_DisplayText
	virtual HRESULT STDMETHODCALLTYPE put_AlwaysDisplayButtonText(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c AnchorHighlighting property</em>
	///
	/// Retrieves whether the highlighted button remains highlighted until another button is highlighted.
	/// If set to \c VARIANT_TRUE, the highlighted button remains highlighted even if the mouse cursor is
	/// moved outside the control's client area. The button remains highligthed until another button is
	/// highlighted. If set to \c VARIANT_FALSE, the highlighting is removed as soon as the mouse cursor
	/// leaves the button's bounding rectangle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AnchorHighlighting, get_HotButton
	virtual HRESULT STDMETHODCALLTYPE get_AnchorHighlighting(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AnchorHighlighting property</em>
	///
	/// Sets whether the highlighted button remains highlighted until another button is highlighted.
	/// If set to \c VARIANT_TRUE, the highlighted button remains highlighted even if the mouse cursor is
	/// moved outside the control's client area. The button remains highligthed until another button is
	/// highlighted. If set to \c VARIANT_FALSE, the highlighting is removed as soon as the mouse cursor
	/// leaves the button's bounding rectangle.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AnchorHighlighting, putref_HotButton
	virtual HRESULT STDMETHODCALLTYPE put_AnchorHighlighting(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Appearance property</em>
	///
	/// Retrieves the kind of border that is drawn around the control. Any of the values defined by
	/// the \c AppearanceConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Appearance, get_BorderStyle, TBarCtlsLibU::AppearanceConstants
	/// \else
	///   \sa put_Appearance, get_BorderStyle, TBarCtlsLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Appearance(AppearanceConstants* pValue);
	/// \brief <em>Sets the \c Appearance property</em>
	///
	/// Sets the kind of border that is drawn around the control. Any of the values defined by the
	/// \c AppearanceConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Appearance, put_BorderStyle, TBarCtlsLibU::AppearanceConstants
	/// \else
	///   \sa get_Appearance, put_BorderStyle, TBarCtlsLibA::AppearanceConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Appearance(AppearanceConstants newValue);
	/// \brief <em>Retrieves the control's application ID</em>
	///
	/// Retrieves the control's application ID. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppID(SHORT* pValue);
	/// \brief <em>Retrieves the control's application name</em>
	///
	/// Retrieves the control's application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppName(BSTR* pValue);
	/// \brief <em>Retrieves the control's short application name</em>
	///
	/// Retrieves the control's short application name. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The short application name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_Build, get_CharSet, get_IsRelease, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_AppShortName(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c BackStyle property</em>
	///
	/// Retrieves how the control's background is drawn. Any of the values defined by the
	/// \c BackStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c BackStyle property should be set to \c bksOpaque, if the \c Orientation property is
	///          set to \c oVertical and the control is not placed inside a rebar control.
	///
	/// \if UNICODE
	///   \sa put_BackStyle, get_ButtonStyle, get_Orientation, TBarCtlsLibU::BackStyleConstants
	/// \else
	///   \sa put_BackStyle, get_ButtonStyle, get_Orientation, TBarCtlsLibA::BackStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BackStyle(BackStyleConstants* pValue);
	/// \brief <em>Sets the \c BackStyle property</em>
	///
	/// Sets how the control's background is drawn. Any of the values defined by the
	/// \c BackStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c BackStyle property should be set to \c bksOpaque, if the \c Orientation property is
	///          set to \c oVertical and the control is not placed inside a rebar control.
	///
	/// \if UNICODE
	///   \sa get_BackStyle, put_ButtonStyle, put_Orientation, TBarCtlsLibU::BackStyleConstants
	/// \else
	///   \sa get_BackStyle, put_ButtonStyle, put_Orientation, TBarCtlsLibA::BackStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BackStyle(BackStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c BorderStyle property</em>
	///
	/// Retrieves the kind of inner border that is drawn around the control. Any of the values defined
	/// by the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_BorderStyle, get_Appearance, TBarCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa put_BorderStyle, get_Appearance, TBarCtlsLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_BorderStyle(BorderStyleConstants* pValue);
	/// \brief <em>Sets the \c BorderStyle property</em>
	///
	/// Sets the kind of inner border that is drawn around the control. Any of the values defined by
	/// the \c BorderStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_BorderStyle, put_Appearance, TBarCtlsLibU::BorderStyleConstants
	/// \else
	///   \sa get_BorderStyle, put_Appearance, TBarCtlsLibA::BorderStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_BorderStyle(BorderStyleConstants newValue);
	/// \brief <em>Retrieves the control's build number</em>
	///
	/// Retrieves the control's build number. This property is part of the fingerprint that
	/// uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The build number.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_CharSet, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Build(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ButtonHeight property</em>
	///
	/// Retrieves the height (in pixels) of a button. If set to 0, the system's default button height is
	/// used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ButtonHeight, get_ButtonWidth, get_IdealHeight
	virtual HRESULT STDMETHODCALLTYPE get_ButtonHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ButtonHeight property</em>
	///
	/// Sets the height (in pixels) of a button. If set to 0, the system's default button height is
	/// used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ButtonHeight, put_ButtonWidth, get_IdealHeight
	virtual HRESULT STDMETHODCALLTYPE put_ButtonHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c ButtonRowCount property</em>
	///
	/// Retrieves the number of rows used to display the tool bar buttons.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa SetButtonRowCount, get_WrapButtons
	virtual HRESULT STDMETHODCALLTYPE get_ButtonRowCount(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Buttons property</em>
	///
	/// Retrieves a collection object wrapping the control's buttons.
	///
	/// \param[out] ppButtons Receives the collection object's \c IToolBarButtons implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ToolBarButtons
	virtual HRESULT STDMETHODCALLTYPE get_Buttons(IToolBarButtons** ppButtons);
	/// \brief <em>Retrieves the current setting of the \c ButtonStyle property</em>
	///
	/// Retrieves the appearance of the tool bar buttons. Any of the values defined by the
	/// \c ButtonStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c ButtonStyle property should be set to \c bst3D, if the \c Orientation property is
	///          set to \c oVertical and the control is not placed inside a rebar control.
	///
	/// \if UNICODE
	///   \sa put_ButtonStyle, get_ButtonTextPosition, get_Orientation, TBarCtlsLibU::ButtonStyleConstants
	/// \else
	///   \sa put_ButtonStyle, get_ButtonTextPosition, get_Orientation, TBarCtlsLibA::ButtonStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ButtonStyle(ButtonStyleConstants* pValue);
	/// \brief <em>Sets the \c ButtonStyle property</em>
	///
	/// Sets the appearance of the tool bar buttons. Any of the values defined by the
	/// \c ButtonStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c ButtonStyle property should be set to \c bst3D, if the \c Orientation property is
	///          set to \c oVertical and the control is not placed inside a rebar control.
	///
	/// \if UNICODE
	///   \sa get_ButtonStyle, put_ButtonTextPosition, put_Orientation, TBarCtlsLibU::ButtonStyleConstants
	/// \else
	///   \sa get_ButtonStyle, put_ButtonTextPosition, put_Orientation, TBarCtlsLibA::ButtonStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ButtonStyle(ButtonStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ButtonTextPosition property</em>
	///
	/// Retrieves the position of the button texts relative to the button icons. Any of the values defined by
	/// the \c ButtonTextPositionConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ButtonTextPosition, get_ButtonStyle, ToolBarButton::get_Text,
	///       TBarCtlsLibU::ButtonTextPositionConstants
	/// \else
	///   \sa put_ButtonTextPosition, get_ButtonStyle, ToolBarButton::get_Text,
	///       TBarCtlsLibA::ButtonTextPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ButtonTextPosition(ButtonTextPositionConstants* pValue);
	/// \brief <em>Sets the \c ButtonTextPosition property</em>
	///
	/// Sets the position of the button texts relative to the button icons. Any of the values defined by
	/// the \c ButtonTextPositionConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ButtonTextPosition, put_ButtonStyle, ToolBarButton::put_Text
	///       TBarCtlsLibU::ButtonTextPositionConstants
	/// \else
	///   \sa get_ButtonTextPosition, put_ButtonStyle, ToolBarButton::put_Text,
	///       TBarCtlsLibA::ButtonTextPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ButtonTextPosition(ButtonTextPositionConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ButtonWidth property</em>
	///
	/// Retrieves the width (in pixels) of a button. If set to 0, the system's default button width is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ButtonWidth, get_ButtonHeight, get_MinimumButtonWidth, get_MaximumButtonWidth
	virtual HRESULT STDMETHODCALLTYPE get_ButtonWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c ButtonWidth property</em>
	///
	/// Sets the width (in pixels) of a button. If set to 0, the system's default button width is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ButtonWidth, put_ButtonHeight, put_MinimumButtonWidth, put_MaximumButtonWidth
	virtual HRESULT STDMETHODCALLTYPE put_ButtonWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the control's character set</em>
	///
	/// Retrieves the control's character set (Unicode/ANSI). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The character set.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_IsRelease, get_Programmer,
	///     get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_CharSet(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c CommandEnabled property</em>
	///
	/// Retrieves whether the specified command is enabled or disabled for execution. If set to
	/// \c VARIANT_TRUE, the control will raise the \c ExecuteCommand event if the command's associated
	/// button or menu entry is clicked or its associated hotkey is pressed; otherwise not.
	///
	/// \param[in] commandID The unique ID of the command for which to retrieve the current setting.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CommandEnabled, get_Enabled, ToolBarButton::get_Enabled, RegisterHotkey, Raise_ExecuteCommand
	virtual HRESULT STDMETHODCALLTYPE get_CommandEnabled(LONG commandID, VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CommandEnabled property</em>
	///
	/// Retrieves whether the specified command is enabled or disabled for execution. If set to
	/// \c VARIANT_TRUE, the control will raise the \c ExecuteCommand event if the command's associated
	/// button or menu entry is clicked or its associated hotkey is pressed; otherwise not.
	///
	/// \param[in] commandID The unique ID of the command to enable or disable.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If a tool bar button with the specified command ID exists, it will be enabled or disabled
	///          when setting this property.
	///
	/// \sa get_CommandEnabled, put_Enabled, ToolBarButton::put_Enabled, RegisterHotkey, Raise_ExecuteCommand
	virtual HRESULT STDMETHODCALLTYPE put_CommandEnabled(LONG commandID, VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisabledEvents property</em>
	///
	/// Retrieves the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DisabledEvents, TBarCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa put_DisabledEvents, TBarCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisabledEvents(DisabledEventsConstants* pValue);
	/// \brief <em>Sets the \c DisabledEvents property</em>
	///
	/// Sets the events that won't be fired. Disabling events increases performance. Any
	/// combination of the values defined by the \c DisabledEventsConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DisabledEvents, TBarCtlsLibU::DisabledEventsConstants
	/// \else
	///   \sa get_DisabledEvents, TBarCtlsLibA::DisabledEventsConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisabledEvents(DisabledEventsConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayMenuDivider property</em>
	///
	/// Retrieves whether a thin line is drawn on top of the control, that visually separates the tool bar
	/// from the menu. If set to \c VARIANT_TRUE, the divider is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property may destroy and recreate the control window.
	///
	/// \sa put_DisplayMenuDivider
	virtual HRESULT STDMETHODCALLTYPE get_DisplayMenuDivider(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayMenuDivider property</em>
	///
	/// Sets whether a thin line is drawn on top of the control, that visually separates the tool bar
	/// from the menu. If set to \c VARIANT_TRUE, the divider is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property may destroy and recreate the control window.
	///
	/// \sa get_DisplayMenuDivider
	virtual HRESULT STDMETHODCALLTYPE put_DisplayMenuDivider(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayPartiallyClippedButtons property</em>
	///
	/// Retrieves whether partially clipped buttons are displayed or hidden. If set to \c VARIANT_TRUE,
	/// partially clipped buttons are displayed; otherwise hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c WrapButtons property is set to \c VARIANT_TRUE.\n
	///          This property cannot be set to \c VARIANT_FALSE, if the \c Orientation property is set to
	///          \c oVertical.
	///
	/// \if UNICODE
	///   \sa put_DisplayPartiallyClippedButtons, get_Buttons, get_WrapButtons, get_Orientation,
	///       TBarCtlsLibU::OrientationConstants
	/// \else
	///   \sa put_DisplayPartiallyClippedButtons, get_Buttons, get_WrapButtons, get_Orientation,
	///       TBarCtlsLibA::OrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DisplayPartiallyClippedButtons(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayPartiallyClippedButtons property</em>
	///
	/// Sets whether partially clipped buttons are displayed or hidden. If set to \c VARIANT_TRUE,
	/// partially clipped buttons are displayed; otherwise hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c WrapButtons property is set to \c VARIANT_TRUE.\n
	///          This property cannot be set to \c VARIANT_FALSE, if the \c Orientation property is set to
	///          \c oVertical.
	///
	/// \if UNICODE
	///   \sa get_DisplayPartiallyClippedButtons, get_Buttons, put_WrapButtons, put_Orientation,
	///       TBarCtlsLibU::OrientationConstants
	/// \else
	///   \sa get_DisplayPartiallyClippedButtons, get_Buttons, put_WrapButtons, put_Orientation,
	///       TBarCtlsLibA::OrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DisplayPartiallyClippedButtons(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DontRedraw property</em>
	///
	/// Retrieves whether automatic redrawing of the control is enabled or disabled. If set to
	/// \c VARIANT_FALSE, the control will redraw itself automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE get_DontRedraw(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DontRedraw property</em>
	///
	/// Enables or disables automatic redrawing of the control. Disabling redraw while doing large changes
	/// on the control may increase performance. If set to \c VARIANT_FALSE, the control will redraw itself
	/// automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DontRedraw
	virtual HRESULT STDMETHODCALLTYPE put_DontRedraw(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DragClickTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be placed over a button during a
	/// drag'n'drop operation before this button will be clicked automatically. If set to 0,
	/// auto-clicking is disabled. If set to -1, the system's double-click time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DragClickTime, get_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE get_DragClickTime(LONG* pValue);
	/// \brief <em>Sets the \c DragClickTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be placed over a button during a
	/// drag'n'drop operation before this button will be clicked automatically. If set to 0,
	/// auto-clicking is disabled. If set to -1, the system's double-click time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DragClickTime, put_RegisterForOLEDragDrop, Raise_OLEDragMouseMove
	virtual HRESULT STDMETHODCALLTYPE put_DragClickTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c DragDropCustomizationModifierKey property</em>
	///
	/// Retrieves the modifier key that must be pressed to customize the tool bar (i. e. move buttons) using
	/// drag'n'drop. Any of the values defined by the \c DragDropCustomizationModifierKeyConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c AllowCustomization property is set to \c VARIANT_FALSE, this property is ignored.
	///
	/// \if UNICODE
	///   \sa put_DragDropCustomizationModifierKey, get_AllowCustomization, Customize,
	///       TBarCtlsLibU::DragDropCustomizationModifierKeyConstants
	/// \else
	///   \sa put_DragDropCustomizationModifierKey, get_AllowCustomization, Customize,
	///       TBarCtlsLibA::DragDropCustomizationModifierKeyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DragDropCustomizationModifierKey(DragDropCustomizationModifierKeyConstants* pValue);
	/// \brief <em>Sets the \c DragDropCustomizationModifierKey property</em>
	///
	/// Sets the modifier key that must be pressed to customize the tool bar (i. e. move buttons) using
	/// drag'n'drop. Any of the values defined by the \c DragDropCustomizationModifierKeyConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c AllowCustomization property is set to \c VARIANT_FALSE, this property is ignored.
	///
	/// \if UNICODE
	///   \sa get_DragDropCustomizationModifierKey, put_AllowCustomization, Customize,
	///       TBarCtlsLibU::DragDropCustomizationModifierKeyConstants
	/// \else
	///   \sa get_DragDropCustomizationModifierKey, put_AllowCustomization, Customize,
	///       TBarCtlsLibA::DragDropCustomizationModifierKeyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DragDropCustomizationModifierKey(DragDropCustomizationModifierKeyConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DraggedButtons property</em>
	///
	/// Retrieves a collection object wrapping the control's buttons that are currently dragged. These
	/// are the same buttons that were passed to the \c OLEDrag method.
	///
	/// \param[out] ppButtons Receives the collection object's \c IToolBarButtonContainer implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa OLEDrag, ToolBarButtonContainer, get_Buttons
	virtual HRESULT STDMETHODCALLTYPE get_DraggedButtons(IToolBarButtonContainer** ppButtons);
	/// \brief <em>Retrieves the current setting of the \c DropDownGap property</em>
	///
	/// Retrieves the extra space in pixels between a button's text and its drop down section. If set to -1,
	/// the system's default value is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c ButtonTextPosition property is set to a value other
	///          than \c btpRightToIcon or if the \c NormalDropDownButtonStyle property is set to a value
	///          other than \c nddbsSplitButton.
	///
	/// \sa put_DropDownGap, get_HorizontalButtonPadding, get_HorizontalIconCaptionGap,
	///     get_ButtonTextPosition, get_NormalDropDownButtonStyle
	virtual HRESULT STDMETHODCALLTYPE get_DropDownGap(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c DropDownGap property</em>
	///
	/// Sets the extra space in pixels between a button's text and its drop down section. If set to -1,
	/// the system's default value is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c ButtonTextPosition property is set to a value other
	///          than \c btpRightToIcon or if the \c NormalDropDownButtonStyle property is set to a value
	///          other than \c nddbsSplitButton.
	///
	/// \sa get_DropDownGap, put_HorizontalButtonPadding, put_HorizontalIconCaptionGap,
	///     put_ButtonTextPosition, put_NormalDropDownButtonStyle
	virtual HRESULT STDMETHODCALLTYPE put_DropDownGap(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the control is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled, get_CommandEnabled, ToolBarButton::get_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Enables or disables the control for user input. If set to \c VARIANT_TRUE, the control reacts
	/// to user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled, put_CommandEnabled, ToolBarButton::put_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FirstButtonIndentation property</em>
	///
	/// Retrieves the number of pixels between the control's left edge and the left edge of the first button
	/// in a row.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FirstButtonIndentation, get_HorizontalButtonSpacing
	virtual HRESULT STDMETHODCALLTYPE get_FirstButtonIndentation(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c FirstButtonIndentation property</em>
	///
	/// Sets the number of pixels between the control's left edge and the left edge of the first button
	/// in a row.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FirstButtonIndentation, put_HorizontalButtonSpacing
	virtual HRESULT STDMETHODCALLTYPE put_FirstButtonIndentation(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c FocusOnClick property</em>
	///
	/// Retrieves whether the control accepts the keyboard focus always or only if it is required to support
	/// keyboard navigation. If set to \c VARIANT_TRUE, the control accepts the keyboard focus always.
	/// If set to \c VARIANT_FALSE, keyboard focus is accepted in the following situations:
	/// - The control is in menu mode and the user closes the currently displayed menus by pressing
	///   the [ESC] button. In this case the control will actively steal the focus. It will set the focus
	///   back to the previously focused control if the user presses the [ESC] button again.
	/// - The focus is set to the control by walking through the dialog's tab order (using the [TAB]
	///   button).
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FocusOnClick, get_MenuMode
	virtual HRESULT STDMETHODCALLTYPE get_FocusOnClick(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c FocusOnClick property</em>
	///
	/// Sets whether the control accepts the keyboard focus always or only if it is required to support
	/// keyboard navigation. If set to \c VARIANT_TRUE, the control accepts the keyboard focus always.
	/// If set to \c VARIANT_FALSE, keyboard focus is accepted in the following situations:
	/// - The control is in menu mode and the user closes the currently displayed menus by pressing
	///   the [ESC] button. In this case the control will actively steal the focus. It will set the focus
	///   back to the previously focused control if the user presses the [ESC] button again.
	/// - The focus is set to the control by walking through the dialog's tab order (using the [TAB]
	///   button).
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FocusOnClick, get_MenuMode
	virtual HRESULT STDMETHODCALLTYPE put_FocusOnClick(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the button captions.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tool bars (except on Windows XP).
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_Font(IFontDisp** ppFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the button captions.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.\n
	///          This property isn't supported for themed tool bars (except on Windows XP).
	///
	/// \sa get_Font, putref_Font, put_UseSystemFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_Font(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the button captions.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tool bars (except on Windows XP).
	///
	/// \sa get_Font, put_Font, put_UseSystemFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c hChevronMenu property</em>
	///
	/// The chevron popup that the control may display is a simple popup menu (if the control is in menu
	/// mode) or a popup window that contains a tool bar control. This property retrieves the handle of the
	/// popup menu.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hWndChevronToolBar, DisplayChevronPopupWindow
	virtual HRESULT STDMETHODCALLTYPE get_hChevronMenu(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c HighlightColor property</em>
	///
	/// Retrieves the color used by the control to draw highlighted parts of tool bar buttons. If set to -1,
	/// the default color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tool bars.
	///
	/// \sa put_HighlightColor, get_ShadowColor
	virtual HRESULT STDMETHODCALLTYPE get_HighlightColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c HighlightColor property</em>
	///
	/// Sets the color used by the control to draw highlighted parts of tool bar buttons. If set to -1,
	/// the default color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tool bars.
	///
	/// \sa get_HighlightColor, put_ShadowColor
	virtual HRESULT STDMETHODCALLTYPE put_HighlightColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c hImageList property</em>
	///
	/// Retrieves the handle of the specified imagelist.
	///
	/// \param[in] imageList The imageList to retrieve or set. Most of the values defined by the
	///            \c ImageListConstants enumeration is valid.
	/// \param[in] imageListIndex The zero-based index of the image list of the type specified by
	///            \c imageList. The tool bar buttons' icons can be taken from different image lists, e. g.
	///            the 1st button's icon can be from image list A, the 2nd button's icon from image list B
	///            and so on.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa put_hImageList, ToolBarButton::get_IconIndex, ToolBarButton::get_ImageListIndex,
	///       get_ImageListCount, get_SuggestedIconSize, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa put_hImageList, ToolBarButton::get_IconIndex, ToolBarButton::get_ImageListIndex,
	///       get_ImageListCount, get_SuggestedIconSize, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_hImageList(ImageListConstants imageList, LONG imageListIndex = 0, OLE_HANDLE* pValue = NULL);
	/// \brief <em>Sets the \c hImageList property</em>
	///
	/// Sets the handle of the specified imagelist.
	///
	/// \param[in] imageList The imageList to retrieve or set. Most of the values defined by the
	///            \c ImageListConstants enumeration is valid.
	/// \param[in] imageListIndex The zero-based index of the image list of the type specified by
	///            \c imageList. The tool bar buttons' icons can be taken from different image lists, e. g.
	///            the 1st button's icon can be from image list A, the 2nd button's icon from image list B
	///            and so on.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa get_hImageList, ToolBarButton::put_IconIndex, ToolBarButton::put_ImageListIndex,
	///       get_ImageListCount, get_SuggestedIconSize, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa get_hImageList, ToolBarButton::put_IconIndex, ToolBarButton::put_ImageListIndex,
	///       get_ImageListCount, get_SuggestedIconSize, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_hImageList(ImageListConstants imageList, LONG imageListIndex = 0, OLE_HANDLE newValue = NULL);
	/// \brief <em>Retrieves the current setting of the \c HorizontalButtonPadding property</em>
	///
	/// Retrieves the number of pixels in horizontal direction between a button's border and its content. If
	/// set to -1, the system's default value is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Padding is applied to automatically-sized buttons only.\n
	///          This property is ignored if the \c AlwaysDisplayButtonText property is set to
	///          \c VARIANT_FALSE.
	///
	/// \sa put_HorizontalButtonPadding, get_VerticalButtonPadding, get_HorizontalIconCaptionGap,
	///     get_HorizontalButtonSpacing, ToolBarButton::get_AutoSize, get_AlwaysDisplayButtonText
	virtual HRESULT STDMETHODCALLTYPE get_HorizontalButtonPadding(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c HorizontalButtonPadding property</em>
	///
	/// Sets the number of pixels in horizontal direction between a button's border and its content. If
	/// set to -1, the system's default value is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Padding is applied to automatically-sized buttons only.\n
	///          This property is ignored if the \c AlwaysDisplayButtonText property is set to
	///          \c VARIANT_FALSE.
	///
	/// \sa get_HorizontalButtonPadding, put_VerticalButtonPadding, get_HorizontalIconCaptionGap,
	///     put_HorizontalButtonSpacing, ToolBarButton::put_AutoSize, put_AlwaysDisplayButtonText
	virtual HRESULT STDMETHODCALLTYPE put_HorizontalButtonPadding(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c HorizontalButtonSpacing property</em>
	///
	/// Retrieves the number of pixels between buttons in horizontal direction.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa put_HorizontalButtonSpacing, get_VerticalButtonSpacing, get_HorizontalIconCaptionGap,
	///     get_FirstButtonIndentation, get_HorizontalButtonPadding
	virtual HRESULT STDMETHODCALLTYPE get_HorizontalButtonSpacing(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c HorizontalButtonSpacing property</em>
	///
	/// Sets the number of pixels between buttons in horizontal direction.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa get_HorizontalButtonSpacing, put_VerticalButtonSpacing, put_HorizontalIconCaptionGap,
	///     put_FirstButtonIndentation, put_HorizontalButtonPadding
	virtual HRESULT STDMETHODCALLTYPE put_HorizontalButtonSpacing(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c HorizontalIconCaptionGap property</em>
	///
	/// Retrieves the number of pixels in horizontal direction between a button's icon and its text. If set
	/// to -1, the system's default value is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c ButtonTextPosition property is set to a value other
	///          than \c btpRightToIcon.
	///
	/// \sa put_HorizontalIconCaptionGap, get_HorizontalButtonPadding, get_DropDownGap,
	///     get_VerticalButtonSpacing, get_VerticalButtonSpacing, get_ButtonTextPosition
	virtual HRESULT STDMETHODCALLTYPE get_HorizontalIconCaptionGap(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c HorizontalIconCaptionGap property</em>
	///
	/// Sets the number of pixels in horizontal direction between a button's icon and its text. If set
	/// to -1, the system's default value is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c ButtonTextPosition property is set to a value other
	///          than \c btpRightToIcon.
	///
	/// \sa get_HorizontalIconCaptionGap, put_HorizontalButtonPadding, put_DropDownGap,
	///     put_VerticalButtonSpacing, put_VerticalButtonSpacing, put_ButtonTextPosition
	virtual HRESULT STDMETHODCALLTYPE put_HorizontalIconCaptionGap(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c HorizontalTextAlignment property</em>
	///
	/// Retrieves the horizontal alignment of the horizontal alignment of the button texts. Any of the values
	/// defined by the \c HAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_HorizontalTextAlignment, get_VerticalTextAlignment, ToolBarButton::get_Text,
	///       TBarCtlsLibU::HAlignmentConstants
	/// \else
	///   \sa put_HorizontalTextAlignment, get_VerticalTextAlignment, ToolBarButton::get_Text,
	///       TBarCtlsLibA::HAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_HorizontalTextAlignment(HAlignmentConstants* pValue);
	/// \brief <em>Sets the \c HorizontalTextAlignment property</em>
	///
	/// Sets the horizontal alignment of the horizontal alignment of the button texts. Any of the values
	/// defined by the \c HAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_HorizontalTextAlignment, put_VerticalTextAlignment, ToolBarButton::put_Text,
	///       TBarCtlsLibU::HAlignmentConstants, SetDrawTextFlags
	/// \else
	///   \sa get_HorizontalTextAlignment, put_VerticalTextAlignment, ToolBarButton::put_Text,
	///       TBarCtlsLibA::HAlignmentConstants, SetDrawTextFlags
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_HorizontalTextAlignment(HAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c HotButton property</em>
	///
	/// Retrieves the control's hot button. The hot button is the button under the mouse cursor. If set to
	/// \c NULL, the control does not have a hot button.
	///
	/// \param[out] ppHotButton Receives the hot item's \c IToolBarButton implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c ButtonStyle property is set to \c bst3D.
	///
	/// \sa putref_HotButton, SetHotButton, ToolBarButton::get_Hot, Raise_HotButtonChanging
	virtual HRESULT STDMETHODCALLTYPE get_HotButton(IToolBarButton** ppHotButton);
	/// \brief <em>Sets the \c HotButton property</em>
	///
	/// Sets the control's hot button. The hot button is the button under the mouse cursor. If set to
	/// \c NULL, the control does not have a hot button.
	///
	/// \param[in] pNewHotButton The new hot item's \c IToolBarButton implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored if the \c ButtonStyle property is set to \c bst3D.
	///
	/// \sa get_HotButton, SetHotButton, Raise_HotButtonChanging
	virtual HRESULT STDMETHODCALLTYPE putref_HotButton(IToolBarButton* pNewHotButton);
	/// \brief <em>Retrieves the current setting of the \c HoverTime property</em>
	///
	/// Retrieves the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE get_HoverTime(LONG* pValue);
	/// \brief <em>Sets the \c HoverTime property</em>
	///
	/// Sets the number of milliseconds the mouse cursor must be located over the control's client
	/// area before the \c MouseHover event is fired. If set to -1, the system hover time is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HoverTime, Raise_MouseHover
	virtual HRESULT STDMETHODCALLTYPE put_HoverTime(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndChevronToolBar property</em>
	///
	/// The chevron popup that the control may display is a simple popup menu (if the control is in menu
	/// mode) or a popup window that contains a tool bar control. This property retrieves the window handle
	/// of the tool bar control that is part of the chevron popup.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hChevronMenu, DisplayChevronPopupWindow, get_hWnd
	virtual HRESULT STDMETHODCALLTYPE get_hWndChevronToolBar(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndToolTip property</em>
	///
	/// Retrieves the tooltip control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hWndToolTip, get_ShowToolTips
	virtual HRESULT STDMETHODCALLTYPE get_hWndToolTip(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hWndToolTip property</em>
	///
	/// Sets the tooltip control to use.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_hWndToolTip, put_ShowToolTips
	virtual HRESULT STDMETHODCALLTYPE put_hWndToolTip(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c IdealHeight property</em>
	///
	/// Retrieves the control's ideal height, i. e. the height at which all buttons would be displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_IdealWidth, AutoSize, get_MaximumHeight
	virtual HRESULT STDMETHODCALLTYPE get_IdealHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c IdealWidth property</em>
	///
	/// Retrieves the control's ideal width, i. e. the width at which all buttons would be displayed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_IdealHeight, AutoSize, get_MaximumWidth
	virtual HRESULT STDMETHODCALLTYPE get_IdealWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c ImageListCount property</em>
	///
	/// Retrieves the number of image lists associated with the control.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hImageList
	virtual HRESULT STDMETHODCALLTYPE get_ImageListCount(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c InsertMarkColor property</em>
	///
	/// Retrieves the color that is the control's insertion mark is drawn in.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_InsertMarkColor, GetInsertMarkPosition
	virtual HRESULT STDMETHODCALLTYPE get_InsertMarkColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c InsertMarkColor property</em>
	///
	/// Sets the color that is the control's insertion mark is drawn in.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_InsertMarkColor, SetInsertMarkPosition
	virtual HRESULT STDMETHODCALLTYPE put_InsertMarkColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the control's release type</em>
	///
	/// Retrieves the control's release type. This property is part of the fingerprint that uniquely
	/// identifies each software written by Timo "TimoSoft" Kunze. If set to \c VARIANT_TRUE, the
	/// control was compiled for release; otherwise it was compiled for debugging.
	///
	/// \param[out] pValue The release type.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_Programmer, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_IsRelease(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c MaximumButtonWidth property</em>
	///
	/// Retrieves the maximum width of a button. If this width would be exceeded, the button's text is drawn
	/// with ellipsis.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MaximumButtonWidth, get_MinimumButtonWidth, get_ButtonWidth
	virtual HRESULT STDMETHODCALLTYPE get_MaximumButtonWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c MaximumButtonWidth property</em>
	///
	/// Sets the maximum width of a button. If this width would be exceeded, the button's text is drawn
	/// with ellipsis.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MaximumButtonWidth, put_MinimumButtonWidth, put_ButtonWidth
	virtual HRESULT STDMETHODCALLTYPE put_MaximumButtonWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c MaximumHeight property</em>
	///
	/// Retrieves the control's maximum height, i. e. the calculated total height of all visible buttons and
	/// separators together.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MaximumWidth, AutoSize, get_IdealHeight
	virtual HRESULT STDMETHODCALLTYPE get_MaximumHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c MaximumTextRows property</em>
	///
	/// Retrieves the maximum number of lines used to display button texts.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The button text is wrapped at word breaks only, line breaks are ignored.\n
	///          This property is ignored if the \c ButtonTextPosition property is set to
	///          \c btpRightToIcon.
	///
	/// \sa put_MaximumTextRows, ToolBarButton::get_Text, get_ButtonTextPosition
	virtual HRESULT STDMETHODCALLTYPE get_MaximumTextRows(LONG* pValue);
	/// \brief <em>Sets the \c MaximumTextRows property</em>
	///
	/// Sets the maximum number of lines used to display button texts.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The button text is wrapped at word breaks only, line breaks are ignored.\n
	///          This property is ignored if the \c ButtonTextPosition property is set to
	///          \c btpRightToIcon.
	///
	/// \sa get_MaximumTextRows, ToolBarButton::put_Text, put_ButtonTextPosition
	virtual HRESULT STDMETHODCALLTYPE put_MaximumTextRows(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c MaximumWidth property</em>
	///
	/// Retrieves the control's maximum width, i. e. the calculated total width of all visible buttons and
	/// separators together.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MaximumHeight, AutoSize, get_IdealWidth
	virtual HRESULT STDMETHODCALLTYPE get_MaximumWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c MenuBarTheme property</em>
	///
	/// Retrieves how the tool bar buttons are drawn if the control is in menu mode. Any of the values
	/// defined by the \c MenuBarThemeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MenuBarTheme, get_MenuMode, TBarCtlsLibU::MenuBarThemeConstants
	/// \else
	///   \sa put_MenuBarTheme, get_MenuMode, TBarCtlsLibA::MenuBarThemeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MenuBarTheme(MenuBarThemeConstants* pValue);
	/// \brief <em>Sets the \c MenuBarTheme property</em>
	///
	/// Sets how the tool bar buttons are drawn if the control is in menu mode. Any of the values
	/// defined by the \c MenuBarThemeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MenuBarTheme, put_MenuMode, TBarCtlsLibU::MenuBarThemeConstants
	/// \else
	///   \sa get_MenuBarTheme, put_MenuMode, TBarCtlsLibA::MenuBarThemeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MenuBarTheme(MenuBarThemeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c MenuMode property</em>
	///
	/// Retrieves whether the control emulates a menu bar. If set to \c VARIANT_TRUE, the control emulates
	/// a menu bar; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property cannot be set at runtime.\n
	///          Emulating a menu bar activates a couple of restrictions:
	///          - The \c AlwaysDisplayButtonText property has to be set to \c VARIANT_TRUE.
	///          - The \c ButtonTextPosition property has to be set to \c btpRightToIcon.
	///          - The \c HorizontalTextAlignment property has to be set to \c halCenter.
	///          - The \c NormalDropDownButtonStyle property has to be set to \c nddbsWithoutArrow.
	///          - The \c VerticalTextAlignment property has to be set to \c valCenter.
	///
	/// \sa put_MenuMode, get_MenuBarTheme, RegisterHotkey
	virtual HRESULT STDMETHODCALLTYPE get_MenuMode(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c MenuMode property</em>
	///
	/// Sets whether the control emulates a menu bar. If set to \c VARIANT_TRUE, the control emulates
	/// a menu bar; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property cannot be set at runtime.\n
	///          Emulating a menu bar activates a couple of restrictions:
	///          - The \c AlwaysDisplayButtonText property has to be set to \c VARIANT_TRUE.
	///          - The \c ButtonTextPosition property has to be set to \c btpRightToIcon.
	///          - The \c HorizontalTextAlignment property has to be set to \c halCenter.
	///          - The \c NormalDropDownButtonStyle property has to be set to \c nddbsWithoutArrow.
	///          - The \c VerticalTextAlignment property has to be set to \c valCenter.
	///
	/// \sa get_MenuMode, put_MenuBarTheme, RegisterHotkey
	virtual HRESULT STDMETHODCALLTYPE put_MenuMode(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c MinimumButtonWidth property</em>
	///
	/// Retrieves the minimum width of a button. Auto-sized buttons won't become smaller than this.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MinimumButtonWidth, get_MaximumButtonWidth, get_ButtonWidth
	virtual HRESULT STDMETHODCALLTYPE get_MinimumButtonWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c MinimumButtonWidth property</em>
	///
	/// Sets the minimum width of a button. Auto-sized buttons won't become smaller than this.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MinimumButtonWidth, put_MaximumButtonWidth, put_ButtonWidth
	virtual HRESULT STDMETHODCALLTYPE put_MinimumButtonWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c MouseIcon property</em>
	///
	/// Retrieves a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[out] ppMouseIcon Receives the picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, TBarCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa put_MouseIcon, putref_MouseIcon, get_MousePointer, TBarCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MouseIcon(IPictureDisp** ppMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewMouseIcon is cloned.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, TBarCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, putref_MouseIcon, put_MousePointer, TBarCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Sets the \c MouseIcon property</em>
	///
	/// Sets a user-defined mouse cursor. It's used if \c MousePointer is set to \c mpCustom.
	///
	/// \param[in] pNewMouseIcon The new picture object's \c IPictureDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, TBarCtlsLibU::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \else
	///   \sa get_MouseIcon, put_MouseIcon, put_MousePointer, TBarCtlsLibA::MousePointerConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms680762.aspx">IPictureDisp</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_MouseIcon(IPictureDisp* pNewMouseIcon);
	/// \brief <em>Retrieves the current setting of the \c MousePointer property</em>
	///
	/// Retrieves the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_MousePointer, get_MouseIcon, TBarCtlsLibU::MousePointerConstants
	/// \else
	///   \sa put_MousePointer, get_MouseIcon, TBarCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MousePointer(MousePointerConstants* pValue);
	/// \brief <em>Sets the \c MousePointer property</em>
	///
	/// Sets the cursor's type that's used if the mouse cursor is placed within the control's
	/// client area. Any of the values defined by the \c MousePointerConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MousePointer, put_MouseIcon, TBarCtlsLibU::MousePointerConstants
	/// \else
	///   \sa get_MousePointer, put_MouseIcon, TBarCtlsLibA::MousePointerConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_MousePointer(MousePointerConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c MultiColumn property</em>
	///
	/// Retrieves whether the control displays the buttons in multiple columns. If set to \c VARIANT_TRUE,
	/// the buttons are organized in multiple columns; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property cannot be set to \c VARIANT_TRUE, if the \c Orientation property is set to
	///          \c oHorizontal.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa put_MultiColumn, get_Orientation
	virtual HRESULT STDMETHODCALLTYPE get_MultiColumn(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c MultiColumn property</em>
	///
	/// Sets whether the control displays the buttons in multiple columns. If set to \c VARIANT_TRUE,
	/// the buttons are organized in multiple columns; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property cannot be set to \c VARIANT_TRUE, if the \c Orientation property is set to
	///          \c oHorizontal.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \sa get_MultiColumn, put_Orientation
	virtual HRESULT STDMETHODCALLTYPE put_MultiColumn(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c NativeDropTarget property</em>
	///
	/// Retrieves the native tool bar control's \c IDropTarget implementation.
	///
	/// \param[out] ppValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_RegisterForOLEDragDrop
	virtual HRESULT STDMETHODCALLTYPE get_NativeDropTarget(LPUNKNOWN* ppValue);
	/// \brief <em>Retrieves the current setting of the \c NormalDropDownButtonStyle property</em>
	///
	/// Retrieves how buttons with the \c DropDownStyle property being set to \c ddstNormal are drawn.
	/// Any of the values defined by the \c NormalDropDownButtonStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa put_NormalDropDownButtonStyle, ToolBarButton::get_DropDownStyle,
	///       TBarCtlsLibU::NormalDropDownButtonStyleConstants
	/// \else
	///   \sa put_NormalDropDownButtonStyle, ToolBarButton::get_DropDownStyle,
	///       TBarCtlsLibA::NormalDropDownButtonStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_NormalDropDownButtonStyle(NormalDropDownButtonStyleConstants* pValue);
	/// \brief <em>Sets the \c NormalDropDownButtonStyle property</em>
	///
	/// Sets how buttons with the \c DropDownStyle property being set to \c ddstNormal are drawn.
	/// Any of the values defined by the \c NormalDropDownButtonStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_NormalDropDownButtonStyle, ToolBarButton::put_DropDownStyle,
	///       TBarCtlsLibU::NormalDropDownButtonStyleConstants
	/// \else
	///   \sa get_NormalDropDownButtonStyle, ToolBarButton::put_DropDownStyle,
	///       TBarCtlsLibA::NormalDropDownButtonStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_NormalDropDownButtonStyle(NormalDropDownButtonStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c OLEDragImageStyle property</em>
	///
	/// Retrieves the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       TBarCtlsLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa put_OLEDragImageStyle, get_SupportOLEDragImages, OLEDrag,
	///       TBarCtlsLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_OLEDragImageStyle(OLEDragImageStyleConstants* pValue);
	/// \brief <em>Sets the \c OLEDragImageStyle property</em>
	///
	/// Sets the appearance of the OLE drag images generated by the control. Any of the values defined
	/// by the \c OLEDragImageStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       TBarCtlsLibU::OLEDragImageStyleConstants
	/// \else
	///   \sa get_OLEDragImageStyle, put_SupportOLEDragImages, OLEDrag,
	///       TBarCtlsLibA::OLEDragImageStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_OLEDragImageStyle(OLEDragImageStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Orientation property</em>
	///
	/// Retrieves the direction, in which the control displays buttons. Any of the values defined by the
	/// \c OrientationConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c oVertical, the \c DisplayPartiallyClippedButtons property
	///          must be set to \c VARIANT_TRUE. Additionally the \c BackStyle property should be set to
	///          \c bksOpaque and the \c ButtonStyle property should be set to \c bst3D, if the control
	///          is not placed inside a rebar control.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa put_Orientation, get_DisplayPartiallyClippedButtons, get_BackStyle, get_ButtonStyle,
	///       get_MultiColumn, TBarCtlsLibU::OrientationConstants
	/// \else
	///   \sa put_Orientation, get_DisplayPartiallyClippedButtons, get_BackStyle, get_ButtonStyle,
	///       get_MultiColumn, TBarCtlsLibA::OrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Orientation(OrientationConstants* pValue);
	/// \brief <em>Sets the \c Orientation property</em>
	///
	/// Sets the direction, in which the control displays buttons. Any of the values defined by the
	/// \c OrientationConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If this property is set to \c oVertical, the \c DisplayPartiallyClippedButtons property
	///          must be set to \c VARIANT_TRUE. Additionally the \c BackStyle property should be set to
	///          \c bksOpaque and the \c ButtonStyle property should be set to \c bst3D, if the control
	///          is not placed inside a rebar control.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_Orientation, put_DisplayPartiallyClippedButtons, put_BackStyle, put_ButtonStyle,
	///       put_MultiColumn, TBarCtlsLibU::OrientationConstants
	/// \else
	///   \sa get_Orientation, put_DisplayPartiallyClippedButtons, put_BackStyle, put_ButtonStyle,
	///       put_MultiColumn, TBarCtlsLibA::OrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Orientation(OrientationConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ProcessContextMenuKeys property</em>
	///
	/// Retrieves whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10] or
	/// [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the event is fired; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE get_ProcessContextMenuKeys(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ProcessContextMenuKeys property</em>
	///
	/// Sets whether the control fires the \c ContextMenu event if the user presses [SHIFT]+[F10] or
	/// [WINDOWS CONTEXTMENU]. If set to \c VARIANT_TRUE, the event is fired; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ProcessContextMenuKeys, Raise_ContextMenu
	virtual HRESULT STDMETHODCALLTYPE put_ProcessContextMenuKeys(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's programmer(s)</em>
	///
	/// Retrieves the name(s) of the control's programmer(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The programmer.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease, get_Tester
	virtual HRESULT STDMETHODCALLTYPE get_Programmer(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c RaiseCustomDrawEventOnEraseBackground property</em>
	///
	/// Retrieves whether the \c CustomDraw event is raised when erasing the control's background. If this
	/// property is set to \c VARIANT_TRUE, the event is raised when the control erases its background.
	/// Otherwise the event is not raised in this case.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RaiseCustomDrawEventOnEraseBackground, Raise_CustomDraw,
	///       TBarCtlsLibU::CustomDrawStageConstants
	/// \else
	///   \sa put_RaiseCustomDrawEventOnEraseBackground, Raise_CustomDraw,
	///       TBarCtlsLibA::CustomDrawStageConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RaiseCustomDrawEventOnEraseBackground(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c RaiseCustomDrawEventOnEraseBackground property</em>
	///
	/// Sets whether the \c CustomDraw event is raised when erasing the control's background. If this
	/// property is set to \c VARIANT_TRUE, the event is raised when the control erases its background.
	/// Otherwise the event is not raised in this case.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_RaiseCustomDrawEventOnEraseBackground, Raise_CustomDraw,
	///       TBarCtlsLibU::CustomDrawStageConstants
	/// \else
	///   \sa get_RaiseCustomDrawEventOnEraseBackground, Raise_CustomDraw,
	///       TBarCtlsLibA::CustomDrawStageConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RaiseCustomDrawEventOnEraseBackground(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RegisterForOLEDragDrop property</em>
	///
	/// Retrieves whether the control is registered as a target for OLE drag'n'drop. Any of the values
	/// defined by the \c RegisterForOLEDragDropConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RegisterForOLEDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter,
	///       TBarCtlsLibU::RegisterForOLEDragDropConstants
	/// \else
	///   \sa put_RegisterForOLEDragDrop, get_SupportOLEDragImages, Raise_OLEDragEnter,
	///       TBarCtlsLibA::RegisterForOLEDragDropConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants* pValue);
	/// \brief <em>Sets the \c RegisterForOLEDragDrop property</em>
	///
	/// Sets whether the control is registered as a target for OLE drag'n'drop. Any of the values
	/// defined by the \c RegisterForOLEDragDropConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_RegisterForOLEDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter,
	///       TBarCtlsLibU::RegisterForOLEDragDropConstants
	/// \else
	///   \sa get_RegisterForOLEDragDrop, put_SupportOLEDragImages, Raise_OLEDragEnter,
	///       TBarCtlsLibA::RegisterForOLEDragDropConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RegisterForOLEDragDrop(RegisterForOLEDragDropConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c RightToLeft property</em>
	///
	/// Retrieves whether bidirectional features are enabled or disabled. Any combination of the values
	/// defined by the \c RightToLeftConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_RightToLeft, TBarCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa put_RightToLeft, TBarCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_RightToLeft(RightToLeftConstants* pValue);
	/// \brief <em>Sets the \c RightToLeft property</em>
	///
	/// Enables or disables bidirectional features. Any combination of the values defined by the
	/// \c RightToLeftConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Setting or clearing the \c rtlLayout flag destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_RightToLeft, TBarCtlsLibU::RightToLeftConstants
	/// \else
	///   \sa get_RightToLeft, TBarCtlsLibA::RightToLeftConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_RightToLeft(RightToLeftConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c ShadowColor property</em>
	///
	/// Retrieves the color used by the control to draw shadowed parts of tool bar buttons. If set to -1, the
	/// default color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tool bars.
	///
	/// \sa put_ShadowColor, get_HighlightColor
	virtual HRESULT STDMETHODCALLTYPE get_ShadowColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ShadowColor property</em>
	///
	/// Sets the color used by the control to draw shadowed parts of tool bar buttons. If set to -1, the
	/// default color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tool bars.
	///
	/// \sa get_ShadowColor, put_HighlightColor
	virtual HRESULT STDMETHODCALLTYPE put_ShadowColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowDragImage property</em>
	///
	/// Retrieves whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible;
	/// otherwise hidden.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowDragImage, get_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE get_ShowDragImage(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowDragImage property</em>
	///
	/// Sets whether the drag image is visible or hidden. If set to \c VARIANT_TRUE, it is visible; otherwise
	/// hidden.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowDragImage, put_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE put_ShowDragImage(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ShowToolTips property</em>
	///
	/// Retrieves whether the control displays any tooltips. If this property is set to \c VARIANT_TRUE,
	/// the buttons' tooltips retrieved through the \c ButtonGetInfoTipText event are displayed. Otherwise
	/// no tooltips are shown.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowToolTips, get_hWndToolTip, Raise_ButtonGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE get_ShowToolTips(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowToolTips property</em>
	///
	/// Sets whether the control displays any tooltips. If this property is set to \c VARIANT_TRUE,
	/// the buttons' tooltips retrieved through the \c ButtonGetInfoTipText event are displayed. Otherwise
	/// no tooltips are shown.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowToolTips, put_hWndToolTip, Raise_ButtonGetInfoTipText
	virtual HRESULT STDMETHODCALLTYPE put_ShowToolTips(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SuggestedIconSize property</em>
	///
	/// Retrieves a value specifying which icon size should be used for the tool bar button icons. This
	/// suggestion is based upon the system's DPI settings. Any of the values defined by the
	/// \c SuggestedIconSizeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_hImageList, TBarCtlsLibU::SuggestedIconSizeConstants
	/// \else
	///   \sa get_hImageList, TBarCtlsLibA::SuggestedIconSizeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SuggestedIconSize(SuggestedIconSizeConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c SupportOLEDragImages property</em>
	///
	/// Retrieves whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, get_hImageList, get_ShowDragImage,
	///     get_OLEDragImageStyle, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE get_SupportOLEDragImages(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c SupportOLEDragImages property</em>
	///
	/// Sets whether the control creates an \c IDropTargetHelper object, so that a drag image can be
	/// shown during OLE drag'n'drop. If set to \c VARIANT_TRUE, the control creates the object; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires shell32.dll version 5.0 or higher.
	///
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, put_hImageList, put_ShowDragImage,
	///     put_OLEDragImageStyle, FinishOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
	virtual HRESULT STDMETHODCALLTYPE put_SupportOLEDragImages(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the name(s) of the control's tester(s)</em>
	///
	/// Retrieves the name(s) of the control's tester(s). This property is part of the fingerprint
	/// that uniquely identifies each software written by Timo "TimoSoft" Kunze.
	///
	/// \param[out] pValue The name(s) of the tester(s).
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is hidden and read-only.
	///
	/// \sa get_AppID, get_AppName, get_AppShortName, get_Build, get_CharSet, get_IsRelease,
	///     get_Programmer
	virtual HRESULT STDMETHODCALLTYPE get_Tester(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c UseMnemonics property</em>
	///
	/// Retrieves whether the control handles ampersands in button texts as accelerator keys. If set to
	/// \c VARIANT_TRUE, ampersands are handled as marking accelerator keys; otherwise they are handled as
	/// normal characters.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseMnemonics, ToolBarButton::get_Text
	virtual HRESULT STDMETHODCALLTYPE get_UseMnemonics(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseMnemonics property</em>
	///
	/// Sets whether the control handles ampersands in button texts as accelerator keys. If set to
	/// \c VARIANT_TRUE, ampersands are handled as marking accelerator keys; otherwise they are handled as
	/// normal characters.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseMnemonics, ToolBarButton::put_Text
	virtual HRESULT STDMETHODCALLTYPE put_UseMnemonics(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c UseSystemFont property</em>
	///
	/// Retrieves whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tool bars (except on Windows XP).
	///
	/// \sa put_UseSystemFont, get_Font
	virtual HRESULT STDMETHODCALLTYPE get_UseSystemFont(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseSystemFont property</em>
	///
	/// Sets whether the control uses the MS Shell Dlg font (which is mapped to the system's default GUI
	/// font) or the font specified by the \c Font property. If set to \c VARIANT_TRUE, the system font;
	/// otherwise the specified font is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed tool bars (except on Windows XP).
	///
	/// \sa get_UseSystemFont, put_Font, putref_Font
	virtual HRESULT STDMETHODCALLTYPE put_UseSystemFont(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \param[out] pValue The control's version.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	virtual HRESULT STDMETHODCALLTYPE get_Version(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c VerticalButtonPadding property</em>
	///
	/// Retrieves the number of pixels in vertical direction between a button's border and its content. If
	/// set to -1, the system's default value is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Padding is applied to automatically-sized buttons only.\n
	///          This property is ignored if the \c AlwaysDisplayButtonText property is set to
	///          \c VARIANT_FALSE.
	///
	/// \sa put_VerticalButtonPadding, get_HorizontalButtonPadding, get_VerticalButtonSpacing,
	///     ToolBarButton::get_AutoSize, get_AlwaysDisplayButtonText
	virtual HRESULT STDMETHODCALLTYPE get_VerticalButtonPadding(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c VerticalButtonPadding property</em>
	///
	/// Sets the number of pixels in vertical direction between a button's border and its content. If
	/// set to -1, the system's default value is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Padding is applied to automatically-sized buttons only.\n
	///          This property is ignored if the \c AlwaysDisplayButtonText property is set to
	///          \c VARIANT_FALSE.
	///
	/// \sa get_VerticalButtonPadding, put_HorizontalButtonPadding, put_VerticalButtonSpacing,
	///     ToolBarButton::put_AutoSize, put_AlwaysDisplayButtonText
	virtual HRESULT STDMETHODCALLTYPE put_VerticalButtonPadding(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c VerticalButtonSpacing property</em>
	///
	/// Retrieves the number of pixels between buttons in vertical direction.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa get_VerticalButtonSpacing, get_HorizontalButtonSpacing
	virtual HRESULT STDMETHODCALLTYPE get_VerticalButtonSpacing(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c VerticalButtonSpacing property</em>
	///
	/// Sets the number of pixels between buttons in vertical direction.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa put_VerticalButtonSpacing, put_HorizontalButtonSpacing
	virtual HRESULT STDMETHODCALLTYPE put_VerticalButtonSpacing(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c VerticalTextAlignment property</em>
	///
	/// Retrieves the vertical alignment of the horizontal alignment of the button texts. Any of the values
	/// defined by the \c VAlignmentConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_VerticalTextAlignment, get_HorizontalTextAlignment, ToolBarButton::get_Text,
	///       TBarCtlsLibU::VAlignmentConstants
	/// \else
	///   \sa put_VerticalTextAlignment, get_HorizontalTextAlignment, ToolBarButton::get_Text,
	///       TBarCtlsLibA::VAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_VerticalTextAlignment(VAlignmentConstants* pValue);
	/// \brief <em>Sets the \c VerticalTextAlignment property</em>
	///
	/// Sets the vertical alignment of the horizontal alignment of the button texts. Any of the values
	/// defined by the \c VAlignmentConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_VerticalTextAlignment, put_HorizontalTextAlignment, ToolBarButton::put_Text,
	///       TBarCtlsLibU::VAlignmentConstants, SetDrawTextFlags
	/// \else
	///   \sa get_VerticalTextAlignment, put_HorizontalTextAlignment, ToolBarButton::put_Text,
	///       TBarCtlsLibA::VAlignmentConstants, SetDrawTextFlags
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_VerticalTextAlignment(VAlignmentConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c WrapButtons property</em>
	///
	/// Retrieves whether buttons may be displayed on multiple lines, if the tool bar becomes too narrow to
	/// include all buttons on the same line. If set to \c VARIANT_TRUE, buttons may be wrapped; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_WrapButtons, get_DisplayPartiallyClippedButtons, ToolBarButton::get_FollowedByLineBreak,
	///     ToolBarButton::get_Width, get_ButtonRowCount
	virtual HRESULT STDMETHODCALLTYPE get_WrapButtons(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c WrapButtons property</em>
	///
	/// Sets whether buttons may be displayed on multiple lines, if the tool bar becomes too narrow to
	/// include all buttons on the same line. If set to \c VARIANT_TRUE, buttons may be wrapped; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_WrapButtons, put_DisplayPartiallyClippedButtons, ToolBarButton::put_FollowedByLineBreak,
	///     ToolBarButton::put_Width, SetButtonRowCount
	virtual HRESULT STDMETHODCALLTYPE put_WrapButtons(VARIANT_BOOL newValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Advises the control to resize itself to the best extent</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_IdealWidth, get_IdealHeight
	virtual HRESULT STDMETHODCALLTYPE AutoSize(void);
	/// \brief <em>Counts the buttons that use the specified character as accelerator character</em>
	///
	/// Retrieves the number of tool bar buttons that use the specified character as accelerator character.
	///
	/// \param[in] accelerator The accelerator character for which to count the buttons.
	/// \param[out] pCount Receives the number of buttons with the specified accelerator character.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ToolBarButton::get_Text, Raise_MapAccelerator
	virtual HRESULT STDMETHODCALLTYPE CountAcceleratorOccurrences(SHORT accelerator, LONG* pCount);
	/// \brief <em>Creates a new \c ToolBarButtonContainer object</em>
	///
	/// Retrieves a new \c ToolBarButtonContainer object and fills it with the specified buttons.
	///
	/// \param[in] buttons The button(s) to add to the collection. May be either \c Empty, a button ID, a
	///            \c ToolBarButton object or a \c ToolBarButtons collection.
	/// \param[out] ppContainer Receives the new collection object's \c IToolBarButtonContainer
	///             implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ToolBarButtonContainer::Clone, ToolBarButtonContainer::Add
	virtual HRESULT STDMETHODCALLTYPE CreateButtonContainer(VARIANT buttons = _variant_t(DISP_E_PARAMNOTFOUND, VT_ERROR), IToolBarButtonContainer** ppContainer = NULL);
	/// \brief <em>Displays a dialog that allows customizing the tool bar</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowCustomization, get_DragDropCustomizationModifierKey, Raise_QueryInsertButton,
	///     Raise_CustomizedControl
	virtual HRESULT STDMETHODCALLTYPE Customize(void);
	/// \brief <em>Creates and displays a popup window with the hidden tool bar buttons</em>
	///
	/// Creates a popup window that contains the tool bar buttons that are currently not visible and displays
	/// this window at the specified location.
	///
	/// \param[in] x The x-coordinate (in pixels relative to the screen's upper-left corner) at which the
	///            popup window will be displayed.
	/// \param[in] y The y-coordinate (in pixels relative to the screen's upper-left corner) at which the
	///            popup window will be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ReBar::Raise_ChevronClick, DisplayChevronMenu
	virtual HRESULT STDMETHODCALLTYPE DisplayChevronPopupWindow(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Finishes a pending drop operation</em>
	///
	/// During a drag'n'drop operation the drag image is displayed until the \c OLEDragDrop event has been
	/// handled. This order is intended by Microsoft Windows. However, if a message box is displayed from
	/// within the \c OLEDragDrop event, or the drop operation cannot be performed asynchronously and takes
	/// a long time, it may be desirable to remove the drag image earlier.\n
	/// This method will break the intended order and finish the drag'n'drop operation (including removal
	/// of the drag image) immediately.
	///
	/// \remarks This method will fail if not called from the \c OLEDragDrop event handler or if no drag
	///          images are used.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_OLEDragDrop, get_SupportOLEDragImages
	virtual HRESULT STDMETHODCALLTYPE FinishOLEDragDrop(void);
	/// \brief <em>Proposes a position for the control's insertion mark</em>
	///
	/// Retrieves the insertion mark position that is closest to the specified point.
	///
	/// \param[in] x The x-coordinate (in pixels) of the point for which to retrieve the closest
	///            insertion mark position. It must be relative to the control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point for which to retrieve the closest
	///            insertion mark position. It must be relative to the control's upper-left corner.
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified button. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] ppToolButton Receives the \c IToolBarButton implementation of the button, at which
	///             the insertion mark should be displayed.
	/// \param[out] pIsOverButton If \c VARIANT_TRUE, the specified position is over or near the button
	///             specified by \c ppToolButton; otherwise it is over the control's background or outside
	///             the control's client area.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If \c ppToolButton is not \c NULL, but \c relativePosition is \c impNowhere, the
	///          specified position is directly over the button specified by \c ppToolButton.
	///
	/// \if UNICODE
	///   \sa SetInsertMarkPosition, GetInsertMarkPosition, TBarCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetInsertMarkPosition, GetInsertMarkPosition, TBarCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetClosestInsertMarkPosition(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, InsertMarkPositionConstants* pRelativePosition, IToolBarButton** ppToolButton, VARIANT_BOOL* pIsOverButton = NULL);
	/// \brief <em>Retrieves the command ID that is registered for the specified shortcut key combination</em>
	///
	/// Retrieves the command ID that has been registered for the specified shortcut key combination.
	///
	/// \param[in] modifierKeys Specifies which modifier keys must be pressed to trigger the command. Any
	///            combination of the values specified by the \c ModifierKeysConstants enumeration is valid.
	/// \param[in] acceleratorKeyCode Specifies the virtual key code of the key that must be pressed to
	///            trigger the command.
	/// \param[in] pCommandID Receives the command ID that is triggered if the key combination is pressed. If no
	///            command is registered for the specified hot key, -1 is returned.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa RegisterHotkey, UnregisterHotkey, Raise_ExecuteCommand, TBarCtlsLibU::ModifierKeysConstants
	/// \else
	///   \sa RegisterHotkey, UnregisterHotkey, Raise_ExecuteCommand, TBarCtlsLibA::ModifierKeysConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetCommandForHotkey(ModifierKeysConstants modifierKeys, SHORT acceleratorKeyCode, LONG* pCommandID);
	/// \brief <em>Retrieves the position of the control's insertion mark</em>
	///
	/// \param[out] pRelativePosition The insertion mark's position relative to the specified button. The
	///             following values, defined by the \c InsertMarkPositionConstants enumeration, are
	///             valid: \c impBefore, \c impAfter, \c impNowhere.
	/// \param[out] ppToolButton Receives the \c IToolBarButton implementation of the button, at which
	///             the insertion mark is shown.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetInsertMarkPosition, GetClosestInsertMarkPosition, TBarCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa SetInsertMarkPosition, GetClosestInsertMarkPosition, TBarCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetInsertMarkPosition(InsertMarkPositionConstants* pRelativePosition, IToolBarButton** ppToolButton);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in,out] pHitTestDetails Receives a value specifying the exact part of the control the
	///                specified point lies in. Any of the values defined by the \c HitTestConstants
	///                enumeration is valid.
	/// \param[out] ppHitButton Receives the "hit" tool bar button's \c IToolBarButton implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::HitTestConstants, HitTestChevronToolBar, HitTest
	/// \else
	///   \sa TBarCtlsLibA::HitTestConstants, HitTestChevronToolBar, HitTest
	/// \endif
	// \if UNICODE
	//   \sa get_ButtonBoundingBoxDefinition, TBarCtlsLibU::HitTestConstants, HitTestChevronToolBar, HitTest
	// \else
	//   \sa get_ButtonBoundingBoxDefinition, TBarCtlsLibA::HitTestConstants, HitTestChevronToolBar, HitTest
	// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IToolBarButton** ppHitButton);
	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the chevron tool bar control's parts that lie below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the chevron
	///            tool bar control's upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the chevron
	///            tool bar control's upper-left corner.
	/// \param[in,out] pHitTestDetails Receives a value specifying the exact part of the control the
	///                specified point lies in. Any of the values defined by the \c HitTestConstants
	///                enumeration is valid.
	/// \param[out] ppHitButton Receives the "hit" tool bar button's \c IToolBarButton implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::HitTestConstants, HitTest
	/// \else
	///   \sa TBarCtlsLibA::HitTestConstants, HitTest
	/// \endif
	// \if UNICODE
	//   \sa get_ButtonBoundingBoxDefinition, TBarCtlsLibU::HitTestConstants, HitTest
	// \else
	//   \sa get_ButtonBoundingBoxDefinition, TBarCtlsLibA::HitTestConstants, HitTest
	// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTestChevronToolBar(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IToolBarButton** ppHitButton);
	/// \brief <em>Loads a system's default image list</em>
	///
	/// Loads one of the system's default image lists, so that the images can be used as button images.
	/// For valid predefined image indexes see the \c SystemImageIndexConstants enumeration.
	///
	/// \param[in] imageListType The image list to load. Any of the values specified by the
	///            \c SystemImageListTypeConstants enumeration is valid.
	/// \param[out] pImageCount Will be set to the number of images that the loaded image list contains.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_hImageList, ToolBarButton::get_IconIndex, TBarCtlsLibU::SystemImageListTypeConstants,
	///       TBarCtlsLibU::SystemImageIndexConstants
	/// \else
	///   \sa get_hImageList, ToolBarButton::get_IconIndex, TBarCtlsLibA::SystemImageListTypeConstants,
	///       TBarCtlsLibA::SystemImageIndexConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE LoadDefaultImages(SystemImageListTypeConstants imageListType, LONG* pImageCount);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile, LoadToolBarStateFromRegistry
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Loads the control's state from the specified value in the registry</em>
	///
	/// \param[in] subKeyName The name of the sub key (relative to \c hParentKey) in the registry where to
	///            load the tool bar state from.
	/// \param[in] valueName The name of the registry value where to load the tool bar state from.
	/// \param[in] hParentKey The handle of the parent key of \c subKeyName. By default this is
	///            \c HKEY_CURRENT_USER.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveToolBarStateToRegistry, LoadSettingsFromFile
	virtual HRESULT STDMETHODCALLTYPE LoadToolBarStateFromRegistry(BSTR subKeyName, BSTR valueName, OLE_HANDLE hParentKey = 0x80000001, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Enters OLE drag'n'drop mode</em>
	///
	/// \param[in] pDataObject A pointer to the \c IDataObject implementation to use during OLE
	///            drag'n'drop. If not specified, the control's own implementation is used.
	/// \param[in] supportedEffects A bit field defining all drop effects the client wants to support.
	///            Any combination of the values defined by the \c OLEDropEffectConstants enumeration
	///            (except \c odeScroll) is valid.
	/// \param[in] hWndToAskForDragImage The handle of the window, that is awaiting the
	///            \c DI_GETDRAGIMAGE message to specify the drag image to use. If -1, the method
	///            creates the drag image itself. If \c SupportOLEDragImages is set to \c VARIANT_FALSE,
	///            no drag image is used.
	/// \param[in] pDraggedButtons The dragged buttons collection object's \c IToolBarButtonContainer
	///            implementation. It is used to generate the drag image and is ignored if
	///            \c hWndToAskForDragImage is not -1.
	/// \param[in] itemCountToDisplay The number to display in the item count label of Aero drag images.
	///            If set to 0 or 1, no item count label is displayed. If set to -1, the number of buttons
	///            contained in the \c pDraggedButtons collection is displayed in the item count label. If
	///            set to any value larger than 1, this value is displayed in the item count label.
	/// \param[out] pPerformedEffects The performed drop effect. Any of the values defined by the
	///             \c OLEDropEffectConstants enumeration (except \c odeScroll) is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Raise_ButtonBeginDrag, Raise_ButtonBeginRDrag, Raise_OLEStartDrag,
	///       get_SupportOLEDragImages, get_OLEDragImageStyle, TBarCtlsLibU::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \else
	///   \sa Raise_ButtonBeginDrag, Raise_ButtonBeginRDrag, Raise_OLEStartDrag,
	///       get_SupportOLEDragImages, get_OLEDragImageStyle, TBarCtlsLibA::OLEDropEffectConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE OLEDrag(LONG* pDataObject = NULL, OLEDropEffectConstants supportedEffects = odeCopyOrMove, OLE_HANDLE hWndToAskForDragImage = -1, IToolBarButtonContainer* pDraggedButtons = NULL, LONG itemCountToDisplay = -1, OLEDropEffectConstants* pPerformedEffects = NULL);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	/// \brief <em>Connects a command ID with a shortcut key combination</em>
	///
	/// Connects a command ID with a shortcut key combination so that pressing the key combination triggers
	/// the command.
	///
	/// \param[in] modifierKeys Specifies which modifier keys must be pressed to trigger the command. Any
	///            combination of the values specified by the \c ModifierKeysConstants enumeration is valid.
	/// \param[in] acceleratorKeyCode Specifies the virtual key code of the key that must be pressed to
	///            trigger the command.
	/// \param[in] commandID Specifies the command to trigger if the key combination is pressed. If the
	///            command is triggered, the \c ExecuteCommand event is fired.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa UnregisterHotkey, GetCommandForHotkey, Raise_ExecuteCommand,
	///       TBarCtlsLibU::ModifierKeysConstants
	/// \else
	///   \sa UnregisterHotkey, GetCommandForHotkey, Raise_ExecuteCommand,
	///       TBarCtlsLibA::ModifierKeysConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE RegisterHotkey(ModifierKeysConstants modifierKeys, SHORT acceleratorKeyCode, LONG commandID);
	/// \brief <em>Saves the control's settings to the specified file</em>
	///
	/// \param[in] file The file to write to.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadSettingsFromFile, SaveToolBarStateToRegistry
	virtual HRESULT STDMETHODCALLTYPE SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Saves the control's state to the specified value in the registry</em>
	///
	/// \param[in] subKeyName The name of the sub key (relative to \c hParentKey) in the registry where to
	///            store the tool bar state.
	/// \param[in] valueName The name of the registry value where to store the tool bar state.
	/// \param[in] hParentKey The handle of the parent key of \c subKeyName. By default this is
	///            \c HKEY_CURRENT_USER.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadToolBarStateFromRegistry, SaveSettingsToFile, Raise_InitializeToolBarStateRegistryStorage
	virtual HRESULT STDMETHODCALLTYPE SaveToolBarStateToRegistry(BSTR subKeyName, BSTR valueName, OLE_HANDLE hParentKey = 0x80000001, VARIANT_BOOL* pSucceeded = NULL);
	/// \brief <em>Sets the number of rows used to display the tool bar buttons</em>
	///
	/// Sets the number of rows used to display the tool bar buttons.
	///
	/// \param[in] rowCount The number of rows that the control shall display the tool bar buttons in.
	/// \param[in] allowMoreRows If set to \c VARIANT_TRUE, the control is allowed to create more rows than
	///            specified by \c rowCount, for instance if the control's width is too small to display all
	///            buttons with the specified number of rows. If set to \c VARIANT_FALSE, the control won't
	///            create more rows than specified.
	/// \param[out] pControlWidth Receives the control's width (in pixels) after the buttons have been
	///             reorganized.
	/// \param[out] pControlHeight Receives the control's height (in pixels) after the buttons have been
	///             reorganized.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The control does not break button groups, therefore the resulting number of rows might be
	///          different than specified.
	///
	/// \sa get_ButtonRowCount, put_WrapButtons, ToolBarButton::put_FollowedByLineBreak
	virtual HRESULT STDMETHODCALLTYPE SetButtonRowCount(LONG rowCount, VARIANT_BOOL allowMoreRows = VARIANT_TRUE, OLE_XSIZE_PIXELS* pControlWidth = NULL, OLE_YSIZE_PIXELS* pControlHeight = NULL);
	/// \brief <em>Sets the control's hot button</em>
	///
	/// Sets the control's hot button. The hot button is the button under the mouse cursor. If set to
	/// \c NULL, the control does not have a hot button.\n
	/// Other than the \c HotButton property, this method allows specifying additional parameters and
	/// returns the previous hot button.
	///
	/// \param[in] pNewHotButton The tool bar button to set as the control's hot button. If \c NULL, the
	///            control's hot button is cleared.
	/// \param[in] hotButtonChangeReason The reason for changing the hot button. This value is passed to
	///            the \c HotButtonChanging event. Any combination of the values defined by the
	///            \c HotButtonChangingCausedByConstants enumeration is valid.
	/// \param[in] additionalInfo Additional information describing the change of the hot button. This
	///            value is passed to the \c HotButtonChanging event. Any combination of the values defined
	///            by the \c HotButtonChangingAdditionalInfoConstants enumeration is valid.
	/// \param[out] ppPreviousHotButton Receives the \c IToolBarButton implementation of the control's
	///             previous hot button. This may be \c NULL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This method fails if the \c ButtonStyle property is set to \c bst3D.\n
	///          Specifying the \c hbcaiDoDropDownClick flag in the \c additionalInfo parameter causes a
	///          programmatic click on the new hot button's drop-down arrow.
	///
	/// \if UNICODE
	///   \sa putref_HotButton, get_ButtonStyle, Raise_HotButtonChanging,
	///       TBarCtlsLibU::HotButtonChangingCausedByConstants,
	///       TBarCtlsLibU::HotButtonChangingAdditionalInfoConstants, ToolBarButton::get_DropDownStyle
	/// \else
	///   \sa putref_HotButton, get_ButtonStyle, Raise_HotButtonChanging,
	///       TBarCtlsLibA::HotButtonChangingCausedByConstants,
	///       TBarCtlsLibA::HotButtonChangingAdditionalInfoConstants, ToolBarButton::get_DropDownStyle
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetHotButton(IToolBarButton* pNewHotButton, HotButtonChangingCausedByConstants hotButtonChangeReason = hbccbOther, HotButtonChangingAdditionalInfoConstants additionalInfo = static_cast<HotButtonChangingAdditionalInfoConstants>(0), IToolBarButton** ppPreviousHotButton = NULL);
	/// \brief <em>Sets the position of the control's insertion mark</em>
	///
	/// \param[in] relativePosition The insertion mark's position relative to the specified button. Any of
	///            the values defined by the \c InsertMarkPositionConstants enumeration is valid.
	/// \param[in] pToolButton The \c IToolBarButton implementation of the button, at which the insertion
	///            mark will be shown. If set to \c NULL, the insertion mark will be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetInsertMarkPosition, GetClosestInsertMarkPosition, get_InsertMarkColor,
	///       get_RegisterForOLEDragDrop, TBarCtlsLibU::InsertMarkPositionConstants
	/// \else
	///   \sa GetInsertMarkPosition, GetClosestInsertMarkPosition, get_InsertMarkColor,
	///       get_RegisterForOLEDragDrop, TBarCtlsLibA::InsertMarkPositionConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetInsertMarkPosition(InsertMarkPositionConstants relativePosition, IToolBarButton* pToolButton);
	/// \brief <em>Disconnects a command ID from a shortcut key combination</em>
	///
	/// Removes the connection of a shortcut key combination with a command ID so that pressing the key
	/// combination no longer triggers the command.
	///
	/// \param[in] modifierKeys Specifies which modifier keys must be pressed to trigger the command. Any
	///            combination of the values specified by the \c ModifierKeysConstants enumeration is valid.
	/// \param[in] acceleratorKeyCode Specifies the virtual key code of the key that must be pressed to
	///            trigger the command.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa RegisterHotkey, GetCommandForHotkey, TBarCtlsLibU::ModifierKeysConstants
	/// \else
	///   \sa RegisterHotkey, GetCommandForHotkey, TBarCtlsLibA::ModifierKeysConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE UnregisterHotkey(ModifierKeysConstants modifierKeys, SHORT acceleratorKeyCode);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Property object changes
	///
	//@{
	/// \brief <em>Will be called after a property object was changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that was changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnChanged
	HRESULT OnPropertyObjectChanged(DISPID propertyObject, DISPID /*objectProperty*/);
	/// \brief <em>Will be called before a property object is changed</em>
	///
	/// \param[in] propertyObject The \c DISPID of the property object.
	/// \param[in] objectProperty The \c DISPID of the property that is about to be changed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa PropertyNotifySinkImpl::OnRequestEdit
	HRESULT OnPropertyObjectRequestEdit(DISPID /*propertyObject*/, DISPID /*objectProperty*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Called to create the control window</em>
	///
	/// Called to create the control window. This method overrides \c CWindowImpl::Create() and is
	/// used to customize the window styles.
	///
	/// \param[in] hWndParent The control's parent window.
	/// \param[in] rect The control's bounding rectangle.
	/// \param[in] szWindowName The control's window name.
	/// \param[in] dwStyle The control's window style. Will be ignored.
	/// \param[in] dwExStyle The control's extended window style. Will be ignored.
	/// \param[in] MenuOrID The control's ID.
	/// \param[in] lpCreateParam The window creation data. Will be passed to the created window.
	///
	/// \return The created window's handle.
	///
	/// \sa OnCreate, GetStyleBits, GetExStyleBits
	HWND Create(HWND hWndParent, ATL::_U_RECT rect = NULL, LPCTSTR szWindowName = NULL, DWORD dwStyle = 0, DWORD dwExStyle = 0, ATL::_U_MENUorID MenuOrID = 0U, LPVOID lpCreateParam = NULL);
	/// \brief <em>Called to draw the control</em>
	///
	/// Called to draw the control. This method overrides \c CComControlBase::OnDraw() and is used to prevent
	/// the "ATL 9.0" drawing in user mode and replace it in design mode.
	///
	/// \param[in] drawInfo Contains any details like the device context required for drawing.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/056hw3hs.aspx">CComControlBase::OnDraw</a>
	virtual HRESULT OnDraw(ATL_DRAWINFO& drawInfo);
	/// \brief <em>Called after receiving the last message (typically \c WM_NCDESTROY)</em>
	///
	/// \param[in] hWnd The window being destroyed.
	///
	/// \sa OnCreate, OnDestroy
	void OnFinalMessage(HWND /*hWnd*/);
	/// \brief <em>Informs an embedded object of its display location within its container</em>
	///
	/// \param[in] pClientSite The \c IOleClientSite implementation of the container application's
	///            client side.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms684013.aspx">IOleObject::SetClientSite</a>
	virtual HRESULT STDMETHODCALLTYPE IOleObject_SetClientSite(LPOLECLIENTSITE pClientSite);
	/// \brief <em>Notifies the control when the container's document window is activated or deactivated</em>
	///
	/// ATL's implementation of \c OnDocWindowActivate calls \c IOleInPlaceObject_UIDeactivate if the control
	/// is deactivated. This causes a bug in MDI apps. If the control sits on a \c MDI child window and has
	/// the focus and the title bar of another top-level window (not the MDI parent window) of the same
	/// process is clicked, the focus is moved from the ATL based ActiveX control to the next control on the
	/// MDI child before it is moved to the other top-level window that was clicked. If the focus is set back
	/// to the MDI child, the ATL based control no longer has the focus.
	///
	/// \param[in] fActivate If \c TRUE, the document window is activated; otherwise deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/0kz79wfc.aspx">IOleInPlaceActiveObjectImpl::OnDocWindowActivate</a>
	virtual HRESULT STDMETHODCALLTYPE OnDocWindowActivate(BOOL /*fActivate*/);

	/// \brief <em>A keyboard or mouse message needs to be translated</em>
	///
	/// The control's container calls this method if it receives a keyboard or mouse message. It gives
	/// us the chance to customize keystroke translation (i. e. to react to them in a non-default way).
	/// This method overrides \c CComControlBase::PreTranslateAccelerator.
	///
	/// \param[in] pMessage A \c MSG structure containing details about the received window message.
	/// \param[out] hReturnValue A reference parameter of type \c HRESULT which will be set to \c S_OK,
	///             if the message was translated, and to \c S_FALSE otherwise.
	///
	/// \return \c FALSE if the object's container should translate the message; otherwise \c TRUE.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/hxa56938.aspx">CComControlBase::PreTranslateAccelerator</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646373.aspx">TranslateAccelerator</a>
	BOOL PreTranslateAccelerator(LPMSG pMessage, HRESULT& hReturnValue);

	/// \brief <em>A CBT event occured</em>
	///
	/// This function is the callback function that we defined when we installed a CBT hook to trap the
	/// creation and destruction of menu windows.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam Depends on the nCode parameter.
	/// \param[in] lParam Depends on the nCode parameter.
	///
	/// \return The value returned by \c CallNextHookEx or 0 if \c CallNextHookEx was not called.
	///
	/// \sa DoTrackPopupMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644977.aspx">CBTProc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644974.aspx">CallNextHookEx</a>
	static LRESULT CALLBACK CBTProc(int code, WPARAM wParam, LPARAM lParam);

	//////////////////////////////////////////////////////////////////////
	/// \name Drag image creation
	///
	//@{
	/// \brief <em>Creates a legacy drag image for the specified button</em>
	///
	/// Creates a drag image for the specified button in the style of Windows versions prior to Vista. The
	/// drag image is added to an imagelist which is returned.
	///
	/// \param[in] buttonIndex The button for which to create the drag image.
	/// \param[out] pUpperLeftPoint Receives the coordinates (in pixels) of the drag image's upper-left
	///             corner relative to the control's upper-left corner.
	/// \param[out] pBoundingRectangle Receives the drag image's bounding rectangle (in pixels) relative to
	///             the control's upper-left corner.
	///
	/// \return An imagelist containing the drag image.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa CreateLegacyOLEDragImage, ToolBarButtonContainer::CreateDragImage
	HIMAGELIST CreateLegacyDragImage(int buttonIndex, LPPOINT pUpperLeftPoint, LPRECT pBoundingRectangle);
	/// \brief <em>Creates a legacy OLE drag image for the specified buttons</em>
	///
	/// Creates an OLE drag image for the specified buttons in the style of Windows versions prior to Vista.
	///
	/// \param[in] pButtons A \c IToolBarButtonContainer object wrapping the buttons for which to create the
	///            drag image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateVistaOLEDragImage, CreateLegacyDragImage, ToolBarButtonContainer,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateLegacyOLEDragImage(IToolBarButtonContainer* pButtons, __in LPSHDRAGIMAGE pDragImage);
	/// \brief <em>Creates a Vista-like OLE drag image for the specified buttons</em>
	///
	/// Creates an OLE drag image for the specified buttons in the style of Windows Vista and newer.
	///
	/// \param[in] pButtons A \c IToolBarButtonContainer object wrapping the buttons for which to create the
	///            drag image.
	/// \param[out] pDragImage Receives the drag image including transparency information and the coordinates
	///             (in pixels) of the drag image's upper-left corner relative to the control's upper-left
	///             corner.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa OnGetDragImage, CreateLegacyOLEDragImage, CreateLegacyDragImage, ToolBarButtonContainer,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb759778.aspx">SHDRAGIMAGE</a>
	BOOL CreateVistaOLEDragImage(IToolBarButtonContainer* pButtons, __in LPSHDRAGIMAGE pDragImage);
	//@}
	//////////////////////////////////////////////////////////////////////

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropTarget
	///
	//@{
	/// \brief <em>Indicates whether a drop can be accepted, and, if so, the effect of the drop</em>
	///
	/// This method is called by the \c DoDragDrop function to determine the target's preferred drop
	/// effect the first time the user moves the mouse into the control during OLE drag'n'Drop. The
	/// target communicates the \c DoDragDrop function the drop effect it wants to be used on drop.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the dragged
	///            data.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragOver, DragLeave, Drop, Raise_OLEDragEnter,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680106.aspx">IDropTarget::DragEnter</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnter(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Notifies the target that it no longer is the target of the current OLE drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse out of the
	/// control during OLE drag'n'Drop or if the user canceled the operation. The target must release
	/// any references it holds to the data object.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, Drop, Raise_OLEDragLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680110.aspx">IDropTarget::DragLeave</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeave(void);
	/// \brief <em>Communicates the current drop effect to the \c DoDragDrop function</em>
	///
	/// This method is called by the \c DoDragDrop function if the user moves the mouse over the
	/// control during OLE drag'n'Drop. The target communicates the \c DoDragDrop function the drop
	/// effect it wants to be used on drop.
	///
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, the current drop effect (defined by the \c DROPEFFECT
	///                enumeration). On return, this paramter must be set to the drop effect that the
	///                target wants to be used on drop.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragLeave, Drop, Raise_OLEDragMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680129.aspx">IDropTarget::DragOver</a>
	virtual HRESULT STDMETHODCALLTYPE DragOver(DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	/// \brief <em>Incorporates the source data into the target and completes the drag'n'drop operation</em>
	///
	/// This method is called by the \c DoDragDrop function if the user completes the drag'n'drop
	/// operation. The target must incorporate the dragged data into itself and pass the used drop
	/// effect back to the \c DoDragDrop function. The target must release any references it holds to
	/// the data object.
	///
	/// \param[in] pDataObject The \c IDataObject implementation of the object containing the data
	///            to transfer.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, this paramter must be set
	///                to the drop effect that the target finally executed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnter, DragOver, DragLeave, Raise_OLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687242.aspx">IDropTarget::Drop</a>
	virtual HRESULT STDMETHODCALLTYPE Drop(IDataObject* pDataObject, DWORD keyState, POINTL mousePosition, DWORD* pEffect);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSource
	///
	//@{
	/// \brief <em>Notifies the source of the current drop effect during OLE drag'n'drop</em>
	///
	/// This method is called frequently by the \c DoDragDrop function to notify the source of the
	/// last drop effect that the target has chosen. The source should set an appropriate mouse cursor.
	///
	/// \param[in] effect The drop effect chosen by the target. Any of the values defined by the
	///            \c DROPEFFECT enumeration is valid.
	///
	/// \return \c S_OK if the method has set a custom mouse cursor; \c DRAGDROP_S_USEDEFAULTCURSORS to
	///         use default mouse cursors; or an error code otherwise.
	///
	/// \sa QueryContinueDrag, Raise_OLEGiveFeedback,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693723.aspx">IDropSource::GiveFeedback</a>
	virtual HRESULT STDMETHODCALLTYPE GiveFeedback(DWORD effect);
	/// \brief <em>Determines whether a drag'n'drop operation should be continued, canceled or completed</em>
	///
	/// This method is called by the \c DoDragDrop function to determine whether a drag'n'drop
	/// operation should be continued, canceled or completed.
	///
	/// \param[in] pressedEscape Indicates whether the user has pressed the \c ESC key since the
	///            previous call of this method. If \c TRUE, the key has been pressed; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	///
	/// \return \c S_OK if the drag'n'drop operation should continue; \c DRAGDROP_S_DROP if it should
	///         be completed; \c DRAGDROP_S_CANCEL if it should be canceled; or an error code otherwise.
	///
	/// \sa GiveFeedback, Raise_OLEQueryContinueDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690076.aspx">IDropSource::QueryContinueDrag</a>
	virtual HRESULT STDMETHODCALLTYPE QueryContinueDrag(BOOL pressedEscape, DWORD keyState);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IDropSourceNotify
	///
	//@{
	/// \brief <em>Notifies the source that the user drags the mouse cursor into a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor into a potential drop target window.
	///
	/// \param[in] hWndTarget The potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragLeaveTarget, Raise_OLEDragEnterPotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragEnterTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragEnterTarget(HWND hWndTarget);
	/// \brief <em>Notifies the source that the user drags the mouse cursor out of a potential drop target window</em>
	///
	/// This method is called by the \c DoDragDrop function to notify the source that the user is dragging
	/// the mouse cursor out of a potential drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DragEnterTarget, Raise_OLEDragLeavePotentialTarget,
	///     <a href="https://msdn.microsoft.com/en-us/library/aa344050.aspx">IDropSourceNotify::DragLeaveTarget</a>
	virtual HRESULT STDMETHODCALLTYPE DragLeaveTarget(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICategorizeProperties
	///
	//@{
	/// \brief <em>Retrieves a category's name</em>
	///
	/// \param[in] category The ID of the category whose name is requested.
	// \param[in] languageID The locale identifier identifying the language in which name should be
	//            provided.
	/// \param[out] pName The category's name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::GetCategoryName
	virtual HRESULT STDMETHODCALLTYPE GetCategoryName(PROPCAT category, LCID /*languageID*/, BSTR* pName);
	/// \brief <em>Maps a property to a category</em>
	///
	/// \param[in] property The ID of the property whose category is requested.
	/// \param[out] pCategory The category's ID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ICategorizeProperties::MapPropertyToCategory
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToCategory(DISPID property, PROPCAT* pCategory);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ICreditsProvider
	///
	//@{
	/// \brief <em>Retrieves the name of the control's authors</em>
	///
	/// \return A string containing the names of all authors.
	CAtlString GetAuthors(void);
	/// \brief <em>Retrieves the URL of the website that has information about the control</em>
	///
	/// \return A string containing the URL.
	CAtlString GetHomepage(void);
	/// \brief <em>Retrieves the URL of the website where users can donate via Paypal</em>
	///
	/// \return A string containing the URL.
	CAtlString GetPaypalLink(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank especially</em>
	///
	/// \return A string containing the special thanks.
	CAtlString GetSpecialThanks(void);
	/// \brief <em>Retrieves persons, websites, organizations we want to thank</em>
	///
	/// \return A string containing the thanks.
	CAtlString GetThanks(void);
	/// \brief <em>Retrieves the control's version</em>
	///
	/// \return A string containing the version.
	CAtlString GetVersion(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IKeyboardHookHandler
	///
	//@{
	/// \brief <em>Processes a hooked keyboard message</em>
	///
	/// This method is called by the callback function that we defined when we installed a keyboard hook to
	/// trap accelerator key presses for the \c ToolBar control.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam The virtual key code of the pressed key.
	/// \param[in] lParam Details about the key press.
	/// \param[out] consumeMessage If set to \c TRUE, the message will not be passed to any further handlers
	///             and \c CallNextHookEx will not be called.
	///
	/// \return The value to return if consuming the message.
	///
	/// \sa SendConfigurationMessages, KeyboardHookProvider::KeyboardHookProvider, IKeyboardHookHandler,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644984.aspx">KeyboardProc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644974.aspx">CallNextHookEx</a>
	LRESULT HandleKeyboardMessage(int /*code*/, WPARAM wParam, LPARAM lParam, BOOL& consumeMessage);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IMouseHookHandler
	///
	//@{
	/// \brief <em>Processes a hooked mouse message</em>
	///
	/// This method is called by the callback function that we defined when we installed a mouse hook to trap
	/// mouse messages that should dismiss the \c ToolBar control's chevron popup window.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam The identifier of the mouse message.
	/// \param[in] lParam Points to a \c MOUSEHOOKSTRUCT structure that contains more information.
	/// \param[out] consumeMessage If set to \c TRUE, the message will not be passed to any further handlers
	///             and \c CallNextHookEx will not be called.
	///
	/// \return The value to return if consuming the message.
	///
	/// \sa DisplayChevronPopupWindow, MouseHookProvider::MouseHookProc, IMouseHookHandler,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644968.aspx">MOUSEHOOKSTRUCT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644988.aspx">MouseProc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644974.aspx">CallNextHookEx</a>
	LRESULT HandleMouseMessage(int /*code*/, WPARAM wParam, LPARAM lParam, BOOL& consumeMessage);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IPerPropertyBrowsing
	///
	//@{
	/// \brief <em>A displayable string for a property's current value is required</em>
	///
	/// This method is called if the caller's user interface requests a user-friendly description of the
	/// specified property's current value that may be displayed instead of the value itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display name is requested.
	/// \param[out] pDescription The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms688734.aspx">IPerPropertyBrowsing::GetDisplayString</a>
	virtual HRESULT STDMETHODCALLTYPE GetDisplayString(DISPID property, BSTR* pDescription);
	/// \brief <em>Displayable strings for a property's predefined values are required</em>
	///
	/// This method is called if the caller's user interface requests user-friendly descriptions of the
	/// specified property's predefined values that may be displayed instead of the values itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property whose display names are requested.
	/// \param[in,out] pDescriptions A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of strings. This array will be
	///                filled with the display name strings.
	/// \param[in,out] pCookies A caller-allocated, counted array structure containing the element
	///                count and address of a callee-allocated array of \c DWORD values. Each \c DWORD
	///                value identifies a predefined value of the property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedValue, GetDisplayStringForSetting,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms679724.aspx">IPerPropertyBrowsing::GetPredefinedStrings</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedStrings(DISPID property, CALPOLESTR* pDescriptions, CADWORD* pCookies);
	/// \brief <em>A property's predefined value identified by a token is required</em>
	///
	/// This method is called if the caller's user interface requires a property's predefined value that
	/// it has the token of. The token was returned by the \c GetPredefinedStrings method.
	/// We use this method for enumeration-type properties to transform strings like "1 - At Root"
	/// back to the underlying enumeration value (here: \c lsLinesAtRoot).
	///
	/// \param[in] property The ID of the property for which a predefined value is requested.
	/// \param[in] cookie Token identifying which value to return. The token was previously returned
	///            in the \c pCookies array filled by \c IPerPropertyBrowsing::GetPredefinedStrings.
	/// \param[out] pPropertyValue A \c VARIANT that will receive the predefined value.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetPredefinedStrings,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690401.aspx">IPerPropertyBrowsing::GetPredefinedValue</a>
	virtual HRESULT STDMETHODCALLTYPE GetPredefinedValue(DISPID property, DWORD cookie, VARIANT* pPropertyValue);
	/// \brief <em>A property's property page is required</em>
	///
	/// This method is called to request the \c CLSID of the property page used to edit the specified
	/// property.
	///
	/// \param[in] property The ID of the property whose property page is requested.
	/// \param[out] pPropertyPage The property page's \c CLSID.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms694476.aspx">IPerPropertyBrowsing::MapPropertyToPage</a>
	virtual HRESULT STDMETHODCALLTYPE MapPropertyToPage(DISPID property, CLSID* pPropertyPage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Retrieves a displayable string for a specified setting of a specified property</em>
	///
	/// Retrieves a user-friendly description of the specified property's specified setting. This
	/// description may be displayed by the caller's user interface instead of the setting itself.
	/// We use this method for enumeration-type properties to display strings like "1 - At Root" instead
	/// of "1 - lsLinesAtRoot".
	///
	/// \param[in] property The ID of the property for which to retrieve the display name.
	/// \param[in] cookie Token identifying the setting for which to retrieve a description.
	/// \param[out] description The setting's display name.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetDisplayString, GetPredefinedStrings, GetResStringWithNumber
	HRESULT GetDisplayStringForSetting(DISPID property, DWORD cookie, CComBSTR& description);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of ISpecifyPropertyPages
	///
	//@{
	/// \brief <em>The property pages to show are required</em>
	///
	/// This method is called if the property pages, that may be displayed for this object, are required.
	///
	/// \param[out] pPropertyPages A caller-allocated, counted array structure containing the element
	///             count and address of a callee-allocated array of \c GUID structures. Each \c GUID
	///             structure identifies a property page to display.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CommonProperties,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms687276.aspx">ISpecifyPropertyPages::GetPages</a>
	virtual HRESULT STDMETHODCALLTYPE GetPages(CAUUID* pPropertyPages);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler to raise the \c KeyPress event.
	///
	/// \sa OnKeyDown, Raise_KeyPress,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_COMMAND handler</em>
	///
	/// Will be called if a command has been selected from a menu.
	/// We use this handler for menu mode.
	///
	/// \sa get_MenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_COMMAND</a>
	LRESULT OnCommand(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnRButtonDown, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnContextMenu(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_CREATE handler</em>
	///
	/// Will be called right after the control was created.
	/// We use this handler to configure the control window and to raise the \c RecreatedControlWindow event.
	///
	/// \sa OnDestroy, OnFinalMessage, Raise_RecreatedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632619.aspx">WM_CREATE</a>
	LRESULT OnCreate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control is being destroyed.
	/// We use this handler to tidy up and to raise the \c DestroyedControlWindow event.
	///
	/// \sa OnCreate, OnFinalMessage, Raise_DestroyedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_FORWARDMSG handler</em>
	///
	/// Will be called if a message that was intended for another window, has been forwarded to the control.
	/// We use this handler to handle messages that have been forwarded by the message hook.
	///
	/// \sa GetMsgProc,
	///     <a href="https://msdn.microsoft.com/en-us/library/xfsxak9k.aspx">WM_FORWARDMSG</a>
	LRESULT OnForwardMsg(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_INITMENUPOPUP handler</em>
	///
	/// Will be called if a popup menu is about to become active.
	/// We use this handler to raise the \c OpenPopupMenu event.
	///
	/// \sa OnMenuSelect, get_MenuMode, Raise_OpenPopupMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646347.aspx">WM_INITMENUPOPUP</a>
	LRESULT OnInitMenuPopup(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyDown event.
	///
	/// \sa OnKeyUp, OnChar, Raise_KeyDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnKeyDown(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KEYUP handler</em>
	///
	/// Will be called if a nonsystem key is released while the control has the keyboard focus.
	/// We use this handler to raise the \c KeyUp event.
	///
	/// \sa OnKeyDown, OnChar, Raise_KeyUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646281.aspx">WM_KEYUP</a>
	LRESULT OnKeyUp(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_KILLFOCUS handler</em>
	///
	/// Will be called after the control lost the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnSetFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646282.aspx">WM_KILLFOCUS</a>
	LRESULT OnKillFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the left mouse
	/// button.
	/// We use this handler to raise the \c DblClick event.
	///
	/// \sa OnMButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_DblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645606.aspx">WM_LBUTTONDBLCLK</a>
	LRESULT OnLButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnClickNotification, OnMButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MENUCHAR handler</em>
	///
	/// Will be called when a menu is active and the user presses a key that does not correspond to any
	/// mnemonic or accelerator key.
	/// We use this handler to support keyboard navigation in menu mode.
	///
	/// \sa get_MenuMode, DoDropDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646349.aspx">WM_MENUCHAR</a>
	LRESULT OnMenuChar(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MEASUREITEM and \c WM_DRAWITEM handler</em>
	///
	/// Will be called when a owner-drawn menu item needs to be drawn or its size is required.
	/// We use this handler to raise the \c RawMenuMessage event.
	///
	/// \sa get_MenuMode, Raise_RawMenuMessage,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775923.aspx">WM_DRAWITEM</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775925.aspx">WM_MEASUREITEM</a>
	LRESULT OnMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MENUSELECT handler</em>
	///
	/// Will be called if the user selects a menu item in a menu owned by us.
	/// We use this handler to support keyboard navigation in menu mode and to raise the \c SelectedMenuItem
	/// event.
	///
	/// \sa OnInitMenuPopup, get_MenuMode, DoDropDown, Raise_SelectedMenuItem,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646352.aspx">WM_MENUSELECT</a>
	LRESULT OnMenuSelect(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEACTIVATE handler</em>
	///
	/// Will be called if the control is inactive and the user clicked in its client area.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnSetFocus, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645612.aspx">WM_MOUSEACTIVATE</a>
	LRESULT OnMouseActivate(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEHOVER handler</em>
	///
	/// Will be called if the mouse cursor has been located over the control's client area for a previously
	/// specified number of milliseconds.
	/// We use this handler to raise the \c MouseHover event.
	///
	/// \sa Raise_MouseHover,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645613.aspx">WM_MOUSEHOVER</a>
	LRESULT OnMouseHover(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NOTIFY handler</em>
	///
	/// Will be called if the control receives a notification from a child-window (probably the tooltip
	/// control).
	/// We use this handler to allow the user to cancel info tip popups.
	///
	/// \sa OnToolTipGetDispInfoNotificationW, OnChevronPopupToolbarNotify,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672614.aspx">WM_NOTIFY</a>
	LRESULT OnNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_PAINT handler</em>
	///
	/// Will be called if the control needs to be drawn.
	/// We use this handler to avoid the control being drawn by \c CComControl.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms534901.aspx">WM_PAINT</a>
	LRESULT OnPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_RBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the right mouse
	/// button.
	/// We use this handler to raise the \c RDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnXButtonDblClk, Raise_RDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646241.aspx">WM_RBUTTONDBLCLK</a>
	LRESULT OnRButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(LONG index, UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFOCUS handler</em>
	///
	/// Will be called after the control gained the keyboard focus.
	/// We use this handler to make VB's Validate event work.
	///
	/// \sa OnMouseActivate, OnKillFocus, Flags::uiActivationPending,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646283.aspx">WM_SETFOCUS</a>
	LRESULT OnSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETFONT handler</em>
	///
	/// Will be called if the control's font shall be set.
	/// We use this handler to synchronize our font settings with the new font.
	///
	/// \sa get_Font,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632642.aspx">WM_SETFONT</a>
	LRESULT OnSetFont(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETREDRAW handler</em>
	///
	/// Will be called if the control's redraw state shall be changed.
	/// We use this handler for proper handling of the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms534853.aspx">WM_SETREDRAW</a>
	LRESULT OnSetRedraw(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETTINGCHANGE handler</em>
	///
	/// Will be called if a system setting was changed.
	/// We use this handler to update our appearance.
	///
	/// \attention This message is posted to top-level windows only, so actually we'll never receive it.
	///
	/// \sa OnThemeChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms725497.aspx">WM_SETTINGCHANGE</a>
	LRESULT OnSettingChange(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_THEMECHANGED handler</em>
	///
	/// Will be called on themable systems if the theme was changed.
	/// We use this handler to update our appearance.
	///
	/// \sa OnSettingChange, Flags::usingThemes,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632650.aspx">WM_THEMECHANGED</a>
	LRESULT OnThemeChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_TIMER handler</em>
	///
	/// Will be called when a timer expires that's associated with the control.
	/// We use this handler for the \c DontRedraw property.
	///
	/// \sa get_DontRedraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644902.aspx">WM_TIMER</a>
	LRESULT OnTimer(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGED handler</em>
	///
	/// Will be called if the control was moved.
	/// We use this handler to resize the control on COM level and to raise the \c ResizedControlWindow
	/// event.
	///
	/// \sa Raise_ResizedControlWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using one of the extended
	/// mouse buttons.
	/// We use this handler to raise the \c XDblClick event.
	///
	/// \sa OnLButtonDblClk, OnMButtonDblClk, OnRButtonDblClk, Raise_XDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646244.aspx">WM_XBUTTONDBLCLK</a>
	LRESULT OnXButtonDblClk(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(LONG index, UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NOTIFY handler</em>
	///
	/// Will be called if the control's parent window receives a notification from the control.
	/// We use this handler for the custom draw support.
	///
	/// \sa OnCustomDrawNotification, OnReflectedNotifyFormat,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775583.aspx">WM_NOTIFY</a>
	LRESULT OnReflectedNotify(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NOTIFYFORMAT handler</em>
	///
	/// Will be called if the control asks its parent window which format (Unicode/ANSI) the \c WM_NOTIFY
	/// notifications should have.
	/// We use this handler for proper Unicode support.
	///
	/// \sa OnReflectedNotify,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775584.aspx">WM_NOTIFYFORMAT</a>
	LRESULT OnReflectedNotifyFormat(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c DI_GETDRAGIMAGE handler</em>
	///
	/// Will be called during OLE drag'n'drop if the control is queried for a drag image.
	///
	/// \sa OLEDrag, CreateLegacyOLEDragImage, CreateVistaOLEDragImage,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646443.aspx">DI_GETDRAGIMAGE</a>
	LRESULT OnGetDragImage(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c Bar_InternalGetInstanceMessage handler</em>
	///
	/// Will be called if the control is being asked for the handle of the associated \c ToolBar control.
	/// We use this handler in the message hook to retrieve the associated \c ToolBar window.
	///
	/// \sa GetMsgProc, GetBarMessage
	LRESULT OnGetToolbar(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c ToolBar_InternalAutoPopupMessage handler</em>
	///
	/// Will be called if the control is being asked to display the drop-down menu of a tool bar button in
	/// menu mode.
	/// We use this handler to support switching between sub-menus by moving the mouse.
	///
	/// \sa get_MenuMode, OnHookMouseMove, DoDropDown, GetAutoPopupMessage
	LRESULT OnAutoPopup(LONG index, UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c ReBar_InternalMDIChildWindowStateChangedMessage handler</em>
	///
	/// Will be called if the control is being notified that the active MDI child window has been maximized
	/// or restored.
	/// We use this handler to support switching between the MDI child window's system menu and the tool
	/// bar's drop-down menus.
	///
	/// \sa get_MenuMode, OnHookKeyDown, GetMDIChildWindowStateChangedMessage
	LRESULT OnMDIChildWindowStateChanged(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c ToolBar_DisplayChevronPopupMessage handler</em>
	///
	/// Will be called if the control is being asked to display the chevron popup window.
	/// We use this handler to support displaying the chevron popup window without help of the client
	/// application.
	///
	/// \sa DisplayChevronPopupWindow, ReBar::OnChevronPushedNotification, GetDisplayChevronPopupMessage
	LRESULT OnDisplayChevronPopup(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c ToolBar_DeferredSetButtonTextMessage handler</em>
	///
	/// Will be called if the control is being asked to raise the \c ButtonGetDisplayInfo event and set a
	/// button's text.
	/// We use this handler to work around a problematic implementation of the \c TBN_GETBUTTONINFOA
	/// notification in the native tool bar control.
	///
	/// \sa Raise_ButtonGetDisplayInfo, OnGetButtonInfoNotification, GetDeferredSetButtonTextMessage
	LRESULT OnDeferredSetButtonText(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_ADDBUTTONS handler</em>
	///
	/// Will be called if new tool bar buttons shall be inserted into the control.
	/// We use this handler mainly to raise the \c InsertingButton and \c InsertedButton events.
	///
	/// \sa OnInsertButton, OnDeleteButton, Raise_InsertingButton, Raise_InsertedButton,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787291.aspx">TB_ADDBUTTONS</a>
	LRESULT OnAddButtons(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_DELETEBUTTON handler</em>
	///
	/// Will be called if a tool bar button shall be removed from the control.
	/// We use this handler mainly to raise the \c RemovingButton and \c RemovedButton events.
	///
	/// \sa OnDeletingButtonNotification, OnAddButtons, OnInsertButton, Raise_RemovingButton,
	///     Raise_RemovedButton,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787309.aspx">TB_DELETEBUTTON</a>
	LRESULT OnDeleteButton(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_INSERTBUTTON handler</em>
	///
	/// Will be called if a new tool bar button shall be inserted into the control.
	/// We use this handler mainly to raise the \c InsertingButton and \c InsertedButton events.
	///
	/// \sa OnAddButtons, OnDeleteButton, Raise_InsertingButton, Raise_InsertedButton,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787364.aspx">TB_INSERTBUTTON</a>
	LRESULT OnInsertButton(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_SETBUTTONINFO handler</em>
	///
	/// Will be called if a button's settings shall be changed.
	/// We use this handler to synchronize our map of placeholder buttons and to raise the
	/// \c ButtonSelectionStateChanged event.
	///
	/// \sa placeholderButtons, OnSetCmdId, Raise_ButtonSelectionStateChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787413.aspx">TB_SETBUTTONINFO</a>
	LRESULT OnSetButtonInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_SETBUTTONWIDTH handler</em>
	///
	/// Will be called if the minimum and maximum width of buttons shall be changed.
	/// We use this handler to synchronize the \c MaximumButtonWidth and \c MinimumButtonWidth properties.
	///
	/// \sa get_MaximumButtonWidth, get_MinimumButtonWidth,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787417.aspx">TB_SETBUTTONWIDTH</a>
	LRESULT OnSetButtonWidth(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_SETCMDID handler</em>
	///
	/// Will be called if the unique identifier of a button is being changed.
	/// We use this handler to synchronize our map of placeholder buttons.
	///
	/// \sa placeholderButtons, OnSetButtonInfo,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787419.aspx">TB_SETCMDID</a>
	LRESULT OnSetCmdId(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_SETDRAWTEXTFLAGS handler</em>
	///
	/// Will be called if the flags used when calling the \c DrawText function shall be changed.
	/// We use this handler to buffer the settings, because there is no way to retrieve the settings from the
	/// control.
	///
	/// \sa Properties::drawTextFlags, Properties::drawTextFlagsMask,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787425.aspx">TB_SETDRAWTEXTFLAGS</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/dd162498.aspx">DrawText</a>
	LRESULT OnSetDrawTextFlags(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_SETDROPDOWNGAP handler</em>
	///
	/// Will be called if the spacing between buttons' text and drop-down section shall be changed.
	/// We use this handler to synchronize the \c DropDownGap property.
	///
	/// \sa OnSetListGap, OnSetIndent, get_DropDownGap,
	///     TB_SETDROPDOWNGAP
	LRESULT OnSetDropDownGap(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_SETINDENT handler</em>
	///
	/// Will be called if the indentation of the first button in a row shall be changed.
	/// We use this handler to synchronize the \c FirstButtonIndentation property.
	///
	/// \sa OnSetListGap, get_FirstButtonIndentation,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787435.aspx">TB_SETINDENT</a>
	LRESULT OnSetIndent(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TB_SETLISTGAP handler</em>
	///
	/// Will be called if the spacing between buttons' icon and text shall be changed.
	/// We use this handler to synchronize the \c HorizontalIconCaptionGap property.
	///
	/// \sa OnSetDropDownGap, OnSetIndent, get_HorizontalIconCaptionGap,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787441.aspx">TB_SETLISTGAP</a>
	LRESULT OnSetListGap(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);

	/// \brief <em>\c WM_ACTIVATE handler</em>
	///
	/// Will be called if the control's top-level parent window is being activated or deactivated.
	/// We use this handler in menu mode to redraw the control.
	///
	/// \sa get_MenuMode, get_MenuBarTheme,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646274.aspx">WM_ACTIVATE</a>
	LRESULT OnParentActivate(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SYSCOMMAND handler</em>
	///
	/// Will be called if the control's top-level parent window is asked to display its system menu.
	/// We use this handler in menu mode to show and hide keyboard cues.
	///
	/// \sa get_MenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646360.aspx">WM_SYSCOMMAND</a>
	LRESULT OnParentSysCommand(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);

	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler for menu mode.
	///
	/// \sa get_MenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnHookMouseMove(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SYSKEYDOWN handler</em>
	///
	/// Will be called if a system key is pressed while the control has the keyboard focus.
	/// We use this handler for menu mode.
	///
	/// \sa OnHookSysKeyUp, OnHookSysChar, OnHookKeyDown, get_MenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646286.aspx">WM_SYSKEYDOWN</a>
	LRESULT OnHookSysKeyDown(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SYSKEYUP handler</em>
	///
	/// Will be called if a system key is released while the control has the keyboard focus.
	/// We use this handler for menu mode.
	///
	/// \sa OnHookSysKeyDown, OnHookSysChar, get_MenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646287.aspx">WM_SYSKEYUP</a>
	LRESULT OnHookSysKeyUp(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_SYSCHAR handler</em>
	///
	/// Will be called if a \c WM_SYSKEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler for menu mode.
	///
	/// \sa OnHookSysKeyDown, OnHookSysKeyUp, get_MenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646357.aspx">WM_SYSCHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnHookSysChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_KEYDOWN handler</em>
	///
	/// Will be called if a nonsystem key is pressed while the control has the keyboard focus.
	/// We use this handler for menu mode.
	///
	/// \sa OnHookSysKeyDown, OnHookChar, get_MenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646280.aspx">WM_KEYDOWN</a>
	LRESULT OnHookKeyDown(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_NEXTMENU handler</em>
	///
	/// Will be called if the right or left arrow key is used to switch between the menu bar and the system
	/// menu.
	/// We use this handler for menu mode.
	///
	/// \sa get_MenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647612.aspx">WM_NEXTMENU</a>
	LRESULT OnHookNextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_CHAR handler</em>
	///
	/// Will be called if a \c WM_KEYDOWN message was translated by \c TranslateMessage.
	/// We use this handler for menu mode.
	///
	/// \sa OnHookKeyDown, get_MenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646276.aspx">WM_CHAR</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644955.aspx">TranslateMessage</a>
	LRESULT OnHookChar(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);

	/// \brief <em>\c WM_ACTIVATEAPP and \c WM_CANCELMODE handler</em>
	///
	/// Will be called if a \c WM_ACTIVATEAPP or \c WM_CANCELMODE message was sent to the
	/// \c parentWindowChevronPopupMenu window.
	/// We use this handler to dismiss the chevron popup window.
	///
	/// \sa DisplayChevronPopupWindow, parentWindowChevronPopupMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632614.aspx">WM_ACTIVATEAPP</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632615.aspx">WM_CANCELMODE</a>
	LRESULT OnChevronPopupMenuDismiss(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);

	/// \brief <em>\c WM_NOTIFY handler</em>
	///
	/// Will be called if the chevron popup tool bar control receives a notification from a child-window
	/// (probably the tooltip control).
	/// We use this handler to allow the user to cancel info tip popups.
	///
	/// \sa OnToolTipGetDispInfoNotificationW, OnNotify,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms672614.aspx">WM_NOTIFY</a>
	LRESULT OnChevronPopupToolbarNotify(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c TBN_BEGINADJUST handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user is beginning customizing the
	/// tool bar.
	/// We use this handler to raise the \c BeginCustomization event.
	///
	/// \sa OnEndAdjustNotification, OnInitCustomizeNotification, OnQueryInsertNotification,
	///     OnQueryDeleteNotification, Raise_BeginCustomization,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760509.aspx">TBN_BEGINADJUST</a>
	LRESULT OnBeginAdjustNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_BEGINDRAG handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user wants to drag a tool bar
	/// button.
	/// We use this handler to raise the \c ButtonBeginDrag and \c ButtonBeginRDrag events.
	///
	/// \sa Raise_ButtonBeginDrag, Raise_ButtonBeginRDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760512.aspx">TBN_BEGINDRAG</a>
	LRESULT OnBeginDragNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_CUSTHELP handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user clicked the help button in
	/// the control's customization dialog.
	/// We use this handler to raise the \c DisplayCustomizationHelp event.
	///
	/// \sa OnInitCustomizeNotification, Raise_DisplayCustomizationHelp,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760521.aspx">TBN_CUSTHELP</a>
	LRESULT OnCustHelpNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_DELETINGBUTTON handler</em>
	///
	/// Will be called if the control's parent window is notified, that a tool bar button is being removed.
	/// We use this handler to raise the \c FreeButtonData event.
	///
	/// \sa OnDeleteButton, Raise_FreeButtonData,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760523.aspx">TBN_DELETINGBUTTON</a>
	LRESULT OnDeletingButtonNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_DROPDOWN handler</em>
	///
	/// Will be called if the control's parent window is notified, that a drop-down button's drop-down menu
	/// should be displayed.
	/// We use this handler to raise the \c DropDown event.
	///
	/// \sa OnReflectedClicked, Raise_DropDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760527.aspx">TBN_DROPDOWN</a>
	LRESULT OnDropDownNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_DUPACCELERATOR handler</em>
	///
	/// Will be called if the control needs to know whether an accelerator character maps to more than one
	/// button.
	/// We use this handler to raise the \c IsDuplicateAccelerator event.
	///
	/// \sa OnMapAcceleratorNotification, OnWrapAcceleratorNotification, Raise_IsDuplicateAccelerator,
	///     <a href="https://msdn.microsoft.com/en-us/library/​cc835029.aspx">TBN_DUPACCELERATOR</a>
	LRESULT OnDupAcceleratorNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_ENDADJUST handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user has stopped customizing the
	/// tool bar.
	/// We use this handler to raise the \c EndCustomization event.
	///
	/// \sa OnBeginAdjustNotification, Raise_EndCustomization,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760529.aspx">TBN_ENDADJUST</a>
	LRESULT OnEndAdjustNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_GETBUTTONINFO handler</em>
	///
	/// Will be called if the control requests information about a tool bar button from its parent window.
	/// We use this handler to raise the \c GetAvailableButton event.
	///
	/// \sa Raise_GetAvailableButton,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787266.aspx">TBN_GETBUTTONINFO</a>
	LRESULT OnGetButtonInfoNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_GETDISPINFO handler</em>
	///
	/// Will be called if the control requests display information about a tool bar button from its parent
	/// window.
	/// We use this handler to raise the \c ButtonGetDisplayInfo event.
	///
	/// \sa Raise_ButtonGetDisplayInfo,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787268.aspx">TBN_GETDISPINFO</a>
	LRESULT OnGetDispInfoNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_GETINFOTIP handler</em>
	///
	/// Will be called if the control requests a tool bar button's tooltip text from its parent window.
	/// We use this handler to raise the \c ButtonGetInfoTipText event.
	///
	/// \sa Raise_ButtonGetInfoTipText,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787270.aspx">TBN_GETINFOTIP</a>
	LRESULT OnGetInfoTipNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_GETOBJECT handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user is dragging data over a
	/// button.
	/// We use this handler to make \c TBSTYLE_REGISTERDROP work together with our implementation of
	/// \c IDropTarget.
	///
	/// \sa put_RegisterForOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787272.aspx">TBN_GETOBJECT</a>
	LRESULT OnGetObjectNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_HOTITEMCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's hot button is being
	/// changed. We use this handler to raise the \c HotButtonChanging event.
	///
	/// \sa OnWrapHotItemNotification, Raise_HotButtonChanging,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787274.aspx">TBN_HOTITEMCHANGE</a>
	LRESULT OnHotItemChangeNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& wasHandled);
	/// \brief <em>\c TBN_INITCUSTOMIZE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's customization dialog
	/// is being initialized. We use this handler to raise the \c InitializeCustomizationDialog event.
	///
	/// \sa OnBeginAdjustNotification, OnCustHelpNotification, Raise_InitializeCustomizationDialog,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787276.aspx">TBN_INITCUSTOMIZE</a>
	LRESULT OnInitCustomizeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_MAPACCELERATOR handler</em>
	///
	/// Will be called if the control needs to know the button to which an accelerator character corresponds.
	/// We use this handler to raise the \c MapAccelerator event.
	///
	/// \sa OnDupAcceleratorNotification, OnWrapAcceleratorNotification, Raise_MapAccelerator,
	///     <a href="https://msdn.microsoft.com/en-us/library/cc835030.aspx">TBN_MAPACCELERATOR</a>
	LRESULT OnMapAcceleratorNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_QUERYDELETE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's customization dialog
	/// needs to know whether a button shall be removable by the user. We use this handler to raise the
	/// \c QueryRemoveButton event.
	///
	/// \sa Customize, Raise_QueryRemoveButton,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787278.aspx">TBN_QUERYDELETE</a>
	LRESULT OnQueryDeleteNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_QUERYINSERT handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's customization dialog
	/// needs to know whether a button can be inserted to the left of another button. We use this handler to
	/// raise the \c QueryInsertButton event.
	///
	/// \sa Customize, Raise_QueryInsertButton,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787280.aspx">TBN_QUERYINSERT</a>
	LRESULT OnQueryInsertNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_RESET handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user clicked the reset button in
	/// the control's customization dialog.
	/// We use this handler to raise the \c ResetCustomizations event.
	///
	/// \sa Customize, Raise_ResetCustomizations,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787281.aspx">TBN_RESET</a>
	LRESULT OnResetNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_RESTORE handler</em>
	///
	/// Will be called if the control's parent window is notified, that a button's state is being restored
	/// from the registry.
	/// We use this handler to raise the \c InitializeToolBarStateRegistryRestorage and
	/// \c RestoreButtonFromRegistryStream events.
	///
	/// \sa OnSaveNotification, Raise_InitializeToolBarStateRegistryRestorage,
	///     Raise_RestoreButtonFromRegistryStream,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787283.aspx">TBN_RESTORE</a>
	LRESULT OnRestoreNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_SAVE handler</em>
	///
	/// Will be called if the control's parent window is notified, that a button's state is being written
	/// to the registry.
	/// We use this handler to raise the \c InitializeToolBarStateRegistryStorage and
	/// \c SaveButtonToRegistryStream events.
	///
	/// \sa OnRestoreNotification, Raise_InitializeToolBarStateRegistryStorage,
	///     Raise_SaveButtonToRegistryStream,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787285.aspx">TBN_SAVE</a>
	LRESULT OnSaveNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_TOOLBARCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the tool bar has been customized.
	/// We use this handler to raise the \c CustomizedControl event.
	///
	/// \sa Raise_CustomizedControl,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787287.aspx">TBN_TOOLBARCHANGE</a>
	LRESULT OnToolBarChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_WRAPACCELERATOR handler</em>
	///
	/// Will be called if the control tries to map an accelerator character to a tool bar button and has
	/// reached the last button so that it will continue with the first button. The client application can
	/// return its own mapping.
	/// We use this handler to raise the \c MapAccelerator event.
	///
	/// \sa OnDupAcceleratorNotification, OnMapAcceleratorNotification, Raise_MapAccelerator,
	///     <a href="https://msdn.microsoft.com/en-us/library/​cc835031.aspx">TBN_WRAPACCELERATOR</a>
	LRESULT OnWrapAcceleratorNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c TBN_WRAPHOTITEM handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's hot button will be
	/// changed from the last button to the first or from the first button to the last.
	/// We use this handler to raise the \c HotButtonChangeWrapping event.
	///
	/// \sa OnHotItemChangeNotification, Raise_HotButtonChangeWrapping,
	///     <a href="https://msdn.microsoft.com/en-us/library/cc835032.aspx">TBN_WRAPHOTITEM</a>
	LRESULT OnWrapHotItemNotification(LONG index, int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	/// \brief <em>\c BN_CLICKED handler</em>
	///
	/// Will be called if the control's parent window is notified that a tool bar button's associated command
	/// shall be executed.
	/// We use this handler to raise the \c ExecuteCommand event.
	///
	/// \sa OnDropDownNotification, Raise_ExecuteCommand,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms673572.aspx">BN_CLICKED</a>
	LRESULT OnReflectedClicked(LONG index, WORD /*notifyCode*/, WORD controlID, HWND /*hWnd*/, BOOL& wasHandled);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>\c NM_CUSTOMDRAW handler</em>
	///
	/// Will be called by the \c OnReflectedNotify method if a custom draw notification was received for
	/// the tool bar.
	/// We use this handler to raise the \c CustomDraw event.
	///
	/// \sa OnReflectedNotify, Raise_CustomDraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760492.aspx">NM_CUSTOMDRAW (tool bar)</a>
	LRESULT OnCustomDrawNotification(LONG index, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOA handler</em>
	///
	/// Will be called by the \c OnNotify and the \c OnChevronPopupToolbarNotify method if the
	/// \c TTN_GETDISPINFO notification was received.
	/// We use this handler to allow the user to cancel info tip popups.
	///
	/// \sa OnToolTipGetDispInfoNotificationW, OnNotify, OnChevronPopupToolbarNotify,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760269.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationA(BOOL chevronPopup, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c TTN_GETDISPINFOW handler</em>
	///
	/// Will be called by the \c OnNotify and the \c OnChevronPopupToolbarNotify method if the
	/// \c TTN_GETDISPINFO notification was received.
	/// We use this handler to allow the user to cancel info tip popups.
	///
	/// \sa OnToolTipGetDispInfoNotificationA, OnNotify, OnChevronPopupToolbarNotify,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb760269.aspx">TTN_GETDISPINFO</a>
	LRESULT OnToolTipGetDispInfoNotificationW(BOOL chevronPopup, UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c BeforeDisplayChevronPopup event</em>
	///
	/// \param[in] hPopup The handle of popup. In menu mode, this is the handle of the chevron popup
	///            menu; otherwise this is the window handle of the popup tool bar.
	/// \param[in] x The x-coordinate (in twips) of the menu's proposed position relative to the
	///            screen's upper-left corner.
	/// \param[in] y The y-coordinate (in twips) of the menu's proposed position relative to the
	///            screen's upper-left corner.
	/// \param[in] isMenu If \c VARIANT_TRUE, the handle specified by \c hPopup is a menu handle; otherwise
	///            it is a window handle.
	/// \param[out] pCancelPopup If set to \c VARIANT_TRUE, the popup won't be displayed.
	/// \param[out] pCommandToExecute If the popup is canceled, this parameter can be set to the unique ID
	///             of a command that will be executed. If set to \c 0, no command will be executed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_BeforeDisplayChevronPopup,
	///       TBarCtlsLibU::_IToolBarEvents::BeforeDisplayChevronPopup, Raise_DestroyingChevronPopup,
	///       get_DisplayChevronPopupWindow, get_MenuMode
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_BeforeDisplayChevronPopup,
	///       TBarCtlsLibA::_IToolBarEvents::BeforeDisplayChevronPopup, Raise_DestroyingChevronPopup,
	///       get_DisplayChevronPopupWindow, get_MenuMode
	/// \endif
	inline HRESULT Raise_BeforeDisplayChevronPopup(LONG hPopup, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL isMenu, VARIANT_BOOL* pCancelPopup, LONG* pCommandToExecute);
	/// \brief <em>Raises the \c BeginCustomization event</em>
	///
	/// \param[in] restoringFromRegistry If \c VARIANT_TRUE, the event is fired, because the control's state
	///            is being restored from the registry.
	/// \param[in,out] pDontRestoreFromRegistry If set to \c VARIANT_TRUE, the caller should ignore the data
	///                stored in the registry when restoring the control's state; otherwise it should use
	///                this data. This parameter should be ignored if \c restoringFromRegistry is set to
	///                \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_BeginCustomization,
	///       TBarCtlsLibU::_IToolBarEvents::BeginCustomization, Raise_InitializeCustomizationDialog,
	///       Raise_EndCustomization, Raise_QueryInsertButton, Raise_QueryRemoveButton,
	///       Raise_InitializeToolBarStateRegistryRestorage, get_AllowCustomization, Customize,
	///       LoadToolBarStateFromRegistry
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_BeginCustomization,
	///       TBarCtlsLibA::_IToolBarEvents::BeginCustomization, Raise_InitializeCustomizationDialog,
	///       Raise_EndCustomization, Raise_QueryInsertButton, Raise_QueryRemoveButton,
	///       Raise_InitializeToolBarStateRegistryRestorage, get_AllowCustomization, Customize,
	///       LoadToolBarStateFromRegistry
	/// \endif
	inline HRESULT Raise_BeginCustomization(VARIANT_BOOL restoringFromRegistry, VARIANT_BOOL* pDontRestoreFromRegistry);
	/// \brief <em>Raises the \c ButtonBeginDrag event</em>
	///
	/// \param[in] pToolButton The button that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbLeftButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ButtonBeginDrag, TBarCtlsLibU::_IToolBarEvents::ButtonBeginDrag,
	///       OLEDrag, Raise_ButtonBeginRDrag, TBarCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ButtonBeginDrag, TBarCtlsLibA::_IToolBarEvents::ButtonBeginDrag,
	///       OLEDrag, Raise_ButtonBeginRDrag, TBarCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ButtonBeginDrag(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c ButtonBeginRDrag event</em>
	///
	/// \param[in] pToolButton The button that the user wants to drag.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid, but usually it should be just
	///            \c vbRightButton.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies in.
	///            Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ButtonBeginRDrag, TBarCtlsLibU::_IToolBarEvents::ButtonBeginRDrag,
	///       OLEDrag, Raise_ButtonBeginDrag, TBarCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ButtonBeginRDrag, TBarCtlsLibA::_IToolBarEvents::ButtonBeginRDrag,
	///       OLEDrag, Raise_ButtonBeginDrag, TBarCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ButtonBeginRDrag(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c ButtonGetDisplayInfo event</em>
	///
	/// \param[in] pToolButton The button whose properties are requested.
	/// \param[in] requestedInfo Specifies which properties' values are required. Any combination of
	///            the values defined by the \c RequestedInfoConstants enumeration is valid.
	/// \param[out] pIconIndex The zero-based index of the button's icon. The icon is taken from the
	///             image lists specified by the \c pImageListIndex parameter. If the \c requestedInfo
	///             parameter doesn't include \c riIconIndex, the caller should ignore this value.
	/// \param[out] pImageListIndex The zero-based index of the control's image lists that the specified icon
	///             image lists specified by the \c pImageListIndex parameter. If the \c requestedInfo
	///             is taken from. The icon is taken from the parameter doesn't include \c riIconIndex, the
	///             caller should ignore this value.
	/// \param[in] maxButtonTextLength The maximum number of characters the button's or text may consist of.
	///            If the \c requestedInfo parameter doesn't include \c riButtonText, the client should
	///            ignore this value.
	/// \param[out] pButtonText The button's text. If the \c requestedInfo parameter doesn't include
	///             \c riButtonText, the caller should ignore this value.
	/// \param[in,out] pDontAskAgain If \c VARIANT_TRUE, the caller should always use the same settings and
	///                never fire this event again for these properties of this button; otherwise it
	///                shouldn't make the values persistent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ButtonGetDisplayInfo,
	///       TBarCtlsLibU::_IToolBarEvents::ButtonGetDisplayInfo, ToolBarButton::put_IconIndex,
	///       ToolBarButton::put_ImageListIndex, put_hImageList, ToolBarButton::put_Text,
	///       TBarCtlsLibU::RequestedInfoConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ButtonGetDisplayInfo,
	///       TBarCtlsLibA::_IToolBarEvents::ButtonGetDisplayInfo, ToolBarButton::put_IconIndex,
	///       ToolBarButton::put_ImageListIndex, put_hImageList, ToolBarButton::put_Text,
	///       TBarCtlsLibA::RequestedInfoConstants
	/// \endif
	inline HRESULT Raise_ButtonGetDisplayInfo(IToolBarButton* pToolButton, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pImageListIndex, LONG maxButtonTextLength, BSTR* pButtonText, VARIANT_BOOL* pDontAskAgain);
	/// \brief <em>Raises the \c ButtonGetDropDownMenu event</em>
	///
	/// \param[in] pToolButton The button that the menu is required for.
	/// \param[out] phMenu Receives the handle of the popup menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ButtonGetDropDownMenu,
	///       TBarCtlsLibU::_IToolBarEvents::ButtonGetDropDownMenu, get_MenuMode, Raise_DropDown
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ButtonGetDropDownMenu,
	///       TBarCtlsLibA::_IToolBarEvents::ButtonGetDropDownMenu, get_MenuMode, Raise_DropDown
	/// \endif
	inline HRESULT Raise_ButtonGetDropDownMenu(IToolBarButton* pToolButton, LONG* phMenu);
	/// \brief <em>Raises the \c ButtonGetInfoTipText event</em>
	///
	/// \param[in] pToolButton The button whose info tip text is requested.
	/// \param[in] maxInfoTipLength The maximum number of characters the info tip text may consist of.
	/// \param[out] pInfoTipText The item's info tip text.
	/// \param[in,out] pAbortToolTip If \c VARIANT_TRUE, the caller should NOT show the tooltip;
	///                otherwise it should.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ButtonGetInfoTipText,
	///       TBarCtlsLibU::_IToolBarEvents::ButtonGetInfoTipText, get_ShowToolTips, ToolBarButton::get_Text
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ButtonGetInfoTipText,
	///       TBarCtlsLibA::_IToolBarEvents::ButtonGetInfoTipText, get_ShowToolTips, ToolBarButton::get_Text
	/// \endif
	inline HRESULT Raise_ButtonGetInfoTipText(IToolBarButton* pToolButton, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip);
	/// \brief <em>Raises the \c ButtonMouseEnter event</em>
	///
	/// \param[in] pToolButton The tool bar button that was entered.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ButtonMouseEnter, TBarCtlsLibU::_IToolBarEvents::ButtonMouseEnter,
	///       Raise_ButtonMouseLeave, Raise_MouseMove, TBarCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ButtonMouseEnter, TBarCtlsLibA::_IToolBarEvents::ButtonMouseEnter,
	///       Raise_ButtonMouseLeave, Raise_MouseMove, TBarCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ButtonMouseEnter(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c ButtonMouseLeave event</em>
	///
	/// \param[in] pToolButton The tool bar button that was left.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Most of the values defined by the \c HitTestConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ButtonMouseLeave, TBarCtlsLibU::_IToolBarEvents::ButtonMouseLeave,
	///       Raise_ButtonMouseEnter, Raise_MouseMove, TBarCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ButtonMouseLeave, TBarCtlsLibA::_IToolBarEvents::ButtonMouseLeave,
	///       Raise_ButtonMouseEnter, Raise_MouseMove, TBarCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_ButtonMouseLeave(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c ButtonSelectionStateChanged event</em>
	///
	/// \param[in] pToolButton The tool bar button for which the selection state was changed.
	/// \param[in] previousSelectionState The tool bar button's previous selection state.
	/// \param[in] newSelectionState The tool bar button's new selection state.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ButtonSelectionStateChanged,
	///       TBarCtlsLibU::_IToolBarEvents::ButtonSelectionStateChanged, Raise_ExecuteCommand,
	///       TBarCtlsLibU::SelectionStateConstants, ToolBarButton::get_SelectionState
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ButtonSelectionStateChanged,
	///       TBarCtlsLibA::_IToolBarEvents::ButtonSelectionStateChanged, Raise_ExecuteCommand,
	///       TBarCtlsLibA::SelectionStateConstants, ToolBarButton::get_SelectionState
	/// \endif
	inline HRESULT Raise_ButtonSelectionStateChanged(IToolBarButton* pToolButton, SelectionStateConstants previousSelectionState, SelectionStateConstants newSelectionState);
	/// \brief <em>Raises the \c Click event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_Click, TBarCtlsLibU::_IToolBarEvents::Click, Raise_DblClick,
	///       Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_Click, TBarCtlsLibA::_IToolBarEvents::Click, Raise_DblClick,
	///       Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_Click(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c ContextMenu event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ContextMenu, TBarCtlsLibU::_IToolBarEvents::ContextMenu,
	///       Raise_RClick, get_ProcessContextMenuKeys
	///      
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ContextMenu, TBarCtlsLibA::_IToolBarEvents::ContextMenu,
	///       Raise_RClick, get_ProcessContextMenuKeys
	/// \endif
	inline HRESULT Raise_ContextMenu(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CustomDraw event</em>
	///
	/// \param[in] pToolButton The tool bar button that the notification refers to. May be \c NULL.
	/// \param[in,out] pNormalTextColor An \c OLE_COLOR value specifying the color to draw the button's text
	///                in if the button is neither in hot nor in marked state. The client may change this
	///                value.
	/// \param[in,out] pNormalButtonBackColor An \c OLE_COLOR value specifying the color to fill the button's
	///                background with if the button is neither in hot nor in marked state. The client may
	///                change this value.
	/// \param[in,out] pNormalBackgroundMode Specifies how to draw the background of the button's text if the
	///                button is neither in hot nor in marked state. Any of the values defined by the
	///                \c StringBackgroundModeConstants enumeration are valid.
	/// \param[in,out] pHotTextColor An \c OLE_COLOR value specifying the color to draw the button's text
	///                in if the button is in hot state. The client may change this value.
	/// \param[in,out] pHotButtonBackColor An \c OLE_COLOR value specifying the color to fill the button's
	///                background with if the button is in hot state. The client may change this value.
	/// \param[in,out] pMarkedTextBackColor An \c OLE_COLOR value specifying the color to draw the button's
	///                text background in if the button is in marked state. The client may change this value.
	/// \param[in,out] pMarkedButtonBackColor An \c OLE_COLOR value specifying the color to fill the button's
	///                background with if the button is in marked state. The client may change this value.
	/// \param[in,out] pMarkedBackgroundMode Specifies how to draw the background of the button's text if the
	///                button is in marked state. Any of the values defined by the
	///                \c StringBackgroundModeConstants enumeration are valid.
	/// \param[in] drawStage The stage of custom drawing this event is raised for. Most of the values
	///            defined by the \c CustomDrawStageConstants enumeration are valid.
	/// \param[in] buttonState The tool bar button's current state (focused, selected etc.). Most of the
	///            values defined by the \c CustomDrawItemStateConstants enumeration are valid.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	/// \param[in,out] pTextRectangle The bounding rectangle of the button's text. The \c Right and \c Bottom
	///                members may be changed by the client application.
	/// \param[in,out] pHorizontalIconCaptionGap Specifies the number of pixels in horizontal direction
	///                between the button's icon and its text.
	/// \param[in,out] pFurtherProcessing A value controlling further drawing. Most of the values defined
	///                by the \c CustomDrawReturnValuesConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c horizontalIconCaptionGap parameter is ignored, if the \c ButtonTextPosition
	///          property is set to a value other than \c btpRightToIcon.\n
	///          The \c horizontalIconCaptionGap parameter requires comctl32.dll version 6.0 or higher.\n
	///          This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_CustomDraw, TBarCtlsLibU::_IToolBarEvents::CustomDraw,
	///       TBarCtlsLibU::RECTANGLE, TBarCtlsLibU::StringBackgroundModeConstants,
	///       TBarCtlsLibU::CustomDrawStageConstants, TBarCtlsLibU::CustomDrawItemStateConstants,
	///       TBarCtlsLibU::CustomDrawReturnValuesConstants, get_HotButton, ToolBarButton::get_Marked,
	///       get_ButtonTextPosition, get_DisabledEvents,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb760492.aspx">NM_CUSTOMDRAW (tool bar)</a>
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_CustomDraw, TBarCtlsLibA::_IToolBarEvents::CustomDraw,
	///       TBarCtlsLibA::RECTANGLE, TBarCtlsLibA::StringBackgroundModeConstants,
	///       TBarCtlsLibA::CustomDrawStageConstants, TBarCtlsLibA::CustomDrawItemStateConstants,
	///       TBarCtlsLibA::CustomDrawReturnValuesConstants, get_HotButton, ToolBarButton::get_Marked,
	///       get_ButtonTextPosition, get_DisabledEvents,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb760492.aspx">NM_CUSTOMDRAW (tool bar)</a>
	/// \endif
	inline HRESULT Raise_CustomDraw(IToolBarButton* pToolButton, OLE_COLOR* pNormalTextColor, OLE_COLOR* pNormalButtonBackColor, StringBackgroundModeConstants* pNormalBackgroundMode, OLE_COLOR* pHotTextColor, OLE_COLOR* pHotButtonBackColor, OLE_COLOR* pMarkedTextBackColor, OLE_COLOR* pMarkedButtonBackColor, StringBackgroundModeConstants* pMarkedBackgroundMode, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants buttonState, LONG hDC, RECTANGLE* pDrawingRectangle, RECTANGLE* pTextRectangle, OLE_XSIZE_PIXELS* pHorizontalIconCaptionGap, CustomDrawReturnValuesConstants* pFurtherProcessing);
	/// \brief <em>Raises the \c CustomizedControl event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_CustomizedControl, TBarCtlsLibU::_IToolBarEvents::CustomizedControl,
	///       Raise_BeginCustomization, Raise_EndCustomization, get_AllowCustomization, Customize
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_CustomizedControl, TBarCtlsLibA::_IToolBarEvents::CustomizedControl,
	///       Raise_BeginCustomization, Raise_EndCustomization, get_AllowCustomization, Customize
	/// \endif
	inline HRESULT Raise_CustomizedControl(void);
	/// \brief <em>Raises the \c CustomizeDialogRemovingButton event</em>
	///
	/// \param[in] pToolButton The tool bar button that is being removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_CustomizeDialogRemovingButton,
	///       TBarCtlsLibU::_IToolBarEvents::CustomizeDialogRemovingButton, Raise_FreeButtonData,
	///       Raise_QueryRemoveButton, Customize, ToolBarButtons::Remove
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_CustomizeDialogRemovingButton,
	///       TBarCtlsLibA::_IToolBarEvents::CustomizeDialogRemovingButton, Raise_FreeButtonData,
	///       Raise_QueryRemoveButton, Customize, ToolBarButtons::Remove
	/// \endif
	inline HRESULT Raise_CustomizeDialogRemovingButton(IToolBarButton* pToolButton);
	/// \brief <em>Raises the \c DblClick event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_DblClick, TBarCtlsLibU::_IToolBarEvents::DblClick, Raise_Click,
	///       Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_DblClick, TBarCtlsLibA::_IToolBarEvents::DblClick, Raise_Click,
	///       Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_DblClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_DestroyedControlWindow,
	///       TBarCtlsLibU::_IToolBarEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_DestroyedControlWindow,
	///       TBarCtlsLibA::_IToolBarEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c DestroyingChevronPopup event</em>
	///
	/// \param[in] hPopup The handle of popup. In menu mode, this is the handle of the chevron popup
	///            menu; otherwise this is the window handle of the popup tool bar.
	/// \param[in] isMenu If \c VARIANT_TRUE, the handle specified by \c hPopup is a menu handle; otherwise
	///            it is a window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_DestroyingChevronPopup,
	///       TBarCtlsLibU::_IToolBarEvents::DestroyingChevronPopup, Raise_BeforeDisplayChevronPopup,
	///       get_DisplayChevronPopupWindow, get_MenuMode
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_DestroyingChevronPopup,
	///       TBarCtlsLibA::_IToolBarEvents::DestroyingChevronPopup, Raise_BeforeDisplayChevronPopup,
	///       get_DisplayChevronPopupWindow, get_MenuMode
	/// \endif
	inline HRESULT Raise_DestroyingChevronPopup(LONG hPopup, VARIANT_BOOL isMenu);
	/// \brief <em>Raises the \c DisplayCustomizationHelp event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_DisplayCustomizationHelp,
	///       TBarCtlsLibU::_IToolBarEvents::DisplayCustomizationHelp, Raise_InitializeCustomizationDialog,
	///       get_AllowCustomization, Customize
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_DisplayCustomizationHelp,
	///       TBarCtlsLibA::_IToolBarEvents::DisplayCustomizationHelp, Raise_InitializeCustomizationDialog,
	///       get_AllowCustomization, Customize
	/// \endif
	inline HRESULT Raise_DisplayCustomizationHelp(void);
	/// \brief <em>Raises the \c DropDown event</em>
	///
	/// \param[in] pToolButton The tool bar button whose drop-down menu shall be displayed.
	/// \param[in] pButtonRectangle A \c RECTANGLE structure specifying the tool bar button's bounding
	///            rectangle in coordinates relative to the control's upper-left corner.
	/// \param[in,out] pFurtherProcessing A value controlling further processing of the user input. Most of
	///                the values defined by the \c DropDownReturnValuesConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_DropDown, TBarCtlsLibU::_IToolBarEvents::DropDown,
	///       Raise_ExecuteCommand, Raise_Click, TBarCtlsLibU::RECTANGLE,
	///       TBarCtlsLibU::DropDownReturnValuesConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_DropDown, TBarCtlsLibA::_IToolBarEvents::DropDown,
	///       Raise_ExecuteCommand, Raise_Click, TBarCtlsLibA::RECTANGLE,
	///       TBarCtlsLibA::DropDownReturnValuesConstants
	/// \endif
	inline HRESULT Raise_DropDown(IToolBarButton* pToolButton, RECTANGLE* pButtonRectangle, DropDownReturnValuesConstants* pFurtherProcessing);
	/// \brief <em>Raises the \c EndCustomization event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_EndCustomization, TBarCtlsLibU::_IToolBarEvents::EndCustomization,
	///       Raise_BeginCustomization, Raise_CustomizedControl, get_AllowCustomization, Customize
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_EndCustomization, TBarCtlsLibA::_IToolBarEvents::EndCustomization,
	///       Raise_BeginCustomization, Raise_CustomizedControl, get_AllowCustomization, Customize
	/// \endif
	inline HRESULT Raise_EndCustomization(void);
	/// \brief <em>Raises the \c ExecuteCommand event</em>
	///
	/// \param[in] commandID The unique ID of the command that shall be executed.
	/// \param[in] pToolButton The tool bar button whose associated command shall be executed. May be
	///            \c NULL.
	/// \param[in] commandOrigin Specifies whether the command was triggered by a button, a menu item or a
	///            hotkey. Any of the values defined by the \c CommandOriginConstants enumeration is valid.
	/// \param[in,out] pForwardMessage If set to \c VARIANT_TRUE, the command will be forwarded to the parent
	///                top-level window; otherwise not.
	/// \param[out] pCommandIsDisabled If set to \c TRUE, the event has not been fired, because the command
	///             is disabled.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ExecuteCommand, TBarCtlsLibU::_IToolBarEvents::ExecuteCommand,
	///       Raise_ButtonSelectionStateChanged, Raise_Click, Raise_DropDown,
	///       TBarCtlsLibU::CommandOriginConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ExecuteCommand, TBarCtlsLibA::_IToolBarEvents::ExecuteCommand,
	///       Raise_ButtonSelectionStateChanged, Raise_Click, Raise_DropDown,
	///       TBarCtlsLibA::CommandOriginConstants
	/// \endif
	inline HRESULT Raise_ExecuteCommand(LONG commandID, IToolBarButton* pToolButton, CommandOriginConstants commandOrigin, VARIANT_BOOL* pForwardMessage, __in_opt LPBOOL pCommandIsDisabled);
	/// \brief <em>Raises the \c FreeButtonData event</em>
	///
	/// \param[in] pToolButton The tool bar button whose associated data shall be freed. If \c NULL, all
	///            buttons' associated data shall be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_FreeButtonData, TBarCtlsLibU::_IToolBarEvents::FreeButtonData,
	///       Raise_RemovingButton, Raise_RemovedButton, ToolBarButton::put_ButtonData
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_FreeButtonData, TBarCtlsLibA::_IToolBarEvents::FreeButtonData,
	///       Raise_RemovingButton, Raise_RemovedButton, ToolBarButton::put_ButtonData
	/// \endif
	inline HRESULT Raise_FreeButtonData(IToolBarButton* pToolButton);
	/// \brief <em>Raises the \c GetAvailableButton event</em>
	///
	/// \param[in] pToolButton The tool bar button hat is being requested. The application needs to set the
	///            object's properties.
	/// \param[in,out] pNoMoreButtons If \c VARIANT_TRUE, the button specified by \c pToolButton is the last
	///                button and the event won't be fired for any further buttons; otherwise the event will
	///                be fired once more to request the next button.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_GetAvailableButton,
	///       TBarCtlsLibU::_IToolBarEvents::GetAvailableButton, Raise_QueryInsertButton,
	///       Raise_BeginCustomization, Raise_EndCustomization, get_AllowCustomization, Customize,
	///       VirtualToolBarButton
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_GetAvailableButton,
	///       TBarCtlsLibA::_IToolBarEvents::GetAvailableButton, Raise_QueryInsertButton,
	///       Raise_BeginCustomization, Raise_EndCustomization, get_AllowCustomization, Customize,
	///       VirtualToolBarButton
	/// \endif
	inline HRESULT Raise_GetAvailableButton(IVirtualToolBarButton* pToolButton, VARIANT_BOOL* pNoMoreButtons);
	/// \brief <em>Raises the \c HotButtonChangeWrapping event</em>
	///
	/// \param[in] pPreviousHotButton The previous hot button. May be \ NULL.
	/// \param[in] wrappingDirection The direction into which the hot button will be changed. Any of the
	///            values defined by the \c WrappingDirectionConstants enumeration is valid.
	/// \param[in] causedBy The reason for the hot button change. Any combination of the values defined by
	///            the \c HotButtonChangingCausedByConstants enumeration is valid.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort the hot button change; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_HotButtonChangeWrapping, TBarCtlsLibU::_IToolBarEvents::HotButtonChangeWrapping,
	///       get_HotButton, Raise_HotButtonChanging, TBarCtlsLibU::HotButtonChangingCausedByConstants,
	///       TBarCtlsLibU::WrappingDirectionConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_HotButtonChangeWrapping, TBarCtlsLibA::_IToolBarEvents::HotButtonChangeWrapping,
	///       get_HotButton, Raise_HotButtonChanging, TBarCtlsLibA::HotButtonChangingCausedByConstants,
	///       TBarCtlsLibA::WrappingDirectionConstants
	/// \endif
	inline HRESULT Raise_HotButtonChangeWrapping(IToolBarButton* pPreviousHotButton, WrappingDirectionConstants wrappingDirection, HotButtonChangingCausedByConstants causedBy, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c HotButtonChanging event</em>
	///
	/// \param[in] pPreviousHotButton The previous hot button. May be \ NULL.
	/// \param[in] pNewHotButton The new hot button. May be \ NULL.
	/// \param[in] causedBy The reason for the hot button change. Any combination of the values defined by
	///            the \c HotButtonChangingCausedByConstants enumeration is valid.
	/// \param[in] additionalInfo Additional information over the hot button change. Any combination of the
	///            values defined by the \c HotButtonChangingAdditionalInfoConstants enumeration is valid.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort the hot button change; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_HotButtonChanging, TBarCtlsLibU::_IToolBarEvents::HotButtonChanging,
	///       get_HotButton, Raise_HotButtonChangeWrapping, TBarCtlsLibU::HotButtonChangingCausedByConstants,
	///       TBarCtlsLibU::HotButtonChangingAdditionalInfoConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_HotButtonChanging, TBarCtlsLibA::_IToolBarEvents::HotButtonChanging,
	///       get_HotButton, Raise_HotButtonChangeWrapping, TBarCtlsLibA::HotButtonChangingCausedByConstants,
	///       TBarCtlsLibA::HotButtonChangingAdditionalInfoConstants
	/// \endif
	inline HRESULT Raise_HotButtonChanging(IToolBarButton* pPreviousHotButton, IToolBarButton* pNewHotButton, HotButtonChangingCausedByConstants causedBy, HotButtonChangingAdditionalInfoConstants additionalInfo, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c InitializeCustomizationDialog event</em>
	///
	/// \param[in] hCustomizationDialog The window handle of the customization dialog.
	/// \param[in,out] pDisplayHelpButton If \c VARIANT_TRUE, the customization dialog shall be displayed
	///                with a help button; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_InitializeCustomizationDialog,
	///       TBarCtlsLibU::_IToolBarEvents::InitializeCustomizationDialog, Raise_BeginCustomization,
	///       Raise_EndCustomization, Raise_DisplayCustomizationHelp, get_AllowCustomization, Customize
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_InitializeCustomizationDialog,
	///       TBarCtlsLibA::_IToolBarEvents::InitializeCustomizationDialog, Raise_BeginCustomization,
	///       Raise_EndCustomization, Raise_DisplayCustomizationHelp, get_AllowCustomization, Customize
	/// \endif
	inline HRESULT Raise_InitializeCustomizationDialog(LONG hCustomizationDialog, VARIANT_BOOL* pDisplayHelpButton);
	/// \brief <em>Raises the \c InitializeToolBarStateRegistryRestorage event</em>
	///
	/// \param[in,out] pNumberOfButtonsToLoad The number of tool bar buttons that will be restored from the
	///                data stream. The client application must ensure that this parameter is set to the
	///                correct value.
	/// \param[in] totalDataSize The total size of the data stream in bytes.
	/// \param[in] perButtonDataSize The number of bytes that hold the control-defined part of a single
	///            button's state. The stream may contain additional application-defined data.
	/// \param[in] pDataStream The memory address where the tool bar state and the application-defined data
	///            are loaded from.
	/// \param[in,out] ppStartOfNextDataBlock The (absolute) memory address of the first byte within the data
	///                stream that contains application-defined data. When leaving the event handler,
	///                \c ppStartOfNextDataBlock must point to the first byte after the application-defined
	///                data.
	/// \param[in,out] pCancelLoading If set to \c VARIANT_TRUE, restoring the tool bar state is aborted;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If an application does not need to load additional data when restoring the tool bar state,
	///          it does not need to handle this event.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_InitializeToolBarStateRegistryRestorage,
	///       TBarCtlsLibU::_IToolBarEvents::InitializeToolBarStateRegistryRestorage,
	///       Raise_InitializeToolBarStateRegistryStorage, Raise_RestoreButtonFromRegistryStream,
	///       LoadToolBarStateFromRegistry
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_InitializeToolBarStateRegistryRestorage,
	///       TBarCtlsLibA::_IToolBarEvents::InitializeToolBarStateRegistryRestorage,
	///       Raise_InitializeToolBarStateRegistryStorage, Raise_RestoreButtonFromRegistryStream,
	///       LoadToolBarStateFromRegistry
	/// \endif
	inline HRESULT Raise_InitializeToolBarStateRegistryRestorage(LONG* pNumberOfButtonsToLoad, LONG totalDataSize, LONG perButtonDataSize, LONG pDataStream, LONG* ppStartOfNextDataBlock, VARIANT_BOOL* pCancelLoading);
	/// \brief <em>Raises the \c InitializeToolBarStateRegistryStorage event</em>
	///
	/// \param[in,out] pNumberOfButtonsToSave The number of tool bar buttons that will be saved into the data
	///                stream. The client application must ensure that this parameter is set to the correct
	///                value.
	/// \param[in,out] pTotalDataSize The total size of the data stream in bytes. The control sets this
	///                parameter to the number of bytes that it needs to store the tool bar state. The client
	///                application must add the number of bytes that it needs to store additional data.
	/// \param[in,out] ppDataStream The memory address where the tool bar state and the application-defined
	///                data will be written to. This memory must be allocated by the client application and
	///                it must be large enough to hold the number of bytes specified by \c pTotalDataSize.
	/// \param[in,out] ppStartOfUnusedSpace The (absolute) memory address of the first unused byte within the
	///                allocated data stream. After allocating the memory, the application can write data
	///                into it. When leaving the event handler, \c ppStartOfUnusedSpace must point to the
	///                first byte that the control can write to without overwriting the application-defined
	///                data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If an application does not need to store additional data when storing the tool bar state, it
	///          does not need to handle this event.\n
	///          Memory allocated by the client application must be freed by the client application. This
	///          should be done after the \c SaveToolBarStateToRegistry method returned.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_InitializeToolBarStateRegistryStorage,
	///       TBarCtlsLibU::_IToolBarEvents::InitializeToolBarStateRegistryStorage,
	///       Raise_InitializeToolBarStateRegistryRestorage, Raise_SaveButtonToRegistryStream,
	///       SaveToolBarStateToRegistry
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_InitializeToolBarStateRegistryStorage,
	///       TBarCtlsLibA::_IToolBarEvents::InitializeToolBarStateRegistryStorage,
	///       Raise_InitializeToolBarStateRegistryRestorage, Raise_SaveButtonToRegistryStream,
	///       SaveToolBarStateToRegistry
	/// \endif
	inline HRESULT Raise_InitializeToolBarStateRegistryStorage(LONG* pNumberOfButtonsToSave, LONG* pTotalDataSize, LONG* ppDataStream, LONG* ppStartOfUnusedSpace);
	/// \brief <em>Raises the \c InsertedButton event</em>
	///
	/// \param[in] pToolButton The inserted tool bar button.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_InsertedButton, TBarCtlsLibU::_IToolBarEvents::InsertedButton,
	///       Raise_InsertingButton, Raise_RemovedButton
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_InsertedButton, TBarCtlsLibA::_IToolBarEvents::InsertedButton,
	///       Raise_InsertingButton, Raise_RemovedButton
	/// \endif
	inline HRESULT Raise_InsertedButton(IToolBarButton* pToolButton);
	/// \brief <em>Raises the \c InsertingButton event</em>
	///
	/// \param[in] pToolButton The tool bar button being inserted.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort insertion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_InsertingButton, TBarCtlsLibU::_IToolBarEvents::InsertingButton,
	///       Raise_InsertedButton, Raise_RemovingButton
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_InsertingButton, TBarCtlsLibA::_IToolBarEvents::InsertingButton,
	///       Raise_InsertedButton, Raise_RemovingButton
	/// \endif
	inline HRESULT Raise_InsertingButton(IVirtualToolBarButton* pToolButton, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c IsDuplicateAccelerator event</em>
	///
	/// \param[in] accelerator The accelerator character to map.
	/// \param[out] pIsDuplicate If set to \c VARIANT_TRUE, the accelerator character is mapped to more than
	///             one tool bar button; otherwise not.
	/// \param[out] pHandledEvent If set to \c VARIANT_TRUE, the control should use the value of the
	///             \c pIsDuplicate parameter. Otherwise the control should try to detect duplicate mappings
	///             itself.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_IsDuplicateAccelerator,
	///       TBarCtlsLibU::_IToolBarEvents::IsDuplicateAccelerator, Raise_MapAccelerator
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_IsDuplicateAccelerator,
	///       TBarCtlsLibA::_IToolBarEvents::IsDuplicateAccelerator, Raise_MapAccelerator
	/// \endif
	inline HRESULT Raise_IsDuplicateAccelerator(SHORT accelerator, VARIANT_BOOL* pIsDuplicate, VARIANT_BOOL* pHandledEvent);
	/// \brief <em>Raises the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_KeyDown, TBarCtlsLibU::_IToolBarEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_KeyDown, TBarCtlsLibA::_IToolBarEvents::KeyDown, Raise_KeyUp,
	///       Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyDown(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_KeyPress, TBarCtlsLibU::_IToolBarEvents::KeyPress, Raise_KeyDown,
	///       Raise_KeyUp
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_KeyPress, TBarCtlsLibA::_IToolBarEvents::KeyPress, Raise_KeyDown,
	///       Raise_KeyUp
	/// \endif
	inline HRESULT Raise_KeyPress(SHORT* pKeyAscii);
	/// \brief <em>Raises the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_KeyUp, TBarCtlsLibU::_IToolBarEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_KeyUp, TBarCtlsLibA::_IToolBarEvents::KeyUp, Raise_KeyDown,
	///       Raise_KeyPress
	/// \endif
	inline HRESULT Raise_KeyUp(SHORT* pKeyCode, SHORT shift);
	/// \brief <em>Raises the \c MapAccelerator event</em>
	///
	/// \param[in] accelerator The accelerator character to map.
	/// \param[in] pStartingPointOfSearch The button after which the search shall be started. Usually this is
	///            the current hot button. The \c ppMatchingButton parameter must be set to a button that has
	///            a higher index than this button.
	/// \param[in] resumingSearchWithFirstButton If set to \c VARIANT_TRUE, the event is raised during
	///            automatic mapping of the accelerator character and notifies the client application that
	///            the last button has been reached and the search will continue with the first button. The
	///            client application has another chance to provide a custom mapping.
	/// \param[out] ppMatchingButton Must be set to the button to which the accelerator was mapped. This can
	///             also be \c NULL.
	/// \param[out] pHandledEvent If set to \c VARIANT_TRUE, the control should use the button specified by
	///             the \c ppMatchingButton parameter. The client application must set this parameter to
	///             \c VARIANT_TRUE, if it handles this event, even if it sets \c ppMatchingButton to
	///             \c NULL. Otherwise the control should try to find a mapping itself.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_MapAccelerator, TBarCtlsLibU::_IToolBarEvents::MapAccelerator,
	///       Raise_KeyDown, Raise_KeyUp, Raise_KeyPress, Raise_ExecuteCommand, Raise_IsDuplicateAccelerator
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_MapAccelerator, TBarCtlsLibA::_IToolBarEvents::MapAccelerator,
	///       Raise_KeyDown, Raise_KeyUp, Raise_KeyPress, Raise_ExecuteCommand, Raise_IsDuplicateAccelerator
	/// \endif
	inline HRESULT Raise_MapAccelerator(SHORT accelerator, IToolBarButton* pStartingPointOfSearch, VARIANT_BOOL resumingSearchWithFirstButton, IToolBarButton** ppMatchingButton, VARIANT_BOOL* pHandledEvent);
	/// \brief <em>Raises the \c MClick event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_MClick, TBarCtlsLibU::_IToolBarEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_MClick, TBarCtlsLibA::_IToolBarEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MDblClick event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_MDblClick, TBarCtlsLibU::_IToolBarEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_MDblClick, TBarCtlsLibA::_IToolBarEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_MDblClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_MouseDown, TBarCtlsLibU::_IToolBarEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_MouseDown, TBarCtlsLibA::_IToolBarEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseDown(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_MouseEnter, TBarCtlsLibU::_IToolBarEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_ButtonMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_MouseEnter, TBarCtlsLibA::_IToolBarEvents::MouseEnter,
	///       Raise_MouseLeave, Raise_ButtonMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseEnter(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseHover event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_MouseHover, TBarCtlsLibU::_IToolBarEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_MouseHover, TBarCtlsLibA::_IToolBarEvents::MouseHover,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \endif
	inline HRESULT Raise_MouseHover(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseLeave event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_MouseLeave, TBarCtlsLibU::_IToolBarEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_ButtonMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_MouseLeave, TBarCtlsLibA::_IToolBarEvents::MouseLeave,
	///       Raise_MouseEnter, Raise_ButtonMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseLeave(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_MouseMove, TBarCtlsLibU::_IToolBarEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_MouseMove, TBarCtlsLibA::_IToolBarEvents::MouseMove,
	///       Raise_MouseEnter, Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp
	/// \endif
	inline HRESULT Raise_MouseMove(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseUp event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_MouseUp, TBarCtlsLibU::_IToolBarEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_MouseUp, TBarCtlsLibA::_IToolBarEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseUp(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLECompleteDrag, TBarCtlsLibU::_IToolBarEvents::OLECompleteDrag,
	///       Raise_OLEStartDrag, TBarCtlsLibU::IOLEDataObject::GetData, OLEDrag
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLECompleteDrag, TBarCtlsLibA::_IToolBarEvents::OLECompleteDrag,
	///       Raise_OLEStartDrag, TBarCtlsLibA::IOLEDataObject::GetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect);
	/// \brief <em>Raises the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client finally executed.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragDrop, TBarCtlsLibU::_IToolBarEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragDrop, TBarCtlsLibA::_IToolBarEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data. If \c NULL, the cached data object is used. We use this when
	///            we call this method from other places than \c DragEnter.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragEnter, TBarCtlsLibU::_IToolBarEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragEnter, TBarCtlsLibA::_IToolBarEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The potential drop target window's handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragEnterPotentialTarget,
	///       TBarCtlsLibU::_IToolBarEvents::OLEDragEnterPotentialTarget, Raise_OLEDragLeavePotentialTarget,
	///       OLEDrag
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragEnterPotentialTarget,
	///       TBarCtlsLibA::_IToolBarEvents::OLEDragEnterPotentialTarget, Raise_OLEDragLeavePotentialTarget,
	///       OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragLeave, TBarCtlsLibU::_IToolBarEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragLeave, TBarCtlsLibA::_IToolBarEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(void);
	/// \brief <em>Raises the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragLeavePotentialTarget,
	///       TBarCtlsLibU::_IToolBarEvents::OLEDragLeavePotentialTarget, Raise_OLEDragEnterPotentialTarget,
	///       OLEDrag
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragLeavePotentialTarget,
	///       TBarCtlsLibA::_IToolBarEvents::OLEDragLeavePotentialTarget, Raise_OLEDragEnterPotentialTarget,
	///       OLEDrag
	/// \endif
	inline HRESULT Raise_OLEDragLeavePotentialTarget(void);
	/// \brief <em>Raises the \c OLEDragMouseMove event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragMouseMove, TBarCtlsLibU::_IToolBarEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEDragMouseMove, TBarCtlsLibA::_IToolBarEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition);
	/// \brief <em>Raises the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c DROPEFFECT enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEGiveFeedback, TBarCtlsLibU::_IToolBarEvents::OLEGiveFeedback,
	///       Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEGiveFeedback, TBarCtlsLibA::_IToolBarEvents::OLEGiveFeedback,
	///       Raise_OLEQueryContinueDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEGiveFeedback(DWORD effect, VARIANT_BOOL* pUseDefaultCursors);
	/// \brief <em>Raises the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c TRUE, the user has pressed the \c ESC key since the last time
	///            this event was raised; otherwise not.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue (\c S_OK), cancel
	///                (\c DRAGDROP_S_CANCEL) or complete (\c DRAGDROP_S_DROP) the drag'n'drop operation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEQueryContinueDrag,
	///       TBarCtlsLibU::_IToolBarEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEQueryContinueDrag,
	///       TBarCtlsLibA::_IToolBarEvents::OLEQueryContinueDrag, Raise_OLEGiveFeedback
	/// \endif
	inline HRESULT Raise_OLEQueryContinueDrag(BOOL pressedEscape, DWORD keyState, HRESULT* pActionToContinueWith);
	/// \brief <em>Raises the \c OLEReceivedNewData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the data object has received data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEReceivedNewData,
	///       TBarCtlsLibU::_IToolBarEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEReceivedNewData,
	///       TBarCtlsLibA::_IToolBarEvents::OLEReceivedNewData, Raise_OLESetData,
	///       SourceOLEDataObject::GetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLESetData event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] formatID An integer value specifying the format the drop target is requesting data for.
	///            Valid values are those defined by VB's \c ClipBoardConstants enumeration, but also any
	///            other format that has been registered using the \c RegisterClipboardFormat API function.
	/// \param[in] index An integer value that is assigned to the internal \c FORMATETC struct's \c lindex
	///            member. Usually it is -1, but some formats like \c CFSTR_FILECONTENTS require multiple
	///            \c FORMATETC structs for the same format. In such cases each struct of this format will
	///            have a separate index.
	/// \param[in] dataOrViewAspect An integer value that is assigned to the internal \c FORMATETC struct's
	///            \c dwAspect member. Any of the \c DVASPECT_* values defined by the Microsoft&reg;
	///            Windows&reg; SDK are valid. The default is \c DVASPECT_CONTENT.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLESetData, TBarCtlsLibU::_IToolBarEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLESetData, TBarCtlsLibA::_IToolBarEvents::OLESetData,
	///       Raise_OLEStartDrag, SourceOLEDataObject::SetData, OLEDrag,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms649049.aspx">RegisterClipboardFormat</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms682177.aspx">FORMATETC</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776902.aspx#CFSTR_FILECONTENTS">CFSTR_FILECONTENTS</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms690318.aspx">DVASPECT</a>
	/// \endif
	HRESULT Raise_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect);
	/// \brief <em>Raises the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OLEStartDrag, TBarCtlsLibU::_IToolBarEvents::OLEStartDrag,
	///       Raise_OLESetData, Raise_OLECompleteDrag, SourceOLEDataObject::SetData, OLEDrag
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OLEStartDrag, TBarCtlsLibA::_IToolBarEvents::OLEStartDrag,
	///       Raise_OLESetData, Raise_OLECompleteDrag, SourceOLEDataObject::SetData, OLEDrag
	/// \endif
	inline HRESULT Raise_OLEStartDrag(IOLEDataObject* pData);
	/// \brief <em>Raises the \c OpenPopupMenu event</em>
	///
	/// \param[in] hMenu The handle of the menu that is about to be displayed.
	/// \param[in] parentMenuItemIndex The zero-based index of the parent menu item within its containing
	///            menu.
	/// \param[in] isSystemMenu If \c VARIANT_TRUE, \c hMenu is a window's system menu; otherwise it is a
	///            normal menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_OpenPopupMenu, TBarCtlsLibU::_IToolBarEvents::OpenPopupMenu,
	///       Raise_SelectedMenuItem, Raise_ButtonGetDropDownMenu, get_MenuMode
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_OpenPopupMenu, TBarCtlsLibA::_IToolBarEvents::OpenPopupMenu,
	///       Raise_SelectedMenuItem, Raise_ButtonGetDropDownMenu, get_MenuMode
	/// \endif
	inline HRESULT Raise_OpenPopupMenu(LONG hMenu, LONG parentMenuItemIndex, VARIANT_BOOL isSystemMenu);
	/// \brief <em>Raises the \c QueryInsertButton event</em>
	///
	/// \param[in] pToolButton The tool bar button for which the control needs to know whether it shall
	///            allow inserting a new button to the left of this button.
	/// \param[in,out] pAllowInsertionToLeft If set to \c VARIANT_TRUE, the customization dialog allows
	///                inserting a button to the left of the specified button; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_QueryInsertButton, TBarCtlsLibU::_IToolBarEvents::QueryInsertButton,
	///       Raise_QueryRemoveButton, Raise_BeginCustomization, Raise_EndCustomization,
	///       Raise_GetAvailableButton, get_AllowCustomization, Customize
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_QueryInsertButton, TBarCtlsLibA::_IToolBarEvents::QueryInsertButton,
	///       Raise_QueryRemoveButton, Raise_BeginCustomization, Raise_EndCustomization,
	///       Raise_GetAvailableButton, get_AllowCustomization, Customize
	/// \endif
	inline HRESULT Raise_QueryInsertButton(IToolBarButton* pToolButton, VARIANT_BOOL* pAllowInsertionToLeft);
	/// \brief <em>Raises the \c QueryRemoveButton event</em>
	///
	/// \param[in] pToolButton The tool bar button for which the control needs to know whether it may be
	///            removed by the user.
	/// \param[in,out] pAllowRemoval If set to \c VARIANT_TRUE, the customization dialog allows removing the
	///                specified button; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_QueryRemoveButton, TBarCtlsLibU::_IToolBarEvents::QueryRemoveButton,
	///       Raise_QueryInsertButton, Raise_CustomizeDialogRemovingButton, Raise_BeginCustomization,
	///       Raise_EndCustomization, get_AllowCustomization, Customize
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_QueryRemoveButton, TBarCtlsLibA::_IToolBarEvents::QueryRemoveButton,
	///       Raise_QueryInsertButton, Raise_CustomizeDialogRemovingButton, Raise_BeginCustomization,
	///       Raise_EndCustomization, get_AllowCustomization, Customize
	/// \endif
	inline HRESULT Raise_QueryRemoveButton(IToolBarButton* pToolButton, VARIANT_BOOL* pAllowRemoval);
	/// \brief <em>Raises the \c RawMenuMessage event</em>
	///
	/// \param[in] message The message to handle. The event is fired for the following messages:
	///            - \c WM_DRAWITEM
	///            - \c WM_INITMENUPOPUP
	///            - \c WM_MEASUREITEM
	///            - \c WM_MENUCHAR
	///            - \c WM_MENUSELECT
	///            - \c WM_NEXTMENU
	/// \param[in] wParam The message's \c wParam parameter.
	/// \param[in] lParam The message's \c lParam parameter.
	/// \param[out] pResult Receives the value to return for this message if the client application handles
	///             the message.
	/// \param[out] pHandledEvent If set to \c VARIANT_TRUE, the client application has handled the message;
	///             otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_RawMenuMessage, TBarCtlsLibU::_IToolBarEvents::RawMenuMessage,
	///       get_MenuMode, Raise_SelectedMenuItem, Raise_OpenPopupMenu, Raise_ButtonGetDropDownMenu,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776085.aspx">IContextMenu3::HandleMenuMsg2</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775923.aspx">WM_DRAWITEM</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646347.aspx">WM_INITMENUPOPUP</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775925.aspx">WM_MEASUREITEM</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646349.aspx">WM_MENUCHAR</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646352.aspx">WM_MENUSELECT</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647612.aspx">WM_NEXTMENU</a>
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_RawMenuMessage, TBarCtlsLibA::_IToolBarEvents::RawMenuMessage,
	///       get_MenuMode, Raise_SelectedMenuItem, Raise_OpenPopupMenu, Raise_ButtonGetDropDownMenu,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb776085.aspx">IContextMenu3::HandleMenuMsg2</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775923.aspx">WM_DRAWITEM</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646347.aspx">WM_INITMENUPOPUP</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775925.aspx">WM_MEASUREITEM</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646349.aspx">WM_MENUCHAR</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646352.aspx">WM_MENUSELECT</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647612.aspx">WM_NEXTMENU</a>
	/// \endif
	inline HRESULT Raise_RawMenuMessage(LONG message, LONG wParam, LONG lParam, LONG* pResult, VARIANT_BOOL* pHandledEvent);
	/// \brief <em>Raises the \c RClick event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_RClick, TBarCtlsLibU::_IToolBarEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_RClick, TBarCtlsLibA::_IToolBarEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_RClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RDblClick event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_RDblClick, TBarCtlsLibU::_IToolBarEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_RDblClick, TBarCtlsLibA::_IToolBarEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_RDblClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_RecreatedControlWindow,
	///       TBarCtlsLibU::_IToolBarEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_RecreatedControlWindow,
	///       TBarCtlsLibA::_IToolBarEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c RemovedButton event</em>
	///
	/// \param[in] pToolButton The removed button. If \c NULL, all buttons were removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_RemovedButton, TBarCtlsLibU::_IToolBarEvents::RemovedButton,
	///       Raise_RemovingButton, Raise_InsertedButton
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_RemovedButton, TBarCtlsLibA::_IToolBarEvents::RemovedButton,
	///       Raise_RemovingButton, Raise_InsertedButton
	/// \endif
	inline HRESULT Raise_RemovedButton(IVirtualToolBarButton* pToolButton);
	/// \brief <em>Raises the \c RemovingButton event</em>
	///
	/// \param[in] pToolButton The button being removed. If \c NULL, all buttons are removed.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort deletion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_RemovingButton, TBarCtlsLibU::_IToolBarEvents::RemovingButton,
	///       Raise_RemovedButton, Raise_InsertingButton
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_RemovingButton, TBarCtlsLibA::_IToolBarEvents::RemovingButton,
	///       Raise_RemovedButton, Raise_InsertingButton
	/// \endif
	inline HRESULT Raise_RemovingButton(IToolBarButton* pToolButton, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ResetCustomizations event</em>
	///
	/// \param[in] hCustomizationDialog The window handle of the customization dialog.
	/// \param[in,out] pEndCustomization If \c VARIANT_TRUE, the caller should end customization; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ResetCustomizations,
	///       TBarCtlsLibU::_IToolBarEvents::ResetCustomizations, Raise_BeginCustomization,
	///       Raise_EndCustomization, Customize
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ResetCustomizations,
	///       TBarCtlsLibA::_IToolBarEvents::ResetCustomizations, Raise_BeginCustomization,
	///       Raise_EndCustomization, Customize
	/// \endif
	inline HRESULT Raise_ResetCustomizations(LONG hCustomizationDialog, VARIANT_BOOL* pEndCustomization);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_ResizedControlWindow,
	///       TBarCtlsLibU::_IToolBarEvents::ResizedControlWindow
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_ResizedControlWindow,
	///       TBarCtlsLibA::_IToolBarEvents::ResizedControlWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c RestoreButtonFromRegistryStream event</em>
	///
	/// \param[in] pToolButton The tool bar button that is being restored.
	/// \param[in] numberOfButtonsToLoad The number of tool bar buttons that is being restored from the data
	///            stream.
	/// \param[in] totalDataSize The total size of the data stream in bytes.
	/// \param[in] perButtonDataSize The number of bytes that hold the control-defined part of a single
	///            button's state. The stream may contain additional application-defined data.
	/// \param[in] pDataStream The memory address where the tool bar state and the application-defined data
	///            are loaded from.
	/// \param[in,out] ppStartOfNextDataBlock The (absolute) memory address of the first byte within the data
	///                stream that contains application-defined data for the specified button. When leaving
	///                the event handler, \c ppStartOfNextDataBlock must point to the first byte after this
	///                application-defined data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If an application does not need to load additional data when restoring the tool bar state,
	///          it does not need to handle this event.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_RestoreButtonFromRegistryStream,
	///       TBarCtlsLibU::_IToolBarEvents::RestoreButtonFromRegistryStream,
	///       Raise_SaveButtonToRegistryStream, Raise_InitializeToolBarStateRegistryRestorage,
	///       LoadToolBarStateFromRegistry
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_RestoreButtonFromRegistryStream,
	///       TBarCtlsLibA::_IToolBarEvents::RestoreButtonFromRegistryStream,
	///       Raise_SaveButtonToRegistryStream, Raise_InitializeToolBarStateRegistryRestorage,
	///       LoadToolBarStateFromRegistry
	/// \endif
	inline HRESULT Raise_RestoreButtonFromRegistryStream(IVirtualToolBarButton* pToolButton, LONG numberOfButtonsToLoad, LONG totalDataSize, LONG perButtonDataSize, LONG pDataStream, LONG* ppStartOfNextDataBlock);
	/// \brief <em>Raises the \c SaveButtonToRegistryStream event</em>
	///
	/// \param[in] pToolButton The tool bar button that needs to be stored.
	/// \param[in] totalDataSize The total size of the data stream in bytes.
	/// \param[in] pDataStream The memory address where the tool bar state and the application-defined data
	///            is written to.
	/// \param[in,out] ppStartOfUnusedSpace The (absolute) memory address of the first unused byte within
	///                the data stream. The application can write data to this address. When leaving the
	///                event handler, \c ppStartOfUnusedSpace must point to the first byte that the control
	///                can write to without overwriting the application-defined data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If an application does not need to store additional data when storing the tool bar state, it
	///          does not need to handle this event.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_SaveButtonToRegistryStream,
	///       TBarCtlsLibU::_IToolBarEvents::SaveButtonToRegistryStream,
	///       Raise_RestoreButtonFromRegistryStream, Raise_InitializeToolBarStateRegistryStorage,
	///       SaveToolBarStateToRegistry
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_SaveButtonToRegistryStream,
	///       TBarCtlsLibA::_IToolBarEvents::SaveButtonToRegistryStream,
	///       Raise_RestoreButtonFromRegistryStream, Raise_InitializeToolBarStateRegistryStorage,
	///       SaveToolBarStateToRegistry
	/// \endif
	inline HRESULT Raise_SaveButtonToRegistryStream(IVirtualToolBarButton* pToolButton, LONG totalDataSize, LONG pDataStream, LONG* ppStartOfUnusedSpace);
	/// \brief <em>Raises the \c SelectedMenuItem event</em>
	///
	/// \param[in] commandIDOrSubMenuIndex The unique ID of the selected menu command. If the menu item
	///            opens a sub-menu, this parameter contains the zero-based index of the menu item within
	///            its containing menu. If the menu is closed, this parameter will be 0.
	/// \param[in] menuItemState A bit-field describing the state of the menu item. Any combination of
	///            the values defined by the \c MenuItemStateConstants enumeration is valid. If the menu
	///            is closed, this parameter will be 0xFFFF.
	/// \param[in] hMenu The handle of the menu that contains the selected menu item. If the menu is closed,
	///            this parameter will be \c NULL.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_SelectedMenuItem, TBarCtlsLibU::_IToolBarEvents::SelectedMenuItem,
	///       Raise_ButtonGetDropDownMenu, Raise_OpenPopupMenu, get_MenuMode,
	///       TBarCtlsLibU::MenuItemStateConstants
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_SelectedMenuItem, TBarCtlsLibA::_IToolBarEvents::SelectedMenuItem,
	///       Raise_ButtonGetDropDownMenu, Raise_OpenPopupMenu, get_MenuMode,
	///       TBarCtlsLibA::MenuItemStateConstants
	/// \endif
	inline HRESULT Raise_SelectedMenuItem(LONG commandIDOrSubMenuIndex, MenuItemStateConstants menuItemState, LONG hMenu);
	/// \brief <em>Raises the \c XClick event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_XClick, TBarCtlsLibU::_IToolBarEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_XClick, TBarCtlsLibA::_IToolBarEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick
	/// \endif
	inline HRESULT Raise_XClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c XDblClick event</em>
	///
	/// \param[in] messageMapIndex The index of the message map that the call originates from.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IToolBarEvents::Fire_XDblClick, TBarCtlsLibU::_IToolBarEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick
	/// \else
	///   \sa Proxy_IToolBarEvents::Fire_XDblClick, TBarCtlsLibA::_IToolBarEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick
	/// \endif
	inline HRESULT Raise_XDblClick(LONG messageMapIndex, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c Font property.
	///
	/// \sa get_Font
	void ApplyFont(void);
	/// \brief <em>Sends the \c TB_SETDRAWTEXTFLAGS message combining various properties</em>
	///
	/// \sa put_HorizontalTextAlignment, put_VerticalTextAlignment,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb787425.aspx">TB_SETDRAWTEXTFLAGS</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/dd162498.aspx">DrawText</a>
	void SetDrawTextFlags(void);

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleObject
	///
	//@{
	/// \brief <em>Implementation of \c IOleObject::DoVerb</em>
	///
	/// Will be called if one of the control's registered actions (verbs) shall be executed; e. g.
	/// executing the "About" verb will display the About dialog.
	///
	/// \param[in] verbID The requested verb's ID.
	/// \param[in] pMessage A \c MSG structure describing the event (such as a double-click) that
	///            invoked the verb.
	/// \param[in] pActiveSite The \c IOleClientSite implementation of the control's active client site
	///            where the event occurred that invoked the verb.
	/// \param[in] reserved Reserved; must be zero.
	/// \param[in] hWndParent The handle of the document window containing the control.
	/// \param[in] pBoundingRectangle A \c RECT structure containing the coordinates and size in pixels
	///            of the control within the window specified by \c hWndParent.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa EnumVerbs,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms694508.aspx">IOleObject::DoVerb</a>
	virtual HRESULT STDMETHODCALLTYPE DoVerb(LONG verbID, LPMSG pMessage, IOleClientSite* pActiveSite, LONG reserved, HWND hWndParent, LPCRECT pBoundingRectangle);
	/// \brief <em>Implementation of \c IOleObject::EnumVerbs</em>
	///
	/// Will be called if the control's container requests the control's registered actions (verbs); e. g.
	/// we provide a verb "About" that will display the About dialog on execution.
	///
	/// \param[out] ppEnumerator Receives the \c IEnumOLEVERB implementation of the verbs' enumerator.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692781.aspx">IOleObject::EnumVerbs</a>
	virtual HRESULT STDMETHODCALLTYPE EnumVerbs(IEnumOLEVERB** ppEnumerator);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IOleControl
	///
	//@{
	/// \brief <em>Implementation of \c IOleControl::GetControlInfo</em>
	///
	/// Will be called if the container requests details about the control's keyboard mnemonics and keyboard
	/// behavior.
	///
	/// \param[in, out] pControlInfo The requested details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	// \sa OnMnemonic, Properties::hAcceleratorTable,
	//     <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	virtual HRESULT STDMETHODCALLTYPE GetControlInfo(LPCONTROLINFO pControlInfo);
	// \brief <em>Implementation of \c IOleControl::OnMnemonic</em>
	//
	// Will be called if the user has pressed a keystroke that matches one of the entries in the control's
	// accelerator table.
	//
	// \param[in] pMessage The \c MSG structure describing the keystroke.
	//
	// \return An \c HRESULT error code.
	//
	// \sa GetControlInfo, Properties::hAcceleratorTable,
	//     <a href="https://msdn.microsoft.com/en-us/library/ms680699.aspx">IOleControl::OnMnemonic</a>
	//virtual HRESULT STDMETHODCALLTYPE OnMnemonic(LPMSG pMessage);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Executes the About verb</em>
	///
	/// Will be called if the control's registered actions (verbs) "About" shall be executed. We'll
	/// display the About dialog.
	///
	/// \param[in] hWndParent The window to use as parent for any user interface.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa DoVerb, About
	HRESULT DoVerbAbout(HWND hWndParent);

	//////////////////////////////////////////////////////////////////////
	/// \name MFC clones
	///
	//@{
	/// \brief <em>A rewrite of MFC's \c COleControl::RecreateControlWindow</em>
	///
	/// Destroys and re-creates the control window.
	///
	/// \remarks This rewrite probably isn't complete.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/5tx8h2fd.aspx">COleControl::RecreateControlWindow</a>
	void RecreateControlWindow(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Control window configuration
	///
	//@{
	/// \brief <em>Calculates the extended window style bits</em>
	///
	/// Calculates the extended window style bits that the control must have set to match the current
	/// property settings.
	///
	/// \return A bit field of \c WS_EX_* constants specifying the required extended window styles.
	///
	/// \sa GetStyleBits, SendConfigurationMessages, Create
	DWORD GetExStyleBits(void);
	/// \brief <em>Calculates the window style bits</em>
	///
	/// Calculates the window style bits that the control must have set to match the current property
	/// settings.
	///
	/// \return A bit field of \c WS_* constants specifying the required window styles.
	///
	/// \sa GetExStyleBits, SendConfigurationMessages, Create
	DWORD GetStyleBits(void);
	/// \brief <em>Configures the control by sending messages</em>
	///
	/// Sends \c WM_* and \c TB_* messages to the control window to make it match the current property
	/// settings. Will be called out of \c Raise_RecreatedControlWindow.
	///
	/// \sa GetExStyleBits, GetStyleBits, Raise_RecreatedControlWindow
	void SendConfigurationMessages(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Value translation
	///
	//@{
	/// \brief <em>Translates a \c MousePointerConstants type mouse cursor into a \c HCURSOR type mouse cursor</em>
	///
	/// \param[in] mousePointer The \c MousePointerConstants type mouse cursor to translate.
	///
	/// \return The translated \c HCURSOR type mouse cursor.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \else
	///   \sa TBarCtlsLibA::MousePointerConstants, OnSetCursor, put_MousePointer
	/// \endif
	HCURSOR MousePointerConst2hCursor(MousePointerConstants mousePointer);
	/// \brief <em>Translates a button's unique ID to its index</em>
	///
	/// \param[in] ID The unique ID of the button whose index is requested.
	/// \param[in] hWndToUse The window handle of the tool bar control for which to retrieve the button
	///            index. If set to \c NULL, the control itself is used.
	///
	/// \return The requested button's index if successful; -1 otherwise.
	///
	/// \sa ButtonIndexToID, ToolBarButtons::get_Item, ToolBarButtons::Remove
	int IDToButtonIndex(LONG ID, HWND hWndToUse = NULL);
	/// \brief <em>Translates a button's index to its unique ID</em>
	///
	/// \param[in] buttonIndex The index of the button whose unique ID is requested.
	///
	/// \return The requested button's unique ID if successful; -1 otherwise.
	///
	/// \sa IDToButtonIndex, ToolBarButton::get_ID
	LONG ButtonIndexToID(int buttonIndex);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Menu mode
	///
	//@{
	/// \brief <em>Message handling method for the control's top-level parent window</em>
	///
	/// This method is the callback method used when subclassing the control's top-level parent window to
	/// make the control run in menu mode.
	///
	/// \param[in] hWnd The handle of the window that the message has been sent to. This always should be the
	///            top-level parent window.
	/// \param[in] message The message to handle.
	/// \param[in] wParam The message's \c wParam parameter.
	/// \param[in] lParam The message's \c lParam parameter.
	/// \param[in] idSubclass A unique ID identifying the instance of the tool bar control that this message
	///            is meant for. This ID must have been passed to \c SetWindowSubclass when subclassing the
	///            window. We pass the instance pointer (\c this).
	/// \param[in] refData An application defined data that can be passed to \c SetWindowSubclass and will be
	///            passed back to the message handling method. We do not use this feature.
	///
	/// \return A message-dependent return value.
	///
	/// \sa SendConfigurationMessages, Raise_DestroyedControlWindow, parentWindowMenuMode,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762102.aspx">SetWindowSubclass</a>
	static LRESULT CALLBACK ParentWindowSubclass_MenuMode(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/);
	/// \brief <em>Message handling method for the control's top-level parent window</em>
	///
	/// This method is the callback method used when subclassing the control's top-level parent window to
	/// make the chevron popup menu work.
	///
	/// \param[in] hWnd The handle of the window that the message has been sent to. This always should be the
	///            top-level parent window.
	/// \param[in] message The message to handle.
	/// \param[in] wParam The message's \c wParam parameter.
	/// \param[in] lParam The message's \c lParam parameter.
	/// \param[in] idSubclass A unique ID identifying the instance of the tool bar control that this message
	///            is meant for. This ID must have been passed to \c SetWindowSubclass when subclassing the
	///            window. We pass the instance pointer (\c this).
	/// \param[in] refData An application defined data that can be passed to \c SetWindowSubclass and will be
	///            passed back to the message handling method. We do not use this feature.
	///
	/// \return A message-dependent return value.
	///
	/// \sa DisplayChevronPopupWindow, Raise_DestroyedControlWindow, parentWindowChevronPopupMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762102.aspx">SetWindowSubclass</a>
	static LRESULT CALLBACK ParentWindowSubclass_ChevronPopupMenu(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/);
	/// \brief <em>Synchronizes some internal variables with the system settings</em>
	///
	/// \sa OnSettingChange
	void GetSystemSettings(void);
	/// \brief <em>Sets the input focus to the tool bar control and stores the handle of the window that had the focus previously</em>
	///
	/// \sa GiveFocusBack, get_MenuMode, OnDropDownNotification
	void TakeFocus(void);
	/// \brief <em>Sets the input focus back to the window that had it when we called \c TakeFocus</em>
	///
	/// \sa TakeFocus, get_MenuMode, DoPopupMenu
	void GiveFocusBack(void);
	/// \brief <em>Activates or deactivates drawing button texts with keyboard cues</em>
	///
	/// \param[in] show If \c TRUE, keyboard cues will be displayed; otherwise not.
	///
	/// \sa get_MenuMode
	void ShowKeyboardCues(BOOL show);
	/// \brief <em>Retrieves the zero-based index of the next menu item button to the left of the specified button</em>
	///
	/// \param[in] buttonIndex Zero-based index of the button at which to start the search.
	///
	/// \return The zero-based index of the next menu item button to the left of the button specified by
	///         \c buttonIndex. -1 in case of an error, -2 if the chevron menu is next.
	///
	/// \sa GetNextMenuItem, OnHookKeyDown
	int GetPreviousMenuItem(int buttonIndex);
	/// \brief <em>Retrieves the zero-based index of the next menu item button to the right of the specified button</em>
	///
	/// \param[in] buttonIndex Zero-based index of the button at which to start the search.
	///
	/// \return The zero-based index of the next menu item button to the right of the button specified by
	///         \c buttonIndex. -1 in case of an error, -2 if the chevron menu is next.
	///
	/// \sa GetPreviousMenuItem, OnHookKeyDown
	int GetNextMenuItem(int buttonIndex);
	/// \brief <em>Makes a parent rebar control display the chevron menu</em>
	///
	/// \return \c TRUE if the chevron button was pressed; otherwise \c FALSE.
	///
	/// \sa OnHookKeyDown, DisplayChevronPopupWindow
	BOOL DisplayChevronMenu(void);
	/// \brief <em>Displays the specified popup menu for the specified tool bar button</em>
	///
	/// \param[in] buttonIndex The zero-based index of the button for which to display the menu.
	/// \param[in] hPopupMenu The menu to display.
	/// \param[in] animate If \c TRUE, displaying the menu will be animated.
	///
	/// \sa get_MenuMode, OnDropDownNotification, DoTrackPopupMenu
	void DoPopupMenu(int buttonIndex, HMENU hPopupMenu, BOOL animate);
	/// \brief <em>Displays the specified popup menu at the specified position with the specified parameters</em>
	///
	/// \param[in] hPopupMenu The menu to display.
	/// \param[in] flags The flags to pass to \c TrackPopupMenuEx.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the screen's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the screen's
	///            upper-left corner.
	/// \param[in] pParams The extra parameters to pass to \c TrackPopupMenuEx.
	///
	/// \return \c TRUE on success; otherwise \c FALSE.
	///
	/// \sa get_MenuMode, DoPopupMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648003.aspx">TrackPopupMenuEx</a>
	BOOL DoTrackPopupMenu(HMENU hPopupMenu, UINT flags, int x, int y, LPTPMPARAMS pParams = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the tool bar button that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pFlags A bit field of TBHT_* flags, that holds further details about the control's
	///             part below the specified point.
	/// \param[in] hWndToUse The window handle of the tool bar control for which to hit-test the point.
	///            If set to \c NULL, the control itself is used.
	// \param[in] ignoreBoundingBoxDefinition If \c TRUE the setting of the \c ButtonBoundingBoxDefinition
	//            property is ignored; otherwise the returned tool bar button index is set to -1, if the
	//            \c ButtonBoundingBoxDefinition property's setting and the \c pFlags parameter don't
	///            match.
	///
	/// \return The "hit" tool bar button's zero-based index.
	///
	/// \sa HitTest
	// \sa HitTest, get_ButtonBoundingBoxDefinition
	int HitTest(LONG x, LONG y, UINT* pFlags, HWND hWndToUse = NULL/*, BOOL ignoreBoundingBoxDefinition = FALSE*/);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);
	// \brief <em>Recreates the control's accelerator table</em>
	//
	// \sa Properties::hAcceleratorTable, OnAddButtons, OnInsertButton, OnDeleteButton, GetControlInfo,
	//     OnMnemonic
	//void RebuildAcceleratorTable(void);

	/// \brief <em>Retrieves whether the logical left mouse button is held down</em>
	///
	/// \return \c TRUE if the logical left mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsRightMouseButtonDown
	BOOL IsLeftMouseButtonDown(void);
	/// \brief <em>Retrieves whether the logical right mouse button is held down</em>
	///
	/// \return \c TRUE if the logical right mouse button is held down; otherwise \c FALSE.
	///
	/// \sa IsLeftMouseButtonDown
	BOOL IsRightMouseButtonDown(void);


	/// \brief <em>Holds a control instance's properties' settings</em>
	typedef struct Properties
	{
		/// \brief <em>Holds a font property's settings</em>
		typedef struct FontProperty
		{
		protected:
			/// \brief <em>Holds the control's default font</em>
			///
			/// \sa GetDefaultFont
			static FONTDESC defaultFont;

		public:
			/// \brief <em>Holds whether we're listening for events fired by the font object</em>
			///
			/// If greater than 0, we're advised to the \c IFontDisp object identified by \c pFontDisp. I. e.
			/// we will be notified if a property of the font object changes. If 0, we won't receive any events
			/// fired by the \c IFontDisp object.
			///
			/// \sa pFontDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>Flag telling \c OnSetFont not to retrieve the current font if set to \c TRUE</em>
			///
			/// \sa OnSetFont
			UINT dontGetFontObject : 1;
			/// \brief <em>The control's current font</em>
			///
			/// \sa ApplyFont, owningFontResource
			CFont currentFont;
			/// \brief <em>If \c TRUE, \c currentFont may destroy the font resource; otherwise not</em>
			///
			/// \sa currentFont
			UINT owningFontResource : 1;
			/// \brief <em>A pointer to the font object's implementation of \c IFontDisp</em>
			IFontDisp* pFontDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<ToolBar> >* pPropertyNotifySink;

			FontProperty()
			{
				watching = 0;
				dontGetFontObject = FALSE;
				owningFontResource = TRUE;
				pFontDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~FontProperty()
			{
				Release();
			}

			FontProperty& operator =(const FontProperty& source)
			{
				Release();

				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				pFontDisp = source.pFontDisp;
				if(pFontDisp) {
					pFontDisp->AddRef();
				}
				owningFontResource = source.owningFontResource;
				if(!owningFontResource) {
					currentFont.Attach(source.currentFont.m_hFont);
				}
				dontGetFontObject = source.dontGetFontObject;

				if(source.watching > 0) {
					StartWatching();
				}

				return *this;
			}

			/// \brief <em>Retrieves a default font that may be used</em>
			///
			/// \return A \c FONTDESC structure containing the default font.
			///
			/// \sa defaultFont
			static FONTDESC GetDefaultFont(void)
			{
				return defaultFont;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(ToolBar* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ToolBar> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(owningFontResource) {
					if(!currentFont.IsNull()) {
						currentFont.DeleteObject();
					}
				} else {
					currentFont.Detach();
				}
				if(pFontDisp) {
					pFontDisp->Release();
					pFontDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pFontDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pFontDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} FontProperty;

		/// \brief <em>Holds a picture property's settings</em>
		typedef struct PictureProperty
		{
			/// \brief <em>Holds whether we're listening for events fired by the picture object</em>
			///
			/// If greater than 0, we're advised to the \c IPictureDisp object identified by \c pPictureDisp.
			/// I. e. we will be notified if a property of the picture object changes. If 0, we won't receive any
			/// events fired by the \c IPictureDisp object.
			///
			/// \sa pPictureDisp, pPropertyNotifySink
			int watching;
			/// \brief <em>A pointer to the picture object's implementation of \c IPictureDisp</em>
			IPictureDisp* pPictureDisp;
			/// \brief <em>Receives notifications on changes to this property object's settings</em>
			///
			/// \sa InitializePropertyWatcher, PropertyNotifySinkImpl
			CComObject< PropertyNotifySinkImpl<ToolBar> >* pPropertyNotifySink;

			PictureProperty()
			{
				watching = 0;
				pPictureDisp = NULL;
				pPropertyNotifySink = NULL;
			}

			~PictureProperty()
			{
				Release();
			}

			PictureProperty& operator =(const PictureProperty& source)
			{
				Release();

				pPictureDisp = source.pPictureDisp;
				if(pPictureDisp) {
					pPictureDisp->AddRef();
				}
				InitializePropertyWatcher(source.pPropertyNotifySink->properties.pObjectToNotify, source.pPropertyNotifySink->properties.propertyToWatch);
				if(source.watching > 0) {
					StartWatching();
				}
				return *this;
			}

			/// \brief <em>Initializes an object that will watch this property for changes</em>
			///
			/// \param[in] pObjectToNotify The object to notify on changes.
			/// \param[in] propertyToWatch The property to watch for changes.
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, StopWatching
			HRESULT InitializePropertyWatcher(ToolBar* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ToolBar> >::CreateInstance(&pPropertyNotifySink);
				ATLASSUME(pPropertyNotifySink);
				pPropertyNotifySink->AddRef();
				return pPropertyNotifySink->Initialize(pObjectToNotify, propertyToWatch);
			}

			/// \brief <em>Prepares the object for destruction</em>
			void Release(void)
			{
				if(pPropertyNotifySink) {
					StopWatching();
					pPropertyNotifySink->Release();
					pPropertyNotifySink = NULL;
				}
				ATLASSERT(watching == 0);
				if(pPictureDisp) {
					pPictureDisp->Release();
					pPictureDisp = NULL;
				}
			}

			/// \brief <em>Starts watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StopWatching, InitializePropertyWatcher
			HRESULT StartWatching(void)
			{
				if(pPictureDisp) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StartWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						++watching;
					}
					return hr;
				}
				return E_FAIL;
			}

			/// \brief <em>Stops watching the property object for changes</em>
			///
			/// \return An \c HRESULT error code.
			///
			/// \sa StartWatching, InitializePropertyWatcher
			HRESULT StopWatching(void)
			{
				if(watching > 0) {
					ATLASSUME(pPropertyNotifySink);
					HRESULT hr = pPropertyNotifySink->StopWatching(pPictureDisp);
					ATLASSERT(SUCCEEDED(hr));
					if(SUCCEEDED(hr)) {
						--watching;
					}
					return hr;
				}
				return E_FAIL;
			}
		} PictureProperty;

		/// \brief <em>Holds the \c AllowCustomization property's setting</em>
		///
		/// \sa get_AllowCustomization, put_AllowCustomization
		UINT allowCustomization : 1;
		/// \brief <em>Holds the \c AlwaysDisplayButtonText property's setting</em>
		///
		/// \sa get_AlwaysDisplayButtonText, put_AlwaysDisplayButtonText
		UINT alwaysDisplayButtonText : 1;
		/// \brief <em>Holds the \c AnchorHighlighting property's setting</em>
		///
		/// \sa get_AnchorHighlighting, put_AnchorHighlighting
		UINT anchorHighlighting : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c BackStyle property's setting</em>
		///
		/// \sa get_BackStyle, put_BackStyle
		BackStyleConstants backStyle;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c ButtonHeight property's setting</em>
		///
		/// \sa get_ButtonHeight, put_ButtonHeight
		OLE_YSIZE_PIXELS buttonHeight;
		/// \brief <em>Holds the \c ButtonStyle property's setting</em>
		///
		/// \sa get_ButtonStyle, put_ButtonStyle
		ButtonStyleConstants buttonStyle;
		/// \brief <em>Holds the \c ButtonTextPosition property's setting</em>
		///
		/// \sa get_ButtonTextPosition, put_ButtonTextPosition
		ButtonTextPositionConstants buttonTextPosition;
		/// \brief <em>Holds the \c ButtonWidth property's setting</em>
		///
		/// \sa get_ButtonWidth, put_ButtonWidth
		OLE_XSIZE_PIXELS buttonWidth;
		#ifdef USE_STL
			/// \brief <em>Holds the currently disabled command IDs</em>
			///
			/// \sa get_CommandEnabled, put_CommandEnabled
			std::vector<LONG> disabledCommands;
		#else
			/// \brief <em>Holds the currently disabled command IDs</em>
			///
			/// \sa get_CommandEnabled, put_CommandEnabled
			CAtlArray<LONG> disabledCommands;
		#endif
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DisplayMenuDivider property's setting</em>
		///
		/// \sa get_DisplayMenuDivider, put_DisplayMenuDivider
		UINT displayMenuDivider : 1;
		/// \brief <em>Holds the \c DisplayPartiallyClippedButtons property's setting</em>
		///
		/// \sa get_DisplayPartiallyClippedButtons, put_DisplayPartiallyClippedButtons
		UINT displayPartiallyClippedButtons : 1;
		/// \brief <em>Holds the \c DontRedraw property's setting</em>
		///
		/// \sa get_DontRedraw, put_DontRedraw
		UINT dontRedraw : 1;
		/// \brief <em>Holds the \c DragClickTime property's setting</em>
		///
		/// \sa get_DragClickTime, put_DragClickTime
		long dragClickTime;
		/// \brief <em>Holds the \c DragDropCustomizationModifierKey property's setting</em>
		///
		/// \sa get_DragDropCustomizationModifierKey, put_DragDropCustomizationModifierKey
		DragDropCustomizationModifierKeyConstants dragDropCustomizationModifierKey;
		/// \brief <em>Holds the \c DropDownGap property's setting</em>
		///
		/// \sa get_DropDownGap, put_DropDownGap
		OLE_XSIZE_PIXELS dropDownGap;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c FirstButtonIndentation property's setting</em>
		///
		/// \sa get_FirstButtonIndentation, put_FirstButtonIndentation
		OLE_XSIZE_PIXELS firstButtonIndentation;
		/// \brief <em>Holds the \c FocusOnClick property's setting</em>
		///
		/// \sa get_FocusOnClick, put_FocusOnClick
		UINT focusOnClick : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		// \brief <em>Holds the control's accelerator table</em>
		//
		// \sa RegisterHotkey, pAccelerators
		// \sa RebuildAcceleratorTable, GetControlInfo, OnMnemonic
		//HACCEL hAcceleratorTable;
		/// \brief <em>Holds the control's accelerators</em>
		///
		/// \sa RegisterHotkey, numberOfAccelerators
		LPACCEL pAccelerators;
		/// \brief <em>Holds the count of the control's accelerators</em>
		///
		/// \sa RegisterHotkey, pAccelerators
		UINT numberOfAccelerators;
		#ifdef USE_STL
			/// \brief <em>Holds the \c ilHighResolution imagelists</em>
			///
			/// \sa get_hImageList, put_hImageList
			std::unordered_map<LONG, HIMAGELIST> hHighResImageLists;
		#else
			/// \brief <em>Holds the \c ilHighResolution imagelists</em>
			///
			/// \sa get_hImageList, put_hImageList
			CAtlMap<LONG, HIMAGELIST> hHighResImageLists;
		#endif
		/// \brief <em>Holds the \c HighlightColor property's setting</em>
		///
		/// \sa get_HighlightColor, put_HighlightColor
		OLE_COLOR highlightColor;
		/// \brief <em>Holds the \c HorizontalButtonPadding property's setting</em>
		///
		/// \sa get_HorizontalButtonPadding, put_HorizontalButtonPadding
		OLE_XSIZE_PIXELS horizontalButtonPadding;
		/// \brief <em>Holds the \c HorizontalButtonSpacing property's setting</em>
		///
		/// \sa get_HorizontalButtonSpacing, put_HorizontalButtonSpacing
		OLE_XSIZE_PIXELS horizontalButtonSpacing;
		/// \brief <em>Holds the \c HorizontalIconCaptionGap property's setting</em>
		///
		/// \sa get_HorizontalIconCaptionGap, put_HorizontalIconCaptionGap
		OLE_XSIZE_PIXELS horizontalIconCaptionGap;
		/// \brief <em>Holds the \c HorizontalTextAlignment property's setting</em>
		///
		/// \sa get_HorizontalTextAlignment, put_HorizontalTextAlignment
		HAlignmentConstants horizontalTextAlignment;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c InsertMarkColor property's setting</em>
		///
		/// \sa get_InsertMarkColor, put_InsertMarkColor
		OLE_COLOR insertMarkColor;
		/// \brief <em>Holds the \c MaximumButtonWidth property's setting</em>
		///
		/// \sa get_MaximumButtonWidth, put_MaximumButtonWidth
		OLE_XSIZE_PIXELS maximumButtonWidth;
		/// \brief <em>Holds the \c MaximumTextRows property's setting</em>
		///
		/// \sa get_MaximumTextRows, put_MaximumTextRows
		LONG maximumTextRows;
		/// \brief <em>Holds the \c MenuBarTheme property's setting</em>
		///
		/// \sa get_MenuBarTheme, put_MenuBarTheme
		MenuBarThemeConstants menuBarTheme;
		/// \brief <em>Holds the \c MenuMode property's setting</em>
		///
		/// \sa get_MenuMode, put_MenuMode
		UINT menuMode : 1;
		/// \brief <em>Holds the \c MinimumButtonWidth property's setting</em>
		///
		/// \sa get_MinimumButtonWidth, put_MinimumButtonWidth
		OLE_XSIZE_PIXELS minimumButtonWidth;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c MultiColumn property's setting</em>
		///
		/// \sa get_MultiColumn, put_MultiColumn
		UINT multiColumn : 1;
		/// \brief <em>Holds the \c NormalDropDownButtonStyle property's setting</em>
		///
		/// \sa get_NormalDropDownButtonStyle, put_NormalDropDownButtonStyle
		NormalDropDownButtonStyleConstants normalDropDownButtonStyle;
		/// \brief <em>Holds the \c OLEDragImageStyle property's setting</em>
		///
		/// \sa get_OLEDragImageStyle, put_OLEDragImageStyle
		OLEDragImageStyleConstants oleDragImageStyle;
		/// \brief <em>Holds the \c Orientation property's setting</em>
		///
		/// \sa get_Orientation, put_Orientation
		OrientationConstants orientation;
		/// \brief <em>Holds the \c ProcessContextMenuKeys property's setting</em>
		///
		/// \sa get_ProcessContextMenuKeys, put_ProcessContextMenuKeys
		UINT processContextMenuKeys : 1;
		/// \brief <em>Holds the \c RaiseCustomDrawEventOnEraseBackground property's setting</em>
		///
		/// \sa get_RaiseCustomDrawEventOnEraseBackground, put_RaiseCustomDrawEventOnEraseBackground
		UINT raiseCustomDrawEventOnEraseBackground : 1;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		RegisterForOLEDragDropConstants registerForOLEDragDrop;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c ShadowColor property's setting</em>
		///
		/// \sa get_ShadowColor, put_ShadowColor
		OLE_COLOR shadowColor;
		/// \brief <em>Holds the \c ShowToolTips property's setting</em>
		///
		/// \sa get_ShowToolTips, put_ShowToolTips
		UINT showToolTips : 1;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the \c UseMnemonics property's setting</em>
		///
		/// \sa get_UseMnemonics, put_UseMnemonics
		UINT useMnemonics : 1;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;
		/// \brief <em>Holds the \c VerticalButtonPadding property's setting</em>
		///
		/// \sa get_VerticalButtonPadding, put_VerticalButtonPadding
		OLE_YSIZE_PIXELS verticalButtonPadding;
		/// \brief<em>Holds the default values for the \c HorizontalButtonPadding and \c VerticalButtonPadding properties as specified by comctl32.dll</em>
		///
		/// \sa horizontalButtonPadding, verticalButtonPadding
		LPARAM defaultButtonPadding;
		/// \brief <em>Holds the \c VerticalButtonSpacing property's setting</em>
		///
		/// \sa get_VerticalButtonSpacing, put_VerticalButtonSpacing
		OLE_YSIZE_PIXELS verticalButtonSpacing;
		/// \brief <em>Holds the \c VerticalTextAlignment property's setting</em>
		///
		/// \sa get_VerticalTextAlignment, put_VerticalTextAlignment
		VAlignmentConstants verticalTextAlignment;
		/// \brief <em>Holds the \c WrapButtons property's setting</em>
		///
		/// \sa get_WrapButtons, put_WrapButtons
		UINT wrapButtons : 1;
		/// \brief <em>Holds the \c DT_* flags set by the \c TB_SETDRAWTEXTFLAGS message</em>
		///
		/// \sa OnSetDrawTextFlags
		UINT drawTextFlags;
		/// \brief <em>Holds the \c DT_* flags mask set by the \c TB_SETDRAWTEXTFLAGS message</em>
		///
		/// \sa OnSetDrawTextFlags
		UINT drawTextFlagsMask;
		/// \brief <em>Holds the \c DT_* flags that are the default for the current setting of the \c ButtonTextPosition property</em>
		///
		/// Depending on the position of the button text, the tool bar control uses different combinations of
		/// \c DT_* flags as the default when drawing text. To calculate the resulting flags in
		/// \c OnSetDrawTextFlags we store the correct defaults when setting the \c ButtonTextPosition
		/// property.
		///
		/// \sa get_ButtonTextPosition, OnSetDrawTextFlags
		UINT defaultDrawTextFlags;
		/// \brief <em>Holds the chevron popup menu if we are in menu mode</em>
		HMENU hChevronPopupMenu;
		/// \brief <em>Holds the accelerator character of the tool bar button that should be preselected in the chevron popup</em>
		USHORT acceleratorToSendToChevronMenu;

		Properties()
		{
			ResetToDefaults();
		}

		~Properties()
		{
			Release();
		}

		/// \brief <em>Prepares the object for destruction</em>
		void Release(void)
		{
			font.Release();
			/*if(hAcceleratorTable) {
				DestroyAcceleratorTable(hAcceleratorTable);
				hAcceleratorTable = NULL;
			}*/
			if(pAccelerators) {
				HeapFree(GetProcessHeap(), 0, pAccelerators);
				pAccelerators = NULL;
			}
			numberOfAccelerators = 0;
			if(hChevronPopupMenu) {
				DestroyMenu(hChevronPopupMenu);
			}
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			allowCustomization = FALSE;
			alwaysDisplayButtonText = FALSE;
			anchorHighlighting = FALSE;
			appearance = a2D;
			backStyle = bksTransparent;
			borderStyle = bsNone;
			buttonHeight = 0;
			buttonStyle = bstFlat;
			buttonTextPosition = btpBelowIcon;
			buttonWidth = 0;
			focusOnClick = FALSE;
			#ifdef USE_STL
				disabledCommands.clear();
			#else
				disabledCommands.RemoveAll();
			#endif
			disabledEvents = static_cast<DisabledEventsConstants>(deMouseEvents | deClickEvents | deKeyboardEvents | deButtonInsertionEvents | deButtonDeletionEvents | deFreeButtonData | deCustomDraw | deHotButtonChangeEvents | deAcceleratorEvents | deRawMenuMessage);
			displayMenuDivider = FALSE;
			displayPartiallyClippedButtons = FALSE;
			dontRedraw = FALSE;
			dragClickTime = -1;
			dragDropCustomizationModifierKey = ddcmkShift;
			dropDownGap = -1;
			enabled = TRUE;
			firstButtonIndentation = 0;
			//hAcceleratorTable = NULL;
			pAccelerators = NULL;
			numberOfAccelerators = 0;
			highlightColor = static_cast<OLE_COLOR>(-1);
			horizontalButtonPadding = -1;
			horizontalButtonSpacing = 0;
			horizontalIconCaptionGap = -1;
			horizontalTextAlignment = halLeft;
			hoverTime = -1;
			insertMarkColor = 0;
			maximumButtonWidth = 0;
			maximumTextRows = 1;
			menuBarTheme = mbtNativeMenuBar;
			menuMode = FALSE;
			minimumButtonWidth = 0;
			mousePointer = mpDefault;
			multiColumn = FALSE;
			normalDropDownButtonStyle = nddbsSplitButton;
			oleDragImageStyle = odistClassic;
			orientation = oHorizontal;
			processContextMenuKeys = TRUE;
			raiseCustomDrawEventOnEraseBackground = FALSE;
			registerForOLEDragDrop = rfoddNoDragDrop;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			shadowColor = static_cast<OLE_COLOR>(-1);
			showToolTips = TRUE;
			supportOLEDragImages = TRUE;
			useMnemonics = TRUE;
			useSystemFont = TRUE;
			verticalButtonPadding = -1;
			defaultButtonPadding = MAKELPARAM(-1, -1);
			verticalButtonSpacing = 0;
			verticalTextAlignment = valTop;
			wrapButtons = FALSE;
			drawTextFlags = 0;
			drawTextFlagsMask = 0;
			defaultDrawTextFlags = 0;
			hChevronPopupMenu = NULL;
			acceleratorToSendToChevronMenu = 0;
		}
	} Properties;
	/// \brief <em>Holds the control's properties' settings</em>
	Properties properties;

	/// \brief <em>Holds the control's flags</em>
	struct Flags
	{
		/// \brief <em>If \c TRUE, the control has been activated by mouse and needs to be UI-activated by \c OnSetFocus</em>
		///
		/// ATL always UI-activates the control in \c OnMouseActivate. If the control is activated by mouse,
		/// \c WM_SETFOCUS is sent after \c WM_MOUSEACTIVATE, but Visual Basic 6 won't raise the \c Validate
		/// event if the control already is UI-activated when it receives the focus. Therefore we need to delay
		/// UI-activation.
		///
		/// \sa OnMouseActivate, OnSetFocus, OnKillFocus
		UINT uiActivationPending : 1;
		/// \brief <em>If \c TRUE, we're using themes</em>
		///
		/// \sa OnThemeChanged
		UINT usingThemes : 1;
		/// \brief <em>If \c TRUE, we will avoid the info tip being displayed</em>
		///
		/// \sa OnGetInfoTipNotification, OnToolTipGetDispInfoNotificationW
		UINT cancelToolTip : 1;
		/// \brief <em>If \c TRUE, the tool bar is currently being customized</em>
		///
		/// \sa OnBeginAdjustNotification, OnEndAdjustNotification
		UINT isBeingCustomized : 1;
		/// \brief <em>If \c TRUE, the tool bar is currently being restored from the registry</em>
		///
		/// \sa OnBeginAdjustNotification, Raise_InitializeToolBarStateRegistryRestorage,
		///     Raise_EndCustomization
		UINT isBeingRestored : 1;
		/// \brief <em>If \c TRUE, the current \c TBN_DELETINGBUTTON notification was sent because a button was removed through the customization dialog</em>
		///
		/// \sa OnBeginAdjustNotification, OnEndAdjustNotification, OnDeleteButton
		UINT buttonIsRemovedByCustomizationDialog : 1;
		/// \brief <em>If \c TRUE, the \c HandleKeyboardMessage method already is being executed</em>
		///
		/// \sa HandleKeyboardMessage
		UINT recursiveHookCall : 1;
		/// \brief <em>If \c TRUE, the \c HandleKeyboardMessage method already is executing a hotkey command</em>
		///
		/// \sa HandleKeyboardMessage
		UINT executingHotKeyCommand : 1;
		/// \brief <em>If \c TRUE, we allow receiving the keyboard focus even though the \c CanGetFocus property is set to \c VARIANT_FALSE</em>
		///
		/// \sa Properties::focusOnClick, TakeFocus
		UINT allowSetFocus : 1;
		/// \brief <em>If \c TRUE, the chevron popup window is visible; otherwise not</em>
		///
		/// \sa DisplayChevronPopupWindow, OnChevronPopupMenuDismiss
		UINT chevronPopupVisible : 1;
		/// \brief <em>If \c TRUE, \c OnWindowPosChanged won't report the new control rectangle to the COM container</em>
		///
		/// \sa SendConfigurationMessages, OnWindowPosChanged
		UINT ignoreResize : 1;
		/// \brief <em>If \c TRUE, we force the \c TBSTYLE_CUSTOMERASE style and draw the background using custom draw; otherwise not</em>
		///
		/// \sa put_RaiseCustomDrawEventOnEraseBackground, OnCustomDrawNotification, OnThemeChanged,
		///     Raise_RecreatedControlWindow
		UINT applyBackgroundHack : 1;

		Flags()
		{
			uiActivationPending = FALSE;
			usingThemes = FALSE;
			cancelToolTip = FALSE;
			isBeingCustomized = FALSE;
			isBeingRestored = FALSE;
			buttonIsRemovedByCustomizationDialog = FALSE;
			recursiveHookCall = FALSE;
			executingHotKeyCommand = FALSE;
			allowSetFocus = FALSE;
			chevronPopupVisible = FALSE;
			ignoreResize = FALSE;
			applyBackgroundHack = FALSE;
		}
	} flags;

	//////////////////////////////////////////////////////////////////////
	/// \name Button management
	///
	//@{
	/// \brief <em>Hold details about a placeholder button</em>
	typedef struct PLACEHOLDERBUTTON
	{
		/// \brief <em>The handle of the window that is positioned at the placeholder button's position</em>
		HWND hContainedWindow;
		/// \brief <em>The handle of the window that originally contained the window specified by \c hContainedWindow</em>
		HWND hOriginalParentWindow;
		/// \brief <em>The horizontal alignment of the window specified by \c hContainedWindow inside the placeholder button's rectangle</em>
		HAlignmentConstants horizontalChildWindowAlignment;
		/// \brief <em>The vertical alignment of the window specified by \c hContainedWindow inside the placeholder button's rectangle</em>
		VAlignmentConstants verticalChildWindowAlignment;

		PLACEHOLDERBUTTON()
		{
			hContainedWindow = NULL;
			hOriginalParentWindow = NULL;
			horizontalChildWindowAlignment = halCenter;
			verticalChildWindowAlignment = valCenter;
		}
	} PLACEHOLDERBUTTON, *LPPLACEHOLDERBUTTON;
	#ifdef USE_STL
		/// \brief <em>A map of all \c ToolBarButtonContainer objects that we've created</em>
		///
		/// Holds pointers to all \c ToolBarButtonContainer objects that we've created. We use this map to
		/// inform the containers of removed buttons. The container's ID is stored as key; the container's
		/// \c IButtonContainer implementation is stored as value.
		///
		/// \sa CreateButtonContainer, RegisterButtonContainer, ToolBarButtonContainer
		std::unordered_map<DWORD, IButtonContainer*> buttonContainers;
		/// \brief <em>A map of all buttons of type \c btyPlaceholder</em>
		///
		/// Holds all placeholder buttons. We use this map to distinguish placeholders from command buttons.
		/// The button's ID is stored as key; the value can be set to a \c PLACEHOLDERBUTTON struct that
		/// contains details about the child window that is positioned at the placeholder button's position.
		///
		/// \sa RegisterPlaceholderButton, DeregisterPlaceholderButton, IsPlaceholderButton, PLACEHOLDERBUTTON
		std::unordered_map<LONG, LPPLACEHOLDERBUTTON> placeholderButtons;
	#else
		/// \brief <em>A map of all \c ToolBarButtonContainer objects that we've created</em>
		///
		/// Holds pointers to all \c ToolBarButtonContainer objects that we've created. We use this map to
		/// inform the containers of removed buttons. The container's ID is stored as key; the container's
		/// \c IButtonContainer implementation is stored as value.
		///
		/// \sa CreateButtonContainer, RegisterButtonContainer, ToolBarButtonContainer
		CAtlMap<DWORD, IButtonContainer*> buttonContainers;
		/// \brief <em>A map of all buttons of type \c btyPlaceholder</em>
		///
		/// Holds all placeholder buttons. We use this map to distinguish placeholders from command buttons.
		/// The button's ID is stored as key; the value can be set to a \c PLACEHOLDERBUTTON struct that
		/// contains details about the child window that is positioned at the placeholder button's position.
		///
		/// \sa RegisterPlaceholderButton, DeregisterPlaceholderButton, IsPlaceholderButton, PLACEHOLDERBUTTON
		CAtlMap<LONG, LPPLACEHOLDERBUTTON> placeholderButtons;
	#endif
	/// \brief <em>Registers the specified \c ToolBarButtonContainer collection</em>
	///
	/// Registers the specified \c ToolBarButtonContainer collection so that it is informed of button
	/// deletions.
	///
	/// \param[in] pContainer The container's \c IButtonContainer implementation.
	///
	/// \sa DeregisterButtonContainer, buttonContainers, RemoveButtonFromButtonContainers
	void RegisterButtonContainer(IButtonContainer* pContainer);
	/// \brief <em>De-registers the specified \c ToolBarButtonContainer collection</em>
	///
	/// De-registers the specified \c ToolBarButtonContainer collection so that it no longer is informed of
	/// button deletions.
	///
	/// \param[in] containerID The container's ID.
	///
	/// \sa RegisterButtonContainer, buttonContainers
	void DeregisterButtonContainer(DWORD containerID);
	/// \brief <em>Removes the specified button from all registered \c ToolBarButtonContainer collections</em>
	///
	/// \param[in] buttonID The unique ID of the button to remove. If -1, all buttons are removed.
	///
	/// \sa buttonContainers, RegisterButtonContainer, OnDeleteButton, OnDeletingButtonNotification,
	///     Raise_DestroyedControlWindow
	void RemoveButtonFromButtonContainers(LONG buttonID);
	/// \brief <em>Registers the specified button as a placeholder button</em>
	///
	/// \param[in] buttonID The unique ID of the button to register.
	/// \param[in] hWnd The handle of the window that is positioned above the placeholder.
	///
	/// \sa DeregisterPlaceholderButton, IsPlaceholderButton, GetPlaceholderButtonChildWindow,
	///     SetPlaceholderButtonChildWindow, placeholderButtons
	void RegisterPlaceholderButton(LONG buttonID, HWND hWnd = NULL);
	/// \brief <em>De-registers the specified button as a placeholder button</em>
	///
	/// \param[in] buttonID The unique ID of the button to de-register. If -1, all buttons are de-registered.
	///
	/// \sa RegisterPlaceholderButton, IsPlaceholderButton, SetPlaceholderButtonChildWindow,
	///     placeholderButtons
	void DeregisterPlaceholderButton(LONG buttonID);
	/// \brief <em>Retrieves whether the specified button is a placeholder button</em>
	///
	/// \param[in] buttonID The unique ID of the button to check.
	///
	/// \return \c TRUE, if the button is a placeholder button; otherwise \c FALSE.
	///
	/// \sa placeholderButtons, RegisterPlaceholderButton, GetPlaceholderButtonChildWindow,
	///     DeregisterPlaceholderButton
	BOOL IsPlaceholderButton(LONG buttonID);
	/// \brief <em>Retrieves the handle of the window that is displayed at the position of the specified placeholder button</em>
	///
	/// \param[in] buttonID The unique ID of the button for which to retrieve the child window.
	/// \param[out] horizontalAlignment The horizontal alignment of the window specified by \c hWnd inside
	///             the placeholder button's rectangle. Any of the values defined by the
	///             \c HAlignmentConstants enumeration is valid.
	/// \param[out] verticalAlignment The vertical alignment of the window specified by \c hWnd inside the
	///             placeholder button's rectangle. Any of the values defined by the \c VAlignmentConstants
	///             enumeration is valid.
	///
	/// \return The handle of the window that is displayed at the position of the placeholder button.
	///
	/// \if UNICODE
	///   \sa SetPlaceholderButtonChildWindow, RegisterPlaceholderButton, IsPlaceholderButton,
	///       TBarCtlsLibU::HAlignmentConstants, TBarCtlsLibU::VAlignmentConstants
	/// \else
	///   \sa SetPlaceholderButtonChildWindow, RegisterPlaceholderButton, IsPlaceholderButton,
	///       TBarCtlsLibA::HAlignmentConstants, TBarCtlsLibA::VAlignmentConstants
	/// \endif
	HWND GetPlaceholderButtonChildWindow(LONG buttonID, HAlignmentConstants& horizontalAlignment, VAlignmentConstants& verticalAlignment);
	/// \brief <em>Sets the handle of the window that is displayed at the position of the specified placeholder button</em>
	///
	/// \param[in] buttonID The unique ID of the button for which to set the child window.
	/// \param[in] hWnd The handle of the window that is positioned at the position of the placeholder
	///            button.
	/// \param[in] horizontalAlignment The horizontal alignment of the window specified by \c hWnd inside the
	///            placeholder button's rectangle. Any of the values defined by the \c HAlignmentConstants
	///            enumeration is valid.
	/// \param[in] verticalAlignment The vertical alignment of the window specified by \c hWnd inside the
	///            placeholder button's rectangle. Any of the values defined by the \c VAlignmentConstants
	///            enumeration is valid.
	///
	/// \return \c TRUE, if the button is a placeholder button; otherwise \c FALSE.
	///
	/// \if UNICODE
	///   \sa GetPlaceholderButtonChildWindow, RegisterPlaceholderButton, DeregisterPlaceholderButton,
	///       IsPlaceholderButton, TBarCtlsLibU::HAlignmentConstants, TBarCtlsLibU::VAlignmentConstants
	/// \else
	///   \sa GetPlaceholderButtonChildWindow, RegisterPlaceholderButton, DeregisterPlaceholderButton,
	///       IsPlaceholderButton, TBarCtlsLibA::HAlignmentConstants, TBarCtlsLibA::VAlignmentConstants
	/// \endif
	BOOL SetPlaceholderButtonChildWindow(LONG buttonID, HWND hWnd, HAlignmentConstants horizontalAlignment, VAlignmentConstants verticalAlignment);
	/// \brief <em>Retrieves whether the specified button is currently dropped down</em>
	///
	/// \param[in] buttonIndex The zero-based index of the button to check.
	///
	/// \return \c TRUE, if the button is currently dropped down; otherwise \c FALSE.
	///
	/// \sa droppedDownButton, ToolBarButton::get_DroppedDown
	BOOL IsDroppedDownButton(int buttonIndex);
	///@}
	//////////////////////////////////////////////////////////////////////


	/// \brief <em>Holds the index of the button below the mouse cursor</em>
	///
	/// \attention This member is not reliable with \c deMouseEvents being set.
	int buttonUnderMouse;
	/// \brief <em>Holds the index of the button that is currently dropped down</em>
	///
	/// \sa IsDroppedDownButton
	int droppedDownButton;

	/// \brief <em>Holds mouse status variables</em>
	typedef struct MouseStatus
	{
	protected:
		/// \brief <em>Holds all mouse buttons that may cause a click event in the immediate future</em>
		///
		/// A bit field of \c SHORT values representing those mouse buttons that are currently pressed and
		/// may cause a click event in the immediate future.
		///
		/// \sa StoreClickCandidate, IsClickCandidate, RemoveClickCandidate, Raise_Click, Raise_MClick,
		///     Raise_RClick, Raise_XClick
		SHORT clickCandidates;

	public:
		/// \brief <em>If \c TRUE, the \c MouseEnter event already was raised</em>
		///
		/// \sa Raise_MouseEnter
		UINT enteredControl : 1;
		/// \brief <em>If \c TRUE, the \c MouseHover event already was raised</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa Raise_MouseHover
		UINT hoveredControl : 1;
		/// \brief <em>Holds the mouse cursor's last position</em>
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		POINT lastPosition;
		/// \brief <em>Holds the index of the last clicked button</em>
		///
		/// Holds the index of the last clicked button. We use this to ensure that the \c *DblClick events
		/// are not raised accidently.
		///
		/// \attention This member is not reliable with \c deClickEvents being set.
		///
		/// \sa Raise_Click, Raise_DblClick, Raise_MClick, Raise_MDblClick, Raise_RClick, Raise_RDblClick
		int lastClickedButton;

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
			lastClickedButton = -1;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
			lastClickedButton = -1;
		}

		/// \brief <em>Changes flags to indicate the \c MouseHover event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl, LeaveControl
		void HoverControl(void)
		{
			enteredControl = TRUE;
			hoveredControl = TRUE;
		}

		/// \brief <em>Changes flags to indicate the \c MouseLeave event was just raised</em>
		///
		/// \sa enteredControl, hoveredControl, EnterControl
		void LeaveControl(void)
		{
			enteredControl = FALSE;
			hoveredControl = FALSE;
			lastClickedButton = -1;
		}

		/// \brief <em>Stores a mouse button as click candidate</em>
		///
		/// param[in] button The mouse button to store.
		///
		/// \sa clickCandidates, IsClickCandidate, RemoveClickCandidate
		void StoreClickCandidate(SHORT button)
		{
			// avoid combined click events
			if(clickCandidates == 0) {
				clickCandidates |= button;
			}
		}

		/// \brief <em>Retrieves whether a mouse button is a click candidate</em>
		///
		/// \param[in] button The mouse button to check.
		///
		/// \return \c TRUE if the button is stored as a click candidate; otherwise \c FALSE.
		///
		/// \attention This member is not reliable with \c deMouseEvents being set.
		///
		/// \sa clickCandidates, StoreClickCandidate, RemoveClickCandidate
		BOOL IsClickCandidate(SHORT button)
		{
			return (clickCandidates & button);
		}

		/// \brief <em>Removes a mouse button from the list of click candidates</em>
		///
		/// \param[in] button The mouse button to remove.
		///
		/// \sa clickCandidates, RemoveAllClickCandidates, StoreClickCandidate, IsClickCandidate
		void RemoveClickCandidate(SHORT button)
		{
			clickCandidates &= ~button;
		}

		/// \brief <em>Clears the list of click candidates</em>
		///
		/// \sa clickCandidates, RemoveClickCandidate, StoreClickCandidate, IsClickCandidate
		void RemoveAllClickCandidates(void)
		{
			clickCandidates = 0;
		}
	} MouseStatus;

	/// \brief <em>Holds the control's mouse status</em>
	MouseStatus mouseStatus;

	//////////////////////////////////////////////////////////////////////
	/// \name Drag'n'Drop
	///
	//@{
	/// \brief <em>The \c CLSID_WICImagingFactory object used to create WIC objects that are required during drag image creation</em>
	///
	/// \sa OnGetDragImage, CreateThumbnail
	CComPtr<IWICImagingFactory> pWICImagingFactory;
	/// \brief <em>Creates a thumbnail of the specified icon in the specified size</em>
	///
	/// \param[in] hIcon The icon to create the thumbnail for.
	/// \param[in] size The thumbnail's size in pixels.
	/// \param[in,out] pBits The thumbnail's DIB bits.
	/// \param[in] doAlphaChannelPostProcessing WIC has problems to handle the alpha channel of the icon
	///            specified by \c hIcon. If this parameter is set to \c TRUE, some post-processing is done
	///            to correct the pixel failures. Otherwise the failures are not corrected.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa OnGetDragImage, pWICImagingFactory
	HRESULT CreateThumbnail(HICON hIcon, SIZE& size, LPRGBQUAD pBits, BOOL doAlphaChannelPostProcessing);

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		/// \brief <em>The \c IToolBarButtonContainer implementation of the collection of the dragged buttons</em>
		IToolBarButtonContainer* pDraggedButtons;
		/// \brief <em>The handle of the imagelist containing the drag image</em>
		HIMAGELIST hDragImageList;
		/// \brief <em>Enables or disables auto-destruction of \c hDragImageList</em>
		///
		/// Controls whether the imagelist defined by \c hDragImageList is destroyed automatically. If set to
		/// \c TRUE, it is destroyed in \c EndDrag; otherwise not.
		///
		/// \sa hDragImageList
		UINT autoDestroyImgLst : 1;
		/// \brief <em>Indicates whether the drag image is visible or hidden</em>
		///
		/// If this value is 0, the drag image is visible; otherwise not.
		///
		/// \sa get_ShowDragImage, put_ShowDragImage, ShowDragImage, HideDragImage, IsDragImageVisible
		int dragImageIsHidden;
		/// \brief <em>The zero-based index of the last drop target</em>
		int lastDropTarget;
		/// \brief <em>The zero-based index of the last automatically clicked button</em>
		int lastClickedButton;

		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>The currently dragged data for the case that the we're the drag source</em>
		CComPtr<IDataObject> pSourceDataObject;
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
		/// \brief <em>Holds the mouse button (as \c MK_* constant) that the drag'n'drop operation is performed with</em>
		DWORD draggingMouseButton;
		/// \brief <em>If \c TRUE, we'll hide and re-show the drag image in \c IDropTarget::DragEnter so that the item count label is displayed</em>
		///
		/// \sa DragEnter, OLEDrag
		UINT useItemCountLabelHack : 1;
		/// \brief <em>Holds the \c IDataObject to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms688421.aspx">IDataObject</a>,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		IDataObject* drop_pDataObject;
		/// \brief <em>Holds the mouse position to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		POINT drop_mousePosition;
		/// \brief <em>Holds the drop effect to pass to \c IDropTargetHelper::Drop in \c FinishOLEDragDrop</em>
		///
		/// \sa FinishOLEDragDrop, Drop,
		///     <a href="https://msdn.microsoft.com/en-us/library/bb762027.aspx">IDropTargetHelper::Drop</a>
		DWORD drop_effect;
		//@}
		//////////////////////////////////////////////////////////////////////

		DragDropStatus()
		{
			pActiveDataObject = NULL;
			pSourceDataObject = NULL;
			pDropTargetHelper = NULL;
			draggingMouseButton = 0;

			pDraggedButtons = NULL;
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
			lastDropTarget = -1;
			lastClickedButton = -1;
			drop_pDataObject = NULL;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
				pDropTargetHelper = NULL;
			}
			ATLASSERT(!pDraggedButtons);
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			if(this->pDraggedButtons) {
				this->pDraggedButtons->Release();
				this->pDraggedButtons = NULL;
			}
			if(hDragImageList && autoDestroyImgLst) {
				ImageList_Destroy(hDragImageList);
			}
			hDragImageList = NULL;
			autoDestroyImgLst = FALSE;
			dragImageIsHidden = 1;
			lastDropTarget = -1;
			lastClickedButton = -1;

			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			if(this->pSourceDataObject) {
				this->pSourceDataObject = NULL;
			}
			draggingMouseButton = 0;
			useItemCountLabelHack = FALSE;
			drop_pDataObject = NULL;
		}

		/// \brief <em>Decrements the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, HideDragImage, IsDragImageVisible
		void ShowDragImage(BOOL commonDragDropOnly)
		{
			if(hDragImageList) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					ImageList_DragShowNolock(TRUE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				--dragImageIsHidden;
				if(dragImageIsHidden == 0) {
					pDropTargetHelper->Show(TRUE);
				}
			}
		}

		/// \brief <em>Increments the \c dragImageIsHidden flag</em>
		///
		/// \param[in] commonDragDropOnly If \c TRUE, the method does nothing if we're within an OLE
		///            drag'n'drop operation.
		///
		/// \sa dragImageIsHidden, ShowDragImage, IsDragImageVisible
		void HideDragImage(BOOL commonDragDropOnly)
		{
			if(hDragImageList) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					ImageList_DragShowNolock(FALSE);
				}
			} else if(pDropTargetHelper && !commonDragDropOnly) {
				++dragImageIsHidden;
				if(dragImageIsHidden == 1) {
					pDropTargetHelper->Show(FALSE);
				}
			}
		}

		/// \brief <em>Retrieves whether we're currently displaying a drag image</em>
		///
		/// \return \c TRUE, if we're displaying a drag image; otherwise \c FALSE.
		///
		/// \sa dragImageIsHidden, ShowDragImage, HideDragImage
		BOOL IsDragImageVisible(void)
		{
			return (dragImageIsHidden == 0);
		}

		/// \brief <em>Retrieves whether we're in drag'n'drop mode</em>
		///
		/// \return \c TRUE if we're in drag'n'drop mode; otherwise \c FALSE.
		BOOL IsDragging(void)
		{
			return (pDraggedButtons != NULL);
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			this->lastDropTarget = -1;
			this->lastClickedButton = -1;
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			this->lastDropTarget = -1;
			this->lastClickedButton = -1;
		}
	} dragDropStatus;
	///@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds data and flags related to menu mode</em>
	struct MenuModeState
	{
		/// \brief <em>Holds the handle of the window from which we stole input focus</em>
		HWND hFocusedWindow;
		/// \brief <em>Holds the zero-based index of the tool bar button whose popup menu currently is being displayed</em>
		int activePopupMenuButtonIndex;
		/// \brief <em>Holds the zero-based index of the tool bar button whose popup menu shall be displayed next</em>
		int nextPopupMenuButtonIndex;
		/// \brief <em>If \c TRUE, we called \c AddMessageHookClient</em>
		UINT attachedToMessageHook : 1;
		/// \brief <em>If \c TRUE, the control's parent window currently is active</em>
		UINT parentWindowIsActive : 1;
		/// \brief <em>If \c TRUE, we are currently in the progress of displaying a popup menu</em>
		UINT displayingPopupMenu : 1;
		/// \brief <em>If \c TRUE, we are currently displaying a popup menu</em>
		UINT menuIsActive : 1;
		/// \brief <em>If \c TRUE, the [ESC] key has been pressed</em>
		UINT escapePressed : 1;
		/// \brief <em>If \c TRUE, the currently selected menu item has a sub menu</em>
		UINT selectedMenuItemHasSubMenu : 1;
		/// \brief <em>If \c TRUE, the system is configured to display keyboard cues on keyboard input only</em>
		UINT useContextSensitiveKeyboardCues : 1;
		/// \brief <em>If \c TRUE, keyboard cues currently should be displayed</em>
		UINT displayKeyboardCues : 1;
		/// \brief <em>If \c TRUE, keyboard cues may be displayed</em>
		UINT allowDisplayingKeyboardCues : 1;
		/// \brief <em>If \c TRUE, the control will skip posting a \c WM_KEYDOWN message for \c VK_DOWN once</em>
		UINT skipPostingKeyDown : 1;
		/// \brief <em>If \c TRUE, the last input was keyboard input</em>
		UINT keyboardInput : 1;
		/// \brief <em>If \c TRUE, \c OnParentSysCommand will handle certain keys differently</em>
		///
		/// \sa OnParentSysCommand, OnHookSysKeyDown
		UINT skipSysCommandMessage : 1;
		/// \brief <em>If \c TRUE, the currently active MDI child window is maximized; otherwise not</em>
		UINT mdiChildIsMaximized : 1;
		/// \brief <em>A stack of the menu windows created by us</em>
		CSimpleStack<HWND> menuWindowStack;
		/// \brief <em>Holds the position of the mouse cursor during the last call of \c OnHookMouseMove</em>
		///
		/// \sa OnHookMouseMove
		POINT lastMouseMovePosition;
		/// \brief <em>The currently pressed system key</em>
		///
		/// \sa OnHookSysKeyDown
		UINT pressedSysKey;
		/// \brief <em>Holds the handle of the window that has been the target of the last forwarded message</em>
		HWND hWndForwardedMessage;

		MenuModeState()
		{
			hFocusedWindow = NULL;
			activePopupMenuButtonIndex = -1;
			nextPopupMenuButtonIndex = -1;
			attachedToMessageHook = FALSE;
			parentWindowIsActive = TRUE;
			menuIsActive = FALSE;
			displayingPopupMenu = FALSE;
			escapePressed = FALSE;
			selectedMenuItemHasSubMenu = FALSE;
			useContextSensitiveKeyboardCues = FALSE;
			displayKeyboardCues = FALSE;
			allowDisplayingKeyboardCues = TRUE;
			skipPostingKeyDown = FALSE;
			keyboardInput = FALSE;
			skipSysCommandMessage = FALSE;
			mdiChildIsMaximized = FALSE;
			lastMouseMovePosition.x = -1;
			lastMouseMovePosition.y = -1;
			pressedSysKey = 0;
			hWndForwardedMessage = NULL;
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			hFocusedWindow = NULL;
			activePopupMenuButtonIndex = -1;
			nextPopupMenuButtonIndex = -1;
			attachedToMessageHook = FALSE;
			parentWindowIsActive = TRUE;
			menuIsActive = FALSE;
			displayingPopupMenu = FALSE;
			escapePressed = FALSE;
			selectedMenuItemHasSubMenu = FALSE;
			useContextSensitiveKeyboardCues = FALSE;
			displayKeyboardCues = FALSE;
			allowDisplayingKeyboardCues = TRUE;
			skipPostingKeyDown = FALSE;
			keyboardInput = FALSE;
			skipSysCommandMessage = FALSE;
			mdiChildIsMaximized = FALSE;
			lastMouseMovePosition.x = -1;
			lastMouseMovePosition.y = -1;
			pressedSysKey = 0;
			hWndForwardedMessage = NULL;
		}
	} menuModeState;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;
		/// \brief <em>The ID of the timer that is used to auto-click buttons during drag'n'drop</em>
		static const UINT_PTR ID_DRAGCLICK = 14;
		/// \brief <em>The ID of the timer that is used to preselect an item in the chevron popup menu</em>
		static const UINT_PTR ID_PRESELECTCHEVRONMENUITEM = 15;

		/// \brief <em>The interval of the timer that is used to redraw the control window after recreation</em>
		static const UINT INT_REDRAW = 10;
		/// \brief <em>The interval of the timer that is used to preselect an item in the chevron popup menu</em>
		static const UINT INT_PRESELECTCHEVRONMENUITEM = 400;
	} timers;

	//////////////////////////////////////////////////////////////////////
	/// \name Version information
	///
	//@{
	/// \brief <em>Retrieves whether we're using at least version 6.10 of comctl32.dll</em>
	///
	/// \return \c TRUE if we're using comctl32.dll version 6.10 or higher; otherwise \c FALSE.
	BOOL IsComctl32Version610OrNewer(void);
	//@}
	//////////////////////////////////////////////////////////////////////

private:
};     // ToolBar

OBJECT_ENTRY_AUTO(__uuidof(ToolBar), ToolBar)