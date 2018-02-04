//////////////////////////////////////////////////////////////////////
/// \class ReBar
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Superclasses \c ReBarWindow32</em>
///
/// This class superclasses \c ReBarWindow32 and makes it accessible by COM.
///
/// \todo Move the OLE drag'n'drop flags into their own struct?
/// \todo Verify documentation of \c PreTranslateAccelerator.
///
/// \if UNICODE
///   \sa TBarCtlsLibU::IReBar
/// \else
///   \sa TBarCtlsLibA::IReBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "UndocComctl32.h"
#include "CWindowEx2.h"
#include "_IReBarEvents_CP.h"
#include "ICategorizeProperties.h"
#include "ICreditsProvider.h"
#include "helpers.h"
#include "EnumOLEVERB.h"
#include "PropertyNotifySinkImpl.h"
#include "AboutDlg.h"
#include "ReBarBand.h"
#include "VirtualReBarBand.h"
#include "ReBarBands.h"
#include "CommonProperties.h"
#include "TargetOLEDataObject.h"


class ATL_NO_VTABLE ReBar : 
    public CComObjectRootEx<CComSingleThreadModel>,
    #ifdef UNICODE
    	public IDispatchImpl<IReBar, &IID_IReBar, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IReBar, &IID_IReBar, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPersistStreamInitImpl<ReBar>,
    public IOleControlImpl<ReBar>,
    public IOleObjectImpl<ReBar>,
    public IOleInPlaceActiveObjectImpl<ReBar>,
    public IViewObjectExImpl<ReBar>,
    public IOleInPlaceObjectWindowlessImpl<ReBar>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ReBar>,
    public Proxy_IReBarEvents<ReBar>,
    public IPersistStorageImpl<ReBar>,
    public IPersistPropertyBagImpl<ReBar>,
    public ISpecifyPropertyPages,
    public IQuickActivateImpl<ReBar>,
    #ifdef UNICODE
    	public IProvideClassInfo2Impl<&CLSID_ReBar, &__uuidof(_IReBarEvents), &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IProvideClassInfo2Impl<&CLSID_ReBar, &__uuidof(_IReBarEvents), &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IPropertyNotifySinkCP<ReBar>,
    public CComCoClass<ReBar, &CLSID_ReBar>,
    public CComControl<ReBar>,
    public IPerPropertyBrowsingImpl<ReBar>,
    public IDropTarget,
    public ICategorizeProperties,
    public ICreditsProvider
{
	friend class ReBarBand;
	friend class ReBarBands;
	friend class VirtualReBarBand;

public:
	/// \brief <em>The \c MDIClient window that we subclass for the \c ReplaceMDIFrameMenu property</em>
	///
	/// \sa get_ReplaceMDIFrameMenu, mdiFrame
	CContainedWindow mdiClient;
	/// \brief <em>The top-level window that we subclass for the \c ReplaceMDIFrameMenu property</em>
	///
	/// \sa get_ReplaceMDIFrameMenu, mdiClient
	CContainedWindow mdiFrame;
	/// \brief <em>The window contained by the rebar band that replaces the MDI frame window's menu bar</em>
	///
	/// \sa get_ReplaceMDIFrameMenu, get_MDIFrameMenuBand
	CContainedWindow menuBandWindow;
	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	ReBar();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_OLEMISC_STATUS(OLEMISC_ACTIVATEWHENVISIBLE | OLEMISC_ALIGNABLE | OLEMISC_CANTLINKINSIDE | OLEMISC_INSIDEOUT | OLEMISC_RECOMPOSEONRESIZE | OLEMISC_SETCLIENTSITEFIRST)
		DECLARE_REGISTRY_RESOURCEID(IDR_REBAR)

		#ifdef UNICODE
			DECLARE_WND_SUPERCLASS(TEXT("ReBarU"), REBARCLASSNAMEW)
		#else
			DECLARE_WND_SUPERCLASS(TEXT("ReBarA"), REBARCLASSNAMEA)
		#endif

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		// we have a solid background and draw the entire rectangle
		DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

		BEGIN_COM_MAP(ReBar)
			COM_INTERFACE_ENTRY(IReBar)
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
		END_COM_MAP()

		BEGIN_PROP_MAP(ReBar)
			// NOTE: Don't forget to update Load and Save! This is for property bags only, not for streams!
			PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
			PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
			PROP_ENTRY_TYPE("AllowBandReordering", DISPID_RB_ALLOWBANDREORDERING, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Appearance", DISPID_RB_APPEARANCE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("AutoUpdateLayout", DISPID_RB_AUTOUPDATELAYOUT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("BackColor", DISPID_RB_BACKCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("BorderStyle", DISPID_RB_BORDERSTYLE, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisabledEvents", DISPID_RB_DISABLEDEVENTS, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("DisplayBandSeparators", DISPID_RB_DISPLAYBANDSEPARATORS, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DisplaySplitter", DISPID_RB_DISPLAYSPLITTER, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("DontRedraw", DISPID_RB_DONTREDRAW, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Enabled", DISPID_RB_ENABLED, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("FixedBandHeight", DISPID_RB_FIXEDBANDHEIGHT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("Font", DISPID_RB_FONT, CLSID_StockFontPage, VT_DISPATCH)
			PROP_ENTRY_TYPE("ForeColor", DISPID_RB_FORECOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("HighlightColor", DISPID_RB_HIGHLIGHTCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("HoverTime", DISPID_RB_HOVERTIME, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("MouseIcon", DISPID_RB_MOUSEICON, CLSID_StockPicturePage, VT_DISPATCH)
			PROP_ENTRY_TYPE("MousePointer", DISPID_RB_MOUSEPOINTER, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("Orientation", DISPID_RB_ORIENTATION, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ReplaceMDIFrameMenu", DISPID_RB_REPLACEMDIFRAMEMENU, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RegisterForOLEDragDrop", DISPID_RB_REGISTERFOROLEDRAGDROP, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("RightToLeft", DISPID_RB_RIGHTTOLEFT, CLSID_NULL, VT_I4)
			PROP_ENTRY_TYPE("ShadowColor", DISPID_RB_SHADOWCOLOR, CLSID_StockColorPage, VT_I4)
			PROP_ENTRY_TYPE("SupportOLEDragImages", DISPID_RB_SUPPORTOLEDRAGIMAGES, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("ToggleOnDoubleClick", DISPID_RB_TOGGLEONDOUBLECLICK, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("UseSystemFont", DISPID_RB_USESYSTEMFONT, CLSID_NULL, VT_BOOL)
			PROP_ENTRY_TYPE("VerticalSizingGripsOnVerticalOrientation", DISPID_RB_VERTICALSIZINGGRIPSONVERTICALORIENTATION, CLSID_NULL, VT_BOOL)
		END_PROP_MAP()

		BEGIN_CONNECTION_POINT_MAP(ReBar)
			CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
			CONNECTION_POINT_ENTRY(__uuidof(_IReBarEvents))
		END_CONNECTION_POINT_MAP()

		BEGIN_MSG_MAP(ReBar)
			MESSAGE_HANDLER(WM_CONTEXTMENU, OnContextMenu)
			MESSAGE_HANDLER(WM_CREATE, OnCreate)
			MESSAGE_HANDLER(WM_DESTROY, OnDestroy)
			MESSAGE_HANDLER(WM_DRAWITEM, OnMenuMessage)
			MESSAGE_HANDLER(WM_FORWARDMSG, OnForwardMsg)
			MESSAGE_HANDLER(WM_INITMENUPOPUP, OnMenuMessage)
			MESSAGE_HANDLER(WM_KILLFOCUS, OnKillFocus)
			MESSAGE_HANDLER(WM_LBUTTONDBLCLK, OnLButtonDblClk)
			MESSAGE_HANDLER(WM_LBUTTONDOWN, OnLButtonDown)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnLButtonUp)
			MESSAGE_HANDLER(WM_MBUTTONDBLCLK, OnMButtonDblClk)
			MESSAGE_HANDLER(WM_MBUTTONDOWN, OnMButtonDown)
			MESSAGE_HANDLER(WM_MBUTTONUP, OnMButtonUp)
			MESSAGE_HANDLER(WM_MEASUREITEM, OnMenuMessage)
			MESSAGE_HANDLER(WM_MENUCHAR, OnMenuMessage)
			MESSAGE_HANDLER(WM_MENUSELECT, OnMenuSelect)
			MESSAGE_HANDLER(WM_MOUSEACTIVATE, OnMouseActivate)
			MESSAGE_HANDLER(WM_MOUSEHOVER, OnMouseHover)
			MESSAGE_HANDLER(WM_MOUSELEAVE, OnMouseLeave)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnMouseMove)
			MESSAGE_HANDLER(WM_PAINT, OnPaint)
			MESSAGE_HANDLER(WM_RBUTTONDBLCLK, OnRButtonDblClk)
			MESSAGE_HANDLER(WM_RBUTTONDOWN, OnRButtonDown)
			MESSAGE_HANDLER(WM_RBUTTONUP, OnRButtonUp)
			MESSAGE_HANDLER(WM_SETCURSOR, OnSetCursor)
			MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
			MESSAGE_HANDLER(WM_SETFONT, OnSetFont)
			MESSAGE_HANDLER(WM_SETREDRAW, OnSetRedraw)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			MESSAGE_HANDLER(WM_THEMECHANGED, OnThemeChanged)
			MESSAGE_HANDLER(WM_TIMER, OnTimer)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnWindowPosChanged)
			MESSAGE_HANDLER(WM_XBUTTONDBLCLK, OnXButtonDblClk)
			MESSAGE_HANDLER(WM_XBUTTONDOWN, OnXButtonDown)
			MESSAGE_HANDLER(WM_XBUTTONUP, OnXButtonUp)
			MESSAGE_HANDLER(GetBarMessage(), OnGetRebar)
			MESSAGE_HANDLER(GetAutoPopupMessage(), OnAutoPopup)

			MESSAGE_HANDLER(OCM_NOTIFY, OnReflectedNotify)
			MESSAGE_HANDLER(OCM__BASE + WM_NOTIFYFORMAT, OnReflectedNotifyFormat)

			MESSAGE_HANDLER(RB_DELETEBAND, OnDeleteBand)
			MESSAGE_HANDLER(RB_GETBANDINFOA, OnGetBandInfo)
			MESSAGE_HANDLER(RB_GETBANDINFOW, OnGetBandInfo)
			MESSAGE_HANDLER(RB_INSERTBANDA, OnInsertBand)
			MESSAGE_HANDLER(RB_INSERTBANDW, OnInsertBand)
			MESSAGE_HANDLER(RB_SETBANDINFOA, OnSetBandInfo)
			MESSAGE_HANDLER(RB_SETBANDINFOW, OnSetBandInfo)

			REFLECTED_NOTIFY_CODE_HANDLER(NM_NCHITTEST, OnNCHitTestNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(NM_RELEASEDCAPTURE, OnReleasedCaptureNotification)

			REFLECTED_NOTIFY_CODE_HANDLER(RBN_AUTOBREAK, OnAutoBreakNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_AUTOSIZE, OnAutoSizeNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_BEGINDRAG, OnBeginDragNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_CHEVRONPUSHED, OnChevronPushedNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_CHILDSIZE, OnChildSizeNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_DELETINGBAND, OnDeletingBandNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_ENDDRAG, OnEndDragNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_GETOBJECT, OnGetObjectNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_HEIGHTCHANGE, OnHeightChangeNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_LAYOUTCHANGED, OnLayoutChangedNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_MINMAX, OnMinMaxNotification)
			REFLECTED_NOTIFY_CODE_HANDLER(RBN_SPLITTERDRAG, OnSplitterDragNotification)

			CHAIN_MSG_MAP(CComControl<ReBar>)

			ALT_MSG_MAP(1)     // mdiClient
			MESSAGE_HANDLER(WM_MDISETMENU, OnMDIClientMDISetMenu)

			ALT_MSG_MAP(2)     // mdiFrame
			MESSAGE_HANDLER(WM_ACTIVATE, OnMDIFrameActivate)
			MESSAGE_HANDLER(WM_SETTINGCHANGE, OnSettingChange)
			MESSAGE_HANDLER(GetBarMessage(), OnGetRebar)

			ALT_MSG_MAP(3)     // menuBandWindow
			MESSAGE_HANDLER(WM_CAPTURECHANGED, OnMenuBandChildCaptureChanged)
			MESSAGE_HANDLER(WM_LBUTTONUP, OnMenuBandChildLButtonUp)
			MESSAGE_HANDLER(WM_MOUSEMOVE, OnMenuBandChildMouseMove)
			MESSAGE_HANDLER(WM_NCCALCSIZE, OnMenuBandChildNCCalcSize)
			MESSAGE_HANDLER(WM_NCHITTEST, OnMenuBandChildNCHitTest)
			MESSAGE_HANDLER(WM_NCLBUTTONDBLCLK, OnMenuBandChildLButtonDblClk)
			MESSAGE_HANDLER(WM_NCLBUTTONDOWN, OnMenuBandChildNCLButtonDown)
			MESSAGE_HANDLER(WM_NCPAINT, OnMenuBandChildNCPaint)
			MESSAGE_HANDLER(WM_SIZE, OnMenuBandChildSize)

			ALT_MSG_MAP(4)     // message hook
			MESSAGE_RANGE_HANDLER(0, 0xFFFF, OnAllHookMessages)
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
	/// \name Implementation of IReBar
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AllowBandReordering property</em>
	///
	/// Retrieves whether the user can bring the bands into a different order using drag'n'drop. If set to
	/// \c VARIANT_TRUE, reordering is possible; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AllowBandReordering, ReBarBand::get_SizingGripVisibility
	virtual HRESULT STDMETHODCALLTYPE get_AllowBandReordering(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AllowBandReordering property</em>
	///
	/// Sets whether the user can bring the bands into a different order using drag'n'drop. If set to
	/// \c VARIANT_TRUE, reordering is possible; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AllowBandReordering, ReBarBand::put_SizingGripVisibility
	virtual HRESULT STDMETHODCALLTYPE put_AllowBandReordering(VARIANT_BOOL newValue);
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
	/// \brief <em>Retrieves the current setting of the \c AutoUpdateLayout property</em>
	///
	/// Retrieves whether the bands are layouted automatically if the control's size or position changes. If
	/// set to \c VARIANT_TRUE, the layout is changed automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AutoUpdateLayout, Raise_AutoSized, Raise_LayoutChanged
	virtual HRESULT STDMETHODCALLTYPE get_AutoUpdateLayout(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutoUpdateLayout property</em>
	///
	/// Sets whether the bands are layouted automatically if the control's size or position changes. If
	/// set to \c VARIANT_TRUE, the layout is changed automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AutoUpdateLayout, Raise_AutoSized, Raise_LayoutChanged
	virtual HRESULT STDMETHODCALLTYPE put_AutoUpdateLayout(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the control's background color. If set to -1, the system's default is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa put_BackColor, get_ForeColor, get_HighlightColor, get_ShadowColor, ReBarBand::get_BackColor,
	///     ReBarBand::get_hBackgroundBitmap
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the control's background color. If set to -1, the system's default is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa get_BackColor, put_ForeColor, put_HighlightColor, put_ShadowColor, ReBarBand::put_BackColor,
	///     ReBarBand::put_hBackgroundBitmap
	virtual HRESULT STDMETHODCALLTYPE put_BackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c Bands property</em>
	///
	/// Retrieves a collection object wrapping the control's bands.
	///
	/// \param[out] ppBands Receives the collection object's \c IReBarBands implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ReBarBands
	virtual HRESULT STDMETHODCALLTYPE get_Bands(IReBarBands** ppBands);
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
	/// \brief <em>Retrieves the current setting of the \c ControlHeight property</em>
	///
	/// Retrieves the control's height in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ReBarBand::get_RowHeight
	virtual HRESULT STDMETHODCALLTYPE get_ControlHeight(OLE_YSIZE_PIXELS* pValue);
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
	/// \brief <em>Retrieves the current setting of the \c DisplayBandSeparators property</em>
	///
	/// Retrieves whether adjacent bands are separated by a thin border. If set to \c VARIANT_TRUE, a
	/// border is drawn between adjacent bands; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_DisplayBandSeparators, get_DisplaySplitter, get_HighlightColor, get_ShadowColor,
	///     ReBarBand::get_AddMarginsAroundChild
	virtual HRESULT STDMETHODCALLTYPE get_DisplayBandSeparators(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayBandSeparators property</em>
	///
	/// Sets whether adjacent bands are separated by a thin border. If set to \c VARIANT_TRUE, a
	/// border is drawn between adjacent bands; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_DisplayBandSeparators, put_DisplaySplitter, put_HighlightColor, put_ShadowColor,
	///     ReBarBand::put_AddMarginsAroundChild
	virtual HRESULT STDMETHODCALLTYPE put_DisplayBandSeparators(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplaySplitter property</em>
	///
	/// Retrieves whether a splitter is displayed at the bottom of the control. If the mouse is moved over
	/// this splitter, the cursor is changed to a sizing cursor. Dragging the splitter doesn't resize the
	/// control, but raises the \c DraggingSplitter event. If set to \c VARIANT_TRUE, the splitter is
	/// displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa put_DisplaySplitter, get_DisplayBandSeparators, Raise_DraggingSplitter
	virtual HRESULT STDMETHODCALLTYPE get_DisplaySplitter(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplaySplitter property</em>
	///
	/// Sets whether a splitter is displayed at the bottom of the control. If the mouse is moved over
	/// this splitter, the cursor is changed to a sizing cursor. Dragging the splitter doesn't resize the
	/// control, but raises the \c DraggingSplitter event. If set to \c VARIANT_TRUE, the splitter is
	/// displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_DisplaySplitter, put_DisplayBandSeparators, Raise_DraggingSplitter
	virtual HRESULT STDMETHODCALLTYPE put_DisplaySplitter(VARIANT_BOOL newValue);
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
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the control is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled
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
	/// \sa get_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FixedBandHeight property</em>
	///
	/// Retrieves whether the rebar bands are sized using each individual band's minimum height or using the
	/// tallest band's height. If set to \c VARIANT_TRUE, the tallest band's height is used for all bands;
	/// otherwise each band is sized individually using its minimum height.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FixedBandHeight, ReBarBand::get_CurrentHeight
	virtual HRESULT STDMETHODCALLTYPE get_FixedBandHeight(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c FixedBandHeight property</em>
	///
	/// Sets whether the rebar bands are sized using each individual band's minimum height or using the
	/// tallest band's height. If set to \c VARIANT_TRUE, the tallest band's height is used for all bands;
	/// otherwise each band is sized individually using its minimum height.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FixedBandHeight, ReBarBand::put_CurrentHeight
	virtual HRESULT STDMETHODCALLTYPE put_FixedBandHeight(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Font property</em>
	///
	/// Retrieves the control's font. It's used to draw the band captions.
	///
	/// \param[out] ppFont Receives the font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa put_Font, putref_Font, get_UseSystemFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE get_Font(IFontDisp** ppFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the band captions.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The object identified by \c pNewFont is cloned.\n
	///          This property isn't supported for themed rebars.
	///
	/// \sa get_Font, putref_Font, put_UseSystemFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE put_Font(IFontDisp* pNewFont);
	/// \brief <em>Sets the \c Font property</em>
	///
	/// Sets the control's font. It's used to draw the band captions.
	///
	/// \param[in] pNewFont The new font object's \c IFontDisp implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa get_Font, put_Font, put_UseSystemFont,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms692695.aspx">IFontDisp</a>
	virtual HRESULT STDMETHODCALLTYPE putref_Font(IFontDisp* pNewFont);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the control's text color. If set to -1, the system's default is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa put_ForeColor, get_BackColor, get_HighlightColor, get_ShadowColor, ReBarBand::get_ForeColor
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the control's text color. If set to -1, the system's default is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa get_ForeColor, put_BackColor, put_HighlightColor, put_ShadowColor, ReBarBand::put_ForeColor
	virtual HRESULT STDMETHODCALLTYPE put_ForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c HighlightColor property</em>
	///
	/// Retrieves the color used by the control to draw highlighted parts of 3D elements like band borders
	/// and sizing grips. If set to -1, the default color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa put_HighlightColor, get_ShadowColor, get_BackColor, get_ForeColor, get_DisplayBandSeparators,
	///     ReBarBand::get_SizingGripVisibility
	virtual HRESULT STDMETHODCALLTYPE get_HighlightColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c HighlightColor property</em>
	///
	/// Sets the color used by the control to draw highlighted parts of 3D elements like band borders
	/// and sizing grips. If set to -1, the default color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa get_HighlightColor, put_ShadowColor, put_BackColor, put_ForeColor, put_DisplayBandSeparators,
	///     ReBarBand::put_SizingGripVisibility
	virtual HRESULT STDMETHODCALLTYPE put_HighlightColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c hImageList property</em>
	///
	/// Retrieves the handle of the specified imagelist.
	///
	/// \param[in] imageList The imageList to retrieve. Some of the values defined by the
	///            \c ImageListConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa put_hImageList, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa put_hImageList, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_hImageList(ImageListConstants imageList, OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hImageList property</em>
	///
	/// Sets the handle of the specified imagelist.
	///
	/// \param[in] imageList The imageList to set. Some of the values defined by the \c ImageListConstants
	///            enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set image list does NOT get destroyed automatically.
	///
	/// \if UNICODE
	///   \sa get_hImageList, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa get_hImageList, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_hImageList(ImageListConstants imageList, OLE_HANDLE newValue);
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
	/// \brief <em>Retrieves the current setting of the \c hPalette property</em>
	///
	/// Retrieves the palette used by the control when drawing itself.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hPalette, ReBarBand::get_hBackgroundBitmap
	virtual HRESULT STDMETHODCALLTYPE get_hPalette(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hPalette property</em>
	///
	/// Sets the palette used by the control when drawing itself.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The previously set palette does NOT get destroyed automatically.
	///
	/// \sa get_hPalette, ReBarBand::put_hBackgroundBitmap
	virtual HRESULT STDMETHODCALLTYPE put_hPalette(OLE_HANDLE newValue);
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
	/// \sa get_hWndNotificationReceiver, Raise_RecreatedControlWindow, Raise_DestroyedControlWindow
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndNotificationReceiver property</em>
	///
	/// Retrieves the handle of the window that the control sends notifications to.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hWndNotificationReceiver, get_hWnd
	virtual HRESULT STDMETHODCALLTYPE get_hWndNotificationReceiver(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hWndNotificationReceiver property</em>
	///
	/// Sets the handle of the window that the control sends notifications to.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_hWndNotificationReceiver, get_hWnd
	virtual HRESULT STDMETHODCALLTYPE put_hWndNotificationReceiver(OLE_HANDLE newValue);
	// \brief <em>Retrieves the current setting of the \c hWndToolTip property</em>
	//
	// Retrieves the tooltip control's window handle.
	//
	// \param[out] pValue The property's current setting.
	//
	// \return An \c HRESULT error code.
	//
	// \sa put_hWndToolTip, get_ShowToolTips
	//virtual HRESULT STDMETHODCALLTYPE get_hWndToolTip(OLE_HANDLE* pValue);
	// \brief <em>Sets the \c hWndToolTip property</em>
	//
	// Sets the tooltip control to use.
	//
	// \param[in] newValue The setting to apply.
	//
	// \return An \c HRESULT error code.
	//
	// \remarks The previous tooltip window does NOT get auto-destroyed.
	//
	// \sa get_hWndToolTip, put_ShowToolTips
	//virtual HRESULT STDMETHODCALLTYPE put_hWndToolTip(OLE_HANDLE newValue);
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
	/// \brief <em>Retrieves the current setting of the \c MDIFrameMenuBand property</em>
	///
	/// Retrieves the rebar band that replaces the parent window's menu bar if the parent window is a MDI
	/// frame window and the \c ReplaceMDIFrameMenu property is set to \c rmfmFullReplace.
	///
	/// \param[out] ppMenuBand Receives the menu band's \c IReBarBand implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa putref_MDIFrameMenuBand, get_ReplaceMDIFrameMenu, get_Bands, ReBarBand,
	///       TBarCtlsLibU::ReplaceMDIFrameMenuConstants
	/// \else
	///   \sa putref_MDIFrameMenuBand, get_ReplaceMDIFrameMenu, get_Bands, ReBarBand,
	///       TBarCtlsLibA::ReplaceMDIFrameMenuConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_MDIFrameMenuBand(IReBarBand** ppMenuBand);
	/// \brief <em>Sets the \c MDIFrameMenuBand property</em>
	///
	/// Sets the rebar band that replaces the parent window's menu bar if the parent window is a MDI
	/// frame window and the \c ReplaceMDIFrameMenu property is set to \c rmfmFullReplace.
	///
	/// \param[in] pNewMenuBand The new menu band's \c IReBarBand implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_MDIFrameMenuBand, put_ReplaceMDIFrameMenu, get_Bands, ReBarBand,
	///       TBarCtlsLibU::ReplaceMDIFrameMenuConstants
	/// \else
	///   \sa get_MDIFrameMenuBand, put_ReplaceMDIFrameMenu, get_Bands, ReBarBand,
	///       TBarCtlsLibA::ReplaceMDIFrameMenuConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE putref_MDIFrameMenuBand(IReBarBand* pNewMenuBand);
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
	/// \brief <em>Retrieves the current setting of the \c NativeDropTarget property</em>
	///
	/// Retrieves the native rebar control's \c IDropTarget implementation.
	///
	/// \param[out] ppValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_RegisterForOLEDragDrop
	virtual HRESULT STDMETHODCALLTYPE get_NativeDropTarget(LPUNKNOWN* ppValue);
	/// \brief <em>Retrieves the current setting of the \c NumberOfRows property</em>
	///
	/// Retrieves the number of band rows.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ReBarBand::get_RowHeight
	virtual HRESULT STDMETHODCALLTYPE get_NumberOfRows(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Orientation property</em>
	///
	/// Retrieves the direction, in which the control displays bands. Any of the values defined by the
	/// \c OrientationConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Orientation, get_VerticalSizingGripsOnVerticalOrientation,
	///       TBarCtlsLibU::OrientationConstants
	/// \else
	///   \sa put_Orientation, get_VerticalSizingGripsOnVerticalOrientation,
	///       TBarCtlsLibA::OrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Orientation(OrientationConstants* pValue);
	/// \brief <em>Sets the \c Orientation property</em>
	///
	/// Sets the direction, in which the control displays bands. Any of the values defined by the
	/// \c OrientationConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Changing this property destroys and recreates the control window.
	///
	/// \if UNICODE
	///   \sa get_Orientation, put_VerticalSizingGripsOnVerticalOrientation,
	///       TBarCtlsLibU::OrientationConstants
	/// \else
	///   \sa get_Orientation, put_VerticalSizingGripsOnVerticalOrientation,
	///       TBarCtlsLibA::OrientationConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Orientation(OrientationConstants newValue);
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
	/// \brief <em>Retrieves the current setting of the \c ReplaceMDIFrameMenu property</em>
	///
	/// Retrieves whether the control replaces the parent window's menu bar if the parent window is a MDI
	/// frame window. Any of the values defined by the \c ReplaceMDIFrameMenuConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property cannot be set at runtime.
	///
	/// \if UNICODE
	///   \sa put_ReplaceMDIFrameMenu, get_MDIFrameMenuBand, TBarCtlsLibU::ReplaceMDIFrameMenuConstants
	/// \else
	///   \sa put_ReplaceMDIFrameMenu, get_MDIFrameMenuBand, TBarCtlsLibA::ReplaceMDIFrameMenuConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ReplaceMDIFrameMenu(ReplaceMDIFrameMenuConstants* pValue);
	/// \brief <em>Sets the \c ReplaceMDIFrameMenu property</em>
	///
	/// Sets whether the control replaces the parent window's menu bar if the parent window is a MDI
	/// frame window. Any of the values defined by the \c ReplaceMDIFrameMenuConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property cannot be set at runtime.
	///
	/// \if UNICODE
	///   \sa get_ReplaceMDIFrameMenu, putref_MDIFrameMenuBand, TBarCtlsLibU::ReplaceMDIFrameMenuConstants
	/// \else
	///   \sa get_ReplaceMDIFrameMenu, putref_MDIFrameMenuBand, TBarCtlsLibA::ReplaceMDIFrameMenuConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ReplaceMDIFrameMenu(ReplaceMDIFrameMenuConstants newValue);
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
	/// Retrieves the color used by the control to draw shadowed parts of 3D elements like band borders and
	/// sizing grips. If set to -1, the default color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa put_ShadowColor, get_HighlightColor, get_BackColor, get_ForeColor, get_DisplayBandSeparators,
	///     ReBarBand::get_SizingGripVisibility
	virtual HRESULT STDMETHODCALLTYPE get_ShadowColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ShadowColor property</em>
	///
	/// Sets the color used by the control to draw shadowed parts of 3D elements like band borders and
	/// sizing grips. If set to -1, the default color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa get_ShadowColor, put_HighlightColor, put_BackColor, put_ForeColor, put_DisplayBandSeparators,
	///     ReBarBand::put_SizingGripVisibility
	virtual HRESULT STDMETHODCALLTYPE put_ShadowColor(OLE_COLOR newValue);
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
	/// \sa put_SupportOLEDragImages, get_RegisterForOLEDragDrop, FinishOLEDragDrop,
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
	/// \sa get_SupportOLEDragImages, put_RegisterForOLEDragDrop, FinishOLEDragDrop,
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
	/// \brief <em>Retrieves the current setting of the \c ToggleOnDoubleClick property</em>
	///
	/// Retrieves whether a double-click is required to toggle a band's minimized or maximized state. If set
	/// to \c VARIANT_TRUE, a band is minimized or maximized if it is double-clicked; otherwise if it is
	/// single-clicked.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ToggleOnDoubleClick, ReBarBand::Maximize, ReBarBand::Minimize, Raise_DblClick,
	///     Raise_TogglingBand
	virtual HRESULT STDMETHODCALLTYPE get_ToggleOnDoubleClick(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ToggleOnDoubleClick property</em>
	///
	/// Sets whether a double-click is required to toggle a band's minimized or maximized state. If set
	/// to \c VARIANT_TRUE, a band is minimized or maximized if it is double-clicked; otherwise if it is
	/// single-clicked.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ToggleOnDoubleClick, ReBarBand::Maximize, ReBarBand::Minimize, Raise_DblClick,
	///     Raise_TogglingBand
	virtual HRESULT STDMETHODCALLTYPE put_ToggleOnDoubleClick(VARIANT_BOOL newValue);
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
	/// \remarks This property isn't supported for themed rebars.
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
	/// \remarks This property isn't supported for themed rebars.
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
	/// \brief <em>Retrieves the current setting of the \c VerticalSizingGripsOnVerticalOrientation property</em>
	///
	/// Retrieves whether sizing grips are drawn vertically if the \c Orientation property is set to
	/// \c oVertical. If set to \c VARIANT_TRUE, sizing grips are drawn vertically then; otherwise
	/// horizontally.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_VerticalSizingGripsOnVerticalOrientation, get_Orientation,
	///     ReBarBand::get_SizingGripVisibility
	virtual HRESULT STDMETHODCALLTYPE get_VerticalSizingGripsOnVerticalOrientation(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c VerticalSizingGripsOnVerticalOrientation property</em>
	///
	/// Sets whether sizing grips are drawn vertically if the \c Orientation property is set to
	/// \c oVertical. If set to \c VARIANT_TRUE, sizing grips are drawn vertically then; otherwise
	/// horizontally.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_VerticalSizingGripsOnVerticalOrientation, put_Orientation,
	///     ReBarBand::put_SizingGripVisibility
	virtual HRESULT STDMETHODCALLTYPE put_VerticalSizingGripsOnVerticalOrientation(VARIANT_BOOL newValue);

	/// \brief <em>Displays the control's credits</em>
	///
	/// Displays some information about this control and its author.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa AboutDlg
	virtual HRESULT STDMETHODCALLTYPE About(void);
	/// \brief <em>Starts dragging a band</em>
	///
	/// Puts the control into drag'n'drop mode.
	///
	/// \param[in] pBandToDrag The band to drag.
	/// \param[in] xMousePosition The x-coordinate (in pixels) of the mouse cursor's position relative to
	///            the control's upper-left corner.
	/// \param[in] yMousePosition The y-coordinate (in pixels) of the mouse cursor's position relative to
	///            the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If \c xMousePosition or \c yMousePosition is set to -1, the position, that the mouse cursor
	///          had on the last processed message, is used. If \c xMousePosition or \c yMousePosition is set
	///          to -2, the mouse cursor's position on the last \c MouseDown message is used.\n
	///          The \c BandBeginDrag event won't be raised.
	///
	/// \sa DragMoveBand, EndDragBand, Raise_BandBeginDrag
	virtual HRESULT STDMETHODCALLTYPE BeginDragBand(IReBarBand* pBandToDrag, LONG xMousePosition = -1, LONG yMousePosition = -1);
	/// \brief <em>Sets the drag position while dragging a band</em>
	///
	/// Sets the drag position is the control is in drag'n'drop mode.
	///
	/// \param[in] xMousePosition The x-coordinate (in pixels) of the mouse cursor's position relative to
	///            the control's upper-left corner.
	/// \param[in] yMousePosition The y-coordinate (in pixels) of the mouse cursor's position relative to
	///            the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If \c xMousePosition or \c yMousePosition is set to -1, the position, that the mouse cursor
	///          had on the last processed message, is used.
	///
	/// \sa BeginDragBand, EndDragBand
	virtual HRESULT STDMETHODCALLTYPE DragMoveBand(LONG xMousePosition = -1, LONG yMousePosition = -1);
	/// \brief <em>Ends drag'n'drop mode</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The \c BandEndDrag event won't be raised.
	///
	/// \sa BeginDragBand, EndDragBand, Raise_BandEndDrag
	virtual HRESULT STDMETHODCALLTYPE EndDragBand(void);
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
	/// \brief <em>Retrieves the margins around a band</em>
	///
	/// Retrieves the size (in pixels) of the margins around a band.
	///
	/// \param[out] pLeftMargin The width (in pixels) of the margin on the left of a band.
	/// \param[out] pTopMargin The height (in pixels) of the margin on the top of a band.
	/// \param[out] pRightMargin The width (in pixels) of the margin on the right of a band.
	/// \param[out] pBottomMargin The height (in pixels) of the margin on the bottom of a band.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \sa ReBarBand::get_AddMarginsAroundChild, ReBarBand::GetBorderSizes, ReBarBand::GetRectangle
	virtual HRESULT STDMETHODCALLTYPE GetMargins(OLE_XSIZE_PIXELS* pLeftMargin = NULL, OLE_YSIZE_PIXELS* pTopMargin = NULL, OLE_XSIZE_PIXELS* pRightMargin = NULL, OLE_YSIZE_PIXELS* pBottomMargin = NULL);
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
	/// \param[out] ppHitBand Receives the "hit" band's \c IReBarBand implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::HitTestConstants, HitTest
	/// \else
	///   \sa TBarCtlsLibA::HitTestConstants, HitTest
	/// \endif
	// \if UNICODE
	//   \sa get_BandBoundingBoxDefinition, TBarCtlsLibU::HitTestConstants, HitTest
	// \else
	//   \sa get_BandBoundingBoxDefinition, TBarCtlsLibA::HitTestConstants, HitTest
	// \endif
	virtual HRESULT STDMETHODCALLTYPE HitTest(OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants* pHitTestDetails, IReBarBand** ppHitBand);
	/// \brief <em>Loads the control's settings from the specified file</em>
	///
	/// \param[in] file The file to read from.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveSettingsToFile
	virtual HRESULT STDMETHODCALLTYPE LoadSettingsFromFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Advises the control to redraw itself</em>
	///
	/// \return An \c HRESULT error code.
	virtual HRESULT STDMETHODCALLTYPE Refresh(void);
	/// \brief <em>Saves the control's settings to the specified file</em>
	///
	/// \param[in] file The file to write to.
	/// \param[out] pSucceeded Will be set to \c VARIANT_TRUE on success and to \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadSettingsFromFile
	virtual HRESULT STDMETHODCALLTYPE SaveSettingsToFile(BSTR file, VARIANT_BOOL* pSucceeded);
	/// \brief <em>Resizes the control and changes its layout to best fit a given rectangle</em>
	///
	/// Resizes the control and reorganizes the bands so that the control's bounding rectangle matches the
	/// specified rectangle as good as possible.
	///
	/// \param[in] xLeft The x-coordinate (in pixels) of the rectangle's upper-left corner.
	/// \param[in] yTop The y-coordinate (in pixels) of the rectangle's upper-left corner.
	/// \param[in] xRight The x-coordinate (in pixels) of the rectangle's lower-right corner.
	/// \param[in] yBottom The y-coordinate (in pixels) of the rectangle's lower-right corner.
	/// \param[out] pLayoutChanged Will be set to \c VARIANT_TRUE if the layout has been changed and to
	///             \c VARIANT_FALSE otherwise.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Raise_ResizedControlWindow, Raise_LayoutChanged
	virtual HRESULT STDMETHODCALLTYPE SizeToRectangle(OLE_XPOS_PIXELS xLeft, OLE_YPOS_PIXELS yTop, OLE_XPOS_PIXELS xRight, OLE_YPOS_PIXELS yBottom, VARIANT_BOOL* pLayoutChanged);
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
	/// \brief <em>\c WM_CONTEXTMENU handler</em>
	///
	/// Will be called if the control's context menu should be displayed.
	/// We use this handler to raise the \c ContextMenu event.
	///
	/// \sa OnRButtonDown, Raise_ContextMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms647592.aspx">WM_CONTEXTMENU</a>
	LRESULT OnContextMenu(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
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
	LRESULT OnForwardMsg(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& /*wasHandled*/);
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
	LRESULT OnLButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnMButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645607.aspx">WM_LBUTTONDOWN</a>
	LRESULT OnLButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnMButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnLButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the control's client area using the middle mouse
	/// button.
	/// We use this handler to raise the \c MDblClick event.
	///
	/// \sa OnLButtonDblClk, OnRButtonDblClk, OnXButtonDblClk, Raise_MDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645609.aspx">WM_MBUTTONDBLCLK</a>
	LRESULT OnMButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnRButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645610.aspx">WM_MBUTTONDOWN</a>
	LRESULT OnMButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MBUTTONUP handler</em>
	///
	/// Will be called if the user releases the middle mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnRButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645611.aspx">WM_MBUTTONUP</a>
	LRESULT OnMButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MENUCHAR, \c WM_INITMENUPOPUP, \c WM_MEASUREITEM and \c WM_DRAWITEM handler</em>
	///
	/// Will be called when a owner-drawn menu item needs to be drawn or its size is required.
	/// We use this handler to raise the \c RawMenuMessage event.
	///
	/// \sa get_ReplaceMDIFrameMenu, Raise_RawMenuMessage,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775923.aspx">WM_DRAWITEM</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646347.aspx">WM_INITMENUPOPUP</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb775925.aspx">WM_MEASUREITEM</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646349.aspx">WM_MENUCHAR</a>
	LRESULT OnMenuMessage(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MENUSELECT handler</em>
	///
	/// Will be called if the user selects a menu item in a menu owned by us.
	/// We use this handler to support keyboard navigation in menu mode and to raise the \c SelectedMenuItem
	/// event.
	///
	/// \sa OnMenuMessage, get_ReplaceMDIFrameMenu, Raise_SelectedMenuItem,
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
	LRESULT OnMouseHover(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSELEAVE handler</em>
	///
	/// Will be called if the user moves the mouse cursor out of the control's client area.
	/// We use this handler to raise the \c MouseLeave event.
	///
	/// \sa Raise_MouseLeave,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645615.aspx">WM_MOUSELEAVE</a>
	LRESULT OnMouseLeave(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the control's
	/// client area.
	/// We use this handler to raise the \c MouseMove event.
	///
	/// \sa OnLButtonDown, OnLButtonUp, Raise_MouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMouseMove(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
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
	LRESULT OnRButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnXButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646242.aspx">WM_RBUTTONDOWN</a>
	LRESULT OnRButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_RBUTTONUP handler</em>
	///
	/// Will be called if the user releases the right mouse button while the mouse cursor is located over
	/// the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnXButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646243.aspx">WM_RBUTTONUP</a>
	LRESULT OnRButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SETCURSOR handler</em>
	///
	/// Will be called if the mouse cursor type is required that shall be used while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to set the mouse cursor type.
	///
	/// \sa get_MouseIcon, get_MousePointer,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms648382.aspx">WM_SETCURSOR</a>
	LRESULT OnSetCursor(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
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
	/// \sa OnAutoSizeNotification, Raise_ResizedControlWindow,
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
	LRESULT OnXButtonDblClk(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseDown event.
	///
	/// \sa OnLButtonDown, OnMButtonDown, OnRButtonDown, Raise_MouseDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646245.aspx">WM_XBUTTONDOWN</a>
	LRESULT OnXButtonDown(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_XBUTTONUP handler</em>
	///
	/// Will be called if the user releases one of the extended mouse buttons while the mouse cursor is
	/// located over the control's client area.
	/// We use this handler to raise the \c MouseUp event.
	///
	/// \sa OnLButtonUp, OnMButtonUp, OnRButtonUp, Raise_MouseUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646246.aspx">WM_XBUTTONUP</a>
	LRESULT OnXButtonUp(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c Bar_InternalGetInstanceMessage handler</em>
	///
	/// Will be called if the control is being asked for the handle of the associated \c ReBar control.
	/// We use this handler in the message hook to retrieve the associated \c ReBar window.
	///
	/// \sa GetMsgProc, GetBarMessage
	LRESULT OnGetRebar(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c ToolBar_InternalAutoPopupMessage handler</em>
	///
	/// Will be called if the control is being asked to display the active MDI child window's system menu.
	/// We use this handler to support switching between the MDI child window's system menu and the tool
	/// bar's drop-down menus.
	///
	/// \sa ToolBar::OnHookKeyDown, GetAutoPopupMessage
	LRESULT OnAutoPopup(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
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
	/// \brief <em>\c RB_DELETEBAND handler</em>
	///
	/// Will be called if a rebar band shall be removed from the control.
	/// We use this handler mainly to raise the \c RemovingBand and \c RemovedBand events.
	///
	/// \sa OnDeletingBandNotification, OnInsertBand, Raise_RemovingBand, Raise_RemovedBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774431.aspx">RB_DELETEBAND</a>
	LRESULT OnDeleteBand(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c RB_GETBANDINFO handler</em>
	///
	/// Will be called if the control is requested to retrieve some or all of a rebar band's attributes.
	/// We use this handler to support retrieving the chevron rectangle for older systems.
	///
	/// \sa OnSetBandInfo, ReBarBand::GetChevronRectangle,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774451.aspx">RB_GETBANDINFO</a>
	LRESULT OnGetBandInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c RB_INSERTBAND handler</em>
	///
	/// Will be called if a new rebar band shall be inserted into the control.
	/// We use this handler mainly to raise the \c InsertingBand and \c InsertedBand events.
	///
	/// \sa OnDeleteBand, Raise_InsertingBand, Raise_InsertedBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774498.aspx">RB_INSERTBAND</a>
	LRESULT OnInsertBand(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c RB_SETBANDINFO handler</em>
	///
	/// Will be called if the control is requested to update some or all of a rebar band's attributes.
	/// We use this handler to update the control's accelerator table.
	///
	/// \sa OnGetBandInfo, OnInsertBand, OnDeleteBand, Raise_InsertedBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774508.aspx">RB_SETBANDINFO</a>
	LRESULT OnSetBandInfo(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);

	/// \brief <em>\c WM_MDISETMENU handler</em>
	///
	/// Will be called if the MDI frame window's menu (or parts of it) shall be replaced.
	/// We use this handler for the \c ReplaceMDIFrameMenu property.
	///
	/// \sa get_ReplaceMDIFrameMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644915.aspx">WM_MDISETMENU</a>
	LRESULT OnMDIClientMDISetMenu(UINT message, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);

	/// \brief <em>\c WM_ACTIVATE handler</em>
	///
	/// Will be called if the MDI frame window is being activated or deactivated.
	/// We use this handler for the \c ReplaceMDIFrameMenu property.
	///
	/// \sa get_ReplaceMDIFrameMenu,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646274.aspx">WM_ACTIVATE</a>
	LRESULT OnMDIFrameActivate(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);

	/// \brief <em>\c WM_CAPTURECHANGED handler</em>
	///
	/// Will be called if the child window of the rebar band, that has been chosen to replace the MDI frame
	/// window's menu bar is losing the mouse capture.
	/// We use this handler for the \c MDIFrameMenuBand property.
	///
	/// \sa get_MDIFrameMenuBand, OnMenuBandChildNCLButtonDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645605.aspx">WM_CAPTURECHANGED</a>
	LRESULT OnMenuBandChildCaptureChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_LBUTTONUP handler</em>
	///
	/// Will be called if the user releases the left mouse button while the mouse cursor is located over
	/// the window's client area.
	/// We use this handler for the \c MDIFrameMenuBand property.
	///
	/// \sa get_MDIFrameMenuBand, OnMenuBandChildNCLButtonDown, OnMenuBandChildMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645608.aspx">WM_LBUTTONUP</a>
	LRESULT OnMenuBandChildLButtonUp(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_MOUSEMOVE handler</em>
	///
	/// Will be called if the user moves the mouse while the mouse cursor is located over the window's
	/// client area.
	/// We use this handler for the \c MDIFrameMenuBand property.
	///
	/// \sa get_MDIFrameMenuBand, OnMenuBandChildNCLButtonDown, OnMenuBandChildLButtonUp,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645616.aspx">WM_MOUSEMOVE</a>
	LRESULT OnMenuBandChildMouseMove(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NCCALCSIZE handler</em>
	///
	/// Will be called if the size of the non-client area of the child window of the rebar band, that has
	/// been chosen to replace the MDI frame window's menu bar, shall be retrieved.
	/// We use this handler for the \c MDIFrameMenuBand property.
	///
	/// \sa get_MDIFrameMenuBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632634.aspx">WM_NCCALCSIZE</a>
	LRESULT OnMenuBandChildNCCalcSize(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c NM_NCHITTEST handler</em>
	///
	/// Will be called to determine what part of the window corresponds to a particular screen coordinate.
	/// We use this handler for the \c MDIFrameMenuBand property.
	///
	/// \sa get_MDIFrameMenuBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645618.aspx">NM_NCHITTEST</a>
	LRESULT OnMenuBandChildNCHitTest(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_NCLBUTTONDBLCLK handler</em>
	///
	/// Will be called if the user double-clicked into the non-client area of the child window of the rebar
	/// band, that has been chosen to replace the MDI frame window's menu bar, using the left mouse button.
	/// We use this handler for the \c MDIFrameMenuBand property.
	///
	/// \sa get_MDIFrameMenuBand, OnMenuBandChildNCLButtonDown,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645619.aspx">WM_NCLBUTTONDBLCLK</a>
	LRESULT OnMenuBandChildLButtonDblClk(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NCLBUTTONDOWN handler</em>
	///
	/// Will be called if the user presses the left mouse button while the mouse cursor is located over the
	/// non-client area of the child window of the rebar band, that has been chosen to replace the MDI frame
	/// window's menu bar.
	/// We use this handler for the \c MDIFrameMenuBand property.
	///
	/// \sa get_MDIFrameMenuBand, OnMenuBandChildMouseMove,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645620.aspx">WM_NCLBUTTONDOWN</a>
	LRESULT OnMenuBandChildNCLButtonDown(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_NCPAINT handler</em>
	///
	/// Will be called if the non-client area of the child window of the rebar band, that has been chosen
	/// to replace the MDI frame window's menu bar, needs to be drawn.
	/// We use this handler for the \c MDIFrameMenuBand property.
	///
	/// \sa get_MDIFrameMenuBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/dd145212.aspx">WM_NCPAINT</a>
	LRESULT OnMenuBandChildNCPaint(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SIZE handler</em>
	///
	/// Will be called if the child window of the rebar band, that has been chosen to replace the MDI frame
	/// window's menu bar, is being resized.
	/// We use this handler for the \c MDIFrameMenuBand property.
	///
	/// \sa get_MDIFrameMenuBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632646.aspx">WM_SIZE</a>
	LRESULT OnMenuBandChildSize(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);

	/// \brief <em>Message handler handling any messages caught by the message hook</em>
	///
	/// Will be called if a message needs to be processed that has been forwarded to the control by the
	/// message hook.
	/// We use this handler for the \c ReplaceMDIFrameMenu property.
	///
	/// \sa get_ReplaceMDIFrameMenu, GetMsgProc, OnForwardMsg
	LRESULT OnAllHookMessages(UINT message, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Notification handlers
	///
	//@{
	/// \brief <em>\c NM_NCHITTEST handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control has received a
	/// \c WM_NCHITTEST message.
	/// We use this handler to raise the \c NonClientHitTest event.
	///
	/// \sa OnMouseMove, Raise_NonClientHitTest,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645618.aspx">WM_NCHITTEST</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774399.aspx">NM_NCHITTEST</a>
	LRESULT OnNCHitTestNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c NM_RELEASEDCAPTURE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control has released mouse
	/// capture.
	/// We use this handler to raise the \c ReleasedMouseCapture event.
	///
	/// \sa Raise_ReleasedMouseCapture,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774401.aspx">NM_RELEASEDCAPTURE</a>
	LRESULT OnReleasedCaptureNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_AUTOBREAK handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control will break a band onto a
	/// new line.
	/// We use this handler to raise the \c AutoBreakingBand event.
	///
	/// \sa Raise_AutoBreakingBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774403.aspx">RBN_AUTOBREAK</a>
	LRESULT OnAutoBreakNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_AUTOSIZE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control has resized itself
	/// automatically.
	/// We use this handler to raise the \c AutoSized event.
	///
	/// \sa OnLayoutChangedNotification, OnHeightChangeNotification, OnWindowPosChanged, Raise_AutoSized,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774405.aspx">RBN_AUTOSIZE</a>
	LRESULT OnAutoSizeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_BEGINDRAG handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user wants to drag a rebar band.
	/// We use this handler to raise the \c BandBeginDrag and \c BandBeginRDrag events.
	///
	/// \sa OnEndDragNotification, OnLayoutChangedNotification, Raise_BandBeginDrag, Raise_BandBeginRDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774407.aspx">RBN_BEGINDRAG</a>
	LRESULT OnBeginDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_CHEVRONPUSHED handler</em>
	///
	/// Will be called if the control's parent window is notified, that a band's chevron button has been
	/// pushed.
	/// We use this handler to raise the \c ChevronClick event.
	///
	/// \sa ReBarBand::get_UseChevron, Raise_ChevronClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774409.aspx">RBN_CHEVRONPUSHED</a>
	LRESULT OnChevronPushedNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_CHILDSIZE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control is resizing a window
	/// contained in a band.
	/// We use this handler to raise the \c ResizingContainedWindow event.
	///
	/// \sa OnWindowPosChanged, OnLayoutChangedNotification, Raise_ResizingContainedWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774411.aspx">RBN_CHILDSIZE</a>
	LRESULT OnChildSizeNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_DELETINGBAND handler</em>
	///
	/// Will be called if the control's parent window is notified, that a rebar band is being removed.
	/// We use this handler to raise the \c FreeBandData event.
	///
	/// \sa OnDeleteBand, Raise_FreeBandData,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774415.aspx">RBN_DELETINGBAND</a>
	LRESULT OnDeletingBandNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_ENDDRAG handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user has stopped dragging a rebar
	/// band.
	/// We use this handler to raise the \c BandEndDrag event.
	///
	/// \sa OnBeginDragNotification, OnLayoutChangedNotification, Raise_BandEndDrag,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774417.aspx">RBN_ENDDRAG</a>
	LRESULT OnEndDragNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_GETOBJECT handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user is dragging data over a
	/// band.
	/// We use this handler to make \c RBS_REGISTERDROP work together with our implementation of
	/// \c IDropTarget.
	///
	/// \sa put_RegisterForOLEDragDrop,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774419.aspx">RBN_GETOBJECT</a>
	LRESULT OnGetObjectNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_HEIGHTCHANGE handler</em>
	///
	/// Will be called if the control's parent window is notified, that the control's height has changed.
	/// We use this handler to raise the \c HeightChanged event.
	///
	/// \sa OnAutoSizeNotification, OnLayoutChangedNotification, OnWindowPosChanged, Raise_HeightChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774421.aspx">RBN_HEIGHTCHANGE</a>
	LRESULT OnHeightChangeNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_LAYOUTCHANGED handler</em>
	///
	/// Will be called if the control's parent window is notified, that the layout of the bands has been
	/// changed.
	/// We use this handler to raise the \c LayoutChanged event.
	///
	/// \sa OnChildSizeNotification, OnAutoSizeNotification, OnHeightChangeNotification, Raise_LayoutChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774423.aspx">RBN_LAYOUTCHANGED</a>
	LRESULT OnLayoutChangedNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_MINMAX handler</em>
	///
	/// Will be called if the control's parent window is notified, that a band's minimized or maximized state
	/// is about to be toggled.
	/// We use this handler to raise the \c TogglingBand event.
	///
	/// \sa Raise_TogglingBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774425.aspx">RBN_MINMAX</a>
	LRESULT OnMinMaxNotification(int /*controlID*/, LPNMHDR pNotificationDetails, BOOL& /*wasHandled*/);
	/// \brief <em>\c RBN_SPLITTERDRAG handler</em>
	///
	/// Will be called if the control's parent window is notified, that the user is dragging the control's
	/// splitter.
	/// We use this handler to raise the \c DraggingSplitter event.
	///
	/// \sa Raise_DraggingSplitter,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774427.aspx">RBN_SPLITTERDRAG</a>
	LRESULT OnSplitterDragNotification(int /*controlID*/, LPNMHDR /*pNotificationDetails*/, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Command handlers
	///
	//@{
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>\c NM_CUSTOMDRAW handler</em>
	///
	/// Will be called by the \c OnReflectedNotify method if a custom draw notification was received for
	/// the rebar.
	/// We use this handler to raise the \c CustomDraw event.
	///
	/// \sa OnReflectedNotify, Raise_CustomDraw,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb774397.aspx">NM_CUSTOMDRAW (rebar)</a>
	LRESULT OnCustomDrawNotification(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c AutoBreakingBand event</em>
	///
	/// \param[in] pBand The band that is about to be moved.
	/// \param[in,out] pDoAutoBreak If \c VARIANT_FALSE, the caller should keep the band in the current row;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.0 or higher.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_AutoBreakingBand, TBarCtlsLibU::_IReBarEvents::AutoBreakingBand,
	///       Raise_LayoutChanged, ReBarBand::get_NewLine
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_AutoBreakingBand, TBarCtlsLibA::_IReBarEvents::AutoBreakingBand,
	///       Raise_LayoutChanged, ReBarBand::get_NewLine
	/// \endif
	inline HRESULT Raise_AutoBreakingBand(IReBarBand* pBand, VARIANT_BOOL* pDoAutoBreak);
	/// \brief <em>Raises the \c AutoSized event</em>
	///
	/// \param[in] pTargetRectangle A \c RECTANGLE structure specifying the rectangle to which the control
	///            tried to size itself.
	/// \param[in] pActualRectangle A \c RECTANGLE structure specifying the rectangle to which the control
	///            actually sized itself.
	/// \param[in] changedBandHeightOrStyle Indicates if the bands' height or style has changed. If set to
	///            \c VARIANT_TRUE, it has changed; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_AutoSized, TBarCtlsLibU::_IReBarEvents::AutoSized,
	///       TBarCtlsLibU::RECTANGLE, get_AutoUpdateLayout, Raise_HeightChanged, Raise_LayoutChanged,
	///       Raise_ResizedControlWindow
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_AutoSized, TBarCtlsLibA::_IReBarEvents::AutoSized,
	///       TBarCtlsLibU::RECTANGLE, get_AutoUpdateLayout, Raise_HeightChanged, Raise_LayoutChanged,
	///       Raise_ResizedControlWindow
	/// \endif
	inline HRESULT Raise_AutoSized(RECTANGLE* pTargetRectangle, RECTANGLE* pActualRectangle, VARIANT_BOOL changedBandHeightOrStyle);
	/// \brief <em>Raises the \c BandBeginDrag event</em>
	///
	/// \param[in] pBand The band that the user wants to drag.
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
	/// \param[in,out] pCancelDrag If \c VARIANT_TRUE, the caller should abort drag'n'drop; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_BandBeginDrag, TBarCtlsLibU::_IReBarEvents::BandBeginDrag,
	///       Raise_BandBeginRDrag, Raise_BandEndDrag, BeginDragBand, get_AllowBandReordering,
	///       TBarCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_BandBeginDrag, TBarCtlsLibA::_IReBarEvents::BandBeginDrag,
	///       Raise_BandBeginRDrag, Raise_BandEndDrag, BeginDragBand, get_AllowBandReordering,
	///       TBarCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_BandBeginDrag(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails, VARIANT_BOOL* pCancelDrag);
	/// \brief <em>Raises the \c BandBeginRDrag event</em>
	///
	/// \param[in] pBand The band that the user wants to drag.
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
	/// \param[in,out] pCancelDrag If \c VARIANT_TRUE, the caller should abort drag'n'drop; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_BandBeginRDrag, TBarCtlsLibU::_IReBarEvents::BandBeginRDrag,
	///       Raise_BandBeginDrag, Raise_BandEndDrag, BeginDragBand, get_AllowBandReordering,
	///       TBarCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_BandBeginRDrag, TBarCtlsLibA::_IReBarEvents::BandBeginRDrag,
	///       Raise_BandBeginDrag, Raise_BandEndDrag, BeginDragBand, get_AllowBandReordering,
	///       TBarCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_BandBeginRDrag(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails, VARIANT_BOOL* pCancelDrag);
	/// \brief <em>Raises the \c BandEndDrag event</em>
	///
	/// \param[in] pBand The band that the user has dragged.
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
	///   \sa Proxy_IReBarEvents::Fire_BandEndDrag, TBarCtlsLibU::_IReBarEvents::BandEndDrag,
	///       Raise_BandBeginDrag, Raise_BandBeginRDrag, EndDragBand, get_AllowBandReordering,
	///       TBarCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_BandEndDrag, TBarCtlsLibA::_IReBarEvents::BandEndDrag,
	///       Raise_BandBeginDrag, Raise_BandBeginRDrag, EndDragBand, get_AllowBandReordering,
	///       TBarCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_BandEndDrag(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c BandMouseEnter event</em>
	///
	/// \param[in] pBand The band that was entered.
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
	///   \sa Proxy_IReBarEvents::Fire_BandMouseEnter, TBarCtlsLibU::_IReBarEvents::BandMouseEnter,
	///       Raise_BandMouseLeave, Raise_MouseMove, TBarCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_BandMouseEnter, TBarCtlsLibA::_IReBarEvents::BandMouseEnter,
	///       Raise_BandMouseLeave, Raise_MouseMove, TBarCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_BandMouseEnter(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c BandMouseLeave event</em>
	///
	/// \param[in] pBand The band that was left.
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
	///   \sa Proxy_IReBarEvents::Fire_BandMouseLeave, TBarCtlsLibU::_IReBarEvents::BandMouseLeave,
	///       Raise_BandMouseEnter, Raise_MouseMove, TBarCtlsLibU::HitTestConstants
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_BandMouseLeave, TBarCtlsLibA::_IReBarEvents::BandMouseLeave,
	///       Raise_BandMouseEnter, Raise_MouseMove, TBarCtlsLibA::HitTestConstants
	/// \endif
	inline HRESULT Raise_BandMouseLeave(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, UINT hitTestDetails);
	/// \brief <em>Raises the \c BeforeDisplayMDIChildSystemMenu event</em>
	///
	/// \param[in] hActiveMDIChild The handle of the currently active MDI child window of which the
	///            system menu is about to be displayed.
	/// \param[in] hMenu The handle of the system menu that is about to be displayed.
	/// \param[in] x The x-coordinate (in twips) of the menu's proposed position relative to the
	///            control's upper-left corner.
	/// \param[in] y The y-coordinate (in twips) of the menu's proposed position relative to the
	///            control's upper-left corner.
	/// \param[out] pCancelMenu If set to \c VARIANT_TRUE, the menu won't be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_BeforeDisplayMDIChildSystemMenu,
	///       TBarCtlsLibU::_IReBarEvents::BeforeDisplayMDIChildSystemMenu, Raise_CleanupMDIChildSystemMenu,
	///       get_ReplaceMDIFrameMenu
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_BeforeDisplayMDIChildSystemMenu,
	///       TBarCtlsLibA::_IReBarEvents::BeforeDisplayMDIChildSystemMenu, Raise_CleanupMDIChildSystemMenu,
	///       get_ReplaceMDIFrameMenu
	/// \endif
	inline HRESULT Raise_BeforeDisplayMDIChildSystemMenu(LONG hActiveMDIChild, LONG hMenu, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pCancelMenu);
	/// \brief <em>Raises the \c ChevronClick event</em>
	///
	/// \param[in] pBand The band whose chevron button has been clicked.
	/// \param[in] pChevronRectangle The chevron button's bounding rectangle.
	/// \param[in] userData The \c LONG value that has been passed to the \c ClickChevron method. If the
	///            chevron button has been clicked by the user, this parameter will be 0.
	/// \param[out] pDoDefault If set to \c VARIANT_TRUE and the band's child window is a \c ToolBar control,
	///             the rebar control creates and displays a chevron popup automatically.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_ChevronClick, TBarCtlsLibU::_IReBarEvents::ChevronClick,
	///       ReBarBand::get_UseChevron, ReBarBand::ClickChevron
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_ChevronClick, TBarCtlsLibA::_IReBarEvents::ChevronClick,
	///       ReBarBand::get_UseChevron, ReBarBand::ClickChevron
	/// \endif
	inline HRESULT Raise_ChevronClick(IReBarBand* pBand, RECTANGLE* pChevronRectangle, LONG userData, VARIANT_BOOL* pDoDefault);
	/// \brief <em>Raises the \c CleanupMDIChildSystemMenu event</em>
	///
	/// \param[in] hActiveMDIChild The handle of the currently active MDI child window of which the
	///            system menu is about to be displayed.
	/// \param[in] hMenu The handle of the system menu that is about to be displayed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_CleanupMDIChildSystemMenu,
	///       TBarCtlsLibU::_IReBarEvents::CleanupMDIChildSystemMenu, Raise_BeforeDisplayMDIChildSystemMenu,
	///       get_ReplaceMDIFrameMenu
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_CleanupMDIChildSystemMenu,
	///       TBarCtlsLibA::_IReBarEvents::CleanupMDIChildSystemMenu, Raise_BeforeDisplayMDIChildSystemMenu,
	///       get_ReplaceMDIFrameMenu
	/// \endif
	inline HRESULT Raise_CleanupMDIChildSystemMenu(LONG hActiveMDIChild, LONG hMenu);
	/// \brief <em>Raises the \c Click event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_Click, TBarCtlsLibU::_IReBarEvents::Click, Raise_DblClick,
	///       Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_Click, TBarCtlsLibA::_IReBarEvents::Click, Raise_DblClick,
	///       Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_Click(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c ContextMenu event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_ContextMenu, TBarCtlsLibU::_IReBarEvents::ContextMenu, Raise_RClick
	///      
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_ContextMenu, TBarCtlsLibA::_IReBarEvents::ContextMenu, Raise_RClick
	/// \endif
	inline HRESULT Raise_ContextMenu(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c CustomDraw event</em>
	///
	/// \param[in] pBand The band that the notification refers to. May be \c NULL.
	/// \param[in] drawStage The stage of custom drawing this event is raised for. Most of the values
	///            defined by the \c CustomDrawStageConstants enumeration are valid.
	/// \param[in] bandState The band's current state. For current versions of Windows this seems to be
	///            always 0.
	/// \param[in] hDC The handle of the device context in which all drawing shall take place.
	/// \param[in] pDrawingRectangle A \c RECTANGLE structure specifying the bounding rectangle of the
	///            area that needs to be drawn.
	/// \param[in,out] pFurtherProcessing A value controlling further drawing. Most of the values defined
	///                by the \c CustomDrawReturnValuesConstants enumeration are valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_CustomDraw, TBarCtlsLibU::_IReBarEvents::CustomDraw,
	///       TBarCtlsLibU::RECTANGLE, TBarCtlsLibU::CustomDrawStageConstants,
	///       TBarCtlsLibU::CustomDrawItemStateConstants, TBarCtlsLibU::CustomDrawReturnValuesConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb774397.aspx">NM_CUSTOMDRAW (rebar)</a>
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_CustomDraw, TBarCtlsLibA::_IReBarEvents::CustomDraw,
	///       TBarCtlsLibA::RECTANGLE, TBarCtlsLibA::CustomDrawStageConstants,
	///       TBarCtlsLibA::CustomDrawItemStateConstants, TBarCtlsLibA::CustomDrawReturnValuesConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb774397.aspx">NM_CUSTOMDRAW (rebar)</a>
	/// \endif
	inline HRESULT Raise_CustomDraw(IReBarBand* pBand, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants bandState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing);
	/// \brief <em>Raises the \c DblClick event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_DblClick, TBarCtlsLibU::_IReBarEvents::DblClick, Raise_Click,
	///       Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_DblClick, TBarCtlsLibA::_IReBarEvents::DblClick, Raise_Click,
	///       Raise_MDblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_DblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_DestroyedControlWindow,
	///       TBarCtlsLibU::_IReBarEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_DestroyedControlWindow,
	///       TBarCtlsLibA::_IReBarEvents::DestroyedControlWindow, Raise_RecreatedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_DestroyedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c DraggingSplitter event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_DraggingSplitter, TBarCtlsLibU::_IReBarEvents::DraggingSplitter,
	///       get_DisplaySplitter
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_DraggingSplitter, TBarCtlsLibA::_IReBarEvents::DraggingSplitter,
	///       get_DisplaySplitter
	/// \endif
	inline HRESULT Raise_DraggingSplitter(void);
	/// \brief <em>Raises the \c FreeBandData event</em>
	///
	/// \param[in] pBand The band whose associated data shall be freed. If \c NULL, all bands' associated data
	///            shall be removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_FreeBandData, TBarCtlsLibU::_IReBarEvents::FreeBandData,
	///       Raise_RemovingBand, Raise_RemovedBand, ReBarBand::put_BandData
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_FreeBandData, TBarCtlsLibA::_IReBarEvents::FreeBandData,
	///       Raise_RemovingBand, Raise_RemovedBand, ReBarBand::put_BandData
	/// \endif
	inline HRESULT Raise_FreeBandData(IReBarBand* pBand);
	/// \brief <em>Raises the \c HeightChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_HeightChanged, TBarCtlsLibU::_IReBarEvents::HeightChanged,
	///       Raise_AutoSized, Raise_LayoutChanged, Raise_ResizedControlWindow
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_HeightChanged, TBarCtlsLibA::_IReBarEvents::HeightChanged,
	///       Raise_AutoSized, Raise_LayoutChanged, Raise_ResizedControlWindow
	/// \endif
	inline HRESULT Raise_HeightChanged(void);
	/// \brief <em>Raises the \c InsertedBand event</em>
	///
	/// \param[in] pBand The inserted band.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_InsertedBand, TBarCtlsLibU::_IReBarEvents::InsertedBand,
	///       Raise_InsertingBand, Raise_RemovedBand
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_InsertedBand, TBarCtlsLibA::_IReBarEvents::InsertedBand,
	///       Raise_InsertingBand, Raise_RemovedBand
	/// \endif
	inline HRESULT Raise_InsertedBand(IReBarBand* pBand);
	/// \brief <em>Raises the \c InsertingBand event</em>
	///
	/// \param[in] pBand The band being inserted.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort insertion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_InsertingBand, TBarCtlsLibU::_IReBarEvents::InsertingBand,
	///       Raise_InsertedBand, Raise_RemovingBand
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_InsertingBand, TBarCtlsLibA::_IReBarEvents::InsertingBand,
	///       Raise_InsertedBand, Raise_RemovingBand
	/// \endif
	inline HRESULT Raise_InsertingBand(IVirtualReBarBand* pBand, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c LayoutChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_LayoutChanged, TBarCtlsLibU::_IReBarEvents::LayoutChanged,
	///       Raise_ResizingContainedWindow, Raise_AutoSized, Raise_HeightChanged, Raise_ResizedControlWindow
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_LayoutChanged, TBarCtlsLibA::_IReBarEvents::LayoutChanged,
	///       Raise_ResizingContainedWindow, Raise_AutoSized, Raise_HeightChanged, Raise_ResizedControlWindow
	/// \endif
	inline HRESULT Raise_LayoutChanged(void);
	/// \brief <em>Raises the \c MClick event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_MClick, TBarCtlsLibU::_IReBarEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_MClick, TBarCtlsLibA::_IReBarEvents::MClick, Raise_MDblClick,
	///       Raise_Click, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MDblClick event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_MDblClick, TBarCtlsLibU::_IReBarEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_MDblClick, TBarCtlsLibA::_IReBarEvents::MDblClick, Raise_MClick,
	///       Raise_DblClick, Raise_RDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_MDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseDown event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_MouseDown, TBarCtlsLibU::_IReBarEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_MouseDown, TBarCtlsLibA::_IReBarEvents::MouseDown, Raise_MouseUp,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseDown(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseEnter event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_MouseEnter, TBarCtlsLibU::_IReBarEvents::MouseEnter, Raise_MouseLeave,
	///       Raise_BandMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_MouseEnter, TBarCtlsLibA::_IReBarEvents::MouseEnter, Raise_MouseLeave,
	///       Raise_BandMouseEnter, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseEnter(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseHover event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_MouseHover, TBarCtlsLibU::_IReBarEvents::MouseHover, Raise_MouseEnter,
	///       Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_MouseHover, TBarCtlsLibA::_IReBarEvents::MouseHover, Raise_MouseEnter,
	///       Raise_MouseLeave, Raise_MouseMove, put_HoverTime
	/// \endif
	inline HRESULT Raise_MouseHover(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseLeave event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_MouseLeave, TBarCtlsLibU::_IReBarEvents::MouseLeave, Raise_MouseEnter,
	///       Raise_BandMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_MouseLeave, TBarCtlsLibA::_IReBarEvents::MouseLeave, Raise_MouseEnter,
	///       Raise_BandMouseLeave, Raise_MouseHover, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_MouseLeave(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseMove event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_MouseMove, TBarCtlsLibU::_IReBarEvents::MouseMove, Raise_MouseEnter,
	///       Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_MouseMove, TBarCtlsLibA::_IReBarEvents::MouseMove, Raise_MouseEnter,
	///       Raise_MouseLeave, Raise_MouseDown, Raise_MouseUp
	/// \endif
	inline HRESULT Raise_MouseMove(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c MouseUp event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_MouseUp, TBarCtlsLibU::_IReBarEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_MouseUp, TBarCtlsLibA::_IReBarEvents::MouseUp, Raise_MouseDown,
	///       Raise_Click, Raise_MClick, Raise_RClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_MouseUp(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c NonClientHitTest event</em>
	///
	/// \param[in] pBand The band that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	/// \param[out] pReturnValue The value to return in response of the \c WM_NCHITTEST message. Any of
	///             the \c HT* values specified on
	///             <a href="https://msdn.microsoft.com/en-us/library/ms645618.aspx">MSDN</a> is valid.\n
	///             If set to 0, the control should process this message itself.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_NonClientHitTest, TBarCtlsLibU::_IReBarEvents::NonClientHitTest,
	///       Raise_MouseMove
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_NonClientHitTest, TBarCtlsLibA::_IReBarEvents::NonClientHitTest,
	///       Raise_MouseMove
	/// \endif
	inline HRESULT Raise_NonClientHitTest(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, LONG* pReturnValue);
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
	/// \param[out] pCallDropTargetHelper If set to \c TRUE, the caller should call the appropriate
	///             \c IDropTargetHelper method; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_OLEDragDrop, TBarCtlsLibU::_IReBarEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_OLEDragDrop, TBarCtlsLibA::_IReBarEvents::OLEDragDrop,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_MouseUp,
	///       get_RegisterForOLEDragDrop, FinishOLEDragDrop,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragDrop(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition, BOOL* /*pCallDropTargetHelper*/);
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
	/// \param[out] pCallDropTargetHelper If set to \c TRUE, the caller should call the appropriate
	///             \c IDropTargetHelper method; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_OLEDragEnter, TBarCtlsLibU::_IReBarEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_OLEDragEnter, TBarCtlsLibA::_IReBarEvents::OLEDragEnter,
	///       Raise_OLEDragMouseMove, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseEnter,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragEnter(IDataObject* pData, DWORD* pEffect, DWORD keyState, POINTL mousePosition, BOOL* /*pCallDropTargetHelper*/);
	/// \brief <em>Raises the \c OLEDragLeave event</em>
	///
	/// \param[out] pCallDropTargetHelper If set to \c TRUE, the caller should call the appropriate
	///             \c IDropTargetHelper method; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_OLEDragLeave, TBarCtlsLibU::_IReBarEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_OLEDragLeave, TBarCtlsLibA::_IReBarEvents::OLEDragLeave,
	///       Raise_OLEDragEnter, Raise_OLEDragMouseMove, Raise_OLEDragDrop, Raise_MouseLeave
	/// \endif
	inline HRESULT Raise_OLEDragLeave(BOOL* /*pCallDropTargetHelper*/);
	/// \brief <em>Raises the \c OLEDragMouseMove event</em>
	///
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the \c DROPEFFECT
	///                enumeration) supported by the drag source. On return, the drop effect that the
	///                client wants to be used on drop.
	/// \param[in] keyState The pressed modifier keys (Shift, Ctrl, Alt) and mouse buttons. Any
	///            combination of the pressed button's and key's \c MK_* flags is valid.
	/// \param[in] mousePosition The mouse cursor's position (in pixels) relative to the screen's
	///            upper-left corner.
	/// \param[out] pCallDropTargetHelper If set to \c TRUE, the caller should call the appropriate
	///             \c IDropTargetHelper method; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_OLEDragMouseMove, TBarCtlsLibU::_IReBarEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_OLEDragMouseMove, TBarCtlsLibA::_IReBarEvents::OLEDragMouseMove,
	///       Raise_OLEDragEnter, Raise_OLEDragLeave, Raise_OLEDragDrop, Raise_MouseMove,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms693457.aspx">DROPEFFECT</a>
	/// \endif
	inline HRESULT Raise_OLEDragMouseMove(DWORD* pEffect, DWORD keyState, POINTL mousePosition, BOOL* /*pCallDropTargetHelper*/);
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
	/// \remarks This event may be disabled.\n
	///          This event won't be fired if the \c ReplaceMDIFrameMenu property is set to
	///          \c rmfmDontReplace.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_RawMenuMessage, TBarCtlsLibU::_IReBarEvents::RawMenuMessage,
	///       get_ReplaceMDIFrameMenu, Raise_SelectedMenuItem, Raise_BeforeDisplayMDIChildSystemMenu,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775923.aspx">WM_DRAWITEM</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646347.aspx">WM_INITMENUPOPUP</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/bb775925.aspx">WM_MEASUREITEM</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646349.aspx">WM_MENUCHAR</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms646352.aspx">WM_MENUSELECT</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647612.aspx">WM_NEXTMENU</a>
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_RawMenuMessage, TBarCtlsLibA::_IReBarEvents::RawMenuMessage,
	///       get_ReplaceMDIFrameMenu, Raise_SelectedMenuItem, Raise_BeforeDisplayMDIChildSystemMenu,
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
	///   \sa Proxy_IReBarEvents::Fire_RClick, TBarCtlsLibU::_IReBarEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_RClick, TBarCtlsLibA::_IReBarEvents::RClick, Raise_ContextMenu,
	///       Raise_RDblClick, Raise_Click, Raise_MClick, Raise_XClick
	/// \endif
	inline HRESULT Raise_RClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RDblClick event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_RDblClick, TBarCtlsLibU::_IReBarEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_RDblClick, TBarCtlsLibA::_IReBarEvents::RDblClick, Raise_RClick,
	///       Raise_DblClick, Raise_MDblClick, Raise_XDblClick
	/// \endif
	inline HRESULT Raise_RDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_RecreatedControlWindow,
	///       TBarCtlsLibU::_IReBarEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_RecreatedControlWindow,
	///       TBarCtlsLibA::_IReBarEvents::RecreatedControlWindow, Raise_DestroyedControlWindow,
	///       get_hWnd
	/// \endif
	inline HRESULT Raise_RecreatedControlWindow(LONG hWnd);
	/// \brief <em>Raises the \c ReleasedMouseCapture event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_ReleasedMouseCapture,
	///       TBarCtlsLibU::_IReBarEvents::ReleasedMouseCapture, Raise_MouseMove
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_ReleasedMouseCapture,
	///       TBarCtlsLibA::_IReBarEvents::ReleasedMouseCapture, Raise_MouseMove
	/// \endif
	inline HRESULT Raise_ReleasedMouseCapture(void);
	/// \brief <em>Raises the \c RemovedBand event</em>
	///
	/// \param[in] pBand The removed band. If \c NULL, all bands were removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_RemovedBand, TBarCtlsLibU::_IReBarEvents::RemovedBand,
	///       Raise_RemovingBand, Raise_InsertedBand
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_RemovedBand, TBarCtlsLibA::_IReBarEvents::RemovedBand,
	///       Raise_RemovingBand, Raise_InsertedBand
	/// \endif
	inline HRESULT Raise_RemovedBand(IVirtualReBarBand* pBand);
	/// \brief <em>Raises the \c RemovingBand event</em>
	///
	/// \param[in] pBand The band being removed. If \c NULL, all bands are removed.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort deletion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This event may be disabled.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_RemovingBand, TBarCtlsLibU::_IReBarEvents::RemovingBand,
	///       Raise_RemovedBand, Raise_InsertingBand
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_RemovingBand, TBarCtlsLibA::_IReBarEvents::RemovingBand,
	///       Raise_RemovedBand, Raise_InsertingBand
	/// \endif
	inline HRESULT Raise_RemovingBand(IReBarBand* pBand, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_ResizedControlWindow,
	///       TBarCtlsLibU::_IReBarEvents::ResizedControlWindow, Raise_HeightChanged, Raise_LayoutChanged,
	///       Raise_ResizingContainedWindow
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_ResizedControlWindow,
	///       TBarCtlsLibA::_IReBarEvents::ResizedControlWindow, Raise_HeightChanged, Raise_LayoutChanged,
	///       Raise_ResizingContainedWindow
	/// \endif
	inline HRESULT Raise_ResizedControlWindow(void);
	/// \brief <em>Raises the \c ResizingContainedWindow event</em>
	///
	/// \param[in] pBand The band whose contained window is about to be resized.
	/// \param[in] pBandClientRectangle A \c RECTANGLE structure specifying the rectangle surrounding the
	///            area that the contained window may occupy.
	/// \param[in,out] pContainedWindowRectangle A \c RECTANGLE structure specifying the rectangle to which
	///                the contained window is being resized. This rectangle may be changed by the client.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_ResizingContainedWindow,
	///       TBarCtlsLibU::_IReBarEvents::ResizingContainedWindow, Raise_ResizedControlWindow,
	///       Raise_LayoutChanged, ReBarBand::get_hContainedWindow
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_ResizingContainedWindow,
	///       TBarCtlsLibA::_IReBarEvents::ResizingContainedWindow, Raise_ResizedControlWindow,
	///       Raise_LayoutChanged, ReBarBand::get_hContainedWindow
	/// \endif
	inline HRESULT Raise_ResizingContainedWindow(IReBarBand* pBand, RECTANGLE* pBandClientRectangle, RECTANGLE* pContainedWindowRectangle);
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
	///   \sa Proxy_IReBarEvents::Fire_SelectedMenuItem, TBarCtlsLibU::_IReBarEvents::SelectedMenuItem,
	///       Raise_RawMenuMessage, Raise_BeforeDisplayMDIChildSystemMenu, get_ReplaceMDIFrameMenu,
	///       TBarCtlsLibU::MenuItemStateConstants
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_SelectedMenuItem, TBarCtlsLibA::_IReBarEvents::SelectedMenuItem,
	///       Raise_RawMenuMessage, Raise_BeforeDisplayMDIChildSystemMenu, get_ReplaceMDIFrameMenu,
	///       TBarCtlsLibA::MenuItemStateConstants
	/// \endif
	inline HRESULT Raise_SelectedMenuItem(LONG commandIDOrSubMenuIndex, MenuItemStateConstants menuItemState, LONG hMenu);
	/// \brief <em>Raises the \c TogglingBand event</em>
	///
	/// \param[in] pBand The band whose minimized or maximized state is about to be toggled.
	/// \param[in,out] pCancel If \c VARIANT_TRUE, the caller should abort toggling, otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IReBarEvents::Fire_TogglingBand, TBarCtlsLibU::_IReBarEvents::TogglingBand,
	///       get_ToggleOnDoubleClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_TogglingBand, TBarCtlsLibA::_IReBarEvents::TogglingBand,
	///       get_ToggleOnDoubleClick
	/// \endif
	inline HRESULT Raise_TogglingBand(IReBarBand* pBand, VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c XClick event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_XClick, TBarCtlsLibU::_IReBarEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_XClick, TBarCtlsLibA::_IReBarEvents::XClick,
	///       Raise_XDblClick, Raise_Click, Raise_MClick, Raise_RClick
	/// \endif
	inline HRESULT Raise_XClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	/// \brief <em>Raises the \c XDblClick event</em>
	///
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
	///   \sa Proxy_IReBarEvents::Fire_XDblClick, TBarCtlsLibU::_IReBarEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick
	/// \else
	///   \sa Proxy_IReBarEvents::Fire_XDblClick, TBarCtlsLibA::_IReBarEvents::XDblClick,
	///       Raise_XClick, Raise_DblClick, Raise_MDblClick, Raise_RDblClick
	/// \endif
	inline HRESULT Raise_XDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

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

	/// \brief <em>Applies the font to ourselves</em>
	///
	/// This method sets our font to the font specified by the \c Font property.
	///
	/// \sa get_Font
	void ApplyFont(void);

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
	/// \sa OnMnemonic, Properties::hAcceleratorTable,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693730.aspx">IOleControl::GetControlInfo</a>
	virtual HRESULT STDMETHODCALLTYPE GetControlInfo(LPCONTROLINFO pControlInfo);
	/// \brief <em>Implementation of \c IOleControl::OnMnemonic</em>
	///
	/// Will be called if the user has pressed a keystroke that matches one of the entries in the control's
	/// accelerator table.
	///
	/// \param[in] pMessage The \c MSG structure describing the keystroke.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetControlInfo, Properties::hAcceleratorTable,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms680699.aspx">IOleControl::OnMnemonic</a>
	virtual HRESULT STDMETHODCALLTYPE OnMnemonic(LPMSG pMessage);
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
	/// Sends \c WM_* and \c RB_* messages to the control window to make it match the current property
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
	/// \brief <em>Translates a band's unique ID to its index</em>
	///
	/// \param[in] ID The unique ID of the band whose index is requested.
	///
	/// \return The requested band's index if successful; -1 otherwise.
	///
	/// \sa BandIndexToID, ReBarBands::get_Item, ReBarBands::Remove
	int IDToBandIndex(LONG ID);
	/// \brief <em>Translates a band's index to its unique ID</em>
	///
	/// \param[in] bandIndex The index of the band whose unique ID is requested.
	///
	/// \return The requested band's unique ID if successful; -1 otherwise.
	///
	/// \sa IDToBandIndex, ReBarBand::get_ID
	LONG BandIndexToID(int bandIndex);
	/// \brief <em>Retrieves the rebar band that contains the specified window</em>
	///
	/// \param[in] hWndChild The handle of the child window to search.
	///
	/// \return The zero-based index of the band that contains the specified window; -1 otherwise.
	///
	/// \sa get_MDIFrameMenuBand
	int FindBandByChildWindow(__in HWND hWndChild);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name MDI frame window menu replacement
	///
	//@{
	/// \brief <em>Message handling method for the MDI client window</em>
	///
	/// This method is the callback method used when subclassing the MDI client window.
	///
	/// \param[in] hWnd The handle of the window that the message has been sent to. This always should be the
	///            MDI client window.
	/// \param[in] message The message to handle.
	/// \param[in] wParam The message's \c wParam parameter.
	/// \param[in] lParam The message's \c lParam parameter.
	/// \param[in] idSubclass A unique ID identifying the instance of the rebar control that this message
	///            is meant for. This ID must have been passed to \c SetWindowSubclass when subclassing the
	///            window. We pass the instance pointer (\c this).
	/// \param[in] refData An application defined data that can be passed to \c SetWindowSubclass and will be
	///            passed back to the message handling method. We do not use this feature.
	///
	/// \return A message-dependent return value.
	///
	/// \sa SendConfigurationMessages, Raise_DestroyedControlWindow, mdiClient,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762102.aspx">SetWindowSubclass</a>
	static LRESULT CALLBACK MDIClientSubclass(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/);
	/// \brief <em>Message handling method for the MDI frame window</em>
	///
	/// This method is the callback method used when subclassing the MDI frame window.
	///
	/// \param[in] hWnd The handle of the window that the message has been sent to. This always should be the
	///            MDI frame window.
	/// \param[in] message The message to handle.
	/// \param[in] wParam The message's \c wParam parameter.
	/// \param[in] lParam The message's \c lParam parameter.
	/// \param[in] idSubclass A unique ID identifying the instance of the rebar control that this message
	///            is meant for. This ID must have been passed to \c SetWindowSubclass when subclassing the
	///            window. We pass the instance pointer (\c this).
	/// \param[in] refData An application defined data that can be passed to \c SetWindowSubclass and will be
	///            passed back to the message handling method. We do not use this feature.
	///
	/// \return A message-dependent return value.
	///
	/// \sa SendConfigurationMessages, Raise_DestroyedControlWindow, mdiFrame,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762102.aspx">SetWindowSubclass</a>
	static LRESULT CALLBACK MDIFrameSubclass(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/);
	/// \brief <em>Message handling method for the window contained by the rebar band that replaces the MDI frame window's menu bar</em>
	///
	/// This method is the callback method used when subclassing the menu band's child window.
	///
	/// \param[in] hWnd The handle of the window that the message has been sent to. This always should be the
	///            menu band's child window.
	/// \param[in] message The message to handle.
	/// \param[in] wParam The message's \c wParam parameter.
	/// \param[in] lParam The message's \c lParam parameter.
	/// \param[in] idSubclass A unique ID identifying the instance of the rebar control that this message
	///            is meant for. This ID must have been passed to \c SetWindowSubclass when subclassing the
	///            window. We pass the instance pointer (\c this).
	/// \param[in] refData An application defined data that can be passed to \c SetWindowSubclass and will be
	///            passed back to the message handling method. We do not use this feature.
	///
	/// \return A message-dependent return value.
	///
	/// \sa SendConfigurationMessages, Raise_DestroyedControlWindow, putref_MDIFrameMenuBand, OnInsertBand,
	///     OnSetBandInfo, menuBandWindow,
	///     <a href="https://msdn.microsoft.com/en-us/library/bb762102.aspx">SetWindowSubclass</a>
	static LRESULT CALLBACK MenuBandWindowSubclass(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR idSubclass, DWORD_PTR /*refData*/);
	/// \brief <em>Calculates the size of the MDI child window icon and the MDI child window system command buttons based on system settings</em>
	///
	/// \sa AdjustMDINonClientButtonSize
	void UpdateMDINonClientAreaSizes(void);
	/// \brief <em>Adjusts the height of the MDI child window system command buttons</em>
	///
	/// \param[in] height The buttons' maximum height in pixels.
	///
	/// \sa UpdateMDINonClientAreaSizes
	void AdjustMDINonClientButtonSize(int height);
	/// \brief <em>Calculates the rectangle into which the active MDI child window's icon shall be drawn</em>
	///
	/// \param[in] availableWidth The maximum width (in pixels) of the rectangle.
	/// \param[in] availableHeight The maximum height (in pixels) of the rectangle.
	/// \param[out] pIconRectangle Receives the icon's bounding rectangle.
	/// \param[in] invertX If \c TRUE, the icon is right-aligned instead of left-aligned. In case of
	///            right-to-left window layout, this parameter should be set to \c TRUE.
	///
	/// \sa CalcMDIChildButtonRectangles, OnMenuBandChildNCPaint
	void CalcMDIChildIconRectangle(int availableWidth, int availableHeight, __out LPRECT pIconRectangle, BOOL invertX = FALSE) const;
	/// \brief <em>Calculates the rectangles into which the active MDI child window's system command buttons shall be drawn</em>
	///
	/// \param[in] availableWidth The maximum width (in pixels) of the rectangle.
	/// \param[in] availableHeight The maximum height (in pixels) of the rectangle.
	/// \param[out] buttonRectangles Receives the system command buttons' bounding rectangles.
	/// \param[in] invertX If \c TRUE, the buttons are right-aligned instead of left-aligned. In case of
	///            right-to-left window layout, this parameter should be set to \c TRUE.
	///
	/// \sa DrawMDIChildButton, CalcMDIChildIconRectangle, OnMenuBandChildNCPaint
	void CalcMDIChildButtonRectangles(int availableWidth, int availableHeight, RECT buttonRectangles[3], BOOL invertX = FALSE) const;
	/// \brief <em>Draws the specified MDI system command button into the specified rectangle into the specified device context</em>
	///
	/// \param[in] memoryDC The device context into which to draw.
	/// \param[in] buttonRectangles The bounding rectangles of the MDI child window's system command buttons.
	/// \param[in] button The zero-based index of the button to draw:
	///            - 0: close button
	///            - 1: restore button
	///            - 2: minimize button
	///
	/// \sa CalcMDIChildButtonRectangles, OnMenuBandChildNCPaint
	void DrawMDIChildButton(CMemoryDC& memoryDC, RECT buttonRectangles[3], int button);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Hit-tests the specified point</em>
	///
	/// Retrieves the band that lies below the point ('x'; 'y').
	///
	/// \param[in] x The x-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the point to check. It must be relative to the control's
	///            upper-left corner.
	/// \param[out] pFlags A bit field of RBHT_* flags, that holds further details about the control's
	///             part below the specified point.
	// \param[in] ignoreBoundingBoxDefinition If \c TRUE the setting of the \c BandBoundingBoxDefinition
	//            property is ignored; otherwise the returned band index is set to -1, if the
	//            \c BandBoundingBoxDefinition property's setting and the \c pFlags parameter don't
	//            match.
	///
	/// \return The "hit" band's zero-based index.
	///
	/// \sa HitTest
	// \sa HitTest, get_BandBoundingBoxDefinition
	int HitTest(LONG x, LONG y, UINT* pFlags/*, BOOL ignoreBoundingBoxDefinition = FALSE*/);
	/// \brief <em>Retrieves whether we're in design mode or in user mode</em>
	///
	/// \return \c TRUE if the control is in design mode (i. e. is placed on a window which is edited
	///         by a form editor); \c FALSE if the control is in user mode (i. e. is placed on a window
	///         getting used by an end-user).
	BOOL IsInDesignMode(void);
	/// \brief <em>Recreates the control's accelerator table</em>
	///
	/// \sa Properties::hAcceleratorTable, OnInsertBand, OnDeleteBand, OnSetBandInfo, GetControlInfo,
	///     OnMnemonic
	void RebuildAcceleratorTable(void);


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
			CComObject< PropertyNotifySinkImpl<ReBar> >* pPropertyNotifySink;

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
			HRESULT InitializePropertyWatcher(ReBar* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ReBar> >::CreateInstance(&pPropertyNotifySink);
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
			CComObject< PropertyNotifySinkImpl<ReBar> >* pPropertyNotifySink;

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
			HRESULT InitializePropertyWatcher(ReBar* pObjectToNotify, DISPID propertyToWatch)
			{
				CComObject< PropertyNotifySinkImpl<ReBar> >::CreateInstance(&pPropertyNotifySink);
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

		/// \brief <em>Holds the \c AllowBandReordering property's setting</em>
		///
		/// \sa get_AllowBandReordering, put_AllowBandReordering
		UINT allowBandReordering : 1;
		/// \brief <em>Holds the \c Appearance property's setting</em>
		///
		/// \sa get_Appearance, put_Appearance
		AppearanceConstants appearance;
		/// \brief <em>Holds the \c AutoUpdateLayout property's setting</em>
		///
		/// \sa get_AutoUpdateLayout, put_AutoUpdateLayout
		UINT autoUpdateLayout : 1;
		/// \brief <em>Holds the \c BackColor property's setting</em>
		///
		/// \sa get_BackColor, put_BackColor
		OLE_COLOR backColor;
		/// \brief <em>Holds the \c BorderStyle property's setting</em>
		///
		/// \sa get_BorderStyle, put_BorderStyle
		BorderStyleConstants borderStyle;
		/// \brief <em>Holds the \c DisabledEvents property's setting</em>
		///
		/// \sa get_DisabledEvents, put_DisabledEvents
		DisabledEventsConstants disabledEvents;
		/// \brief <em>Holds the \c DisplayBandSeparators property's setting</em>
		///
		/// \sa get_DisplayBandSeparators, put_DisplayBandSeparators
		UINT displayBandSeparators : 1;
		/// \brief <em>Holds the \c DisplaySplitter property's setting</em>
		///
		/// \sa get_DisplaySplitter, put_DisplaySplitter
		UINT displaySplitter : 1;
		/// \brief <em>Holds the \c DontRedraw property's setting</em>
		///
		/// \sa get_DontRedraw, put_DontRedraw
		UINT dontRedraw : 1;
		/// \brief <em>Holds the \c Enabled property's setting</em>
		///
		/// \sa get_Enabled, put_Enabled
		UINT enabled : 1;
		/// \brief <em>Holds the \c FixedBandHeight property's setting</em>
		///
		/// \sa get_FixedBandHeight, put_FixedBandHeight
		UINT fixedBandHeight : 1;
		/// \brief <em>Holds the \c Font property's settings</em>
		///
		/// \sa get_Font, put_Font, putref_Font
		FontProperty font;
		/// \brief <em>Holds the \c ForeColor property's setting</em>
		///
		/// \sa get_ForeColor, put_ForeColor
		OLE_COLOR foreColor;
		/// \brief <em>Holds the control's accelerator table</em>
		///
		/// \sa RebuildAcceleratorTable, GetControlInfo, OnMnemonic
		HACCEL hAcceleratorTable;
		/// \brief <em>Holds the \c HighlightColor property's setting</em>
		///
		/// \sa get_HighlightColor, put_HighlightColor
		OLE_COLOR highlightColor;
		/// \brief <em>Holds the \c HoverTime property's setting</em>
		///
		/// \sa get_HoverTime, put_HoverTime
		long hoverTime;
		/// \brief <em>Holds the \c MouseIcon property's settings</em>
		///
		/// \sa get_MouseIcon, put_MouseIcon, putref_MouseIcon
		PictureProperty mouseIcon;
		/// \brief <em>Holds the \c MousePointer property's setting</em>
		///
		/// \sa get_MousePointer, put_MousePointer
		MousePointerConstants mousePointer;
		/// \brief <em>Holds the \c Orientation property's setting</em>
		///
		/// \sa get_Orientation, put_Orientation
		OrientationConstants orientation;
		/// \brief <em>Holds the \c RegisterForOLEDragDrop property's setting</em>
		///
		/// \sa get_RegisterForOLEDragDrop, put_RegisterForOLEDragDrop
		RegisterForOLEDragDropConstants registerForOLEDragDrop;
		/// \brief <em>Holds the \c ReplaceMDIFrameMenu property's setting</em>
		///
		/// \sa get_ReplaceMDIFrameMenu, put_ReplaceMDIFrameMenu
		ReplaceMDIFrameMenuConstants replaceMDIFrameMenu;
		/// \brief <em>Holds the \c RightToLeft property's setting</em>
		///
		/// \sa get_RightToLeft, put_RightToLeft
		RightToLeftConstants rightToLeft;
		/// \brief <em>Holds the \c ShadowColor property's setting</em>
		///
		/// \sa get_ShadowColor, put_ShadowColor
		OLE_COLOR shadowColor;
		/// \brief <em>Holds the \c SupportOLEDragImages property's setting</em>
		///
		/// \sa get_SupportOLEDragImages, put_SupportOLEDragImages
		UINT supportOLEDragImages : 1;
		/// \brief <em>Holds the \c ToggleOnDoubleClick property's setting</em>
		///
		/// \sa get_ToggleOnDoubleClick, put_ToggleOnDoubleClick
		UINT toggleOnDoubleClick : 1;
		/// \brief <em>Holds the \c UseSystemFont property's setting</em>
		///
		/// \sa get_UseSystemFont, put_UseSystemFont
		UINT useSystemFont : 1;
		/// \brief <em>Holds the \c VerticalSizingGripsOnVerticalOrientation property's setting</em>
		///
		/// \sa get_VerticalSizingGripsOnVerticalOrientation, put_VerticalSizingGripsOnVerticalOrientation
		UINT verticalSizingGripsOnVerticalOrientation : 1;

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
			if(hAcceleratorTable) {
				DestroyAcceleratorTable(hAcceleratorTable);
				hAcceleratorTable = NULL;
			}
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			allowBandReordering = TRUE;
			appearance = a2D;
			autoUpdateLayout = TRUE;
			backColor = static_cast<OLE_COLOR>(-1);
			borderStyle = bsNone;
			disabledEvents = static_cast<DisabledEventsConstants>(deMouseEvents | deClickEvents | deKeyboardEvents | deBandInsertionEvents | deBandDeletionEvents | deFreeBandData | deCustomDraw | deRawMenuMessage);
			displayBandSeparators = TRUE;
			displaySplitter = FALSE;
			dontRedraw = FALSE;
			enabled = TRUE;
			fixedBandHeight = FALSE;
			foreColor = static_cast<OLE_COLOR>(-1);
			hAcceleratorTable = NULL;
			highlightColor = static_cast<OLE_COLOR>(-1);
			hoverTime = -1;
			mousePointer = mpDefault;
			orientation = oHorizontal;
			registerForOLEDragDrop = rfoddNoDragDrop;
			replaceMDIFrameMenu = rmfmDontReplace;
			rightToLeft = static_cast<RightToLeftConstants>(0);
			shadowColor = static_cast<OLE_COLOR>(-1);
			supportOLEDragImages = TRUE;
			toggleOnDoubleClick = TRUE;
			useSystemFont = TRUE;
			verticalSizingGripsOnVerticalOrientation = FALSE;
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
		/// \brief <em>If \c TRUE, this instance is a customizer of the message hook used for the \c ReplaceMDIFrameMenu feature</em>
		UINT customizerOfMdiMessageHook : 1;

		Flags()
		{
			uiActivationPending = FALSE;
			usingThemes = FALSE;
			customizerOfMdiMessageHook = FALSE;
		}
	} flags;

	//////////////////////////////////////////////////////////////////////
	/// \name Band management
	///
	//@{
	/// \brief <em>Retrieves a new unique band ID at each call</em>
	///
	/// \return A new unique band ID.
	///
	/// \sa ReBarBand::get_ID
	UINT GetNewBandID(void);
	///@}
	//////////////////////////////////////////////////////////////////////


	/// \brief <em>Holds the zero-based index of the band below the mouse cursor</em>
	///
	/// \attention This member is not reliable with \c deMouseEvents being set.
	int bandUnderMouse;
	/// \brief <em>Holds the zero-based index of the band whose chevron popup window is currently displayed</em>
	int bandChevronPushed;

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
		/// \brief <em>Holds the index of the last clicked band</em>
		///
		/// Holds the index of the last clicked band. We use this to ensure that the \c *DblClick events
		/// are not raised accidently.
		///
		/// \attention This member is not reliable with \c deClickEvents being set.
		///
		/// \sa Raise_Click, Raise_DblClick, Raise_MClick, Raise_MDblClick, Raise_RClick, Raise_RDblClick
		int lastClickedBand;

		MouseStatus()
		{
			clickCandidates = 0;
			enteredControl = FALSE;
			hoveredControl = FALSE;
			lastClickedBand = -1;
		}

		/// \brief <em>Changes flags to indicate the \c MouseEnter event was just raised</em>
		///
		/// \sa enteredControl, HoverControl, LeaveControl
		void EnterControl(void)
		{
			RemoveAllClickCandidates();
			enteredControl = TRUE;
			lastClickedBand = -1;
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
			lastClickedBand = -1;
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

	/// \brief <em>Holds data and flags related to drag'n'drop</em>
	struct DragDropStatus
	{
		//////////////////////////////////////////////////////////////////////
		/// \name OLE Drag'n'Drop
		///
		//@{
		/// \brief <em>The currently dragged data</em>
		CComPtr<IOLEDataObject> pActiveDataObject;
		/// \brief <em>Holds the mouse cursors last position (in screen coordinates)</em>
		POINTL lastMousePosition;
		/// \brief <em>The \c IDropTargetHelper object used for drag image support</em>
		///
		/// \sa put_SupportOLEDragImages,
		///     <a href="https://msdn.microsoft.com/en-us/library/ms646238.aspx">IDropTargetHelper</a>
		IDropTargetHelper* pDropTargetHelper;
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
			pDropTargetHelper = NULL;
			drop_pDataObject = NULL;
		}

		~DragDropStatus()
		{
			if(pDropTargetHelper) {
				pDropTargetHelper->Release();
				pDropTargetHelper = NULL;
			}
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			if(this->pActiveDataObject) {
				this->pActiveDataObject = NULL;
			}
			drop_pDataObject = NULL;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragEnter is called</em>
		///
		/// \return An \c HRESULT error code.
		///
		/// \sa OLEDragLeaveOrDrop
		HRESULT OLEDragEnter(void)
		{
			return S_OK;
		}

		/// \brief <em>Performs any tasks that must be done if \c IDropTarget::DragLeave or \c IDropTarget::Drop is called</em>
		///
		/// \sa OLEDragEnter
		void OLEDragLeaveOrDrop(void)
		{
			//
		}
	} dragDropStatus;

	/// \brief <em>Holds data and flags related to MDI frame menu replacement</em>
	struct MDIStatus
	{
		/// \brief <em>If \c TRUE, the MDI frame window is active; otherwise not</em>
		UINT mdiFrameIsActive : 1;
		/// \brief <em>If \c TRUE, the currently active MDI child window is maximized; otherwise not</em>
		UINT activeMDIChildMaximized : 1;
		/// \brief <em>Holds the window icon of the currently active MDI child window</em>
		HICON hIconChildMaximized;
		/// \brief <em>Holds the handle of the currently active MDI child window</em>
		HWND hWndChildMaximized;
		/// \brief <em>The width (in pixels) required to display the MDI child window's icon</em>
		int mdiChildWindowIconWidth;
		/// \brief <em>The width (in pixels) required to display the MDI child window's system command buttons</em>
		int mdiChildWindowButtonAreaWidth;
		/// \brief <em>The size (in pixels) of a small icon according to the system settings</em>
		SIZE smallIconSize;
		/// \brief <em>The sizes (in pixels) of the MDI child window system command buttons</em>
		SIZE buttonSize[3];
		/// \brief <em>Holds the zero-based index of the MDI child window system command button currently being pressed</em>
		int pressedButton;
		/// \brief <em>Holds the zero-based index of the MDI child window system command button that has been pressed most recently</em>
		int lastPressedButton;

		MDIStatus()
		{
			mdiFrameIsActive = TRUE;
			activeMDIChildMaximized = FALSE;
			hIconChildMaximized = NULL;
			hWndChildMaximized = NULL;
			mdiChildWindowIconWidth = 20;
			mdiChildWindowButtonAreaWidth = 55;
			smallIconSize.cx = 16;
			smallIconSize.cy = 16;
			buttonSize[0].cx = 17;
			buttonSize[0].cy = 15;
			buttonSize[1] = buttonSize[0];
			buttonSize[2] = buttonSize[0];
			pressedButton = -1;
			lastPressedButton = -1;
		}

		/// \brief <em>Resets all member variables to their defaults</em>
		void Reset(void)
		{
			mdiFrameIsActive = TRUE;
			activeMDIChildMaximized = FALSE;
			hIconChildMaximized = NULL;
			hWndChildMaximized = NULL;
			mdiChildWindowIconWidth = 20;
			mdiChildWindowButtonAreaWidth = 55;
			smallIconSize.cx = 16;
			smallIconSize.cy = 16;
			buttonSize[0].cx = 17;
			buttonSize[0].cy = 15;
			buttonSize[1] = buttonSize[0];
			buttonSize[2] = buttonSize[0];
			pressedButton = -1;
			lastPressedButton = -1;
		}
	} mdiStatus;

	/// \brief <em>Holds IDs and intervals of timers that we use</em>
	///
	/// \sa OnTimer
	static struct Timers
	{
		/// \brief <em>The ID of the timer that is used to redraw the control window after recreation</em>
		static const UINT_PTR ID_REDRAW = 12;

		/// \brief <em>The interval of the timer that is used to redraw the control window after recreation</em>
		static const UINT INT_REDRAW = 10;
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
};     // ReBar

OBJECT_ENTRY_AUTO(__uuidof(ReBar), ReBar)