//////////////////////////////////////////////////////////////////////
/// \class ToolBarButton
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing tool bar button</em>
///
/// This class is a wrapper around a tool bar button that - unlike a button wrapped by
/// \c VirtualToolBarButton - really exists in the control.
///
/// \if UNICODE
///   \sa TBarCtlsLibU::IToolBarButton, VirtualToolBarButton, ToolBarButtons, ToolBar
/// \else
///   \sa TBarCtlsLibA::IToolBarButton, VirtualToolBarButton, ToolBarButtons, ToolBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "_IToolBarButtonEvents_CP.h"
#include "helpers.h"
#include "ToolBar.h"


class ATL_NO_VTABLE ToolBarButton : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ToolBarButton, &CLSID_ToolBarButton>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ToolBarButton>,
    public Proxy_IToolBarButtonEvents<ToolBarButton>, 
    #ifdef UNICODE
    	public IDispatchImpl<IToolBarButton, &IID_IToolBarButton, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IToolBarButton, &IID_IToolBarButton, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ToolBar;
	friend class ToolBarButtons;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TOOLBARBUTTON)

		BEGIN_COM_MAP(ToolBarButton)
			COM_INTERFACE_ENTRY(IToolBarButton)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ToolBarButton)
			CONNECTION_POINT_ENTRY(__uuidof(_IToolBarButtonEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()
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
	/// \name Implementation of IToolBarButton
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AutoSize property</em>
	///
	/// Retrieves whether the button's width is adjusted automatically depending on its text and icon. If
	/// set to \c VARIANT_TRUE, the button's width is adjusted automatically; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AutoSize, get_Width, get_Text, get_IconIndex
	virtual HRESULT STDMETHODCALLTYPE get_AutoSize(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutoSize property</em>
	///
	/// Sets whether the button's width is adjusted automatically depending on its text and icon. If
	/// set to \c VARIANT_TRUE, the button's width is adjusted automatically; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AutoSize, put_Width, put_Text, put_IconIndex
	virtual HRESULT STDMETHODCALLTYPE put_AutoSize(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ButtonData property</em>
	///
	/// Retrieves the \c LONG value associated with the button. Use this property to associate any data
	/// with the button.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ButtonData, ToolBar::Raise_FreeButtonData
	virtual HRESULT STDMETHODCALLTYPE get_ButtonData(LONG* pValue);
	/// \brief <em>Sets the \c ButtonData property</em>
	///
	/// Sets the \c LONG value associated with the button. Use this property to associate any data
	/// with the button.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ButtonData, ToolBar::Raise_FreeButtonData
	virtual HRESULT STDMETHODCALLTYPE put_ButtonData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ButtonType property</em>
	///
	/// Retrieves which kind of button this button is. Any of the values defined by the
	/// \c ButtonTypeConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ButtonType, get_PartOfGroup, TBarCtlsLibU::ButtonTypeConstants
	/// \else
	///   \sa put_ButtonType, get_PartOfGroup, TBarCtlsLibA::ButtonTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ButtonType(ButtonTypeConstants* pValue);
	/// \brief <em>Sets the \c ButtonType property</em>
	///
	/// Sets which kind of button this button is. Any of the values defined by the
	/// \c ButtonTypeConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ButtonType, put_PartOfGroup, TBarCtlsLibU::ButtonTypeConstants
	/// \else
	///   \sa get_ButtonType, put_PartOfGroup, TBarCtlsLibA::ButtonTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ButtonType(ButtonTypeConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DisplayText property</em>
	///
	/// Retrieves whether the button's text is displayed. If set to \c VARIANT_TRUE, the button's text is
	/// displayed next to the icon; otherwise it is not displayed next to the icon, but as tool tip if the
	/// mouse cursor is moved over the button.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c IToolBar::ButtonTextPosition property is not set to
	///          \c btpRightToIcon or the \c IToolBar::AlwaysDisplayButtonText property is set to
	///          \c VARIANT_TRUE.
	///
	/// \sa put_DisplayText, get_Text, ToolBar::get_AlwaysDisplayButtonText, ToolBar::get_ButtonTextPosition
	virtual HRESULT STDMETHODCALLTYPE get_DisplayText(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayText property</em>
	///
	/// Sets whether the button's text is displayed. If set to \c VARIANT_TRUE, the button's text is
	/// displayed next to the icon; otherwise it is not displayed next to the icon, but as tool tip if the
	/// mouse cursor is moved over the button.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c IToolBar::ButtonTextPosition property is not set to
	///          \c btpRightToIcon or the \c IToolBar::AlwaysDisplayButtonText property is set to
	///          \c VARIANT_TRUE.
	///
	/// \sa get_DisplayText, put_Text, ToolBar::put_AlwaysDisplayButtonText, ToolBar::put_ButtonTextPosition
	virtual HRESULT STDMETHODCALLTYPE put_DisplayText(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c DropDownStyle property</em>
	///
	/// Retrieves whether the button is a drop-down button and how the drop-down arrow is displayed. Any of
	/// the values defined by the \c DropDownStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DropDownStyle, ToolBar::get_NormalDropDownButtonStyle, get_DroppedDown,
	///       TBarCtlsLibU::DropDownStyleConstants
	/// \else
	///   \sa put_DropDownStyle, ToolBar::get_NormalDropDownButtonStyle, get_DroppedDown,
	///       TBarCtlsLibA::DropDownStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DropDownStyle(DropDownStyleConstants* pValue);
	/// \brief <em>Sets the \c DropDownStyle property</em>
	///
	/// Sets whether the button is a drop-down button and how the drop-down arrow is displayed. Any of
	/// the values defined by the \c DropDownStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DropDownStyle, ToolBar::put_NormalDropDownButtonStyle, get_DroppedDown,
	///       TBarCtlsLibU::DropDownStyleConstants
	/// \else
	///   \sa get_DropDownStyle, ToolBar::put_NormalDropDownButtonStyle, get_DroppedDown,
	///       TBarCtlsLibA::DropDownStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DropDownStyle(DropDownStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c DropDownStyle property</em>
	///
	/// Retrieves whether the button is currently dropped down. If this property is set to \c VARIANT_TRUE,
	/// the button is currently dropped down; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_DropDownStyle
	virtual HRESULT STDMETHODCALLTYPE get_DroppedDown(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the button is enabled or disabled for user input. If set to \c VARIANT_TRUE,
	/// it reacts to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled, ToolBar::get_Enabled, ToolBar::get_CommandEnabled
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Enables or disables the button for user input. If set to \c VARIANT_TRUE, the control reacts
	/// to user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled, ToolBar::put_Enabled, ToolBar::put_CommandEnabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FollowedByLineBreak property</em>
	///
	/// Retrieves whether the next button after this one is placed on a new line. If set to \c VARIANT_TRUE,
	/// the next button is placed on a new line; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FollowedByLineBreak, ToolBar::get_WrapButtons
	virtual HRESULT STDMETHODCALLTYPE get_FollowedByLineBreak(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c FollowedByLineBreak property</em>
	///
	/// Sets whether the next button after this one is placed on a new line. If set to \c VARIANT_TRUE,
	/// the next button is placed on a new line; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FollowedByLineBreak, ToolBar::put_WrapButtons
	virtual HRESULT STDMETHODCALLTYPE put_FollowedByLineBreak(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Hot property</em>
	///
	/// Retrieves whether the button is the control's hot button. The hot button is the button under the
	/// mouse cursor. If it is the hot button, this property is set to \c VARIANT_TRUE; otherwise it's set to
	/// \c VARIANT_FALSE.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Hot, get_Pushed, get_Marked, ToolBar::get_HotButton
	virtual HRESULT STDMETHODCALLTYPE get_Hot(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the button's icon in the control's image lists. If set to -1, the
	/// control will fire the \c ButtonGetDisplayInfo event each time this property's value is required. If
	/// set to -2, no icon is displayed for this button.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IconIndex, get_ImageListIndex, ToolBar::get_hImageList,
	///       ToolBar::Raise_ButtonGetDisplayInfo, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa put_IconIndex, get_ImageListIndex, ToolBar::get_hImageList,
	///       ToolBar::Raise_ButtonGetDisplayInfo, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Sets the \c IconIndex property</em>
	///
	/// Sets the zero-based index of the button's icon in the control's image lists. If set to -1, the
	/// control will fire the \c ButtonGetDisplayInfo event each time this property's value is required. If
	/// set to -2, no icon is displayed for this button.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IconIndex, put_ImageListIndex, ToolBar::put_hImageList,
	///       ToolBar::Raise_ButtonGetDisplayInfo, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa get_IconIndex, put_ImageListIndex, ToolBar::put_hImageList,
	///       ToolBar::Raise_ButtonGetDisplayInfo, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this tool bar button.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IToolBarButton interface.
	///
	/// \if UNICODE
	///   \sa put_ID, get_Index, TBarCtlsLibU::ButtonIdentifierTypeConstants
	/// \else
	///   \sa put_ID, get_Index, TBarCtlsLibA::ButtonIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Sets the \c ID property</em>
	///
	/// Sets an unique ID identifying this tool bar button.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This is the default property of the \c IToolBarButton interface.
	///
	/// \if UNICODE
	///   \sa get_ID, put_Index, TBarCtlsLibU::ButtonIdentifierTypeConstants
	/// \else
	///   \sa get_ID, put_Index, TBarCtlsLibA::ButtonIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ID(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ImageListIndex property</em>
	///
	/// Retrieves the zero-based index of the control's image lists that the button's icons will be taken
	/// from. If set to -1, the control will fire or has fired the \c ButtonGetDisplayInfo event each time
	/// this property's value is required.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ImageListIndex, get_IconIndex, ToolBar::get_hImageList,
	///       ToolBar::Raise_ButtonGetDisplayInfo, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa put_ImageListIndex, get_IconIndex, ToolBar::get_hImageList,
	///       ToolBar::Raise_ButtonGetDisplayInfo, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ImageListIndex(LONG* pValue);
	/// \brief <em>Sets the \c ImageListIndex property</em>
	///
	/// Sets the zero-based index of the control's image lists that the button's icons will be taken
	/// from. If set to -1, the control will fire or has fired the \c ButtonGetDisplayInfo event each time
	/// this property's value is required.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ImageListIndex, put_IconIndex, ToolBar::put_hImageList,
	///       ToolBar::Raise_ButtonGetDisplayInfo, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa get_ImageListIndex, put_IconIndex, ToolBar::put_hImageList,
	///       ToolBar::Raise_ButtonGetDisplayInfo, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ImageListIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index identifying this tool bar button.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing buttons changes other buttons' indexes, the index is the best
	///          (and fastest) option to identify a button.
	///
	/// \if UNICODE
	///   \sa put_Index, get_ID, TBarCtlsLibU::ButtonIdentifierTypeConstants
	/// \else
	///   \sa put_Index, get_ID, TBarCtlsLibA::ButtonIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Sets the \c Index property</em>
	///
	/// Sets the zero-based index identifying this tool bar button. Setting this property moves the button.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing buttons changes other buttons' indexes, the index is the best
	///          (and fastest) option to identify a button.
	///
	/// \if UNICODE
	///   \sa get_Index, put_ID, TBarCtlsLibU::ButtonIdentifierTypeConstants
	/// \else
	///   \sa get_Index, put_ID, TBarCtlsLibA::ButtonIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Index(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Marked property</em>
	///
	/// Retrieves whether the button is tagged as "Marked". It's up to the application how marked buttons are
	/// interpreted. If set to \c VARIANT_TRUE, the button is tagged as "Marked"; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Marked, get_Hot, get_SelectionState
	virtual HRESULT STDMETHODCALLTYPE get_Marked(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Marked property</em>
	///
	/// Sets whether the button is tagged as "Marked". It's up to the application how marked buttons are
	/// interpreted. If set to \c VARIANT_TRUE, the button is tagged as "Marked"; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Marked, put_SelectionState
	virtual HRESULT STDMETHODCALLTYPE put_Marked(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c PartOfChevronToolBar property</em>
	///
	/// Retrieves whether the button is a part of the chevron popup tool bar control. If this property is
	/// set to \c VARIANT_TRUE, the button is a part of the chevron popup tool bar control; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ShouldBeDisplayedInChevronPopup, ToolBar::DisplayChevronPopupWindow
	virtual HRESULT STDMETHODCALLTYPE get_PartOfChevronToolBar(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c PartOfGroup property</em>
	///
	/// Retrieves whether the button is part of a button group. If set to \c VARIANT_TRUE, it is part of a
	/// group; otherwise not.\n
	/// Among all consecutive buttons that are part of a group, only one button can be checked. A check
	/// button that is part of a group cannot be unchecked by simply clicking it a second time. This can be
	/// used to create option-button like behavior.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_PartOfGroup, get_ButtonType, get_SelectionState
	virtual HRESULT STDMETHODCALLTYPE get_PartOfGroup(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c PartOfGroup property</em>
	///
	/// Sets whether the button is part of a button group. If set to \c VARIANT_TRUE, it is part of a
	/// group; otherwise not.\n
	/// Among all consecutive buttons that are part of a group, only one button can be checked. A check
	/// button that is part of a group cannot be unchecked by simply clicking it a second time. This can be
	/// used to create option-button like behavior.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_PartOfGroup, put_ButtonType, put_SelectionState
	virtual HRESULT STDMETHODCALLTYPE put_PartOfGroup(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Pushed property</em>
	///
	/// Retrieves whether the button is drawn like being pushed. If set to \c VARIANT_TRUE, it is drawn like
	/// being pushed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Pushed, get_Hot, get_SelectionState
	virtual HRESULT STDMETHODCALLTYPE get_Pushed(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Pushed property</em>
	///
	/// Sets whether the button is drawn like being pushed. If set to \c VARIANT_TRUE, it is drawn like
	/// being pushed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Pushed, put_SelectionState
	virtual HRESULT STDMETHODCALLTYPE put_Pushed(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SelectionState property</em>
	///
	/// Retrieves the state of check buttons. Any of the values defined by the \c SelectionStateConstants
	/// enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SelectionState, get_ButtonType, get_Pushed,
	///       TBarCtlsLibU::SelectionStateConstants, ToolBar::Raise_ButtonSelectionStateChanged
	/// \else
	///   \sa put_SelectionState, get_ButtonType, get_Pushed,
	///       TBarCtlsLibA::SelectionStateConstants, ToolBar::Raise_ButtonSelectionStateChanged
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SelectionState(SelectionStateConstants* pValue);
	/// \brief <em>Sets the \c SelectionState property</em>
	///
	/// Sets the state of check buttons. Any of the values defined by the \c SelectionStateConstants
	/// enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SelectionState, put_ButtonType, put_Pushed, TBarCtlsLibU::SelectionStateConstants,
	///       ToolBar::Raise_ButtonSelectionStateChanged
	/// \else
	///   \sa get_SelectionState, put_ButtonType, put_Pushed, TBarCtlsLibA::SelectionStateConstants,
	///       ToolBar::Raise_ButtonSelectionStateChanged
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SelectionState(SelectionStateConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c PartOfChevronToolBar property</em>
	///
	/// Retrieves whether the button should be displayed in the chevron popup window. If this property is
	/// set to \c VARIANT_TRUE, the button should be displayed in the chevron popup window; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_PartOfChevronToolBar, ToolBar::DisplayChevronPopupWindow
	virtual HRESULT STDMETHODCALLTYPE get_ShouldBeDisplayedInChevronPopup(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ShowingEllipsis property</em>
	///
	/// Retrieves whether the button's text is cut off. If set to \c VARIANT_TRUE, the text is cut off and
	/// an ellipsis is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowingEllipsis, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_ShowingEllipsis(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowingEllipsis property</em>
	///
	/// Sets whether the button's text is cut off. If set to \c VARIANT_TRUE, the text is cut off and
	/// an ellipsis is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowingEllipsis, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_ShowingEllipsis(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the button's text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, get_IconIndex, get_UseMnemonic, get_ShowingEllipsis
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the button's text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, put_IconIndex, put_UseMnemonic, put_ShowingEllipsis
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c UseMnemonic property</em>
	///
	/// Retrieves whether the button has an accelerator prefix associated with it. If set to \c VARIANT_TRUE,
	/// the first character of the button's text that is prefixed by an ampersand (&amp;), is used as the
	/// button's accelerator prefix and this character is underlined. If set to \c VARIANT_FALSE, ampersands
	/// in the button's text are interpreted as normal text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseMnemonic, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_UseMnemonic(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseMnemonic property</em>
	///
	/// Sets whether the button has an accelerator prefix associated with it. If set to \c VARIANT_TRUE,
	/// the first character of the button's text that is prefixed by an ampersand (&amp;), is used as the
	/// button's accelerator prefix and this character is underlined. If set to \c VARIANT_FALSE, ampersands
	/// in the button's text are interpreted as normal text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseMnemonic, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_UseMnemonic(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Visible property</em>
	///
	/// Retrieves whether the button is visible. If set to \c VARIANT_TRUE, the button will be displayed;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Visible, ToolBarButtons::Remove, ToolBarButtons::Add
	virtual HRESULT STDMETHODCALLTYPE get_Visible(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Visible property</em>
	///
	/// Sets whether the button is visible. If set to \c VARIANT_TRUE, the button will be displayed;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Visible, ToolBarButtons::Remove, ToolBarButtons::Add
	virtual HRESULT STDMETHODCALLTYPE put_Visible(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Width property</em>
	///
	/// Retrieves the button's current width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Width, get_AutoSize
	virtual HRESULT STDMETHODCALLTYPE get_Width(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c Width property</em>
	///
	/// Sets the button's current width in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Width, put_AutoSize
	virtual HRESULT STDMETHODCALLTYPE put_Width(OLE_XSIZE_PIXELS newValue);

	/// \brief <em>Sets the window contained in the placeholder button</em>
	///
	/// Sets the handle of the window that is displayed at the position of the button if the button is a
	/// placeholder button.
	///
	/// \param[out] pHorizontalAlignment The horizontal alignment of the window inside the placeholder
	///             button's rectangle. Any of the values defined by the \c HAlignmentConstants enumeration
	///             is valid.
	/// \param[out] pVerticalAlignment The vertical alignment of the window inside the placeholder button's
	///             rectangle. Any of the values defined by the \c VAlignmentConstants enumeration is
	///             valid.
	/// \param[out] pHWnd The handle of the window that is displayed at the position of the placeholder
	///             button.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa SetContainedWindow, get_ButtonType, TBarCtlsLibU::HAlignmentConstants,
	///       TBarCtlsLibU::VAlignmentConstants
	/// \else
	///   \sa SetContainedWindow, get_ButtonType, TBarCtlsLibA::HAlignmentConstants,
	///       TBarCtlsLibA::VAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetContainedWindow(HAlignmentConstants* pHorizontalAlignment = NULL, VAlignmentConstants* pVerticalAlignment = NULL, OLE_HANDLE* pHWnd = NULL);
	/// \brief <em>Retrieves the bounding rectangle of either the button or a part of it</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's client area) of either the
	/// button or a part of it.
	///
	/// \param[in] rectangleType The rectangle to retrieve. Any of the values defined by the
	///            \c ButtonRectangleTypeConstants enumeration is valid.
	/// \param[out] PXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
	///             relative to the control's upper-left corner.
	/// \param[out] pYTop The y-coordinate (in pixels) of the bounding rectangle's top border
	///             relative to the control's upper-left corner.
	/// \param[out] pXRight The x-coordinate (in pixels) of the bounding rectangle's right border
	///             relative to the control's upper-left corner.
	/// \param[out] pYBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border
	///             relative to the control's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::ButtonRectangleTypeConstants
	/// \else
	///   \sa TBarCtlsLibA::ButtonRectangleTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(ButtonRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Sets the window contained in the placeholder button</em>
	///
	/// Sets the handle of the window that is displayed at the position of the button if the button is a
	/// placeholder button.
	///
	/// \param[in] hWnd The handle of the window that is positioned at the position of the placeholder
	///            button.
	/// \param[in] horizontalAlignment The horizontal alignment of the window inside the placeholder
	///            button's rectangle. Any of the values defined by the \c HAlignmentConstants enumeration
	///            is valid.
	/// \param[in] verticalAlignment The vertical alignment of the window inside the placeholder button's
	///            rectangle. Any of the values defined by the \c VAlignmentConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa GetContainedWindow, get_ButtonType, TBarCtlsLibU::HAlignmentConstants,
	///       TBarCtlsLibU::VAlignmentConstants
	/// \else
	///   \sa GetContainedWindow, get_ButtonType, TBarCtlsLibA::HAlignmentConstants,
	///       TBarCtlsLibA::VAlignmentConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE SetContainedWindow(OLE_HANDLE hWnd, HAlignmentConstants horizontalAlignment = halCenter, VAlignmentConstants verticalAlignment = valCenter);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given button</em>
	///
	/// Attaches this object to a given button, so that the button's properties can be retrieved and set
	/// using this object's methods.
	///
	/// \param[in] buttonIndex The button to attach to.
	/// \param[in] partOfChevronPopup If \c TRUE, the button is part of the chevron popup tool bar control.
	///            In this case all methods of this object will work on the chevron popup tool bar instead
	///            of the main tool bar.
	///
	/// \sa Detach, ToolBar::DisplayChevronPopupWindow, ToolBar::get_hWndChevronToolBar
	void Attach(int buttonIndex, BOOL partOfChevronPopup);
	/// \brief <em>Detaches this object from a button</em>
	///
	/// Detaches this object from the button it currently wraps, so that it doesn't wrap any button anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// Applies the settings from a given source to the button wrapped by this object.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(LPTBBUTTON pSource);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// \overload
	HRESULT LoadState(VirtualToolBarButton* pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndTBar The rebar window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(LPTBBUTTON pTarget, HWND hWndTBar = NULL);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \overload
	HRESULT SaveState(VirtualToolBarButton* pTarget);
	/// \brief <em>Sets the owner of this button</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerTBar
	void SetOwner(__in_opt ToolBar* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ToolBar object that owns this button</em>
		///
		/// \sa SetOwner
		ToolBar* pOwnerTBar;
		/// \brief <em>The index of the button wrapped by this object</em>
		int buttonIndex;
		/// \brief <em>If \c TRUE, the button is part of the chevron popup tool bar control and all methods of this object will work on the chevron popup tool bar</em>
		BOOL partOfChevronPopup;

		Properties()
		{
			pOwnerTBar = NULL;
			buttonIndex = -1;
			partOfChevronPopup = FALSE;
		}

		~Properties();

		/// \brief <em>Retrieves the owning tool bar's window handle</em>
		///
		/// \return The window handle of the tool bar that contains this button.
		///
		/// \sa pOwnerTBar
		HWND GetTBarHWnd(void);
	} properties;

	/// \brief <em>Writes a given object's settings to a given target</em>
	///
	/// \param[in] buttonIndex The button for which to save the settings.
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndTBar The tool bar window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(int buttonIndex, LPTBBUTTON pTarget, HWND hWndTBar = NULL);
};     // ToolBarButton

OBJECT_ENTRY_AUTO(__uuidof(ToolBarButton), ToolBarButton)