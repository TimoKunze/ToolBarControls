//////////////////////////////////////////////////////////////////////
/// \class VirtualToolBarButton
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing tool bar button</em>
///
/// This class is a wrapper around a tool bar button that does not yet or not anymore exist in the control.
///
/// \if UNICODE
///   \sa TBarCtlsLibU::IVirtualToolBarButton, ToolBarButton, ToolBar
/// \else
///   \sa TBarCtlsLibA::IVirtualToolBarButton, ToolBarButton, ToolBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "_IVirtualToolBarButtonEvents_CP.h"
#include "helpers.h"
#include "ToolBar.h"


class ATL_NO_VTABLE VirtualToolBarButton : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualToolBarButton, &CLSID_VirtualToolBarButton>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualToolBarButton>,
    public Proxy_IVirtualToolBarButtonEvents< VirtualToolBarButton>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualToolBarButton, &IID_IVirtualToolBarButton, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualToolBarButton, &IID_IVirtualToolBarButton, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ToolBar;
	friend class ToolBarButton;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALTOOLBARBUTTON)

		BEGIN_COM_MAP(VirtualToolBarButton)
			COM_INTERFACE_ENTRY(IVirtualToolBarButton)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualToolBarButton)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualToolBarButtonEvents))
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
	/// \name Implementation of IVirtualToolBarButton
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AutoSize property</em>
	///
	/// Retrieves whether the button's width will be or was adjusted automatically depending on its text and
	/// icon. If set to \c VARIANT_TRUE, the button's width will be or was adjusted automatically; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AutoSize, get_Width, get_Text, get_IconIndex
	virtual HRESULT STDMETHODCALLTYPE get_AutoSize(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AutoSize property</em>
	///
	/// Sets whether the button's width will be or was adjusted automatically depending on its text and
	/// icon. If set to \c VARIANT_TRUE, the button's width will be or was adjusted automatically; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AutoSize, put_Width, put_Text, put_IconIndex
	virtual HRESULT STDMETHODCALLTYPE put_AutoSize(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ButtonData property</em>
	///
	/// Retrieves the \c LONG value that will be or was associated with the button. Use this property to
	/// associate any data with the button.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ButtonData, ToolBar::Raise_FreeButtonData
	virtual HRESULT STDMETHODCALLTYPE get_ButtonData(LONG* pValue);
	/// \brief <em>Sets the \c ButtonData property</em>
	///
	/// Sets the \c LONG value that will be or was associated with the button. Use this property to
	/// associate any data with the button.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ButtonData, ToolBar::Raise_FreeButtonData
	virtual HRESULT STDMETHODCALLTYPE put_ButtonData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ButtonType property</em>
	///
	/// Retrieves which kind of button this button will be or was. Any of the values defined by the
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
	/// Sets which kind of button this button will be or was. Any of the values defined by the
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
	/// Retrieves whether the button's text will be or was displayed. If set to \c VARIANT_TRUE, the button's
	/// text will be or was displayed next to the icon; otherwise it will be or was not displayed next to the
	/// icon, but as tool tip if the mouse cursor is moved over the button.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is ignored, if the \c IToolBar::ButtonTextPosition is not set to
	///          \c btpRightToIcon or the \c IToolBar::AlwaysDisplayButtonText is set to \c VARIANT_TRUE.
	///
	/// \sa put_DisplayText, get_Text, ToolBar::get_AlwaysDisplayButtonText, ToolBar::get_ButtonTextPosition
	virtual HRESULT STDMETHODCALLTYPE get_DisplayText(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c DisplayText property</em>
	///
	/// Sets whether the button's text will be or was displayed. If set to \c VARIANT_TRUE, the button's
	/// text will be or was displayed next to the icon; otherwise it will be or was not displayed next to the
	/// icon, but as tool tip if the mouse cursor is moved over the button.
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
	/// Retrieves whether the button will be or was a drop-down button and how the drop-down arrow will be or
	/// was displayed. Any of the values defined by the \c DropDownStyleConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_DropDownStyle, ToolBar::get_NormalDropDownButtonStyle, TBarCtlsLibU::DropDownStyleConstants
	/// \else
	///   \sa put_DropDownStyle, ToolBar::get_NormalDropDownButtonStyle, TBarCtlsLibA::DropDownStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_DropDownStyle(DropDownStyleConstants* pValue);
	/// \brief <em>Sets the \c DropDownStyle property</em>
	///
	/// Sets whether the button will be or was a drop-down button and how the drop-down arrow will be or
	/// was displayed. Any of the values defined by the \c DropDownStyleConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_DropDownStyle, ToolBar::put_NormalDropDownButtonStyle, TBarCtlsLibU::DropDownStyleConstants
	/// \else
	///   \sa get_DropDownStyle, ToolBar::put_NormalDropDownButtonStyle, TBarCtlsLibA::DropDownStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_DropDownStyle(DropDownStyleConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Enabled property</em>
	///
	/// Retrieves whether the button will be or was enabled or disabled for user input. If set to
	/// \c VARIANT_TRUE, it will react or has reacted to user input; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Enabled, ToolBar::get_Enabled
	virtual HRESULT STDMETHODCALLTYPE get_Enabled(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Enabled property</em>
	///
	/// Sets whether the button will be or was enabled or disabled for user input. If set to
	/// \c VARIANT_TRUE, it will react or has reacted to user input; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Enabled, ToolBar::put_Enabled
	virtual HRESULT STDMETHODCALLTYPE put_Enabled(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c FollowedByLineBreak property</em>
	///
	/// Retrieves whether the next button after this one will be or was placed on a new line. If set to
	/// \c VARIANT_TRUE, the next button will be or was placed on a new line; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FollowedByLineBreak, ToolBar::get_WrapButtons
	virtual HRESULT STDMETHODCALLTYPE get_FollowedByLineBreak(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c FollowedByLineBreak property</em>
	///
	/// Sets whether the next button after this one will be or was placed on a new line. If set to
	/// \c VARIANT_TRUE, the next button will be or was placed on a new line; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FollowedByLineBreak, ToolBar::put_WrapButtons
	virtual HRESULT STDMETHODCALLTYPE put_FollowedByLineBreak(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the button's icon in the control's image lists. If set to -1, the
	/// control will fire or has fired the \c ButtonGetDisplayInfo event each time this property's value is
	/// required. If set to -2, no icon will be or was displayed for this button.
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
	/// control will fire or has fired the \c ButtonGetDisplayInfo event each time this property's value is
	/// required. If set to -2, no icon will be or was displayed for this button.
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
	/// \remarks This is the default property of the \c IVirtualToolBarButton interface.
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
	///   \sa get_ID, TBarCtlsLibU::ButtonIdentifierTypeConstants
	/// \else
	///   \sa get_ID, TBarCtlsLibA::ButtonIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ID(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ImageListIndex property</em>
	///
	/// Retrieves the zero-based index of the control's image lists that the button's icons will be or has
	/// been taken from. If set to -1, the control will fire or has fired the \c ButtonGetDisplayInfo event
	/// each time this property's value will be or was required.
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
	/// Sets the zero-based index of the control's image lists that the button's icons will be or has
	/// been taken from. If set to -1, the control will fire or has fired the \c ButtonGetDisplayInfo event
	/// each time this property's value will be or was required.
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
	///          (and fastest) option to identify a button.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ID, TBarCtlsLibU::ButtonIdentifierTypeConstants
	/// \else
	///   \sa get_ID, TBarCtlsLibA::ButtonIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c Marked property</em>
	///
	/// Retrieves whether the button will be or was tagged as "Marked". It's up to the application how marked
	/// buttons are interpreted. If set to \c VARIANT_TRUE, the button will be or was tagged as "Marked";
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Marked, get_SelectionState
	virtual HRESULT STDMETHODCALLTYPE get_Marked(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Marked property</em>
	///
	/// Sets whether the button will be or was tagged as "Marked". It's up to the application how marked
	/// buttons are interpreted. If set to \c VARIANT_TRUE, the button will be or was tagged as "Marked";
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Marked, put_SelectionState
	virtual HRESULT STDMETHODCALLTYPE put_Marked(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c PartOfGroup property</em>
	///
	/// Retrieves whether the button will be or was part of a button group. If set to \c VARIANT_TRUE, it will
	/// be or was part of a group; otherwise not.\n
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
	/// Sets whether the button will be or was part of a button group. If set to \c VARIANT_TRUE, it will
	/// be or was part of a group; otherwise not.\n
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
	/// Retrieves whether the button will be or was drawn like being pushed. If set to \c VARIANT_TRUE, it
	/// will be or was drawn like being pushed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Pushed, get_SelectionState
	virtual HRESULT STDMETHODCALLTYPE get_Pushed(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Pushed property</em>
	///
	/// Sets whether the button will be or was drawn like being pushed. If set to \c VARIANT_TRUE, it
	/// will be or was drawn like being pushed; otherwise not.
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
	///   \sa put_SelectionState, get_ButtonType, get_Pushed, TBarCtlsLibU::SelectionStateConstants,
	///       ToolBar::Raise_ButtonSelectionStateChanged
	/// \else
	///   \sa put_SelectionState, get_ButtonType, get_Pushed, TBarCtlsLibA::SelectionStateConstants,
	///       ToolBar::Raise_ButtonSelectionStateChanged
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
	/// \brief <em>Retrieves the current setting of the \c ShowingEllipsis property</em>
	///
	/// Retrieves whether the button's text will be or was cut off. If set to \c VARIANT_TRUE, the text will
	/// be or was cut off and an ellipsis will be or was displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowingEllipsis, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_ShowingEllipsis(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowingEllipsis property</em>
	///
	/// Sets whether the button's text will be or was cut off. If set to \c VARIANT_TRUE, the text will
	/// be or was cut off and an ellipsis will be or was displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowingEllipsis, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_ShowingEllipsis(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the button's text. The maximum number of characters in this text is specified by
	/// \c MAX_BUTTONTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, MAX_BUTTONTEXTLENGTH, get_IconIndex, get_UseMnemonic, get_ShowingEllipsis
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the button's text. The maximum number of characters in this text is specified by
	/// \c MAX_BUTTONTEXTLENGTH.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, MAX_BUTTONTEXTLENGTH, put_IconIndex, put_UseMnemonic, put_ShowingEllipsis
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c UseMnemonic property</em>
	///
	/// Retrieves whether the button has or had an accelerator prefix associated with it. If set to
	/// \c VARIANT_TRUE, the first character of the button's text that is prefixed by an ampersand (&amp;),
	/// will be or was used as the button's accelerator prefix and this character will be or was underlined.
	/// If set to \c VARIANT_FALSE, ampersands in the button's text will be or have been interpreted as
	/// normal text.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseMnemonic, get_Text
	virtual HRESULT STDMETHODCALLTYPE get_UseMnemonic(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseMnemonic property</em>
	///
	/// Sets whether the button has or had an accelerator prefix associated with it. If set to
	/// \c VARIANT_TRUE, the first character of the button's text that is prefixed by an ampersand (&amp;),
	/// will be or was used as the button's accelerator prefix and this character will be or was underlined.
	/// If set to \c VARIANT_FALSE, ampersands in the button's text will be or have been interpreted as
	/// normal text.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseMnemonic, put_Text
	virtual HRESULT STDMETHODCALLTYPE put_UseMnemonic(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Visible property</em>
	///
	/// Retrieves whether the button will be or was visible. If set to \c VARIANT_TRUE, the button will be or
	/// was displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Visible, ToolBarButtons::Remove, ToolBarButtons::Add
	virtual HRESULT STDMETHODCALLTYPE get_Visible(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Visible property</em>
	///
	/// Sets whether the button will be or was visible. If set to \c VARIANT_TRUE, the button will be or
	/// was displayed; otherwise not.
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
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Initializes this object with given data</em>
	///
	/// Initializes this object with the settings from a given source.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	/// \param[in] buttonIndex The button's zero-based index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(__in LPTBBUTTON pSource, int buttonIndex);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[out] buttonIndex The button's zero-based index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(__in LPTBBUTTON pTarget, int& buttonIndex);
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
		/// \brief <em>A structure holding this button's settings</em>
		TBBUTTON settings;
		/// \brief <em>The button's zero-based index</em>
		int buttonIndex;
		/// \brief <em>If \c TRUE, the string specified by \c settings.iString will be freed automatically</em>
		UINT freeString : 1;
		/// \brief <em>If \c TRUE, all properties will be read-only</em>
		UINT readOnly : 1;

		Properties()
		{
			pOwnerTBar = NULL;
			ZeroMemory(&settings, sizeof(settings));
			buttonIndex = -1;
			freeString = FALSE;
			readOnly = TRUE;
		}

		~Properties();

		/// \brief <em>Retrieves the owning tool bar's window handle</em>
		///
		/// \return The window handle of the tool bar that contains this button.
		///
		/// \sa pOwnerTBar
		HWND GetTBarHWnd(void);
	} properties;
};     // VirtualToolBarButton

OBJECT_ENTRY_AUTO(__uuidof(VirtualToolBarButton), VirtualToolBarButton)