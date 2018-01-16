//////////////////////////////////////////////////////////////////////
/// \class Proxy_IToolBarEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IToolBarEvents interface</em>
///
/// \if UNICODE
///   \sa ToolBar, TBarCtlsLibU::_IToolBarEvents
/// \else
///   \sa ToolBar, TBarCtlsLibA::_IToolBarEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IToolBarEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IToolBarEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c BeforeDisplayChevronPopup event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::BeforeDisplayChevronPopup,
	///       ToolBar::Raise_BeforeDisplayChevronPopup, Fire_DestroyingChevronPopup
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::BeforeDisplayChevronPopup,
	///       ToolBar::Raise_BeforeDisplayChevronPopup, Fire_DestroyingChevronPopup
	/// \endif
	HRESULT Fire_BeforeDisplayChevronPopup(LONG hPopup, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL isMenu, VARIANT_BOOL* pCancelPopup, LONG* pCommandToExecute)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = hPopup;
				p[4] = x;													p[4].vt = VT_XPOS_PIXELS;
				p[3] = y;													p[3].vt = VT_YPOS_PIXELS;
				p[2] = isMenu;
				p[1].pboolVal = pCancelPopup;			p[1].vt = VT_BOOL | VT_BYREF;
				p[0].plVal = pCommandToExecute;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_BEFOREDISPLAYCHEVRONPOPUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BeginCustomization event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::BeginCustomization, ToolBar::Raise_BeginCustomization,
	///       Fire_InitializeCustomizationDialog, Fire_EndCustomization, Fire_QueryInsertButton,
	///       Fire_QueryRemoveButton, Fire_InitializeToolBarStateRegistryRestorage
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::BeginCustomization, ToolBar::Raise_BeginCustomization,
	///       Fire_InitializeCustomizationDialog, Fire_EndCustomization, Fire_QueryInsertButton,
	///       Fire_QueryRemoveButton, Fire_InitializeToolBarStateRegistryRestorage
	/// \endif
	HRESULT Fire_BeginCustomization(VARIANT_BOOL restoringFromRegistry, VARIANT_BOOL* pDontRestoreFromRegistry)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = restoringFromRegistry;
				p[0].pboolVal = pDontRestoreFromRegistry;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_BEGINCUSTOMIZATION, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ButtonBeginDrag event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::ButtonBeginDrag, ToolBar::Raise_ButtonBeginDrag,
	///       Fire_ButtonBeginRDrag
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ButtonBeginDrag, ToolBar::Raise_ButtonBeginDrag,
	///       Fire_ButtonBeginRDrag
	/// \endif
	HRESULT Fire_ButtonBeginDrag(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_BUTTONBEGINDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ButtonBeginRDrag event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::ButtonBeginRDrag, ToolBar::Raise_ButtonBeginRDrag,
	///       Fire_ButtonBeginDrag
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ButtonBeginRDrag, ToolBar::Raise_ButtonBeginRDrag,
	///       Fire_ButtonBeginDrag
	/// \endif
	HRESULT Fire_ButtonBeginRDrag(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_BUTTONBEGINRDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ButtonGetDisplayInfo event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::ButtonGetDisplayInfo, ToolBar::Raise_ButtonGetDisplayInfo
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ButtonGetDisplayInfo, ToolBar::Raise_ButtonGetDisplayInfo
	/// \endif
	HRESULT Fire_ButtonGetDisplayInfo(IToolBarButton* pToolButton, RequestedInfoConstants requestedInfo, LONG* pIconIndex, LONG* pImageListIndex, LONG maxButtonTextLength, BSTR* pButtonText, VARIANT_BOOL* pDontAskAgain)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pToolButton;
				p[5] = static_cast<LONG>(requestedInfo);		p[5].vt = VT_I4;
				p[4].plVal = pIconIndex;										p[4].vt = VT_I4 | VT_BYREF;
				p[3].plVal = pImageListIndex;								p[3].vt = VT_I4 | VT_BYREF;
				p[2] = maxButtonTextLength;
				p[1].pbstrVal = pButtonText;								p[1].vt = VT_BSTR | VT_BYREF;
				p[0].pboolVal = pDontAskAgain;							p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_TBE_BUTTONGETDISPLAYINFO, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ButtonGetDropDownMenu event</em>
	///
	/// \param[in] pToolButton The button that the menu is required for.
	/// \param[out] phMenu Receives the handle of the popup menu.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::ButtonGetDropDownMenu, ToolBar::Raise_ButtonGetDropDownMenu,
	///       Fire_DropDown
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ButtonGetDropDownMenu, ToolBar::Raise_ButtonGetDropDownMenu,
	///       Fire_DropDown
	/// \endif
	HRESULT Fire_ButtonGetDropDownMenu(IToolBarButton* pToolButton, LONG* phMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pToolButton;
				p[0].plVal = reinterpret_cast<PLONG>(phMenu);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_BUTTONGETDROPDOWNMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ButtonGetInfoTipText event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::ButtonGetInfoTipText, ToolBar::Raise_ButtonGetInfoTipText
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ButtonGetInfoTipText, ToolBar::Raise_ButtonGetInfoTipText
	/// \endif
	HRESULT Fire_ButtonGetInfoTipText(IToolBarButton* pToolButton, LONG maxInfoTipLength, BSTR* pInfoTipText, VARIANT_BOOL* pAbortToolTip)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pToolButton;
				p[2] = maxInfoTipLength;
				p[1].pbstrVal = pInfoTipText;			p[1].vt = VT_BSTR | VT_BYREF;
				p[0].pboolVal = pAbortToolTip;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TBE_BUTTONGETINFOTIPTEXT, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ButtonMouseEnter event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::ButtonMouseEnter, ToolBar::Raise_ButtonMouseEnter,
	///       Fire_ButtonMouseLeave, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ButtonMouseEnter, ToolBar::Raise_ButtonMouseEnter,
	///       Fire_ButtonMouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_ButtonMouseEnter(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_BUTTONMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ButtonMouseLeave event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::ButtonMouseLeave, ToolBar::Raise_ButtonMouseLeave,
	///       Fire_ButtonMouseEnter, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ButtonMouseLeave, ToolBar::Raise_ButtonMouseLeave,
	///       Fire_ButtonMouseEnter, Fire_MouseMove
	/// \endif
	HRESULT Fire_ButtonMouseLeave(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_BUTTONMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ButtonSelectionStateChanged event</em>
	///
	/// \param[in] pToolButton The tool bar button for which the selection state was changed.
	/// \param[in] previousSelectionState The tool bar button's previous selection state.
	/// \param[in] newSelectionState The tool bar button's new selection state.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::ButtonSelectionStateChanged,
	///       ToolBar::Raise_ButtonSelectionStateChanged, Fire_ExecuteCommand
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ButtonSelectionStateChanged,
	///       ToolBar::Raise_ButtonSelectionStateChanged, Fire_ExecuteCommand
	/// \endif
	HRESULT Fire_ButtonSelectionStateChanged(IToolBarButton* pToolButton, SelectionStateConstants previousSelectionState, SelectionStateConstants newSelectionState)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pToolButton;
				p[1] = previousSelectionState;
				p[0] = newSelectionState;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_TBE_BUTTONSELECTIONSTATECHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Click event</em>
	///
	/// \param[in] pToolButton The clicked tool bar button. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::Click, ToolBar::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::Click, ToolBar::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_Click(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_CLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ContextMenu event</em>
	///
	/// \param[in] pToolButton The tool bar button that the context menu refers to. May be \c NULL.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the menu's proposed position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that the menu's proposed position lies in.
	///            Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::ContextMenu, ToolBar::Raise_ContextMenu, Fire_RClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ContextMenu, ToolBar::Raise_ContextMenu, Fire_RClick
	/// \endif
	HRESULT Fire_ContextMenu(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_CONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CustomDraw event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::CustomDraw, ToolBar::Raise_CustomDraw
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::CustomDraw, ToolBar::Raise_CustomDraw
	/// \endif
	HRESULT Fire_CustomDraw(IToolBarButton* pToolButton, OLE_COLOR* pNormalTextColor, OLE_COLOR* pNormalButtonBackColor, StringBackgroundModeConstants* pNormalBackgroundMode, OLE_COLOR* pHotTextColor, OLE_COLOR* pHotButtonBackColor, OLE_COLOR* pMarkedTextBackColor, OLE_COLOR* pMarkedButtonBackColor, StringBackgroundModeConstants* pMarkedBackgroundMode, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants buttonState, LONG hDC, RECTANGLE* pDrawingRectangle, RECTANGLE* pTextRectangle, OLE_XSIZE_PIXELS* pHorizontalIconCaptionGap, CustomDrawReturnValuesConstants* pFurtherProcessing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[16];
				p[15] = pToolButton;
				p[14].plVal = reinterpret_cast<PLONG>(pNormalTextColor);					p[14].vt = VT_I4 | VT_BYREF;
				p[13].plVal = reinterpret_cast<PLONG>(pNormalButtonBackColor);		p[13].vt = VT_I4 | VT_BYREF;
				p[12].plVal = reinterpret_cast<PLONG>(pNormalBackgroundMode);			p[12].vt = VT_I4 | VT_BYREF;
				p[11].plVal = reinterpret_cast<PLONG>(pHotTextColor);							p[11].vt = VT_I4 | VT_BYREF;
				p[10].plVal = reinterpret_cast<PLONG>(pHotButtonBackColor);				p[10].vt = VT_I4 | VT_BYREF;
				p[9].plVal = reinterpret_cast<PLONG>(pMarkedTextBackColor);				p[9].vt = VT_I4 | VT_BYREF;
				p[8].plVal = reinterpret_cast<PLONG>(pMarkedButtonBackColor);			p[8].vt = VT_I4 | VT_BYREF;
				p[7].plVal = reinterpret_cast<PLONG>(pMarkedBackgroundMode);			p[7].vt = VT_I4 | VT_BYREF;
				p[6].lVal = static_cast<LONG>(drawStage);													p[6].vt = VT_I4;
				p[5].lVal = static_cast<LONG>(buttonState);												p[5].vt = VT_I4;
				p[4] = hDC;
				p[1].plVal = static_cast<PLONG>(pHorizontalIconCaptionGap);				p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = reinterpret_cast<PLONG>(pFurtherProcessing);					p[0].vt = VT_I4 | VT_BYREF;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo1 = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{AC1E3771-638F-4e25-B405-682C2A2389D1}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo1)));
				#else
					LPOLESTR clsid = OLESTR("{4111318D-17B8-4728-9AC5-C8B3C6A6498F}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo1)));
				#endif
				VariantInit(&p[3]);
				p[3].vt = VT_RECORD | VT_BYREF;
				p[3].pRecInfo = pRecordInfo1;
				p[3].pvRecord = pRecordInfo1->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[3].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[3].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[3].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[3].pvRecord)->Top = pDrawingRectangle->Top;

				// pack the pTextRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo2 = NULL;
				#ifdef UNICODE
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo2)));
				#else
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo2)));
				#endif
				VariantInit(&p[2]);
				p[2].vt = VT_RECORD | VT_BYREF;
				p[2].pRecInfo = pRecordInfo2;
				p[2].pvRecord = pRecordInfo2->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Bottom = pTextRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Left = pTextRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Right = pTextRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Top = pTextRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 16, 0};
				hr = pConnection->Invoke(DISPID_TBE_CUSTOMDRAW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(SUCCEEDED(hr)) {
					pTextRectangle->Bottom = reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Bottom;
					pTextRectangle->Left = reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Left;
					pTextRectangle->Right = reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Right;
					pTextRectangle->Top = reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Top;
				}

				if(pRecordInfo1) {
					pRecordInfo1->RecordDestroy(p[3].pvRecord);
				}
				if(pRecordInfo2) {
					pRecordInfo2->RecordDestroy(p[2].pvRecord);
				}
			}
		}

		/* Although the p*Color parameters are of type OLE_COLOR, they may contain RGB colors only. So convert
		   OLE_COLOR colors into RGB colors. */
		if(*pNormalTextColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pNormalTextColor, NULL, &color);
			*pNormalTextColor = static_cast<OLE_COLOR>(color);
		}
		if(*pNormalButtonBackColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pNormalButtonBackColor, NULL, &color);
			*pNormalButtonBackColor = static_cast<OLE_COLOR>(color);
		}
		if(*pHotTextColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pHotTextColor, NULL, &color);
			*pHotTextColor = static_cast<OLE_COLOR>(color);
		}
		if(*pHotButtonBackColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pHotButtonBackColor, NULL, &color);
			*pHotButtonBackColor = static_cast<OLE_COLOR>(color);
		}
		if(*pMarkedTextBackColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pMarkedTextBackColor, NULL, &color);
			*pMarkedTextBackColor = static_cast<OLE_COLOR>(color);
		}
		if(*pMarkedButtonBackColor & 0x80000000) {
			COLORREF color = RGB(0x00, 0x00, 0x00);
			OleTranslateColor(*pMarkedButtonBackColor, NULL, &color);
			*pMarkedButtonBackColor = static_cast<OLE_COLOR>(color);
		}

		return hr;
	}

	/// \brief <em>Fires the \c CustomizedControl event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::CustomizedControl, ToolBar::Raise_CustomizedControl,
	///       Fire_BeginCustomization, Fire_EndCustomization
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::CustomizedControl, ToolBar::Raise_CustomizedControl,
	///       Fire_BeginCustomization, Fire_EndCustomization
	/// \endif
	HRESULT Fire_CustomizedControl(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_TBE_CUSTOMIZEDCONTROL, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CustomizeDialogRemovingButton event</em>
	///
	/// \param[in] pToolButton The tool bar button that is being removed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::CustomizeDialogRemovingButton,
	///       ToolBar::Raise_CustomizeDialogRemovingButton, Fire_FreeButtonData, Fire_QueryRemoveButton
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::CustomizeDialogRemovingButton,
	///       ToolBar::Raise_CustomizeDialogRemovingButton, Fire_FreeButtonData, Fire_QueryRemoveButton
	/// \endif
	HRESULT Fire_CustomizeDialogRemovingButton(IToolBarButton* pToolButton)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pToolButton;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_CUSTOMIZEDIALOGREMOVINGBUTTON, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DblClick event</em>
	///
	/// \param[in] pToolButton The double-clicked tool bar button. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::DblClick, ToolBar::Raise_DblClick, Fire_Click, Fire_MDblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::DblClick, ToolBar::Raise_DblClick, Fire_Click, Fire_MDblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_DblClick(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_DBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::DestroyedControlWindow,
	///       ToolBar::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::DestroyedControlWindow,
	///       ToolBar::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \endif
	HRESULT Fire_DestroyedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_DESTROYEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DestroyingChevronPopup event</em>
	///
	/// \param[in] hPopup The handle of popup. In menu mode, this is the handle of the chevron popup
	///            menu; otherwise this is the window handle of the popup tool bar.
	/// \param[in] isMenu If \c VARIANT_TRUE, the handle specified by \c hPopup is a menu handle; otherwise
	///            it is a window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::DestroyingChevronPopup, ToolBar::Raise_DestroyingChevronPopup,
	///       Fire_BeforeDisplayChevronPopup
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::DestroyingChevronPopup, ToolBar::Raise_DestroyingChevronPopup,
	///       Fire_BeforeDisplayChevronPopup
	/// \endif
	HRESULT Fire_DestroyingChevronPopup(LONG hPopup, VARIANT_BOOL isMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = hPopup;
				p[0] = isMenu;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_DESTROYINGCHEVRONPOPUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DisplayCustomizationHelp event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::DisplayCustomizationHelp,
	///       ToolBar::Raise_DisplayCustomizationHelp, Fire_InitializeCustomizationDialog
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::DisplayCustomizationHelp,
	///       ToolBar::Raise_DisplayCustomizationHelp, Fire_InitializeCustomizationDialog
	/// \endif
	HRESULT Fire_DisplayCustomizationHelp(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_DISPLAYCUSTOMIZATIONHELP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DropDown event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::DropDown, ToolBar::Raise_DropDown, Fire_ExecuteCommand,
	///       Fire_Click
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::DropDown, ToolBar::Raise_DropDown, Fire_ExecuteCommand,
	///       Fire_Click
	/// \endif
	HRESULT Fire_DropDown(IToolBarButton* pToolButton, RECTANGLE* pButtonRectangle, DropDownReturnValuesConstants* pFurtherProcessing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pToolButton;
				p[0].plVal = reinterpret_cast<PLONG>(pFurtherProcessing);		p[0].vt = VT_I4 | VT_BYREF;

				// pack the pButtonRectangle parameter into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{AC1E3771-638F-4e25-B405-682C2A2389D1}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#else
					LPOLESTR clsid = OLESTR("{4111318D-17B8-4728-9AC5-C8B3C6A6498F}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo)));
				#endif
				VariantInit(&p[1]);
				p[1].vt = VT_RECORD | VT_BYREF;
				p[1].pRecInfo = pRecordInfo;
				p[1].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Bottom = pButtonRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Left = pButtonRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Right = pButtonRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Top = pButtonRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_TBE_DROPDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[1].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c EndCustomization event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::EndCustomization, ToolBar::Raise_EndCustomization,
	///       Fire_BeginCustomization, Fire_CustomizedControl
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::EndCustomization, ToolBar::Raise_EndCustomization,
	///       Fire_BeginCustomization, Fire_CustomizedControl
	/// \endif
	HRESULT Fire_EndCustomization(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_TBE_ENDCUSTOMIZATION, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ExecuteCommand event</em>
	///
	/// \param[in] commandID The unique ID of the command that shall be executed.
	/// \param[in] pToolButton The tool bar button whose associated command shall be executed. May be
	///            \c NULL.
	/// \param[in] commandOrigin Specifies whether the command was triggered by a button, a menu item or a
	///            hotkey. Any of the values defined by the \c CommandOriginConstants enumeration is valid.
	/// \param[in,out] pForwardMessage If set to \c VARIANT_TRUE, the command will be forwarded to the parent
	///                top-level window; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::ExecuteCommand, ToolBar::Raise_ExecuteCommand,
	///       Fire_ButtonSelectionStateChanged, Fire_Click, Fire_DropDown
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ExecuteCommand, ToolBar::Raise_ExecuteCommand,
	///       Fire_ButtonSelectionStateChanged, Fire_Click, Fire_DropDown
	/// \endif
	HRESULT Fire_ExecuteCommand(LONG commandID, IToolBarButton* pToolButton, CommandOriginConstants commandOrigin, VARIANT_BOOL* pForwardMessage)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = commandID;
				p[2] = pToolButton;
				p[1] = commandOrigin;
				p[0].pboolVal = pForwardMessage;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TBE_EXECUTECOMMAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FreeButtonData event</em>
	///
	/// \param[in] pToolButton The tool bar button whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::FreeButtonData, ToolBar::Raise_FreeButtonData,
	///       Fire_RemovingButton, Fire_RemovedButton
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::FreeButtonData, ToolBar::Raise_FreeButtonData,
	///       Fire_RemovingButton, Fire_RemovedButton
	/// \endif
	HRESULT Fire_FreeButtonData(IToolBarButton* pToolButton)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pToolButton;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_FREEBUTTONDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c GetAvailableButton event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::GetAvailableButton, ToolBar::Raise_GetAvailableButton,
	///       Fire_QueryInsertButton, Fire_BeginCustomization, Fire_EndCustomization
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::GetAvailableButton, ToolBar::Raise_GetAvailableButton,
	///       Fire_QueryInsertButton, Fire_BeginCustomization, Fire_EndCustomization
	/// \endif
	HRESULT Fire_GetAvailableButton(IVirtualToolBarButton* pToolButton, VARIANT_BOOL* pNoMoreButtons)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pToolButton;
				p[0].pboolVal = pNoMoreButtons;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_GETAVAILABLEBUTTON, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HotButtonChangeWrapping event</em>
	///
	/// \param[in] pPreviousHotButton The previous hot button. May be \ NULL.
	/// \param[in] direction The direction into which the hot button will be changed. Any of the values
	///            defined by the \c WrappingDirectionConstants enumeration is valid.
	/// \param[in] causedBy The reason for the hot button change. Any combination of the values defined by
	///            the \c HotButtonChangingCausedByConstants enumeration is valid.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should abort the hot button change;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::HotButtonChangeWrapping, ToolBar::Raise_HotButtonChangeWrapping,
	///       Fire_HotButtonChanging
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::HotButtonChangeWrapping, ToolBar::Raise_HotButtonChangeWrapping,
	///       Fire_HotButtonChanging
	/// \endif
	HRESULT Fire_HotButtonChangeWrapping(IToolBarButton* pPreviousHotButton, WrappingDirectionConstants direction, HotButtonChangingCausedByConstants causedBy, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pPreviousHotButton;
				p[2] = direction;
				p[1] = causedBy;
				p[0].pboolVal = pCancelChange;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TBE_HOTBUTTONCHANGEWRAPPING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HotButtonChanging event</em>
	///
	/// \param[in] pPreviousHotButton The previous hot button. May be \ NULL.
	/// \param[in] pNewHotButton The new hot button. May be \ NULL.
	/// \param[in] causedBy The reason for the hot button change. Any combination of the values defined by
	///            the \c HotButtonChangingCausedByConstants enumeration is valid.
	/// \param[in] additionalInfo Additional information over the hot button change. Any combination of the
	///            values defined by the \c HotButtonChangingAdditionalInfoConstants enumeration is valid.
	/// \param[in,out] pCancelChange If \c VARIANT_TRUE, the caller should abort the hot button change;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::HotButtonChanging, ToolBar::Raise_HotButtonChanging,
	///       Fire_HotButtonChangeWrapping
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::HotButtonChanging, ToolBar::Raise_HotButtonChanging,
	///       Fire_HotButtonChangeWrapping
	/// \endif
	HRESULT Fire_HotButtonChanging(IToolBarButton* pPreviousHotButton, IToolBarButton* pNewHotButton, HotButtonChangingCausedByConstants causedBy, HotButtonChangingAdditionalInfoConstants additionalInfo, VARIANT_BOOL* pCancelChange)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = pPreviousHotButton;
				p[3] = pNewHotButton;
				p[2] = causedBy;
				p[1] = additionalInfo;
				p[0].pboolVal = pCancelChange;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_TBE_HOTBUTTONCHANGING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InitializeCustomizationDialog event</em>
	///
	/// \param[in] hCustomizationDialog The window handle of the customization dialog.
	/// \param[in,out] pDisplayHelpButton If \c VARIANT_TRUE, the customization dialog shall be displayed
	///                with a help button; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::InitializeCustomizationDialog,
	///       ToolBar::Raise_InitializeCustomizationDialog, Fire_BeginCustomization, Fire_EndCustomization,
	///       Fire_DisplayCustomizationHelp
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::InitializeCustomizationDialog,
	///       ToolBar::Raise_InitializeCustomizationDialog, Fire_BeginCustomization, Fire_EndCustomization,
	///       Fire_DisplayCustomizationHelp
	/// \endif
	HRESULT Fire_InitializeCustomizationDialog(LONG hCustomizationDialog, VARIANT_BOOL* pDisplayHelpButton)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = hCustomizationDialog;
				p[0].pboolVal = pDisplayHelpButton;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_INITIALIZECUSTOMIZATIONDIALOG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InitializeToolBarStateRegistryRestorage event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::InitializeToolBarStateRegistryRestorage,
	///       ToolBar::Raise_InitializeToolBarStateRegistryRestorage,
	///       Fire_InitializeToolBarStateRegistryStorage, Fire_RestoreButtonFromRegistryStream
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::InitializeToolBarStateRegistryRestorage,
	///       ToolBar::Raise_InitializeToolBarStateRegistryRestorage,
	///       Fire_InitializeToolBarStateRegistryStorage, Fire_RestoreButtonFromRegistryStream
	/// \endif
	HRESULT Fire_InitializeToolBarStateRegistryRestorage(LONG* pNumberOfButtonsToLoad, LONG totalDataSize, LONG perButtonDataSize, LONG pDataStream, LONG* ppStartOfNextDataBlock, VARIANT_BOOL* pCancelLoading)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5].plVal = pNumberOfButtonsToLoad;		p[5].vt = VT_I4 | VT_BYREF;
				p[4] = totalDataSize;
				p[3] = perButtonDataSize;
				p[2] = pDataStream;
				p[1].plVal = ppStartOfNextDataBlock;		p[1].vt = VT_I4 | VT_BYREF;
				p[0].pboolVal = pCancelLoading;					p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_INITIALIZETOOLBARSTATEREGISTRYRESTORAGE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InitializeToolBarStateRegistryStorage event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::InitializeToolBarStateRegistryStorage,
	///       ToolBar::Raise_InitializeToolBarStateRegistryStorage,
	///       Fire_InitializeToolBarStateRegistryRestorage, Fire_SaveButtonToRegistryStream
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::InitializeToolBarStateRegistryStorage,
	///       ToolBar::Raise_InitializeToolBarStateRegistryStorage,
	///       Fire_InitializeToolBarStateRegistryRestorage, Fire_SaveButtonToRegistryStream
	/// \endif
	HRESULT Fire_InitializeToolBarStateRegistryStorage(LONG* pNumberOfButtonsToSave, LONG* pTotalDataSize, LONG* ppDataStream, LONG* ppStartOfUnusedSpace)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3].plVal = pNumberOfButtonsToSave;		p[3].vt = VT_I4 | VT_BYREF;
				p[2].plVal = pTotalDataSize;						p[2].vt = VT_I4 | VT_BYREF;
				p[1].plVal = ppDataStream;							p[1].vt = VT_I4 | VT_BYREF;
				p[0].plVal = ppStartOfUnusedSpace;			p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TBE_INITIALIZETOOLBARSTATEREGISTRYSTORAGE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedButton event</em>
	///
	/// \param[in] pToolButton The inserted tool bar button.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::InsertedButton, ToolBar::Raise_InsertedButton,
	///       Fire_InsertingButton, Fire_RemovedButton
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::InsertedButton, ToolBar::Raise_InsertedButton,
	///       Fire_InsertingButton, Fire_RemovedButton
	/// \endif
	HRESULT Fire_InsertedButton(IToolBarButton* pToolButton)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pToolButton;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_INSERTEDBUTTON, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingButton event</em>
	///
	/// \param[in] pToolButton The tool bar button being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::InsertingButton, ToolBar::Raise_InsertingButton,
	///       Fire_InsertedButton, Fire_RemovingButton
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::InsertingButton, ToolBar::Raise_InsertingButton,
	///       Fire_InsertedButton, Fire_RemovingButton
	/// \endif
	HRESULT Fire_InsertingButton(IVirtualToolBarButton* pToolButton, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pToolButton;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_INSERTINGBUTTON, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c IsDuplicateAccelerator event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::IsDuplicateAccelerator, ToolBar::Raise_IsDuplicateAccelerator,
	///       Fire_MapAccelerator
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::IsDuplicateAccelerator, ToolBar::Raise_IsDuplicateAccelerator,
	///       Fire_MapAccelerator
	/// \endif
	HRESULT Fire_IsDuplicateAccelerator(SHORT accelerator, VARIANT_BOOL* pIsDuplicate, VARIANT_BOOL* pHandledEvent)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = accelerator;								p[2].vt = VT_I2;
				p[1].pboolVal = pIsDuplicate;			p[1].vt = VT_BOOL | VT_BYREF;
				p[0].pboolVal = pHandledEvent;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_TBE_ISDUPLICATEACCELERATOR, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyDown event</em>
	///
	/// \param[in,out] pKeyCode The pressed key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYDOWN message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::KeyDown, ToolBar::Raise_KeyDown, Fire_KeyUp, Fire_KeyPress
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::KeyDown, ToolBar::Raise_KeyDown, Fire_KeyUp, Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyDown(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_KEYDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyPress event</em>
	///
	/// \param[in,out] pKeyAscii The pressed key's ASCII code. If set to 0, the caller should eat the
	///                \c WM_CHAR message.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::KeyPress, ToolBar::Raise_KeyPress, Fire_KeyDown, Fire_KeyUp
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::KeyPress, ToolBar::Raise_KeyPress, Fire_KeyDown, Fire_KeyUp
	/// \endif
	HRESULT Fire_KeyPress(SHORT* pKeyAscii)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].piVal = pKeyAscii;		p[0].vt = VT_I2 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_KEYPRESS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c KeyUp event</em>
	///
	/// \param[in,out] pKeyCode The released key. Any of the values defined by VB's \c KeyCodeConstants
	///                enumeration is valid. If set to 0, the caller should eat the \c WM_KEYUP message.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::KeyUp, ToolBar::Raise_KeyUp, Fire_KeyDown, Fire_KeyPress
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::KeyUp, ToolBar::Raise_KeyUp, Fire_KeyDown, Fire_KeyPress
	/// \endif
	HRESULT Fire_KeyUp(SHORT* pKeyCode, SHORT shift)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1].piVal = pKeyCode;		p[1].vt = VT_I2 | VT_BYREF;
				p[0] = shift;							p[0].vt = VT_I2;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_KEYUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MapAccelerator event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::MapAccelerator, ToolBar::Raise_MapAccelerator, Fire_KeyDown,
	///       Fire_KeyUp, Fire_KeyPress, Fire_ExecuteCommand, Fire_IsDuplicateAccelerator
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::MapAccelerator, ToolBar::Raise_MapAccelerator, Fire_KeyDown,
	///       Fire_KeyUp, Fire_KeyPress, Fire_ExecuteCommand, Fire_IsDuplicateAccelerator
	/// \endif
	HRESULT Fire_MapAccelerator(SHORT accelerator, IToolBarButton* pStartingPointOfSearch, VARIANT_BOOL resumingSearchWithFirstButton, IToolBarButton** ppMatchingButton, VARIANT_BOOL* pHandledEvent)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = accelerator;																									p[4].vt = VT_I2;
				p[3] = pStartingPointOfSearch;
				p[2] = resumingSearchWithFirstButton;
				p[1].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppMatchingButton);		p[1].vt = VT_DISPATCH | VT_BYREF;
				p[0].pboolVal = pHandledEvent;																			p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_TBE_MAPACCELERATOR, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MClick event</em>
	///
	/// \param[in] pToolButton The clicked tool bar button. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::MClick, ToolBar::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::MClick, ToolBar::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MClick(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_MCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MDblClick event</em>
	///
	/// \param[in] pToolButton The double-clicked tool bar button. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbMiddleButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::MDblClick, ToolBar::Raise_MDblClick, Fire_MClick, Fire_DblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::MDblClick, ToolBar::Raise_MDblClick, Fire_MClick, Fire_DblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_MDblClick(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_MDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseDown event</em>
	///
	/// \param[in] pToolButton The tool bar button that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The pressed mouse button. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::MouseDown, ToolBar::Raise_MouseDown, Fire_MouseUp, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::MouseDown, ToolBar::Raise_MouseDown, Fire_MouseUp, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseDown(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_MOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseEnter event</em>
	///
	/// \param[in] pToolButton The tool bar button that the mouse cursor is located over. May be \c NULL.
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
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::MouseEnter, ToolBar::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_ButtonMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::MouseEnter, ToolBar::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_ButtonMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseEnter(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_MOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseHover event</em>
	///
	/// \param[in] pToolButton The tool bar button that the mouse cursor is located over. May be \c NULL.
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
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::MouseHover, ToolBar::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::MouseHover, ToolBar::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseHover(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_MOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseLeave event</em>
	///
	/// \param[in] pToolButton The tool bar button that the mouse cursor is located over. May be \c NULL.
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
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::MouseLeave, ToolBar::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_ButtonMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::MouseLeave, ToolBar::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_ButtonMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseLeave(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_MOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseMove event</em>
	///
	/// \param[in] pToolButton The tool bar button that the mouse cursor is located over. May be \c NULL.
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
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::MouseMove, ToolBar::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::MouseMove, ToolBar::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp
	/// \endif
	HRESULT Fire_MouseMove(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_MOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseUp event</em>
	///
	/// \param[in] pToolButton The tool bar button that the mouse cursor is located over. May be \c NULL.
	/// \param[in] button The released mouse buttons. Any of the values defined by VB's
	///            \c MouseButtonConstants enumeration or the \c ExtendedMouseButtonConstants enumeration
	///            is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the mouse cursor's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The exact part of the control that the mouse cursor's position lies
	///            in. Any of the values defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::MouseUp, ToolBar::Raise_MouseUp, Fire_MouseDown, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::MouseUp, ToolBar::Raise_MouseUp, Fire_MouseDown, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseUp(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_MOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLECompleteDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	/// \param[in] performedEffect The performed drop effect. Any of the values (except \c odeScroll)
	///            defined by the \c OLEDropEffectConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLECompleteDrag, ToolBar::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLECompleteDrag, ToolBar::Raise_OLECompleteDrag,
	///       Fire_OLEStartDrag
	/// \endif
	HRESULT Fire_OLECompleteDrag(IOLEDataObject* pData, OLEDropEffectConstants performedEffect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pData;
				p[0].lVal = static_cast<LONG>(performedEffect);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLECOMPLETEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragDrop event</em>
	///
	/// \param[in] pData The dropped data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client finally executed.
	/// \param[in] pDropTarget The button object that is the nearest one from the mouse cursor's position.
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEDragDrop, ToolBar::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEDragDrop, ToolBar::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \endif
	HRESULT Fire_OLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IToolBarButton* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6].plVal = reinterpret_cast<PLONG>(pEffect);		p[6].vt = VT_I4 | VT_BYREF;
				p[5] = pDropTarget;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnter event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The button that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another button.
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
	/// \param[in,out] pAutoClickButton If set to \c VARIANT_TRUE, the caller should auto-click the button
	///                specified by \c ppDropTarget; otherwise it should cancel auto-clicking. See the
	///                following <strong>remarks</strong> section for details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Auto-clicking is timered, i. e. the caller should start the timer after this event, if the
	///          \c ppDropTarget parameter specifies another button than the last time this event was fired.
	///          If the timer isn't canceled, the caller should click the button when the timer expires.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEDragEnter, ToolBar::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEDragEnter, ToolBar::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \endif
	HRESULT Fire_OLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IToolBarButton** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoClickButton)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7].plVal = reinterpret_cast<PLONG>(pEffect);									p[7].vt = VT_I4 | VT_BYREF;
				p[6].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[6].vt = VT_DISPATCH | VT_BYREF;
				p[5] = button;																									p[5].vt = VT_I2;
				p[4] = shift;																										p[4].vt = VT_I2;
				p[3] = x;																												p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																												p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);									p[1].vt = VT_I4;
				p[0].pboolVal = pAutoClickButton;																p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragEnterPotentialTarget event</em>
	///
	/// \param[in] hWndPotentialTarget The handle of the potential drag'n'drop target window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEDragEnterPotentialTarget,
	///       ToolBar::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEDragEnterPotentialTarget,
	///       ToolBar::Raise_OLEDragEnterPotentialTarget, Fire_OLEDragLeavePotentialTarget
	/// \endif
	HRESULT Fire_OLEDragEnterPotentialTarget(LONG hWndPotentialTarget)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndPotentialTarget;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLEDRAGENTERPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] pDropTarget The button that is the current target of the drag'n'drop operation.
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
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEDragLeave, ToolBar::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEDragLeave, ToolBar::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \endif
	HRESULT Fire_OLEDragLeave(IOLEDataObject* pData, IToolBarButton* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pData;
				p[5] = pDropTarget;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeavePotentialTarget event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEDragLeavePotentialTarget,
	///       ToolBar::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEDragLeavePotentialTarget,
	///       ToolBar::Raise_OLEDragLeavePotentialTarget, Fire_OLEDragEnterPotentialTarget
	/// \endif
	HRESULT Fire_OLEDragLeavePotentialTarget(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLEDRAGLEAVEPOTENTIALTARGET, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragMouseMove event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in,out] pEffect On entry, a bit field of the drop effects (defined by the
	///                \c OLEDropEffectConstants enumeration) supported by the drag source. On return,
	///                the drop effect that the client wants to be used on drop.
	/// \param[in,out] ppDropTarget The button that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another button.
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
	/// \param[in,out] pAutoClickButton If set to \c VARIANT_TRUE, the caller should auto-click the button
	///                specified by \c ppDropTarget; otherwise it should cancel auto-clicking. See the
	///                following <strong>remarks</strong> section for details.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Auto-clicking is timered, i. e. the caller should start the timer after this event, if the
	///          \c ppDropTarget parameter specifies another button than the last time this event was fired.
	///          If the timer isn't canceled, the caller should click the button when the timer expires.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEDragMouseMove, ToolBar::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEDragMouseMove, ToolBar::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove
	/// \endif
	HRESULT Fire_OLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IToolBarButton** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pAutoClickButton)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[9];
				p[8] = pData;
				p[7].plVal = reinterpret_cast<PLONG>(pEffect);									p[7].vt = VT_I4 | VT_BYREF;
				p[6].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[6].vt = VT_DISPATCH | VT_BYREF;
				p[5] = button;																									p[5].vt = VT_I2;
				p[4] = shift;																										p[4].vt = VT_I2;
				p[3] = x;																												p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																												p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);									p[1].vt = VT_I4;
				p[0].pboolVal = pAutoClickButton;																p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 9, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEGiveFeedback event</em>
	///
	/// \param[in] effect The current drop effect. It is chosen by the potential drop target. Any of
	///            the values defined by the \c OLEDropEffectConstants enumeration is valid.
	/// \param[in,out] pUseDefaultCursors If set to \c VARIANT_TRUE, the system's default mouse cursors
	///                shall be used to visualize the various drop effects. If set to \c VARIANT_FALSE,
	///                the client has set a custom mouse cursor.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEGiveFeedback, ToolBar::Raise_OLEGiveFeedback,
	///       Fire_OLEQueryContinueDrag
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEGiveFeedback, ToolBar::Raise_OLEGiveFeedback,
	///       Fire_OLEQueryContinueDrag
	/// \endif
	HRESULT Fire_OLEGiveFeedback(OLEDropEffectConstants effect, VARIANT_BOOL* pUseDefaultCursors)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = static_cast<LONG>(effect);			p[1].vt = VT_I4;
				p[0].pboolVal = pUseDefaultCursors;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLEGIVEFEEDBACK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEQueryContinueDrag event</em>
	///
	/// \param[in] pressedEscape If \c VARIANT_TRUE, the user has pressed the \c ESC key since the last
	///            time this event was fired; otherwise not.
	/// \param[in] button The pressed mouse buttons. Any combination of the values defined by VB's
	///            \c MouseButtonConstants enumeration is valid.
	/// \param[in] shift The pressed modifier keys (Shift, Ctrl, Alt). Any combination of the values
	///            defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in,out] pActionToContinueWith Indicates whether to continue, cancel or complete the
	///                drag'n'drop operation. Any of the values defined by the
	///                \c OLEActionToContinueWithConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEQueryContinueDrag, ToolBar::Raise_OLEQueryContinueDrag,
	///       Fire_OLEGiveFeedback
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEQueryContinueDrag, ToolBar::Raise_OLEQueryContinueDrag,
	///       Fire_OLEGiveFeedback
	/// \endif
	HRESULT Fire_OLEQueryContinueDrag(VARIANT_BOOL pressedEscape, SHORT button, SHORT shift, OLEActionToContinueWithConstants* pActionToContinueWith)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pressedEscape;
				p[2] = button;																									p[2].vt = VT_I2;
				p[1] = shift;																										p[1].vt = VT_I2;
				p[0].plVal = reinterpret_cast<PLONG>(pActionToContinueWith);		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLEQUERYCONTINUEDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEReceivedNewData event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEReceivedNewData, ToolBar::Raise_OLEReceivedNewData,
	///       Fire_OLESetData
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEReceivedNewData, ToolBar::Raise_OLEReceivedNewData,
	///       Fire_OLESetData
	/// \endif
	HRESULT Fire_OLEReceivedNewData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLERECEIVEDNEWDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLESetData event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLESetData, ToolBar::Raise_OLESetData, Fire_OLEStartDrag
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLESetData, ToolBar::Raise_OLESetData, Fire_OLEStartDrag
	/// \endif
	HRESULT Fire_OLESetData(IOLEDataObject* pData, LONG formatID, LONG index, LONG dataOrViewAspect)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pData;
				p[2] = formatID;
				p[1] = index;
				p[0] = dataOrViewAspect;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLESETDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEStartDrag event</em>
	///
	/// \param[in] pData The object that holds the dragged data.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::OLEStartDrag, ToolBar::Raise_OLEStartDrag, Fire_OLESetData,
	///       Fire_OLECompleteDrag
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OLEStartDrag, ToolBar::Raise_OLEStartDrag, Fire_OLESetData,
	///       Fire_OLECompleteDrag
	/// \endif
	HRESULT Fire_OLEStartDrag(IOLEDataObject* pData)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pData;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_OLESTARTDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OpenPopupMenu event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::OpenPopupMenu, ToolBar::Raise_OpenPopupMenu,
	///       Fire_SelectedMenuItem, Fire_ButtonGetDropDownMenu
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::OpenPopupMenu, ToolBar::Raise_OpenPopupMenu,
	///       Fire_SelectedMenuItem, Fire_ButtonGetDropDownMenu
	/// \endif
	HRESULT Fire_OpenPopupMenu(LONG hMenu, LONG parentMenuItemIndex, VARIANT_BOOL isSystemMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = hMenu;
				p[1] = parentMenuItemIndex;
				p[0] = isSystemMenu;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_TBE_OPENPOPUPMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c QueryInsertButton event</em>
	///
	/// \param[in] pToolButton The tool bar button for which the control needs to know whether it shall
	///            allow inserting a new button to the left of this button.
	/// \param[in,out] pAllowInsertionToLeft If set to \c VARIANT_TRUE, the customization dialog allows
	///                inserting a button to the left of the specified button; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::QueryInsertButton, ToolBar::Raise_QueryInsertButton,
	///       Fire_QueryRemoveButton, Fire_BeginCustomization, Fire_EndCustomization, Fire_GetAvailableButton
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::QueryInsertButton, ToolBar::Raise_QueryInsertButton,
	///       Fire_QueryRemoveButton, Fire_BeginCustomization, Fire_EndCustomization, Fire_GetAvailableButton
	/// \endif
	HRESULT Fire_QueryInsertButton(IToolBarButton* pToolButton, VARIANT_BOOL* pAllowInsertionToLeft)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pToolButton;
				p[0].pboolVal = pAllowInsertionToLeft;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_QUERYINSERTBUTTON, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c QueryRemoveButton event</em>
	///
	/// \param[in] pToolButton The tool bar button for which the control needs to know whether it may be
	///            removed by the user.
	/// \param[in,out] pAllowRemoval If set to \c VARIANT_TRUE, the customization dialog allows removing the
	///                specified button; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::QueryRemoveButton, ToolBar::Raise_QueryRemoveButton,
	///       Fire_QueryInsertButton, Fire_CustomizeDialogRemovingButton, Fire_BeginCustomization,
	///       Fire_EndCustomization
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::QueryRemoveButton, ToolBar::Raise_QueryRemoveButton,
	///       Fire_QueryInsertButton, Fire_CustomizeDialogRemovingButton, Fire_BeginCustomization,
	///       Fire_EndCustomization
	/// \endif
	HRESULT Fire_QueryRemoveButton(IToolBarButton* pToolButton, VARIANT_BOOL* pAllowRemoval)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pToolButton;
				p[0].pboolVal = pAllowRemoval;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_QUERYREMOVEBUTTON, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RawMenuMessage event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::RawMenuMessage, ToolBar::Raise_RawMenuMessage,
	///       Fire_SelectedMenuItem, Fire_OpenPopupMenu, Fire_ButtonGetDropDownMenu
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::RawMenuMessage, ToolBar::Raise_RawMenuMessage,
	///       Fire_SelectedMenuItem, Fire_OpenPopupMenu, Fire_ButtonGetDropDownMenu
	/// \endif
	HRESULT Fire_RawMenuMessage(LONG message, LONG wParam, LONG lParam, LONG* pResult, VARIANT_BOOL* pHandledEvent)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = message;
				p[3] = wParam;
				p[2] = lParam;
				p[1].plVal = pResult;							p[1].vt = VT_I4 | VT_BYREF;
				p[0].pboolVal = pHandledEvent;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_TBE_RAWMENUMESSAGE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RClick event</em>
	///
	/// \param[in] pToolButton The clicked tool bar button. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was clicked. Any of the values defined
	///            by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::RClick, ToolBar::Raise_RClick, Fire_ContextMenu, Fire_RDblClick,
	///       Fire_Click, Fire_MClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::RClick, ToolBar::Raise_RClick, Fire_ContextMenu, Fire_RDblClick,
	///       Fire_Click, Fire_MClick, Fire_XClick
	/// \endif
	HRESULT Fire_RClick(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_RCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RDblClick event</em>
	///
	/// \param[in] pToolButton The double-clicked tool bar button. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbRightButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::RDblClick, ToolBar::Raise_RDblClick, Fire_RClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_XDblClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::RDblClick, ToolBar::Raise_RDblClick, Fire_RClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_RDblClick(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_RDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RecreatedControlWindow event</em>
	///
	/// \param[in] hWnd The control's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::RecreatedControlWindow,
	///       ToolBar::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::RecreatedControlWindow,
	///       ToolBar::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \endif
	HRESULT Fire_RecreatedControlWindow(LONG hWnd)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWnd;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_RECREATEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovedButton event</em>
	///
	/// \param[in] pToolButton The removed button.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::RemovedButton, ToolBar::Raise_RemovedButton,
	///       Fire_RemovingButton, Fire_InsertedButton
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::RemovedButton, ToolBar::Raise_RemovedButton,
	///       Fire_RemovingButton, Fire_InsertedButton
	/// \endif
	HRESULT Fire_RemovedButton(IVirtualToolBarButton* pToolButton)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pToolButton;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_TBE_REMOVEDBUTTON, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingButton event</em>
	///
	/// \param[in] pToolButton The button being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::RemovingButton, ToolBar::Raise_RemovingButton,
	///       Fire_RemovedButton, Fire_InsertingButton
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::RemovingButton, ToolBar::Raise_RemovingButton,
	///       Fire_RemovedButton, Fire_InsertingButton
	/// \endif
	HRESULT Fire_RemovingButton(IToolBarButton* pToolButton, VARIANT_BOOL* pCancelDeletion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pToolButton;
				p[0].pboolVal = pCancelDeletion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_REMOVINGBUTTON, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResetCustomizations event</em>
	///
	/// \param[in] hCustomizationDialog The window handle of the customization dialog.
	/// \param[in,out] pEndCustomization If \c VARIANT_TRUE, the caller should end customization; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::ResetCustomizations, ToolBar::Raise_ResetCustomizations,
	///       Fire_BeginCustomization, Fire_EndCustomization
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ResetCustomizations, ToolBar::Raise_ResetCustomizations,
	///       Fire_BeginCustomization, Fire_EndCustomization
	/// \endif
	HRESULT Fire_ResetCustomizations(LONG hCustomizationDialog, VARIANT_BOOL* pEndCustomization)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = hCustomizationDialog;
				p[0].pboolVal = pEndCustomization;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_TBE_RESETCUSTOMIZATIONS, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::ResizedControlWindow,
	///       ToolBar::Raise_ResizedControlWindow
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::ResizedControlWindow,
	///       ToolBar::Raise_ResizedControlWindow
	/// \endif
	HRESULT Fire_ResizedControlWindow(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_TBE_RESIZEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RestoreButtonFromRegistryStream event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::RestoreButtonFromRegistryStream,
	///       ToolBar::Raise_RestoreButtonFromRegistryStream, Fire_SaveButtonToRegistryStream,
	///       Fire_InitializeToolBarStateRegistryRestorage
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::RestoreButtonFromRegistryStream,
	///       ToolBar::Raise_RestoreButtonFromRegistryStream, Fire_SaveButtonToRegistryStream,
	///       Fire_InitializeToolBarStateRegistryRestorage
	/// \endif
	HRESULT Fire_RestoreButtonFromRegistryStream(IVirtualToolBarButton* pToolButton, LONG numberOfButtonsToLoad, LONG totalDataSize, LONG perButtonDataSize, LONG pDataStream, LONG* ppStartOfNextDataBlock)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = numberOfButtonsToLoad;
				p[3] = totalDataSize;
				p[2] = perButtonDataSize;
				p[1] = pDataStream;
				p[0].plVal = ppStartOfNextDataBlock;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_RESTOREBUTTONFROMREGISTRYSTREAM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SaveButtonToRegistryStream event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::SaveButtonToRegistryStream,
	///       ToolBar::Raise_SaveButtonToRegistryStream, Fire_RestoreButtonFromRegistryStream,
	///       Fire_InitializeToolBarStateRegistryStorage
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::SaveButtonToRegistryStream,
	///       ToolBar::Raise_SaveButtonToRegistryStream, Fire_RestoreButtonFromRegistryStream,
	///       Fire_InitializeToolBarStateRegistryStorage
	/// \endif
	HRESULT Fire_SaveButtonToRegistryStream(IVirtualToolBarButton* pToolButton, LONG totalDataSize, LONG pDataStream, LONG* ppStartOfUnusedSpace)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pToolButton;
				p[2] = totalDataSize;
				p[1] = pDataStream;
				p[0].plVal = ppStartOfUnusedSpace;		p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_TBE_SAVEBUTTONTOREGISTRYSTREAM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c SelectedMenuItem event</em>
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
	///   \sa TBarCtlsLibU::_IToolBarEvents::SelectedMenuItem, ToolBar::Raise_SelectedMenuItem,
	///       Fire_OpenPopupMenu, Fire_ButtonGetDropDownMenu
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::SelectedMenuItem, ToolBar::Raise_SelectedMenuItem,
	///       Fire_OpenPopupMenu, Fire_ButtonGetDropDownMenu
	/// \endif
	HRESULT Fire_SelectedMenuItem(LONG commandIDOrSubMenuIndex, MenuItemStateConstants menuItemState, LONG hMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = commandIDOrSubMenuIndex;
				p[1] = menuItemState;
				p[0] = hMenu;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_TBE_SELECTEDMENUITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XClick event</em>
	///
	/// \param[in] pToolButton The clicked tool bar button. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the click. This should always be
	///            constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the click. Any
	///            combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::XClick, ToolBar::Raise_XClick, Fire_XDblClick, Fire_Click,
	///       Fire_MClick, Fire_RClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::XClick, ToolBar::Raise_XClick, Fire_XDblClick, Fire_Click,
	///       Fire_MClick, Fire_RClick
	/// \endif
	HRESULT Fire_XClick(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_XCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XDblClick event</em>
	///
	/// \param[in] pToolButton The double-clicked tool bar button. May be \c NULL.
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be a constant defined by the \c ExtendedMouseButtonConstants enumeration.
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the control's
	///            upper-left corner.
	/// \param[in] hitTestDetails The part of the control that was double-clicked. Any of the values
	///            defined by the \c HitTestConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IToolBarEvents::XDblClick, ToolBar::Raise_XDblClick, Fire_XClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_RDblClick
	/// \else
	///   \sa TBarCtlsLibA::_IToolBarEvents::XDblClick, ToolBar::Raise_XDblClick, Fire_XClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_RDblClick
	/// \endif
	HRESULT Fire_XDblClick(IToolBarButton* pToolButton, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pToolButton;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_TBE_XDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IToolBarEvents