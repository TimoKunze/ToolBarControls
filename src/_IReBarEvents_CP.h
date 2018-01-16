//////////////////////////////////////////////////////////////////////
/// \class Proxy_IReBarEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IReBarEvents interface</em>
///
/// \if UNICODE
///   \sa ReBar, TBarCtlsLibU::_IReBarEvents
/// \else
///   \sa ReBar, TBarCtlsLibA::_IReBarEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IReBarEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IReBarEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c AutoBreakingBand event</em>
	///
	/// \param[in] pBand The band that is about to be moved.
	/// \param[in,out] pDoAutoBreak If \c VARIANT_FALSE, the caller should keep the band in the current row;
	///                otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::AutoBreakingBand, ReBar::Raise_AutoBreakingBand,
	///       Fire_LayoutChanged
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::AutoBreakingBand, ReBar::Raise_AutoBreakingBand,
	///       Fire_LayoutChanged
	/// \endif
	HRESULT Fire_AutoBreakingBand(IReBarBand* pBand, VARIANT_BOOL* pDoAutoBreak)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pBand;
				p[0].pboolVal = pDoAutoBreak;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_RBE_AUTOBREAKINGBAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c AutoSized event</em>
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
	///   \sa TBarCtlsLibU::_IReBarEvents::AutoSized, ReBar::Raise_AutoSized, Fire_HeightChanged,
	///       Fire_LayoutChanged, Fire_ResizedControlWindow
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::AutoSized, ReBar::Raise_AutoSized, Fire_HeightChanged,
	///       Fire_LayoutChanged, Fire_ResizedControlWindow
	/// \endif
	HRESULT Fire_AutoSized(RECTANGLE* pTargetRectangle, RECTANGLE* pActualRectangle, VARIANT_BOOL changedBandHeightOrStyle)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];

				// pack the pTargetRectangle and pActualRectangle parameters into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo1 = NULL;
				CComPtr<IRecordInfo> pRecordInfo2 = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{AC1E3771-638F-4e25-B405-682C2A2389D1}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo1)));
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo2)));
				#else
					LPOLESTR clsid = OLESTR("{4111318D-17B8-4728-9AC5-C8B3C6A6498F}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo1)));
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo2)));
				#endif
				VariantInit(&p[2]);
				p[2].vt = VT_RECORD | VT_BYREF;
				p[2].pRecInfo = pRecordInfo1;
				p[2].pvRecord = pRecordInfo1->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Bottom = pTargetRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Left = pTargetRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Right = pTargetRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Top = pTargetRectangle->Top;
				VariantInit(&p[1]);
				p[1].vt = VT_RECORD | VT_BYREF;
				p[1].pRecInfo = pRecordInfo2;
				p[1].pvRecord = pRecordInfo2->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Bottom = pActualRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Left = pActualRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Right = pActualRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Top = pActualRectangle->Top;
				p[0].boolVal = changedBandHeightOrStyle;		p[0].vt = VT_BOOL;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_RBE_AUTOSIZED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo1) {
					pRecordInfo1->RecordDestroy(p[2].pvRecord);
				}
				if(pRecordInfo2) {
					pRecordInfo2->RecordDestroy(p[1].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BandBeginDrag event</em>
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
	///   \sa TBarCtlsLibU::_IReBarEvents::BandBeginDrag, ReBar::Raise_BandBeginDrag, Fire_BandBeginRDrag,
	///       Fire_BandEndDrag
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::BandBeginDrag, ReBar::Raise_BandBeginDrag, Fire_BandBeginRDrag,
	///       Fire_BandEndDrag
	/// \endif
	HRESULT Fire_BandBeginDrag(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pCancelDrag)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pBand;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pCancelDrag;											p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_RBE_BANDBEGINDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BandBeginRDrag event</em>
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
	///   \sa TBarCtlsLibU::_IReBarEvents::BandBeginRDrag, ReBar::Raise_BandBeginRDrag, Fire_BandBeginDrag,
	///       Fire_BandEndDrag
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::BandBeginRDrag, ReBar::Raise_BandBeginRDrag, Fire_BandBeginDrag,
	///       Fire_BandEndDrag
	/// \endif
	HRESULT Fire_BandBeginRDrag(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, VARIANT_BOOL* pCancelDrag)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pBand;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].pboolVal = pCancelDrag;											p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_RBE_BANDBEGINRDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BandEndDrag event</em>
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
	///   \sa TBarCtlsLibU::_IReBarEvents::BandEndDrag, ReBar::Raise_BandEndDrag, Fire_BandBeginDrag,
	///       Fire_BandBeginRDrag
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::BandEndDrag, ReBar::Raise_BandEndDrag, Fire_BandBeginDrag,
	///       Fire_BandBeginRDrag
	/// \endif
	HRESULT Fire_BandEndDrag(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_BANDENDDRAG, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BandMouseEnter event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::BandMouseEnter, ReBar::Raise_BandMouseEnter, Fire_BandMouseLeave,
	///       Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::BandMouseEnter, ReBar::Raise_BandMouseEnter, Fire_BandMouseLeave,
	///       Fire_MouseMove
	/// \endif
	HRESULT Fire_BandMouseEnter(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_BANDMOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BandMouseLeave event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::BandMouseLeave, ReBar::Raise_BandMouseLeave, Fire_BandMouseEnter,
	///       Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::BandMouseLeave, ReBar::Raise_BandMouseLeave, Fire_BandMouseEnter,
	///       Fire_MouseMove
	/// \endif
	HRESULT Fire_BandMouseLeave(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_BANDMOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c BeforeDisplayMDIChildSystemMenu event</em>
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
	///   \sa TBarCtlsLibU::_IReBarEvents::BeforeDisplayMDIChildSystemMenu,
	///       ReBar::Raise_BeforeDisplayMDIChildSystemMenu, Fire_CleanupMDIChildSystemMenu
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::BeforeDisplayMDIChildSystemMenu,
	///       ReBar::Raise_BeforeDisplayMDIChildSystemMenu, Fire_CleanupMDIChildSystemMenu
	/// \endif
	HRESULT Fire_BeforeDisplayMDIChildSystemMenu(LONG hActiveMDIChild, LONG hMenu, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, VARIANT_BOOL* pCancelMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[5];
				p[4] = hActiveMDIChild;
				p[3] = hMenu;
				p[2] = x;												p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;												p[1].vt = VT_YPOS_PIXELS;
				p[0].pboolVal = pCancelMenu;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 5, 0};
				hr = pConnection->Invoke(DISPID_RBE_BEFOREDISPLAYMDICHILDSYSTEMMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ChevronClick event</em>
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
	///   \sa TBarCtlsLibU::_IReBarEvents::ChevronClick, ReBar::Raise_ChevronClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::ChevronClick, ReBar::Raise_ChevronClick
	/// \endif
	HRESULT Fire_ChevronClick(IReBarBand* pBand, RECTANGLE* pChevronRectangle, LONG userData, VARIANT_BOOL* pDoDefault)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = pBand;

				// pack the pBandRectangle and pContainedWindowRectangle parameters into a VARIANT of type VT_RECORD
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
				VariantInit(&p[2]);
				p[2].vt = VT_RECORD | VT_BYREF;
				p[2].pRecInfo = pRecordInfo;
				p[2].pvRecord = pRecordInfo->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Bottom = pChevronRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Left = pChevronRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Right = pChevronRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[2].pvRecord)->Top = pChevronRectangle->Top;
				p[1] = userData;
				p[0].pboolVal = pDoDefault;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_RBE_CHEVRONCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[2].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CleanupMDIChildSystemMenu event</em>
	///
	/// \param[in] hActiveMDIChild The handle of the currently active MDI child window of which the
	///            system menu should be cleaned up.
	/// \param[in] hMenu The handle of the system menu that should be cleaned up.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::CleanupMDIChildSystemMenu, ReBar::Raise_CleanupMDIChildSystemMenu,
	///       Fire_BeforeDisplayMDIChildSystemMenu
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::CleanupMDIChildSystemMenu, ReBar::Raise_CleanupMDIChildSystemMenu,
	///       Fire_BeforeDisplayMDIChildSystemMenu
	/// \endif
	HRESULT Fire_CleanupMDIChildSystemMenu(LONG hActiveMDIChild, LONG hMenu)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = hActiveMDIChild;
				p[0] = hMenu;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_RBE_CLEANUPMDICHILDSYSTEMMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Click event</em>
	///
	/// \param[in] pBand The clicked band. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::Click, ReBar::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::Click, ReBar::Raise_Click, Fire_DblClick, Fire_MClick,
	///       Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_Click(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_CLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ContextMenu event</em>
	///
	/// \param[in] pBand The band that the context menu refers to. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::ContextMenu, ReBar::Raise_ContextMenu, Fire_RClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::ContextMenu, ReBar::Raise_ContextMenu, Fire_RClick
	/// \endif
	HRESULT Fire_ContextMenu(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_CONTEXTMENU, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c CustomDraw event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::CustomDraw, ReBar::Raise_CustomDraw
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::CustomDraw, ReBar::Raise_CustomDraw
	/// \endif
	HRESULT Fire_CustomDraw(IReBarBand* pBand, CustomDrawStageConstants drawStage, CustomDrawItemStateConstants bandState, LONG hDC, RECTANGLE* pDrawingRectangle, CustomDrawReturnValuesConstants* pFurtherProcessing)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4].lVal = static_cast<LONG>(drawStage);										p[4].vt = VT_I4;
				p[3].lVal = static_cast<LONG>(bandState);										p[3].vt = VT_I4;
				p[2] = hDC;
				p[0].plVal = reinterpret_cast<PLONG>(pFurtherProcessing);		p[0].vt = VT_I4 | VT_BYREF;

				// pack the pDrawingRectangle parameter into a VARIANT of type VT_RECORD
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
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Bottom = pDrawingRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Left = pDrawingRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Right = pDrawingRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Top = pDrawingRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_CUSTOMDRAW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				if(pRecordInfo) {
					pRecordInfo->RecordDestroy(p[1].pvRecord);
				}
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DblClick event</em>
	///
	/// \param[in] pBand The double-clicked band. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::DblClick, ReBar::Raise_DblClick, Fire_Click, Fire_MDblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::DblClick, ReBar::Raise_DblClick, Fire_Click, Fire_MDblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_DblClick(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_DBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa TBarCtlsLibU::_IReBarEvents::DestroyedControlWindow,
	///       ReBar::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::DestroyedControlWindow,
	///       ReBar::Raise_DestroyedControlWindow, Fire_RecreatedControlWindow
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
				hr = pConnection->Invoke(DISPID_RBE_DESTROYEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c DraggingSplitter event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::DraggingSplitter, ReBar::Raise_DraggingSplitter
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::DraggingSplitter, ReBar::Raise_DraggingSplitter
	/// \endif
	HRESULT Fire_DraggingSplitter(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_RBE_DRAGGINGSPLITTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c FreeBandData event</em>
	///
	/// \param[in] pBand The band whose associated data shall be freed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::FreeBandData, ReBar::Raise_FreeBandData, Fire_RemovingBand,
	///       Fire_RemovedBand
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::FreeBandData, ReBar::Raise_FreeBandData, Fire_RemovingBand,
	///       Fire_RemovedBand
	/// \endif
	HRESULT Fire_FreeBandData(IReBarBand* pBand)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pBand;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_RBE_FREEBANDDATA, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c HeightChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::HeightChanged, ReBar::Raise_HeightChanged,
	///       Fire_AutoSized, Fire_LayoutChanged, Fire_ResizedControlWindow
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::HeightChanged, ReBar::Raise_HeightChanged,
	///       Fire_AutoSized, Fire_LayoutChanged, Fire_ResizedControlWindow
	/// \endif
	HRESULT Fire_HeightChanged(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_RBE_HEIGHTCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertedBand event</em>
	///
	/// \param[in] pBand The inserted band.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::InsertedBand, ReBar::Raise_InsertedBand, Fire_InsertingBand,
	///       Fire_RemovedBand
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::InsertedBand, ReBar::Raise_InsertedBand, Fire_InsertingBand,
	///       Fire_RemovedBand
	/// \endif
	HRESULT Fire_InsertedBand(IReBarBand* pBand)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pBand;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_RBE_INSERTEDBAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c InsertingBand event</em>
	///
	/// \param[in] pBand The band being inserted.
	/// \param[in,out] pCancelInsertion If \c VARIANT_TRUE, the caller should abort insertion; otherwise
	///                not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::InsertingBand, ReBar::Raise_InsertingBand, Fire_InsertedBand,
	///       Fire_RemovingBand
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::InsertingBand, ReBar::Raise_InsertingBand, Fire_InsertedBand,
	///       Fire_RemovingBand
	/// \endif
	HRESULT Fire_InsertingBand(IVirtualReBarBand* pBand, VARIANT_BOOL* pCancelInsertion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pBand;
				p[0].pboolVal = pCancelInsertion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_RBE_INSERTINGBAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c LayoutChanged event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::LayoutChanged, ReBar::Raise_LayoutChanged,
	///       Fire_ResizingContainedWindow, Fire_AutoSized, Fire_HeightChanged, Fire_ResizedControlWindow
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::LayoutChanged, ReBar::Raise_LayoutChanged,
	///       Fire_ResizingContainedWindow, Fire_AutoSized, Fire_HeightChanged, Fire_ResizedControlWindow
	/// \endif
	HRESULT Fire_LayoutChanged(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_RBE_LAYOUTCHANGED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MClick event</em>
	///
	/// \param[in] pBand The clicked band. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::MClick, ReBar::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::MClick, ReBar::Raise_MClick, Fire_MDblClick, Fire_Click,
	///       Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MClick(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_MCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MDblClick event</em>
	///
	/// \param[in] pBand The double-clicked band. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::MDblClick, ReBar::Raise_MDblClick, Fire_MClick, Fire_DblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::MDblClick, ReBar::Raise_MDblClick, Fire_MClick, Fire_DblClick,
	///       Fire_RDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_MDblClick(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_MDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseDown event</em>
	///
	/// \param[in] pBand The band that the mouse cursor is located over. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::MouseDown, ReBar::Raise_MouseDown, Fire_MouseUp, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::MouseDown, ReBar::Raise_MouseDown, Fire_MouseUp, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseDown(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_MOUSEDOWN, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseEnter event</em>
	///
	/// \param[in] pBand The band that the mouse cursor is located over. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::MouseEnter, ReBar::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_BandMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::MouseEnter, ReBar::Raise_MouseEnter, Fire_MouseLeave,
	///       Fire_BandMouseEnter, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseEnter(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_MOUSEENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseHover event</em>
	///
	/// \param[in] pBand The band that the mouse cursor is located over. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::MouseHover, ReBar::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::MouseHover, ReBar::Raise_MouseHover, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseHover(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_MOUSEHOVER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseLeave event</em>
	///
	/// \param[in] pBand The band that the mouse cursor is located over. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::MouseLeave, ReBar::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_BandMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::MouseLeave, ReBar::Raise_MouseLeave, Fire_MouseEnter,
	///       Fire_BandMouseLeave, Fire_MouseHover, Fire_MouseMove
	/// \endif
	HRESULT Fire_MouseLeave(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_MOUSELEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseMove event</em>
	///
	/// \param[in] pBand The band that the mouse cursor is located over. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::MouseMove, ReBar::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::MouseMove, ReBar::Raise_MouseMove, Fire_MouseEnter,
	///       Fire_MouseLeave, Fire_MouseDown, Fire_MouseUp
	/// \endif
	HRESULT Fire_MouseMove(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_MOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c MouseUp event</em>
	///
	/// \param[in] pBand The band that the mouse cursor is located over. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::MouseUp, ReBar::Raise_MouseUp, Fire_MouseDown, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::MouseUp, ReBar::Raise_MouseUp, Fire_MouseDown, Fire_Click,
	///       Fire_MClick, Fire_RClick, Fire_XClick
	/// \endif
	HRESULT Fire_MouseUp(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_MOUSEUP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c NonClientHitTest event</em>
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
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::NonClientHitTest, ReBar::Raise_NonClientHitTest, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::NonClientHitTest, ReBar::Raise_NonClientHitTest, Fire_MouseMove
	/// \endif
	HRESULT Fire_NonClientHitTest(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails, LONG* pReturnValue)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[7];
				p[6] = pBand;
				p[5] = button;																		p[5].vt = VT_I2;
				p[4] = shift;																			p[4].vt = VT_I2;
				p[3] = x;																					p[3].vt = VT_XPOS_PIXELS;
				p[2] = y;																					p[2].vt = VT_YPOS_PIXELS;
				p[1].lVal = static_cast<LONG>(hitTestDetails);		p[1].vt = VT_I4;
				p[0].plVal = pReturnValue;												p[0].vt = VT_I4 | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 7, 0};
				hr = pConnection->Invoke(DISPID_RBE_NONCLIENTHITTEST, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	/// \param[in] pDropTarget The band object that is the nearest one from the mouse cursor's position.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::OLEDragDrop, ReBar::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::OLEDragDrop, ReBar::Raise_OLEDragDrop, Fire_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_MouseUp
	/// \endif
	HRESULT Fire_OLEDragDrop(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IReBarBand* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_RBE_OLEDRAGDROP, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	/// \param[in,out] ppDropTarget The band that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another band.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::OLEDragEnter, ReBar::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::OLEDragEnter, ReBar::Raise_OLEDragEnter,
	///       Fire_OLEDragMouseMove, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseEnter
	/// \endif
	HRESULT Fire_OLEDragEnter(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IReBarBand** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6].plVal = reinterpret_cast<PLONG>(pEffect);									p[6].vt = VT_I4 | VT_BYREF;
				p[5].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[5].vt = VT_DISPATCH | VT_BYREF;
				p[4] = button;																									p[4].vt = VT_I2;
				p[3] = shift;																										p[3].vt = VT_I2;
				p[2] = x;																												p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																												p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);									p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_RBE_OLEDRAGENTER, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c OLEDragLeave event</em>
	///
	/// \param[in] pData The dragged data.
	/// \param[in] pDropTarget The band that is the current target of the drag'n'drop operation.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::OLEDragLeave, ReBar::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::OLEDragLeave, ReBar::Raise_OLEDragLeave,
	///       Fire_OLEDragEnter, Fire_OLEDragMouseMove, Fire_OLEDragDrop, Fire_MouseLeave
	/// \endif
	HRESULT Fire_OLEDragLeave(IOLEDataObject* pData, IReBarBand* pDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
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
				hr = pConnection->Invoke(DISPID_RBE_OLEDRAGLEAVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	/// \param[in,out] ppDropTarget The band that is the current target of the drag'n'drop operation.
	///                The client may set this parameter to another band.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::OLEDragMouseMove, ReBar::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::OLEDragMouseMove, ReBar::Raise_OLEDragMouseMove,
	///       Fire_OLEDragEnter, Fire_OLEDragLeave, Fire_OLEDragDrop, Fire_MouseMove
	/// \endif
	HRESULT Fire_OLEDragMouseMove(IOLEDataObject* pData, OLEDropEffectConstants* pEffect, IReBarBand** ppDropTarget, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[8];
				p[7] = pData;
				p[6].plVal = reinterpret_cast<PLONG>(pEffect);									p[6].vt = VT_I4 | VT_BYREF;
				p[5].ppdispVal = reinterpret_cast<LPDISPATCH*>(ppDropTarget);		p[5].vt = VT_DISPATCH | VT_BYREF;
				p[4] = button;																									p[4].vt = VT_I2;
				p[3] = shift;																										p[3].vt = VT_I2;
				p[2] = x;																												p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																												p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);									p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 8, 0};
				hr = pConnection->Invoke(DISPID_RBE_OLEDRAGMOUSEMOVE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa TBarCtlsLibU::_IReBarEvents::RawMenuMessage, ReBar::Raise_RawMenuMessage,
	///       Fire_SelectedMenuItem, Fire_BeforeDisplayMDIChildSystemMenu
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::RawMenuMessage, ReBar::Raise_RawMenuMessage,
	///       Fire_SelectedMenuItem, Fire_BeforeDisplayMDIChildSystemMenu
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
				hr = pConnection->Invoke(DISPID_RBE_RAWMENUMESSAGE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RClick event</em>
	///
	/// \param[in] pBand The clicked band. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::RClick, ReBar::Raise_RClick, Fire_ContextMenu, Fire_RDblClick,
	///       Fire_Click, Fire_MClick, Fire_XClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::RClick, ReBar::Raise_RClick, Fire_ContextMenu, Fire_RDblClick,
	///       Fire_Click, Fire_MClick, Fire_XClick
	/// \endif
	HRESULT Fire_RClick(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_RCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RDblClick event</em>
	///
	/// \param[in] pBand The double-clicked band. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::RDblClick, ReBar::Raise_RDblClick, Fire_RClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_XDblClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::RDblClick, ReBar::Raise_RDblClick, Fire_RClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_XDblClick
	/// \endif
	HRESULT Fire_RDblClick(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_RDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
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
	///   \sa TBarCtlsLibU::_IReBarEvents::RecreatedControlWindow,
	///       ReBar::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::RecreatedControlWindow,
	///       ReBar::Raise_RecreatedControlWindow, Fire_DestroyedControlWindow
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
				hr = pConnection->Invoke(DISPID_RBE_RECREATEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ReleasedMouseCapture event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::ReleasedMouseCapture,
	///       ReBar::Raise_ReleasedMouseCapture, Fire_MouseMove
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::ReleasedMouseCapture,
	///       ReBar::Raise_ReleasedMouseCapture, Fire_MouseMove
	/// \endif
	HRESULT Fire_ReleasedMouseCapture(void)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				DISPPARAMS params = {NULL, NULL, 0, 0};
				hr = pConnection->Invoke(DISPID_RBE_RELEASEDMOUSECAPTURE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovedBand event</em>
	///
	/// \param[in] pBand The removed band.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::RemovedBand, ReBar::Raise_RemovedBand, Fire_RemovingBand,
	///       Fire_InsertedBand
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::RemovedBand, ReBar::Raise_RemovedBand, Fire_RemovingBand,
	///       Fire_InsertedBand
	/// \endif
	HRESULT Fire_RemovedBand(IVirtualReBarBand* pBand)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = pBand;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_RBE_REMOVEDBAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c RemovingBand event</em>
	///
	/// \param[in] pBand The band being removed.
	/// \param[in,out] pCancelDeletion If \c VARIANT_TRUE, the caller should abort deletion; otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::RemovingBand, ReBar::Raise_RemovingBand, Fire_RemovedBand,
	///       Fire_InsertingBand
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::RemovingBand, ReBar::Raise_RemovingBand, Fire_RemovedBand,
	///       Fire_InsertingBand
	/// \endif
	HRESULT Fire_RemovingBand(IReBarBand* pBand, VARIANT_BOOL* pCancelDeletion)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pBand;
				p[0].pboolVal = pCancelDeletion;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_RBE_REMOVINGBAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizedControlWindow event</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::ResizedControlWindow,
	///       ReBar::Raise_ResizedControlWindow, Fire_HeightChanged, Fire_LayoutChanged,
	///       Fire_ResizingContainedWindow
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::ResizedControlWindow,
	///       ReBar::Raise_ResizedControlWindow, Fire_HeightChanged, Fire_LayoutChanged,
	///       Fire_ResizingContainedWindow
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
				hr = pConnection->Invoke(DISPID_RBE_RESIZEDCONTROLWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c ResizingContainedWindow event</em>
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
	///   \sa TBarCtlsLibU::_IReBarEvents::ResizingContainedWindow, ReBar::Raise_ResizingContainedWindow,
	///       Fire_ResizedControlWindow, Fire_LayoutChanged
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::ResizingContainedWindow, ReBar::Raise_ResizingContainedWindow,
	///       Fire_ResizedControlWindow, Fire_LayoutChanged
	/// \endif
	HRESULT Fire_ResizingContainedWindow(IReBarBand* pBand, RECTANGLE* pBandClientRectangle, RECTANGLE* pContainedWindowRectangle)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[3];
				p[2] = pBand;

				// pack the pBandRectangle and pContainedWindowRectangle parameters into a VARIANT of type VT_RECORD
				CComPtr<IRecordInfo> pRecordInfo1 = NULL;
				CComPtr<IRecordInfo> pRecordInfo2 = NULL;
				CLSID clsidRECTANGLE = {0};
				#ifdef UNICODE
					LPOLESTR clsid = OLESTR("{AC1E3771-638F-4e25-B405-682C2A2389D1}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo1)));
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibU, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo2)));
				#else
					LPOLESTR clsid = OLESTR("{4111318D-17B8-4728-9AC5-C8B3C6A6498F}");
					CLSIDFromString(clsid, &clsidRECTANGLE);
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo1)));
					ATLVERIFY(SUCCEEDED(GetRecordInfoFromGuids(LIBID_TBarCtlsLibA, VERSION_MAJOR, VERSION_MINOR, GetUserDefaultLCID(), static_cast<REFGUID>(clsidRECTANGLE), &pRecordInfo2)));
				#endif
				VariantInit(&p[1]);
				p[1].vt = VT_RECORD | VT_BYREF;
				p[1].pRecInfo = pRecordInfo1;
				p[1].pvRecord = pRecordInfo1->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Bottom = pBandClientRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Left = pBandClientRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Right = pBandClientRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[1].pvRecord)->Top = pBandClientRectangle->Top;
				VariantInit(&p[0]);
				p[0].vt = VT_RECORD | VT_BYREF;
				p[0].pRecInfo = pRecordInfo2;
				p[0].pvRecord = pRecordInfo2->RecordCreate();
				// transfer data
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Bottom = pContainedWindowRectangle->Bottom;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Left = pContainedWindowRectangle->Left;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Right = pContainedWindowRectangle->Right;
				reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Top = pContainedWindowRectangle->Top;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_RBE_RESIZINGCONTAINEDWINDOW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);

				pContainedWindowRectangle->Bottom = reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Bottom;
				pContainedWindowRectangle->Left = reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Left;
				pContainedWindowRectangle->Right = reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Right;
				pContainedWindowRectangle->Top = reinterpret_cast<RECTANGLE*>(p[0].pvRecord)->Top;

				if(pRecordInfo1) {
					pRecordInfo1->RecordDestroy(p[1].pvRecord);
				}
				if(pRecordInfo2) {
					pRecordInfo2->RecordDestroy(p[0].pvRecord);
				}
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
	///   \sa TBarCtlsLibU::_IReBarEvents::SelectedMenuItem, ReBar::Raise_SelectedMenuItem,
	///       Fire_RawMenuMessage, Fire_BeforeDisplayMDIChildSystemMenu
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::SelectedMenuItem, ReBar::Raise_SelectedMenuItem,
	///       Fire_RawMenuMessage, Fire_BeforeDisplayMDIChildSystemMenu
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
				hr = pConnection->Invoke(DISPID_RBE_SELECTEDMENUITEM, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TogglingBand event</em>
	///
	/// \param[in] pBand The band whose minimized or maximized state is about to be toggled.
	/// \param[in,out] pCancelToggling If \c VARIANT_TRUE, the caller should abort toggling, otherwise not.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IReBarEvents::TogglingBand, ReBar::Raise_TogglingBand
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::TogglingBand, ReBar::Raise_TogglingBand
	/// \endif
	HRESULT Fire_TogglingBand(IReBarBand* pBand, VARIANT_BOOL* pCancelToggling)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = pBand;
				p[0].pboolVal = pCancelToggling;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_RBE_TOGGLINGBAND, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XClick event</em>
	///
	/// \param[in] pBand The clicked band. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::XClick, ReBar::Raise_XClick, Fire_XDblClick, Fire_Click,
	///       Fire_MClick, Fire_RClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::XClick, ReBar::Raise_XClick, Fire_XDblClick, Fire_Click,
	///       Fire_MClick, Fire_RClick
	/// \endif
	HRESULT Fire_XClick(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_XCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c XDblClick event</em>
	///
	/// \param[in] pBand The double-clicked band. May be \c NULL.
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
	///   \sa TBarCtlsLibU::_IReBarEvents::XDblClick, ReBar::Raise_XDblClick, Fire_XClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_RDblClick
	/// \else
	///   \sa TBarCtlsLibA::_IReBarEvents::XDblClick, ReBar::Raise_XDblClick, Fire_XClick, Fire_DblClick,
	///       Fire_MDblClick, Fire_RDblClick
	/// \endif
	HRESULT Fire_XDblClick(IReBarBand* pBand, SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y, HitTestConstants hitTestDetails)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[6];
				p[5] = pBand;
				p[4] = button;																		p[4].vt = VT_I2;
				p[3] = shift;																			p[3].vt = VT_I2;
				p[2] = x;																					p[2].vt = VT_XPOS_PIXELS;
				p[1] = y;																					p[1].vt = VT_YPOS_PIXELS;
				p[0].lVal = static_cast<LONG>(hitTestDetails);		p[0].vt = VT_I4;

				// invoke the event
				DISPPARAMS params = {p, NULL, 6, 0};
				hr = pConnection->Invoke(DISPID_RBE_XDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IReBarEvents