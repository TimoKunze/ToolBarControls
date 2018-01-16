//////////////////////////////////////////////////////////////////////
/// \class Proxy_IControlHostWindowEvents
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Fires events specified by the \c _IControlHostWindowEvents interface</em>
///
/// \if UNICODE
///   \sa ControlHostWindow, TBarCtlsLibU::_IControlHostWindowEvents
/// \else
///   \sa ControlHostWindow, TBarCtlsLibA::_IControlHostWindowEvents
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "DispIDs.h"


template<class TypeOfTrigger>
class Proxy_IControlHostWindowEvents :
    public IConnectionPointImpl<TypeOfTrigger, &__uuidof(_IControlHostWindowEvents), CComDynamicUnkArray>
{
public:
	/// \brief <em>Fires the \c Activate event</em>
	///
	/// \param[in] activatedByMouseClick If \c VARIANT_TRUE, the window is being activated in response to a
	///            mouse click; otherwise it is activated by using the keyboard or by code.
	/// \param[in] hWndBeingDeactivated The handle of the window that is being deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::Activate, ControlHostWindow::Raise_Activate,
	///       Fire_Deactivate
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::Activate, ControlHostWindow::Raise_Activate,
	///       Fire_Deactivate
	/// \endif
	HRESULT Fire_Activate(VARIANT_BOOL activatedByMouseClick, LONG hWndBeingDeactivated)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[2];
				p[1] = activatedByMouseClick;
				p[0] = hWndBeingDeactivated;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_CHWE_ACTIVATE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Closing event</em>
	///
	/// \param[out] pCancel If set to \c VARIANT_TRUE, the window shall be kept open; otherwise it shall be
	///             closed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::Closing, ControlHostWindow::Raise_Closing,
	///       Fire_Destroyed, Fire_Created
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::Closing, ControlHostWindow::Raise_Closing,
	///       Fire_Destroyed, Fire_Created
	/// \endif
	HRESULT Fire_Closing(VARIANT_BOOL* pCancel)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0].pboolVal = pCancel;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_CHWE_CLOSING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Created event</em>
	///
	/// \param[in] hWnd The control host's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::Created, ControlHostWindow::Raise_Created,
	///       Fire_Destroyed
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::Created, ControlHostWindow::Raise_Created,
	///       Fire_Destroyed
	/// \endif
	HRESULT Fire_Created(LONG hWnd)
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
				hr = pConnection->Invoke(DISPID_CHWE_CREATED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Deactivate event</em>
	///
	/// \param[in] hWndBeingActivated The handle of the window that is being deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::Deactivate, ControlHostWindow::Raise_Deactivate,
	///       Fire_Activate
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::Deactivate, ControlHostWindow::Raise_Deactivate,
	///       Fire_Activate
	/// \endif
	HRESULT Fire_Deactivate(LONG hWndBeingActivated)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[1];
				p[0] = hWndBeingActivated;

				// invoke the event
				DISPPARAMS params = {p, NULL, 1, 0};
				hr = pConnection->Invoke(DISPID_CHWE_DEACTIVATE, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Destroyed event</em>
	///
	/// \param[in] hWnd The control host's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::Destroyed, ControlHostWindow::Raise_Destroyed,
	///       Fire_Closing, Fire_Created
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::Destroyed, ControlHostWindow::Raise_Destroyed,
	///       Fire_Closing, Fire_Created
	/// \endif
	HRESULT Fire_Destroyed(LONG hWnd)
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
				hr = pConnection->Invoke(DISPID_CHWE_DESTROYED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Moved event</em>
	///
	/// \param[in] left The x-coordinate (in pixels) of the window's position relative to the screen's
	///            upper-left corner.
	/// \param[in] top The y-coordinate (in pixels) of the window's position relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::Moved, ControlHostWindow::Raise_Moved, Fire_Moving,
	///       Fire_Resized
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::Moved, ControlHostWindow::Raise_Moved, Fire_Moving,
	///       Fire_Resized
	/// \endif
	HRESULT Fire_Moved(OLE_XPOS_PIXELS left, OLE_YPOS_PIXELS top)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				CComVariant p[2];
				p[1] = left;		p[1].vt = VT_XPOS_PIXELS;
				p[0] = top;			p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_CHWE_MOVED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Moving event</em>
	///
	/// \param[in,out] pLeft The x-coordinate (in pixels) of the window's position relative to the screen's
	///                upper-left corner. This value can be changed by the client application.
	/// \param[in,out] pTop The y-coordinate (in pixels) of the window's position relative to the screen's
	///                upper-left corner. This value can be changed by the client application.
	/// \param[out] pPreventMove If set to \c VARIANT_TRUE, the change to the window's position shall be
	///             refused; otherwise it shall be accepted.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::Moving, ControlHostWindow::Raise_Moving, Fire_Moved,
	///       Fire_Resizing
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::Moving, ControlHostWindow::Raise_Moving, Fire_Moved,
	///       Fire_Resizing
	/// \endif
	HRESULT Fire_Moving(OLE_XPOS_PIXELS* pLeft, OLE_YPOS_PIXELS* pTop, VARIANT_BOOL* pPreventMove)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				CComVariant p[3];
				p[2].plVal = pLeft;							p[2].vt = VT_XPOS_PIXELS | VT_BYREF;
				p[1].plVal = pTop;							p[1].vt = VT_YPOS_PIXELS | VT_BYREF;
				p[0].pboolVal = pPreventMove;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_CHWE_MOVING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Resized event</em>
	///
	/// \param[in] width The control host window's width in pixels.
	/// \param[in] height The control host window's height in pixels.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::Resized, ControlHostWindow::Raise_Resized,
	///        Fire_Resizing, Fire_Moved
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::Resized, ControlHostWindow::Raise_Resized,
	///        Fire_Resizing, Fire_Moved
	/// \endif
	HRESULT Fire_Resized(OLE_XSIZE_PIXELS width, OLE_YSIZE_PIXELS height)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				CComVariant p[2];
				p[1] = width;			p[1].vt = VT_XSIZE_PIXELS;
				p[0] = height;		p[0].vt = VT_YSIZE_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 2, 0};
				hr = pConnection->Invoke(DISPID_CHWE_RESIZED, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c Resizing event</em>
	///
	/// \param[in,out] pWidth The control host window's width in pixels. This value can be changed by the
	///                client application.
	/// \param[in,out] pHeight The control host window's height in pixels. This value can be changed by the
	///                client application.
	/// \param[out] pPreventResize If set to \c VARIANT_TRUE, the change to the window's size shall be
	///             refused; otherwise it shall be accepted.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::Resizing, ControlHostWindow::Raise_Resizing,
	///        Fire_Resized, Fire_Moving
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::Resizing, ControlHostWindow::Raise_Resizing,
	///        Fire_Resized, Fire_Moving
	/// \endif
	HRESULT Fire_Resizing(OLE_XSIZE_PIXELS* pWidth, OLE_YSIZE_PIXELS* pHeight, VARIANT_BOOL* pPreventResize)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				// invoke the event
				CComVariant p[3];
				p[2].plVal = pWidth;							p[2].vt = VT_XSIZE_PIXELS | VT_BYREF;
				p[1].plVal = pHeight;							p[1].vt = VT_YSIZE_PIXELS | VT_BYREF;
				p[0].pboolVal = pPreventResize;		p[0].vt = VT_BOOL | VT_BYREF;

				// invoke the event
				DISPPARAMS params = {p, NULL, 3, 0};
				hr = pConnection->Invoke(DISPID_CHWE_RESIZING, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}

	/// \brief <em>Fires the \c TitleDblClick event</em>
	///
	/// \param[in] button The mouse buttons that were pressed during the double-click. This should
	///            always be \c vbLeftButton (defined by VB's \c MouseButtonConstants enumeration).
	/// \param[in] shift The modifier keys (Shift, Ctrl, Alt) that were pressed during the double-click.
	///            Any combination of the values defined by VB's \c ShiftConstants enumeration is valid.
	/// \param[in] x The x-coordinate (in pixels) of the double-click's position relative to the screen's
	///            upper-left corner.
	/// \param[in] y The y-coordinate (in pixels) of the double-click's position relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa TBarCtlsLibU::_IControlHostWindowEvents::TitleDblClick, ControlHostWindow::Raise_TitleDblClick
	/// \else
	///   \sa TBarCtlsLibA::_IControlHostWindowEvents::TitleDblClick, ControlHostWindow::Raise_TitleDblClick
	/// \endif
	HRESULT Fire_TitleDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y)
	{
		HRESULT hr = S_OK;
		// invoke the event for each connection point
		for(int i = 0; i < m_vec.GetSize(); ++i) {
			LPDISPATCH pConnection = static_cast<LPDISPATCH>(m_vec.GetAt(i));
			if(pConnection) {
				CComVariant p[4];
				p[3] = button;		p[3].vt = VT_I2;
				p[2] = shift;			p[2].vt = VT_I2;
				p[1] = x;					p[1].vt = VT_XPOS_PIXELS;
				p[0] = y;					p[0].vt = VT_YPOS_PIXELS;

				// invoke the event
				DISPPARAMS params = {p, NULL, 4, 0};
				hr = pConnection->Invoke(DISPID_CHWE_TITLEDBLCLICK, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, NULL, NULL, NULL);
			}
		}
		return hr;
	}
};     // Proxy_IControlHostWindowEvents