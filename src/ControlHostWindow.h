//////////////////////////////////////////////////////////////////////
/// \class ControlHostWindow
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Creates a dialog for hosting a free-floating control</em>
///
/// This class creates a dialog that behaves like a free-floating tool window of a parent window. It is
/// designed to host exactly one control and can be used to provide e.g. free-floating tool bars.
///
/// \sa ControlHostDialog
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "_IControlHostWindowEvents_CP.h"
#include "helpers.h"
#include "ControlHostDialog.h"


/// \brief <em>The handle of the \c WndProc hook</em>
extern HHOOK hCallWndProcHook;
/// \brief <em>The number of \c ControlHostWindow objects currently being active</em>
extern volatile UINT controlHostWindowsCount;


class ATL_NO_VTABLE ControlHostWindow : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ControlHostWindow, &CLSID_ControlHostWindow>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ControlHostWindow>,
    public Proxy_IControlHostWindowEvents<ControlHostWindow>,
    #ifdef UNICODE
    	public IDispatchImpl<IControlHostWindow, &IID_IControlHostWindow, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IControlHostWindow, &IID_IControlHostWindow, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public CMessageMap
{
	friend class ClassFactory;

public:
	/// \brief <em>The parent window of the control host dialog</em>
	CContainedWindow parentWindow;
	/// \brief <em>The control host dialog window</em>
	CContainedWindow freeFloatDialog;

	/// \brief <em>The constructor of this class</em>
	///
	/// Used for initialization.
	ControlHostWindow();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_CONTROLHOSTWINDOW)

		BEGIN_COM_MAP(ControlHostWindow)
			COM_INTERFACE_ENTRY(IControlHostWindow)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ControlHostWindow)
			CONNECTION_POINT_ENTRY(__uuidof(_IControlHostWindowEvents))
		END_CONNECTION_POINT_MAP()

		DECLARE_PROTECT_FINAL_CONSTRUCT()

		BEGIN_MSG_MAP(ControlHostWindow)
			ALT_MSG_MAP(1)
			MESSAGE_HANDLER(WM_ACTIVATE, OnHostDialogActivate)
			MESSAGE_HANDLER(WM_CLOSE, OnHostDialogClose)
			MESSAGE_HANDLER(WM_DESTROY, OnHostDialogDestroy)
			MESSAGE_HANDLER(WM_NCACTIVATE, OnHostDialogNCActivate)
			MESSAGE_HANDLER(WM_NCLBUTTONDBLCLK, OnHostDialogNCLButtonDblClk)
			MESSAGE_HANDLER(WM_SYSCOMMAND, OnHostDialogSysCommand)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGED, OnHostDialogWindowPosChanged)
			MESSAGE_HANDLER(WM_WINDOWPOSCHANGING, OnHostDialogWindowPosChanging)
			MESSAGE_HANDLER(GetControlHostWindowMessage(), OnHostDialogGetControlHostWindow)
			ALT_MSG_MAP(2)
			MESSAGE_HANDLER(WM_SETFOCUS, OnParentSetFocus)
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
	/// \name Implementation of IControlHostWindow
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c hWnd property</em>
	///
	/// Retrieves the window's handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa Create, get_hWndChild, Raise_Created, Raise_Destroyed
	virtual HRESULT STDMETHODCALLTYPE get_hWnd(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hWndChild property</em>
	///
	/// Retrieves the child control's window handle.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hWndChild, get_hWnd
	virtual HRESULT STDMETHODCALLTYPE get_hWndChild(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hWndChild property</em>
	///
	/// Sets the child control's window handle.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_hWndChild, get_hWnd
	virtual HRESULT STDMETHODCALLTYPE put_hWndChild(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c Resizeable property</em>
	///
	/// Retrieves whether the window can be resized. If set to \c VARIANT_TRUE, the window can be resized;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Resizeable, get_Visible, Raise_Resized
	virtual HRESULT STDMETHODCALLTYPE get_Resizeable(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Resizeable property</em>
	///
	/// Sets whether the window can be resized. If set to \c VARIANT_TRUE, the window can be resized;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Resizeable, put_Visible, Raise_Resized
	virtual HRESULT STDMETHODCALLTYPE put_Resizeable(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Visible property</em>
	///
	/// Retrieves whether the window is visible. If set to \c VARIANT_TRUE, the window is visible; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Visible, get_WindowTitle, get_Resizeable
	virtual HRESULT STDMETHODCALLTYPE get_Visible(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Visible property</em>
	///
	/// Sets whether the window is visible. If set to \c VARIANT_TRUE, the window is visible; otherwise
	/// not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Visible, put_WindowTitle, put_Resizeable
	virtual HRESULT STDMETHODCALLTYPE put_Visible(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c WindowTitle property</em>
	///
	/// Retrieves the text that is displayed as the window's caption.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_WindowTitle, get_Visible
	virtual HRESULT STDMETHODCALLTYPE get_WindowTitle(BSTR* pValue);
	/// \brief <em>Sets the \c WindowTitle property</em>
	///
	/// Sets the text that is displayed as the window's caption.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_WindowTitle, put_Visible
	virtual HRESULT STDMETHODCALLTYPE put_WindowTitle(BSTR newValue);

	/// \brief <em>Calculates the window size that is required to achieve the specified client size</em>
	///
	/// Calculates the window size that is required to achieve the specified client size. The calculation
	/// takes the current window styles into account, e. g. it adds space for the window borders.
	///
	/// \param[in] clientWidth The width (in pixels) of the client area that the window shall have.
	/// \param[in] clientHeight The height (in pixels) of the client area that the window shall have.
	/// \param[out] pWindowWidth Receives the width (in pixels) that the window has to be sized to in order
	///             to achieve the specified client area width.
	/// \param[out] pWindowHeight Receives the height (in pixels) that the window has to be sized to in order
	///             to achieve the specified client area height.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetClientRectangle, GetRectangle, Move
	virtual HRESULT STDMETHODCALLTYPE CalculateWindowSize(long clientWidth = 0x80000000, long clientHeight = 0x80000000, OLE_XSIZE_PIXELS* pWindowWidth = NULL, OLE_YSIZE_PIXELS* pWindowHeight = NULL);
	/// \brief <em>Creates a window that can be used to host other windows</em>
	///
	/// Creates a window that can be used to host other windows.
	///
	/// \param[in] hWndParent The window that will be the parent window of the created window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Destroy, get_hWnd, put_hWndChild, put_Visible, Raise_Created
	virtual HRESULT STDMETHODCALLTYPE Create(OLE_HANDLE hWndParent);
	/// \brief <em>Destroys the control host window</em>
	///
	/// Destroys the control host window.
	///
	/// \remarks This method will fail as long as a child window is placed on the control host window.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Create, get_hWnd, put_hWndChild, Raise_Destroyed
	virtual HRESULT STDMETHODCALLTYPE Destroy(void);
	/// \brief <em>Retrieves the bounding rectangle of the control host window's client area</em>
	///
	/// Retrieves the bounding rectangle (in pixels) of the control host window's client area. The retrieved
	/// rectangle is in client coordinates.
	///
	/// \param[out] pXLeft The x-coordinate (in pixels) of the bounding rectangle's left border.
	/// \param[out] pYTop The y-coordinate (in pixels) of the bounding rectangle's top border.
	/// \param[out] pXRight The x-coordinate (in pixels) of the bounding rectangle's right border.
	/// \param[out] pYBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetRectangle, CalculateWindowSize, Raise_Resizing, Raise_Resized
	virtual HRESULT STDMETHODCALLTYPE GetClientRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Retrieves the bounding rectangle of the control host window</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the screen's upper-left corner) of the
	/// control host window.
	///
	/// \param[out] pXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
	///             relative to the screen's upper-left corner.
	/// \param[out] pYTop The y-coordinate (in pixels) of the bounding rectangle's top border
	///             relative to the screen's upper-left corner.
	/// \param[out] pXRight The x-coordinate (in pixels) of the bounding rectangle's right border
	///             relative to the screen's upper-left corner.
	/// \param[out] pYBottom The y-coordinate (in pixels) of the bounding rectangle's bottom border
	///             relative to the screen's upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetClientRectangle, CalculateWindowSize, Move, Raise_Moving, Raise_Resizing, Raise_Moved,
	///     Raise_Resized
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Moves and resizes the control host window.</em>
	///
	/// Moves and resizes the control host window.
	///
	/// \param[in] left The x-coordinate (in pixels) of the window's position relative to the screen's
	///            upper-left corner.
	/// \param[in] top The y-coordinate (in pixels) of the window's position relative to the screen's
	///            upper-left corner.
	/// \param[in] width The control host window's width in pixels.
	/// \param[in] height The control host window's height in pixels.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa CalculateWindowSize, GetRectangle, Raise_Moving, Raise_Resizing, Raise_Moved, Raise_Resized
	virtual HRESULT STDMETHODCALLTYPE Move(long left = 0x80000000, long top = 0x80000000, long width = 0x80000000, long height = 0x80000000);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>A message is to be processed</em>
	///
	/// This function is the callback function that we defined when we installed a \c WndProc hook to trap
	/// any calls to a window procedure for the \c ControlHostWindow COM class.
	///
	/// \param[in] code A code the hook procedure uses to determine how to process the message.
	/// \param[in] wParam Specifies whether the message was sent by the current thread. If the message was
	///            sent by the current thread, it is nonzero; otherwise, it is zero.
	/// \param[in] lParam A pointer to a \c CWPSTRUCT structure that contains details about the message.
	///
	/// \return The value returned by \c CallNextHookEx or 0 if \c CallNextHookEx was not called.
	///
	/// \sa Create, OnParentSetFocus,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644975.aspx">CallWndProc</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644964.aspx">CWPSTRUCT</a>
	///     <a href="https://msdn.microsoft.com/en-us/library/ms644974.aspx">CallNextHookEx</a>
	static LRESULT CALLBACK CallWndProc(int code, WPARAM wParam, LPARAM lParam);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Message handlers
	///
	//@{
	/// \brief <em>\c WM_ACTIVATE handler</em>
	///
	/// Will be called if the control host dialog is being activated or deactivated.
	/// We use this handler to tidy up and to raise the \c Activate and \c Deactivate events.
	///
	/// \sa Raise_Activate, Raise_Deactivate,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646274.aspx">WM_ACTIVATE</a>
	LRESULT OnHostDialogActivate(UINT /*message*/, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_CLOSE handler</em>
	///
	/// Will be called if the control host dialog shall be closed.
	/// We use this handler to destroy the dialog window.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms632617.aspx">WM_CLOSE</a>
	LRESULT OnHostDialogClose(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_DESTROY handler</em>
	///
	/// Will be called while the control host dialog is being destroyed.
	/// We use this handler to tidy up and to raise the \c Destroyed event.
	///
	/// \sa Raise_Destroyed,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632620.aspx">WM_DESTROY</a>
	LRESULT OnHostDialogDestroy(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_NCACTIVATE handler</em>
	///
	/// Will be called if the control host dialog is asked to change the activation style of its title bar.
	/// This happens if the window is activated or deactivated.
	/// We use this handler to change the dialog's parent window to activated state if the dialog is set
	/// to activated state, so that the dialog appears more integrated.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms632633.aspx">WM_NCACTIVATE</a>
	LRESULT OnHostDialogNCActivate(UINT /*message*/, WPARAM wParam, LPARAM /*lParam*/, BOOL& wasHandled);
	/// \brief <em>\c WM_NCLBUTTONDBLCLK handler</em>
	///
	/// Will be called if the control host dialog's title bar is double-clicked using the left mouse button.
	/// We use this handler to raise the \c TitleDblClick event.
	///
	/// \sa Raise_TitleDblClick,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms645619.aspx">WM_NCLBUTTONDBLCLK</a>
	LRESULT OnHostDialogNCLButtonDblClk(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_SYSCOMMAND handler</em>
	///
	/// Will be called if the control host dialog is asked to change the activation style of its title bar.
	/// This happens if the window is activated or deactivated.
	/// We use this handler to change the dialog's parent window to activated state if the dialog is set
	/// to activated state, so that the dialog appears more integrated.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms646360.aspx">WM_SYSCOMMAND</a>
	LRESULT OnHostDialogSysCommand(UINT message, WPARAM wParam, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGED handler</em>
	///
	/// Will be called if the control host window was moved or resized.
	/// We use this handler to raise the \c Moved and \c Resized events.
	///
	/// \sa Raise_Moved, Raise_Resized, OnHostDialogWindowPosChanging,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGED</a>
	LRESULT OnHostDialogWindowPosChanged(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c WM_WINDOWPOSCHANGING handler</em>
	///
	/// Will be called if the control host window is being moved or resized.
	/// We use this handler to raise the \c Moving and \c Resizing events.
	///
	/// \sa Raise_Moving, Raise_Resizing, OnHostDialogWindowPosChanged,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms632652.aspx">WM_WINDOWPOSCHANGING</a>
	LRESULT OnHostDialogWindowPosChanging(UINT /*message*/, WPARAM /*wParam*/, LPARAM lParam, BOOL& wasHandled);
	/// \brief <em>\c ControlHostWindow_InternalGetInstanceMessage handler</em>
	///
	/// Will be called if the control host window is being asked for the pointer to the \c ControlHostWindow
	/// object that controls the window.
	/// We use this handler in the \c WndProc hook to retrieve the associated \c ControlHostWindow object.
	///
	/// \sa CallWndProc, GetControlHostWindowMessage
	LRESULT OnHostDialogGetControlHostWindow(UINT /*message*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*wasHandled*/);
	/// \brief <em>\c WM_SETFOCUS handler</em>
	///
	/// Will be called if the control host dialog's parent window gained keyboard focus.
	/// We use this handler to enable the control host dialog window again, after it has been disabled by
	/// \c OnParentMouseActivate.
	///
	/// \sa CallWndProc,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms646283.aspx">WM_SETFOCUS</a>
	LRESULT OnParentSetFocus(UINT message, WPARAM wParam, LPARAM lParam, BOOL& /*wasHandled*/);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Event raisers
	///
	//@{
	/// \brief <em>Raises the \c Activate event</em>
	///
	/// \param[in] activatedByMouseClick If \c VARIANT_TRUE, the window is being activated in response to a
	///            mouse click; otherwise it is activated by using the keyboard or by code.
	/// \param[in] hWndBeingDeactivated The handle of the window that is being deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IControlHostWindowEvents::Fire_Activate,
	///       TBarCtlsLibU::_IControlHostWindowEvents::Activate, Raise_Deactivate
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_Activate,
	///       TBarCtlsLibA::_IControlHostWindowEvents::Activate, Raise_Deactivate
	/// \endif
	inline HRESULT Raise_Activate(VARIANT_BOOL activatedByMouseClick, LONG hWndBeingDeactivated);
	/// \brief <em>Raises the \c Closing event</em>
	///
	/// \param[out] pCancel If set to \c VARIANT_TRUE, the window shall be kept open; otherwise it shall be
	///             closed.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IControlHostWindowEvents::Fire_Closing, TBarCtlsLibU::_IControlHostWindowEvents::Closing,
	///       Raise_Destroyed, Raise_Created
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_Closing, TBarCtlsLibA::_IControlHostWindowEvents::Closing,
	///       Raise_Destroyed, Raise_Created
	/// \endif
	inline HRESULT Raise_Closing(VARIANT_BOOL* pCancel);
	/// \brief <em>Raises the \c Created event</em>
	///
	/// \param[in] hWnd The control hosts's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IControlHostWindowEvents::Fire_Created, TBarCtlsLibU::_IControlHostWindowEvents::Created,
	///       Raise_Destroyed, get_hWnd
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_Created, TBarCtlsLibA::_IControlHostWindowEvents::Created,
	///       Raise_Destroyed, get_hWnd
	/// \endif
	inline HRESULT Raise_Created(LONG hWnd);
	/// \brief <em>Raises the \c Deactivate event</em>
	///
	/// \param[in] hWndBeingActivated The handle of the window that is being deactivated.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IControlHostWindowEvents::Fire_Deactivate,
	///       TBarCtlsLibU::_IControlHostWindowEvents::Deactivate, Raise_Activate
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_Deactivate,
	///       TBarCtlsLibA::_IControlHostWindowEvents::Deactivate, Raise_Activate
	/// \endif
	inline HRESULT Raise_Deactivate(LONG hWndBeingActivated);
	/// \brief <em>Raises the \c Destroyed event</em>
	///
	/// \param[in] hWnd The control host's window handle.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IControlHostWindowEvents::Fire_Destroyed,
	///       TBarCtlsLibU::_IControlHostWindowEvents::Destroyed, Raise_Closing, Raise_Created, get_hWnd
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_Destroyed,
	///       TBarCtlsLibA::_IControlHostWindowEvents::Destroyed, Raise_Closing, Raise_Created, get_hWnd
	/// \endif
	inline HRESULT Raise_Destroyed(LONG hWnd);
	/// \brief <em>Raises the \c Moved event</em>
	///
	/// \param[in] left The x-coordinate (in pixels) of the window's position relative to the screen's
	///            upper-left corner.
	/// \param[in] top The y-coordinate (in pixels) of the window's position relative to the screen's
	///            upper-left corner.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IControlHostWindowEvents::Fire_Moved, TBarCtlsLibU::_IControlHostWindowEvents::Moved,
	///       Raise_Moving, Raise_Resized, Move
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_Moved, TBarCtlsLibA::_IControlHostWindowEvents::Moved,
	///       Raise_Moving, Raise_Resized, Move
	/// \endif
	inline HRESULT Raise_Moved(OLE_XPOS_PIXELS left, OLE_YPOS_PIXELS top);
	/// \brief <em>Raises the \c Moving event</em>
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
	///   \sa Proxy_IControlHostWindowEvents::Fire_Moving, TBarCtlsLibU::_IControlHostWindowEvents::Moving,
	///       Raise_Moved, Raise_Resizing, Move
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_Moving, TBarCtlsLibA::_IControlHostWindowEvents::Moving,
	///       Raise_Moved, Raise_Resizing, Move
	/// \endif
	inline HRESULT Raise_Moving(OLE_XPOS_PIXELS* pLeft, OLE_YPOS_PIXELS* pTop, VARIANT_BOOL* pPreventMove);
	/// \brief <em>Raises the \c Resized event</em>
	///
	/// \param[in] width The control host window's width in pixels.
	/// \param[in] height The control host window's height in pixels.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Proxy_IControlHostWindowEvents::Fire_Resized, TBarCtlsLibU::_IControlHostWindowEvents::Resized,
	///       Raise_Resizing, Raise_Moved
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_Resized, TBarCtlsLibA::_IControlHostWindowEvents::Resized,
	///       Raise_Resizing, Raise_Moved
	/// \endif
	inline HRESULT Raise_Resized(OLE_XSIZE_PIXELS width, OLE_YSIZE_PIXELS height);
	/// \brief <em>Raises the \c Resizing event</em>
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
	///   \sa Proxy_IControlHostWindowEvents::Fire_Resizing,
	///       TBarCtlsLibU::_IControlHostWindowEvents::Resizing, Raise_Resized, Raise_Moving
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_Resizing,
	///       TBarCtlsLibA::_IControlHostWindowEvents::Resizing, Raise_Resized, Raise_Moving
	/// \endif
	inline HRESULT Raise_Resizing(OLE_XSIZE_PIXELS* pWidth, OLE_YSIZE_PIXELS* pHeight, VARIANT_BOOL* pPreventResize);
	/// \brief <em>Raises the \c TitleDblClick event</em>
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
	///   \sa Proxy_IControlHostWindowEvents::Fire_TitleDblClick,
	///       TBarCtlsLibU::_IControlHostWindowEvents::TitleDblClick
	/// \else
	///   \sa Proxy_IControlHostWindowEvents::Fire_TitleDblClick,
	///       TBarCtlsLibA::_IControlHostWindowEvents::TitleDblClick
	/// \endif
	inline HRESULT Raise_TitleDblClick(SHORT button, SHORT shift, OLE_XPOS_PIXELS x, OLE_YPOS_PIXELS y);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>Holds the \c Resizeable property's setting</em>
		///
		/// \sa get_Resizeable, put_Resizeable
		UINT resizeable : 1;
		/// \brief <em>Holds the \c Visible property's setting</em>
		///
		/// \sa get_Visible, put_Visible
		UINT visible : 1;
		/// \brief <em>Holds the \c WindowTitle property's setting</em>
		///
		/// \sa get_WindowTitle, put_WindowTitle
		CComBSTR windowTitle;

		/// \brief <em>The dialog window hosting the control</em>
		ControlHostDialog hostDialog;
		/// \brief <em>The hosted control's previous parent window</em>
		HWND hWndOldParent;

		Properties()
		{
			ResetToDefaults();
		}

		/// \brief <em>Resets all properties to their defaults</em>
		void ResetToDefaults(void)
		{
			resizeable = TRUE;
			visible = TRUE;
			windowTitle = L"";

			hWndOldParent = NULL;
		}
	} properties;
};     // ControlHostWindow

OBJECT_ENTRY_AUTO(__uuidof(ControlHostWindow), ControlHostWindow)