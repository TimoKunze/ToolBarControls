//////////////////////////////////////////////////////////////////////
/// \class ToolBarButtonContainer
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ToolBarButton objects</em>
///
/// This class provides easy access to collections of \c ToolBarButton objects. While a
/// \c ToolBarButtons object is used to group tool bar buttons that have certain properties in common, a
/// \c ToolBarButtonContainer object is used to group any buttons and acts more like a clipboard.
///
/// \if UNICODE
///   \sa TBarCtlsLibU::IToolBarButtonContainer, ToolBarButton, ToolBarButtons, ToolBar
/// \else
///   \sa TBarCtlsLibA::IToolBarButtonContainer, ToolBarButton, ToolBarButtons, ToolBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "_IToolBarButtonContainerEvents_CP.h"
#include "IButtonContainer.h"
#include "helpers.h"
#include "ToolBar.h"
#include "ToolBarButton.h"


class ATL_NO_VTABLE ToolBarButtonContainer : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ToolBarButtonContainer, &CLSID_ToolBarButtonContainer>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ToolBarButtonContainer>,
    public Proxy_IToolBarButtonContainerEvents<ToolBarButtonContainer>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IToolBarButtonContainer, &IID_IToolBarButtonContainer, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #else
    	public IDispatchImpl<IToolBarButtonContainer, &IID_IToolBarButtonContainer, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>,
    #endif
    public IButtonContainer
{
	friend class ToolBar;
	friend class ToolBarButtonContainer;

public:
	/// \brief <em>The constructor of this class</em>
	///
	/// Used to generate and set this object's ID.
	///
	/// \sa ~ToolBarButtonContainer
	ToolBarButtonContainer();
	/// \brief <em>The destructor of this class</em>
	///
	/// Used to deregister the container.
	///
	/// \sa ToolBarButtonContainer
	~ToolBarButtonContainer();

	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TOOLBARBUTTONCONTAINER)

		BEGIN_COM_MAP(ToolBarButtonContainer)
			COM_INTERFACE_ENTRY(IToolBarButtonContainer)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ToolBarButtonContainer)
			CONNECTION_POINT_ENTRY(__uuidof(_IToolBarButtonContainerEvents))
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
	/// \name Implementation of IEnumVARIANT
	///
	//@{
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the buttons</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ToolBarButton objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x buttons</em>
	///
	/// Retrieves the next \c numberOfMaxButtons buttons from the iterator.
	///
	/// \param[in] numberOfMaxButtons The maximum number of buttons the array identified by \c pButtons can
	///            contain.
	/// \param[in,out] pButtons An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to an button's \c IToolBarButton implementation.
	/// \param[out] pNumberOfButtonsReturned The number of buttons that actually were copied to the array
	///             identified by \c pButtons.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ToolBarButton,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxButtons, VARIANT* pButtons, ULONG* pNumberOfButtonsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// button in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x buttons</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfButtonsToSkip buttons.
	///
	/// \param[in] numberOfButtonsToSkip The number of buttons to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfButtonsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IToolBarButtonContainer
	///
	//@{
	/// \brief <em>Retrieves a \c ToolBarButton object from the collection</em>
	///
	/// Retrieves a \c ToolBarButton object from the collection that wraps the tool bar button identified
	/// by \c buttonID.
	///
	/// \param[in] buttonID The unique ID of the tool bar button to retrieve.
	/// \param[out] ppButton Receives the item's \c IToolBarButton implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ToolBarButton::get_ID, Add, Remove
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG buttonID, IToolBarButton** ppButton);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ToolBarButton objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator Receives the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds the specified button(s) to the collection</em>
	///
	/// \param[in] buttons The button(s) to add. May be either a button ID, a \c ToolBarButton object or a
	///            \c ToolBarButtons collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ToolBarButton::get_ID, Count, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Add(VARIANT buttons);
	/// \brief <em>Clones the collection object</em>
	///
	/// Retrieves an exact copy of the collection.
	///
	/// \param[out] ppClone Receives the clone's \c IToolBarButtonContainer implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ToolBar::CreateButtonContainer
	virtual HRESULT STDMETHODCALLTYPE Clone(IToolBarButtonContainer** ppClone);
	/// \brief <em>Counts the buttons in the collection</em>
	///
	/// Retrieves the number of \c ToolBarButton objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Retrieves an imagelist containing the buttons' common drag image</em>
	///
	/// Retrieves the handle to an imagelist containing a bitmap that can be used to visualize
	/// dragging of the buttons of this collection.
	///
	/// \param[out] pXUpperLeft The x-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] pYUpperLeft The y-coordinate (in pixels) of the drag image's upper-left corner relative
	///             to the control's upper-left corner.
	/// \param[out] phImageList The handle to the imagelist.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The caller is responsible for destroying the imagelist.
	///
	/// \sa ToolBar::CreateLegacyDragImage
	virtual HRESULT STDMETHODCALLTYPE CreateDragImage(OLE_XPOS_PIXELS* pXUpperLeft = NULL, OLE_YPOS_PIXELS* pYUpperLeft = NULL, OLE_HANDLE* phImageList = NULL);
	/// \brief <em>Removes the specified button from the collection</em>
	///
	/// \param[in] buttonID The unique ID of the tool bar button to remove.
	/// \param[in] removePhysically If \c VARIANT_TRUE, the button is removed from the control, too.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa ToolBarButton::get_ID, Add, Count, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG buttonID, VARIANT_BOOL removePhysically = VARIANT_FALSE);
	/// \brief <em>Removes all buttons from the collection</em>
	///
	/// \param[in] removePhysically If \c VARIANT_TRUE, the buttons get removed from the control, too.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(VARIANT_BOOL removePhysically = VARIANT_FALSE);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Sets the owner of this collection</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerTBar
	void SetOwner(__in_opt ToolBar* pOwner);

protected:
	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IButtonContainer
	///
	//@{
	/// \brief <em>A button was removed and the button container should check its content</em>
	///
	/// \param[in] buttonID The unique ID of the removed button. If -1, all buttons were removed.
	///
	/// \sa ToolBar::RemoveButtonFromButtonContainers
	void RemovedButton(LONG buttonID);
	/// \brief <em>Retrieves the container's ID used to identify it</em>
	///
	/// \return The container's ID.
	///
	/// \sa ToolBar::DeregisterButtonContainer, containerID
	DWORD GetID(void);
	/// \brief <em>Holds the container's ID</em>
	DWORD containerID;
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ToolBar object that owns this collection</em>
		ToolBar* pOwnerTBar;
		#ifdef USE_STL
			/// \brief <em>Holds the buttons that this object contains</em>
			std::vector<LONG> buttons;
		#else
			/// \brief <em>Holds the buttons that this object contains</em>
			//CAtlArray<LONG> buttons;
		#endif
		/// \brief <em>Points to the next enumerated button</em>
		int nextEnumeratedButton;

		Properties()
		{
			pOwnerTBar = NULL;
			nextEnumeratedButton = 0;
		}

		~Properties();

		/// \brief <em>Retrieves the owning tool bar's window handle</em>
		///
		/// \return The window handle of the tool bar that contains the buttons in this collection.
		///
		/// \sa pOwnerTBar
		HWND GetTBarHWnd(void);
	} properties;

	/* TODO: If we move this one into the Properties struct, the compiler complains that the private
	         = operator of CAtlArray cannot be accessed?! */
	#ifndef USE_STL
		/// \brief <em>Holds the buttons that this object contains</em>
		CAtlArray<LONG> buttons;
	#endif

	/// \brief <em>Incremented by the constructor and used as the constructed object's ID</em>
	///
	/// \sa ToolBarButtonContainer, containerID
	static DWORD nextID;
};     // ToolBarButtonContainer

OBJECT_ENTRY_AUTO(__uuidof(ToolBarButtonContainer), ToolBarButtonContainer)