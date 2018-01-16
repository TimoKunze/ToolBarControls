//////////////////////////////////////////////////////////////////////
/// \class ToolBarButtons
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ToolBarButton objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ToolBarButton objects.
/// A \c ToolBarButtons object is used to group buttons that have certain properties in common.
///
/// \if UNICODE
///   \sa TBarCtlsLibU::IToolBarButtons, ToolBarButton, ToolBar
/// \else
///   \sa TBarCtlsLibA::IToolBarButtons, ToolBarButton, ToolBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "_IToolBarButtonsEvents_CP.h"
#include "helpers.h"
#include "ToolBar.h"
#include "ToolBarButton.h"


class ATL_NO_VTABLE ToolBarButtons : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ToolBarButtons, &CLSID_ToolBarButtons>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ToolBarButtons>,
    public Proxy_IToolBarButtonsEvents<ToolBarButtons>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IToolBarButtons, &IID_IToolBarButtons, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IToolBarButtons, &IID_IToolBarButtons, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ToolBar;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_TOOLBARBUTTONS)

		BEGIN_COM_MAP(ToolBarButtons)
			COM_INTERFACE_ENTRY(IToolBarButtons)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ToolBarButtons)
			CONNECTION_POINT_ENTRY(__uuidof(_IToolBarButtonsEvents))
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
	///                the pointer to a button's \c IToolBarButton implementation.
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
	/// \name Implementation of IToolBarButtons
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c CaseSensitiveFilters property</em>
	///
	/// Retrieves whether string comparisons, that are done when applying the filters on a button, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CaseSensitiveFilters, get_Filter, get_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE get_CaseSensitiveFilters(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c CaseSensitiveFilters property</em>
	///
	/// Sets whether string comparisons, that are done when applying the filters on a button, are case
	/// sensitive. If this property is set to \c VARIANT_TRUE, string comparisons are case sensitive;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CaseSensitiveFilters, put_Filter, put_ComparisonFunction
	virtual HRESULT STDMETHODCALLTYPE put_CaseSensitiveFilters(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ComparisonFunction property</em>
	///
	/// Retrieves a button filter's comparison function. This property takes the address of a function
	/// having the following signature:\n
	/// \code
	///   BOOL IsEqual(T bandProperty, T pattern);
	/// \endcode
	/// where T stands for the filtered property's type (\c VARIANT_BOOL, \c LONG or \c BSTR). This function
	/// must compare its arguments and return a non-zero value if the arguments are equal and zero
	/// otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       TBarCtlsLibU::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa put_ComparisonFunction, get_Filter, get_CaseSensitiveFilters,
	///       TBarCtlsLibA::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG* pValue);
	/// \brief <em>Sets the \c ComparisonFunction property</em>
	///
	/// Sets a button filter's comparison function. This property takes the address of a function
	/// having the following signature:\n
	/// \code
	///   BOOL IsEqual(T itemProperty, T pattern);
	/// \endcode
	/// where T stands for the filtered property's type (\c VARIANT_BOOL, \c LONG or \c BSTR). This function
	/// must compare its arguments and return a non-zero value if the arguments are equal and zero
	/// otherwise.\n
	/// If this property is set to 0, the control compares the values itself using the "==" operator
	/// (\c lstrcmp and \c lstrcmpi for string filters).
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       TBarCtlsLibU::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \else
	///   \sa get_ComparisonFunction, put_Filter, put_CaseSensitiveFilters,
	///       TBarCtlsLibA::FilteredPropertyConstants,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647488.aspx">lstrcmp</a>,
	///       <a href="https://msdn.microsoft.com/en-us/library/ms647489.aspx">lstrcmpi</a>
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_ComparisonFunction(FilteredPropertyConstants filteredProperty, LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c Filter property</em>
	///
	/// Retrieves a button filter.\n
	/// An \c IToolBarButtons collection can be filtered by any of \c IToolBarButton's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction, TBarCtlsLibU::FilteredPropertyConstants
	/// \else
	///   \sa put_Filter, get_FilterType, get_ComparisonFunction, TBarCtlsLibA::FilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Filter(FilteredPropertyConstants filteredProperty, VARIANT* pValue);
	/// \brief <em>Sets the \c Filter property</em>
	///
	/// Sets a button filter.\n
	/// An \c IToolBarButtons collection can be filtered by any of \c IToolBarButton's properties, that
	/// the \c FilteredPropertyConstants enumeration defines a constant for. Combinations of multiple
	/// filters are possible, too. A filter is a \c VARIANT containing an array whose elements are of
	/// type \c VARIANT. Each element of this array contains a valid value for the property, that the
	/// filter refers to.\n
	/// When applying the filter, the elements of the array are connected using the logical Or operator.\n\n
	/// Setting this property to \c Empty or any other value, that doesn't match the described structure,
	/// deactivates the filter.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       TBarCtlsLibU::FilteredPropertyConstants
	/// \else
	///   \sa get_Filter, put_FilterType, put_ComparisonFunction, IsPartOfCollection,
	///       TBarCtlsLibA::FilteredPropertyConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Filter(FilteredPropertyConstants filteredProperty, VARIANT newValue);
	/// \brief <em>Retrieves the current setting of the \c FilterType property</em>
	///
	/// Retrieves a button filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_FilterType, get_Filter, TBarCtlsLibU::FilteredPropertyConstants,
	///       TBarCtlsLibU::FilterTypeConstants
	/// \else
	///   \sa put_FilterType, get_Filter, TBarCtlsLibA::FilteredPropertyConstants,
	///       TBarCtlsLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants* pValue);
	/// \brief <em>Sets the \c FilterType property</em>
	///
	/// Sets a button filter's type.
	///
	/// \param[in] filteredProperty A value specifying the property that the filter refers to. Any of the
	///            values defined by the \c FilteredPropertyConstants enumeration is valid.
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_FilterType, put_Filter, IsPartOfCollection, TBarCtlsLibU::FilteredPropertyConstants,
	///       TBarCtlsLibU::FilterTypeConstants
	/// \else
	///   \sa get_FilterType, put_Filter, IsPartOfCollection, TBarCtlsLibA::FilteredPropertyConstants,
	///       TBarCtlsLibA::FilterTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_FilterType(FilteredPropertyConstants filteredProperty, FilterTypeConstants newValue);
	/// \brief <em>Retrieves a \c ToolBarButton object from the collection</em>
	///
	/// Retrieves a \c ToolBarButton object from the collection that wraps the tool bar button identified by
	/// \c buttonIdentifier.
	///
	/// \param[in] buttonIdentifier A value that identifies the tool bar button to be retrieved.
	/// \param[in] buttonIdentifierType A value specifying the meaning of \c buttonIdentifier. Any of the
	///            values defined by the \c ButtonIdentifierTypeConstants enumeration is valid.
	/// \param[in] getChevronToolBarButton If \c VARIANT_TRUE, the retrieved button will refer to the chevron
	///            popup tool bar control.
	/// \param[out] ppButton Receives the button's \c IToolBarButton implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ToolBarButton, Add, Remove, Contains, TBarCtlsLibU::ButtonIdentifierTypeConstants
	/// \else
	///   \sa ToolBarButton, Add, Remove, Contains, TBarCtlsLibA::ButtonIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG buttonIdentifier, ButtonIdentifierTypeConstants buttonIdentifierType = btitIndex, VARIANT_BOOL getChevronToolBarButton = VARIANT_FALSE, IToolBarButton** ppButton = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ToolBarButton objects
	/// managed by this collection object. This iterator is used by Visual Basic's \c For...Each
	/// construct.
	///
	/// \param[out] ppEnumerator A pointer to the iterator's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only and hidden.
	///
	/// \sa <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>
	virtual HRESULT STDMETHODCALLTYPE get__NewEnum(IUnknown** ppEnumerator);

	/// \brief <em>Adds a button to the tool bar control</em>
	///
	/// Adds a button with the specified properties at the specified position in the control and returns a
	/// \c ToolBarButton object wrapping the inserted button.
	///
	/// \param[in] id A unique ID identifying the new button. Actually button IDs do not need to be unique,
	///            but it is strongly recommended that they are.
	/// \param[in] insertAt The new button's zero-based index. If set to -1, the button will be inserted
	///            as the last button.
	/// \param[in] iconIndex The zero-based index of the new button's icon in the control's image lists.
	///            If set to -1, the control will fire the \c ButtonGetDisplayInfo event each time this
	///            property's value is required. If set to -2, no icon is displayed for this button.
	/// \param[in] imageListIndex The zero-based index of the control's image lists that the button's icons
	///            will be taken from. If set to -1, the control will fire the \c ButtonGetDisplayInfo event
	///            each time this property's value is required.
	/// \param[in] text The new button's text.
	/// \param[in] displayText If set to \c VARIANT_TRUE, the new button's text is displayed next to its
	///            icon; otherwise it is displayed as tool tip.
	/// \param[in] useMnemonic If set to \c VARIANT_TRUE, the first character of the new button's text that
	///            is prefixed by an ampersand (&amp;), is used as the button's accelerator prefix and this
	///            character is underlined. If set to \c VARIANT_FALSE, ampersands in the button's text are
	///            interpreted as normal text.
	/// \param[in] buttonType Specifies which kind of button the new button is. Any of the values defined
	///            by the \c ButtonTypeConstants enumeration is valid.
	/// \param[in] partOfGroup If set to \c VARIANT_TRUE, the new button is part of a button group; otherwise
	///            not.\n
	///            Among all consecutive buttons that are part of a group, only one button can be checked. A
	///            check button that is part of a group cannot be unchecked by simply clicking it a second
	///            time. This can be used to create option-button like behavior.
	/// \param[in] selectionState Specifies the state of the new button if it is a check button. Any of
	///            the values defined by the \c SelectionStateConstants enumeration is valid.
	/// \param[in] dropDownStyle Specifies whether the new button is a drop-down button and how the
	///            drop-down arrow is displayed. Any of the values defined by the \c DropDownStyleConstants
	///            enumeration is valid.
	/// \param[in] buttonData A \c LONG value that will be associated with the new button.
	/// \param[in] visible If set to \c VARIANT_TRUE, the new button is made visible; otherwise not.
	/// \param[in] enabled If set to \c VARIANT_TRUE, the new button reacts to user input; otherwise not.
	/// \param[in] autoSize If set to \c VARIANT_TRUE, the new button's width is adjusted automatically
	///            depending on its text and icon; otherwise not.
	/// \param[in] width The new button's width in pixels.
	/// \param[in] followedByLineBreak If set to \c VARIANT_TRUE, the next button after the new button is
	///            placed on a new line; otherwise not.
	/// \param[in] showingEllipsis If set to \c VARIANT_TRUE, the new button's text is cut off and an
	///            ellipsis is displayed; otherwise not.
	/// \param[in] marked If set to \c VARIANT_TRUE, the new button is tagged as "Marked"; otherwise not.
	///            It's up to the application how marked buttons are interpreted.
	/// \param[out] ppAddedButton Receives the added button's \c IToolBarButton implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks The parameter \c displayText is ignored, if the \c IToolBar::ButtonTextPosition property
	///          is not set to \c btpRightToIcon or the \c IToolBar::AlwaysDisplayButtonText property is
	///          set to \c VARIANT_TRUE.\n
	///          When inserting a separator, the \c autoSize parameter should be set to \c VARIANT_FALSE.\n
	///          For placeholder buttons (buttons of type \c btyPlaceholder), only the following parameters
	///          are used:
	///          - \c id
	///          - \c insertAt
	///          - \c buttonType
	///          - \c partOfGroup
	///          - \c buttonData
	///          - \c visible
	///          - \c width
	///          - \c followedByLineBreak
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, ToolBarButton::get_ID, ToolBarButton::get_IconIndex,
	///       ToolBarButton::get_ImageListIndex, ToolBarButton::get_Text, ToolBarButton::get_DisplayText,
	///       ToolBarButton::get_UseMnemonic, ToolBarButton::get_ButtonType, ToolBarButton::get_PartOfGroup,
	///       ToolBarButton::get_SelectionState, ToolBarButton::get_DropDownStyle,
	///       ToolBarButton::get_ButtonData, ToolBarButton::get_Visible, ToolBarButton::get_Enabled,
	///       ToolBarButton::get_AutoSize, ToolBarButton::get_Width, ToolBarButton::get_FollowedByLineBreak,
	///       ToolBarButton::get_ShowingEllipsis, ToolBarButton::get_Marked,
	///       TBarCtlsLibU::ButtonTypeConstants, TBarCtlsLibU::SelectionStateConstants,
	///       TBarCtlsLibU::DropDownStyleConstants
	/// \else
	///   \sa Count, Remove, RemoveAll, ToolBarButton::get_ID, ToolBarButton::get_IconIndex,
	///       ToolBarButton::get_ImageListIndex, ToolBarButton::get_Text, ToolBarButton::get_DisplayText,
	///       ToolBarButton::get_UseMnemonic, ToolBarButton::get_ButtonType, ToolBarButton::get_PartOfGroup,
	///       ToolBarButton::get_SelectionState, ToolBarButton::get_DropDownStyle,
	///       ToolBarButton::get_ButtonData, ToolBarButton::get_Visible, ToolBarButton::get_Enabled,
	///       ToolBarButton::get_AutoSize, ToolBarButton::get_Width, ToolBarButton::get_FollowedByLineBreak,
	///       ToolBarButton::get_ShowingEllipsis, ToolBarButton::get_Marked,
	///       TBarCtlsLibA::ButtonTypeConstants, TBarCtlsLibA::SelectionStateConstants,
	///       TBarCtlsLibA::DropDownStyleConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(LONG id, LONG insertAt = -1, LONG iconIndex = -2, LONG imageListIndex = 0, BSTR text = L"", VARIANT_BOOL displayText = VARIANT_FALSE, VARIANT_BOOL useMnemonic = VARIANT_TRUE, ButtonTypeConstants buttonType = btyCommandButton, VARIANT_BOOL partOfGroup = VARIANT_FALSE, SelectionStateConstants selectionState = ssUnchecked, DropDownStyleConstants dropDownStyle = ddstNoDropDown, LONG buttonData = 0, VARIANT_BOOL visible = VARIANT_TRUE, VARIANT_BOOL enabled = VARIANT_TRUE, VARIANT_BOOL autoSize = VARIANT_TRUE, OLE_XSIZE_PIXELS width = -1, VARIANT_BOOL followedByLineBreak = VARIANT_FALSE, VARIANT_BOOL showingEllipsis = VARIANT_FALSE, VARIANT_BOOL marked = VARIANT_FALSE, IToolBarButton** ppAddedButton = NULL);
	/// \brief <em>Retrieves whether the specified button is part of the button collection</em>
	///
	/// \param[in] buttonIdentifier A value that identifies the button to be checked.
	/// \param[in] buttonIdentifierType A value specifying the meaning of \c buttonIdentifier. Any of the
	///            values defined by the \c ButtonIdentifierTypeConstants enumeration is valid.
	/// \param[in] checkChevronToolBarButton If \c VARIANT_TRUE, the method will check the chevron popup
	///            tool bar control.
	/// \param[out] pValue \c VARIANT_TRUE, if the button is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, Add, Remove, TBarCtlsLibU::ButtonIdentifierTypeConstants
	/// \else
	///   \sa get_Filter, Add, Remove, TBarCtlsLibA::ButtonIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(LONG buttonIdentifier, ButtonIdentifierTypeConstants buttonIdentifierType = btitIndex, VARIANT_BOOL checkChevronToolBarButton = VARIANT_FALSE, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the buttons in the collection</em>
	///
	/// Retrieves the number of \c ToolBarButton objects in the collection.
	///
	/// \param[in] countChevronToolBarButtons If \c VARIANT_TRUE, the method will count the buttons of the
	///            chevron popup tool bar control.
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(VARIANT_BOOL countChevronToolBarButtons = VARIANT_FALSE, LONG* pValue = NULL);
	/// \brief <em>Removes the specified button in the collection from the tool bar</em>
	///
	/// \param[in] buttonIdentifier A value that identifies the tool bar button to be removed.
	/// \param[in] buttonIdentifierType A value specifying the meaning of \c buttonIdentifier. Any of the
	///            values defined by the \c ButtonIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, TBarCtlsLibU::ButtonIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, TBarCtlsLibA::ButtonIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG buttonIdentifier, ButtonIdentifierTypeConstants buttonIdentifierType = btitIndex);
	/// \brief <em>Removes all buttons in the collection from the tool bar</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Count, Remove
	virtual HRESULT STDMETHODCALLTYPE RemoveAll(void);
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
	/// \name Filter validation
	///
	//@{
	/// \brief <em>Validates a filter for a \c VARIANT_BOOL type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type \c VARIANT_BOOL.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidIntegerFilter, IsValidStringFilter, put_Filter
	BOOL IsValidBooleanFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c LONG or compatible.
	///
	/// \param[in] filter The \c VARIANT to check.
	/// \param[in] minValue The minimum value that the corresponding property would accept.
	/// \param[in] maxValue The maximum value that the corresponding property would accept.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidStringFilter, put_Filter
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue, int maxValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type
	/// \c LONG or compatible.
	///
	/// \param[in] filter The \c VARIANT to check.
	/// \param[in] minValue The minimum value that the corresponding property would accept.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidStringFilter, put_Filter
	BOOL IsValidIntegerFilter(VARIANT& filter, int minValue);
	/// \brief <em>Validates a filter for a \c LONG (or compatible) type property</em>
	///
	/// \overload
	BOOL IsValidIntegerFilter(VARIANT& filter);
	/// \brief <em>Validates a filter for a \c BSTR type property</em>
	///
	/// Retrieves whether a \c VARIANT value can be used as a filter for a property of type \c BSTR.
	///
	/// \param[in] filter The \c VARIANT to check.
	///
	/// \return \c TRUE, if the filter is valid; otherwise \c FALSE.
	///
	/// \sa IsValidBooleanFilter, IsValidIntegerFilter, put_Filter
	BOOL IsValidStringFilter(VARIANT& filter);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Filter appliance
	///
	//@{
	/// \brief <em>Retrieves the control's first button that might be in the collection</em>
	///
	/// \param[in] numberOfButtons The number of tool bar buttons in the control.
	///
	/// \return The button being the collection's first candidate. -1 if no button was found.
	///
	/// \sa GetNextButtonToProcess, Count, RemoveAll, Next
	int GetFirstButtonToProcess(int numberOfButtons);
	/// \brief <em>Retrieves the control's next button that might be in the collection</em>
	///
	/// \param[in] previousButton The button at which to start the search for a new collection candidate.
	/// \param[in] numberOfButtons The number of buttons in the control.
	///
	/// \return The next button being a candidate for the collection. -1 if no button was found.
	///
	/// \sa GetFirstButtonToProcess, Count, RemoveAll, Next
	int GetNextButtonToProcess(int previousButton, int numberOfButtons);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified boolean value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsIntegerInSafeArray, IsStringInSafeArray, get_ComparisonFunction
	BOOL IsBooleanInSafeArray(LPSAFEARRAY pSafeArray, VARIANT_BOOL value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified integer value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsStringInSafeArray, get_ComparisonFunction
	BOOL IsIntegerInSafeArray(LPSAFEARRAY pSafeArray, int value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether the specified \c SAFEARRAY contains the specified \c BSTR value</em>
	///
	/// \param[in] pSafeArray The \c SAFEARRAY to search.
	/// \param[in] value The value to search for.
	/// \param[in] pComparisonFunction The address of the comparison function to use.
	///
	/// \return \c TRUE, if the array contains the value; otherwise \c FALSE.
	///
	/// \sa IsPartOfCollection, IsBooleanInSafeArray, IsIntegerInSafeArray, get_ComparisonFunction
	BOOL IsStringInSafeArray(LPSAFEARRAY pSafeArray, BSTR value, LPVOID pComparisonFunction);
	/// \brief <em>Retrieves whether a button is part of the collection (applying the filters)</em>
	///
	/// \param[in] buttonIndex The button to check.
	/// \param[in] hWndTBar The tool bar window the method will work on.
	///
	/// \return \c TRUE, if the button is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(int buttonIndex, HWND hWndTBar = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Shortens a filter as much as possible</em>
	///
	/// Optimizes a filter by detecting redundancies, tautologies and so on.
	///
	/// \param[in] filteredProperty The filter to optimize. Any of the values defined by the
	///            \c FilteredPropertyConstants enumeration is valid.
	///
	/// \sa put_Filter, put_FilterType
	void OptimizeFilter(FilteredPropertyConstants filteredProperty);
	#ifdef USE_STL
		/// \brief <em>Removes the specified buttons</em>
		///
		/// \param[in] buttonsToRemove A list containing all buttons to remove.
		/// \param[in] hWndTBar The tool bar window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveButtons(std::list<int>& buttonsToRemove, HWND hWndTBar);
	#else
		/// \brief <em>Removes the specified buttons</em>
		///
		/// \param[in] buttonsToRemove A list containing all buttons to remove.
		/// \param[in] hWndTBar The tool bar window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveButtons(CAtlList<int>& buttonsToRemove, HWND hWndTBar);
	#endif

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		#define NUMBEROFFILTERS_TB 53
		/// \brief <em>Holds the \c CaseSensitiveFilters property's setting</em>
		///
		/// \sa get_CaseSensitiveFilters, put_CaseSensitiveFilters
		UINT caseSensitiveFilters : 1;
		/// \brief <em>Holds the \c ComparisonFunction property's setting</em>
		///
		/// \sa get_ComparisonFunction, put_ComparisonFunction
		LPVOID comparisonFunction[NUMBEROFFILTERS_TB];
		/// \brief <em>Holds the \c Filter property's setting</em>
		///
		/// \sa get_Filter, put_Filter
		VARIANT filter[NUMBEROFFILTERS_TB];
		/// \brief <em>Holds the \c FilterType property's setting</em>
		///
		/// \sa get_FilterType, put_FilterType
		FilterTypeConstants filterType[NUMBEROFFILTERS_TB];

		/// \brief <em>The \c ToolBar object that owns this collection</em>
		///
		/// \sa SetOwner
		ToolBar* pOwnerTBar;
		/// \brief <em>Holds the last enumerated button</em>
		int lastEnumeratedButton;
		/// \brief <em>If \c TRUE, we must filter the buttons</em>
		///
		/// \sa put_Filter, put_FilterType
		UINT usingFilters : 1;

		Properties()
		{
			caseSensitiveFilters = FALSE;
			pOwnerTBar = NULL;
			lastEnumeratedButton = -1;

			for(int i = 0; i < NUMBEROFFILTERS_TB; ++i) {
				VariantInit(&filter[i]);
				filterType[i] = ftDeactivated;
				comparisonFunction[i] = NULL;
			}
			usingFilters = FALSE;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);

		/// \brief <em>Retrieves the owning tool bar's window handle</em>
		///
		/// \param[in] getChevronPopup If \c TRUE, the method will retrieve the window handle of the popup
		///            chevron tool bar.
		///
		/// \return The window handle of the tool bar that contains the buttons in this collection.
		///
		/// \sa pOwnerTBar
		HWND GetTBarHWnd(BOOL getChevronPopup = FALSE);
	} properties;
};     // ToolBarButtons

OBJECT_ENTRY_AUTO(__uuidof(ToolBarButtons), ToolBarButtons)