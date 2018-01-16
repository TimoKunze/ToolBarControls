//////////////////////////////////////////////////////////////////////
/// \class ReBarBands
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Manages a collection of \c ReBarBand objects</em>
///
/// This class provides easy access (including filtering) to collections of \c ReBarBand objects.
/// A \c ReBarBands object is used to group bands that have certain properties in common.
///
/// \if UNICODE
///   \sa TBarCtlsLibU::IReBarBands, ReBarBand, ReBar
/// \else
///   \sa TBarCtlsLibA::IReBarBands, ReBarBand, ReBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "_IReBarBandsEvents_CP.h"
#include "helpers.h"
#include "ReBar.h"
#include "ReBarBand.h"


class ATL_NO_VTABLE ReBarBands : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ReBarBands, &CLSID_ReBarBands>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ReBarBands>,
    public Proxy_IReBarBandsEvents<ReBarBands>,
    public IEnumVARIANT,
    #ifdef UNICODE
    	public IDispatchImpl<IReBarBands, &IID_IReBarBands, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IReBarBands, &IID_IReBarBands, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ReBar;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_REBARBANDS)

		BEGIN_COM_MAP(ReBarBands)
			COM_INTERFACE_ENTRY(IReBarBands)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
			COM_INTERFACE_ENTRY(IEnumVARIANT)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ReBarBands)
			CONNECTION_POINT_ENTRY(__uuidof(_IReBarBandsEvents))
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
	/// \brief <em>Clones the \c VARIANT iterator used to iterate the bands</em>
	///
	/// Clones the \c VARIANT iterator including its current state. This iterator is used to iterate
	/// the \c ReBarBand objects managed by this collection object.
	///
	/// \param[out] ppEnumerator Receives the clone's \c IEnumVARIANT implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Next, Reset, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms221053.aspx">IEnumVARIANT</a>,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690336.aspx">IEnumXXXX::Clone</a>
	virtual HRESULT STDMETHODCALLTYPE Clone(IEnumVARIANT** ppEnumerator);
	/// \brief <em>Retrieves the next x bands</em>
	///
	/// Retrieves the next \c numberOfMaxBands bands from the iterator.
	///
	/// \param[in] numberOfMaxBands The maximum number of bands the array identified by \c pBands can
	///            contain.
	/// \param[in,out] pBands An array of \c VARIANT values. On return, each \c VARIANT will contain
	///                the pointer to a band's \c IReBarBand implementation.
	/// \param[out] pNumberOfBandsReturned The number of bands that actually were copied to the array
	///             identified by \c pBands.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Reset, Skip, ReBarBand,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms695273.aspx">IEnumXXXX::Next</a>
	virtual HRESULT STDMETHODCALLTYPE Next(ULONG numberOfMaxBands, VARIANT* pBands, ULONG* pNumberOfBandsReturned);
	/// \brief <em>Resets the \c VARIANT iterator</em>
	///
	/// Resets the \c VARIANT iterator so that the next call of \c Next or \c Skip starts at the first
	/// band in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Skip,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms693414.aspx">IEnumXXXX::Reset</a>
	virtual HRESULT STDMETHODCALLTYPE Reset(void);
	/// \brief <em>Skips the next x bands</em>
	///
	/// Instructs the \c VARIANT iterator to skip the next \c numberOfBandsToSkip bands.
	///
	/// \param[in] numberOfBandsToSkip The number of bands to skip.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Clone, Next, Reset,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms690392.aspx">IEnumXXXX::Skip</a>
	virtual HRESULT STDMETHODCALLTYPE Skip(ULONG numberOfBandsToSkip);
	//@}
	//////////////////////////////////////////////////////////////////////

	//////////////////////////////////////////////////////////////////////
	/// \name Implementation of IReBarBands
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c CaseSensitiveFilters property</em>
	///
	/// Retrieves whether string comparisons, that are done when applying the filters on a band, are case
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
	/// Sets whether string comparisons, that are done when applying the filters on a band, are case
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
	/// Retrieves a band filter's comparison function. This property takes the address of a function
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
	/// Sets a band filter's comparison function. This property takes the address of a function
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
	/// Retrieves a band filter.\n
	/// An \c IReBarBands collection can be filtered by any of \c IReBarBand's properties, that
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
	/// Sets a band filter.\n
	/// An \c IReBarBands collection can be filtered by any of \c IReBarBand's properties, that
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
	/// Retrieves a band filter's type.
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
	/// Sets a band filter's type.
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
	/// \brief <em>Retrieves a \c ReBarBand object from the collection</em>
	///
	/// Retrieves a \c ReBarBand object from the collection that wraps the rebar band identified by
	/// \c bandIdentifier.
	///
	/// \param[in] bandIdentifier A value that identifies the rebar band to be retrieved.
	/// \param[in] bandIdentifierType A value specifying the meaning of \c bandIdentifier. Any of the
	///            values defined by the \c BandIdentifierTypeConstants enumeration is valid.
	/// \param[out] ppBand Receives the band's \c IReBarBand implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ReBarBand, Add, Remove, Contains, TBarCtlsLibU::BandIdentifierTypeConstants
	/// \else
	///   \sa ReBarBand, Add, Remove, Contains, TBarCtlsLibA::BandIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Item(LONG bandIdentifier, BandIdentifierTypeConstants bandIdentifierType = bitIndex, IReBarBand** ppBand = NULL);
	/// \brief <em>Retrieves a \c VARIANT enumerator</em>
	///
	/// Retrieves a \c VARIANT enumerator that may be used to iterate the \c ReBarBand objects
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

	/// \brief <em>Adds a band to the rebar control</em>
	///
	/// Adds a band with the specified properties at the specified position in the control and returns a
	/// \c ReBarBand object wrapping the inserted band.
	///
	/// \param[in] insertAt The new band's zero-based index. If set to -1, the band will be inserted
	///            as the last band.
	/// \param[in] hContainedWindow The handle of the window that will be contained in the band.
	/// \param[in] newLine If set to \c VARIANT_TRUE, the band will be displayed in a new row within the
	///            control; otherwise it will be displayed in the same row as its preceding band.
	/// \param[in] sizingGripVisibility Specifies when the new band's sizing grip is displayed. Any of the
	///            values defined by the \c SizingGripVisibilityConstants enumeration is valid.
	/// \param[in] resizable If set to \c VARIANT_TRUE, the new band can be resized by the user; otherwise
	///            not.
	/// \param[in] keepInFirstRow If set to \c VARIANT_TRUE, the new band is always displayed within the
	///            first row; otherwise it may be moved to another row.
	/// \param[in] visible If set to \c VARIANT_TRUE, the new band is made visible; otherwise not.
	/// \param[in] idealWidth The width (in pixels), to which the control tries to resize the new band, if it
	///            auto-sizes it.
	/// \param[in] minimumWidth The new band's minimum width in pixels. The band won't be sized smaller than
	///            that.
	/// \param[in] minimumHeight The new band's minimum height in pixels. The band won't be sized smaller
	///            than that.
	/// \param[in] maximumHeight The new band's maximum height in pixels. The band won't be sized larger than
	///            that.
	/// \param[in] heightChangeStepSize The number of pixels by which the new band's height can be changed in
	///            a single step. If the control changes the band's height, the difference between old and
	///            new height will be a multiple of this value.
	/// \param[in] variableHeight If set to \c VARIANT_TRUE, the new band's height can be changed by the
	///            control; otherwise not.
	/// \param[in] showTitle If set to \c VARIANT_TRUE, the new band's caption, which may consist of text and
	///            an icon, is displayed; otherwise not.
	/// \param[in] text The new band's text. The maximum number of characters in this text is specified by
	///            \c MAX_BANDTEXTLENGTH.
	/// \param[in] iconIndex The zero-based index of the new band's icon in the control's \c ilBands image
	///            list. If set to -2, no icon is displayed for this band.
	/// \param[in] bandData A \c LONG value that will be associated with the new band.
	/// \param[in] titleWidth The width (in pixels) of the new band's caption, which may consist of text and
	///            an icon. If set to -1, the caption is sized automatically.
	/// \param[in] hideIfVertical If set to \c VARIANT_TRUE, the new band is displayed if the control's
	///            orientation is vertical; otherwise it is displayed on horizontal orientation only.
	/// \param[in] addMarginsAroundChild If set to \c VARIANT_TRUE, a small margin is inserted at the top and
	///            bottom of the new band's contained window; otherwise not.
	/// \param[out] ppAddedBand Receives the added band's \c IReBarBand implementation.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the parameter \c variableHeight is set to \c VARIANT_FALSE, the parameters
	///          \c maximumHeight and \c heightChangeStepSize are ignored.
	///
	/// \if UNICODE
	///   \sa Count, Remove, RemoveAll, ReBarBand::get_hContainedWindow, ReBarBand::get_NewLine,
	///       ReBarBand::get_SizingGripVisibility, ReBarBand::get_Resizable, ReBarBand::get_KeepInFirstRow,
	///       ReBarBand::get_Visible, ReBarBand::get_IdealWidth, ReBarBand::get_MinimumWidth,
	///       ReBarBand::get_MinimumHeight, ReBarBand::get_MaximumHeight,
	///       ReBarBand::get_HeightChangeStepSize, ReBarBand::get_VariableHeight, ReBarBand::get_ShowTitle,
	///       ReBarBand::get_Text, MAX_BANDTEXTLENGTH, ReBarBand::get_IconIndex, ReBar::get_hImageList,
	///       ReBarBand::get_BandData, ReBarBand::get_TitleWidth, ReBarBand::get_HideIfVertical,
	///       ReBar::get_Orientation, ReBarBand::get_AddMarginsAroundChild,
	///       TBarCtlsLibU::SizingGripVisibilityConstants
	/// \else
	///   \sa Count, Remove, RemoveAll, ReBarBand::get_hContainedWindow, ReBarBand::get_NewLine,
	///       ReBarBand::get_SizingGripVisibility, ReBarBand::get_Resizable, ReBarBand::get_KeepInFirstRow,
	///       ReBarBand::get_Visible, ReBarBand::get_IdealWidth, ReBarBand::get_MinimumWidth,
	///       ReBarBand::get_MinimumHeight, ReBarBand::get_MaximumHeight,
	///       ReBarBand::get_HeightChangeStepSize, ReBarBand::get_VariableHeight, ReBarBand::get_ShowTitle,
	///       ReBarBand::get_Text, MAX_BANDTEXTLENGTH, ReBarBand::get_IconIndex, ReBar::get_hImageList,
	///       ReBarBand::get_BandData, ReBarBand::get_TitleWidth, ReBarBand::get_HideIfVertical,
	///       ReBar::get_Orientation, ReBarBand::get_AddMarginsAroundChild,
	///       TBarCtlsLibA::SizingGripVisibilityConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Add(LONG insertAt = -1, OLE_HANDLE hContainedWindow = NULL, VARIANT_BOOL newLine = VARIANT_FALSE, SizingGripVisibilityConstants sizingGripVisibility = sgvAutomatic, VARIANT_BOOL resizable = VARIANT_TRUE, VARIANT_BOOL keepInFirstRow = VARIANT_FALSE, VARIANT_BOOL visible = VARIANT_TRUE, OLE_XSIZE_PIXELS idealWidth = 0, OLE_XSIZE_PIXELS minimumWidth = -1, OLE_YSIZE_PIXELS minimumHeight = -1, OLE_YSIZE_PIXELS maximumHeight = -1, OLE_YSIZE_PIXELS heightChangeStepSize = -1, VARIANT_BOOL variableHeight = VARIANT_FALSE, VARIANT_BOOL showTitle = VARIANT_TRUE, BSTR text = L"", LONG iconIndex = -2, LONG bandData = 0, OLE_XSIZE_PIXELS titleWidth = -1, VARIANT_BOOL hideIfVertical = VARIANT_FALSE, VARIANT_BOOL addMarginsAroundChild = VARIANT_FALSE, IReBarBand** ppAddedBand = NULL);
	/// \brief <em>Retrieves whether the specified band is part of the band collection</em>
	///
	/// \param[in] bandIdentifier A value that identifies the band to be checked.
	/// \param[in] bandIdentifierType A value specifying the meaning of \c bandIdentifier. Any of the
	///            values defined by the \c BandIdentifierTypeConstants enumeration is valid.
	/// \param[out] pValue \c VARIANT_TRUE, if the band is part of the collection; otherwise
	///             \c VARIANT_FALSE.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_Filter, Add, Remove, TBarCtlsLibU::BandIdentifierTypeConstants
	/// \else
	///   \sa get_Filter, Add, Remove, TBarCtlsLibA::BandIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Contains(LONG bandIdentifier, BandIdentifierTypeConstants bandIdentifierType = bitIndex, VARIANT_BOOL* pValue = NULL);
	/// \brief <em>Counts the bands in the collection</em>
	///
	/// Retrieves the number of \c ReBarBand objects in the collection.
	///
	/// \param[out] pValue The number of elements in the collection.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Add, Remove, RemoveAll
	virtual HRESULT STDMETHODCALLTYPE Count(LONG* pValue);
	/// \brief <em>Removes the specified band in the collection from the rebar</em>
	///
	/// \param[in] bandIdentifier A value that identifies the rebar band to be removed.
	/// \param[in] bandIdentifierType A value specifying the meaning of \c bandIdentifier. Any of the
	///            values defined by the \c BandIdentifierTypeConstants enumeration is valid.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa Add, Count, RemoveAll, Contains, TBarCtlsLibU::BandIdentifierTypeConstants
	/// \else
	///   \sa Add, Count, RemoveAll, Contains, TBarCtlsLibA::BandIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE Remove(LONG bandIdentifier, BandIdentifierTypeConstants bandIdentifierType = bitIndex);
	/// \brief <em>Removes all bands in the collection from the rebar</em>
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
	/// \sa Properties::pOwnerRBar
	void SetOwner(__in_opt ReBar* pOwner);

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
	/// \overload
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
	/// \brief <em>Retrieves the control's first band that might be in the collection</em>
	///
	/// \param[in] numberOfBands The number of rebar bands in the control.
	///
	/// \return The band being the collection's first candidate. -1 if no band was found.
	///
	/// \sa GetNextBandToProcess, Count, RemoveAll, Next
	int GetFirstBandToProcess(int numberOfBands);
	/// \brief <em>Retrieves the control's next band that might be in the collection</em>
	///
	/// \param[in] previousBand The band at which to start the search for a new collection candidate.
	/// \param[in] numberOfBands The number of bands in the control.
	///
	/// \return The next band being a candidate for the collection. -1 if no band was found.
	///
	/// \sa GetFirstBandToProcess, Count, RemoveAll, Next
	int GetNextBandToProcess(int previousBand, int numberOfBands);
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
	/// \brief <em>Retrieves whether a band is part of the collection (applying the filters)</em>
	///
	/// \param[in] bandIndex The band to check.
	/// \param[in] hWndRBar The rebar window the method will work on.
	///
	/// \return \c TRUE, if the band is part of the collection; otherwise \c FALSE.
	///
	/// \sa Contains, Count, Remove, RemoveAll, Next
	BOOL IsPartOfCollection(int bandIndex, HWND hWndRBar = NULL);
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
		/// \brief <em>Removes the specified bands</em>
		///
		/// \param[in] bandsToRemove A list containing all bands to remove.
		/// \param[in] hWndRBar The rebar window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveBands(std::list<int>& bandsToRemove, HWND hWndRBar);
	#else
		/// \brief <em>Removes the specified bands</em>
		///
		/// \param[in] bandsToRemove A list containing all bands to remove.
		/// \param[in] hWndRBar The rebar window the method will work on.
		///
		/// \return An \c HRESULT error code.
		HRESULT RemoveBands(CAtlList<int>& bandsToRemove, HWND hWndRBar);
	#endif

	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		#define NUMBEROFFILTERS_RB 44
		/// \brief <em>Holds the \c CaseSensitiveFilters property's setting</em>
		///
		/// \sa get_CaseSensitiveFilters, put_CaseSensitiveFilters
		UINT caseSensitiveFilters : 1;
		/// \brief <em>Holds the \c ComparisonFunction property's setting</em>
		///
		/// \sa get_ComparisonFunction, put_ComparisonFunction
		LPVOID comparisonFunction[NUMBEROFFILTERS_RB];
		/// \brief <em>Holds the \c Filter property's setting</em>
		///
		/// \sa get_Filter, put_Filter
		VARIANT filter[NUMBEROFFILTERS_RB];
		/// \brief <em>Holds the \c FilterType property's setting</em>
		///
		/// \sa get_FilterType, put_FilterType
		FilterTypeConstants filterType[NUMBEROFFILTERS_RB];

		/// \brief <em>The \c ReBar object that owns this collection</em>
		///
		/// \sa SetOwner
		ReBar* pOwnerRBar;
		/// \brief <em>Holds the last enumerated band</em>
		int lastEnumeratedBand;
		/// \brief <em>If \c TRUE, we must filter the bands</em>
		///
		/// \sa put_Filter, put_FilterType
		UINT usingFilters : 1;

		Properties()
		{
			caseSensitiveFilters = FALSE;
			pOwnerRBar = NULL;
			lastEnumeratedBand = -1;

			for(int i = 0; i < NUMBEROFFILTERS_RB; ++i) {
				VariantInit(&filter[i]);
				filterType[i] = ftDeactivated;
				comparisonFunction[i] = NULL;
			}
			usingFilters = FALSE;
		}

		~Properties();

		/// \brief <em>Copies this struct's content to another \c Properties struct</em>
		void CopyTo(Properties* pTarget);

		/// \brief <em>Retrieves the owning rebar's window handle</em>
		///
		/// \return The window handle of the rebar that contains the bands in this collection.
		///
		/// \sa pOwnerRBar
		HWND GetRBarHWnd(void);
	} properties;
};     // ReBarBands

OBJECT_ENTRY_AUTO(__uuidof(ReBarBands), ReBarBands)