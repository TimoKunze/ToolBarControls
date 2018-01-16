//////////////////////////////////////////////////////////////////////
/// \class ReBarBand
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps an existing rebar band</em>
///
/// This class is a wrapper around a rebar band that - unlike a band wrapped by
/// \c VirtualReBarBand - really exists in the control.
///
/// \if UNICODE
///   \sa TBarCtlsLibU::IReBarBand, VirtualReBarBand, ReBarBands, ReBar
/// \else
///   \sa TBarCtlsLibA::IReBarBand, VirtualReBarBand, ReBarBands, ReBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "_IReBarBandEvents_CP.h"
#include "helpers.h"
#include "ReBar.h"


class ATL_NO_VTABLE ReBarBand : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<ReBarBand, &CLSID_ReBarBand>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<ReBarBand>,
    public Proxy_IReBarBandEvents<ReBarBand>, 
    #ifdef UNICODE
    	public IDispatchImpl<IReBarBand, &IID_IReBarBand, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IReBarBand, &IID_IReBarBand, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ReBar;
	friend class ReBarBands;
	friend class ClassFactory;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_REBARBAND)

		BEGIN_COM_MAP(ReBarBand)
			COM_INTERFACE_ENTRY(IReBarBand)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(ReBarBand)
			CONNECTION_POINT_ENTRY(__uuidof(_IReBarBandEvents))
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
	/// \name Implementation of IReBarBand
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AddMarginsAroundChild property</em>
	///
	/// Retrieves whether the band's height is increased by a small margin. If set to \c VARIANT_TRUE, a
	/// small margin is inserted at the top and bottom of the band's contained window; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_AddMarginsAroundChild, get_CurrentHeight, ReBar::GetMargins, ReBar::get_DisplayBandSeparators
	virtual HRESULT STDMETHODCALLTYPE get_AddMarginsAroundChild(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c AddMarginsAroundChild property</em>
	///
	/// Sets whether the band's height is increased by a small margin. If set to \c VARIANT_TRUE, a
	/// small margin is inserted at the top and bottom of the band's contained window; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_AddMarginsAroundChild, put_CurrentHeight, ReBar::GetMargins, ReBar::put_DisplayBandSeparators
	virtual HRESULT STDMETHODCALLTYPE put_AddMarginsAroundChild(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the band's background color. If set to -1, the control's background color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa put_BackColor, get_hBackgroundBitmap, get_ForeColor, ReBar::get_BackColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c BackColor property</em>
	///
	/// Sets the band's background color. If set to -1, the control's background color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa get_BackColor, put_hBackgroundBitmap, put_ForeColor, ReBar::put_BackColor
	virtual HRESULT STDMETHODCALLTYPE put_BackColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c BandData property</em>
	///
	/// Retrieves the \c LONG value associated with the band. Use this property to associate any data
	/// with the band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_BandData, ReBar::Raise_FreeBandData
	virtual HRESULT STDMETHODCALLTYPE get_BandData(LONG* pValue);
	/// \brief <em>Sets the \c BandData property</em>
	///
	/// Sets the \c LONG value associated with the band. Use this property to associate any data
	/// with the band.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_BandData, ReBar::Raise_FreeBandData
	virtual HRESULT STDMETHODCALLTYPE put_BandData(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ChevronButtonObjectState property</em>
	///
	/// Retrieves the accessibility object state of the band's chevron button. For a list of possible object
	/// states see the <a href="https://msdn.microsoft.com/en-us/library/ms697270.aspx">MSDN article</a>.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Requires comctl32.dll version 6.10 or higher.\n
	///          This property is read-only.
	///
	/// \sa get_UseChevron,
	///     <a href="https://msdn.microsoft.com/en-us/library/ms697270.aspx">Accessibility Object State Constants</a>
	virtual HRESULT STDMETHODCALLTYPE get_ChevronButtonObjectState(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c ChevronHot property</em>
	///
	/// Retrieves whether the mouse cursor is located over the band's chevron button. If \c VARIANT_TRUE, the
	/// chevron button is hot (i.e. below the mouse); otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_UseChevron, get_ChevronVisible, get_ChevronPushed
	virtual HRESULT STDMETHODCALLTYPE get_ChevronHot(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ChevronPushed property</em>
	///
	/// Retrieves whether the band's chevron button currently is pushed. If \c VARIANT_TRUE, the chevron
	/// button is pushed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_UseChevron, get_ChevronVisible, get_ChevronHot
	virtual HRESULT STDMETHODCALLTYPE get_ChevronPushed(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ChevronVisible property</em>
	///
	/// Retrieves whether the band's chevron button is visible. If \c VARIANT_TRUE, the chevron button is
	/// being displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_UseChevron, get_IdealWidth, get_ChevronHot, get_ChevronPushed
	virtual HRESULT STDMETHODCALLTYPE get_ChevronVisible(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c CurrentHeight property</em>
	///
	/// Retrieves the band's current height in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c VariableHeight property is set to \c VARIANT_FALSE, this property is ignored.
	///
	/// \sa put_CurrentHeight, get_CurrentWidth, get_MinimumHeight, get_MaximumHeight,
	///     get_HeightChangeStepSize, get_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE get_CurrentHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c CurrentHeight property</em>
	///
	/// Sets the band's current height in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c VariableHeight property is set to \c VARIANT_FALSE, this property is ignored.
	///
	/// \sa get_CurrentHeight, put_CurrentWidth, put_MinimumHeight, put_MaximumHeight,
	///     put_HeightChangeStepSize, put_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE put_CurrentHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c CurrentWidth property</em>
	///
	/// Retrieves the band's current width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_CurrentWidth, get_CurrentHeight, get_MinimumWidth, get_IdealWidth, get_TitleWidth, Maximize,
	///     Minimize
	virtual HRESULT STDMETHODCALLTYPE get_CurrentWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c CurrentWidth property</em>
	///
	/// Sets the band's current width in pixels.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_CurrentWidth, put_CurrentHeight, put_MinimumWidth, put_IdealWidth, put_TitleWidth, Maximize,
	///     Minimize
	virtual HRESULT STDMETHODCALLTYPE put_CurrentWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c FixedBackgroundBitmapOrigin property</em>
	///
	/// Retrieves whether the band's background bitmap begins at the control's upper-left corner instead of
	/// the band's upper-left corner. If set to \c VARIANT_TRUE, the bitmap's origin is the control's
	/// upper-left corner, so that the background image doesn't move if the band is resized or moved (instead
	/// the visible part of the background image changes). If set to \c VARIANT_FALSE, the origin is the
	/// band's upper-left corner, so that the background image moves with the band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_FixedBackgroundBitmapOrigin, get_hBackgroundBitmap
	virtual HRESULT STDMETHODCALLTYPE get_FixedBackgroundBitmapOrigin(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c FixedBackgroundBitmapOrigin property</em>
	///
	/// Sets whether the band's background bitmap begins at the control's upper-left corner instead of
	/// the band's upper-left corner. If set to \c VARIANT_TRUE, the bitmap's origin is the control's
	/// upper-left corner, so that the background image doesn't move if the band is resized or moved (instead
	/// the visible part of the background image changes). If set to \c VARIANT_FALSE, the origin is the
	/// band's upper-left corner, so that the background image moves with the band.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_FixedBackgroundBitmapOrigin, put_hBackgroundBitmap
	virtual HRESULT STDMETHODCALLTYPE put_FixedBackgroundBitmapOrigin(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the band's text color. If set to -1, the control's text color is used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa put_ForeColor, get_BackColor, ReBar::get_ForeColor
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Sets the \c ForeColor property</em>
	///
	/// Sets the band's text color. If set to -1, the control's text color is used.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.
	///
	/// \sa get_ForeColor, put_BackColor, ReBar::put_ForeColor
	virtual HRESULT STDMETHODCALLTYPE put_ForeColor(OLE_COLOR newValue);
	/// \brief <em>Retrieves the current setting of the \c hBackgroundBitmap property</em>
	///
	/// Retrieves the handle of the bitmap used as the background for the band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hBackgroundBitmap, get_BackColor, ReBar::get_hPalette, get_FixedBackgroundBitmapOrigin
	virtual HRESULT STDMETHODCALLTYPE get_hBackgroundBitmap(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hBackgroundBitmap property</em>
	///
	/// Sets the handle of the bitmap used as the background for the band.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_hBackgroundBitmap, put_BackColor, ReBar::put_hPalette, put_FixedBackgroundBitmapOrigin
	virtual HRESULT STDMETHODCALLTYPE put_hBackgroundBitmap(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c hContainedWindow property</em>
	///
	/// Retrieves the handle of the window contained in the band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_hContainedWindow, get_MinimumHeight, get_CurrentHeight, get_MaximumHeight, get_MinimumWidth,
	///     get_CurrentWidth
	virtual HRESULT STDMETHODCALLTYPE get_hContainedWindow(OLE_HANDLE* pValue);
	/// \brief <em>Sets the \c hContainedWindow property</em>
	///
	/// Sets the handle of the window contained in the band.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \attention Setting this property to zero hides the previously contained window. However, COM might
	///            still treat the window as visible, so it's best to explicitly set the contained control's
	///            \c Visible property to \c False (if it is a COM control) after setting the
	///            \c hContainedWindow property to zero. To make the control visible after setting the
	///            \c hContainedWindow property to zero set the control's \c Visible property to \c False
	///            and then to \c True.
	///
	/// \sa get_hContainedWindow, put_MinimumHeight, put_CurrentHeight, put_MaximumHeight, put_MinimumWidth,
	///     put_CurrentWidth
	virtual HRESULT STDMETHODCALLTYPE put_hContainedWindow(OLE_HANDLE newValue);
	/// \brief <em>Retrieves the current setting of the \c HeightChangeStepSize property</em>
	///
	/// Retrieves the number of pixels by which the band's height can be changed in a single step. If the
	/// control changes the band's height, the difference between old and new height will be a multiple of
	/// this value.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c VariableHeight property is set to \c VARIANT_FALSE, this property is ignored.
	///
	/// \sa put_HeightChangeStepSize, get_MinimumHeight, get_CurrentHeight, get_MaximumHeight,
	///     get_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE get_HeightChangeStepSize(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c HeightChangeStepSize property</em>
	///
	/// Sets the number of pixels by which the band's height can be changed in a single step. If the
	/// control changes the band's height, the difference between old and new height will be a multiple of
	/// this value.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c VariableHeight property is set to \c VARIANT_FALSE, this property is ignored.
	///
	/// \sa get_HeightChangeStepSize, put_MinimumHeight, put_CurrentHeight, put_MaximumHeight,
	///     put_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE put_HeightChangeStepSize(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c HideIfVertical property</em>
	///
	/// Retrieves whether the band is visible if the control's orientation is vertical. If set to
	/// \c VARIANT_TRUE, the band is displayed if the control's orientation is vertical; otherwise it is
	/// displayed on horizontal orientation only.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_HideIfVertical, get_Visible, ReBar::get_Orientation
	virtual HRESULT STDMETHODCALLTYPE get_HideIfVertical(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c HideIfVertical property</em>
	///
	/// Sets whether the band is visible if the control's orientation is vertical. If set to
	/// \c VARIANT_TRUE, the band is displayed if the control's orientation is vertical; otherwise it is
	/// displayed on horizontal orientation only.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_HideIfVertical, put_Visible, ReBar::put_Orientation
	virtual HRESULT STDMETHODCALLTYPE put_HideIfVertical(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the band's icon in the control's \c ilBands image list. If set to
	/// -2, no icon is displayed for this band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_IconIndex, ReBar::get_hImageList, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa put_IconIndex, ReBar::get_hImageList, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
	/// \brief <em>Sets the \c IconIndex property</em>
	///
	/// Sets the zero-based index of the band's icon in the control's \c ilBands image list. If set to
	/// -2, no icon is displayed for this band.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_IconIndex, ReBar::put_hImageList, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa get_IconIndex, ReBar::put_hImageList, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_IconIndex(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c ID property</em>
	///
	/// Retrieves an unique ID identifying this band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks A band's ID will never change.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_Index, TBarCtlsLibU::BandIdentifierTypeConstants
	/// \else
	///   \sa get_Index, TBarCtlsLibA::BandIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_ID(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c IdealWidth property</em>
	///
	/// Retrieves the width (in pixels), to which the control tries to resize the band, if it auto-sizes it.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_IdealWidth, get_MinimumWidth, get_CurrentWidth, Maximize
	virtual HRESULT STDMETHODCALLTYPE get_IdealWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c IdealWidth property</em>
	///
	/// Sets the width (in pixels), to which the control tries to resize the band, if it auto-sizes it.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_IdealWidth, put_MinimumWidth, put_CurrentWidth, Maximize
	virtual HRESULT STDMETHODCALLTYPE put_IdealWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index identifying this band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing bands changes other bands' indexes, the index is the best
	///          (and fastest) option to identify a band.
	///
	/// \if UNICODE
	///   \sa put_Index, get_ID, TBarCtlsLibU::BandIdentifierTypeConstants
	/// \else
	///   \sa put_Index, get_ID, TBarCtlsLibA::BandIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Sets the \c Index property</em>
	///
	/// Sets the zero-based index identifying this band. Setting this property moves the band.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing bands changes other bands' indexes, the index is the best
	///          (and fastest) option to identify a band.
	///
	/// \if UNICODE
	///   \sa get_Index, get_ID, TBarCtlsLibU::BandIdentifierTypeConstants
	/// \else
	///   \sa get_Index, get_ID, TBarCtlsLibA::BandIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_Index(LONG newValue);
	/// \brief <em>Retrieves the current setting of the \c KeepInFirstRow property</em>
	///
	/// Retrieves whether the band is always displayed within the first row. If set to \c VARIANT_TRUE, the
	/// band is locked to the first row; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_KeepInFirstRow, get_NewLine, ReBar::get_NumberOfRows
	virtual HRESULT STDMETHODCALLTYPE get_KeepInFirstRow(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c KeepInFirstRow property</em>
	///
	/// Sets whether the band is always displayed within the first row. If set to \c VARIANT_TRUE, the
	/// band is locked to the first row; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_KeepInFirstRow, put_NewLine
	virtual HRESULT STDMETHODCALLTYPE put_KeepInFirstRow(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c MaximumHeight property</em>
	///
	/// Retrieves the band's maximum height in pixels. The band won't be sized larger than that.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c VariableHeight property is set to \c VARIANT_FALSE, this property is ignored.
	///
	/// \sa put_MaximumHeight, get_MinimumHeight, get_CurrentHeight, get_HeightChangeStepSize,
	///     get_MinimumWidth, get_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE get_MaximumHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c MaximumHeight property</em>
	///
	/// Sets the band's maximum height in pixels. The band won't be sized larger than that.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c VariableHeight property is set to \c VARIANT_FALSE, this property is ignored.
	///
	/// \sa get_MaximumHeight, put_MinimumHeight, put_CurrentHeight, put_HeightChangeStepSize,
	///     put_MinimumWidth, put_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE put_MaximumHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c MinimumHeight property</em>
	///
	/// Retrieves the band's minimum height in pixels. The band won't be sized smaller than that.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MinimumHeight, get_MaximumHeight, get_CurrentHeight, get_HeightChangeStepSize,
	///     get_MinimumWidth
	virtual HRESULT STDMETHODCALLTYPE get_MinimumHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c MinimumHeight property</em>
	///
	/// Sets the band's minimum height in pixels. The band won't be sized smaller than that.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MinimumHeight, put_MaximumHeight, put_CurrentHeight, put_HeightChangeStepSize,
	///     put_MinimumWidth
	virtual HRESULT STDMETHODCALLTYPE put_MinimumHeight(OLE_YSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c MinimumWidth property</em>
	///
	/// Retrieves the band's minimum width in pixels. The band won't be sized smaller than that.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_MinimumWidth, get_CurrentWidth, get_MinimumHeight, get_MaximumHeight
	virtual HRESULT STDMETHODCALLTYPE get_MinimumWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c MinimumWidth property</em>
	///
	/// Sets the band's minimum width in pixels. The band won't be sized smaller than that.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_MinimumWidth, put_CurrentWidth, put_MinimumHeight, put_MaximumHeight
	virtual HRESULT STDMETHODCALLTYPE put_MinimumWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c NewLine property</em>
	///
	/// Retrieves whether the band is displayed in a new row within the control. If set to \c VARIANT_TRUE,
	/// the band is displayed on a new line; otherwise it is displayed in the same row as the previous band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_NewLine, get_KeepInFirstRow, get_IdealWidth, get_CurrentWidth, ReBar::get_NumberOfRows,
	///     ReBar::Raise_AutoBreakingBand
	virtual HRESULT STDMETHODCALLTYPE get_NewLine(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c NewLine property</em>
	///
	/// Sets whether the band is displayed in a new row within the control. If set to \c VARIANT_TRUE,
	/// the band is displayed on a new line; otherwise it is displayed in the same row as the previous band.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_NewLine, put_KeepInFirstRow, put_IdealWidth, put_CurrentWidth, ReBar::Raise_AutoBreakingBand
	virtual HRESULT STDMETHODCALLTYPE put_NewLine(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Resizable property</em>
	///
	/// Retrieves whether the band can be resized by the user. If set to \c VARIANT_TRUE, a sizing grip is
	/// displayed with which the band can be resized; otherwise no such grip is displayed and the band has a
	/// fixed size.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Resizable, get_SizingGripVisibility, get_IdealWidth, get_CurrentWidth, get_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE get_Resizable(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Resizable property</em>
	///
	/// Sets whether the band can be resized by the user. If set to \c VARIANT_TRUE, a sizing grip is
	/// displayed with which the band can be resized; otherwise no such grip is displayed and the band has a
	/// fixed size.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Resizable, put_SizingGripVisibility, put_IdealWidth, put_CurrentWidth, put_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE put_Resizable(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c RowHeight property</em>
	///
	/// Retrieves the height in pixels of the row that contains this band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa GetRectangle, ReBar::get_NumberOfRows
	virtual HRESULT STDMETHODCALLTYPE get_RowHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c ShowTitle property</em>
	///
	/// Retrieves whether the band's caption, which may consist of text and an icon, is displayed. If set to
	/// \c VARIANT_TRUE, the band's caption is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_ShowTitle, get_Text, get_IconIndex, get_TitleWidth
	virtual HRESULT STDMETHODCALLTYPE get_ShowTitle(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c ShowTitle property</em>
	///
	/// Sets whether the band's caption, which may consist of text and an icon, is displayed. If set to
	/// \c VARIANT_TRUE, the band's caption is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_ShowTitle, put_Text, put_IconIndex, put_TitleWidth
	virtual HRESULT STDMETHODCALLTYPE put_ShowTitle(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c SizingGripVisibility property</em>
	///
	/// Retrieves when the band's sizing grip is displayed. Any of the values defined by the
	/// \c SizingGripVisibilityConstants enumeration is valid.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa put_SizingGripVisibility, get_Resizable, ReBar::get_HighlightColor, ReBar::get_ShadowColor,
	///       TBarCtlsLibU::SizingGripVisibilityConstants
	/// \else
	///   \sa put_SizingGripVisibility, get_Resizable, ReBar::get_HighlightColor, ReBar::get_ShadowColor,
	///       TBarCtlsLibA::SizingGripVisibilityConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SizingGripVisibility(SizingGripVisibilityConstants* pValue);
	/// \brief <em>Sets the \c SizingGripVisibility property</em>
	///
	/// Sets when the band's sizing grip is displayed. Any of the values defined by the
	/// \c SizingGripVisibilityConstants enumeration is valid.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \if UNICODE
	///   \sa get_SizingGripVisibility, put_Resizable, ReBar::put_HighlightColor, ReBar::put_ShadowColor,
	///       TBarCtlsLibU::SizingGripVisibilityConstants
	/// \else
	///   \sa get_SizingGripVisibility, put_Resizable, ReBar::put_HighlightColor, ReBar::put_ShadowColor,
	///       TBarCtlsLibA::SizingGripVisibilityConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE put_SizingGripVisibility(SizingGripVisibilityConstants newValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the band's text. The maximum number of characters in this text is specified by
	/// \c MAX_BANDTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Text, ReBarBands::Add, MAX_BANDTEXTLENGTH
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Sets the \c Text property</em>
	///
	/// Sets the band's text. The maximum number of characters in this text is specified by
	/// \c MAX_BANDTEXTLENGTH.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Text, ReBarBands::Add, MAX_BANDTEXTLENGTH
	virtual HRESULT STDMETHODCALLTYPE put_Text(BSTR newValue);
	/// \brief <em>Retrieves the current setting of the \c TitleWidth property</em>
	///
	/// Retrieves the width (in pixels) of the band's caption, which may consist of text and an icon. If set
	/// to -1, the caption is sized automatically.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_TitleWidth, get_ShowTitle, get_Text, get_IconIndex, get_CurrentWidth
	virtual HRESULT STDMETHODCALLTYPE get_TitleWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Sets the \c TitleWidth property</em>
	///
	/// Sets the width (in pixels) of the band's caption, which may consist of text and an icon. If set
	/// to -1, the caption is sized automatically.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_TitleWidth, put_ShowTitle, put_Text, put_IconIndex, put_CurrentWidth
	virtual HRESULT STDMETHODCALLTYPE put_TitleWidth(OLE_XSIZE_PIXELS newValue);
	/// \brief <em>Retrieves the current setting of the \c UseChevron property</em>
	///
	/// Retrieves whether a chevron button is displayed if the band's current width is smaller than its ideal
	/// width. If set to \c VARIANT_TRUE, a chevron button is displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_UseChevron, get_IdealWidth, get_CurrentWidth, get_ChevronVisible,
	///     get_ChevronButtonObjectState, GetChevronRectangle, ClickChevron
	virtual HRESULT STDMETHODCALLTYPE get_UseChevron(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c UseChevron property</em>
	///
	/// Sets whether a chevron button is displayed if the band's current width is smaller than its ideal
	/// width. If set to \c VARIANT_TRUE, a chevron button is displayed; otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseChevron, put_IdealWidth, put_CurrentWidth, get_ChevronVisible,
	///     get_ChevronButtonObjectState, GetChevronRectangle, ClickChevron
	virtual HRESULT STDMETHODCALLTYPE put_UseChevron(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c VariableHeight property</em>
	///
	/// Retrieves whether the band's height can be changed by the control. If set to \c VARIANT_TRUE, the
	/// band's height can be changed according to the settings specified by the \c MinimumHeight,
	/// \c MaximumHeight and \c HeightChangeStepSize properties; otherwise its height is fixed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_VariableHeight, get_MinimumHeight, get_MaximumHeight, get_HeightChangeStepSize, get_Resizable
	virtual HRESULT STDMETHODCALLTYPE get_VariableHeight(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c VariableHeight property</em>
	///
	/// Sets whether the band's height can be changed by the control. If set to \c VARIANT_TRUE, the
	/// band's height can be changed according to the settings specified by the \c MinimumHeight,
	/// \c MaximumHeight and \c HeightChangeStepSize properties; otherwise its height is fixed.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_VariableHeight, put_MinimumHeight, put_MaximumHeight, put_HeightChangeStepSize, put_Resizable
	virtual HRESULT STDMETHODCALLTYPE put_VariableHeight(VARIANT_BOOL newValue);
	/// \brief <em>Retrieves the current setting of the \c Visible property</em>
	///
	/// Retrieves whether the band is visible. If set to \c VARIANT_TRUE, the band will be displayed;
	/// otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa put_Visible, Hide, Show, get_ChevronVisible, ReBarBands::Remove, ReBarBands::Add
	virtual HRESULT STDMETHODCALLTYPE get_Visible(VARIANT_BOOL* pValue);
	/// \brief <em>Sets the \c Visible property</em>
	///
	/// Sets whether the band is visible. If set to \c VARIANT_TRUE, the band will be displayed;
	/// otherwise not.
	///
	/// \param[in] newValue The setting to apply.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_Visible, Hide, Show, get_ChevronVisible, ReBarBands::Remove, ReBarBands::Add
	virtual HRESULT STDMETHODCALLTYPE put_Visible(VARIANT_BOOL newValue);

	/// \brief <em>Clicks the band's chevron programmatically</em>
	///
	/// Clicks the band's chevron programmatically. The \c ChevronClick event will be raised with the
	/// \c userData parameter being set to the specified value.
	///
	/// \param[in] userData A \c Long value that will be passed to the \c ChevronClick event handler.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa get_UseChevron, ReBar::Raise_ChevronClick
	virtual HRESULT STDMETHODCALLTYPE ClickChevron(LONG userData = 0);
	/// \brief <em>Retrieves the size of the border around the band</em>
	///
	/// Retrieves the width (in pixels) of the border around the band.
	///
	/// \param[out] pLeftBorder The width (in pixels) of the border on the left of the band.
	/// \param[out] pTopBorder The height (in pixels) of the border on the top of the band.
	/// \param[out] pRightBorder The width (in pixels) of the border on the right of the band.
	/// \param[out] pBottomBorder The height (in pixels) of the border on the bottom of the band.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa GetRectangle, ReBar::get_DisplayBandSeparators, ReBar::GetMargins
	virtual HRESULT STDMETHODCALLTYPE GetBorderSizes(OLE_XSIZE_PIXELS* pLeftBorder = NULL, OLE_YSIZE_PIXELS* pTopBorder = NULL, OLE_XSIZE_PIXELS* pRightBorder = NULL, OLE_YSIZE_PIXELS* pBottomBorder = NULL);
	/// \brief <em>Retrieves the bounding rectangle of the band's chevron button</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's client area) of the band's
	/// chevron button.
	///
	/// \param[out] pXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
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
	/// \sa get_UseChevron, GetRectangle
	virtual HRESULT STDMETHODCALLTYPE GetChevronRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Retrieves the bounding rectangle of either the band or a part of it</em>
	///
	/// Retrieves the bounding rectangle (in pixels relative to the control's client area) of either the
	/// band or a part of it.
	///
	/// \param[in] rectangleType The rectangle to retrieve. Any of the values defined by the
	///            \c BandRectangleTypeConstants enumeration is valid.
	/// \param[out] pXLeft The x-coordinate (in pixels) of the bounding rectangle's left border
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
	///   \sa GetBorderSizes, GetChevronRectangle, ReBar::GetMargins,
	///       TBarCtlsLibU::BandRectangleTypeConstants
	/// \else
	///   \sa GetBorderSizes, GetChevronRectangle, ReBar::GetMargins,
	///       TBarCtlsLibA::BandRectangleTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE GetRectangle(BandRectangleTypeConstants rectangleType, OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	/// \brief <em>Hides the band</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Show, put_Visible
	virtual HRESULT STDMETHODCALLTYPE Hide(void);
	/// \brief <em>Maximizes the band's width</em>
	///
	/// Maximizes the band to either its ideal width as specified by the \c IdealWidth property, or to the
	/// largest possible width.
	///
	/// \param[in] useIdealWidth If set to \c VARIANT_TRUE, the band is maximized to its ideal width, as
	///            specified by the \c IdealWidth property; otherwise it is made as large as possible.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Minimize, get_IdealWidth, get_CurrentWidth
	virtual HRESULT STDMETHODCALLTYPE Maximize(VARIANT_BOOL useIdealWidth = VARIANT_TRUE);
	/// \brief <em>Minimizes the band's width</em>
	///
	/// Minimizes the band to its minimum width as specified by the \c MinimumWidth property.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Maximize, get_MinimumWidth, get_IdealWidth, get_CurrentWidth
	virtual HRESULT STDMETHODCALLTYPE Minimize(void);
	/// \brief <em>Unhides the band</em>
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa Hide, put_Visible
	virtual HRESULT STDMETHODCALLTYPE Show(void);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Attaches this object to a given band</em>
	///
	/// Attaches this object to a given band, so that the band's properties can be retrieved and set
	/// using this object's methods.
	///
	/// \param[in] bandIndex The band to attach to.
	///
	/// \sa Detach
	void Attach(int bandIndex);
	/// \brief <em>Detaches this object from a band</em>
	///
	/// Detaches this object from the band it currently wraps, so that it doesn't wrap any band anymore.
	///
	/// \sa Attach
	void Detach(void);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// Applies the settings from a given source to the band wrapped by this object.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(LPREBARBANDINFO pSource);
	/// \brief <em>Sets this object's properties to given values</em>
	///
	/// \overload
	HRESULT LoadState(VirtualReBarBand* pSource);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndRBar The rebar window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(LPREBARBANDINFO pTarget, HWND hWndRBar = NULL);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \overload
	HRESULT SaveState(VirtualReBarBand* pTarget);
	/// \brief <em>Sets the owner of this band</em>
	///
	/// \param[in] pOwner The owner to set.
	///
	/// \sa Properties::pOwnerRBar
	void SetOwner(__in_opt ReBar* pOwner);

protected:
	/// \brief <em>Holds the object's properties' settings</em>
	struct Properties
	{
		/// \brief <em>The \c ReBar object that owns this band</em>
		///
		/// \sa SetOwner
		ReBar* pOwnerRBar;
		/// \brief <em>The index of the band wrapped by this object</em>
		int bandIndex;

		Properties()
		{
			pOwnerRBar = NULL;
			bandIndex = -1;
		}

		~Properties();

		/// \brief <em>Retrieves the owning rebar's window handle</em>
		///
		/// \return The window handle of the rebar that contains this band.
		///
		/// \sa pOwnerRBar
		HWND GetRBarHWnd(void);
	} properties;

	/// \brief <em>Writes a given object's settings to a given target</em>
	///
	/// \param[in] bandIndex The band for which to save the settings.
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[in] hWndRBar The rebar window the method will work on.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(int bandIndex, LPREBARBANDINFO pTarget, HWND hWndRBar = NULL);
};     // ReBarBand

OBJECT_ENTRY_AUTO(__uuidof(ReBarBand), ReBarBand)