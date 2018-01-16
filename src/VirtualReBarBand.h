//////////////////////////////////////////////////////////////////////
/// \class VirtualReBarBand
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Wraps a not existing rebar band</em>
///
/// This class is a wrapper around a rebar band that does not yet or not anymore exist in the control.
///
/// \if UNICODE
///   \sa TBarCtlsLibU::IVirtualReBarBand, ReBarBand, ReBar
/// \else
///   \sa TBarCtlsLibA::IVirtualReBarBand, ReBarBand, ReBar
/// \endif
//////////////////////////////////////////////////////////////////////


#pragma once

#include "res/resource.h"
#ifdef UNICODE
	#include "TBarCtlsU.h"
#else
	#include "TBarCtlsA.h"
#endif
#include "_IVirtualReBarBandEvents_CP.h"
#include "helpers.h"
#include "ReBar.h"


class ATL_NO_VTABLE VirtualReBarBand : 
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<VirtualReBarBand, &CLSID_VirtualReBarBand>,
    public ISupportErrorInfo,
    public IConnectionPointContainerImpl<VirtualReBarBand>,
    public Proxy_IVirtualReBarBandEvents< VirtualReBarBand>,
    #ifdef UNICODE
    	public IDispatchImpl<IVirtualReBarBand, &IID_IVirtualReBarBand, &LIBID_TBarCtlsLibU, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #else
    	public IDispatchImpl<IVirtualReBarBand, &IID_IVirtualReBarBand, &LIBID_TBarCtlsLibA, /*wMajor =*/ VERSION_MAJOR, /*wMinor =*/ VERSION_MINOR>
    #endif
{
	friend class ReBar;
	friend class ReBarBand;

public:
	#ifndef DOXYGEN_SHOULD_SKIP_THIS
		DECLARE_REGISTRY_RESOURCEID(IDR_VIRTUALREBARBAND)

		BEGIN_COM_MAP(VirtualReBarBand)
			COM_INTERFACE_ENTRY(IVirtualReBarBand)
			COM_INTERFACE_ENTRY(IDispatch)
			COM_INTERFACE_ENTRY(ISupportErrorInfo)
			COM_INTERFACE_ENTRY(IConnectionPointContainer)
		END_COM_MAP()

		BEGIN_CONNECTION_POINT_MAP(VirtualReBarBand)
			CONNECTION_POINT_ENTRY(__uuidof(_IVirtualReBarBandEvents))
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
	/// \name Implementation of IVirtualReBarBand
	///
	//@{
	/// \brief <em>Retrieves the current setting of the \c AddMarginsAroundChild property</em>
	///
	/// Retrieves whether the band's height will be or was increased by a small margin. If set to
	/// \c VARIANT_TRUE, a small margin will be or was inserted at the top and bottom of the band's
	/// contained window; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_CurrentHeight, ReBar::get_DisplayBandSeparators
	virtual HRESULT STDMETHODCALLTYPE get_AddMarginsAroundChild(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c BackColor property</em>
	///
	/// Retrieves the band's background color. If set to -1, the control's background color will be or was
	/// used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.\n
	///          This property is read-only.
	///
	/// \sa get_hBackgroundBitmap, get_ForeColor, ReBar::get_BackColor
	virtual HRESULT STDMETHODCALLTYPE get_BackColor(OLE_COLOR* pValue);
	/// \brief <em>Retrieves the current setting of the \c BandData property</em>
	///
	/// Retrieves the \c LONG value that will be or was associated with the band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ReBar::Raise_FreeBandData
	virtual HRESULT STDMETHODCALLTYPE get_BandData(LONG* pValue);
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
	/// \brief <em>Retrieves the current setting of the \c CurrentHeight property</em>
	///
	/// Retrieves the band's current height in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c VariableHeight property is set to \c VARIANT_FALSE, this property is ignored.\n
	///          This property is read-only.
	///
	/// \sa get_CurrentWidth, get_MinimumHeight, get_MaximumHeight, get_HeightChangeStepSize,
	///     get_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE get_CurrentHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c CurrentWidth property</em>
	///
	/// Retrieves the band's current width in pixels.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_CurrentHeight, get_MinimumWidth, get_IdealWidth, get_TitleWidth
	virtual HRESULT STDMETHODCALLTYPE get_CurrentWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c FixedBackgroundBitmapOrigin property</em>
	///
	/// Retrieves whether the band's background bitmap will begin or has begun at the control's upper-left
	/// corner instead of the band's upper-left corner. If set to \c VARIANT_TRUE, the bitmap's origin will
	/// be or was the control's upper-left corner, so that the background image will or did not move if the
	/// band will be or was resized or moved (instead the visible part of the background image will change or
	/// has changed). If set to \c VARIANT_FALSE, the origin will be or was the band's upper-left corner, so
	/// that the background image will move or has moved with the band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_hBackgroundBitmap
	virtual HRESULT STDMETHODCALLTYPE get_FixedBackgroundBitmapOrigin(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ForeColor property</em>
	///
	/// Retrieves the band's text color. If set to -1, the control's text color will be or was used.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property isn't supported for themed rebars.\n
	///          This property is read-only.
	///
	/// \sa get_BackColor, ReBar::get_ForeColor
	virtual HRESULT STDMETHODCALLTYPE get_ForeColor(OLE_COLOR* pValue);
	/// \brief <em>Retrieves the current setting of the \c hBackgroundBitmap property</em>
	///
	/// Retrieves the handle of the bitmap that will be or was used as the background for the band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_BackColor, ReBar::get_hPalette, get_FixedBackgroundBitmapOrigin
	virtual HRESULT STDMETHODCALLTYPE get_hBackgroundBitmap(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c hContainedWindow property</em>
	///
	/// Retrieves the handle of the window that will be or was contained in the band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MinimumHeight, get_CurrentHeight, get_MaximumHeight, get_MinimumWidth, get_CurrentWidth
	virtual HRESULT STDMETHODCALLTYPE get_hContainedWindow(OLE_HANDLE* pValue);
	/// \brief <em>Retrieves the current setting of the \c HeightChangeStepSize property</em>
	///
	/// Retrieves the number of pixels by which the band's height can or could be changed in a single step.
	/// If the control changes or has changed the band's height, the difference between old and new height
	/// will be or was a multiple of this value.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c VariableHeight property is set to \c VARIANT_FALSE, this property is ignored.\n
	///          This property is read-only.
	///
	/// \sa get_MinimumHeight, get_CurrentHeight, get_MaximumHeight, get_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE get_HeightChangeStepSize(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c HideIfVertical property</em>
	///
	/// Retrieves whether the band will be or was visible if the control's orientation is vertical. If set to
	/// \c VARIANT_TRUE, the band will be or was displayed if the control's orientation is vertical;
	/// otherwise it will be or was displayed on horizontal orientation only.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Visible, ReBar::get_Orientation
	virtual HRESULT STDMETHODCALLTYPE get_HideIfVertical(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c IconIndex property</em>
	///
	/// Retrieves the zero-based index of the band's icon in the control's \c ilBands image list. If set to
	/// -2, no icon will be or was displayed for this band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \if UNICODE
	///   \sa ReBar::get_hImageList, TBarCtlsLibU::ImageListConstants
	/// \else
	///   \sa ReBar::get_hImageList, TBarCtlsLibA::ImageListConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_IconIndex(LONG* pValue);
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
	/// Retrieves the width (in pixels), to which the control will try or has tried to resize the band, if it
	/// will auto-size or has auto-sized it.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MinimumWidth, get_CurrentWidth
	virtual HRESULT STDMETHODCALLTYPE get_IdealWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c Index property</em>
	///
	/// Retrieves the zero-based index identifying this band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks Although adding or removing bands changes other bands' indexes, the index is the best
	///          (and fastest) option to identify a band.\n
	///          This property is read-only.
	///
	/// \if UNICODE
	///   \sa get_ID, TBarCtlsLibU::BandIdentifierTypeConstants
	/// \else
	///   \sa get_ID, TBarCtlsLibA::BandIdentifierTypeConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_Index(LONG* pValue);
	/// \brief <em>Retrieves the current setting of the \c KeepInFirstRow property</em>
	///
	/// Retrieves whether the band will be or was always displayed within the first row. If set to
	/// \c VARIANT_TRUE, the band will be or was locked to the first row; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_NewLine, ReBar::get_NumberOfRows
	virtual HRESULT STDMETHODCALLTYPE get_KeepInFirstRow(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c MaximumHeight property</em>
	///
	/// Retrieves the band's maximum height in pixels. The band won't be or wasn't sized larger than that.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks If the \c VariableHeight property is set to \c VARIANT_FALSE, this property is ignored.\n
	///          This property is read-only.
	///
	/// \sa get_MinimumHeight, get_CurrentHeight, get_HeightChangeStepSize, get_MinimumWidth,
	///     get_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE get_MaximumHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c MinimumHeight property</em>
	///
	/// Retrieves the band's minimum height in pixels. The band won't be or wasn't sized smaller than that.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MaximumHeight, get_CurrentHeight, get_HeightChangeStepSize, get_MinimumWidth
	virtual HRESULT STDMETHODCALLTYPE get_MinimumHeight(OLE_YSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c MinimumWidth property</em>
	///
	/// Retrieves the band's minimum width in pixels. The band won't be or wasn't sized smaller than that.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_CurrentWidth, get_MinimumHeight, get_MaximumHeight
	virtual HRESULT STDMETHODCALLTYPE get_MinimumWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c NewLine property</em>
	///
	/// Retrieves whether the band will be or was displayed in a new row within the control. If set to
	/// \c VARIANT_TRUE, the band will be or was displayed on a new line; otherwise it will be or was
	/// displayed in the same row as the previous band.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_KeepInFirstRow, get_IdealWidth, get_CurrentWidth, ReBar::get_NumberOfRows,
	///     ReBar::Raise_AutoBreakingBand
	virtual HRESULT STDMETHODCALLTYPE get_NewLine(VARIANT_BOOL* pValue);
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
	/// \remarks This property is read-only.
	///
	/// \sa get_SizingGripVisibility, get_IdealWidth, get_CurrentWidth, get_VariableHeight
	virtual HRESULT STDMETHODCALLTYPE get_Resizable(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c ShowTitle property</em>
	///
	/// Retrieves whether the band's caption, which may consist of text and an icon, will be or was
	/// displayed. If set to \c VARIANT_TRUE, the band's caption will be or was displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_Text, get_IconIndex, get_TitleWidth
	virtual HRESULT STDMETHODCALLTYPE get_ShowTitle(VARIANT_BOOL* pValue);
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
	///   \sa get_Resizable, ReBar::get_HighlightColor, ReBar::get_ShadowColor,
	///       TBarCtlsLibU::SizingGripVisibilityConstants
	/// \else
	///   \sa get_Resizable, ReBar::get_HighlightColor, ReBar::get_ShadowColor,
	///       TBarCtlsLibA::SizingGripVisibilityConstants
	/// \endif
	virtual HRESULT STDMETHODCALLTYPE get_SizingGripVisibility(SizingGripVisibilityConstants* pValue);
	/// \brief <em>Retrieves the current setting of the \c Text property</em>
	///
	/// Retrieves the band's text. The maximum number of characters in this text will be or was specified by
	/// \c MAX_BANDTEXTLENGTH.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ReBarBands::Add, MAX_BANDTEXTLENGTH
	virtual HRESULT STDMETHODCALLTYPE get_Text(BSTR* pValue);
	/// \brief <em>Retrieves the current setting of the \c TitleWidth property</em>
	///
	/// Retrieves the width (in pixels) of the band's caption, which may consist of text and an icon. If set
	/// to -1, the caption will be or was sized automatically.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_ShowTitle, get_Text, get_IconIndex, get_CurrentWidth
	virtual HRESULT STDMETHODCALLTYPE get_TitleWidth(OLE_XSIZE_PIXELS* pValue);
	/// \brief <em>Retrieves the current setting of the \c UseChevron property</em>
	///
	/// Retrieves whether a chevron button will be or was displayed if the band's current width is smaller
	/// than its ideal width. If set to \c VARIANT_TRUE, a chevron button will be or was displayed; otherwise
	/// not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_IdealWidth, get_CurrentWidth, get_ChevronButtonObjectState, GetChevronRectangle
	virtual HRESULT STDMETHODCALLTYPE get_UseChevron(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c VariableHeight property</em>
	///
	/// Retrieves whether the band's height can or could be changed by the control. If set to
	/// \c VARIANT_TRUE, the band's height can or could be changed according to the settings specified by the
	/// \c MinimumHeight, \c MaximumHeight and \c HeightChangeStepSize properties; otherwise its height will
	/// be or was fixed.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa get_MinimumHeight, get_MaximumHeight, get_HeightChangeStepSize, get_Resizable
	virtual HRESULT STDMETHODCALLTYPE get_VariableHeight(VARIANT_BOOL* pValue);
	/// \brief <em>Retrieves the current setting of the \c Visible property</em>
	///
	/// Retrieves whether the band will be or was visible. If set to \c VARIANT_TRUE, the band will be or was
	/// displayed; otherwise not.
	///
	/// \param[out] pValue The property's current setting.
	///
	/// \return An \c HRESULT error code.
	///
	/// \remarks This property is read-only.
	///
	/// \sa ReBarBands::Remove, ReBarBands::Add
	virtual HRESULT STDMETHODCALLTYPE get_Visible(VARIANT_BOOL* pValue);

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
	/// \remarks Requires comctl32.dll version 6.10 or higher.
	///
	/// \sa get_UseChevron
	virtual HRESULT STDMETHODCALLTYPE GetChevronRectangle(OLE_XPOS_PIXELS* pXLeft = NULL, OLE_YPOS_PIXELS* pYTop = NULL, OLE_XPOS_PIXELS* pXRight = NULL, OLE_YPOS_PIXELS* pYBottom = NULL);
	//@}
	//////////////////////////////////////////////////////////////////////

	/// \brief <em>Initializes this object with given data</em>
	///
	/// Initializes this object with the settings from a given source.
	///
	/// \param[in] pSource The data source from which to copy the settings.
	/// \param[in] bandIndex The band's zero-based index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa SaveState
	HRESULT LoadState(__in LPREBARBANDINFO pSource, int bandIndex);
	/// \brief <em>Writes this object's settings to a given target</em>
	///
	/// \param[in] pTarget The target to which to copy the settings.
	/// \param[out] bandIndex The band's zero-based index.
	///
	/// \return An \c HRESULT error code.
	///
	/// \sa LoadState
	HRESULT SaveState(__in LPREBARBANDINFO pTarget, int& bandIndex);
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
		/// \brief <em>A structure holding this band's settings</em>
		REBARBANDINFO settings;
		/// \brief <em>The band's zero-based index</em>
		int bandIndex;

		Properties()
		{
			pOwnerRBar = NULL;
			ZeroMemory(&settings, sizeof(settings));
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
};     // VirtualReBarBand

OBJECT_ENTRY_AUTO(__uuidof(VirtualReBarBand), VirtualReBarBand)