//////////////////////////////////////////////////////////////////////
/// \file UndocComctl32.h
/// \author Timo "TimoSoft" Kunze
/// \brief <em>Undocumented comctl32.dll stuff</em>
///
/// Declaration of some comctl32.dll internals that we're using.
///
/// \todo Improve documentation ("See also" sections).
///
/// \sa ReBar
//////////////////////////////////////////////////////////////////////

#pragma once


/// \brief <em>Displays a splitter at the bottom of the rebar control</em>
///
/// This is an undocumented extended rebar window style that causes the rebar to display a splitter at the
/// bottom of the control. If the mouse is moved over this splitter, the cursor is changed to a sizing
/// cursor. Dragging the splitter doesn't resize the control, but triggers the \c RBN_SPLITTERDRAG
/// notification.
///
/// \remarks Requires comctl32.dll version 6.10 or higher.
#define RBS_EX_SPLITTER 0x02
#ifndef TBSTYLE_EX_MULTICOLUMN
	/// \brief <em>Organizes the buttons of a vertical tool bar in multiple columns</em>
	///
	/// This is an extended tool bar window style that causes the tool bar to organize the buttons in multiple
	/// columns. The \c TBSTYLE_EX_VERTICAL style has to be set as well.\n
	/// The application has to send the \c TB_SETBOUNDINGSIZE message in order to make the style work.
	///
	/// \sa TBSTYLE_EX_VERTICAL, TB_SETBOUNDINGSIZE
	#define TBSTYLE_EX_MULTICOLUMN 0x02
#endif
#ifndef TBSTYLE_EX_VERTICAL
	/// \brief <em>Activates vertical orientation of the tool bar</em>
	///
	/// This is an extended tool bar window style that makes the tool bar a vertical tool bar.
	///
	/// \sa TBSTYLE_EX_MULTICOLUMN
	#define TBSTYLE_EX_VERTICAL 0x04
#endif
#define CCM_TRANSLATEACCELERATOR (CCM_FIRST + 0xA)
/// \brief <em>Sets the bounding size for a multi-column vertical tool bar</em>
///
/// This is a tool bar message that is used with the \c TBSTYLE_EX_MULTICOLUMN style. It is used to specify
/// the bounding rectangle in which the buttons will be organized.\n
/// The \c wParam parameter has to be set to 0, the \c lParam parameter has to be set to the address of a
/// \c SIZE structure specifying the size of the bounding rectangle.
///
/// \sa TBSTYLE_EX_MULTICOLUMN,
///     <a href="https://msdn.microsoft.com/en-us/library/ee663584.aspx">TB_SETBOUNDINGSIZE</a>
#define TB_SETBOUNDINGSIZE (WM_USER + 93)
/// \brief <em>Sets the bounding size for a multi-column vertical tool bar</em>
///
/// This is a tool bar message that is used to retrieve the count of tool bar buttons that have the
/// specified accelerator character.\n
/// The \c wParam parameter has to be set to a \c WCHAR value representing the input accelerator character
/// to test, the \c lParam parameter has to be set to the address of an \c int value that receives the
/// number of buttons that have the accelerator character.
///
/// \sa <a href="https://msdn.microsoft.com/en-us/library/ee663572.aspx">TB_HASACCELERATOR</a>
#define TB_HASACCELERATOR (WM_USER + 95)
/// \brief <em>Sets the bounding size for a multi-column vertical tool bar</em>
///
/// This is an undocumented tool bar message that is used to specify the distance between a button's
/// caption and the button's drop-down portion.\n
/// The \c wParam parameter has to be set to the distance in pixels, the \c lParam parameter has to be set
/// to 0. The distance is measured from the right edge of the caption to the separator between the button
/// portion and the drop-down portion.
#define TB_SETDROPDOWNGAP (WM_USER + 100)

typedef struct tagNMTBCUSTOMIZEDLG
{
	NMHDR hdr;
	HWND hDlg;
} NMTBCUSTOMIZEDLG, *LPNMTBCUSTOMIZEDLG;

typedef struct tagNMTBDUPACCELERATOR
{
	NMHDR hdr;
	UINT ch;
	BOOL fDup;
} NMTBDUPACCELERATOR, *LPNMTBDUPACCELERATOR;

typedef struct tagNMTBWRAPACCELERATOR
{
	NMHDR hdr;
	UINT ch;
	int iButton;
} NMTBWRAPACCELERATOR, *LPNMTBWRAPACCELERATOR;

typedef struct tagNMTBWRAPHOTITEM
{
	NMHDR hdr;
	int iStart;
	int iDir;
	UINT nReason;
} NMTBWRAPHOTITEM, *LPNMTBWRAPHOTITEM;