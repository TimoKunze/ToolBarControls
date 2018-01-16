VERSION 5.00
Object = "{5D0D0ABC-4898-4E46-AB48-291074A737A1}#1.3#0"; "TBarCtlsU.ocx"
Begin VB.UserControl UCBreadcrumb 
   BackColor       =   &H80000004&
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ScaleHeight     =   86
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   Begin TBarCtlsLibUCtl.ToolBar BreadcrumbBar 
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
      _cx             =   2355
      _cy             =   688
      AllowCustomization=   0   'False
      AlwaysDisplayButtonText=   0   'False
      AnchorHighlighting=   0   'False
      Appearance      =   0
      BackStyle       =   0
      BorderStyle     =   0
      ButtonHeight    =   28
      ButtonStyle     =   1
      ButtonTextPosition=   1
      ButtonWidth     =   22
      DisabledEvents  =   917608
      DisplayMenuDivider=   0   'False
      DisplayPartiallyClippedButtons=   -1  'True
      DontRedraw      =   0   'False
      DragClickTime   =   -1
      DragDropCustomizationModifierKey=   0
      DropDownGap     =   -1
      Enabled         =   -1  'True
      FirstButtonIndentation=   0
      FocusOnClick    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   -1
      HorizontalButtonPadding=   -1
      HorizontalButtonSpacing=   0
      HorizontalIconCaptionGap=   -1
      HorizontalTextAlignment=   1
      HoverTime       =   -1
      InsertMarkColor =   0
      MaximumButtonWidth=   0
      MaximumTextRows =   1
      MenuBarTheme    =   1
      MenuMode        =   0   'False
      MinimumButtonWidth=   0
      MousePointer    =   0
      MultiColumn     =   0   'False
      NormalDropDownButtonStyle=   0
      OLEDragImageStyle=   0
      Orientation     =   0
      ProcessContextMenuKeys=   -1  'True
      RaiseCustomDrawEventOnEraseBackground=   0   'False
      RegisterForOLEDragDrop=   1
      RightToLeft     =   0
      ShadowColor     =   -1
      ShowToolTips    =   0   'False
      SupportOLEDragImages=   -1  'True
      UseMnemonics    =   -1  'True
      UseSystemFont   =   -1  'True
      VerticalButtonPadding=   -1
      VerticalButtonSpacing=   0
      VerticalTextAlignment=   1
      WrapButtons     =   0   'False
   End
End
Attribute VB_Name = "UCBreadcrumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  ' This sample was heavily inspired by the "Explorer breadcrumbs in Vista" sample by
  ' Bjarke Viksoe which can be found on http://www.viksoe.dk/code/breadcrumbs.htm


  Implements ISubclassedWindow


  Private Const BTNID_FOLDERICON As Long = 1
  Private Const BTNID_FIRSTPART As Long = 2

  Private Const HWND_DESKTOP As Long = 0
  Private Const LF_FACESIZE As Long = 32

  Private Const HMARGIN_ICON As Long = 4
  Private Const HMARGIN_TEXT As Long = 4
  Private Const VMARGIN_ICON As Long = 4
  Private Const VMARGIN_TEXT As Long = 4


  Private Enum MenuItemTypeConstants
    mitShellItem
    mitOpenThisFolderItem
    mitEmptyItem
    mitSeparator
  End Enum


  Private Type DRAWITEMMETRICS
    rcSelection As RECT
    rcGutter As RECT
    rcIcon As RECT
    rcText As RECT
    rcSeparator As RECT
  End Type

  Private Type DRAWITEMSTRUCT
    ctlType As Long
    ctlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hDC As Long
    rcItem As RECT
    itemData As Long
  End Type

  Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(1 To LF_FACESIZE * 2) As Byte
  End Type

  Private Type MEASUREITEMSTRUCT
    ctlType As Long
    ctlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
  End Type

  Private Type MENUITEMDATA
    cbSize As Long
    itemType As MenuItemTypeConstants
    hasSubMenu As Boolean
    IParentISF As IVBShellFolder
    pIDLToParent As Long
    pIDLToDesktop As Long
  End Type

  Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As Long
    cch As Long
  End Type

  Private Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LOGFONT
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LOGFONT
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LOGFONT
    lfStatusFont As LOGFONT
    lfMessageFont As LOGFONT
    iPaddedBorderWidth As Long
  End Type

  Private Type TEXTMETRIC
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Integer
    tmLastChar As Integer
    tmDefaultChar As Integer
    tmBreakChar As Integer
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
  End Type

  Private Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
  End Type


  Private IID_IImageList As UUID
  Private IID_IShellIconOverlay As UUID
  Private IID_IShellItem As UUID

  Private m_bHasHiddenItems As Boolean
  Private m_currentPIDL As Long
  Private m_hImageList As Long
  Private m_hImageListDummy As Long
  Private m_hMenuFont As Long
  Private m_hTheme As Long
  Private m_hWndNotify As Long
  Private m_iconSize As Size
  Private m_IDesktop As IVBShellFolder
  Private m_nextCommandID As Long
  Private m_selectedMenuItemData As Long
  Private m_subclassed As Boolean
  Private m_textMetrics As TEXTMETRIC
  Private m_themedMenus As Boolean


  Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
  Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, ByVal pClipRect As Long) As Long
  Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal hDC As Long, prc As RECT) As Long
  Private Declare Function DrawThemeText Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal dwTextFlags2 As Long, pRect As RECT) As Long
  Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
  Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As Point) As Long
  Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, ByRef lpMenuItemInfo As MENUITEMINFO) As Long
  Private Declare Function GetProcessHeap Lib "kernel32.dll" () As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetTextMetrics Lib "gdi32.dll" Alias "GetTextMetricsW" (ByVal hDC As Long, lptm As TEXTMETRIC) As Long
  Private Declare Function GetThemeInt Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, piVal As Long) As Long
  Private Declare Function GetThemeMargins Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, ByVal prc As Long, pMargins As MARGINS) As Long
  Private Declare Function GetThemePartSize Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal prc As Long, ByVal eSize As Long, psz As Size) As Long
  Private Declare Function GetThemeTextExtent Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal pszText As Long, ByVal iCharCount As Long, ByVal dwTextFlags As Long, ByVal pBoundingRect As Long, pExtentRect As RECT) As Long
  Private Declare Function HeapAlloc Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
  Private Declare Function HeapFree Lib "kernel32.dll" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
  Private Declare Function ILAppendID Lib "shell32.dll" (ByVal pIDL As Long, ByVal pmkid As Long, ByVal fAppend As Long) As Long
  Private Declare Function ILGetNext Lib "shell32" Alias "#153" (ByVal pIDL As Long) As Long
  Private Declare Function ILIsEqual Lib "shell32.dll" (ByVal pIDL1 As Long, ByVal pIDL2 As Long) As Long
  Private Declare Function ILRemoveLastID Lib "shell32.dll" (ByVal pIDL As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
  Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, ByRef cx As Long, ByRef cy As Long) As Long
  Private Declare Function ImageList_LoadImage Lib "comctl32.dll" Alias "ImageList_LoadImageW" (ByVal hi As Long, ByVal lpbmp As Long, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function InflateRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
  Private Declare Function InvalidateRect Lib "user32.dll" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
  Private Declare Function IsAppThemed Lib "uxtheme.dll" () As Long
  Private Declare Function IsThemeBackgroundPartiallyTransparent Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long) As Long
  Private Declare Function IsThemePartDefined Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long) As Long
  Private Declare Function lstrcpyn Lib "kernel32.dll" Alias "lstrcpynW" (ByVal lpString1 As Long, ByVal lpString2 As Long, ByVal iMaxLength As Long) As Long
  Private Declare Function MapWindowPoints Lib "user32.dll" (ByVal hWndFrom As Long, ByVal hWndTo As Long, ByVal lpPoints As Long, ByVal cPoints As Long) As Long
  Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  Private Declare Function OffsetRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function PtInRect Lib "user32.dll" (ByRef lprc As RECT, ByVal x As Long, ByVal y As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal uPosition As Long, ByVal uFlags As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
  Private Declare Function SetMenuDefaultItem Lib "user32.dll" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
  Private Declare Function SetRect Lib "user32.dll" (lprc As RECT, ByVal xLeft As Long, ByVal yTop As Long, ByVal xRight As Long, ByVal yBottom As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SHCreateItemFromIDList Lib "shell32.dll" (ByVal pIDL As Long, riid As UUID, ppv As Any) As Long
  Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long
  Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal uiAction As Long, ByVal uiParam As Long, ByVal pvParam As Long, ByVal fWinIni As Long) As Long
  Private Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As Long, ByVal fuFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal lptpm As Long) As Long
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal pDestination As Long, ByVal length As Long)


  Public Event ChangeCurrentPIDL(ByVal pIDL As Long)
  Public Event ChangedCurrentPIDL()
  Public Event EmptySpaceClick()


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidUCBreadcrumb
      lRet = HandleMessage_UserControl(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "UCBreadcrumb.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in UCAddressBand.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_UserControl(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const MF_BYPOSITION As Long = &H400
  Const MFS_DISABLED As Long = &H3
  Const MFT_OWNERDRAW As Long = &H100
  Const MIIM_DATA As Long = &H20
  Const MIIM_FTYPE As Long = &H100
  Const MIIM_STATE As Long = &H1
  Const WM_DRAWITEM As Long = &H2B
  Const WM_ERASEBKGND As Long = &H14
  Const WM_INITMENUPOPUP = &H117
  Const WM_MEASUREITEM As Long = &H2C
  Const WM_MENUSELECT = &H11F
  Const WM_PAINT As Long = &HF
  Const WM_PRINTCLIENT As Long = &H318
  Const WM_SIZE As Long = &H5
  Const WM_THEMECHANGED As Long = &H31A
  Dim dispName As STRRET
  Dim drawItemData As DRAWITEMSTRUCT
  Dim hasSubMenu As Boolean
  Dim hDC As Long
  Dim IParentISF As IVBShellFolder
  Dim IParentISF2 As IVBShellFolder
  Dim itemIconIndex As Long
  Dim itemOverlayIndex As Long
  Dim itemText As String
  Dim itemType As MenuItemTypeConstants
  Dim measureItemData As MEASUREITEMSTRUCT
  Dim menuItem As MENUITEMINFO
  Dim miData As MENUITEMDATA
  Dim numberOfPIDLs As Long
  Dim pIDLs() As Long
  Dim pIDLNamespace As Long
  Dim pIDLToDesktop As Long
  Dim pIDLToParent As Long
  Dim pMIData As Long
  Dim ps As PAINTSTRUCT
  Dim rcClient As RECT
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_ERASEBKGND
      ' make sure the control's background is drawn properly
      If Not IsCompositionEnabled Then DoPaint wParam
      lRet = 1
      bCallDefProc = False
    Case WM_PAINT, WM_PRINTCLIENT
      ' make sure the control's background is drawn properly
      If wParam Then
        If IsCompositionEnabled Then DoPaint wParam
      Else
        hDC = BeginPaint(hWnd, ps)
        If IsCompositionEnabled Then DoPaint hDC
        EndPaint hWnd, ps
      End If
      bCallDefProc = False
    Case WM_SIZE
      ' reposition the progress bar and the tool bar
      rcClient.Right = LoWord(lParam)
      rcClient.Bottom = HiWord(lParam)
      rcClient.Top = rcClient.Top + (4 - m_textMetrics.tmHeight / 2)
      MoveWindow BreadcrumbBar.hWnd, rcClient.Left, rcClient.Top, rcClient.Right - rcClient.Left, rcClient.Bottom - rcClient.Top, 1
      ShowHideButtons
      bCallDefProc = False
    Case WM_THEMECHANGED
      CloseThemeData m_hTheme
      m_hTheme = OpenThemeData(hWnd, StrPtr("BreadcrumbBarComposited::BreadcrumbBar"))

    Case WM_INITMENUPOPUP
      If m_selectedMenuItemData Then
        If GetMenuItemCount(wParam) = 1 Then
          ' The menu contains 1 item. Is it our dummy?
          menuItem.cbSize = LenB(menuItem)
          menuItem.fMask = MIIM_DATA
          GetMenuItemInfo wParam, 0, 1, menuItem
          If menuItem.dwItemData = -1 Then
            ' it is
            RemoveMenu wParam, 0, MF_BYPOSITION
          End If
        End If

        If GetMenuItemCount(wParam) = 0 Then
          ' the sub-menu is empty, so fill it
          itemType = ExtractMenuItemData(m_selectedMenuItemData, IParentISF, pIDLToParent, pIDLToDesktop)
          If itemType = MenuItemTypeConstants.mitShellItem Then
            pIDLNamespace = pIDLToDesktop
            IParentISF.BindToObject pIDLToParent, 0, IID_IShellFolder, IParentISF2
            If Not (IParentISF2 Is Nothing) Then
              EnumSubItems IParentISF2, pIDLNamespace, pIDLs, numberOfPIDLs
              If numberOfPIDLs > 0 Then
                ' fill it with shell items
                QuickSort IParentISF2, pIDLs, 0, numberOfPIDLs - 1
                FillShellMenu wParam, IParentISF2, pIDLToDesktop, pIDLs, numberOfPIDLs, True, 0
              Else
                ' no shell items to display
                menuItem.cbSize = LenB(menuItem)
                menuItem.fMask = MIIM_FTYPE Or MIIM_STATE Or MIIM_DATA
                menuItem.fType = MFT_OWNERDRAW
                menuItem.fState = MFS_DISABLED
                miData.cbSize = LenB(miData)
                miData.itemType = MenuItemTypeConstants.mitEmptyItem
                pMIData = HeapAlloc(GetProcessHeap, 0, LenB(miData))
                CopyMemory pMIData, VarPtr(miData), miData.cbSize
                menuItem.dwItemData = pMIData
                InsertMenuItem wParam, 0, 1, menuItem
              End If
              Set IParentISF2 = Nothing
            End If
            Erase pIDLs
          End If
          Set IParentISF = Nothing
        End If
      End If

    Case WM_MENUSELECT
      menuItem.cbSize = LenB(menuItem)
      menuItem.fMask = MIIM_DATA
      GetMenuItemInfo lParam, LoWord(wParam), 1, menuItem
      m_selectedMenuItemData = menuItem.dwItemData

    Case WM_DRAWITEM
      CopyMemory VarPtr(drawItemData), lParam, LenB(drawItemData)
      If drawItemData.itemData Then
        itemType = ExtractMenuItemData(drawItemData.itemData, IParentISF, pIDLToParent, pIDLToDesktop, hasSubMenu)
        If (itemType = MenuItemTypeConstants.mitShellItem) Or (itemType = MenuItemTypeConstants.mitOpenThisFolderItem) Then
          itemIconIndex = GetSysIconIndex(pIDLToDesktop)
          If Not (IParentISF Is Nothing) Then
            itemOverlayIndex = GetOverlayIndex(IParentISF, pIDLToParent)

            dispName.uType = STRRET_OFFSET
            If IParentISF.GetDisplayNameOf(pIDLToParent, SHGDNConstants.SHGDN_INFOLDER, dispName) = S_OK Then
              itemText = STRRETToString(dispName, pIDLToParent)
            End If
            If itemType = MenuItemTypeConstants.mitOpenThisFolderItem Then
              itemText = "Open " & itemText
            End If
          End If
        ElseIf itemType = MenuItemTypeConstants.mitEmptyItem Then
          itemText = "(Empty)"
          itemIconIndex = -1
          itemOverlayIndex = 0
        End If
        Set IParentISF = Nothing

        DrawItem drawItemData.hDC, itemType, drawItemData.rcItem, itemText, itemIconIndex, itemOverlayIndex, hasSubMenu, drawItemData.itemState, drawItemData.itemAction
        lRet = 1
        bCallDefProc = False
      End If

    Case WM_MEASUREITEM
      CopyMemory VarPtr(measureItemData), lParam, LenB(measureItemData)
      If measureItemData.itemData Then
        itemType = ExtractMenuItemData(measureItemData.itemData, IParentISF, pIDLToParent, pIDLToDesktop)
        If (itemType = MenuItemTypeConstants.mitShellItem) Or (itemType = MenuItemTypeConstants.mitOpenThisFolderItem) Then
          itemIconIndex = GetSysIconIndex(pIDLToDesktop)
          If Not (IParentISF Is Nothing) Then
            itemOverlayIndex = GetOverlayIndex(IParentISF, pIDLToParent)

            dispName.uType = STRRET_OFFSET
            If IParentISF.GetDisplayNameOf(pIDLToParent, SHGDNConstants.SHGDN_INFOLDER, dispName) = S_OK Then
              itemText = STRRETToString(dispName, pIDLToParent)
            End If
            If itemType = MenuItemTypeConstants.mitOpenThisFolderItem Then
              itemText = "Open " & itemText
            End If
          End If
        ElseIf itemType = MenuItemTypeConstants.mitEmptyItem Then
          itemText = "(Empty)"
          itemIconIndex = -1
          itemOverlayIndex = 0
        End If
        Set IParentISF = Nothing

        MeasureItem itemText, itemType, measureItemData.itemWidth, measureItemData.itemHeight
        CopyMemory lParam, VarPtr(measureItemData), LenB(measureItemData)
        lRet = 1
        bCallDefProc = False
      End If
  End Select

StdHandler_Ende:
  HandleMessage_UserControl = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in UCBreadcrumb.HandleMessage_UserControl: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub BreadcrumbBar_Click(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If hitTestDetails And htNotInButton Then
    RaiseEvent EmptySpaceClick
  ElseIf Not toolButton Is Nothing Then
    If toolButton.ID >= BTNID_FIRSTPART Then
      RaiseEvent ChangeCurrentPIDL(toolButton.ButtonData)
    End If
  End If
End Sub

Private Sub BreadcrumbBar_CustomDraw(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal buttonState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, textRectangle As TBarCtlsLibUCtl.RECTANGLE, HorizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
  Const ILD_TRANSPARENT As Long = &H1
  Const SBDD_HOT As Long = 2
  Const SBDD_NORMAL As Long = 1
  Const SBDD_PRESSED As Long = 3
  Const SBP_OVERFLOWCHEVRON As Long = 1
  Const TP_SPLITBUTTONDROPDOWN As Long = 4
  Const TS_HOT As Long = 2
  Const TS_NORMAL As Long = 1
  Const TS_PRESSED As Long = 3
  Dim hTheme As Long
  Dim pt As Point
  Dim rcButton As RECT
  Dim stateID As Long

  If drawStage = cdsPrePaint Then
    furtherProcessing = cdrvNotifyItemDraw
  ElseIf drawStage = cdsItemPrePaint Then
    If toolButton.ID = BTNID_FOLDERICON Then
      LSet rcButton = drawingRectangle
      If toolButton.DropDownStyle = ddstNoDropDown Then
        rcButton.Left = (rcButton.Right - rcButton.Left - m_iconSize.cx) / 2
        rcButton.Top = (rcButton.Bottom - rcButton.Top - m_iconSize.cy) / 2
      Else
        toolButton.GetRectangle brtDropDown, rcButton.Left
        rcButton.Left = (rcButton.Left - drawingRectangle.Left - m_iconSize.cx) / 2
        rcButton.Top = (rcButton.Bottom - rcButton.Top - m_iconSize.cy) / 2
      End If
      ImageList_Draw m_hImageList, toolButton.IconIndex, hDC, rcButton.Left, rcButton.Top, ILD_TRANSPARENT

      If toolButton.DropDownStyle = ddstNormal Then
        LSet rcButton = drawingRectangle
        toolButton.GetRectangle brtDropDown, rcButton.Left, , rcButton.Right
        If m_bHasHiddenItems Then
          stateID = SBDD_NORMAL
          If buttonState And cdisHot Then
            GetCursorPos pt
            MapWindowPoints HWND_DESKTOP, BreadcrumbBar.hWnd, VarPtr(pt), 1
            If PtInRect(rcButton, pt.x, pt.y) Then
              stateID = SBDD_HOT
            End If
          End If
          If toolButton.DroppedDown Then stateID = SBDD_PRESSED
          DrawThemeBackground m_hTheme, hDC, SBP_OVERFLOWCHEVRON, stateID, rcButton, 0
        Else
          hTheme = OpenThemeData(UserControl.hWnd, StrPtr("BBComposited::ToolBar"))
          If hTheme Then
            stateID = TS_NORMAL
            If buttonState And cdisHot Then
              GetCursorPos pt
              MapWindowPoints HWND_DESKTOP, BreadcrumbBar.hWnd, VarPtr(pt), 1
              If PtInRect(rcButton, pt.x, pt.y) Then
                stateID = TS_HOT
              End If
            End If
            If toolButton.DroppedDown Then stateID = TS_PRESSED
            DrawThemeBackground hTheme, hDC, TP_SPLITBUTTONDROPDOWN, stateID, rcButton, 0
            CloseThemeData hTheme
          End If
        End If
      End If

      furtherProcessing = cdrvSkipDefault
    End If
  End If
End Sub

Private Sub BreadcrumbBar_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  Dim nextButton As Long
  Dim pIDLOfDefaultItem As Long
  Dim ptMenu As Point
  Dim rcButton As RECT

  ptMenu.x = buttonRectangle.Left
  ptMenu.y = buttonRectangle.Bottom
  toolButton.GetRectangle brtButton, rcButton.Left, rcButton.Top, rcButton.Right, rcButton.Bottom
  nextButton = toolButton.Index + 1
  If nextButton < BreadcrumbBar.Buttons.Count Then
    pIDLOfDefaultItem = BreadcrumbBar.Buttons.Item(nextButton).ButtonData
  End If
  DisplayShellMenu toolButton.ButtonData, ptMenu, rcButton, pIDLOfDefaultItem, m_bHasHiddenItems And (toolButton.ID = BTNID_FOLDERICON)
End Sub

Private Sub BreadcrumbBar_FreeButtonData(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton)
  If toolButton.ButtonData Then ILFree toolButton.ButtonData
End Sub

Private Sub BreadcrumbBar_MouseMove(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Dim rcButton As RECT

  If Not (toolButton Is Nothing) Then
    If toolButton.ID = BTNID_FOLDERICON Then
      toolButton.GetRectangle brtButton, rcButton.Left, rcButton.Top, rcButton.Right, rcButton.Bottom
      InvalidateRect BreadcrumbBar.hWnd, VarPtr(rcButton), 1
    End If
  End If
End Sub

Private Sub UserControl_Initialize()
  Const ILC_COLOR32 As Long = &H20
  Const IMAGE_BITMAP As Long = 0
  Const LR_CREATEDIBSECTION As Long = &H2000&
  Const LR_LOADFROMFILE As Long = &H10
  Const MENU_POPUPITEM As Long = 14
  Const MPI_HOT As Long = 2
  Const SHIL_SMALL As Long = &H1
  Const strIID_IImageList As String = "{46EB5926-582E-4017-9FDF-E8998DAA0950}"
  Const strIID_IShellFolder As String = "{000214E6-0000-0000-C000-000000000046}"
  Const strIID_IShellIconOverlay As String = "{7D688A70-C613-11D0-999B-00C04FD655E1}"
  Const strIID_IShellItem As String = "{43826d1e-e718-42ee-bc55-a1e261c37bfe}"
  Const WM_GETFONT As Long = &H31
  Dim hDC As Long
  Dim hFont As Long
  Dim hPreviousFont As Long
  Dim hTheme As Long
  Dim imgDir As String

  CLSIDFromString StrPtr(strIID_IImageList), IID_IImageList
  CLSIDFromString StrPtr(strIID_IShellFolder), IID_IShellFolder
  CLSIDFromString StrPtr(strIID_IShellIconOverlay), IID_IShellIconOverlay
  CLSIDFromString StrPtr(strIID_IShellItem), IID_IShellItem

  m_hTheme = OpenThemeData(UserControl.hWnd, StrPtr("BreadcrumbBarComposited::BreadcrumbBar"))
  If IsAppThemed Then
    hTheme = OpenThemeData(UserControl.hWnd, StrPtr("MENU"))
    If hTheme Then
      If IsThemePartDefined(hTheme, MENU_POPUPITEM, MPI_HOT) Then
        m_themedMenus = True
      End If
      CloseThemeData hTheme
    End If
  End If

  SHGetDesktopFolder m_IDesktop
  SHGetImageList SHIL_SMALL, IID_IImageList, VarPtr(m_hImageList)
  ImageList_GetIconSize m_hImageList, m_iconSize.cx, m_iconSize.cy

  If Right$(App.Path, 3) = "bin" Then
    imgDir = App.Path & "\..\res\"
  Else
    imgDir = App.Path & "\res\"
  End If
  m_hImageListDummy = ImageList_Create(8, 16, ILC_COLOR32, 0, 0)
  BreadcrumbBar.hImageList(ilNormalButtons) = m_hImageListDummy

  hDC = GetDC(UserControl.hWnd)
  If hDC Then
    hFont = SendMessage(BreadcrumbBar.hWnd, WM_GETFONT, 0, 0)
    hPreviousFont = SelectObject(hDC, hFont)
    GetTextMetrics hDC, m_textMetrics
    SelectObject hDC, hPreviousFont
    ReleaseDC UserControl.hWnd, hDC
  End If

  ' set some special themes for the tool bar
  SetWindowTheme BreadcrumbBar.hWnd, StrPtr("BBComposited"), 0
End Sub

Private Sub UserControl_Resize()
  Dim rcClient As RECT

  If IsInIDE Then
    ' during runtime the child controls will be aligned whenever we receive a WM_SIZE message
    GetClientRect UserControl.hWnd, rcClient
    rcClient.Top = rcClient.Top + (4 - m_textMetrics.tmHeight / 2)
    MoveWindow BreadcrumbBar.hWnd, rcClient.Left, rcClient.Top, rcClient.Right - rcClient.Left, rcClient.Bottom - rcClient.Top, 1
    ShowHideButtons
  ElseIf Not m_subclassed Then
    If SubclassWindow(UserControl.hWnd, Me, EnumSubclassID.escidUCBreadcrumb) Then
      m_subclassed = True
    Else
      Debug.Print "Subclassing failed!"
    End If
  End If
End Sub

Private Sub UserControl_Terminate()
  BreadcrumbBar.Buttons.RemoveAll

  If UnSubclassWindow(UserControl.hWnd, EnumSubclassID.escidUCBreadcrumb) Then
    m_subclassed = False
  Else
    Debug.Print "UnSubclassing failed!"
  End If
  If m_hTheme Then
    CloseThemeData m_hTheme
  End If
  If m_hImageListDummy Then
    ImageList_Destroy m_hImageListDummy
    m_hImageListDummy = 0
  End If
End Sub


Public Property Get BreadcrumbToolBar() As ToolBar
  Set BreadcrumbToolBar = BreadcrumbBar
End Property

Public Property Get CurrentDirectoryName() As String
  With BreadcrumbBar.Buttons
    If .Count > 0 Then
      With .Item(.Count - 1)
        If .ID >= BTNID_FIRSTPART Then
          CurrentDirectoryName = .Text
        End If
      End With
    End If
  End With
End Property

Public Property Get CurrentPIDL() As Long
Attribute CurrentPIDL.VB_ProcData.VB_Invoke_Property = ";Daten"
Attribute CurrentPIDL.VB_MemberFlags = "400"
  CurrentPIDL = m_currentPIDL
End Property

Public Property Let CurrentPIDL(ByVal newValue As Long)
  If m_currentPIDL <> newValue Then
    If m_currentPIDL Then
      ILFree m_currentPIDL
    End If
    m_currentPIDL = newValue
    SetButtons
  End If
End Property

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Property Get hWndToolBar() As Long
  hWndToolBar = BreadcrumbBar.hWnd
End Property

Public Property Get IsDropped() As Boolean
  IsDropped = False
End Property

Public Sub RefreshBreadcrumbBar()
  BreadcrumbBar.Refresh
End Sub

Public Sub SetNotifyWindow(ByVal hWnd As Long)
  m_hWndNotify = hWnd
End Sub


Private Sub ClearMenu(ByVal hMenu As Long)
  Const MF_BYPOSITION As Long = &H400
  Const MIIM_DATA As Long = &H20
  Const MIIM_SUBMENU As Long = &H4
  Dim menuItem As MENUITEMINFO
  Dim miData As MENUITEMDATA

  menuItem.cbSize = LenB(menuItem)
  menuItem.fMask = MIIM_DATA Or MIIM_SUBMENU
  Do While GetMenuItemCount(hMenu) > 0
    GetMenuItemInfo hMenu, 0, 1, menuItem
    If menuItem.hSubMenu Then
      ClearMenu menuItem.hSubMenu
    End If
    If (menuItem.dwItemData <> 0) And (menuItem.dwItemData <> -1) Then
      CopyMemory VarPtr(miData), menuItem.dwItemData, LenB(miData)
      If (miData.itemType = MenuItemTypeConstants.mitShellItem) Or (miData.itemType = MenuItemTypeConstants.mitOpenThisFolderItem) Then
        Set miData.IParentISF = Nothing
        ILFree miData.pIDLToParent
        ILFree miData.pIDLToDesktop
      End If
      HeapFree GetProcessHeap, 0, menuItem.dwItemData
    End If
    RemoveMenu hMenu, 0, MF_BYPOSITION
  Loop
End Sub

Private Function CreateShellMenu(ByVal pIDLNamespace As Long, ByVal withSubMenus As Boolean, ByVal pIDLOfDefaultItem As Long, Optional ByVal bAddHiddenButtons As Boolean = False) As Long
  Dim freePIDLs As Boolean
  Dim hMenu As Long
  Dim i As Long
  Dim IParentISF As IVBShellFolder
  Dim numberOfPIDLs As Long
  Dim pIDLs() As Long

  Set IParentISF = GetISHFolderInterfaceFQ(pIDLNamespace)
  If Not (IParentISF Is Nothing) Then
    EnumSubItems IParentISF, pIDLNamespace, pIDLs, numberOfPIDLs
    If numberOfPIDLs > 0 Then
      freePIDLs = True
      QuickSort IParentISF, pIDLs, 0, numberOfPIDLs - 1
      hMenu = CreatePopupMenu
      If hMenu Then
        FillShellMenu hMenu, IParentISF, pIDLNamespace, pIDLs, numberOfPIDLs, withSubMenus, pIDLOfDefaultItem, bAddHiddenButtons
        freePIDLs = False
      End If
    End If
    Set IParentISF = Nothing
  End If
  If freePIDLs Then
    For i = 0 To numberOfPIDLs - 1
      ILFree pIDLs(i)
    Next i
  End If
  Erase pIDLs

  CreateShellMenu = hMenu
End Function

Private Sub DisplayShellMenu(ByVal pIDL As Long, ptMenu As Point, rcExclude As RECT, ByVal pIDLOfDefaultItem As Long, ByVal bAddHiddenButtons As Boolean)
  Const MIIM_DATA As Long = &H20
  Const SM_MENUDROPALIGNMENT As Long = 40
  Const SPI_GETNONCLIENTMETRICS As Long = 41
  Const TPM_LEFTALIGN As Long = &H0
  Const TPM_LEFTBUTTON As Long = &H0
  Const TPM_RETURNCMD As Long = &H100
  Const TPM_RIGHTALIGN As Long = &H8
  Const TPM_TOPALIGN As Long = &H0
  Dim commandID As Long
  Dim dispName As STRRET
  Dim flags As Long
  Dim hMenu As Long
  Dim IParentISF As IVBShellFolder
  Dim menuItem As MENUITEMINFO
  Dim ncMetrics As NONCLIENTMETRICS
  Dim pIDLToDesktop As Long
  Dim pIDLToParent As Long
  Dim tpmexParams As TPMPARAMS

  If pIDL = 0 Then Exit Sub

  m_nextCommandID = 1
  ncMetrics.cbSize = LenB(ncMetrics)
  If SystemParametersInfo(SPI_GETNONCLIENTMETRICS, ncMetrics.cbSize, VarPtr(ncMetrics), 0) Then
    m_hMenuFont = CreateFontIndirect(ncMetrics.lfMenuFont)
  End If

  hMenu = CreateShellMenu(pIDL, False, pIDLOfDefaultItem, bAddHiddenButtons)
  If hMenu Then
    flags = TPM_TOPALIGN Or TPM_LEFTBUTTON Or TPM_RETURNCMD
    If GetSystemMetrics(SM_MENUDROPALIGNMENT) Then
      flags = flags Or TPM_RIGHTALIGN
    Else
      flags = flags Or TPM_LEFTALIGN
    End If
    MapWindowPoints BreadcrumbBar.hWnd, HWND_DESKTOP, VarPtr(ptMenu), 1
    With tpmexParams
      .cbSize = LenB(tpmexParams)
      .rcExclude = rcExclude
      MapWindowPoints BreadcrumbBar.hWnd, HWND_DESKTOP, VarPtr(.rcExclude), 2
    End With

    commandID = TrackPopupMenuEx(hMenu, flags, ptMenu.x, ptMenu.y, UserControl.hWnd, VarPtr(tpmexParams))
    If commandID > 0 Then
      menuItem.cbSize = LenB(menuItem)
      menuItem.fMask = MIIM_DATA
      GetMenuItemInfo hMenu, commandID, 0, menuItem
      ExtractMenuItemData menuItem.dwItemData, IParentISF, pIDLToParent, pIDLToDesktop
      Set IParentISF = Nothing

      If pIDLToDesktop Then
        RaiseEvent ChangeCurrentPIDL(pIDLToDesktop)
      End If
    End If

    ClearMenu hMenu
    DestroyMenu hMenu
  End If

  If m_hMenuFont Then
    DeleteObject m_hMenuFont
    m_hMenuFont = 0
  End If
End Sub

Private Sub DoPaint(ByVal hDC As Long)
  Dim rcClient As RECT

  ' draw the control background using the theme API
  GetClientRect UserControl.hWnd, rcClient
  DrawThemeParentBackground UserControl.hWnd, hDC, rcClient
End Sub

Private Sub DrawItem(ByVal hDC As Long, itemType As MenuItemTypeConstants, itemRectangle As RECT, itemText As String, itemIconIndex As Long, itemOverlayIndex As Long, drawSubMenuArrow As Boolean, itemState As Long, itemAction As Long)
  Const COLOR_HIGHLIGHT As Long = 13
  Const COLOR_HIGHLIGHTTEXT As Long = 14
  Const COLOR_MENU As Long = 4
  Const COLOR_MENUTEXT As Long = 7
  Const DT_LEFT As Long = &H0
  Const DT_NOPREFIX As Long = &H800
  Const DT_SINGLELINE As Long = &H20
  Const DT_VCENTER As Long = &H4
  Const ILD_TRANSPARENT As Long = &H1
  Const MENU_POPUPBACKGROUND As Long = 9
  Const MENU_POPUPGUTTER As Long = 13
  Const MENU_POPUPITEM As Long = 14
  Const MENU_POPUPSEPARATOR As Long = 15
  Const MENU_POPUPSUBMENU As Long = 16
  Const MPI_DISABLED As Long = 3
  Const MPI_DISABLEDHOT As Long = 4
  Const MPI_HOT As Long = 2
  Const MPI_NORMAL As Long = 1
  Const MSM_DISABLED As Long = 2
  Const MSM_NORMAL As Long = 1
  Const ODS_DISABLED As Long = &H4
  Const ODS_HOTLIGHT As Long = &H40
  Const ODS_INACTIVE As Long = &H80
  Const ODS_SELECTED As Long = &H1
  Const TMT_CONTENTMARGINS As Long = 3602
  Const TRANSPARENT As Long = 1
  Const TS_TRUE As Long = 1
  Const WM_GETFONT As Long = &H31
  Dim arrowRectangle As RECT
  Dim arrowStateID As Long
  Dim cx As Long
  Dim cy As Long
  Dim hOldFont As Long
  Dim hTheme As Long
  Dim itemMetrics As DRAWITEMMETRICS
  Dim itemStateID As Long
  Dim measureRectangle As RECT
  Dim oldColor As Long
  Dim overlayMask As Long
  Dim popupSubmenuArrowMargins As MARGINS
  Dim popupSubmenuArrowSize As Size
  Dim textRectangle As RECT
  Dim x As Long
  Dim y As Long

  hTheme = OpenThemeData(UserControl.hWnd, StrPtr("MENU"))
  If hTheme Then
    If itemState And (ODS_INACTIVE Or ODS_DISABLED) Then
      If itemState And (ODS_HOTLIGHT Or ODS_SELECTED) Then
        itemStateID = MPI_DISABLEDHOT
      Else
        itemStateID = MPI_DISABLED
      End If
      arrowStateID = MSM_DISABLED
    ElseIf itemState And (ODS_HOTLIGHT Or ODS_SELECTED) Then
      itemStateID = MPI_HOT
      arrowStateID = MSM_NORMAL
    Else
      itemStateID = MPI_NORMAL
      arrowStateID = MSM_NORMAL
    End If

    LayoutMenuItem hTheme, itemType, itemRectangle, itemText, itemMetrics

    If IsThemeBackgroundPartiallyTransparent(hTheme, MENU_POPUPITEM, itemStateID) Then
      DrawThemeBackground hTheme, hDC, MENU_POPUPBACKGROUND, 0, itemRectangle, 0
    End If
    DrawThemeBackground hTheme, hDC, MENU_POPUPGUTTER, 0, itemMetrics.rcGutter, 0

    If itemType = MenuItemTypeConstants.mitSeparator Then
      DrawThemeBackground hTheme, hDC, MENU_POPUPSEPARATOR, 0, itemMetrics.rcSeparator, 0
    Else
      DrawThemeBackground hTheme, hDC, MENU_POPUPITEM, itemStateID, itemMetrics.rcSelection, 0

      If itemIconIndex >= 0 Then
        If itemOverlayIndex > 0 Then
          overlayMask = itemOverlayIndex * (2 ^ 8)
        End If
        ImageList_Draw m_hImageList, itemIconIndex, hDC, itemMetrics.rcIcon.Left, itemMetrics.rcIcon.Top, ILD_TRANSPARENT Or overlayMask
      End If
      DrawThemeText hTheme, hDC, MENU_POPUPITEM, itemStateID, StrPtr(itemText), Len(itemText), DT_NOPREFIX Or DT_LEFT Or DT_SINGLELINE, 0, itemMetrics.rcText
    End If

    If drawSubMenuArrow Then
      GetThemePartSize hTheme, 0, MENU_POPUPSUBMENU, 0, 0, TS_TRUE, popupSubmenuArrowSize
      GetThemeMargins hTheme, 0, MENU_POPUPSUBMENU, 0, TMT_CONTENTMARGINS, 0, popupSubmenuArrowMargins

      popupSubmenuArrowSize.cx = popupSubmenuArrowSize.cx + popupSubmenuArrowMargins.cxLeftWidth + popupSubmenuArrowMargins.cxRightWidth
      popupSubmenuArrowSize.cy = popupSubmenuArrowSize.cy + popupSubmenuArrowMargins.cyTopHeight + popupSubmenuArrowMargins.cyBottomHeight

      ' TODO: this calculation is wrong
      cx = popupSubmenuArrowSize.cx
      cy = popupSubmenuArrowSize.cy
      x = itemRectangle.Right - cx
      y = itemRectangle.Top
      SetRect measureRectangle, x, y, x + cx, y + cy
      SetRect arrowRectangle, measureRectangle.Left + popupSubmenuArrowMargins.cxLeftWidth, measureRectangle.Top + popupSubmenuArrowMargins.cyTopHeight, measureRectangle.Right - popupSubmenuArrowMargins.cxRightWidth, measureRectangle.Bottom - popupSubmenuArrowMargins.cyBottomHeight
      OffsetRect arrowRectangle, 0, ((itemRectangle.Bottom - itemRectangle.Top) - cy) / 2
      ' correction for the failure mentioned above
      OffsetRect arrowRectangle, 2, -2

      DrawThemeBackground hTheme, hDC, MENU_POPUPSUBMENU, arrowStateID, arrowRectangle, 0

      ' this will prevent drawing of the default arrow
      ExcludeClipRect hDC, itemRectangle.Left, itemRectangle.Top, itemRectangle.Right, itemRectangle.Bottom
    End If

    CloseThemeData hTheme
    Exit Sub
  End If

  If m_hMenuFont Then
    hOldFont = SelectObject(hDC, m_hMenuFont)
  Else
    hOldFont = SelectObject(hDC, SendMessage(UserControl.hWnd, WM_GETFONT, 0, 0))
  End If

  If itemState And ODS_SELECTED Then
    FillRect hDC, itemRectangle, GetSysColorBrush(COLOR_HIGHLIGHT)
  Else
    FillRect hDC, itemRectangle, GetSysColorBrush(COLOR_MENU)
  End If
  x = itemRectangle.Left + HMARGIN_ICON
  If itemIconIndex >= 0 Then
    y = itemRectangle.Top + (itemRectangle.Bottom - itemRectangle.Top - m_iconSize.cy) / 2
    If itemOverlayIndex > 0 Then
      overlayMask = itemOverlayIndex * (2 ^ 8)
    End If
    ImageList_Draw m_hImageList, itemIconIndex, hDC, x, y, ILD_TRANSPARENT Or overlayMask
  End If
  x = x + m_iconSize.cx + HMARGIN_ICON + HMARGIN_TEXT
  If itemText <> "" Then
    SetBkMode hDC, TRANSPARENT
    If itemState And ODS_SELECTED Then
      oldColor = SetTextColor(hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
    Else
      oldColor = SetTextColor(hDC, GetSysColor(COLOR_MENUTEXT))
    End If

    textRectangle.Left = x
    textRectangle.Top = itemRectangle.Top
    textRectangle.Right = itemRectangle.Right - HMARGIN_TEXT
    textRectangle.Bottom = itemRectangle.Bottom
    DrawText hDC, StrPtr(itemText), Len(itemText), textRectangle, DT_NOPREFIX Or DT_LEFT Or DT_SINGLELINE Or DT_VCENTER

    SetTextColor hDC, oldColor
  End If

  If hOldFont Then
    SelectObject hDC, hOldFont
  End If
End Sub

Private Sub EnumSubItems(IParentISF As IVBShellFolder, ByVal pIDLNamespace As Long, ByRef pIDLs() As Long, ByRef numberOfPIDLs As Long)
  Dim IEnum As IVBEnumIDList
  Dim pIDLSubItem_ToParent As Long

  numberOfPIDLs = 0
  Erase pIDLs

  If Not (IParentISF Is Nothing) Then
    'IParentISF.EnumObjects UserControl.hWnd, SHCONTFConstants.SHCONTF_FOLDERS, IEnum
    IParentISF.EnumObjects 0, SHCONTFConstants.SHCONTF_FOLDERS, IEnum
    If Not (IEnum Is Nothing) Then
      ReDim pIDLs(0 To 0) As Long
      Do While IEnum.Next(1, pIDLSubItem_ToParent, 0) = S_OK
        If numberOfPIDLs = UBound(pIDLs) Then
          ReDim Preserve pIDLs(0 To numberOfPIDLs + 9) As Long
        End If

        pIDLs(numberOfPIDLs) = pIDLSubItem_ToParent
        numberOfPIDLs = numberOfPIDLs + 1
      Loop
      Set IEnum = Nothing
    End If
  End If
End Sub

Private Function ExtractMenuItemData(ByVal pMenuItemData As Long, IParentISF As IVBShellFolder, pIDLToParent As Long, pIDLToDesktop As Long, Optional hasSubMenu As Boolean = False) As MenuItemTypeConstants
  Dim miData As MENUITEMDATA

  If pMenuItemData Then
    CopyMemory VarPtr(miData), pMenuItemData, LenB(miData)
    If (miData.itemType = MenuItemTypeConstants.mitShellItem) Or (miData.itemType = MenuItemTypeConstants.mitOpenThisFolderItem) Then
      Set IParentISF = miData.IParentISF
      pIDLToParent = miData.pIDLToParent
      pIDLToDesktop = miData.pIDLToDesktop
      ZeroMemory VarPtr(miData.IParentISF), 4
    End If
    hasSubMenu = miData.hasSubMenu
    ExtractMenuItemData = miData.itemType
  End If
End Function

Private Sub FillShellMenu(ByVal hMenu As Long, IParentISF As IVBShellFolder, ByVal pIDLParentToDesktop As Long, ByRef pIDLs() As Long, numberOfPIDLs As Long, ByVal withSubMenus As Boolean, ByVal pIDLOfDefaultItem As Long, Optional ByVal bAddHiddenButtons As Boolean = False)
  Const CSIDL_DESKTOP As Long = &H0
  Const MFS_ENABLED As Long = &H0
  Const MFT_OWNERDRAW As Long = &H100
  Const MFT_SEPARATOR As Long = &H800
  Const MFT_STRING As Long = &H0
  Const MIIM_DATA As Long = &H20
  Const MIIM_FTYPE As Long = &H100
  Const MIIM_ID As Long = &H2
  Const MIIM_STATE As Long = &H1
  Const MIIM_STRING As Long = &H40
  Const MIIM_SUBMENU As Long = &H4
  Dim bInsertDesktop As Boolean
  Dim btn As ToolBarButton
  Dim dispName As STRRET
  Dim hSubMenu As Long
  Dim i As Long
  Dim IParentISF2 As IVBShellFolder
  Dim itemAttributes As SFGAOConstants
  Dim menuItem As MENUITEMINFO
  Dim menuItem2 As MENUITEMINFO
  Dim miData As MENUITEMDATA
  Dim numberOfExtraItems As Long
  Dim path As String
  Dim pIDLToDesktop As Long
  Dim pIDLToParent As Long
  Dim pMIData As Long

  If numberOfPIDLs > 0 Then
    menuItem.cbSize = LenB(menuItem)
    menuItem.fMask = MIIM_ID Or MIIM_FTYPE Or MIIM_STATE Or MIIM_STRING Or MIIM_DATA Or MIIM_SUBMENU
    menuItem.fType = MFT_OWNERDRAW
    menuItem.fState = MFS_ENABLED
    miData.cbSize = LenB(miData)
    For i = 0 To numberOfPIDLs - 1
      menuItem.wID = m_nextCommandID
      On Error Resume Next
      m_nextCommandID = m_nextCommandID + 1
      If Err Then
        MsgBox "Too many menu items!"
      End If
      On Error GoTo 0

      miData.itemType = MenuItemTypeConstants.mitShellItem
      Set miData.IParentISF = IParentISF
      miData.pIDLToParent = pIDLs(i)
      If pIDLParentToDesktop Then
        miData.pIDLToDesktop = ILAppendID(ILClone(pIDLParentToDesktop), miData.pIDLToParent, 1)
      Else
        dispName.uType = STRRET_OFFSET
        If miData.IParentISF.GetDisplayNameOf(miData.pIDLToParent, SHGDNConstants.SHGDN_FORPARSING, dispName) = S_OK Then
          path = STRRETToString(dispName, miData.pIDLToParent)
          miData.pIDLToDesktop = ILCreateFromPath(StrPtr(path))
        End If
      End If

      miData.hasSubMenu = False
      If withSubMenus Then
        menuItem.hSubMenu = 0
        itemAttributes = SFGAOConstants.SFGAO_HASSUBFOLDER
        miData.IParentISF.GetAttributesOf 1, miData.pIDLToParent, itemAttributes
        If itemAttributes And SFGAOConstants.SFGAO_HASSUBFOLDER Then
          ' create a sub-menu and insert a dummy item
          hSubMenu = CreatePopupMenu
          menuItem2.cbSize = LenB(menuItem2)
          menuItem2.fMask = MIIM_FTYPE Or MIIM_DATA Or MIIM_STRING
          menuItem2.dwItemData = -1
          menuItem2.fType = MFT_STRING
          menuItem2.dwTypeData = StrPtr("Dummy")
          InsertMenuItem hSubMenu, 0, 1, menuItem2
          menuItem.hSubMenu = hSubMenu
          miData.hasSubMenu = True
        End If
      End If

      pMIData = HeapAlloc(GetProcessHeap, 0, LenB(miData))
      CopyMemory pMIData, VarPtr(miData), miData.cbSize
      menuItem.dwItemData = pMIData
      ZeroMemory VarPtr(miData.IParentISF), 4

      InsertMenuItem hMenu, i, 1, menuItem
      If pIDLOfDefaultItem Then
        If ILIsEqual(miData.pIDLToDesktop, pIDLOfDefaultItem) Then
          SetMenuDefaultItem hMenu, i, 1
          pIDLOfDefaultItem = 0
        End If
      End If
    Next i
  End If

  If bAddHiddenButtons And Not BreadcrumbBar.Buttons(BTNID_FIRSTPART, btitID).Visible Then
    menuItem.cbSize = LenB(menuItem)
    menuItem.fMask = MIIM_FTYPE Or MIIM_DATA
    menuItem.fType = MFT_SEPARATOR
    If m_themedMenus Then
      menuItem.fType = menuItem.fType Or MFT_OWNERDRAW
    End If
    miData.cbSize = LenB(miData)
    If GetMenuItemCount(hMenu) Then
      miData.itemType = MenuItemTypeConstants.mitSeparator
      pMIData = HeapAlloc(GetProcessHeap, 0, LenB(miData))
      CopyMemory pMIData, VarPtr(miData), miData.cbSize
      menuItem.dwItemData = pMIData
      InsertMenuItem hMenu, 0, 1, menuItem
    End If

    menuItem.fMask = MIIM_ID Or MIIM_FTYPE Or MIIM_STATE Or MIIM_STRING Or MIIM_DATA
    menuItem.fType = MFT_OWNERDRAW
    menuItem.fState = MFS_ENABLED
    miData.itemType = MenuItemTypeConstants.mitShellItem
    miData.hasSubMenu = False
    For i = 0 To BreadcrumbBar.Buttons.Count - 1
      Set btn = BreadcrumbBar.Buttons(i)
      If btn.ID >= BTNID_FIRSTPART Then
        If btn.Visible Then Exit For

        If btn.ID = BTNID_FIRSTPART Then
          bInsertDesktop = (ILGetSize(btn.ButtonData) > 2)
        End If
        menuItem.wID = m_nextCommandID
        On Error Resume Next
        m_nextCommandID = m_nextCommandID + 1
        If Err Then
          MsgBox "Too many menu items!"
        End If
        On Error GoTo 0

        pIDLToDesktop = ILClone(btn.ButtonData)
        SplitFullyQualifiedPIDL pIDLToDesktop, IParentISF2, pIDLToParent
        If Not (IParentISF2 Is Nothing) And pIDLToParent <> 0 Then
          Set miData.IParentISF = IParentISF2
          miData.pIDLToParent = pIDLToParent
          miData.pIDLToDesktop = pIDLToDesktop
          pMIData = HeapAlloc(GetProcessHeap, 0, LenB(miData))
          CopyMemory pMIData, VarPtr(miData), miData.cbSize
          menuItem.dwItemData = pMIData
          ZeroMemory VarPtr(miData.IParentISF), 4
          InsertMenuItem hMenu, 0, 1, menuItem
          numberOfExtraItems = numberOfExtraItems + 1
        Else
          ILFree pIDLToDesktop
        End If
        Set IParentISF2 = Nothing
      End If
    Next i
    If bInsertDesktop Then
      menuItem.wID = m_nextCommandID
      On Error Resume Next
      m_nextCommandID = m_nextCommandID + 1
      If Err Then
        MsgBox "Too many menu items!"
      End If
      On Error GoTo 0

      SHGetFolderLocation UserControl.hWnd, CSIDL_DESKTOP, 0, 0, pIDLToDesktop
      Set miData.IParentISF = m_IDesktop
      miData.pIDLToParent = pIDLToDesktop
      miData.pIDLToDesktop = pIDLToDesktop
      pMIData = HeapAlloc(GetProcessHeap, 0, LenB(miData))
      CopyMemory pMIData, VarPtr(miData), miData.cbSize
      menuItem.dwItemData = pMIData
      ZeroMemory VarPtr(miData.IParentISF), 4
      InsertMenuItem hMenu, numberOfExtraItems, 1, menuItem
    End If
  End If
End Sub

Private Function GetISHFolderInterfaceFQ(pIDLToDesktop As Long) As IVBShellFolder
  Dim dummy As IVBShellFolder
  Dim pIDLToParent As Long
  Dim ret As IVBShellFolder

  If ILGetSize(pIDLToDesktop) = 2 Then
    Set GetISHFolderInterfaceFQ = m_IDesktop
  Else
    SplitFullyQualifiedPIDL pIDLToDesktop, dummy, pIDLToParent
    If Not (dummy Is Nothing) Then
      dummy.BindToObject pIDLToParent, 0, IID_IShellFolder, ret
    End If
    Set GetISHFolderInterfaceFQ = ret
    Set dummy = Nothing
    Set ret = Nothing
  End If
End Function

Private Function GetOverlayIndex(IParentISF As IVBShellFolder, pIDLToParent As Long) As Long
  Dim IParentISIO As IVBShellIconOverlay
  Dim itemAttributes As SFGAOConstants
  Dim overlayIndex As Long

  IParentISF.QueryInterface IID_IShellIconOverlay, IParentISIO
  If IParentISIO Is Nothing Then
    itemAttributes = SFGAOConstants.SFGAO_SHARE Or SFGAOConstants.SFGAO_LINK
    IParentISF.GetAttributesOf 1, pIDLToParent, itemAttributes
    If itemAttributes And SFGAOConstants.SFGAO_SHARE Then
      overlayIndex = 1
    End If
    If itemAttributes And SFGAOConstants.SFGAO_LINK Then
      overlayIndex = 2
    End If
  Else
    If IParentISIO.GetOverlayIndex(pIDLToParent, overlayIndex) <> S_OK Then
      overlayIndex = 0
    End If
    Set IParentISIO = Nothing
  End If

  GetOverlayIndex = overlayIndex
End Function

Private Function IsInIDE() As Boolean
  IsInIDE = Not Ambient.UserMode
End Function

Private Function IsLower(IParentISF As IVBShellFolder, firstPIDL As Long, secondPIDL As Long) As Boolean
  IsLower = (IParentISF.CompareIDs(0, firstPIDL, secondPIDL) < 0)
End Function

Private Sub LayoutMenuItem(hTheme As Long, itemType As MenuItemTypeConstants, itemRectangle As RECT, itemText As String, itemMetrics As DRAWITEMMETRICS)
  Const MENU_POPUPBACKGROUND As Long = 9
  Const MENU_POPUPCHECK As Long = 11
  Const MENU_POPUPCHECKBACKGROUND As Long = 12
  Const MENU_POPUPITEM As Long = 14
  Const TMT_BORDERSIZE As Long = 2403
  Const TMT_CONTENTMARGINS As Long = 3602
  Dim cx As Long
  Dim cy As Long
  Dim drawingSize As Size
  Dim iconAreaSize As Size
  Dim itemHeight As Long
  Dim measureRectangle As RECT
  Dim measureSize As Size
  Dim popupBgBorderSize As Long
  Dim popupBorderSize As Long
  Dim popupCheckBgMargins As MARGINS
  Dim popupCheckMargins As MARGINS
  Dim popupItemMargins As MARGINS
  Dim popupTextMargins As MARGINS
  Dim separatorSize As Size
  Dim textSize As Size
  Dim x As Long
  Dim y As Long

  If hTheme Then
    itemHeight = (itemRectangle.Bottom - itemRectangle.Top)

    GetThemeInt hTheme, MENU_POPUPITEM, 0, TMT_BORDERSIZE, popupBorderSize
    GetThemeInt hTheme, MENU_POPUPBACKGROUND, 0, TMT_BORDERSIZE, popupBgBorderSize
    GetThemeMargins hTheme, 0, MENU_POPUPCHECKBACKGROUND, 0, TMT_CONTENTMARGINS, 0, popupCheckBgMargins
    GetThemeMargins hTheme, 0, MENU_POPUPCHECK, 0, TMT_CONTENTMARGINS, 0, popupCheckMargins
    GetThemeMargins hTheme, 0, MENU_POPUPITEM, 0, TMT_CONTENTMARGINS, 0, popupItemMargins
    popupTextMargins = popupItemMargins
    popupTextMargins.cxRightWidth = popupBorderSize
    popupTextMargins.cxLeftWidth = popupBgBorderSize

    MeasureItemParts hTheme, itemType, itemText, iconAreaSize, textSize, separatorSize

    x = itemRectangle.Left + popupItemMargins.cxLeftWidth
    y = itemRectangle.Top

    If iconAreaSize.cx Then
      x = x + popupCheckBgMargins.cxLeftWidth
      cx = iconAreaSize.cx
      cy = iconAreaSize.cy - (popupCheckBgMargins.cyTopHeight + popupCheckBgMargins.cyBottomHeight)
      SetRect measureRectangle, x, y, x + cx, y + cy
      SetRect itemMetrics.rcIcon, measureRectangle.Left + popupCheckMargins.cxLeftWidth, measureRectangle.Top + popupCheckMargins.cyTopHeight, measureRectangle.Right - popupCheckMargins.cxRightWidth, measureRectangle.Bottom - popupCheckMargins.cyBottomHeight
      OffsetRect itemMetrics.rcIcon, 0, (itemHeight - cy) / 2
      x = x + popupCheckBgMargins.cxRightWidth
      x = x + cx
    End If
    If textSize.cx Then
      cx = textSize.cx
      cy = textSize.cy
      SetRect measureRectangle, x, y, x + cx, y + cy
      SetRect itemMetrics.rcText, measureRectangle.Left + popupTextMargins.cxLeftWidth, measureRectangle.Top + popupTextMargins.cyTopHeight, measureRectangle.Right - popupTextMargins.cxRightWidth, measureRectangle.Bottom - popupTextMargins.cyBottomHeight
      OffsetRect itemMetrics.rcText, 0, (itemHeight - cy) / 2
      x = x + cx
    End If

    drawingSize.cx = iconAreaSize.cx
    drawingSize.cy = iconAreaSize.cy - (popupCheckBgMargins.cyTopHeight + popupCheckBgMargins.cyBottomHeight)

    measureSize.cx = drawingSize.cx + popupCheckBgMargins.cxLeftWidth + popupCheckBgMargins.cxRightWidth
    measureSize.cy = drawingSize.cy + popupCheckBgMargins.cyTopHeight + popupCheckBgMargins.cyBottomHeight

    x = itemRectangle.Left

    drawingSize.cx = iconAreaSize.cx
    measureSize.cx = drawingSize.cx + popupCheckBgMargins.cxLeftWidth + popupCheckBgMargins.cxRightWidth
    measureSize.cy = drawingSize.cy + popupCheckBgMargins.cyTopHeight + popupCheckBgMargins.cyBottomHeight

    SetRect itemMetrics.rcGutter, x, y, x + popupItemMargins.cxLeftWidth + measureSize.cx, y + itemHeight

    If itemType = MenuItemTypeConstants.mitSeparator Then
      x = itemMetrics.rcGutter.Right + popupItemMargins.cxLeftWidth

      SetRect measureRectangle, x, y, itemRectangle.Right - popupItemMargins.cxRightWidth, y + separatorSize.cy
      SetRect itemMetrics.rcSeparator, measureRectangle.Left + popupItemMargins.cxLeftWidth, measureRectangle.Top + popupItemMargins.cyTopHeight, measureRectangle.Right - popupItemMargins.cxRightWidth, measureRectangle.Bottom - popupItemMargins.cyBottomHeight
      OffsetRect itemMetrics.rcSeparator, 0, (itemHeight - separatorSize.cy) / 2
    Else
      x = itemRectangle.Left + popupItemMargins.cxLeftWidth

      SetRect itemMetrics.rcSelection, x, y, itemRectangle.Right - popupItemMargins.cxRightWidth, y + itemHeight
    End If
  End If
End Sub

Private Sub MeasureItem(itemText As String, itemType As MenuItemTypeConstants, cx As Long, cy As Long)
  Const DT_CALCRECT As Long = &H400
  Const DT_LEFT As Long = &H0
  Const DT_NOPREFIX As Long = &H800
  Const DT_SINGLELINE As Long = &H20
  Const MENU_POPUPCHECKBACKGROUND As Long = 12
  Const MENU_POPUPITEM As Long = 14
  Const TMT_CONTENTMARGINS As Long = 3602
  Const WM_GETFONT As Long = &H31
  Dim iconAreaSize As Size
  Dim hDC As Long
  Dim hDCCompatible As Long
  Dim hDCToUse As Long
  Dim hOldFont As Long
  Dim hTheme As Long
  Dim popupCheckBgMargins As MARGINS
  Dim popupItemMargins As MARGINS
  Dim separatorSize As Size
  Dim textRectangle As RECT
  Dim textSize As Size

  hTheme = OpenThemeData(UserControl.hWnd, StrPtr("MENU"))
  If hTheme Then
    GetThemeMargins hTheme, 0, MENU_POPUPCHECKBACKGROUND, 0, TMT_CONTENTMARGINS, 0, popupCheckBgMargins
    GetThemeMargins hTheme, 0, MENU_POPUPITEM, 0, TMT_CONTENTMARGINS, 0, popupItemMargins

    MeasureItemParts hTheme, itemType, itemText, iconAreaSize, textSize, separatorSize
    cx = iconAreaSize.cx + popupCheckBgMargins.cxLeftWidth + popupCheckBgMargins.cxRightWidth
    If itemText <> "" Then
      cx = cx + textSize.cx
    End If
    cx = cx + popupItemMargins.cxLeftWidth + popupItemMargins.cxRightWidth
    If itemType = MenuItemTypeConstants.mitSeparator Then
      cy = separatorSize.cy
    Else
      cy = IIf(iconAreaSize.cy > textSize.cy, iconAreaSize.cy, textSize.cy)
    End If

    CloseThemeData hTheme
    Exit Sub
  End If

  If m_hMenuFont Then
    hDCCompatible = GetDC(0)
    If hDCCompatible Then
      hDC = CreateCompatibleDC(hDCCompatible)
      hDCToUse = hDC
    End If
  End If
  If hDCToUse Then
    hOldFont = SelectObject(hDCToUse, m_hMenuFont)
  Else
    hDCToUse = UserControl.hDC
    hOldFont = SelectObject(hDCToUse, SendMessage(UserControl.hWnd, WM_GETFONT, 0, 0))
  End If

  If itemText <> "" Then
    DrawText hDCToUse, StrPtr(itemText), Len(itemText), textRectangle, DT_CALCRECT Or DT_NOPREFIX Or DT_LEFT Or DT_SINGLELINE
    cx = textRectangle.Right - textRectangle.Left + 2 * HMARGIN_TEXT
    cy = textRectangle.Bottom - textRectangle.Top + 2 * VMARGIN_TEXT
  End If
  cx = cx + m_iconSize.cx + 2 * HMARGIN_ICON
  If cy < m_iconSize.cy + VMARGIN_ICON Then
    cy = m_iconSize.cy + VMARGIN_ICON
  End If

  If hOldFont Then
    SelectObject hDCToUse, hOldFont
  End If
  If hDC Then
    DeleteDC hDC
  End If
  If hDCCompatible Then
    ReleaseDC 0, hDCCompatible
  End If
End Sub

Private Sub MeasureItemParts(hTheme As Long, itemType As MenuItemTypeConstants, itemText As String, iconAreaSize As Size, textSize As Size, separatorSize As Size)
  Const DT_LEFT As Long = &H0
  Const DT_NOPREFIX As Long = &H800
  Const DT_SINGLELINE As Long = &H20
  Const MENU_POPUPBACKGROUND As Long = 9
  Const MENU_POPUPCHECK As Long = 11
  Const MENU_POPUPCHECKBACKGROUND As Long = 12
  Const MENU_POPUPITEM As Long = 14
  Const MENU_POPUPSEPARATOR As Long = 15
  Const TMT_BORDERSIZE As Long = 2403
  Const TMT_CONTENTMARGINS As Long = 3602
  Const TS_TRUE As Long = 1
  Dim drawingSize As Size
  Dim hDC As Long
  Dim hDCCompatible As Long
  Dim popupBgBorderSize As Long
  Dim popupBorderSize As Long
  Dim popupCheckBgMargins As MARGINS
  Dim popupCheckMargins As MARGINS
  Dim popupItemMargins As MARGINS
  Dim popupSeparatorSize As Size
  Dim popupTextMargins As MARGINS
  Dim textRectangle As RECT

  If hTheme Then
    hDCCompatible = GetDC(0)
    If hDCCompatible Then
      hDC = CreateCompatibleDC(hDCCompatible)
      If hDC Then
        GetThemePartSize hTheme, 0, MENU_POPUPSEPARATOR, 0, 0, TS_TRUE, popupSeparatorSize
        GetThemeInt hTheme, MENU_POPUPITEM, 0, TMT_BORDERSIZE, popupBorderSize
        GetThemeInt hTheme, MENU_POPUPBACKGROUND, 0, TMT_BORDERSIZE, popupBgBorderSize
        GetThemeMargins hTheme, 0, MENU_POPUPCHECK, 0, TMT_CONTENTMARGINS, 0, popupCheckMargins
        GetThemeMargins hTheme, 0, MENU_POPUPCHECKBACKGROUND, 0, TMT_CONTENTMARGINS, 0, popupCheckBgMargins
        GetThemeMargins hTheme, 0, MENU_POPUPITEM, 0, TMT_CONTENTMARGINS, 0, popupItemMargins

        popupTextMargins = popupItemMargins
        popupTextMargins.cxRightWidth = popupBorderSize
        popupTextMargins.cxLeftWidth = popupBgBorderSize

        drawingSize = m_iconSize
        drawingSize.cy = drawingSize.cy + popupCheckBgMargins.cyTopHeight + popupCheckBgMargins.cyBottomHeight
        iconAreaSize.cx = drawingSize.cx + popupCheckMargins.cxLeftWidth + popupCheckMargins.cxRightWidth
        iconAreaSize.cy = drawingSize.cy + popupCheckMargins.cyTopHeight + popupCheckMargins.cyBottomHeight

        If itemType = MenuItemTypeConstants.mitSeparator Then
          drawingSize = m_iconSize
          drawingSize.cy = popupSeparatorSize.cy
          separatorSize.cx = drawingSize.cx + popupItemMargins.cxLeftWidth + popupItemMargins.cxRightWidth
          separatorSize.cy = drawingSize.cy + popupItemMargins.cyTopHeight + popupItemMargins.cyBottomHeight
        ElseIf itemText <> "" Then
          GetThemeTextExtent hTheme, hDC, MENU_POPUPITEM, 0, StrPtr(itemText), Len(itemText), DT_NOPREFIX Or DT_LEFT Or DT_SINGLELINE, 0, textRectangle
          textSize.cx = textRectangle.Right - textRectangle.Left + popupTextMargins.cxLeftWidth + popupTextMargins.cxRightWidth
          textSize.cy = textRectangle.Bottom - textRectangle.Top + popupTextMargins.cyTopHeight + popupTextMargins.cyBottomHeight
        End If

        DeleteDC hDC
      End If
      ReleaseDC 0, hDCCompatible
    End If
  End If
End Sub

Private Sub QuickSort(IParentISF As IVBShellFolder, arrayToSort() As Long, ByVal lowerIndex As Long, ByVal upperIndex As Long)
  Dim i As Long
  Dim j As Long
  Dim pivot As Long
  Dim tmp As Long

  i = lowerIndex
  j = upperIndex
  pivot = arrayToSort((lowerIndex + upperIndex) / 2)

  Do
    While IsLower(IParentISF, arrayToSort(i), pivot)
      i = i + 1
    Wend
    While IsLower(IParentISF, pivot, arrayToSort(j))
      j = j - 1
    Wend

    If i <= j Then
      tmp = arrayToSort(i)
      arrayToSort(i) = arrayToSort(j)
      arrayToSort(j) = tmp
      i = i + 1
      j = j - 1
    End If
  Loop Until i > j

  If (lowerIndex < j) Then
    QuickSort IParentISF, arrayToSort, lowerIndex, j
  End If
  If (i < upperIndex) Then
    QuickSort IParentISF, arrayToSort, i, upperIndex
  End If
End Sub

Private Sub SetButtons()
  Dim buttonText As String
  Dim hasSubItems As Boolean
  Dim IShellItem As IVBShellItem
  Dim itemAttributes As SFGAOConstants
  Dim itemCount As Long
  Dim l As Long
  Dim nextID As Long
  Dim pDisplayName As Long
  Dim pIDL As Long

  With BreadcrumbBar.Buttons
    .RemoveAll
    If m_currentPIDL Then
      .Add BTNID_FOLDERICON, , GetSysIconIndex(m_currentPIDL), , "  ", DropDownStyle:=IIf(ILGetSize(m_currentPIDL) > 2, ddstNormal, ddstNoDropDown), AutoSize:=False, Width:=18

      pIDL = m_currentPIDL
      ' the final iteration will find the 2 NULL bytes
      itemCount = -1
      Do While pIDL
        itemCount = itemCount + 1
        pIDL = ILGetNext(pIDL)
      Loop
      If ILGetSize(m_currentPIDL) > 2 Then itemCount = itemCount - 1
      nextID = BTNID_FIRSTPART + itemCount
      pIDL = m_currentPIDL
      Do
        If SHCreateItemFromIDList(pIDL, IID_IShellItem, IShellItem) >= S_OK Then
          If IShellItem.GetDisplayName(SIGDN_PARENTRELATIVE, pDisplayName) >= S_OK Then
            If pDisplayName Then
              l = lstrlen(pDisplayName)
              buttonText = String(l + 1, Chr$(0))
              lstrcpyn StrPtr(buttonText), pDisplayName, l + 1
              CoTaskMemFree pDisplayName
              pDisplayName = 0
            End If
          End If

          hasSubItems = False
          If IShellItem.GetAttributes(SFGAOConstants.SFGAO_HASSUBFOLDER, itemAttributes) >= S_OK Then
            hasSubItems = ((itemAttributes And SFGAOConstants.SFGAO_HASSUBFOLDER) = SFGAOConstants.SFGAO_HASSUBFOLDER)
          End If
          Call BreadcrumbBar.Buttons.Add(nextID, 1, -2, , buttonText, True, DropDownStyle:=IIf(hasSubItems, ddstNormal, ddstNoDropDown), ButtonData:=pIDL)
          Set IShellItem = Nothing
        End If
        pIDL = ILClone(pIDL)
        ILRemoveLastID pIDL
        nextID = nextID - 1
      Loop Until ILGetSize(pIDL) = 2
      .Item(BTNID_FOLDERICON, btitID).ButtonData = pIDL

      ShowHideButtons
    End If
  End With
  RaiseEvent ChangedCurrentPIDL
End Sub

Private Sub ShowHideButtons()
  Dim i As Long
  Dim rcClient As RECT

  m_bHasHiddenItems = False
  GetClientRect UserControl.hWnd, rcClient
  With BreadcrumbBar.Buttons
    For i = 0 To .Count - 1
      If .Item(i).ID >= BTNID_FIRSTPART Then
        .Item(i).Visible = True
      End If
    Next i
    For i = 0 To .Count - 1
      If .Item(i).ID >= BTNID_FIRSTPART Then
        If BreadcrumbBar.IdealWidth < rcClient.Right - rcClient.Left Then
          .Item(i).Visible = True
        Else
          .Item(i).Visible = False
          m_bHasHiddenItems = True
        End If
      End If
    Next i
  End With
End Sub
