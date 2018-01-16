VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVisualManagerOffice2000"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  Implements IVisualManager


  Private Const LF_FACESIZE = 32

  Private Const HMARGIN_SUBMENUARROW = 16
  Private Const HMARGIN_TEXT_LEFT = 6
  Private Const HMARGIN_TEXT_RIGHT = 4
  Private Const MENUSIDEBARWIDTH = 22
  Private Const VMARGIN_ICON = 4
  Private Const VMARGIN_TEXT = 4


  Private Type POINT
    x As Long
    y As Long
  End Type

  Private Type SIZE
    cx As Long
    cy As Long
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

  Private Type IMAGELISTDRAWPARAMS
    cbSize As Long
    himl As Long
    i As Long
    hDCDst As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    xBitmap As Long
    yBitmap As Long
    rgbBk As Long
    rgbFg As Long
    fStyle As Long
    dwRop As Long
    fState As Long
    iFrame As Long
    crEffect As Long
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

  Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 256) As Byte
  End Type


  Private bIsVistaOrNewer As Boolean
  Private hCheckMarkBkgndBrush As Long
  Private hMenuFont As Long
  Private menuFontSize As Long


  Private themeableOS As Boolean


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDestination As Long, ByVal pSource As Long, ByVal Length As Long)
  Private Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal cPlanes As Long, ByVal cBitsPerPel As Long, ByVal lpvBits As Long) As Long
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
  Private Declare Function CreatePatternBrush Lib "gdi32.dll" (ByVal hBitmap As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByVal lpRect As Long, ByVal uFormat As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
  Private Declare Function ImageList_DrawIndirect Lib "comctl32.dll" (pimldp As IMAGELISTDRAWPARAMS) As Long
  Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, cx As Long, cy As Long) As Long
  Private Declare Function InflateRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iBkMode As Long) As Long
  Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
  Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal uiAction As Long, ByVal uiParam As Long, ByVal pvParam As Long, ByVal fWinIni As Long) As Long


Private Sub Class_Initialize()
  Const SPI_GETNONCLIENTMETRICS = 41
  Dim hBmp As Long
  Dim hDC_Mem As Long
  Dim hMod As Long
  Dim hPreviousBmp As Long
  Dim ncMetrics As NONCLIENTMETRICS

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  hDC_Mem = CreateCompatibleDC(0)
  If hDC_Mem Then
    hBmp = CreateBitmap(2, 2, 1, 1, 0)
    If hBmp Then
      hPreviousBmp = SelectObject(hDC_Mem, hBmp)

      SetPixelV hDC_Mem, 0, 0, vbWhite
      SetPixelV hDC_Mem, 1, 1, vbWhite
      SetPixelV hDC_Mem, 0, 1, vbBlack
      SetPixelV hDC_Mem, 1, 0, vbBlack
      hCheckMarkBkgndBrush = CreatePatternBrush(hBmp)

      SelectObject hDC_Mem, hPreviousBmp
      DeleteObject hBmp
    End If
    DeleteDC hDC_Mem
  End If

  ncMetrics.cbSize = LenB(ncMetrics) - IIf(bIsVistaOrNewer, 0, LenB(ncMetrics.iPaddedBorderWidth))
  If SystemParametersInfo(SPI_GETNONCLIENTMETRICS, ncMetrics.cbSize, VarPtr(ncMetrics), 0) Then
    menuFontSize = ncMetrics.lfMenuFont.lfHeight
    hMenuFont = CreateFontIndirect(ncMetrics.lfMenuFont)
  End If
End Sub

Private Sub Class_Terminate()
  If hMenuFont Then
    DeleteObject hMenuFont
    hMenuFont = 0
  End If
  If hCheckMarkBkgndBrush Then
    DeleteObject hCheckMarkBkgndBrush
    hCheckMarkBkgndBrush = 0
  End If
End Sub

Private Sub IVisualManager_InitializeReBar(ByVal ctl As IReBar)
  If themeableOS Then
    SetWindowTheme ctl.hWnd, StrPtr(""), StrPtr("")
  End If
End Sub

Private Sub IVisualManager_InitializeToolBar(ByVal ctl As IToolBar)
  If themeableOS Then
    SetWindowTheme ctl.hWnd, StrPtr(""), StrPtr("")
  End If
End Sub

Private Sub IVisualManager_ReBarCustomDraw(ByVal ctl As IReBar, ByVal band As IReBarBand, ByVal drawStage As CustomDrawStageConstants, ByVal bandState As CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As RECTANGLE, furtherProcessing As CustomDrawReturnValuesConstants)
  '
End Sub

Private Sub IVisualManager_ReBarRawMenuMessage(ByVal ctl As IReBar, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  Const WM_DRAWITEM = &H2B
  Const WM_MEASUREITEM = &H2C

  If message = WM_MEASUREITEM Then
    MeasureItem message, wParam, lParam, result, handledEvent
  ElseIf message = WM_DRAWITEM Then
    DrawItem message, wParam, lParam, result, handledEvent
  End If
End Sub

Private Sub IVisualManager_SetMenuManager(ByVal manager As IMenuManager)
  '
End Sub

Private Sub IVisualManager_ToolBarBeforeDisplayChevronPopup(ByVal ctl As IToolBar, ByVal hPopup As Long, ByVal x As Single, ByVal y As Single, ByVal isMenu As Boolean, cancelPopup As Boolean, commandToExecute As Long)
  If themeableOS And Not isMenu Then
    SetWindowTheme hPopup, StrPtr(""), StrPtr("")
  End If
End Sub

Private Sub IVisualManager_ToolBarCustomDraw(ByVal ctl As ToolBar, ByVal toolButton As ToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As StringBackgroundModeConstants, ByVal drawStage As CustomDrawStageConstants, ByVal buttonState As CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As RECTANGLE, textRectangle As RECTANGLE, HorizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As CustomDrawReturnValuesConstants)
  '
End Sub

Private Sub IVisualManager_ToolBarDestroyingChevronPopup(ByVal ctl As IToolBar, ByVal hPopup As Long, ByVal isMenu As Boolean)
  '
End Sub

Private Sub IVisualManager_ToolBarDropDown(ByVal ctl As IToolBar, ByVal toolButton As IToolBarButton, buttonRectangle As RECTANGLE, furtherProcessing As DropDownReturnValuesConstants)
  '
End Sub

Private Sub IVisualManager_ToolBarRawMenuMessage(ByVal ctl As TBarCtlsLibUCtl.IToolBar, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  Const WM_DRAWITEM = &H2B
  Const WM_MEASUREITEM = &H2C

  If message = WM_MEASUREITEM Then
    MeasureItem message, wParam, lParam, result, handledEvent
  ElseIf message = WM_DRAWITEM Then
    DrawItem message, wParam, lParam, result, handledEvent
  End If
End Sub

Private Sub DrawItem(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  Const BDR_RAISEDINNER = &H4
  Const BDR_RAISEDOUTER = &H1
  Const BDR_SUNKENINNER = &H8
  Const BDR_SUNKENOUTER = &H2
  Const BF_BOTTOM = &H8
  Const BF_LEFT = &H1
  Const BF_RIGHT = &H4
  Const BF_TOP = &H2
  Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
  Const COLOR_BTNFACE = 15
  Const COLOR_BTNHIGHLIGHT = 20
  Const COLOR_GRAYTEXT = 17
  Const COLOR_HIGHLIGHTTEXT = 14
  Const COLOR_MENUHILIGHT = 29
  Const COLOR_MENUTEXT = 7
  Const DT_EXPANDTABS = &H40
  Const DT_HIDEPREFIX = &H100000
  Const DT_LEFT = &H0
  Const DT_NOPREFIX = &H800
  Const DT_RIGHT = &H2
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const ILD_TRANSPARENT = &H1
  Const ILS_SATURATE = &H4
  Const ODS_CHECKED = &H8
  Const ODS_GRAYED = &H2
  Const ODS_HOTLIGHT = &H40
  Const ODS_INACTIVE = &H80
  Const ODS_NOACCEL = &H100
  Const ODS_SELECTED = &H1
  Const TRANSPARENT = 1
  Dim displayKeyboardCues As Boolean
  Dim drawItemData As DRAWITEMSTRUCT
  Dim drawParams As IMAGELISTDRAWPARAMS
  Dim hDC As Long
  Dim hPreviousFont As Long
  Dim iconRectangle As RECT
  Dim iconSize As SIZE
  Dim itemText As String
  Dim positionOfTab As Long
  Dim previousBkColor As Long
  Dim previousBkMode As Long
  Dim previousTextColor As Long
  Dim rc As RECT
  Dim tagObject As CMenuItemData
  Dim textRectangle As RECT

  CopyMemory VarPtr(drawItemData), lParam, LenB(drawItemData)

  If hMenuFont Then
    hDC = drawItemData.hDC
    If hDC Then
      hPreviousFont = SelectObject(hDC, hMenuFont)

      iconSize.cx = 16
      iconSize.cy = 16
      Set tagObject = ExtractMenuTag(drawItemData.itemData)
      If Not tagObject Is Nothing Then
        If tagObject.hImageList <> 0 And tagObject.IconIndex >= 0 Then
          ImageList_GetIconSize tagObject.hImageList, iconSize.cx, iconSize.cy
        End If
      End If

      LSet rc = drawItemData.rcItem
      FillRect hDC, rc, GetSysColorBrush(COLOR_BTNFACE)
      If drawItemData.itemState And ODS_CHECKED Then
        rc.Left = rc.Left + MENUSIDEBARWIDTH
      ElseIf Not tagObject Is Nothing Then
        If tagObject.hImageList <> 0 And tagObject.IconIndex >= 0 Then
          rc.Left = rc.Left + MENUSIDEBARWIDTH
        End If
      End If
      If (drawItemData.itemState And (ODS_SELECTED Or ODS_GRAYED)) = ODS_SELECTED Then
        FillRect hDC, rc, GetSysColorBrush(COLOR_MENUHILIGHT)
      End If

      LSet iconRectangle = drawItemData.rcItem
      iconRectangle.Right = iconRectangle.Left + drawItemData.rcItem.Bottom - drawItemData.rcItem.Top + 1
      LSet textRectangle = drawItemData.rcItem
      textRectangle.Left = iconRectangle.Left + MENUSIDEBARWIDTH + HMARGIN_TEXT_LEFT

      If drawItemData.itemState And ODS_CHECKED Then
        DrawEdge hDC, iconRectangle, BDR_SUNKENOUTER, BF_RECT
      ElseIf (drawItemData.itemState And (ODS_SELECTED Or ODS_GRAYED Or ODS_CHECKED)) = ODS_SELECTED Then
        If Not tagObject Is Nothing Then
          If tagObject.hImageList <> 0 And tagObject.IconIndex >= 0 Then
            DrawEdge hDC, iconRectangle, BDR_RAISEDINNER, BF_RECT
          End If
        End If
      End If

      If Not tagObject Is Nothing Then
        If tagObject.Separator Then
          LSet rc = drawItemData.rcItem
          rc.Top = rc.Top + (drawItemData.rcItem.Bottom - drawItemData.rcItem.Top) / 2 - 1
          rc.Bottom = rc.Top + 2
          DrawEdge hDC, rc, BDR_SUNKENOUTER, BF_RECT
        Else
          itemText = tagObject.Text
          If itemText <> "" Then
            previousBkMode = SetBkMode(hDC, TRANSPARENT)
            If drawItemData.itemState And ODS_GRAYED Then
              previousTextColor = SetTextColor(hDC, GetSysColor(COLOR_GRAYTEXT))
            ElseIf drawItemData.itemState And ODS_SELECTED Then
              previousTextColor = SetTextColor(hDC, GetSysColor(COLOR_HIGHLIGHTTEXT))
            Else
              previousTextColor = SetTextColor(hDC, GetSysColor(COLOR_MENUTEXT))
            End If

            displayKeyboardCues = ((drawItemData.itemState And ODS_NOACCEL = 0))
            positionOfTab = InStrRev(itemText, vbTab)
            If positionOfTab > 0 Then
              Mid(itemText, positionOfTab, Len(vbTab)) = Chr$(0)
            End If
            DrawText hDC, StrPtr(itemText), -1, VarPtr(textRectangle), DT_EXPANDTABS Or DT_VCENTER Or DT_LEFT Or DT_SINGLELINE Or IIf(displayKeyboardCues, 0, DT_HIDEPREFIX)
            If positionOfTab > 0 Then
              itemText = Mid$(itemText, positionOfTab + 1)
              textRectangle.Right = textRectangle.Right - HMARGIN_SUBMENUARROW - HMARGIN_TEXT_RIGHT
              DrawText hDC, StrPtr(itemText), -1, VarPtr(textRectangle), DT_VCENTER Or DT_RIGHT Or DT_SINGLELINE Or DT_NOPREFIX
            End If

            SetTextColor hDC, previousTextColor
            SetBkMode hDC, previousBkMode
          End If
        End If

        If (drawItemData.itemState And (ODS_CHECKED Or ODS_GRAYED Or ODS_SELECTED)) = ODS_CHECKED Then
          InflateRect iconRectangle, -2, -2
          previousBkColor = SetBkColor(hDC, GetSysColor(COLOR_BTNFACE))
          previousTextColor = SetTextColor(hDC, GetSysColor(COLOR_BTNHIGHLIGHT))

          FillRect hDC, iconRectangle, hCheckMarkBkgndBrush

          SetTextColor hDC, previousTextColor
          SetBkColor hDC, previousBkColor
        End If

        If tagObject.hImageList <> 0 And tagObject.IconIndex >= 0 Then
          With drawParams
            .cbSize = LenB(drawParams)
            .himl = tagObject.hImageList
            .i = tagObject.IconIndex
            .hDCDst = hDC
            .x = iconRectangle.Left + (iconRectangle.Right - iconRectangle.Left - iconSize.cx) / 2
            .y = iconRectangle.Top + (iconRectangle.Bottom - iconRectangle.Top - iconSize.cy) / 2
            .fStyle = ILD_TRANSPARENT
            If drawItemData.itemState And ODS_GRAYED Then
              .fState = ILS_SATURATE
            End If
          End With
          ImageList_DrawIndirect drawParams
        ElseIf drawItemData.itemState And ODS_CHECKED Then
          DrawCheckMark menuFontSize, hDC, iconRectangle, (drawItemData.itemState And ODS_GRAYED) = 0, tagObject.RadioCheckMark
        End If
      End If

      SelectObject hDC, hPreviousFont
    End If
  End If

  result = 1
  handledEvent = True
End Sub

Private Sub MeasureItem(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  Const DT_CALCRECT = &H400
  Const DT_EXPANDTABS = &H40
  Const DT_LEFT = &H0
  Const DT_SINGLELINE = &H20
  Dim hDC As Long
  Dim hPreviousFont As Long
  Dim iconSize As SIZE
  Dim itemText As String
  Dim measureItemData As MEASUREITEMSTRUCT
  Dim tagObject As CMenuItemData
  Dim textRectangle As RECT

  CopyMemory VarPtr(measureItemData), lParam, LenB(measureItemData)

  If hMenuFont Then
    hDC = CreateCompatibleDC(0)
    If hDC Then
      hPreviousFont = SelectObject(hDC, hMenuFont)

      iconSize.cx = 16
      iconSize.cy = 16
      Set tagObject = ExtractMenuTag(measureItemData.itemData)
      If Not tagObject Is Nothing Then
        If tagObject.Separator Then
          measureItemData.itemWidth = 10
          measureItemData.itemHeight = 6
        Else
          itemText = tagObject.Text
          If itemText <> "" Then
            DrawText hDC, StrPtr(itemText), -1, VarPtr(textRectangle), DT_CALCRECT Or DT_EXPANDTABS Or DT_LEFT Or DT_SINGLELINE
            measureItemData.itemWidth = textRectangle.Right - textRectangle.Left + HMARGIN_TEXT_LEFT + HMARGIN_TEXT_RIGHT + HMARGIN_SUBMENUARROW
            measureItemData.itemHeight = textRectangle.Bottom - textRectangle.Top + VMARGIN_TEXT
          End If
          If tagObject.hImageList <> 0 And tagObject.IconIndex >= 0 Then
            ImageList_GetIconSize tagObject.hImageList, iconSize.cx, iconSize.cy
          End If
        End If
      End If

      measureItemData.itemWidth = measureItemData.itemWidth + MENUSIDEBARWIDTH
      If measureItemData.itemHeight < iconSize.cy + VMARGIN_ICON Then
        If Not tagObject.Separator Then
          measureItemData.itemHeight = iconSize.cy + VMARGIN_ICON
        End If
      End If

      SelectObject hDC, hPreviousFont
      DeleteDC hDC
    End If
  End If

  CopyMemory lParam, VarPtr(measureItemData), LenB(measureItemData)
  result = 1
  handledEvent = True
End Sub
