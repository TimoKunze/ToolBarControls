VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVisualManagerOfficeXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  Implements IHook
  Implements ISubclassedWindow
  Implements IVisualManager


  Private Const LF_FACESIZE = 32

  Private Const HMARGIN_SUBMENUARROW = 16
  Private Const HMARGIN_TEXT_LEFT = 6
  Private Const HMARGIN_TEXT_RIGHT = 4
  Private Const MENUSIDEBARWIDTH = 23
  Private Const VMARGIN_ICON = 4
  Private Const VMARGIN_TEXT = 4


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

  Private Type WINDOWPOS
    hWnd As Long
    hWndInsertAfter As Long
    x As Long
    y As Long
    cx As Long
    cy As Long
    flags As Long
  End Type


  Private bHasHandledInitialWindowPosChanging As Boolean
  Private bIsVistaOrNewer As Boolean
  Private buttonBorderColor As Long
  Private checkedButtonBackColor As Long
  Private droppedDownButtonFaceColor As Long
  Private hMenuFont As Long
  Private menuFontSize As Long
  Private menuHierarchyDepth As Long
  Private menuHierarchyDepthCorrection As Long
  Private menuLightColor As Long
  Private menuManager As IMenuManager
  Private pressedButtonBackColor As Long
  Private rebarBandGripperColor As Long
  Private separatorColor As Long
  Private subclassedMenuWindows As Collection


  Private lastDroppedButtonWidth As Long
  Private themeableOS As Boolean


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDestination As Long, ByVal pSource As Long, ByVal Length As Long)
  Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function CreateFontIndirect Lib "gdi32.dll" Alias "CreateFontIndirectW" (ByRef lpLogFont As LOGFONT) As Long
  Private Declare Function CreatePen Lib "gdi32.dll" (ByVal fnPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
  Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
  Private Declare Function DrawRectangle Lib "gdi32.dll" Alias "Rectangle" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByVal lpRect As Long, ByVal uFormat As Long) As Long
  Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameW" (ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
  Private Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECT) As Long
  Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
  Private Declare Function ImageList_DrawIndirect Lib "comctl32.dll" (pimldp As IMAGELISTDRAWPARAMS) As Long
  Private Declare Function ImageList_GetIconSize Lib "comctl32.dll" (ByVal himl As Long, cx As Long, cy As Long) As Long
  Private Declare Function InflateRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function LineTo Lib "gdi32.dll" (ByVal hDC As Long, ByVal nXEnd As Long, ByVal nYEnd As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpPoint As Long) As Long
  Private Declare Function OffsetRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iBkMode As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
  Private Declare Function SystemParametersInfo Lib "user32.dll" Alias "SystemParametersInfoW" (ByVal uiAction As Long, ByVal uiParam As Long, ByVal pvParam As Long, ByVal fWinIni As Long) As Long


Private Sub Class_Initialize()
  Const COLOR_BTNFACE = 15
  Const COLOR_BTNSHADOW = 16
  Const COLOR_HIGHLIGHT = 13
  Const COLOR_SCROLLBAR = 0
  Const COLOR_WINDOW = 5
  Const SPI_GETNONCLIENTMETRICS = 41
  Dim hMod As Long
  Dim faceColor As Long
  Dim highlightColor As Long
  Dim ncMetrics As NONCLIENTMETRICS
  Dim windowColor As Long

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  windowColor = GetSysColor(COLOR_WINDOW)
  highlightColor = GetSysColor(COLOR_HIGHLIGHT)
  faceColor = GetSysColor(COLOR_BTNFACE)

  checkedButtonBackColor = BlendColor(highlightColor, windowColor, 50)
  buttonBorderColor = highlightColor
  droppedDownButtonFaceColor = BlendColor(windowColor, faceColor, 40)
  menuLightColor = BlendColor(windowColor, faceColor, 219)
  pressedButtonBackColor = BlendColor(highlightColor, menuLightColor, 70)
  rebarBandGripperColor = BlendColor(GetSysColor(COLOR_BTNSHADOW), faceColor, 205)
  separatorColor = GetSysColor(COLOR_SCROLLBAR)

  ncMetrics.cbSize = LenB(ncMetrics) - IIf(bIsVistaOrNewer, 0, LenB(ncMetrics.iPaddedBorderWidth))
  If SystemParametersInfo(SPI_GETNONCLIENTMETRICS, ncMetrics.cbSize, VarPtr(ncMetrics), 0) Then
    menuFontSize = ncMetrics.lfMenuFont.lfHeight
    hMenuFont = CreateFontIndirect(ncMetrics.lfMenuFont)
  End If

  lastDroppedButtonWidth = 0
  Set subclassedMenuWindows = New Collection

  ' to make drop-down menus appear "connected" to the tool bar button, we need to install a CallWndProc hook and subclass
  ' the menu window
  InstallCallWndProcHook Me
End Sub

Private Sub Class_Terminate()
  RemoveCallWndProcHook Me
  If hMenuFont Then
    DeleteObject hMenuFont
    hMenuFont = 0
  End If
  Set subclassedMenuWindows = Nothing
End Sub

Private Function IHook_CallWndProcAfter(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_CallWndProcBefore(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long, eatIt As Boolean) As Long
  Const WM_CREATE = &H1
  Const WM_DESTROY = &H2
  Dim className As String
  Dim cwpData As CWPSTRUCT
  Dim i As Long

  CopyMemory VarPtr(cwpData), lParam, LenB(cwpData)
  If cwpData.message = WM_CREATE Then
    className = String$(7, 0)
    GetClassName cwpData.hWnd, StrPtr(className), 7
    If Left$(className, 6) = "#32768" Then
      ' this is a menu window, so subclass it
      menuHierarchyDepth = menuHierarchyDepth + 1
      If Not SubclassWindow(cwpData.hWnd, Me, EnumSubclassID.escidMenu) Then
        Debug.Print "Subclassing failed!"
      End If
      subclassedMenuWindows.Add cwpData.hWnd
    End If
  ElseIf cwpData.message = WM_DESTROY Then
    className = String$(7, 0)
    GetClassName cwpData.hWnd, StrPtr(className), 7
    If Left$(className, 6) = "#32768" Then
      ' this is a menu window, so unsubclass it if it has been subclassed
      If Not subclassedMenuWindows Is Nothing Then
        For i = 1 To subclassedMenuWindows.Count
          If subclassedMenuWindows.Item(i) = cwpData.hWnd Then
            If Not UnSubclassWindow(cwpData.hWnd, EnumSubclassID.escidMenu) Then
              Debug.Print "UnSubclassing failed!"
            End If
            subclassedMenuWindows.Remove i
            menuHierarchyDepth = menuHierarchyDepth - 1
            If menuHierarchyDepth = 0 Then
              lastDroppedButtonWidth = 0
            End If
            Exit For
          End If
        Next i
      End If
    End If
  End If
End Function

Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidMenu
      lRet = HandleMessage_Menu(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "CVisualManagerOfficeXP.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in CVisualManagerOfficeXP.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Menu(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const MN_GETHMENU = &H1E1
  Const PRF_NONCLIENT = &H2
  Const SWP_SHOWWINDOW = &H40
  Const WM_NCPAINT = &H85
  Const WM_PAINT = &HF
  Const WM_PRINT = &H317
  Const WM_WINDOWPOSCHANGING = &H46
  Dim hDC As Long
  Dim hMenu As Long
  Dim i As Long
  Dim wndPos As WINDOWPOS
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  If Not (menuManager Is Nothing Or subclassedMenuWindows Is Nothing) Then
    If uMsg <> MN_GETHMENU Then
      hMenu = SendMessageAsLong(hWnd, MN_GETHMENU, 0, 0)
      If hMenu <> 0 Then
        If menuManager.IsSkinnedMenu(hMenu) Then
          Select Case uMsg
            Case WM_NCPAINT, WM_PAINT
              lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
              bCallDefProc = False
              hDC = GetWindowDC(hWnd)
              If hDC Then
                ' draw menu window border
                DrawMenuFrame hWnd, hDC, (menuHierarchyDepth + menuHierarchyDepthCorrection = 1)
                ReleaseDC hWnd, hDC
              End If

            Case WM_PRINT
              If bHasHandledInitialWindowPosChanging Then
                If lParam And PRF_NONCLIENT Then
                  lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
                  bCallDefProc = False

                  hDC = wParam
                  If hDC Then
                    ' draw menu window border
                    DrawMenuFrame hWnd, hDC, (menuHierarchyDepth + menuHierarchyDepthCorrection = 1)
                  End If
                End If
              End If

            Case WM_WINDOWPOSCHANGING
              If Not bHasHandledInitialWindowPosChanging Then
                CopyMemory VarPtr(wndPos), lParam, LenB(wndPos)
                If wndPos.flags And SWP_SHOWWINDOW Then
                  bHasHandledInitialWindowPosChanging = True
                  lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
                  bCallDefProc = False
                End If
              End If
          End Select
        End If
      End If
    End If
  End If

StdHandler_Ende:
  HandleMessage_Menu = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in CVisualManagerOfficeXP.HandleMessage_Menu: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Sub IVisualManager_InitializeReBar(ByVal ctl As IReBar)
  If themeableOS Then
    SetWindowTheme ctl.hWnd, 0, 0
  End If
End Sub

Private Sub IVisualManager_InitializeToolBar(ByVal ctl As IToolBar)
  If themeableOS Then
    SetWindowTheme ctl.hWnd, StrPtr(""), StrPtr("")
  End If
End Sub

Private Sub IVisualManager_ReBarCustomDraw(ByVal ctl As IReBar, ByVal band As IReBarBand, ByVal drawStage As CustomDrawStageConstants, ByVal bandState As CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As RECTANGLE, furtherProcessing As CustomDrawReturnValuesConstants)
  Const COLOR_BTNFACE = 15
  Const PS_SOLID = 0
  Dim chevronRect As RECT
  Dim hBrush As Long
  Dim hitTestDetails As HitTestConstants
  Dim hPen As Long
  Dim hPreviousBrush As Long
  Dim hPreviousPen As Long
  Dim rc As RECT
  Dim x As Long
  Dim y As Long

  Select Case drawStage
    Case cdsPreErase, cdsPrePaint
      LSet rc = drawingRectangle
      FillRect hDC, rc, GetSysColorBrush(COLOR_BTNFACE)

      For Each band In ctl.Bands
        band.GetRectangle brtBand, rc.Left, rc.Top, rc.Right, rc.Bottom

        If band.ChevronVisible Then
          band.GetChevronRectangle chevronRect.Left, chevronRect.Top, chevronRect.Right, chevronRect.Bottom
          chevronRect.Bottom = chevronRect.Bottom - 1

          If band.ChevronHot Or band.ChevronPushed Then
            hBrush = CreateSolidBrush(pressedButtonBackColor)
            If hBrush Then
              hPreviousBrush = SelectObject(hDC, hBrush)
              hPen = CreatePen(PS_SOLID, 1, buttonBorderColor)
              If hPen Then
                hPreviousPen = SelectObject(hDC, hPen)
                DrawRectangle hDC, chevronRect.Left, chevronRect.Top, chevronRect.Right, chevronRect.Bottom

                SelectObject hDC, hPreviousPen
                DeleteObject hPen
              End If
              SelectObject hDC, hPreviousBrush
              DeleteObject hBrush
            End If
          End If
          DrawChevron hDC, chevronRect
        End If

        ' draw gripper
        hPen = CreatePen(PS_SOLID, 1, rebarBandGripperColor)
        If hPen Then
          hPreviousPen = SelectObject(hDC, hPen)

          x = rc.Left + 1
          y = rc.Top + ((rc.Bottom - rc.Top - 12) / 2)
          For y = y To y + 12 Step 2
            MoveToEx hDC, x, y, 0
            LineTo hDC, x + 3, y
          Next y

          SelectObject hDC, hPreviousPen
          DeleteObject hPen
        End If
      Next band
      furtherProcessing = cdrvSkipDefault
  End Select
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
  Set menuManager = manager
End Sub

Private Sub IVisualManager_ToolBarBeforeDisplayChevronPopup(ByVal ctl As TBarCtlsLibUCtl.IToolBar, ByVal hPopup As Long, ByVal x As Single, ByVal y As Single, ByVal isMenu As Boolean, cancelPopup As Boolean, commandToExecute As Long)
  If isMenu Then
    ' The chevron popup is a menu without parent, so usually its frame would be drawn integrated with the tool bar button.
    ' To avoid this we manipulate the menu hierarchy depth counter to effectively start at 2 instead of 1.
    menuHierarchyDepthCorrection = 1
  Else
    ' The chevron popup is a menu window and therefore increases the menu hierarchy counter. We still want the drop-down
    ' menus of tool bar buttons, that are displayed in the chevron popup, to integrate with the tool bar buttons.
    ' To achieve this we manipulate the menu hierarchy depth counter to effectively start at 0 instead of 1.
    menuHierarchyDepthCorrection = -1
  End If
End Sub

Private Sub IVisualManager_ToolBarCustomDraw(ByVal ctl As ToolBar, ByVal toolButton As ToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As StringBackgroundModeConstants, ByVal drawStage As CustomDrawStageConstants, ByVal buttonState As CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As RECTANGLE, textRectangle As RECTANGLE, HorizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As CustomDrawReturnValuesConstants)
  Const BDR_RAISEDOUTER = &H1
  Const BDR_RAISEDINNER = &H4
  Const BDR_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
  Const BF_FLAT = &H4000
  Const BF_LEFT = &H1
  Const BF_RIGHT = &H4
  Const BF_TOP = &H2
  Const COLOR_BTNFACE = 15
  Const COLOR_BTNHIGHLIGHT = 20
  Const COLOR_BTNTEXT = 18
  Const DT_CENTER = &H1
  Const DT_HIDEPREFIX = &H100000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const ILD_TRANSPARENT = &H1
  Const ILS_SATURATE = &H4
  Const PS_SOLID = 0
  Const TRANSPARENT = 1
  Const UISF_HIDEACCEL = &H2
  Const WM_GETFONT = &H31
  Const WM_QUERYUISTATE = &H129
  Dim displayKeyboardCues As Boolean
  Dim drawParams As IMAGELISTDRAWPARAMS
  Dim dropDownRectangle As RECT
  Dim grayScale As Boolean
  Dim hBrush As Long
  Dim hFont As Long
  Dim hImgLst As Long
  Dim hPen As Long
  Dim hPreviousBrush As Long
  Dim hPreviousFont As Long
  Dim hPreviousPen As Long
  Dim icon As Long
  Dim imgLstIndex As Long
  Dim imgHeight As Long
  Dim imgWidth As Long
  Dim isDroppedDown As Boolean
  Dim previousBkMode As Long
  Dim previousTextColor As Long
  Dim rc As RECT
  Dim sep As ToolBarButton
  Dim txt As String
  Dim x As Long
  Dim y As Long

  Select Case drawStage
    Case cdsPrePaint, cdsPreErase
      ' draw the tool bar's background
      GetClientRect ctl.hWnd, rc
      'LSet rc = drawingRectangle
      FillRect hDC, rc, GetSysColorBrush(COLOR_BTNFACE)
      furtherProcessing = cdrvNotifyItemDraw Or cdrvNotifyPostPaint

    Case cdsChevronPopupToolbar Or cdsPrePaint, cdsChevronPopupToolbar Or cdsPreErase
      ' draw the chevron popup tool bar's background
      GetClientRect ctl.hWndChevronToolBar, rc
      hBrush = CreateSolidBrush(droppedDownButtonFaceColor)
      If hBrush Then
        FillRect hDC, rc, hBrush
        DeleteObject hBrush
      End If
      furtherProcessing = cdrvNotifyItemDraw Or cdrvNotifyPostPaint

    Case cdsItemPrePaint, cdsChevronPopupToolbar Or cdsItemPrePaint
      ' draw the tool bar button
      isDroppedDown = toolButton.DroppedDown
      If isDroppedDown Then
        lastDroppedButtonWidth = toolButton.Width
        LSet rc = drawingRectangle
        hBrush = CreateSolidBrush(droppedDownButtonFaceColor)
        If hBrush Then
          FillRect hDC, rc, hBrush
          DeleteObject hBrush
        End If
        DrawEdge hDC, rc, BDR_RAISED, BF_LEFT Or BF_RIGHT Or BF_TOP Or BF_FLAT
      End If
      If Not isDroppedDown And (buttonState And (cdisChecked Or cdisHot)) Then
        If Not toolButton Is Nothing Then
          hBrush = CreateSolidBrush(IIf(buttonState And cdisHot, pressedButtonBackColor, checkedButtonBackColor))
          If hBrush Then
            hPreviousBrush = SelectObject(hDC, hBrush)
            hPen = CreatePen(PS_SOLID, 1, buttonBorderColor)
            If hPen Then
              hPreviousPen = SelectObject(hDC, hPen)
              DrawRectangle hDC, drawingRectangle.Left, drawingRectangle.Top, drawingRectangle.Right, drawingRectangle.Bottom - 1
              If (toolButton.DropDownStyle = ddstNormal) And (ctl.NormalDropDownButtonStyle = nddbsSplitButton) Then
                toolButton.GetRectangle brtDropDown, dropDownRectangle.Left, dropDownRectangle.Top, dropDownRectangle.Right, dropDownRectangle.Bottom
                DrawRectangle hDC, dropDownRectangle.Left, dropDownRectangle.Top, dropDownRectangle.Right, dropDownRectangle.Bottom - 1

                ' draw arrow
                DrawDropDownArrow hDC, dropDownRectangle
              Else
                dropDownRectangle.Left = drawingRectangle.Right
              End If

              icon = toolButton.IconIndex
              If icon >= 0 Then
                imgLstIndex = toolButton.ImageListIndex
                hImgLst = ctl.hImageList(ilHotButtons, imgLstIndex)
              End If

              SelectObject hDC, hPreviousPen
              DeleteObject hPen
            End If
            SelectObject hDC, hPreviousBrush
            DeleteObject hBrush
          End If
        End If

      ElseIf Not isDroppedDown And (buttonState And cdisDisabled) Then
        If Not toolButton Is Nothing Then
          If (toolButton.DropDownStyle = ddstNormal) And (ctl.NormalDropDownButtonStyle = nddbsSplitButton) Then
            toolButton.GetRectangle brtDropDown, dropDownRectangle.Left, dropDownRectangle.Top, dropDownRectangle.Right, dropDownRectangle.Bottom

            ' draw arrow
            DrawDropDownArrow hDC, dropDownRectangle, False
          Else
            dropDownRectangle.Left = drawingRectangle.Right
          End If

          icon = toolButton.IconIndex
          If icon >= 0 Then
            imgLstIndex = toolButton.ImageListIndex
            hImgLst = ctl.hImageList(ilDisabledButtons, imgLstIndex)
          End If
        End If

      Else
        If Not toolButton Is Nothing Then
          If (toolButton.DropDownStyle = ddstNormal) And (ctl.NormalDropDownButtonStyle = nddbsSplitButton) Then
            toolButton.GetRectangle brtDropDown, dropDownRectangle.Left, dropDownRectangle.Top, dropDownRectangle.Right, dropDownRectangle.Bottom

            ' draw arrow
            DrawDropDownArrow hDC, dropDownRectangle
          Else
            dropDownRectangle.Left = drawingRectangle.Right
          End If

          icon = toolButton.IconIndex
          If icon >= 0 Then
            imgLstIndex = toolButton.ImageListIndex
          End If
        End If
      End If

      If hImgLst = 0 Then
        hImgLst = ctl.hImageList(ilNormalButtons, imgLstIndex)
        grayScale = (buttonState And cdisDisabled)
      End If
      If hImgLst Then
        ImageList_GetIconSize hImgLst, imgWidth, imgHeight
        x = drawingRectangle.Left + (dropDownRectangle.Left - drawingRectangle.Left - imgWidth) / 2
        y = drawingRectangle.Top + (drawingRectangle.Bottom - drawingRectangle.Top - imgHeight) / 2
        With drawParams
          .cbSize = LenB(drawParams)
          .himl = hImgLst
          .i = icon
          .hDCDst = hDC
          .x = x
          .y = y
          .fStyle = ILD_TRANSPARENT
          If grayScale Then
            .fState = ILS_SATURATE
          End If
        End With
        ImageList_DrawIndirect drawParams
      End If

      txt = toolButton.Text
      If txt <> "" Then
        LSet rc = drawingRectangle
        rc.Left = rc.Left + textRectangle.Left
        rc.Right = rc.Left + textRectangle.Right - textRectangle.Left

        hPreviousFont = SelectObject(hDC, SendMessageAsLong(ctl.hWnd, WM_GETFONT, 0, 0))
        previousBkMode = SetBkMode(hDC, TRANSPARENT)
        previousTextColor = SetTextColor(hDC, GetSysColor(IIf(buttonState And cdisDisabled, COLOR_BTNHIGHLIGHT, COLOR_BTNTEXT)))

        displayKeyboardCues = ((SendMessageAsLong(ctl.hWnd, WM_QUERYUISTATE, 0, 0) And UISF_HIDEACCEL) = 0)
        DrawText hDC, StrPtr(txt), -1, VarPtr(rc), DT_VCENTER Or DT_CENTER Or DT_SINGLELINE Or IIf(displayKeyboardCues, 0, DT_HIDEPREFIX)

        SetTextColor hDC, previousTextColor
        SetBkMode hDC, previousBkMode
        hPreviousFont = SelectObject(hDC, hPreviousFont)
      End If

      ' unfortunatly Windows ignores cdrvDontDrawButtonEdges for the drop-down portion of a button
      ' therefore we draw the whole button ourselves
      'furtherProcessing = cdrvDontDrawBackground Or cdrvDontDrawButtonEdges
      furtherProcessing = cdrvSkipDefault

    Case cdsPostPaint
      ' draw the separators - but not for the chevron popup tool bar
      For Each sep In ctl.Buttons
        If sep.ButtonType = btySeparator Then
          sep.GetRectangle brtButton, rc.Left, rc.Top, rc.Right, rc.Bottom
          FillRect hDC, rc, GetSysColorBrush(COLOR_BTNFACE)
          hPen = CreatePen(PS_SOLID, 1, separatorColor)
          If hPen Then
            hPreviousPen = SelectObject(hDC, hPen)

            MoveToEx hDC, CLng((rc.Left + rc.Right) / 2), CLng((rc.Top + rc.Bottom) / 2) - 10, 0
            LineTo hDC, CLng((rc.Left + rc.Right) / 2), CLng((rc.Top + rc.Bottom) / 2) + 10

            SelectObject hDC, hPreviousPen
            DeleteObject hPen
          End If
        End If
      Next sep
  End Select
End Sub

Private Sub IVisualManager_ToolBarDestroyingChevronPopup(ByVal ctl As IToolBar, ByVal hPopup As Long, ByVal isMenu As Boolean)
  menuHierarchyDepthCorrection = 0
End Sub

Private Sub IVisualManager_ToolBarDropDown(ByVal ctl As IToolBar, ByVal toolButton As IToolBarButton, buttonRectangle As RECTANGLE, furtherProcessing As DropDownReturnValuesConstants)
  lastDroppedButtonWidth = buttonRectangle.Right - buttonRectangle.Left
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
  Const COLOR_GRAYTEXT = 17
  Const COLOR_MENU = 4
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
  Const PS_SOLID = 0
  Const TRANSPARENT = 1
  Dim displayKeyboardCues As Boolean
  Dim drawItemData As DRAWITEMSTRUCT
  Dim drawParams As IMAGELISTDRAWPARAMS
  Dim hBrush As Long
  Dim hDC As Long
  Dim hPen As Long
  Dim hPreviousBrush As Long
  Dim hPreviousFont As Long
  Dim hPreviousPen As Long
  Dim iconRectangle As RECT
  Dim iconSize As SIZE
  Dim itemText As String
  Dim positionOfTab As Long
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

      If (drawItemData.itemState And (ODS_SELECTED Or ODS_GRAYED)) = ODS_SELECTED Then
        hBrush = CreateSolidBrush(pressedButtonBackColor)
        If hBrush Then
          hPreviousBrush = SelectObject(hDC, hBrush)
          hPen = CreatePen(PS_SOLID, 1, buttonBorderColor)
          If hPen Then
            hPreviousPen = SelectObject(hDC, hPen)
            DrawRectangle hDC, drawItemData.rcItem.Left, drawItemData.rcItem.Top, drawItemData.rcItem.Right, drawItemData.rcItem.Bottom

            SelectObject hDC, hPreviousPen
            DeleteObject hPen
          End If
          SelectObject hDC, hPreviousBrush
          DeleteObject hBrush
        End If
      Else
        LSet rc = drawItemData.rcItem
        rc.Right = rc.Left + MENUSIDEBARWIDTH - 1
        hBrush = CreateSolidBrush(droppedDownButtonFaceColor)
        If hBrush Then
          FillRect hDC, rc, hBrush
          DeleteObject hBrush
        End If
        rc.Left = rc.Right
        rc.Right = drawItemData.rcItem.Right
        hBrush = CreateSolidBrush(menuLightColor)
        If hBrush Then
          FillRect hDC, rc, hBrush
          DeleteObject hBrush
        End If
      End If

      LSet iconRectangle = drawItemData.rcItem
      iconRectangle.Right = iconRectangle.Left + drawItemData.rcItem.Bottom - drawItemData.rcItem.Top + 1
      LSet textRectangle = drawItemData.rcItem
      textRectangle.Left = iconRectangle.Left + MENUSIDEBARWIDTH + HMARGIN_TEXT_LEFT
      InflateRect iconRectangle, -1, -1

      If Not tagObject Is Nothing Then
        If tagObject.Separator Then
          hPen = CreatePen(PS_SOLID, 1, separatorColor)
          If hPen Then
            hPreviousPen = SelectObject(hDC, hPen)

            MoveToEx hDC, textRectangle.Left, textRectangle.Top + (textRectangle.Bottom - textRectangle.Top) / 2, 0
            LineTo hDC, textRectangle.Right, textRectangle.Top + (textRectangle.Bottom - textRectangle.Top) / 2

            SelectObject hDC, hPreviousPen
            DeleteObject hPen
          End If
        Else
          itemText = tagObject.Text
          If itemText <> "" Then
            previousBkMode = SetBkMode(hDC, TRANSPARENT)
            previousTextColor = SetTextColor(hDC, GetSysColor(IIf(drawItemData.itemState And ODS_GRAYED, COLOR_GRAYTEXT, COLOR_MENUTEXT)))

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

        If drawItemData.itemState And ODS_CHECKED Then
          hBrush = CreateSolidBrush(IIf(drawItemData.itemState And (ODS_SELECTED Or ODS_GRAYED) = ODS_SELECTED, pressedButtonBackColor, checkedButtonBackColor))
          If hBrush Then
            hPreviousBrush = SelectObject(hDC, hBrush)
            hPen = CreatePen(PS_SOLID, 1, buttonBorderColor)
            If hPen Then
              hPreviousPen = SelectObject(hDC, hPen)
              DrawRectangle hDC, iconRectangle.Left, iconRectangle.Top, iconRectangle.Right, iconRectangle.Bottom

              SelectObject hDC, hPreviousPen
              DeleteObject hPen
            End If
            SelectObject hDC, hPreviousBrush
            DeleteObject hBrush
          End If
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

        If tagObject.HasSubMenu Then
          ' draw sub menu arrow
          textRectangle.Left = textRectangle.Right - HMARGIN_SUBMENUARROW
          DrawSubMenuArrow menuFontSize, hDC, textRectangle, (drawItemData.itemState And ODS_GRAYED) = 0
        End If
      End If

      SelectObject hDC, hPreviousFont

      ' this will prevent drawing of the default arrow
      ExcludeClipRect hDC, drawItemData.rcItem.Left, drawItemData.rcItem.Top, drawItemData.rcItem.Right, drawItemData.rcItem.Bottom
    End If
  End If

  result = 1
  handledEvent = True
End Sub

Private Sub DrawMenuFrame(ByVal hWnd As Long, ByVal hDC As Long, ByVal isTopLevelMenu As Boolean)
  Const COLOR_BTNSHADOW = 16
  Const PS_SOLID = 0
  Dim hPen As Long
  Dim hPreviousPen As Long
  Dim rc As RECT

  GetWindowRect hWnd, rc
  OffsetRect rc, -rc.Left, -rc.Top

  hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNSHADOW))
  If hPen Then
    hPreviousPen = SelectObject(hDC, hPen)

    MoveToEx hDC, rc.Left, rc.Top, 0
    LineTo hDC, rc.Right - 1, rc.Top
    LineTo hDC, rc.Right - 1, rc.Bottom - 1
    LineTo hDC, rc.Left, rc.Bottom - 1
    LineTo hDC, rc.Left, rc.Top

    SelectObject hDC, hPreviousPen
    DeleteObject hPen
  End If

  If isTopLevelMenu Then
    hPen = CreatePen(PS_SOLID, 1, menuLightColor)
    If hPen Then
      hPreviousPen = SelectObject(hDC, hPen)

      MoveToEx hDC, rc.Left + 1, rc.Top + 1, 0
      LineTo hDC, rc.Right - 2, rc.Top + 1
      LineTo hDC, rc.Right - 2, rc.Bottom - 2
      LineTo hDC, rc.Left + 1, rc.Bottom - 2
      LineTo hDC, rc.Left + 1, rc.Top + 1

      MoveToEx hDC, rc.Left + 2, rc.Top + 2, 0
      LineTo hDC, rc.Right - 3, rc.Top + 2
      LineTo hDC, rc.Right - 3, rc.Bottom - 3
      LineTo hDC, rc.Left + 2, rc.Bottom - 3
      LineTo hDC, rc.Left + 2, rc.Top + 2

      If isTopLevelMenu Then
        MoveToEx hDC, rc.Left + 1, rc.Top, 0
        LineTo hDC, rc.Left + lastDroppedButtonWidth - 1, rc.Top
      End If

      SelectObject hDC, hPreviousPen
      DeleteObject hPen
    End If
  End If
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
          measureItemData.itemHeight = 5
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
