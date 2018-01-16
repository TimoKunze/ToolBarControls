VERSION 5.00
Object = "{5D0D0ABC-4898-4E46-AB48-291074A737A1}#1.3#0"; "TBarCtlsU.ocx"
Object = "{52D76F35-4551-4C56-B53B-A343E42B0AF8}#2.3#0"; "ProgBarU.ocx"
Object = "{FCCB83BF-E483-4317-9FF2-A460758238B5}#1.2#0"; "CBLCtlsU.ocx"
Begin VB.UserControl UCAddressBand 
   ClientHeight    =   1290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ScaleHeight     =   86
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   Begin BreadcrumbBar.UCBreadcrumb Breadcrumb 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
   End
   Begin ProgBarLibUCtl.ProgressBar Progress 
      Height          =   375
      Left            =   1800
      Top             =   240
      Width           =   855
      _cx             =   1508
      _cy             =   661
      ActivateMarquee =   0   'False
      Appearance      =   3
      BackColor       =   -1
      BarColor        =   -1
      BarStyle        =   0
      BorderStyle     =   0
      CurrentValue    =   0
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   3
      DisplayText     =   0   'False
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      HAlignment      =   1
      HoverTime       =   -1
      MarqueeStepDuration=   50
      Maximum         =   100
      Minimum         =   0
      MousePointer    =   0
      Orientation     =   0
      ProgressState   =   1
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      RightToLeftLayout=   0   'False
      SmoothReverse   =   -1  'True
      StepWidth       =   10
      SupportOLEDragImages=   -1  'True
      TextShadowColor =   3158064
      TextShadowOffsetX=   1
      TextShadowOffsetY=   1
      UseSystemFont   =   -1  'True
      DisplayProgressInTaskBar=   0   'False
      Text            =   "UCAddressBand.ctx":0000
   End
   Begin TBarCtlsLibUCtl.ToolBar ToolButtons 
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   855
      _cx             =   1508
      _cy             =   873
      AllowCustomization=   0   'False
      AlwaysDisplayButtonText=   -1  'True
      AnchorHighlighting=   0   'False
      Appearance      =   0
      BackStyle       =   0
      BorderStyle     =   0
      ButtonHeight    =   28
      ButtonStyle     =   1
      ButtonTextPosition=   1
      ButtonWidth     =   22
      DisabledEvents  =   917994
      DisplayMenuDivider=   0   'False
      DisplayPartiallyClippedButtons=   0   'False
      DontRedraw      =   0   'False
      DragClickTime   =   -1
      DragDropCustomizationModifierKey=   0
      DropDownGap     =   -1
      Enabled         =   -1  'True
      FirstButtonIndentation=   0
      FocusOnClick    =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      HorizontalTextAlignment=   0
      HoverTime       =   -1
      InsertMarkColor =   0
      MaximumButtonWidth=   0
      MaximumTextRows =   1
      MenuBarTheme    =   1
      MenuMode        =   0   'False
      MinimumButtonWidth=   0
      MousePointer    =   0
      MultiColumn     =   0   'False
      NormalDropDownButtonStyle=   1
      OLEDragImageStyle=   0
      Orientation     =   0
      ProcessContextMenuKeys=   -1  'True
      RaiseCustomDrawEventOnEraseBackground=   -1  'True
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ShadowColor     =   -1
      ShowToolTips    =   -1  'True
      SupportOLEDragImages=   -1  'True
      UseMnemonics    =   -1  'True
      UseSystemFont   =   -1  'True
      VerticalButtonPadding=   -1
      VerticalButtonSpacing=   0
      VerticalTextAlignment=   1
      WrapButtons     =   0   'False
   End
   Begin CBLCtlsLibUCtl.ImageComboBox PathCombo 
      Height          =   330
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   975
      _cx             =   1720
      _cy             =   582
      AcceptNumbersOnly=   0   'False
      AutoHorizontalScrolling=   -1  'True
      BackColor       =   -2147483643
      CaseSensitiveItemSearching=   0   'False
      DisabledEvents  =   283751
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragDropDownTime=   -1
      DropDownKey     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      HoverTime       =   -1
      IconVisibility  =   0
      IMEMode         =   -1
      ItemHeight      =   -1
      ListAlwaysShowVerticalScrollBar=   0   'False
      ListDragScrollTimeBase=   -1
      ListHeight      =   -1
      ListInsertMarkColor=   0
      ListWidth       =   0
      MaxTextLength   =   -1
      MousePointer    =   0
      OLEDragImageStyle=   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SelectionFieldHeight=   -1
      Sorted          =   0   'False
      Style           =   0
      SupportOLEDragImages=   -1  'True
      TextEndEllipsis =   -1  'True
      UseShellWordBreakFunction=   0   'False
      UseSystemFont   =   -1  'True
      CueBanner       =   "UCAddressBand.ctx":002A
      Text            =   "UCAddressBand.ctx":004A
   End
End
Attribute VB_Name = "UCAddressBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  ' This sample was heavily inspired by the "Explorer breadcrumbs in Vista" sample by
  ' Bjarke Viksoe which can be found on http://www.viksoe.dk/code/breadcrumbs.htm


  Implements ISubclassedWindow


  Private Const BTNID_HISTROY As Long = 1
  Private Const BTNID_GO As Long = 2


  Private WithEvents m_BreadcrumbToolBar As ToolBar
Attribute m_BreadcrumbToolBar.VB_VarHelpID = -1
  Private m_hotTracked As Long
  Private m_hImageList As Long
  Private m_hTheme As Long
  Private m_subclassed As Boolean


  Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hrgnDest As Long, ByVal hrgnSrc1 As Long, ByVal hrgnSrc2 As Long, ByVal fnCombineMode As Long) As Long
  Private Declare Function CreateRectRgnIndirect Lib "gdi32.dll" (lprc As RECT) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal hDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, ByVal pClipRect As Long) As Long
  Private Declare Function DrawThemeParentBackground Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal hDC As Long, prc As RECT) As Long
  Private Declare Function ExcludeClipRect Lib "gdi32.dll" (ByVal hDC As Long, ByVal nLeftRect As Long, ByVal nTopRect As Long, ByVal nRightRect As Long, ByVal nBottomRect As Long) As Long
  Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetThemeInt Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, ByRef piVal As Long) As Long
  Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hWnd As Long) As Long
  Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
  Private Declare Function ImageList_LoadImage Lib "comctl32.dll" Alias "ImageList_LoadImageW" (ByVal hi As Long, ByVal lpbmp As Long, ByVal cx As Long, ByVal cGrow As Long, ByVal crMask As Long, ByVal uType As Long, ByVal uFlags As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function InflateRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function MoveWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
  Private Declare Function OffsetRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hDC As Long) As Long
  Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

  Public Event ChangeCurrentPIDL(ByVal pIDL As Long)


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidUCAddressBand
      lRet = HandleMessage_UserControl(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case EnumSubclassID.escidUCAddressBandProgress
      lRet = HandleMessage_Progress(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "UCAddressBand.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in UCAddressBand.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_UserControl(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const COLOR_WINDOW As Long = 5
  Const RGN_AND As Long = 1
  Const SM_CXBORDER As Long = 5
  Const SM_CXEDGE As Long = 45
  Const SM_CYBORDER As Long = 6
  Const SM_CYEDGE As Long = 46
  Const TME_LEAVE As Long = &H2
  Const TMT_SIZINGBORDERWIDTH As Long = 1201
  Const WM_ERASEBKGND As Long = &H14
  Const WM_NCPAINT As Long = &H85
  Const WM_PAINT As Long = &HF
  Const WM_PRINTCLIENT As Long = &H318
  Const WM_SIZE As Long = &H5
  Const WM_THEMECHANGED As Long = &H31A
  Dim cxBorder As Long
  Dim cxEdge As Long
  Dim cyBorder As Long
  Dim cyEdge As Long
  Dim hDC As Long
  Dim hRgn As Long
  Dim hUpdateRgn As Long
  Dim ps As PAINTSTRUCT
  Dim rcClient As RECT
  Dim rcWnd As RECT
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_ERASEBKGND
      ' make sure the control's background is drawn properly
      If Not IsCompositionEnabled Then DoPaint wParam
      lRet = 1
      bCallDefProc = False
    Case WM_NCPAINT
      ' draws the control's border
      hUpdateRgn = IIf(wParam <> 1, wParam, 0)
      hDC = GetWindowDC(hWnd)
      If hDC Then
        cxBorder = GetSystemMetrics(SM_CXBORDER)
        cyBorder = GetSystemMetrics(SM_CYBORDER)
        If GetThemeInt(m_hTheme, 0, 0, TMT_SIZINGBORDERWIDTH, cxBorder) >= S_OK Then
          cyBorder = cxBorder
        End If

        GetWindowRect hWnd, rcWnd
        cxEdge = GetSystemMetrics(SM_CXEDGE)
        cyEdge = GetSystemMetrics(SM_CYEDGE)
        InflateRect rcWnd, -cxEdge, -cyEdge
        hRgn = CreateRectRgnIndirect(rcWnd)
        If hRgn Then
          If hUpdateRgn Then
            CombineRgn hRgn, hUpdateRgn, hRgn, RGN_AND
          End If
          OffsetRect rcWnd, -rcWnd.Left, -rcWnd.Top
          OffsetRect rcWnd, cxEdge, cyEdge
          ExcludeClipRect hDC, rcWnd.Left, rcWnd.Top, rcWnd.Right, rcWnd.Bottom
          InflateRect rcWnd, cxEdge, cyEdge

          If (cxBorder < cxEdge) And (cyBorder < cyEdge) Then
            InflateRect rcWnd, cxBorder - cxEdge, cyBorder - cyEdge
            FillRect hDC, rcWnd, GetSysColorBrush(COLOR_WINDOW)
          End If
          DefSubclassProc hWnd, WM_NCPAINT, hRgn, 0

          DeleteObject hRgn
        End If

        ReleaseDC hWnd, hDC
      End If
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
      PositionControls rcClient
      bCallDefProc = False
    Case WM_THEMECHANGED
      CloseThemeData m_hTheme
      m_hTheme = OpenThemeData(hWnd, StrPtr("ABComposited::AddressBand"))
  End Select

StdHandler_Ende:
  HandleMessage_UserControl = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in UCAddressBand.HandleMessage_UserControl: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Progress(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_ERASEBKGND As Long = &H14
  Const WM_PAINT As Long = &HF
  Const WM_PRINTCLIENT As Long = &H318
  Dim hDC As Long
  Dim ps As PAINTSTRUCT
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_ERASEBKGND
      ' make sure the control's background is drawn properly
      If Not IsCompositionEnabled Then DoProgressPaint wParam
      lRet = 1
      bCallDefProc = False
    Case WM_PAINT, WM_PRINTCLIENT
      ' make sure the control's background is drawn properly
      If wParam Then
        If IsCompositionEnabled Then DoProgressPaint wParam
      Else
        hDC = BeginPaint(hWnd, ps)
        If IsCompositionEnabled Then DoProgressPaint hDC
        EndPaint hWnd, ps
      End If
      If Progress.CurrentValue > 0 Then DefSubclassProc hWnd, uMsg, wParam, lParam
      lRet = 1
      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Progress = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in UCAddressBand.HandleMessage_Progress: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub Breadcrumb_ChangeCurrentPIDL(ByVal pIDL As Long)
  RaiseEvent ChangeCurrentPIDL(pIDL)
End Sub

Private Sub Breadcrumb_ChangedCurrentPIDL()
  ToolButtons.Buttons(BTNID_GO, btitID).Text = "Refresh """ & Breadcrumb.CurrentDirectoryName & """"
End Sub

Private Sub Breadcrumb_EmptySpaceClick()
  Dim dispName As STRRET
  Dim iconIndex As Long
  Dim IParentISF As IVBShellFolder
  Dim pIDLToParent As Long
  Dim rcClient As RECT

  PathCombo.ComboItems.RemoveAll
  SplitFullyQualifiedPIDL Breadcrumb.CurrentPIDL, IParentISF, pIDLToParent
  If Not (IParentISF Is Nothing) And pIDLToParent <> 0 Then
    dispName.uType = STRRET_OFFSET
    If IParentISF.GetDisplayNameOf(pIDLToParent, SHGDNConstants.SHGDN_FORADDRESSBAR Or SHGDNConstants.SHGDN_FORPARSING, dispName) = S_OK Then
      iconIndex = GetSysIconIndex(Breadcrumb.CurrentPIDL)
      Set PathCombo.SelectedItem = PathCombo.ComboItems.Add(STRRETToString(dispName, pIDLToParent), , iconIndex, , , , ILClone(Breadcrumb.CurrentPIDL))

      Debug.Print "TODO: Insert history items here"

      Breadcrumb.Visible = False
      ToolButtons.Buttons(BTNID_HISTROY, btitID).Visible = False
      GetClientRect UserControl.hWnd, rcClient
      PositionControls rcClient
      PathCombo.Visible = True
      PathCombo.SetFocus
    End If
  End If
  Set IParentISF = Nothing
End Sub

Private Sub m_BreadcrumbToolBar_MouseEnter(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  m_hotTracked = m_hotTracked + 1
  Breadcrumb.RefreshBreadcrumbBar
  ToolButtons.Refresh
End Sub

Private Sub m_BreadcrumbToolBar_MouseLeave(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  m_hotTracked = m_hotTracked - 1
  Breadcrumb.RefreshBreadcrumbBar
  ToolButtons.Refresh
End Sub

Private Sub PathCombo_FreeItemData(ByVal comboItem As CBLCtlsLibUCtl.IImageComboBoxItem)
  If Not (comboItem Is Nothing) Then
    ILFree comboItem.itemData
  End If
End Sub

Private Sub PathCombo_KeyPress(keyAscii As Integer)
  Dim pIDL As Long

  If keyAscii = vbKeyReturn Then
    pIDL = ILCreateFromPath(StrPtr(PathCombo.Text))
    PathCombo_LostFocus
    If pIDL Then
      RaiseEvent ChangeCurrentPIDL(pIDL)
      ILFree pIDL
    Else
      Debug.Print "TODO: Accept paths like 'Desktop' and 'Computer'"
    End If
  End If
End Sub

Private Sub PathCombo_LostFocus()
  Dim rcClient As RECT

  PathCombo.Visible = False
  PathCombo.ComboItems.RemoveAll
  ToolButtons.Buttons(BTNID_HISTROY, btitID).Visible = True
  GetClientRect UserControl.hWnd, rcClient
  PositionControls rcClient
  Breadcrumb.Visible = True
End Sub

Private Sub ToolButtons_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  If commandID = BTNID_HISTROY Then
    Breadcrumb_EmptySpaceClick
    PathCombo.OpenDropDownWindow
  End If
End Sub

Private Sub ToolButtons_MouseEnter(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  m_hotTracked = m_hotTracked + 1
  Breadcrumb.RefreshBreadcrumbBar
  ToolButtons.Refresh
End Sub

Private Sub ToolButtons_MouseLeave(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  m_hotTracked = m_hotTracked - 1
  Breadcrumb.RefreshBreadcrumbBar
  ToolButtons.Refresh
End Sub

Private Sub UserControl_Initialize()
  Const IMAGE_BITMAP As Long = 0
  Const LR_CREATEDIBSECTION As Long = &H2000&
  Const LR_LOADFROMFILE As Long = &H10
  Const SHIL_SMALL As Long = &H1
  Const strIID_IImageList As String = "{46EB5926-582E-4017-9FDF-E8998DAA0950}"
  Dim btn As ToolBarButton
  Dim hSysImageList As Long
  Dim IID_IImageList As UUID
  Dim imgDir As String

  CLSIDFromString StrPtr(strIID_IImageList), IID_IImageList

  m_hTheme = OpenThemeData(UserControl.hWnd, StrPtr("ABComposited::AddressBand"))

  If Right$(App.Path, 3) = "bin" Then
    imgDir = App.Path & "\..\res\"
  Else
    imgDir = App.Path & "\res\"
  End If
  m_hImageList = ImageList_LoadImage(0, StrPtr(imgDir & "GoButton.bmp"), 16, 0, vbBlack, IMAGE_BITMAP, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
  ToolButtons.hImageList(ilNormalButtons) = m_hImageList

  Breadcrumb.SetNotifyWindow UserControl.hWnd
  SetParent Breadcrumb.hWnd, Progress.hWnd
  Set m_BreadcrumbToolBar = Breadcrumb.BreadcrumbToolBar

  SetParent PathCombo.hWnd, Progress.hWnd
  SHGetImageList SHIL_SMALL, IID_IImageList, VarPtr(hSysImageList)
  PathCombo.hImageList(ilItems) = hSysImageList
  SetWindowTheme PathCombo.hWndComboBox, StrPtr("AddressComposited"), 0
  SetWindowTheme PathCombo.hWndEdit, StrPtr("AddressComposited"), 0

  SetParent ToolButtons.hWnd, Progress.hWnd
  Call ToolButtons.Buttons.Add(BTNID_HISTROY, , 3, , "Previous Locations", AutoSize:=False, Width:=19)
  Call ToolButtons.Buttons.Add(BTNID_GO, , 2, , "Refresh", AutoSize:=False, Width:=25)

  ' set some special themes for the tool bar
  SetWindowTheme ToolButtons.hWnd, StrPtr("GoComposited"), 0
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  With PropBag
    Enabled = PropBag.ReadProperty("Enabled", True)
  End With
End Sub

Private Sub UserControl_Resize()
  Dim rcClient As RECT

  If IsInIDE Then
    ' during runtime the child controls will be aligned whenever we receive a WM_SIZE message
    GetClientRect UserControl.hWnd, rcClient
    PositionControls rcClient
  ElseIf Not m_subclassed Then
    If SubclassWindow(UserControl.hWnd, Me, EnumSubclassID.escidUCAddressBand) And SubclassWindow(Progress.hWnd, Me, EnumSubclassID.escidUCAddressBandProgress) Then
      m_subclassed = True
    Else
      Debug.Print "Subclassing failed!"
    End If
  End If
End Sub

Private Sub UserControl_Terminate()
  PathCombo.ComboItems.RemoveAll
  If UnSubclassWindow(UserControl.hWnd, EnumSubclassID.escidUCAddressBand) And UnSubclassWindow(Progress.hWnd, EnumSubclassID.escidUCAddressBandProgress) Then
    m_subclassed = False
  Else
    Debug.Print "UnSubclassing failed!"
  End If
  If m_hTheme Then
    CloseThemeData m_hTheme
  End If
  If m_hImageList Then
    ImageList_Destroy m_hImageList
  End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  With PropBag
    PropBag.WriteProperty "Enabled", Enabled, True
  End With
End Sub


Public Property Get CurrentPIDL() As Long
  CurrentPIDL = Breadcrumb.CurrentPIDL
End Property

Public Property Let CurrentPIDL(ByVal newValue As Long)
  Breadcrumb.CurrentPIDL = newValue
End Property

Public Property Get Enabled() As Boolean
  Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal newValue As Boolean)
  If UserControl.Enabled <> newValue Then
    UserControl.Enabled = newValue
    PropertyChanged "Enabled"
  End If
End Property

Public Property Get hWnd() As Long
  hWnd = UserControl.hWnd
End Property

Public Property Get ProgressBar() As ProgressBar
  Set ProgressBar = Progress
End Property

Public Sub Refresh()
  Progress.Refresh
  Breadcrumb.RefreshBreadcrumbBar
  ToolButtons.Refresh
End Sub


Private Sub DoPaint(ByVal hDC As Long)
  Const SBB_DISABLED As Long = 3
  Const SBB_FOCUSED As Long = 4
  Const SBB_HOT As Long = 2
  Const SBB_NORMAL As Long = 1
  Const SBP_ABBACKGROUND As Long = 1
  Dim rcClient As RECT
  Dim stateID As Long

  ' draw the control background using the theme API
  GetClientRect UserControl.hWnd, rcClient
  stateID = SBB_NORMAL
  If m_hotTracked > 0 Then stateID = SBB_HOT
  If Breadcrumb.IsDropped Then stateID = SBB_FOCUSED
  If Not Me.Enabled Then stateID = SBB_DISABLED
  DrawThemeParentBackground UserControl.hWnd, hDC, rcClient
  DrawThemeBackground m_hTheme, hDC, SBP_ABBACKGROUND, stateID, rcClient, 0
End Sub

Private Sub DoProgressPaint(ByVal hDC As Long)
  Const PBFS_NORMAL As Long = 1
  Const PP_FILL As Long = 5
  Dim barWidth As Long
  Dim hThemeProgress As Long
  Dim rcBar As RECT
  Dim rcClient As RECT
  Dim stateID As Long

  ' draw the control background using the theme API
  GetClientRect Progress.hWnd, rcClient
  DrawThemeParentBackground Progress.hWnd, hDC, rcClient

'  If Progress.CurrentValue > 0 Then
'    hThemeProgress = OpenThemeData(Progress.hWnd, StrPtr("Progress"))
'    If hThemeProgress Then
'      stateID = PBFS_NORMAL
'
'      rcBar = rcClient
'      barWidth = Progress.CurrentValue / (Progress.Maximum - Progress.Minimum) * (rcClient.Right - rcClient.Left)
'      rcBar.Right = rcBar.Left + barWidth
'      rcBar.Top = rcBar.Top + 3
'      rcBar.Bottom = rcBar.Bottom - 3
'      DrawThemeBackground hThemeProgress, hDC, PP_FILL, stateID, rcClient, VarPtr(rcBar)
'      CloseThemeData hThemeProgress
'    End If
'  End If
End Sub

Private Function IsInIDE() As Boolean
  IsInIDE = Not Ambient.UserMode
End Function

Private Sub PositionControls(rcClient As RECT)
  Const SWP_NOACTIVATE As Long = &H10
  Const SWP_NOZORDER As Long = &H4
  Dim buttonsWidth As Long
  Dim rcToolBar As RECT

  buttonsWidth = ToolButtons.IdealWidth
  rcToolBar.Right = rcClient.Right - 2
  rcToolBar.Left = rcToolBar.Right - buttonsWidth
  rcToolBar.Top = 1
  rcToolBar.Bottom = rcClient.Bottom - 1
  MoveWindow Progress.hWnd, rcClient.Left, rcClient.Top, rcClient.Right - rcClient.Left, rcClient.Bottom - rcClient.Top, 1
  MoveWindow PathCombo.hWnd, rcClient.Left + 1, rcClient.Top + 1, rcClient.Right - rcClient.Left - buttonsWidth - 2, rcClient.Bottom - rcClient.Top - 2, 1
  InflateRect rcClient, -2, -2
  MoveWindow Breadcrumb.hWnd, rcClient.Left, rcClient.Top, rcToolBar.Left - rcClient.Left, rcClient.Bottom - rcClient.Top, 1
  SetWindowPos ToolButtons.hWnd, 0, rcToolBar.Left, rcToolBar.Top, rcToolBar.Right - rcToolBar.Left, rcToolBar.Bottom - rcToolBar.Top, SWP_NOACTIVATE Or SWP_NOZORDER
End Sub
