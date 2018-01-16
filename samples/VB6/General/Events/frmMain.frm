VERSION 5.00
Object = "{5D0D0ABC-4898-4e46-AB48-291074A737A1}#1.3#0"; "TBarCtlsU.ocx"
Object = "{C5AB19C0-2E5A-4097-91EF-61EFA3AC27E7}#1.3#0"; "TBarCtlsA.ocx"
Begin VB.Form frmMain 
   Caption         =   "ToolBarControls 1.3 - Events Sample"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   2  'Bildschirmmitte
   Begin TBarCtlsLibUCtl.ToolBar tbMenu 
      Height          =   390
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   4695
      _cx             =   8281
      _cy             =   688
      AllowCustomization=   0   'False
      AlwaysDisplayButtonText=   -1  'True
      AnchorHighlighting=   0   'False
      Appearance      =   0
      BackStyle       =   0
      BorderStyle     =   0
      ButtonHeight    =   0
      ButtonStyle     =   1
      ButtonTextPosition=   1
      ButtonWidth     =   0
      DisabledEvents  =   0
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
      MenuMode        =   -1  'True
      MinimumButtonWidth=   0
      MousePointer    =   0
      MultiColumn     =   0   'False
      NormalDropDownButtonStyle=   1
      OLEDragImageStyle=   0
      Orientation     =   0
      ProcessContextMenuKeys=   -1  'True
      RaiseCustomDrawEventOnEraseBackground=   0   'False
      RegisterForOLEDragDrop=   0
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
   Begin VB.ComboBox cmbAddressA 
      Height          =   315
      Left            =   600
      TabIndex        =   5
      Top             =   3480
      Width           =   1695
   End
   Begin VB.ComboBox cmbAddressU 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin TBarCtlsLibACtl.ToolBar ToolBarA 
      Height          =   390
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   4695
      _cx             =   8281
      _cy             =   688
      AllowCustomization=   0   'False
      AlwaysDisplayButtonText=   0   'False
      AnchorHighlighting=   0   'False
      Appearance      =   0
      BackStyle       =   0
      BorderStyle     =   0
      ButtonHeight    =   0
      ButtonStyle     =   1
      ButtonTextPosition=   1
      ButtonWidth     =   0
      DisabledEvents  =   0
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
      NormalDropDownButtonStyle=   0
      OLEDragImageStyle=   0
      Orientation     =   0
      ProcessContextMenuKeys=   -1  'True
      RaiseCustomDrawEventOnEraseBackground=   0   'False
      RegisterForOLEDragDrop=   1
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
   Begin TBarCtlsLibACtl.ReBar ReBarA 
      Align           =   2  'Unten ausrichten
      Height          =   690
      Left            =   0
      TabIndex        =   3
      Top             =   5250
      Width           =   9600
      _cx             =   16933
      _cy             =   1217
      AllowBandReordering=   -1  'True
      Appearance      =   0
      AutoUpdateLayout=   -1  'True
      BackColor       =   -1
      BorderStyle     =   0
      DisabledEvents  =   0
      DisplayBandSeparators=   -1  'True
      DisplaySplitter =   0   'False
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      FixedBandHeight =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -1
      HighlightColor  =   -1
      HoverTime       =   -1
      MousePointer    =   0
      Orientation     =   0
      ReplaceMDIFrameMenu=   0
      RegisterForOLEDragDrop=   1
      RightToLeft     =   0
      ShadowColor     =   -1
      SupportOLEDragImages=   -1  'True
      ToggleOnDoubleClick=   -1  'True
      UseSystemFont   =   -1  'True
      VerticalSizingGripsOnVerticalOrientation=   0   'False
   End
   Begin TBarCtlsLibUCtl.ReBar ReBarU 
      Align           =   1  'Oben ausrichten
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      _cx             =   16933
      _cy             =   1217
      AllowBandReordering=   -1  'True
      Appearance      =   0
      AutoUpdateLayout=   -1  'True
      BackColor       =   -1
      BorderStyle     =   0
      DisabledEvents  =   0
      DisplayBandSeparators=   -1  'True
      DisplaySplitter =   0   'False
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      FixedBandHeight =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -1
      HighlightColor  =   -1
      HoverTime       =   -1
      MousePointer    =   0
      Orientation     =   0
      ReplaceMDIFrameMenu=   0
      RegisterForOLEDragDrop=   1
      RightToLeft     =   0
      ShadowColor     =   -1
      SupportOLEDragImages=   -1  'True
      ToggleOnDoubleClick=   -1  'True
      UseSystemFont   =   -1  'True
      VerticalSizingGripsOnVerticalOrientation=   0   'False
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   5040
      TabIndex        =   7
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6150
      TabIndex        =   8
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      Height          =   3975
      Left            =   5040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   6
      Top             =   960
      Width           =   4455
   End
   Begin TBarCtlsLibUCtl.ToolBar ToolBarU 
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   4695
      _cx             =   8281
      _cy             =   688
      AllowCustomization=   0   'False
      AlwaysDisplayButtonText=   0   'False
      AnchorHighlighting=   0   'False
      Appearance      =   0
      BackStyle       =   0
      BorderStyle     =   0
      ButtonHeight    =   0
      ButtonStyle     =   1
      ButtonTextPosition=   1
      ButtonWidth     =   0
      DisabledEvents  =   0
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
      NormalDropDownButtonStyle=   0
      OLEDragImageStyle=   0
      Orientation     =   0
      ProcessContextMenuKeys=   -1  'True
      RaiseCustomDrawEventOnEraseBackground=   0   'False
      RegisterForOLEDragDrop=   1
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Const MENUID_FILE = 1
  Private Const MENUID_VIEW = 2
  Private Const MENUID_HELP = 3


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
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
    hBitmap As Long
  End Type

  Private Type POINTAPI
    x As Long
    y As Long
  End Type

  Private Type TOOLBARBUTTONINFO
    ID As Long
    IconIndex As Long
    ButtonData As Long
    Text As String
    hSubMenu As Long
  End Type


  Private BANDID_TOOLBAR As Long

  Private MENUID_FILEEXIT As Long
  Private MENUID_VIEWTOOLBAR As Long
  Private MENUID_VIEWICONTEXT As Long
  Private MENUID_VIEWTOGGLEDOCK As Long
  Private MENUID_VIEWCUSTOMIZE As Long
  Private MENUID_HELPABOUT As Long

  Private bComctl32Version600OrNewer As Boolean
  Private bLog As Boolean
  Private buttonsA(0 To 9) As TOOLBARBUTTONINFO
  Private buttonsU(0 To 9) As TOOLBARBUTTONINFO
  Private WithEvents ctlHostWindow As TBarCtlsLibUCtl.ControlHostWindow
Attribute ctlHostWindow.VB_VarHelpID = -1
  Private hImgLst As Long
  Private menuButtons(0 To 2) As TOOLBARBUTTONINFO
  Private menusToDestroy As Collection
  Private objActiveCtl As Object
  Private themeableOS As Boolean


  Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal sz As Long)
  Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long
  Private Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As Long, ByVal fuFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByVal lptpm As Long) As Long


Private Function ISubclassedWindow_HandleMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case eSubclassID
    Case EnumSubclassID.escidFrmMain
      lRet = HandleMessage_Form(hWnd, uMsg, wParam, lParam, bCallDefProc)
    Case Else
      Debug.Print "frmMain.ISubclassedWindow_HandleMessage: Unknown Subclassing ID " & CStr(eSubclassID)
  End Select

StdHandler_Ende:
  ISubclassedWindow_HandleMessage = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.ISubclassedWindow_HandleMessage (SubclassID=" & CStr(eSubclassID) & ": ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function

Private Function HandleMessage_Form(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
  Const WM_NOTIFYFORMAT = &H55
  Const WM_USER = &H400
  Const OCM__BASE = WM_USER + &H1C00
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_NOTIFYFORMAT
      ' give the control a chance to request Unicode notifications
      lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)

      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub ctlHostWindow_Activate(ByVal activatedByMouseClick As Boolean, ByVal hWndBeingDeactivated As Long)
  AddLogEntry "ctlHostWindow_Activate: activatedByMouseClick=" & activatedByMouseClick & ", hWndBeingDeactivated=0x" & Hex(hWndBeingDeactivated)
End Sub

Private Sub ctlHostWindow_Closing(Cancel As Boolean)
  AddLogEntry "ctlHostWindow_Closing: cancel=" & Cancel

  Cancel = True
  tbMenu_ExecuteCommand MENUID_VIEWTOOLBAR, Nothing, coMenu, False
End Sub

Private Sub ctlHostWindow_Created(ByVal hWnd As Long)
  AddLogEntry "ctlHostWindow_Created: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ctlHostWindow_Deactivate(ByVal hWndBeingActivated As Long)
  AddLogEntry "ctlHostWindow_Deactivate: hWndBeingActivated=0x" & Hex(hWndBeingActivated)
End Sub

Private Sub ctlHostWindow_Destroyed(ByVal hWnd As Long)
  AddLogEntry "ctlHostWindow_Destroyed: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ctlHostWindow_Moved(ByVal Left As stdole.OLE_XPOS_PIXELS, ByVal Top As stdole.OLE_YPOS_PIXELS)
  AddLogEntry "ctlHostWindow_Moved: left=" & Left & ", top=" & Top
End Sub

Private Sub ctlHostWindow_Moving(Left As stdole.OLE_XPOS_PIXELS, Top As stdole.OLE_YPOS_PIXELS, preventMove As Boolean)
  AddLogEntry "ctlHostWindow_Moving: left=" & Left & ", top=" & Top & ", preventMove=" & preventMove
End Sub

Private Sub ctlHostWindow_Resized(ByVal Width As stdole.OLE_XSIZE_PIXELS, ByVal Height As stdole.OLE_YSIZE_PIXELS)
  AddLogEntry "ctlHostWindow_Resized: width=" & Width & ", height=" & Height
End Sub

Private Sub ctlHostWindow_Resizing(Width As stdole.OLE_XSIZE_PIXELS, Height As stdole.OLE_YSIZE_PIXELS, preventResize As Boolean)
  AddLogEntry "ctlHostWindow_Resizing: width=" & Width & ", height=" & Height & ", preventResize=" & preventResize
End Sub

Private Sub ctlHostWindow_TitleDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As stdole.OLE_XPOS_PIXELS, ByVal y As stdole.OLE_YPOS_PIXELS)
  AddLogEntry "ctlHostWindow_TitleDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
  tbMenu_ExecuteCommand MENUID_VIEWTOGGLEDOCK, Nothing, coMenu, False
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim DLLVerData As DLLVERSIONINFO
  Dim hIcon As Long
  Dim hMod As Long
  Dim iconsDir As String
  Dim iconPath As String
  Dim iconSize As Long

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  iconSize = 16
  hImgLst = ImageList_Create(iconSize, iconSize, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 14, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconsDir = iconsDir & iconSize & "x" & iconSize & IIf(bComctl32Version600OrNewer, "x32bpp\normal\", "x8bpp\normal\")
  iconPath = Dir$(iconsDir & "*.ico")
  While iconPath <> ""
    hIcon = LoadImage(0, StrPtr(iconsDir & iconPath), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst, hIcon
      DestroyIcon hIcon
    End If
    iconPath = Dir$
  Wend
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked

  Subclass

  Set menusToDestroy = New Collection
  CreateMenus

  Me.Show

  InsertButtonsA
  InsertButtonsU
  InsertBandsA
  InsertBandsU
End Sub

Private Sub Form_Resize()
  If Me.WindowState <> vbMinimized Then
    On Error Resume Next
    txtLog.Move 0, ReBarU.Height, Me.ScaleWidth, Me.ScaleHeight - ReBarA.Height - ReBarU.Height - cmdAbout.Height - 10
    chkLog.Move 5, txtLog.Top + txtLog.Height + 5
    cmdAbout.Move chkLog.Left + chkLog.Width + ((txtLog.Width - chkLog.Width) - cmdAbout.Width) / 2, txtLog.Top + txtLog.Height + (Me.ScaleHeight - ReBarA.Height - (txtLog.Top + txtLog.Height) - cmdAbout.Height) / 2
  End If
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
  Set ctlHostWindow = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim i As Long

  For i = 1 To menusToDestroy.Count
    DestroyMenu CLng(menusToDestroy(i))
  Next i
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub ReBarA_AutoBreakingBand(ByVal band As TBarCtlsLibACtl.IReBarBand, doAutoBreak As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarA_AutoBreakingBand: band=Nothing, doAutoBreak=" & doAutoBreak
  Else
    AddLogEntry "ReBarA_AutoBreakingBand: band=" & band.Text & ", doAutoBreak=" & doAutoBreak
  End If
End Sub

Private Sub ReBarA_AutoSized(targetRectangle As TBarCtlsLibACtl.RECTANGLE, actualRectangle As TBarCtlsLibACtl.RECTANGLE, ByVal changedBandHeightOrStyle As Boolean)
  AddLogEntry "ReBarA_AutoSized: targetRectangle=(" & targetRectangle.Left & "," & targetRectangle.Top & ")-(" & targetRectangle.Right & "," & targetRectangle.Bottom & "), actualRectangle=(" & actualRectangle.Left & "," & actualRectangle.Top & ")-(" & actualRectangle.Right & "," & actualRectangle.Bottom & "), changedBandHeightOrStyle=" & changedBandHeightOrStyle
End Sub

Private Sub ReBarA_BandBeginDrag(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants, cancelDrag As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarA_BandBeginDrag: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", cancelDrag=" & cancelDrag
  Else
    AddLogEntry "ReBarA_BandBeginDrag: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", cancelDrag=" & cancelDrag
  End If
End Sub

Private Sub ReBarA_BandBeginRDrag(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants, cancelDrag As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarA_BandBeginRDrag: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", cancelDrag=" & cancelDrag
  Else
    AddLogEntry "ReBarA_BandBeginRDrag: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", cancelDrag=" & cancelDrag
  End If
End Sub

Private Sub ReBarA_BandEndDrag(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_BandEndDrag: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_BandEndDrag: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_BandMouseEnter(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_BandMouseEnter: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_BandMouseEnter: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_BandMouseLeave(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_BandMouseLeave: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_BandMouseLeave: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_BeforeDisplayMDIChildSystemMenu(ByVal hActiveMDIChild As Long, ByVal hMenu As Long, ByVal x As Single, ByVal y As Single, cancelMenu As Boolean)
  AddLogEntry "ReBarA_BeforeDisplayMDIChildSystemMenu: hActiveMDIChild=0x" & Hex(hActiveMDIChild) & ", hMenu=0x" & Hex(hMenu) & ", x=" & x & ", y=" & y & ", cancelMenu=" & cancelMenu
End Sub

Private Sub ReBarA_ChevronClick(ByVal band As TBarCtlsLibACtl.IReBarBand, chevronRectangle As TBarCtlsLibACtl.RECTANGLE, ByVal userData As Long, doDefault As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarA_ChevronClick: band=Nothing, chevronRectangle=(" & chevronRectangle.Left & "," & chevronRectangle.Top & ")-(" & chevronRectangle.Right & "," & chevronRectangle.Bottom & "), userData=" & userData & ", doDefault=" & doDefault
  Else
    AddLogEntry "ReBarA_ChevronClick: band=" & band.Text & ", chevronRectangle=(" & chevronRectangle.Left & "," & chevronRectangle.Top & ")-(" & chevronRectangle.Right & "," & chevronRectangle.Bottom & "), userData=" & userData & ", doDefault=" & doDefault
  End If
End Sub

Private Sub ReBarA_CleanupMDIChildSystemMenu(ByVal hActiveMDIChild As Long, ByVal hMenu As Long)
  AddLogEntry "ReBarA_CleanupMDIChildSystemMenu: hActiveMDIChild=0x" & Hex(hActiveMDIChild) & ", hMenu=0x" & Hex(hMenu)
End Sub

Private Sub ReBarA_Click(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_Click: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_Click: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_ContextMenu(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_ContextMenu: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_ContextMenu: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_CustomDraw(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal drawStage As TBarCtlsLibACtl.CustomDrawStageConstants, ByVal bandState As TBarCtlsLibACtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibACtl.RECTANGLE, furtherProcessing As TBarCtlsLibACtl.CustomDrawReturnValuesConstants)
'  If band Is Nothing Then
'    AddLogEntry "ReBarA_CustomDraw: band=Nothing, drawStage=0x" & Hex(drawStage) & ", bandState=0x" & Hex(bandState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  Else
'    AddLogEntry "ReBarA_CustomDraw: band=" & band.Text & ", drawStage=0x" & Hex(drawStage) & ", bandState=0x" & Hex(bandState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ReBarA_DblClick(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_DblClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_DblClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ReBarA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ReBarA_DragDrop(Source As Control, X As Single, Y As Single)
  AddLogEntry "ReBarA_DragDrop"
End Sub

Private Sub ReBarA_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "ReBarA_DragOver"
End Sub

Private Sub ReBarA_DraggingSplitter()
  AddLogEntry "ReBarA_DraggingSplitter"
End Sub

Private Sub ReBarA_FreeBandData(ByVal band As TBarCtlsLibACtl.IReBarBand)
  If band Is Nothing Then
    AddLogEntry "ReBarA_FreeBandData: band=Nothing"
  Else
    AddLogEntry "ReBarA_FreeBandData: band=" & band.Text
  End If
End Sub

Private Sub ReBarA_GotFocus()
  AddLogEntry "ReBarA_GotFocus"
  Set objActiveCtl = ReBarA
End Sub

Private Sub ReBarA_HeightChanged()
  AddLogEntry "ReBarA_HeightChanged"
  Form_Resize
End Sub

Private Sub ReBarA_InsertedBand(ByVal band As TBarCtlsLibACtl.IReBarBand)
  If band Is Nothing Then
    AddLogEntry "ReBarA_InsertedBand: band=Nothing"
  Else
    AddLogEntry "ReBarA_InsertedBand: band=" & band.Text
  End If
End Sub

Private Sub ReBarA_InsertingBand(ByVal band As TBarCtlsLibACtl.IVirtualReBarBand, cancelInsertion As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarA_InsertingBand: band=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ReBarA_InsertingBand: band=" & band.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ReBarA_LayoutChanged()
  AddLogEntry "ReBarA_LayoutChanged"
End Sub

Private Sub ReBarA_LostFocus()
  AddLogEntry "ReBarA_LostFocus"
End Sub

Private Sub ReBarA_MClick(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_MClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_MClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_MDblClick(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_MDblClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_MDblClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_MouseDown(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_MouseDown: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_MouseDown: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_MouseEnter(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_MouseEnter: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_MouseEnter: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_MouseHover(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_MouseHover: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_MouseHover: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_MouseLeave(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_MouseLeave: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_MouseLeave: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_MouseMove(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
'  If band Is Nothing Then
'    AddLogEntry "ReBarA_MouseMove: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  Else
'    AddLogEntry "ReBarA_MouseMove: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  End If
End Sub

Private Sub ReBarA_MouseUp(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_MouseUp: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_MouseUp: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_NonClientHitTest(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants, returnValue As Long)
'  If band Is Nothing Then
'    AddLogEntry "ReBarA_NonClientHitTest: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", returnValue=0x" & Hex(returnValue)
'  Else
'    AddLogEntry "ReBarA_NonClientHitTest: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", returnValue=0x" & Hex(returnValue)
'  End If
End Sub

Private Sub ReBarA_OLEDragDrop(ByVal data As TBarCtlsLibACtl.IOLEDataObject, effect As TBarCtlsLibACtl.OLEDropEffectConstants, ByVal dropTarget As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ReBarA_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ReBarA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ReBarA_OLEDragEnter(ByVal data As TBarCtlsLibACtl.IOLEDataObject, effect As TBarCtlsLibACtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ReBarA_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ReBarA_OLEDragLeave(ByVal data As TBarCtlsLibACtl.IOLEDataObject, ByVal dropTarget As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ReBarA_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ReBarA_OLEDragMouseMove(ByVal data As TBarCtlsLibACtl.IOLEDataObject, effect As TBarCtlsLibACtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ReBarA_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ReBarA_RawMenuMessage(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  AddLogEntry "ReBarA_RawMenuMessage: message=0x" & Hex(message) & ", wParam=0x" & Hex(wParam) & ", lParam=0x" & Hex(lParam) & ", result=" & result & ", handledEvent=" & handledEvent
End Sub

Private Sub ReBarA_RClick(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_RClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_RClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_RDblClick(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_RDblClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_RDblClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ReBarA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertBandsA
End Sub

Private Sub ReBarA_ReleasedMouseCapture()
  AddLogEntry "ReBarA_ReleasedMouseCapture"
End Sub

Private Sub ReBarA_RemovedBand(ByVal band As TBarCtlsLibACtl.IVirtualReBarBand)
  If band Is Nothing Then
    AddLogEntry "ReBarA_RemovedBand: band=Nothing"
  Else
    AddLogEntry "ReBarA_RemovedBand: band=" & band.Text
  End If
End Sub

Private Sub ReBarA_RemovingBand(ByVal band As TBarCtlsLibACtl.IReBarBand, cancelDeletion As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarA_RemovingBand: band=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ReBarA_RemovingBand: band=" & band.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ReBarA_ResizedControlWindow()
  AddLogEntry "ReBarA_ResizedControlWindow"
End Sub

Private Sub ReBarA_ResizingContainedWindow(ByVal band As TBarCtlsLibACtl.IReBarBand, bandClientRectangle As TBarCtlsLibACtl.RECTANGLE, containedWindowRectangle As TBarCtlsLibACtl.RECTANGLE)
  If band Is Nothing Then
    AddLogEntry "ReBarA_ResizingContainedWindow: band=Nothing, bandClientRectangle=(" & bandClientRectangle.Left & "," & bandClientRectangle.Top & ")-(" & bandClientRectangle.Right & "," & bandClientRectangle.Bottom & "), containedWindowRectangle=(" & containedWindowRectangle.Left & "," & containedWindowRectangle.Top & ")-(" & containedWindowRectangle.Right & "," & containedWindowRectangle.Bottom & ")"
  Else
    AddLogEntry "ReBarA_ResizingContainedWindow: band=" & band.Text & ", bandClientRectangle=(" & bandClientRectangle.Left & "," & bandClientRectangle.Top & ")-(" & bandClientRectangle.Right & "," & bandClientRectangle.Bottom & "), containedWindowRectangle=(" & containedWindowRectangle.Left & "," & containedWindowRectangle.Top & ")-(" & containedWindowRectangle.Right & "," & containedWindowRectangle.Bottom & ")"
  End If
End Sub

Private Sub ReBarA_SelectedMenuItem(ByVal commandIDOrSubMenuIndex As Long, ByVal menuItemState As TBarCtlsLibACtl.MenuItemStateConstants, ByVal hMenu As Long)
  AddLogEntry "ReBarA_SelectedMenuItem: commandIDOrSubMenuIndex=" & commandIDOrSubMenuIndex & ", menuItemState=0x" & Hex(menuItemState) & ", hMenu=0x" & Hex(hMenu)
End Sub

Private Sub ReBarA_TogglingBand(ByVal band As TBarCtlsLibACtl.IReBarBand, cancelToggling As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarA_TogglingBand: band=Nothing, cancelToggling=" & cancelToggling
  Else
    AddLogEntry "ReBarA_TogglingBand: band=" & band.Text & ", cancelToggling=" & cancelToggling
  End If
End Sub

Private Sub ReBarA_Validate(Cancel As Boolean)
  AddLogEntry "ReBarA_Validate"
End Sub

Private Sub ReBarA_XClick(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_XClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_XClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarA_XDblClick(ByVal band As TBarCtlsLibACtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarA_XDblClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarA_XDblClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_AutoBreakingBand(ByVal band As TBarCtlsLibUCtl.IReBarBand, doAutoBreak As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarU_AutoBreakingBand: band=Nothing, doAutoBreak=" & doAutoBreak
  Else
    AddLogEntry "ReBarU_AutoBreakingBand: band=" & band.Text & ", doAutoBreak=" & doAutoBreak
  End If
End Sub

Private Sub ReBarU_AutoSized(targetRectangle As TBarCtlsLibUCtl.RECTANGLE, actualRectangle As TBarCtlsLibUCtl.RECTANGLE, ByVal changedBandHeightOrStyle As Boolean)
  AddLogEntry "ReBarU_AutoSized: targetRectangle=(" & targetRectangle.Left & "," & targetRectangle.Top & ")-(" & targetRectangle.Right & "," & targetRectangle.Bottom & "), actualRectangle=(" & actualRectangle.Left & "," & actualRectangle.Top & ")-(" & actualRectangle.Right & "," & actualRectangle.Bottom & "), changedBandHeightOrStyle=" & changedBandHeightOrStyle
End Sub

Private Sub ReBarU_BandBeginDrag(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants, cancelDrag As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarU_BandBeginDrag: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", cancelDrag=" & cancelDrag
  Else
    AddLogEntry "ReBarU_BandBeginDrag: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", cancelDrag=" & cancelDrag
  End If
End Sub

Private Sub ReBarU_BandBeginRDrag(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants, cancelDrag As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarU_BandBeginRDrag: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", cancelDrag=" & cancelDrag
  Else
    AddLogEntry "ReBarU_BandBeginRDrag: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", cancelDrag=" & cancelDrag
  End If
End Sub

Private Sub ReBarU_BandEndDrag(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_BandEndDrag: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_BandEndDrag: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_BandMouseEnter(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_BandMouseEnter: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_BandMouseEnter: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_BandMouseLeave(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_BandMouseLeave: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_BandMouseLeave: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_BeforeDisplayMDIChildSystemMenu(ByVal hActiveMDIChild As Long, ByVal hMenu As Long, ByVal x As Single, ByVal y As Single, cancelMenu As Boolean)
  AddLogEntry "ReBarU_BeforeDisplayMDIChildSystemMenu: hActiveMDIChild=0x" & Hex(hActiveMDIChild) & ", hMenu=0x" & Hex(hMenu) & ", x=" & x & ", y=" & y & ", cancelMenu=" & cancelMenu
End Sub

Private Sub ReBarU_ChevronClick(ByVal band As TBarCtlsLibUCtl.IReBarBand, chevronRectangle As TBarCtlsLibUCtl.RECTANGLE, ByVal userData As Long, doDefault As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarU_ChevronClick: band=Nothing, chevronRectangle=(" & chevronRectangle.Left & "," & chevronRectangle.Top & ")-(" & chevronRectangle.Right & "," & chevronRectangle.Bottom & "), userData=" & userData & ", doDefault=" & doDefault
  Else
    AddLogEntry "ReBarU_ChevronClick: band=" & band.Text & ", chevronRectangle=(" & chevronRectangle.Left & "," & chevronRectangle.Top & ")-(" & chevronRectangle.Right & "," & chevronRectangle.Bottom & "), userData=" & userData & ", doDefault=" & doDefault
  End If
End Sub

Private Sub ReBarU_CleanupMDIChildSystemMenu(ByVal hActiveMDIChild As Long, ByVal hMenu As Long)
  AddLogEntry "ReBarU_CleanupMDIChildSystemMenu: hActiveMDIChild=0x" & Hex(hActiveMDIChild) & ", hMenu=0x" & Hex(hMenu)
End Sub

Private Sub ReBarU_Click(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_Click: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_Click: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_ContextMenu(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_ContextMenu: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_ContextMenu: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_CustomDraw(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal bandState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
'  If band Is Nothing Then
'    AddLogEntry "ReBarU_CustomDraw: band=Nothing, drawStage=0x" & Hex(drawStage) & ", bandState=0x" & Hex(bandState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  Else
'    AddLogEntry "ReBarU_CustomDraw: band=" & band.Text & ", drawStage=0x" & Hex(drawStage) & ", bandState=0x" & Hex(bandState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ReBarU_DblClick(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_DblClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_DblClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ReBarU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ReBarU_DragDrop(Source As Control, X As Single, Y As Single)
  AddLogEntry "ReBarU_DragDrop"
End Sub

Private Sub ReBarU_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "ReBarU_DragOver"
End Sub

Private Sub ReBarU_DraggingSplitter()
  AddLogEntry "ReBarU_DraggingSplitter"
End Sub

Private Sub ReBarU_FreeBandData(ByVal band As TBarCtlsLibUCtl.IReBarBand)
  If band Is Nothing Then
    AddLogEntry "ReBarU_FreeBandData: band=Nothing"
  Else
    AddLogEntry "ReBarU_FreeBandData: band=" & band.Text
  End If
End Sub

Private Sub ReBarU_GotFocus()
  AddLogEntry "ReBarU_GotFocus"
  Set objActiveCtl = ReBarU
End Sub

Private Sub ReBarU_HeightChanged()
  AddLogEntry "ReBarU_HeightChanged"
  Form_Resize
End Sub

Private Sub ReBarU_InsertedBand(ByVal band As TBarCtlsLibUCtl.IReBarBand)
  If band Is Nothing Then
    AddLogEntry "ReBarU_InsertedBand: band=Nothing"
  Else
    AddLogEntry "ReBarU_InsertedBand: band=" & band.Text
  End If
End Sub

Private Sub ReBarU_InsertingBand(ByVal band As TBarCtlsLibUCtl.IVirtualReBarBand, cancelInsertion As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarU_InsertingBand: band=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ReBarU_InsertingBand: band=" & band.Text & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ReBarU_LayoutChanged()
  AddLogEntry "ReBarU_LayoutChanged"
End Sub

Private Sub ReBarU_LostFocus()
  AddLogEntry "ReBarU_LostFocus"
End Sub

Private Sub ReBarU_MClick(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_MClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_MClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_MDblClick(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_MDblClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_MDblClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_MouseDown(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_MouseDown: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_MouseDown: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_MouseEnter(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_MouseEnter: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_MouseEnter: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_MouseHover(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_MouseHover: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_MouseHover: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_MouseLeave(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_MouseLeave: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_MouseLeave: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_MouseMove(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
'  If band Is Nothing Then
'    AddLogEntry "ReBarU_MouseMove: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  Else
'    AddLogEntry "ReBarU_MouseMove: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  End If
End Sub

Private Sub ReBarU_MouseUp(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_MouseUp: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_MouseUp: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_NonClientHitTest(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants, returnValue As Long)
'  If band Is Nothing Then
'    AddLogEntry "ReBarU_NonClientHitTest: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", returnValue=0x" & Hex(returnValue)
'  Else
'    AddLogEntry "ReBarU_NonClientHitTest: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails) & ", returnValue=0x" & Hex(returnValue)
'  End If
End Sub

Private Sub ReBarU_OLEDragDrop(ByVal data As TBarCtlsLibUCtl.IOLEDataObject, effect As TBarCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ReBarU_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ReBarU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ReBarU_OLEDragEnter(ByVal data As TBarCtlsLibUCtl.IOLEDataObject, effect As TBarCtlsLibUCtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ReBarU_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ReBarU_OLEDragLeave(ByVal data As TBarCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ReBarU_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ReBarU_OLEDragMouseMove(ByVal data As TBarCtlsLibUCtl.IOLEDataObject, effect As TBarCtlsLibUCtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ReBarU_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.Text
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ReBarU_RawMenuMessage(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  AddLogEntry "ReBarU_RawMenuMessage: message=0x" & Hex(message) & ", wParam=0x" & Hex(wParam) & ", lParam=0x" & Hex(lParam) & ", result=" & result & ", handledEvent=" & handledEvent
End Sub

Private Sub ReBarU_RClick(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_RClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_RClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_RDblClick(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_RDblClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_RDblClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ReBarU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertBandsU
End Sub

Private Sub ReBarU_ReleasedMouseCapture()
  AddLogEntry "ReBarU_ReleasedMouseCapture"
End Sub

Private Sub ReBarU_RemovedBand(ByVal band As TBarCtlsLibUCtl.IVirtualReBarBand)
  If band Is Nothing Then
    AddLogEntry "ReBarU_RemovedBand: band=Nothing"
  Else
    AddLogEntry "ReBarU_RemovedBand: band=" & band.Text
  End If
End Sub

Private Sub ReBarU_RemovingBand(ByVal band As TBarCtlsLibUCtl.IReBarBand, cancelDeletion As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarU_RemovingBand: band=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ReBarU_RemovingBand: band=" & band.Text & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ReBarU_ResizedControlWindow()
  AddLogEntry "ReBarU_ResizedControlWindow"
End Sub

Private Sub ReBarU_ResizingContainedWindow(ByVal band As TBarCtlsLibUCtl.IReBarBand, bandClientRectangle As TBarCtlsLibUCtl.RECTANGLE, containedWindowRectangle As TBarCtlsLibUCtl.RECTANGLE)
  If band Is Nothing Then
    AddLogEntry "ReBarU_ResizingContainedWindow: band=Nothing, bandClientRectangle=(" & bandClientRectangle.Left & "," & bandClientRectangle.Top & ")-(" & bandClientRectangle.Right & "," & bandClientRectangle.Bottom & "), containedWindowRectangle=(" & containedWindowRectangle.Left & "," & containedWindowRectangle.Top & ")-(" & containedWindowRectangle.Right & "," & containedWindowRectangle.Bottom & ")"
  Else
    AddLogEntry "ReBarU_ResizingContainedWindow: band=" & band.Text & ", bandClientRectangle=(" & bandClientRectangle.Left & "," & bandClientRectangle.Top & ")-(" & bandClientRectangle.Right & "," & bandClientRectangle.Bottom & "), containedWindowRectangle=(" & containedWindowRectangle.Left & "," & containedWindowRectangle.Top & ")-(" & containedWindowRectangle.Right & "," & containedWindowRectangle.Bottom & ")"
  End If
End Sub

Private Sub ReBarU_SelectedMenuItem(ByVal commandIDOrSubMenuIndex As Long, ByVal menuItemState As TBarCtlsLibUCtl.MenuItemStateConstants, ByVal hMenu As Long)
  AddLogEntry "ReBarU_SelectedMenuItem: commandIDOrSubMenuIndex=" & commandIDOrSubMenuIndex & ", menuItemState=0x" & Hex(menuItemState) & ", hMenu=0x" & Hex(hMenu)
End Sub

Private Sub ReBarU_TogglingBand(ByVal band As TBarCtlsLibUCtl.IReBarBand, cancelToggling As Boolean)
  If band Is Nothing Then
    AddLogEntry "ReBarU_TogglingBand: band=Nothing, cancelToggling=" & cancelToggling
  Else
    AddLogEntry "ReBarU_TogglingBand: band=" & band.Text & ", cancelToggling=" & cancelToggling
  End If
End Sub

Private Sub ReBarU_Validate(Cancel As Boolean)
  AddLogEntry "ReBarU_Validate"
End Sub

Private Sub ReBarU_XClick(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_XClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_XClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ReBarU_XDblClick(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal button As Integer, ByVal shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If band Is Nothing Then
    AddLogEntry "ReBarU_XDblClick: band=Nothing, button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ReBarU_XDblClick: band=" & band.Text & ", button=" & button & ", shift=" & shift & ", x=" & X & ", y=" & Y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_BeforeDisplayChevronPopup(ByVal hPopup As Long, ByVal x As Single, ByVal y As Single, ByVal isMenu As Boolean, cancelPopup As Boolean, commandToExecute As Long)
  AddLogEntry "tbMenu_BeforeDisplayChevronPopup: hPopup=0x" & Hex(hPopup) & ", x=" & x & ", y=" & y & ", isMenu=" & isMenu & ", cancelPopup=" & cancelPopup & ", commandToExecute=" & commandToExecute
End Sub

Private Sub tbMenu_BeginCustomization(ByVal restoringFromRegistry As Boolean, dontRestoreFromRegistry As Boolean)
  AddLogEntry "tbMenu_BeginCustomization: restoringFromRegistry=" & restoringFromRegistry & ", dontRestoreFromRegistry=" & dontRestoreFromRegistry
End Sub

Private Sub tbMenu_ButtonBeginDrag(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_ButtonBeginDrag: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_ButtonBeginDrag: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_ButtonBeginRDrag(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_ButtonBeginRDrag: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_ButtonBeginRDrag: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_ButtonGetDisplayInfo(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal requestedInfo As TBarCtlsLibUCtl.RequestedInfoConstants, IconIndex As Long, ImageListIndex As Long, ByVal maxButtonTextLength As Long, buttonText As String, dontAskAgain As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_ButtonGetDisplayInfo: toolButton=Nothing, requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", imageListIndex=" & ImageListIndex & ", maxButtonTextLength=" & maxButtonTextLength & ", buttonText=" & buttonText & ", dontAskAgain=" & dontAskAgain
  Else
    AddLogEntry "tbMenu_ButtonGetDisplayInfo: toolButton=" & toolButton.ID & ", requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", imageListIndex=" & ImageListIndex & ", maxButtonTextLength=" & maxButtonTextLength & ", buttonText=" & buttonText & ", dontAskAgain=" & dontAskAgain
  End If
End Sub

Private Sub tbMenu_ButtonGetDropDownMenu(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, hMenu As Long)
  Dim i As Long

  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_ButtonGetDropDownMenu: toolButton=Nothing, hMenu=0x" & Hex(hMenu)
  Else
    AddLogEntry "tbMenu_ButtonGetDropDownMenu: toolButton=" & toolButton.ID & ", hMenu=0x" & Hex(hMenu)
  End If

  For i = LBound(menuButtons) To UBound(menuButtons)
    If menuButtons(i).ID = toolButton.ID Then
      hMenu = menuButtons(i).hSubMenu
      Exit For
    End If
  Next i
End Sub

Private Sub tbMenu_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_ButtonGetInfoTipText: toolButton=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "tbMenu_ButtonGetInfoTipText: toolButton=" & toolButton.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  End If
End Sub

Private Sub tbMenu_ButtonMouseEnter(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_ButtonMouseEnter: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_ButtonMouseEnter: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_ButtonMouseLeave(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_ButtonMouseLeave: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_ButtonMouseLeave: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_Click(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_Click: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_Click: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_ContextMenu(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_ContextMenu: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_ContextMenu: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_CustomDraw(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal buttonState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, textRectangle As TBarCtlsLibUCtl.RECTANGLE, horizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
'  If toolButton Is Nothing Then
'    AddLogEntry "tbMenu_CustomDraw: toolButton=Nothing, normalTextColor=0x" & Hex(normalTextColor) & ", normalButtonBackColor=0x" & Hex(normalButtonBackColor) & ", normalBackgroundMode=" & normalBackgroundMode & ", hotTextColor=0x" & Hex(hotTextColor) & ", hotButtonBackColor=0x" & Hex(hotButtonBackColor) & ", markedTextBackColor=0x" & Hex(markedTextBackColor) & ", markedButtonBackColor=0x" & Hex(markedButtonBackColor) & ", markedBackgroundMode=" & markedBackgroundMode & ", drawStage=0x" & Hex(drawStage) & ", buttonState=0x" & Hex(buttonState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), horizontalIconCaptionGap=" & horizontalIconCaptionGap & ", furtherProcessing=" & furtherProcessing
'  Else
'    AddLogEntry "tbMenu_CustomDraw: toolButton=" & toolButton.ID & ", normalTextColor=0x" & Hex(normalTextColor) & ", normalButtonBackColor=0x" & Hex(normalButtonBackColor) & ", normalBackgroundMode=" & normalBackgroundMode & ", hotTextColor=0x" & Hex(hotTextColor) & ", hotButtonBackColor=0x" & Hex(hotButtonBackColor) & ", markedTextBackColor=0x" & Hex(markedTextBackColor) & ", markedButtonBackColor=0x" & Hex(markedButtonBackColor) & ", markedBackgroundMode=" & markedBackgroundMode & ", drawStage=0x" & Hex(drawStage) & ", buttonState=0x" & Hex(buttonState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), horizontalIconCaptionGap=" & horizontalIconCaptionGap & ", furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub tbMenu_CustomizedControl()
  AddLogEntry "tbMenu_CustomizedControl"
End Sub

Private Sub tbMenu_CustomizeDialogRemovingButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_CustomizeDialogRemovingButton: toolButton=Nothing"
  Else
    AddLogEntry "tbMenu_CustomizeDialogRemovingButton: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub tbMenu_DblClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_DblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_DblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "tbMenu_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub tbMenu_DestroyingChevronPopup(ByVal hPopup As Long, ByVal isMenu As Boolean)
  AddLogEntry "tbMenu_DestroyingChevronPopup: hPopup=0x" & Hex(hPopup) & ", isMenu=" & isMenu
End Sub

Private Sub tbMenu_DisplayCustomizationHelp()
  AddLogEntry "tbMenu_DisplayCustomizationHelp"
End Sub

Private Sub tbMenu_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "tbMenu_DragDrop"
End Sub

Private Sub tbMenu_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "tbMenu_DragOver"
End Sub

Private Sub tbMenu_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_DropDown: toolButton=Nothing, buttonRectangle=(" & buttonRectangle.Left & "," & buttonRectangle.Top & ")-(" & buttonRectangle.Right & "," & buttonRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
  Else
    AddLogEntry "tbMenu_DropDown: toolButton=" & toolButton.ID & ", buttonRectangle=(" & buttonRectangle.Left & "," & buttonRectangle.Top & ")-(" & buttonRectangle.Right & "," & buttonRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
  End If
End Sub

Private Sub tbMenu_EndCustomization()
  AddLogEntry "tbMenu_EndCustomization"
End Sub

Private Sub tbMenu_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  Const MFS_CHECKED = &H8
  Const MFS_DISABLED = &H3
  Const MFS_ENABLED = &H0
  Const MFS_UNCHECKED = &H0
  Const MFT_STRING = &H0
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_TYPE = &H10
  Dim idealWindowWidth As Long
  Dim isVisible As Boolean
  Dim miiData As MENUITEMINFO
  Dim newMenuItemName As String
  Dim toolbarBand As TBarCtlsLibUCtl.ReBarBand

  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_ExecuteCommand: commandID=" & commandID & ", toolButton=Nothing, commandOrigin=" & commandOrigin & ", forwardMessage=" & forwardMessage
  Else
    AddLogEntry "tbMenu_ExecuteCommand: commandID=" & commandID & ", toolButton=" & toolButton.ID & ", commandOrigin=" & commandOrigin & ", forwardMessage=" & forwardMessage
  End If

  If commandOrigin <> coButton Then
    Select Case commandID
      Case MENUID_FILEEXIT
        Unload Me

      Case MENUID_VIEWTOOLBAR
        ' toggle the setting
        Set toolbarBand = ReBarU.Bands(BANDID_TOOLBAR, bitID)
        If toolbarBand.hContainedWindow Then
          ' currently docked
          toolbarBand.Visible = Not toolbarBand.Visible
          isVisible = toolbarBand.Visible
        ElseIf Not (ctlHostWindow Is Nothing) Then
          ' currently undocked
          ctlHostWindow.Visible = Not ctlHostWindow.Visible
          isVisible = ctlHostWindow.Visible
        End If
        ' update the menu items
        With miiData
          .cbSize = LenB(miiData)
          .fMask = MIIM_STATE
          .fState = IIf(isVisible, MFS_CHECKED, MFS_UNCHECKED)
          SetMenuItemInfo menuButtons(1).hSubMenu, MENUID_VIEWTOOLBAR, 0, miiData
          .fState = IIf(isVisible, MFS_ENABLED, MFS_DISABLED)
          SetMenuItemInfo menuButtons(1).hSubMenu, MENUID_VIEWTOGGLEDOCK, 0, miiData
          SetMenuItemInfo menuButtons(1).hSubMenu, MENUID_VIEWCUSTOMIZE, 0, miiData
        End With

      Case MENUID_VIEWICONTEXT
        ' toggle the setting
        ToolBarU.AlwaysDisplayButtonText = Not ToolBarU.AlwaysDisplayButtonText
        ToolBarA.AlwaysDisplayButtonText = ToolBarU.AlwaysDisplayButtonText
        ' update the band/host window
        If ReBarU.Bands(BANDID_TOOLBAR, bitID).hContainedWindow Then
          ' currently docked
          ReBarU.Bands(BANDID_TOOLBAR, bitID).IdealWidth = ToolBarU.IdealWidth
        Else
          ' currently undocked
          If Not (ctlHostWindow Is Nothing) Then
            ctlHostWindow.CalculateWindowSize ToolBarU.IdealWidth, , idealWindowWidth
            ctlHostWindow.Move , , idealWindowWidth
          End If
        End If
        ' update the menu item
        With miiData
          .cbSize = LenB(miiData)
          .fMask = MIIM_STATE
          .fState = IIf(ToolBarU.AlwaysDisplayButtonText, MFS_CHECKED, MFS_UNCHECKED)
        End With
        SetMenuItemInfo menuButtons(1).hSubMenu, MENUID_VIEWICONTEXT, 0, miiData

      Case MENUID_VIEWTOGGLEDOCK
        If ReBarU.Bands(BANDID_TOOLBAR, bitID).hContainedWindow Then
          ' currently docked
          If ctlHostWindow Is Nothing Then
            Set ctlHostWindow = New TBarCtlsLibUCtl.ControlHostWindow
          End If
          If Not (ctlHostWindow Is Nothing) Then
            ctlHostWindow.WindowTitle = "ToolBar"
            ctlHostWindow.Create Me.hWnd
            ctlHostWindow.hWndChild = ToolBarU.hWnd
            With ReBarU.Bands(BANDID_TOOLBAR, bitID)
              .hContainedWindow = 0
              .Visible = False
            End With
            ' Setting the rebar band's hContainedWindow to 0 hides the previously assigned window!
            ToolBarU.Visible = False
            ToolBarU.Visible = True
            ToolBarU.WrapButtons = True
            newMenuItemName = "&Dock Tool Bar"
          End If
        Else
          ' currently not docked
          ToolBarU.WrapButtons = False
          With ReBarU.Bands(BANDID_TOOLBAR, bitID)
            .Visible = True
            .hContainedWindow = ToolBarU.hWnd
          End With
          ctlHostWindow.Destroy
          newMenuItemName = "&Undock Tool Bar"
        End If
        ' update menu item
        With miiData
          .cbSize = LenB(miiData)
          .fMask = MIIM_TYPE
          .fType = MFT_STRING
          .dwTypeData = StrPtr(newMenuItemName)
          .cch = Len(newMenuItemName)
        End With
        SetMenuItemInfo menuButtons(1).hSubMenu, MENUID_VIEWTOGGLEDOCK, 0, miiData

      Case MENUID_VIEWCUSTOMIZE
        ToolBarU.Customize

      Case MENUID_HELPABOUT
        cmdAbout_Click
    End Select
  End If
End Sub

Private Sub tbMenu_FreeButtonData(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_FreeButtonData: toolButton=Nothing"
  Else
    AddLogEntry "tbMenu_FreeButtonData: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub tbMenu_GetAvailableButton(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton, noMoreButtons As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_GetAvailableButton: toolButton=Nothing, noMoreButtons=" & noMoreButtons
  Else
    AddLogEntry "tbMenu_GetAvailableButton: toolButton=" & toolButton.Index & ", noMoreButtons=" & noMoreButtons
  End If
End Sub

Private Sub tbMenu_GotFocus()
  AddLogEntry "tbMenu_GotFocus"
  Set objActiveCtl = tbMenu
End Sub

Private Sub tbMenu_HotButtonChangeWrapping(ByVal previousHotButton As TBarCtlsLibUCtl.IToolBarButton, ByVal wrappingDirection As TBarCtlsLibUCtl.WrappingDirectionConstants, ByVal causedBy As TBarCtlsLibUCtl.HotButtonChangingCausedByConstants, cancelChange As Boolean)
  Dim str As String

  str = "tbMenu_HotButtonChangeWrapping: previousHotButton="
  If previousHotButton Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & previousHotButton.Text
  End If
  str = str & ", wrappingDirection=" & wrappingDirection & ", causedBy=0x" & Hex(causedBy) & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub tbMenu_HotButtonChanging(ByVal previousHotButton As TBarCtlsLibUCtl.IToolBarButton, ByVal newHotButton As TBarCtlsLibUCtl.IToolBarButton, ByVal causedBy As TBarCtlsLibUCtl.HotButtonChangingCausedByConstants, ByVal additionalInfo As TBarCtlsLibUCtl.HotButtonChangingAdditionalInfoConstants, cancelChange As Boolean)
  Dim str As String

  str = "tbMenu_HotButtonChanging: previousHotButton="
  If previousHotButton Is Nothing Then
    str = str & "Nothing, newHotItem="
  Else
    str = str & previousHotButton.Text & ", newHotButton="
  End If
  If newHotButton Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newHotButton.Text
  End If
  str = str & ", causedBy=0x" & Hex(causedBy) & ", additionalInfo=0x" & Hex(additionalInfo) & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub tbMenu_InitializeCustomizationDialog(ByVal hCustomizationDialog As Long, displayHelpButton As Boolean)
  AddLogEntry "tbMenu_InitializeCustomizationDialog: hCustomizationDialog=0x" & Hex(hCustomizationDialog) & ", displayHelpButton=" & displayHelpButton
End Sub

Private Sub tbMenu_InitializeToolBarStateRegistryRestorage(numberOfButtonsToLoad As Long, ByVal totalDataSize As Long, ByVal perButtonDataSize As Long, ByVal pDataStream As Long, pStartOfNextDataBlock As Long, cancelLoading As Boolean)
  AddLogEntry "tbMenu_InitializeToolBarStateRegistryRestorage: numberOfButtonsToLoad=" & numberOfButtonsToLoad & ", totalDataSize=" & totalDataSize & ", perButtonDataSize=" & perButtonDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfNextDataBlock=0x" & Hex(pStartOfNextDataBlock) & ", cancelLoading=" & cancelLoading
End Sub

Private Sub tbMenu_InitializeToolBarStateRegistryStorage(numberOfButtonsToSave As Long, totalDataSize As Long, pDataStream As Long, pStartOfUnusedSpace As Long)
  AddLogEntry "tbMenu_InitializeToolBarStateRegistryStorage: numberOfButtonsToSave=" & numberOfButtonsToSave & ", totalDataSize=" & totalDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfUnusedSpace=0x" & Hex(pStartOfUnusedSpace)
End Sub

Private Sub tbMenu_InsertedButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_InsertedButton: toolButton=Nothing"
  Else
    AddLogEntry "tbMenu_InsertedButton: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub tbMenu_InsertingButton(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton, cancelInsertion As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_InsertingButton: toolButton=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "tbMenu_InsertingButton: toolButton=" & toolButton.ID & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub tbMenu_IsDuplicateAccelerator(ByVal accelerator As Integer, isDuplicate As Boolean, handledEvent As Boolean)
  AddLogEntry "tbMenu_IsDuplicateAccelerator: accelerator=" & Chr$(accelerator) & ", isDuplicate=" & isDuplicate & ", handledEvent=" & handledEvent
End Sub

Private Sub tbMenu_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "tbMenu_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub tbMenu_KeyPress(keyAscii As Integer)
  AddLogEntry "tbMenu_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub tbMenu_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "tbMenu_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub tbMenu_LostFocus()
  AddLogEntry "tbMenu_LostFocus"
End Sub

Private Sub tbMenu_MapAccelerator(ByVal accelerator As Integer, ByVal startingPointOfSearch As TBarCtlsLibUCtl.IToolBarButton, ByVal resumingSearchWithFirstButton As Boolean, matchingButton As TBarCtlsLibUCtl.IToolBarButton, handledEvent As Boolean)
  Dim str As String

  str = "tbMenu_MapAccelerator: accelerator=" & Chr$(accelerator) & ", startingPointOfSearch="
  If startingPointOfSearch Is Nothing Then
    str = str & "Nothing, resumingSearchWithFirstButton=" & resumingSearchWithFirstButton
  Else
    str = str & startingPointOfSearch.ID & ", resumingSearchWithFirstButton=" & resumingSearchWithFirstButton
  End If
  str = str & ", matchingButton="
  If matchingButton Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & matchingButton.ID
  End If
  str = str & ", handledEvent=" & handledEvent

  AddLogEntry str
End Sub

Private Sub tbMenu_MClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_MClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_MClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_MDblClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_MDblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_MDblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_MouseDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_MouseDown: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_MouseDown: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_MouseEnter(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_MouseEnter: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_MouseEnter: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_MouseHover(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_MouseHover: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_MouseHover: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_MouseLeave(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_MouseLeave: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_MouseLeave: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_MouseMove(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
'  If toolButton Is Nothing Then
'    AddLogEntry "tbMenu_MouseMove: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  Else
'    AddLogEntry "tbMenu_MouseMove: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  End If
End Sub

Private Sub tbMenu_MouseUp(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_MouseUp: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_MouseUp: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_OLECompleteDrag(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject, ByVal performedEffect As TBarCtlsLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "tbMenu_OLECompleteDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub tbMenu_OLEDragDrop(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject, effect As TBarCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "tbMenu_OLEDragDrop: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    tbMenu.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub tbMenu_OLEDragEnter(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject, effect As TBarCtlsLibUCtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants, autoClickButton As Boolean)
  Dim files() As String
  Dim str As String

  str = "tbMenu_OLEDragEnter: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoClickButton=" & autoClickButton

  AddLogEntry str
End Sub

Private Sub tbMenu_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "tbMenu_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub tbMenu_OLEDragLeave(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "tbMenu_OLEDragLeave: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub tbMenu_OLEDragLeavePotentialTarget()
  AddLogEntry "tbMenu_OLEDragLeavePotentialTarget"
End Sub

Private Sub tbMenu_OLEDragMouseMove(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject, effect As TBarCtlsLibUCtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants, autoClickButton As Boolean)
  Dim files() As String
  Dim str As String

  str = "tbMenu_OLEDragMouseMove: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoClickButton=" & autoClickButton

  AddLogEntry str
End Sub

Private Sub tbMenu_OLEGiveFeedback(ByVal effect As TBarCtlsLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "tbMenu_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub tbMenu_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As TBarCtlsLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "tbMenu_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub tbMenu_OLEReceivedNewData(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "tbMenu_OLEReceivedNewData: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub tbMenu_OLESetData(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "tbMenu_OLESetData: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub tbMenu_OLEStartDrag(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "tbMenu_OLEStartDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub tbMenu_OpenPopupMenu(ByVal hMenu As Long, ByVal parentMenuItemIndex As Long, ByVal isSystemMenu As Boolean)
  AddLogEntry "tbMenu_OpenPopupMenu: hMenu=0x" & Hex(hMenu) & ", parentMenuItemIndex=" & parentMenuItemIndex & ", isSystemMenu=" & isSystemMenu
End Sub

Private Sub tbMenu_QueryInsertButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, allowInsertionToLeft As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_QueryInsertButton: toolButton=Nothing, allowInsertionToLeft=" & allowInsertionToLeft
  Else
    AddLogEntry "tbMenu_QueryInsertButton: toolButton=" & toolButton.ID & ", allowInsertionToLeft=" & allowInsertionToLeft
  End If
End Sub

Private Sub tbMenu_QueryRemoveButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, allowRemoval As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_QueryRemoveButton: toolButton=Nothing, allowRemoval=" & allowRemoval
  Else
    AddLogEntry "tbMenu_QueryRemoveButton: toolButton=" & toolButton.ID & ", allowRemoval=" & allowRemoval
  End If
End Sub

Private Sub tbMenu_RawMenuMessage(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  AddLogEntry "tbMenu_RawMenuMessage: message=0x" & Hex(message) & ", wParam=0x" & Hex(wParam) & ", lParam=0x" & Hex(lParam) & ", result=" & result & ", handledEvent=" & handledEvent
End Sub

Private Sub tbMenu_RClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_RClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_RClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_RDblClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_RDblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_RDblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "tbMenu_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertButtonsU
End Sub

Private Sub tbMenu_RemovedButton(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_RemovedButton: toolButton=Nothing"
  Else
    AddLogEntry "tbMenu_RemovedButton: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub tbMenu_RemovingButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, cancelDeletion As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_RemovingButton: toolButton=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "tbMenu_RemovingButton: toolButton=" & toolButton.ID & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub tbMenu_ResetCustomizations(ByVal hCustomizationDialog As Long, EndCustomization As Boolean)
  AddLogEntry "tbMenu_ResetCustomizations: hCustomizationDialog=0x" & Hex(hCustomizationDialog) & ", EndCustomization=" & EndCustomization
  InsertButtonsU
End Sub

Private Sub tbMenu_ResizedControlWindow()
  AddLogEntry "tbMenu_ResizedControlWindow"
End Sub

Private Sub tbMenu_RestoreButtonFromRegistryStream(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton, ByVal numberOfButtonsToLoad As Long, ByVal totalDataSize As Long, ByVal perButtonDataSize As Long, ByVal pDataStream As Long, pStartOfNextDataBlock As Long)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_RestoreButtonFromRegistryStream: toolButton=Nothing, numberOfButtonsToLoad=" & numberOfButtonsToLoad & ", totalDataSize=" & totalDataSize & ", perButtonDataSize=" & perButtonDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfNextDataBlock=0x" & Hex(pStartOfNextDataBlock)
  Else
    AddLogEntry "tbMenu_RestoreButtonFromRegistryStream: toolButton=" & toolButton.ID & ", numberOfButtonsToLoad=" & numberOfButtonsToLoad & ", totalDataSize=" & totalDataSize & ", perButtonDataSize=" & perButtonDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfNextDataBlock=0x" & Hex(pStartOfNextDataBlock)
  End If
End Sub

Private Sub tbMenu_SaveButtonToRegistryStream(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton, ByVal totalDataSize As Long, ByVal pDataStream As Long, pStartOfUnusedSpace As Long)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_SaveButtonToRegistryStream: toolButton=Nothing, totalDataSize=" & totalDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfUnusedSpace=0x" & Hex(pStartOfUnusedSpace)
  Else
    AddLogEntry "tbMenu_SaveButtonToRegistryStream: toolButton=" & toolButton.ID & ", totalDataSize=" & totalDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfUnusedSpace=0x" & Hex(pStartOfUnusedSpace)
  End If
End Sub

Private Sub tbMenu_SelectedMenuItem(ByVal commandIDOrSubMenuIndex As Long, ByVal menuItemState As TBarCtlsLibUCtl.MenuItemStateConstants, ByVal hMenu As Long)
  AddLogEntry "tbMenu_SelectedMenuItem: commandIDOrSubMenuIndex=" & commandIDOrSubMenuIndex & ", menuItemState=0x" & Hex(menuItemState) & ", hMenu=0x" & Hex(hMenu)
End Sub

Private Sub tbMenu_Validate(Cancel As Boolean)
  AddLogEntry "tbMenu_Validate"
End Sub

Private Sub tbMenu_XClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_XClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_XClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub tbMenu_XDblClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "tbMenu_XDblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "tbMenu_XDblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_BeforeDisplayChevronPopup(ByVal hPopup As Long, ByVal x As Single, ByVal y As Single, ByVal isMenu As Boolean, cancelPopup As Boolean, commandToExecute As Long)
  AddLogEntry "ToolBarA_BeforeDisplayChevronPopup: hPopup=0x" & Hex(hPopup) & ", x=" & x & ", y=" & y & ", isMenu=" & isMenu & ", cancelPopup=" & cancelPopup & ", commandToExecute=" & commandToExecute
End Sub

Private Sub ToolBarA_BeginCustomization(ByVal restoringFromRegistry As Boolean, dontRestoreFromRegistry As Boolean)
  AddLogEntry "ToolBarA_BeginCustomization: restoringFromRegistry=" & restoringFromRegistry & ", dontRestoreFromRegistry=" & dontRestoreFromRegistry
End Sub

Private Sub ToolBarA_ButtonBeginDrag(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_ButtonBeginDrag: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_ButtonBeginDrag: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_ButtonBeginRDrag(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_ButtonBeginRDrag: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_ButtonBeginRDrag: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_ButtonGetDisplayInfo(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal requestedInfo As TBarCtlsLibACtl.RequestedInfoConstants, IconIndex As Long, ImageListIndex As Long, ByVal maxButtonTextLength As Long, buttonText As String, dontAskAgain As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_ButtonGetDisplayInfo: toolButton=Nothing, requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", imageListIndex=" & ImageListIndex & ", maxButtonTextLength=" & maxButtonTextLength & ", buttonText=" & buttonText & ", dontAskAgain=" & dontAskAgain
  Else
    AddLogEntry "ToolBarA_ButtonGetDisplayInfo: toolButton=" & toolButton.ID & ", requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", imageListIndex=" & ImageListIndex & ", maxButtonTextLength=" & maxButtonTextLength & ", buttonText=" & buttonText & ", dontAskAgain=" & dontAskAgain
  End If
End Sub

Private Sub ToolBarA_ButtonGetDropDownMenu(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, hMenu As Long)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_ButtonGetDropDownMenu: toolButton=Nothing, hMenu=0x" & Hex(hMenu)
  Else
    AddLogEntry "ToolBarA_ButtonGetDropDownMenu: toolButton=" & toolButton.ID & ", hMenu=0x" & Hex(hMenu)
  End If
End Sub

Private Sub ToolBarA_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_ButtonGetInfoTipText: toolButton=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "ToolBarA_ButtonGetInfoTipText: toolButton=" & toolButton.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
    If infoTipText <> "" Then
      infoTipText = infoTipText & vbNewLine & "ID: " & toolButton.ID & vbNewLine & "ButtonData: " & toolButton.ButtonData
    Else
      infoTipText = "ID: " & toolButton.ID & vbNewLine & "ButtonData: " & toolButton.ButtonData
    End If
  End If
End Sub

Private Sub ToolBarA_ButtonMouseEnter(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_ButtonMouseEnter: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_ButtonMouseEnter: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_ButtonMouseLeave(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_ButtonMouseLeave: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_ButtonMouseLeave: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_Click(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_Click: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_Click: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_ContextMenu(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_ContextMenu: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_ContextMenu: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_CustomDraw(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As TBarCtlsLibACtl.StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As TBarCtlsLibACtl.StringBackgroundModeConstants, ByVal drawStage As TBarCtlsLibACtl.CustomDrawStageConstants, ByVal buttonState As TBarCtlsLibACtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibACtl.RECTANGLE, textRectangle As TBarCtlsLibACtl.RECTANGLE, horizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As TBarCtlsLibACtl.CustomDrawReturnValuesConstants)
'  If toolButton Is Nothing Then
'    AddLogEntry "ToolBarA_CustomDraw: toolButton=Nothing, normalTextColor=0x" & Hex(normalTextColor) & ", normalButtonBackColor=0x" & Hex(normalButtonBackColor) & ", normalBackgroundMode=" & normalBackgroundMode & ", hotTextColor=0x" & Hex(hotTextColor) & ", hotButtonBackColor=0x" & Hex(hotButtonBackColor) & ", markedTextBackColor=0x" & Hex(markedTextBackColor) & ", markedButtonBackColor=0x" & Hex(markedButtonBackColor) & ", markedBackgroundMode=" & markedBackgroundMode & ", drawStage=0x" & Hex(drawStage) & ", buttonState=0x" & Hex(buttonState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), horizontalIconCaptionGap=" & horizontalIconCaptionGap & ", furtherProcessing=" & furtherProcessing
'  Else
'    AddLogEntry "ToolBarA_CustomDraw: toolButton=" & toolButton.ID & ", normalTextColor=0x" & Hex(normalTextColor) & ", normalButtonBackColor=0x" & Hex(normalButtonBackColor) & ", normalBackgroundMode=" & normalBackgroundMode & ", hotTextColor=0x" & Hex(hotTextColor) & ", hotButtonBackColor=0x" & Hex(hotButtonBackColor) & ", markedTextBackColor=0x" & Hex(markedTextBackColor) & ", markedButtonBackColor=0x" & Hex(markedButtonBackColor) & ", markedBackgroundMode=" & markedBackgroundMode & ", drawStage=0x" & Hex(drawStage) & ", buttonState=0x" & Hex(buttonState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), horizontalIconCaptionGap=" & horizontalIconCaptionGap & ", furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ToolBarA_CustomizedControl()
  AddLogEntry "ToolBarA_CustomizedControl"
End Sub

Private Sub ToolBarA_CustomizeDialogRemovingButton(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_CustomizeDialogRemovingButton: toolButton=Nothing"
  Else
    AddLogEntry "ToolBarA_CustomizeDialogRemovingButton: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub ToolBarA_DblClick(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_DblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_DblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ToolBarA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ToolBarA_DestroyingChevronPopup(ByVal hPopup As Long, ByVal isMenu As Boolean)
  AddLogEntry "ToolBarA_DestroyingChevronPopup: hPopup=0x" & Hex(hPopup) & ", isMenu=" & isMenu
End Sub

Private Sub ToolBarA_DisplayCustomizationHelp()
  AddLogEntry "ToolBarA_DisplayCustomizationHelp"
End Sub

Private Sub ToolBarA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "ToolBarA_DragDrop"
End Sub

Private Sub ToolBarA_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "ToolBarA_DragOver"
End Sub

Private Sub ToolBarA_DropDown(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, buttonRectangle As TBarCtlsLibACtl.RECTANGLE, furtherProcessing As TBarCtlsLibACtl.DropDownReturnValuesConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_DropDown: toolButton=Nothing, buttonRectangle=(" & buttonRectangle.Left & "," & buttonRectangle.Top & ")-(" & buttonRectangle.Right & "," & buttonRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
  Else
    AddLogEntry "ToolBarA_DropDown: toolButton=" & toolButton.ID & ", buttonRectangle=(" & buttonRectangle.Left & "," & buttonRectangle.Top & ")-(" & buttonRectangle.Right & "," & buttonRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
  End If
End Sub

Private Sub ToolBarA_EndCustomization()
  AddLogEntry "ToolBarA_EndCustomization"
End Sub

Private Sub ToolBarA_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibACtl.CommandOriginConstants, forwardMessage As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_ExecuteCommand: commandID=" & commandID & ", toolButton=Nothing, commandOrigin=" & commandOrigin & ", forwardMessage=" & forwardMessage
  Else
    AddLogEntry "ToolBarA_ExecuteCommand: commandID=" & commandID & ", toolButton=" & toolButton.ID & ", commandOrigin=" & commandOrigin & ", forwardMessage=" & forwardMessage
  End If
End Sub

Private Sub ToolBarA_FreeButtonData(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_FreeButtonData: toolButton=Nothing"
  Else
    AddLogEntry "ToolBarA_FreeButtonData: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub ToolBarA_GetAvailableButton(ByVal toolButton As TBarCtlsLibACtl.IVirtualToolBarButton, noMoreButtons As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_GetAvailableButton: toolButton=Nothing, noMoreButtons=" & noMoreButtons
  Else
    AddLogEntry "ToolBarA_GetAvailableButton: toolButton=" & toolButton.Index & ", noMoreButtons=" & noMoreButtons
  End If

  If toolButton.Index <= UBound(buttonsA) Then
    toolButton.ButtonData = buttonsA(toolButton.Index).ButtonData
    toolButton.IconIndex = buttonsA(toolButton.Index).IconIndex
    toolButton.ID = buttonsA(toolButton.Index).ID
    toolButton.Text = buttonsA(toolButton.Index).Text
  End If
  noMoreButtons = (toolButton.Index > UBound(buttonsA))
End Sub

Private Sub ToolBarA_GotFocus()
  AddLogEntry "ToolBarA_GotFocus"
  Set objActiveCtl = ToolBarA
End Sub

Private Sub ToolBarA_HotButtonChangeWrapping(ByVal previousHotButton As TBarCtlsLibACtl.IToolBarButton, ByVal wrappingDirection As TBarCtlsLibACtl.WrappingDirectionConstants, ByVal causedBy As TBarCtlsLibACtl.HotButtonChangingCausedByConstants, cancelChange As Boolean)
  Dim str As String

  str = "ToolBarA_HotButtonChangeWrapping: previousHotButton="
  If previousHotButton Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & previousHotButton.Text
  End If
  str = str & ", wrappingDirection=" & wrappingDirection & ", causedBy=0x" & Hex(causedBy) & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub ToolBarA_HotButtonChanging(ByVal previousHotButton As TBarCtlsLibACtl.IToolBarButton, ByVal newHotButton As TBarCtlsLibACtl.IToolBarButton, ByVal causedBy As TBarCtlsLibACtl.HotButtonChangingCausedByConstants, ByVal additionalInfo As TBarCtlsLibACtl.HotButtonChangingAdditionalInfoConstants, cancelChange As Boolean)
  Dim str As String

  str = "ToolBarA_HotButtonChanging: previousHotButton="
  If previousHotButton Is Nothing Then
    str = str & "Nothing, newHotItem="
  Else
    str = str & previousHotButton.Text & ", newHotButton="
  End If
  If newHotButton Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newHotButton.Text
  End If
  str = str & ", causedBy=0x" & Hex(causedBy) & ", additionalInfo=0x" & Hex(additionalInfo) & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub ToolBarA_InitializeCustomizationDialog(ByVal hCustomizationDialog As Long, displayHelpButton As Boolean)
  AddLogEntry "ToolBarA_InitializeCustomizationDialog: hCustomizationDialog=0x" & Hex(hCustomizationDialog) & ", displayHelpButton=" & displayHelpButton
End Sub

Private Sub ToolBarA_InitializeToolBarStateRegistryRestorage(numberOfButtonsToLoad As Long, ByVal totalDataSize As Long, ByVal perButtonDataSize As Long, ByVal pDataStream As Long, pStartOfNextDataBlock As Long, cancelLoading As Boolean)
  AddLogEntry "ToolBarA_InitializeToolBarStateRegistryRestorage: numberOfButtonsToLoad=" & numberOfButtonsToLoad & ", totalDataSize=" & totalDataSize & ", perButtonDataSize=" & perButtonDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfNextDataBlock=0x" & Hex(pStartOfNextDataBlock) & ", cancelLoading=" & cancelLoading
End Sub

Private Sub ToolBarA_InitializeToolBarStateRegistryStorage(numberOfButtonsToSave As Long, totalDataSize As Long, pDataStream As Long, pStartOfUnusedSpace As Long)
  AddLogEntry "ToolBarA_InitializeToolBarStateRegistryStorage: numberOfButtonsToSave=" & numberOfButtonsToSave & ", totalDataSize=" & totalDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfUnusedSpace=0x" & Hex(pStartOfUnusedSpace)
End Sub

Private Sub ToolBarA_InsertedButton(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_InsertedButton: toolButton=Nothing"
  Else
    AddLogEntry "ToolBarA_InsertedButton: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub ToolBarA_InsertingButton(ByVal toolButton As TBarCtlsLibACtl.IVirtualToolBarButton, cancelInsertion As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_InsertingButton: toolButton=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ToolBarA_InsertingButton: toolButton=" & toolButton.ID & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ToolBarA_IsDuplicateAccelerator(ByVal accelerator As Integer, isDuplicate As Boolean, handledEvent As Boolean)
  AddLogEntry "ToolBarA_IsDuplicateAccelerator: accelerator=" & Chr$(accelerator) & ", isDuplicate=" & isDuplicate & ", handledEvent=" & handledEvent
End Sub

Private Sub ToolBarA_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ToolBarA_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ToolBarA_KeyPress(keyAscii As Integer)
  AddLogEntry "ToolBarA_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ToolBarA_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ToolBarA_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ToolBarA_LostFocus()
  AddLogEntry "ToolBarA_LostFocus"
End Sub

Private Sub ToolBarA_MapAccelerator(ByVal accelerator As Integer, ByVal startingPointOfSearch As TBarCtlsLibACtl.IToolBarButton, ByVal resumingSearchWithFirstButton As Boolean, matchingButton As TBarCtlsLibACtl.IToolBarButton, handledEvent As Boolean)
  Dim str As String

  str = "ToolBarA_MapAccelerator: accelerator=" & Chr$(accelerator) & ", startingPointOfSearch="
  If startingPointOfSearch Is Nothing Then
    str = str & "Nothing, resumingSearchWithFirstButton=" & resumingSearchWithFirstButton
  Else
    str = str & startingPointOfSearch.ID & ", resumingSearchWithFirstButton=" & resumingSearchWithFirstButton
  End If
  str = str & ", matchingButton="
  If matchingButton Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & matchingButton.ID
  End If
  str = str & ", handledEvent=" & handledEvent

  AddLogEntry str
End Sub

Private Sub ToolBarA_MClick(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_MClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_MClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_MDblClick(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_MDblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_MDblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_MouseDown(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_MouseDown: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_MouseDown: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_MouseEnter(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_MouseEnter: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_MouseEnter: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_MouseHover(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_MouseHover: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_MouseHover: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_MouseLeave(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_MouseLeave: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_MouseLeave: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_MouseMove(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
'  If toolButton Is Nothing Then
'    AddLogEntry "ToolBarA_MouseMove: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  Else
'    AddLogEntry "ToolBarA_MouseMove: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  End If
End Sub

Private Sub ToolBarA_MouseUp(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_MouseUp: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_MouseUp: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_OLECompleteDrag(ByVal Data As TBarCtlsLibACtl.IOLEDataObject, ByVal performedEffect As TBarCtlsLibACtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ToolBarA_OLECompleteDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ToolBarA_OLEDragDrop(ByVal data As TBarCtlsLibACtl.IOLEDataObject, effect As TBarCtlsLibACtl.OLEDropEffectConstants, ByVal dropTarget As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ToolBarA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ToolBarA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ToolBarA_OLEDragEnter(ByVal data As TBarCtlsLibACtl.IOLEDataObject, effect As TBarCtlsLibACtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants, autoClickButton As Boolean)
  Dim files() As String
  Dim str As String

  str = "ToolBarA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoClickButton=" & autoClickButton

  AddLogEntry str
End Sub

Private Sub ToolBarA_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ToolBarA_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ToolBarA_OLEDragLeave(ByVal data As TBarCtlsLibACtl.IOLEDataObject, ByVal dropTarget As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ToolBarA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ToolBarA_OLEDragLeavePotentialTarget()
  AddLogEntry "ToolBarA_OLEDragLeavePotentialTarget"
End Sub

Private Sub ToolBarA_OLEDragMouseMove(ByVal data As TBarCtlsLibACtl.IOLEDataObject, effect As TBarCtlsLibACtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants, autoClickButton As Boolean)
  Dim files() As String
  Dim str As String

  str = "ToolBarA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoClickButton=" & autoClickButton

  AddLogEntry str
End Sub

Private Sub ToolBarA_OLEGiveFeedback(ByVal effect As TBarCtlsLibACtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ToolBarA_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ToolBarA_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As TBarCtlsLibACtl.OLEActionToContinueWithConstants)
  AddLogEntry "ToolBarA_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ToolBarA_OLEReceivedNewData(ByVal data As TBarCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ToolBarA_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ToolBarA_OLESetData(ByVal Data As TBarCtlsLibACtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ToolBarA_OLESetData: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ToolBarA_OLEStartDrag(ByVal Data As TBarCtlsLibACtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ToolBarA_OLEStartDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ToolBarA_OpenPopupMenu(ByVal hMenu As Long, ByVal parentMenuItemIndex As Long, ByVal isSystemMenu As Boolean)
  AddLogEntry "ToolBarA_OpenPopupMenu: hMenu=0x" & Hex(hMenu) & ", parentMenuItemIndex=" & parentMenuItemIndex & ", isSystemMenu=" & isSystemMenu
End Sub

Private Sub ToolBarA_QueryInsertButton(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, allowInsertionToLeft As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_QueryInsertButton: toolButton=Nothing, allowInsertionToLeft=" & allowInsertionToLeft
  Else
    AddLogEntry "ToolBarA_QueryInsertButton: toolButton=" & toolButton.ID & ", allowInsertionToLeft=" & allowInsertionToLeft
  End If
End Sub

Private Sub ToolBarA_QueryRemoveButton(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, allowRemoval As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_QueryRemoveButton: toolButton=Nothing, allowRemoval=" & allowRemoval
  Else
    AddLogEntry "ToolBarA_QueryRemoveButton: toolButton=" & toolButton.ID & ", allowRemoval=" & allowRemoval
  End If
End Sub

Private Sub ToolBarA_RawMenuMessage(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  AddLogEntry "ToolBarA_RawMenuMessage: message=0x" & Hex(message) & ", wParam=0x" & Hex(wParam) & ", lParam=0x" & Hex(lParam) & ", result=" & result & ", handledEvent=" & handledEvent
End Sub

Private Sub ToolBarA_RClick(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_RClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_RClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_RDblClick(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_RDblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_RDblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ToolBarA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertButtonsA
End Sub

Private Sub ToolBarA_RemovedButton(ByVal toolButton As TBarCtlsLibACtl.IVirtualToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_RemovedButton: toolButton=Nothing"
  Else
    AddLogEntry "ToolBarA_RemovedButton: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub ToolBarA_RemovingButton(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, cancelDeletion As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_RemovingButton: toolButton=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ToolBarA_RemovingButton: toolButton=" & toolButton.ID & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ToolBarA_ResetCustomizations(ByVal hCustomizationDialog As Long, EndCustomization As Boolean)
  AddLogEntry "ToolBarA_ResetCustomizations: hCustomizationDialog=0x" & Hex(hCustomizationDialog) & ", EndCustomization=" & EndCustomization
  InsertButtonsA
End Sub

Private Sub ToolBarA_ResizedControlWindow()
  AddLogEntry "ToolBarA_ResizedControlWindow"
End Sub

Private Sub ToolBarA_RestoreButtonFromRegistryStream(ByVal toolButton As TBarCtlsLibACtl.IVirtualToolBarButton, ByVal numberOfButtonsToLoad As Long, ByVal totalDataSize As Long, ByVal perButtonDataSize As Long, ByVal pDataStream As Long, pStartOfNextDataBlock As Long)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_RestoreButtonFromRegistryStream: toolButton=Nothing, numberOfButtonsToLoad=" & numberOfButtonsToLoad & ", totalDataSize=" & totalDataSize & ", perButtonDataSize=" & perButtonDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfNextDataBlock=0x" & Hex(pStartOfNextDataBlock)
  Else
    AddLogEntry "ToolBarA_RestoreButtonFromRegistryStream: toolButton=" & toolButton.ID & ", numberOfButtonsToLoad=" & numberOfButtonsToLoad & ", totalDataSize=" & totalDataSize & ", perButtonDataSize=" & perButtonDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfNextDataBlock=0x" & Hex(pStartOfNextDataBlock)
  End If
End Sub

Private Sub ToolBarA_SaveButtonToRegistryStream(ByVal toolButton As TBarCtlsLibACtl.IVirtualToolBarButton, ByVal totalDataSize As Long, ByVal pDataStream As Long, pStartOfUnusedSpace As Long)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_SaveButtonToRegistryStream: toolButton=Nothing, totalDataSize=" & totalDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfUnusedSpace=0x" & Hex(pStartOfUnusedSpace)
  Else
    AddLogEntry "ToolBarA_SaveButtonToRegistryStream: toolButton=" & toolButton.ID & ", totalDataSize=" & totalDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfUnusedSpace=0x" & Hex(pStartOfUnusedSpace)
  End If
End Sub

Private Sub ToolBarA_SelectedMenuItem(ByVal commandIDOrSubMenuIndex As Long, ByVal menuItemState As TBarCtlsLibACtl.MenuItemStateConstants, ByVal hMenu As Long)
  AddLogEntry "ToolBarA_SelectedMenuItem: commandIDOrSubMenuIndex=" & commandIDOrSubMenuIndex & ", menuItemState=0x" & Hex(menuItemState) & ", hMenu=0x" & Hex(hMenu)
End Sub

Private Sub ToolBarA_Validate(Cancel As Boolean)
  AddLogEntry "ToolBarA_Validate"
End Sub

Private Sub ToolBarA_XClick(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_XClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_XClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarA_XDblClick(ByVal toolButton As TBarCtlsLibACtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibACtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarA_XDblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarA_XDblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_BeforeDisplayChevronPopup(ByVal hPopup As Long, ByVal x As Single, ByVal y As Single, ByVal isMenu As Boolean, cancelPopup As Boolean, commandToExecute As Long)
  AddLogEntry "ToolBarU_BeforeDisplayChevronPopup: hPopup=0x" & Hex(hPopup) & ", x=" & x & ", y=" & y & ", isMenu=" & isMenu & ", cancelPopup=" & cancelPopup & ", commandToExecute=" & commandToExecute
End Sub

Private Sub ToolBarU_BeginCustomization(ByVal restoringFromRegistry As Boolean, dontRestoreFromRegistry As Boolean)
  AddLogEntry "ToolBarU_BeginCustomization: restoringFromRegistry=" & restoringFromRegistry & ", dontRestoreFromRegistry=" & dontRestoreFromRegistry
End Sub

Private Sub ToolBarU_ButtonBeginDrag(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_ButtonBeginDrag: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_ButtonBeginDrag: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_ButtonBeginRDrag(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_ButtonBeginRDrag: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_ButtonBeginRDrag: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_ButtonGetDisplayInfo(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal requestedInfo As TBarCtlsLibUCtl.RequestedInfoConstants, IconIndex As Long, ImageListIndex As Long, ByVal maxButtonTextLength As Long, buttonText As String, dontAskAgain As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_ButtonGetDisplayInfo: toolButton=Nothing, requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", imageListIndex=" & ImageListIndex & ", maxButtonTextLength=" & maxButtonTextLength & ", buttonText=" & buttonText & ", dontAskAgain=" & dontAskAgain
  Else
    AddLogEntry "ToolBarU_ButtonGetDisplayInfo: toolButton=" & toolButton.ID & ", requestedInfo=" & requestedInfo & ", iconIndex=" & IconIndex & ", imageListIndex=" & ImageListIndex & ", maxButtonTextLength=" & maxButtonTextLength & ", buttonText=" & buttonText & ", dontAskAgain=" & dontAskAgain
  End If
End Sub

Private Sub ToolBarU_ButtonGetDropDownMenu(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, hMenu As Long)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_ButtonGetDropDownMenu: toolButton=Nothing, hMenu=0x" & Hex(hMenu)
  Else
    AddLogEntry "ToolBarU_ButtonGetDropDownMenu: toolButton=" & toolButton.ID & ", hMenu=0x" & Hex(hMenu)
  End If
End Sub

Private Sub ToolBarU_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_ButtonGetInfoTipText: toolButton=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "ToolBarU_ButtonGetInfoTipText: toolButton=" & toolButton.Text & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
    If infoTipText <> "" Then
      infoTipText = infoTipText & vbNewLine & "ID: " & toolButton.ID & vbNewLine & "ButtonData: " & toolButton.ButtonData
    Else
      infoTipText = "ID: " & toolButton.ID & vbNewLine & "ButtonData: " & toolButton.ButtonData
    End If
  End If
End Sub

Private Sub ToolBarU_ButtonMouseEnter(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_ButtonMouseEnter: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_ButtonMouseEnter: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_ButtonMouseLeave(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_ButtonMouseLeave: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_ButtonMouseLeave: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_Click(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_Click: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_Click: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_ContextMenu(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Const TPM_RETURNCMD = &H100&
  Const TPM_TOPALIGN = &H0&
  Const TPM_VERPOSANIMATION = &H1000&
  Const TPM_VERTICAL = &H40&
  Dim commandID As Long
  Dim ptMenu As POINTAPI

  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_ContextMenu: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_ContextMenu: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If

  ptMenu.x = Me.ScaleX(x, vbTwips, vbPixels)
  ptMenu.y = Me.ScaleY(y, vbTwips, vbPixels)
  ClientToScreen ToolBarU.hWnd, ptMenu
  commandID = TrackPopupMenuEx(menuButtons(1).hSubMenu, TPM_TOPALIGN Or TPM_RETURNCMD Or TPM_VERTICAL Or TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, ToolBarU.hWnd, 0)
  If commandID <> 0 Then
    tbMenu_ExecuteCommand commandID, Nothing, coMenu, False
  End If
End Sub

Private Sub ToolBarU_CustomDraw(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal buttonState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, textRectangle As TBarCtlsLibUCtl.RECTANGLE, horizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
'  If toolButton Is Nothing Then
'    AddLogEntry "ToolBarU_CustomDraw: toolButton=Nothing, normalTextColor=0x" & Hex(normalTextColor) & ", normalButtonBackColor=0x" & Hex(normalButtonBackColor) & ", normalBackgroundMode=" & normalBackgroundMode & ", hotTextColor=0x" & Hex(hotTextColor) & ", hotButtonBackColor=0x" & Hex(hotButtonBackColor) & ", markedTextBackColor=0x" & Hex(markedTextBackColor) & ", markedButtonBackColor=0x" & Hex(markedButtonBackColor) & ", markedBackgroundMode=" & markedBackgroundMode & ", drawStage=0x" & Hex(drawStage) & ", buttonState=0x" & Hex(buttonState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), horizontalIconCaptionGap=" & horizontalIconCaptionGap & ", furtherProcessing=" & furtherProcessing
'  Else
'    AddLogEntry "ToolBarU_CustomDraw: toolButton=" & toolButton.ID & ", normalTextColor=0x" & Hex(normalTextColor) & ", normalButtonBackColor=0x" & Hex(normalButtonBackColor) & ", normalBackgroundMode=" & normalBackgroundMode & ", hotTextColor=0x" & Hex(hotTextColor) & ", hotButtonBackColor=0x" & Hex(hotButtonBackColor) & ", markedTextBackColor=0x" & Hex(markedTextBackColor) & ", markedButtonBackColor=0x" & Hex(markedButtonBackColor) & ", markedBackgroundMode=" & markedBackgroundMode & ", drawStage=0x" & Hex(drawStage) & ", buttonState=0x" & Hex(buttonState) & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), textRectangle=(" & textRectangle.Left & "," & textRectangle.Top & ")-(" & textRectangle.Right & "," & textRectangle.Bottom & "), horizontalIconCaptionGap=" & horizontalIconCaptionGap & ", furtherProcessing=" & furtherProcessing
'  End If
End Sub

Private Sub ToolBarU_CustomizedControl()
  AddLogEntry "ToolBarU_CustomizedControl"
End Sub

Private Sub ToolBarU_CustomizeDialogRemovingButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_CustomizeDialogRemovingButton: toolButton=Nothing"
  Else
    AddLogEntry "ToolBarU_CustomizeDialogRemovingButton: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub ToolBarU_DblClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_DblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_DblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ToolBarU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub ToolBarU_DestroyingChevronPopup(ByVal hPopup As Long, ByVal isMenu As Boolean)
  AddLogEntry "ToolBarU_DestroyingChevronPopup: hPopup=0x" & Hex(hPopup) & ", isMenu=" & isMenu
End Sub

Private Sub ToolBarU_DisplayCustomizationHelp()
  AddLogEntry "ToolBarU_DisplayCustomizationHelp"
End Sub

Private Sub ToolBarU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "ToolBarU_DragDrop"
End Sub

Private Sub ToolBarU_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
  AddLogEntry "ToolBarU_DragOver"
End Sub

Private Sub ToolBarU_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_DropDown: toolButton=Nothing, buttonRectangle=(" & buttonRectangle.Left & "," & buttonRectangle.Top & ")-(" & buttonRectangle.Right & "," & buttonRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
  Else
    AddLogEntry "ToolBarU_DropDown: toolButton=" & toolButton.ID & ", buttonRectangle=(" & buttonRectangle.Left & "," & buttonRectangle.Top & ")-(" & buttonRectangle.Right & "," & buttonRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
  End If
End Sub

Private Sub ToolBarU_EndCustomization()
  AddLogEntry "ToolBarU_EndCustomization"
End Sub

Private Sub ToolBarU_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_ExecuteCommand: commandID=" & commandID & ", toolButton=Nothing, commandOrigin=" & commandOrigin & ", forwardMessage=" & forwardMessage
  Else
    AddLogEntry "ToolBarU_ExecuteCommand: commandID=" & commandID & ", toolButton=" & toolButton.ID & ", commandOrigin=" & commandOrigin & ", forwardMessage=" & forwardMessage
  End If
End Sub

Private Sub ToolBarU_FreeButtonData(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_FreeButtonData: toolButton=Nothing"
  Else
    AddLogEntry "ToolBarU_FreeButtonData: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub ToolBarU_GetAvailableButton(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton, noMoreButtons As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_GetAvailableButton: toolButton=Nothing, noMoreButtons=" & noMoreButtons
  Else
    AddLogEntry "ToolBarU_GetAvailableButton: toolButton=" & toolButton.Index & ", noMoreButtons=" & noMoreButtons
  End If

  If toolButton.Index <= UBound(buttonsU) Then
    toolButton.ButtonData = buttonsU(toolButton.Index).ButtonData
    toolButton.IconIndex = buttonsU(toolButton.Index).IconIndex
    toolButton.ID = buttonsU(toolButton.Index).ID
    toolButton.Text = buttonsU(toolButton.Index).Text
  End If
  noMoreButtons = (toolButton.Index > UBound(buttonsU))
End Sub

Private Sub ToolBarU_GotFocus()
  AddLogEntry "ToolBarU_GotFocus"
  Set objActiveCtl = ToolBarU
End Sub

Private Sub ToolBarU_HotButtonChangeWrapping(ByVal previousHotButton As TBarCtlsLibUCtl.IToolBarButton, ByVal wrappingDirection As TBarCtlsLibUCtl.WrappingDirectionConstants, ByVal causedBy As TBarCtlsLibUCtl.HotButtonChangingCausedByConstants, cancelChange As Boolean)
  Dim str As String

  str = "ToolBarU_HotButtonChangeWrapping: previousHotButton="
  If previousHotButton Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & previousHotButton.Text
  End If
  str = str & ", wrappingDirection=" & wrappingDirection & ", causedBy=0x" & Hex(causedBy) & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub ToolBarU_HotButtonChanging(ByVal previousHotButton As TBarCtlsLibUCtl.IToolBarButton, ByVal newHotButton As TBarCtlsLibUCtl.IToolBarButton, ByVal causedBy As TBarCtlsLibUCtl.HotButtonChangingCausedByConstants, ByVal additionalInfo As TBarCtlsLibUCtl.HotButtonChangingAdditionalInfoConstants, cancelChange As Boolean)
  Dim str As String

  str = "ToolBarU_HotButtonChanging: previousHotButton="
  If previousHotButton Is Nothing Then
    str = str & "Nothing, newHotItem="
  Else
    str = str & previousHotButton.Text & ", newHotButton="
  End If
  If newHotButton Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newHotButton.Text
  End If
  str = str & ", causedBy=0x" & Hex(causedBy) & ", additionalInfo=0x" & Hex(additionalInfo) & ", cancelChange=" & cancelChange

  AddLogEntry str
End Sub

Private Sub ToolBarU_InitializeCustomizationDialog(ByVal hCustomizationDialog As Long, displayHelpButton As Boolean)
  AddLogEntry "ToolBarU_InitializeCustomizationDialog: hCustomizationDialog=0x" & Hex(hCustomizationDialog) & ", displayHelpButton=" & displayHelpButton
End Sub

Private Sub ToolBarU_InitializeToolBarStateRegistryRestorage(numberOfButtonsToLoad As Long, ByVal totalDataSize As Long, ByVal perButtonDataSize As Long, ByVal pDataStream As Long, pStartOfNextDataBlock As Long, cancelLoading As Boolean)
  AddLogEntry "ToolBarU_InitializeToolBarStateRegistryRestorage: numberOfButtonsToLoad=" & numberOfButtonsToLoad & ", totalDataSize=" & totalDataSize & ", perButtonDataSize=" & perButtonDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfNextDataBlock=0x" & Hex(pStartOfNextDataBlock) & ", cancelLoading=" & cancelLoading
End Sub

Private Sub ToolBarU_InitializeToolBarStateRegistryStorage(numberOfButtonsToSave As Long, totalDataSize As Long, pDataStream As Long, pStartOfUnusedSpace As Long)
  AddLogEntry "ToolBarU_InitializeToolBarStateRegistryStorage: numberOfButtonsToSave=" & numberOfButtonsToSave & ", totalDataSize=" & totalDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfUnusedSpace=0x" & Hex(pStartOfUnusedSpace)
End Sub

Private Sub ToolBarU_InsertedButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_InsertedButton: toolButton=Nothing"
  Else
    AddLogEntry "ToolBarU_InsertedButton: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub ToolBarU_InsertingButton(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton, cancelInsertion As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_InsertingButton: toolButton=Nothing, cancelInsertion=" & cancelInsertion
  Else
    AddLogEntry "ToolBarU_InsertingButton: toolButton=" & toolButton.ID & ", cancelInsertion=" & cancelInsertion
  End If
End Sub

Private Sub ToolBarU_IsDuplicateAccelerator(ByVal accelerator As Integer, isDuplicate As Boolean, handledEvent As Boolean)
  AddLogEntry "ToolBarU_IsDuplicateAccelerator: accelerator=" & Chr$(accelerator) & ", isDuplicate=" & isDuplicate & ", handledEvent=" & handledEvent
End Sub

Private Sub ToolBarU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ToolBarU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ToolBarU_KeyPress(keyAscii As Integer)
  AddLogEntry "ToolBarU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub ToolBarU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "ToolBarU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub ToolBarU_LostFocus()
  AddLogEntry "ToolBarU_LostFocus"
End Sub

Private Sub ToolBarU_MapAccelerator(ByVal accelerator As Integer, ByVal startingPointOfSearch As TBarCtlsLibUCtl.IToolBarButton, ByVal resumingSearchWithFirstButton As Boolean, matchingButton As TBarCtlsLibUCtl.IToolBarButton, handledEvent As Boolean)
  Dim str As String

  str = "ToolBarU_MapAccelerator: accelerator=" & Chr$(accelerator) & ", startingPointOfSearch="
  If startingPointOfSearch Is Nothing Then
    str = str & "Nothing, resumingSearchWithFirstButton=" & resumingSearchWithFirstButton
  Else
    str = str & startingPointOfSearch.ID & ", resumingSearchWithFirstButton=" & resumingSearchWithFirstButton
  End If
  str = str & ", matchingButton="
  If matchingButton Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & matchingButton.ID
  End If
  str = str & ", handledEvent=" & handledEvent

  AddLogEntry str
End Sub

Private Sub ToolBarU_MClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_MClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_MClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_MDblClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_MDblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_MDblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_MouseDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_MouseDown: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_MouseDown: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_MouseEnter(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_MouseEnter: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_MouseEnter: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_MouseHover(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_MouseHover: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_MouseHover: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_MouseLeave(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_MouseLeave: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_MouseLeave: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_MouseMove(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
'  If toolButton Is Nothing Then
'    AddLogEntry "ToolBarU_MouseMove: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  Else
'    AddLogEntry "ToolBarU_MouseMove: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
'  End If
End Sub

Private Sub ToolBarU_MouseUp(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_MouseUp: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_MouseUp: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_OLECompleteDrag(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject, ByVal performedEffect As TBarCtlsLibUCtl.OLEDropEffectConstants)
  Dim files() As String
  Dim str As String

  str = "ToolBarU_OLECompleteDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", performedEffect=" & performedEffect

  AddLogEntry str

  If Data.GetFormat(vbCFFiles) Then
    files = Data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    MsgBox "Dragged files:" & vbNewLine & str
  End If
End Sub

Private Sub ToolBarU_OLEDragDrop(ByVal data As TBarCtlsLibUCtl.IOLEDataObject, effect As TBarCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ToolBarU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    ToolBarU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub ToolBarU_OLEDragEnter(ByVal data As TBarCtlsLibUCtl.IOLEDataObject, effect As TBarCtlsLibUCtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants, autoClickButton As Boolean)
  Dim files() As String
  Dim str As String

  str = "ToolBarU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoClickButton=" & autoClickButton

  AddLogEntry str
End Sub

Private Sub ToolBarU_OLEDragEnterPotentialTarget(ByVal hWndPotentialTarget As Long)
  AddLogEntry "ToolBarU_OLEDragEnterPotentialTarget: hWndPotentialTarget=0x" & Hex(hWndPotentialTarget)
End Sub

Private Sub ToolBarU_OLEDragLeave(ByVal data As TBarCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "ToolBarU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails

  AddLogEntry str
End Sub

Private Sub ToolBarU_OLEDragLeavePotentialTarget()
  AddLogEntry "ToolBarU_OLEDragLeavePotentialTarget"
End Sub

Private Sub ToolBarU_OLEDragMouseMove(ByVal data As TBarCtlsLibUCtl.IOLEDataObject, effect As TBarCtlsLibUCtl.OLEDropEffectConstants, dropTarget As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants, autoClickButton As Boolean)
  Dim files() As String
  Dim str As String

  str = "ToolBarU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=" & effect & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.ID
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=" & hitTestDetails & ", autoClickButton=" & autoClickButton

  AddLogEntry str
End Sub

Private Sub ToolBarU_OLEGiveFeedback(ByVal effect As TBarCtlsLibUCtl.OLEDropEffectConstants, useDefaultCursors As Boolean)
  AddLogEntry "ToolBarU_OLEGiveFeedback: effect=" & effect & ", useDefaultCursors=" & useDefaultCursors
End Sub

Private Sub ToolBarU_OLEQueryContinueDrag(ByVal pressedEscape As Boolean, ByVal button As Integer, ByVal shift As Integer, actionToContinueWith As TBarCtlsLibUCtl.OLEActionToContinueWithConstants)
  AddLogEntry "ToolBarU_OLEQueryContinueDrag: pressedEscape=" & pressedEscape & ", button=" & button & ", shift=" & shift & ", actionToContinueWith=0x" & Hex(actionToContinueWith)
End Sub

Private Sub ToolBarU_OLEReceivedNewData(ByVal data As TBarCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ToolBarU_OLEReceivedNewData: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ToolBarU_OLESetData(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject, ByVal formatID As Long, ByVal Index As Long, ByVal dataOrViewAspect As Long)
  Dim files() As String
  Dim str As String

  str = "ToolBarU_OLESetData: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", formatID=" & formatID & ", index=" & Index & ", dataOrViewAspect=" & dataOrViewAspect

  AddLogEntry str
End Sub

Private Sub ToolBarU_OLEStartDrag(ByVal Data As TBarCtlsLibUCtl.IOLEDataObject)
  Dim files() As String
  Dim str As String

  str = "ToolBarU_OLEStartDrag: data="
  If Data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = Data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If

  AddLogEntry str
End Sub

Private Sub ToolBarU_OpenPopupMenu(ByVal hMenu As Long, ByVal parentMenuItemIndex As Long, ByVal isSystemMenu As Boolean)
  AddLogEntry "ToolBarU_OpenPopupMenu: hMenu=0x" & Hex(hMenu) & ", parentMenuItemIndex=" & parentMenuItemIndex & ", isSystemMenu=" & isSystemMenu
End Sub

Private Sub ToolBarU_QueryInsertButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, allowInsertionToLeft As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_QueryInsertButton: toolButton=Nothing, allowInsertionToLeft=" & allowInsertionToLeft
  Else
    AddLogEntry "ToolBarU_QueryInsertButton: toolButton=" & toolButton.ID & ", allowInsertionToLeft=" & allowInsertionToLeft
  End If
End Sub

Private Sub ToolBarU_QueryRemoveButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, allowRemoval As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_QueryRemoveButton: toolButton=Nothing, allowRemoval=" & allowRemoval
  Else
    AddLogEntry "ToolBarU_QueryRemoveButton: toolButton=" & toolButton.ID & ", allowRemoval=" & allowRemoval
  End If
End Sub

Private Sub ToolBarU_RawMenuMessage(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  AddLogEntry "ToolBarU_RawMenuMessage: message=0x" & Hex(message) & ", wParam=0x" & Hex(wParam) & ", lParam=0x" & Hex(lParam) & ", result=" & result & ", handledEvent=" & handledEvent
End Sub

Private Sub ToolBarU_RClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_RClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_RClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_RDblClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_RDblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_RDblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "ToolBarU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
  InsertButtonsU
End Sub

Private Sub ToolBarU_RemovedButton(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_RemovedButton: toolButton=Nothing"
  Else
    AddLogEntry "ToolBarU_RemovedButton: toolButton=" & toolButton.ID
  End If
End Sub

Private Sub ToolBarU_RemovingButton(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, cancelDeletion As Boolean)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_RemovingButton: toolButton=Nothing, cancelDeletion=" & cancelDeletion
  Else
    AddLogEntry "ToolBarU_RemovingButton: toolButton=" & toolButton.ID & ", cancelDeletion=" & cancelDeletion
  End If
End Sub

Private Sub ToolBarU_ResetCustomizations(ByVal hCustomizationDialog As Long, EndCustomization As Boolean)
  AddLogEntry "ToolBarU_ResetCustomizations: hCustomizationDialog=0x" & Hex(hCustomizationDialog) & ", EndCustomization=" & EndCustomization
  InsertButtonsU
End Sub

Private Sub ToolBarU_ResizedControlWindow()
  AddLogEntry "ToolBarU_ResizedControlWindow"
End Sub

Private Sub ToolBarU_RestoreButtonFromRegistryStream(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton, ByVal numberOfButtonsToLoad As Long, ByVal totalDataSize As Long, ByVal perButtonDataSize As Long, ByVal pDataStream As Long, pStartOfNextDataBlock As Long)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_RestoreButtonFromRegistryStream: toolButton=Nothing, numberOfButtonsToLoad=" & numberOfButtonsToLoad & ", totalDataSize=" & totalDataSize & ", perButtonDataSize=" & perButtonDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfNextDataBlock=0x" & Hex(pStartOfNextDataBlock)
  Else
    AddLogEntry "ToolBarU_RestoreButtonFromRegistryStream: toolButton=" & toolButton.ID & ", numberOfButtonsToLoad=" & numberOfButtonsToLoad & ", totalDataSize=" & totalDataSize & ", perButtonDataSize=" & perButtonDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfNextDataBlock=0x" & Hex(pStartOfNextDataBlock)
  End If
End Sub

Private Sub ToolBarU_SaveButtonToRegistryStream(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton, ByVal totalDataSize As Long, ByVal pDataStream As Long, pStartOfUnusedSpace As Long)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_SaveButtonToRegistryStream: toolButton=Nothing, totalDataSize=" & totalDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfUnusedSpace=0x" & Hex(pStartOfUnusedSpace)
  Else
    AddLogEntry "ToolBarU_SaveButtonToRegistryStream: toolButton=" & toolButton.ID & ", totalDataSize=" & totalDataSize & ", pDataStream=0x" & Hex(pDataStream) & ", pStartOfUnusedSpace=0x" & Hex(pStartOfUnusedSpace)
  End If
End Sub

Private Sub ToolBarU_SelectedMenuItem(ByVal commandIDOrSubMenuIndex As Long, ByVal menuItemState As TBarCtlsLibUCtl.MenuItemStateConstants, ByVal hMenu As Long)
  AddLogEntry "ToolBarU_SelectedMenuItem: commandIDOrSubMenuIndex=" & commandIDOrSubMenuIndex & ", menuItemState=0x" & Hex(menuItemState) & ", hMenu=0x" & Hex(hMenu)
End Sub

Private Sub ToolBarU_Validate(Cancel As Boolean)
  AddLogEntry "ToolBarU_Validate"
End Sub

Private Sub ToolBarU_XClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_XClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_XClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub

Private Sub ToolBarU_XDblClick(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  If toolButton Is Nothing Then
    AddLogEntry "ToolBarU_XDblClick: toolButton=Nothing, button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  Else
    AddLogEntry "ToolBarU_XDblClick: toolButton=" & toolButton.ID & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)
  End If
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    On Error Resume Next
    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub

Private Sub CreateMenus()
  Const MFS_CHECKED = &H8
  Const MFS_UNCHECKED = &H0
  Const MFT_STRING = &H0
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_SUBMENU = &H4
  Const MIIM_TYPE = &H10
  Dim hMenu As Long
  Dim hSubMenu As Long
  Dim miiData As MENUITEMINFO
  Dim strMenuItem As String

  hMenu = CreatePopupMenu
  menuButtons(0).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
      .fType = MFT_STRING

      .wID = MENUID_FILE * 100
      strMenuItem = "&New" & vbTab & "Ctrl+N"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 1, 1, miiData
      .wID = .wID + 1
      strMenuItem = "&Open" & vbTab & "Ctrl+O"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 2, 1, miiData
      .wID = .wID + 1
      strMenuItem = "&Save" & vbTab & "Ctrl+S"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 3, 1, miiData
      hSubMenu = CreatePopupMenu
      If hSubMenu Then
        menusToDestroy.Add hSubMenu
        .wID = .wID + 1
        strMenuItem = "Test.txt"
        .dwTypeData = StrPtr(strMenuItem)
        .cch = Len(strMenuItem)
        InsertMenuItem hSubMenu, 1, 1, miiData
      End If
      .wID = .wID + 1
      .hSubMenu = hSubMenu
      strMenuItem = "&Recent Files"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 3, 1, miiData
      .wID = .wID + 1
      .hSubMenu = 0
      MENUID_FILEEXIT = .wID
      strMenuItem = "&Exit" & vbTab & "Alt+F4"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 4, 1, miiData
    End With
  End If

  hMenu = CreatePopupMenu
  menuButtons(1).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
      .fType = MFT_STRING

      .wID = MENUID_VIEW * 100
      MENUID_VIEWTOOLBAR = .wID
      .fMask = .fMask Or MIIM_STATE
      .fState = .fState Or MFS_CHECKED
      strMenuItem = "&Tool bar"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 1, 1, miiData
      .fState = .fState And Not MFS_CHECKED
      .fMask = .fMask And Not MIIM_STATE
      .wID = .wID + 1
      MENUID_VIEWICONTEXT = .wID
      strMenuItem = "Icon &Text"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 2, 1, miiData
      .wID = .wID + 1
      MENUID_VIEWTOGGLEDOCK = .wID
      strMenuItem = "&Undock Tool Bar"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 3, 1, miiData
      .wID = .wID + 1
      MENUID_VIEWCUSTOMIZE = .wID
      strMenuItem = "&Customize Tool Bar..."
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 4, 1, miiData
    End With
  End If

  hMenu = CreatePopupMenu
  menuButtons(2).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU
      .fType = MFT_STRING

      .wID = MENUID_HELP * 100
      MENUID_HELPABOUT = .wID
      strMenuItem = "&About"
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 1, 1, miiData
    End With
  End If
End Sub

' returns the higher 16 bits of <value>
Private Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

' makes a 32 bits number out of two 16 bits parts
Private Function MakeDWord(ByVal Lo As Integer, ByVal hi As Integer) As Long
  Dim ret As Long

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(hi), LenB(hi)

  MakeDWord = ret
End Function

Private Sub InsertBandsA()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ReBarA.hWnd, StrPtr("explorer"), 0
  End If

  With ReBarA.Bands
    .Add(, ToolBarA.hWnd, , SizingGripVisibilityConstants.sgvAlways, IdealWidth:=ToolBarA.IdealWidth, MinimumHeight:=22, MaximumHeight:=22, BandData:=1).UseChevron = True
    .Add , cmbAddressA.hWnd, True, SizingGripVisibilityConstants.sgvAlways, MinimumWidth:=100, MinimumHeight:=22, MaximumHeight:=32, HeightChangeStepSize:=1, VariableHeight:=True, Text:="Add&ress:", BandData:=2
  End With
End Sub

Private Sub InsertButtonsA()
  Dim cImages As Long
  Dim i As Long

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ToolBarA.hWnd, StrPtr("explorer"), 0
  End If

  ToolBarA.hImageList(ilNormalButtons) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  For i = LBound(buttonsA) To UBound(buttonsA)
    buttonsA(i).ID = 1 + i
    buttonsA(i).IconIndex = i Mod cImages
    buttonsA(i).ButtonData = i + 1
    buttonsA(i).Text = "Command " & (i + 1)
  Next i

  With ToolBarA.Buttons
    .RemoveAll
    For i = LBound(buttonsA) + 1 To UBound(buttonsA) Step 3
      .Add buttonsA(i).ID, IconIndex:=buttonsA(i).IconIndex, Text:=buttonsA(i).Text, ButtonData:=buttonsA(i).ButtonData
    Next i
  End With
End Sub

Private Sub InsertBandsU()
  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ReBarU.hWnd, StrPtr("explorer"), 0
    'SetWindowTheme ReBarU.hWnd, StrPtr("Communications"), 0
  End If

  With ReBarU.Bands
    .Add(, tbMenu.hWnd, , SizingGripVisibilityConstants.sgvNever, KeepInFirstRow:=True, IdealWidth:=tbMenu.IdealWidth, MinimumHeight:=22, MaximumHeight:=22, BandData:=1).UseChevron = True
    With .Add(, ToolBarU.hWnd, True, SizingGripVisibilityConstants.sgvAlways, IdealWidth:=ToolBarU.IdealWidth, MinimumHeight:=22, MaximumHeight:=22, BandData:=1)
      .UseChevron = True
      BANDID_TOOLBAR = .ID
    End With
    .Add , cmbAddressU.hWnd, True, SizingGripVisibilityConstants.sgvAlways, MinimumWidth:=100, MinimumHeight:=22, MaximumHeight:=32, HeightChangeStepSize:=1, VariableHeight:=True, Text:="A&ddress:", BandData:=2
  End With
End Sub

Private Sub InsertButtonsU()
  Dim cImages As Long
  Dim i As Long
  Dim iIcon As Long

  If themeableOS Then
    ' for Windows Vista
    SetWindowTheme ToolBarU.hWnd, StrPtr("explorer"), 0
  End If

  ToolBarU.hImageList(ilNormalButtons) = hImgLst
  cImages = ImageList_GetImageCount(hImgLst)

  For i = LBound(menuButtons) To UBound(menuButtons)
    menuButtons(i).ID = Choose(i + 1, MENUID_FILE, MENUID_VIEW, MENUID_HELP)
    menuButtons(i).Text = Choose(i + 1, "&File", "&View", "&?")
  Next i

  With tbMenu
    With .Buttons
      .RemoveAll
      For i = LBound(menuButtons) To UBound(menuButtons)
        .Add menuButtons(i).ID, Text:=menuButtons(i).Text, DropDownStyle:=DropDownStyleConstants.ddstNormal
      Next i
    End With
    .RegisterHotkey mkCtrl, vbKeyN, MENUID_FILE * 100 + 0
    .RegisterHotkey mkCtrl, vbKeyO, MENUID_FILE * 100 + 1
    .RegisterHotkey mkCtrl, vbKeyS, MENUID_FILE * 100 + 2
    .RegisterHotkey mkAlt, vbKeyF4, MENUID_FILE * 100 + 5
    .RegisterHotkey 0, vbKeyF1, MENUID_HELP * 100 + 0
  End With

  For i = LBound(buttonsU) To UBound(buttonsU)
    buttonsU(i).ID = 1 + i
    buttonsU(i).IconIndex = i Mod cImages
    buttonsU(i).ButtonData = i + 1
    buttonsU(i).Text = "Command " & (i + 1)
  Next i

  With ToolBarU.Buttons
    .RemoveAll
    For i = LBound(buttonsU) + 1 To UBound(buttonsU) Step 3
      .Add buttonsU(i).ID, IconIndex:=buttonsU(i).IconIndex, Text:=buttonsU(i).Text, ButtonData:=buttonsU(i).ButtonData
    Next i
  End With
End Sub

' subclasses this Form
Private Sub Subclass()
  Const NF_REQUERY = 4
  Const WM_NOTIFYFORMAT = &H55

  #If UseSubClassing Then
    If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
      Debug.Print "Subclassing failed!"
    End If
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong ReBarU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ToolBarU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ReBarA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong ToolBarA.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
