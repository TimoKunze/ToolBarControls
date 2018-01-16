VERSION 5.00
Object = "{5D0D0ABC-4898-4e46-AB48-291074A737A1}#1.3#0"; "TBarCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ToolBarControls 1.3 - ToolBar Menu Sample"
   ClientHeight    =   2625
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   5565
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   175
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   371
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
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
      Left            =   3000
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin TBarCtlsLibUCtl.ToolBar ToolBarU 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      _cx             =   4683
      _cy             =   4048
      AllowCustomization=   0   'False
      AlwaysDisplayButtonText=   -1  'True
      AnchorHighlighting=   0   'False
      Appearance      =   0
      BackStyle       =   1
      BorderStyle     =   0
      ButtonHeight    =   0
      ButtonStyle     =   1
      ButtonTextPosition=   1
      ButtonWidth     =   0
      DisabledEvents  =   917739
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
      Orientation     =   1
      ProcessContextMenuKeys=   -1  'True
      RaiseCustomDrawEventOnEraseBackground=   0   'False
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
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements ISubclassedWindow


  Private Type DLLVERSIONINFO
    cbSize As Long
    dwMajor As Long
    dwMinor As Long
    dwBuildNumber As Long
    dwPlatformId As Long
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


  Private bComctl32Version600OrNewer As Boolean
  Private bLog As Boolean
  Private buttonsU(0 To 3) As TOOLBARBUTTONINFO
  Private hImgLst As Long
  Private themeableOS As Boolean


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal sz As Long)
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pSubAppName As Long, ByVal pSubIDList As Long) As Long


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


Private Sub cmdAbout_Click()
  ToolBarU.About
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
  Subclass

  Me.Show

  InsertButtonsU
End Sub

Private Sub Form_Terminate()
  If hImgLst Then ImageList_Destroy hImgLst
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub ToolBarU_CustomDraw(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal buttonState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, textRectangle As TBarCtlsLibUCtl.RECTANGLE, HorizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
  Select Case drawStage
    Case cdsPrePaint
      furtherProcessing = cdrvNotifyItemDraw
    Case cdsItemPrePaint
      ' according to http://jira.reactos.org/browse/CORE-6813 this should make the toolbar look like a menu
      If buttonState And cdisHot Then
        normalTextColor = vbHighlightText
        hotButtonBackColor = vbHighlight
      Else
        normalTextColor = vbMenuText
      End If
      furtherProcessing = cdrvUseCustomDrawColorsIfThemed Or cdrvUseHotButtonBackColor Or cdrvNoEtchedEffectIfDisabled Or cdrvDontOffsetIfPushed Or cdrvDontDrawBackground
  End Select
End Sub

Private Sub ToolBarU_RecreatedControlWindow(ByVal hWnd As Long)
  InsertButtonsU
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

  For i = LBound(buttonsU) To UBound(buttonsU)
    buttonsU(i).ID = 1 + i
    buttonsU(i).IconIndex = i Mod cImages
    buttonsU(i).ButtonData = 0
    buttonsU(i).Text = Choose(1 + i, "Programs", "Documents", "Settings", "Search")
  Next i

  With ToolBarU.Buttons
    .RemoveAll
    For i = LBound(buttonsU) To UBound(buttonsU)
      .Add buttonsU(i).ID, DropDownStyle:=ddstNormal, AutoSize:=False, IconIndex:=buttonsU(i).IconIndex, Text:=buttonsU(i).Text, ButtonData:=buttonsU(i).ButtonData
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
    SendMessageAsLong ToolBarU.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub
