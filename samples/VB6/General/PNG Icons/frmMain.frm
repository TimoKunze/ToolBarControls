VERSION 5.00
Object = "{5D0D0ABC-4898-4e46-AB48-291074A737A1}#1.3#0"; "TBarCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "ToolBarControls 1.3 - PNG Icons Sample"
   ClientHeight    =   3495
   ClientLeft      =   -15
   ClientTop       =   375
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   233
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin TBarCtlsLibUCtl.ToolBar tbMenu 
      Height          =   390
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
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
      DisabledEvents  =   917995
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
      DisabledEvents  =   524779
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
      Left            =   3593
      TabIndex        =   3
      Top             =   3000
      Width           =   2415
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
      DisabledEvents  =   917995
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

  Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
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
  Private MENUID_VIEWCUSTOMIZE As Long
  Private MENUID_HELPABOUT As Long

  Private bLog As Boolean
  Private buttonsU(0 To 13) As TOOLBARBUTTONINFO
  Private gdiPlusToken As Long
  Private hImgLst As Long
  Private menuButtons(0 To 2) As TOOLBARBUTTONINFO
  Private menusToDestroy As Collection
  Private objActiveCtl As Object
  Private themeableOS As Boolean


  Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal sz As Long)
  Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus.dll" (ByVal FileName As Long, ByRef bmp As Long) As Long
  Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus.dll" (ByVal bmp As Long, ByRef hbmReturn As Long, ByVal background As Long) As Long
  Private Declare Function GdipDisposeImage Lib "gdiplus.dll" (ByVal image As Long) As Long
  Private Declare Sub GdiplusShutdown Lib "gdiplus.dll" (ByVal token As Long)
  Private Declare Function GdiplusStartup Lib "gdiplus.dll" (ByRef token As Long, tInput As GdiplusStartupInput, ByVal pOutput As Long) As Long
  Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal ProcName As String) As Long
  Private Declare Function ImageList_Add Lib "comctl32.dll" (ByVal himl As Long, ByVal hbmImage As Long, ByVal hbmMask As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
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


Private Sub cmdAbout_Click()
  ToolBarU.About
End Sub

Private Sub Form_Initialize()
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Dim hMod As Long
  Dim iconsDir As String
  Dim iconPath As String
  Dim iconSize As Long
  Dim gdiplus As GdiplusStartupInput
  Dim bmp As Long
  Dim hBmp As Long

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If

  gdiplus.GdiplusVersion = 1
  If GdiplusStartup(gdiPlusToken, gdiplus, 0) <> 0 Then
    MsgBox "Could not initialize GDI+"
    End
  End If

  iconSize = 16
  hImgLst = ImageList_Create(iconSize, iconSize, ILC_COLOR32 Or ILC_MASK, 14, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconPath = iconsDir & "Icons.png"

  If GdipCreateBitmapFromFile(StrPtr(iconPath), bmp) = 0 Then
    If GdipCreateHBITMAPFromBitmap(bmp, hBmp, 0) = 0 Then
      ' image loaded
    End If
    GdipDisposeImage bmp
    bmp = 0
  End If

  If hBmp = 0 Then
    MsgBox "Could not toolbar icons"
    ImageList_Destroy hImgLst
    GdiplusShutdown gdiPlusToken
    End
  End If

  ImageList_Add hImgLst, hBmp, 0
  DeleteObject hBmp
  hBmp = 0
End Sub

Private Sub Form_Load()
  Subclass

  Set menusToDestroy = New Collection
  CreateMenus

  Me.Show

  InsertButtonsU
  InsertBandsU
End Sub

Private Sub Form_Terminate()
  If hImgLst Then
    ImageList_Destroy hImgLst
    hImgLst = 0
  End If
  If gdiPlusToken Then
    GdiplusShutdown gdiPlusToken
    gdiPlusToken = 0
  End If
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

Private Sub tbMenu_ButtonGetDropDownMenu(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, hMenu As Long)
  Dim i As Long

  For i = LBound(menuButtons) To UBound(menuButtons)
    If menuButtons(i).ID = toolButton.ID Then
      hMenu = menuButtons(i).hSubMenu
      Exit For
    End If
  Next i
End Sub

Private Sub tbMenu_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  If commandOrigin <> coButton Then
    Select Case commandID
      Case MENUID_FILEEXIT
        Unload Me

      Case MENUID_VIEWCUSTOMIZE
        ToolBarU.Customize

      Case MENUID_HELPABOUT
        cmdAbout_Click
    End Select
  End If
End Sub

Private Sub tbMenu_RecreatedControlWindow(ByVal hWnd As Long)
  InsertButtonsU
End Sub

Private Sub tbMenu_ResetCustomizations(ByVal hCustomizationDialog As Long, EndCustomization As Boolean)
  InsertButtonsU
End Sub

Private Sub ToolBarU_ContextMenu(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As TBarCtlsLibUCtl.HitTestConstants)
  Const TPM_RETURNCMD = &H100&
  Const TPM_TOPALIGN = &H0&
  Const TPM_VERPOSANIMATION = &H1000&
  Const TPM_VERTICAL = &H40&
  Dim commandID As Long
  Dim ptMenu As POINTAPI

  ptMenu.x = Me.ScaleX(x, vbTwips, vbPixels)
  ptMenu.y = Me.ScaleY(y, vbTwips, vbPixels)
  ClientToScreen ToolBarU.hWnd, ptMenu
  commandID = TrackPopupMenuEx(menuButtons(1).hSubMenu, TPM_TOPALIGN Or TPM_RETURNCMD Or TPM_VERTICAL Or TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, ToolBarU.hWnd, 0)
  If commandID <> 0 Then
    tbMenu_ExecuteCommand commandID, Nothing, coMenu, False
  End If
End Sub

Private Sub ToolBarU_GetAvailableButton(ByVal toolButton As TBarCtlsLibUCtl.IVirtualToolBarButton, noMoreButtons As Boolean)
  If toolButton.Index <= UBound(buttonsU) Then
    toolButton.ButtonData = buttonsU(toolButton.Index).ButtonData
    toolButton.IconIndex = buttonsU(toolButton.Index).IconIndex
    toolButton.ID = buttonsU(toolButton.Index).ID
    toolButton.Text = buttonsU(toolButton.Index).Text
  End If
  noMoreButtons = (toolButton.Index > UBound(buttonsU))
End Sub

Private Sub ToolBarU_RecreatedControlWindow(ByVal hWnd As Long)
  InsertButtonsU
End Sub

Private Sub ToolBarU_ResetCustomizations(ByVal hCustomizationDialog As Long, EndCustomization As Boolean)
  InsertButtonsU
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
      MENUID_VIEWCUSTOMIZE = .wID
      strMenuItem = "&Customize Tool Bar..."
      .dwTypeData = StrPtr(strMenuItem)
      .cch = Len(strMenuItem)
      InsertMenuItem hMenu, 1, 1, miiData
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
    For i = LBound(buttonsU) + 1 To UBound(buttonsU) 'Step 3
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
  #End If
End Sub
