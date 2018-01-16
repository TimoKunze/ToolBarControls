VERSION 5.00
Object = "{5D0D0ABC-4898-4E46-AB48-291074A737A1}#1.3#0"; "TBarCtlsU.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "ToolBarControls 1.3 - Custom Draw Sample"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9930
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picContainer 
      Align           =   1  'Oben ausrichten
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      ScaleHeight     =   61
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   658
      TabIndex        =   1
      Top             =   690
      Width           =   9930
      Begin VB.ComboBox cmbFontSize 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7200
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin TBarCtlsLibUCtl.ToolBar tbFormat 
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   120
         Width           =   1815
         _cx             =   3201
         _cy             =   661
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   0   'False
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   0
         ButtonStyle     =   1
         ButtonTextPosition=   0
         ButtonWidth     =   0
         DisabledEvents  =   917739
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
         RegisterForOLEDragDrop=   0
         RightToLeft     =   0
         ShadowColor     =   -1
         ShowToolTips    =   -1  'True
         SupportOLEDragImages=   -1  'True
         UseMnemonics    =   -1  'True
         UseSystemFont   =   -1  'True
         VerticalButtonPadding=   -1
         VerticalButtonSpacing=   0
         VerticalTextAlignment=   0
         WrapButtons     =   0   'False
      End
      Begin TBarCtlsLibUCtl.ToolBar tbCommon 
         Height          =   375
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1935
         _cx             =   3413
         _cy             =   661
         AllowCustomization=   0   'False
         AlwaysDisplayButtonText=   0   'False
         AnchorHighlighting=   0   'False
         Appearance      =   0
         BackStyle       =   0
         BorderStyle     =   0
         ButtonHeight    =   0
         ButtonStyle     =   1
         ButtonTextPosition=   0
         ButtonWidth     =   0
         DisabledEvents  =   393451
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
         MenuMode        =   0   'False
         MinimumButtonWidth=   0
         MousePointer    =   0
         MultiColumn     =   0   'False
         NormalDropDownButtonStyle=   0
         OLEDragImageStyle=   0
         Orientation     =   0
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
         VerticalTextAlignment=   0
         WrapButtons     =   0   'False
      End
      Begin TBarCtlsLibUCtl.ToolBar tbMenu 
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   120
         Width           =   1815
         _cx             =   3201
         _cy             =   661
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
         DisabledEvents  =   393451
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
   End
   Begin TBarCtlsLibUCtl.ReBar ReBarU 
      Align           =   1  'Oben ausrichten
      Height          =   690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9930
      _cx             =   17515
      _cy             =   1217
      AllowBandReordering=   -1  'True
      Appearance      =   0
      AutoUpdateLayout=   -1  'True
      BackColor       =   -1
      BorderStyle     =   0
      DisabledEvents  =   235
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
      ReplaceMDIFrameMenu=   2
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ShadowColor     =   -1
      SupportOLEDragImages=   -1  'True
      ToggleOnDoubleClick=   -1  'True
      UseSystemFont   =   -1  'True
      VerticalSizingGripsOnVerticalOrientation=   0   'False
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IMDIFrame
  Implements IMenuManager
  Implements ISubclassedWindow

  ' keep these in sync with the CMDID_VIEW_SKINOFFICE* constants
  ' CMDID_VIEW_SKINOFFICE* = CMDID_VIEW_FIRSTSKIN + osOffice*
  Private Enum OfficeStyleConstants
    osOffice2000 = 0
    osOfficeXP = 1
    osOffice2003 = 2
  End Enum


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

  Private Type POINT
    x As Long
    y As Long
  End Type

  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Private Type TOOLBARBUTTONINFO
    ID As Long
    IconIndex As Long
    Text As String
    hSubMenu As Long
  End Type

  Private Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
  End Type


  Private BANDID_MENUBAR As Long
  Private BANDID_TOOLBAR_COMMON As Long
  Private BANDID_TOOLBAR_FORMAT As Long

  Private activeStyle As OfficeStyleConstants
  Private bComctl32Version600OrNewer As Boolean
  Private buttonsCommon(0 To 6) As TOOLBARBUTTONINFO
  Private buttonsFormat(0 To 5) As TOOLBARBUTTONINFO
  Private WithEvents ctlHostWindow As TBarCtlsLibUCtl.ControlHostWindow
Attribute ctlHostWindow.VB_VarHelpID = -1
  Private hImgLst(0 To 1) As Long
  Private hMenuRecentFiles As Long
  Private menuButtons(0 To 4) As TOOLBARBUTTONINFO
  Private menusToDestroy As Collection
  Private nextDocumentIndex As Long
  Private openChildren As Collection
  Private visualManager As IVisualManager


  Private Declare Function CheckMenuRadioItem Lib "user32.dll" (ByVal hMenu As Long, ByVal idFirst As Long, ByVal idLast As Long, ByVal idCheck As Long, ByVal uFlags As Long) As Long
  Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, lpPoint As POINT) As Long
  Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
  Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal uPosition As Long, ByVal uFlags As Long) As Long
  Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
  Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function DllGetVersion_comctl32 Lib "comctl32.dll" Alias "DllGetVersion" (Data As DLLVERSIONINFO) As Long
  Private Declare Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal uIDEnableItem As Long, ByVal uEnable As Long) As Long
  Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
  Private Declare Function GetMenuItemID Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
  Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Function ImageList_AddIcon Lib "comctl32.dll" (ByVal himl As Long, ByVal hIcon As Long) As Long
  Private Declare Function ImageList_Create Lib "comctl32.dll" (ByVal cx As Long, ByVal cy As Long, ByVal flags As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
  Private Declare Function ImageList_Destroy Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Function ImageList_GetImageCount Lib "comctl32.dll" (ByVal himl As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageW" (ByVal hInst As Long, ByVal lpszName As Long, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
  Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxW" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
  Private Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As Long, ByVal fuFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, ByRef lptpm As TPMPARAMS) As Long


Private Sub IMDIFrame_CheckCommand(ByVal commandID As Long, ByVal check As Boolean)
  tbFormat.Buttons(commandID, btitID).SelectionState = IIf(check, ssChecked, ssUnchecked)
End Sub

Private Sub IMDIFrame_ClosingMDIChild(ByVal windowID As Long)
  Dim frm As Form

  On Error GoTo InvalidWindowID
  Set frm = openChildren.Item("k" & windowID)
  If Not frm Is Nothing Then
    openChildren.Remove "k" & windowID
    Set frm = Nothing
  End If
  If openChildren.Count = 0 Then
    ReBarU.Bands(BANDID_TOOLBAR_FORMAT, bitID).Visible = False
  End If

InvalidWindowID:
  tbMenu.Buttons(CMDID_WINDOW, btitID).Visible = (openChildren.Count > 0)
  ReBarU.Bands(BANDID_MENUBAR, bitID).IdealWidth = tbMenu.IdealWidth
End Sub

Private Sub IMDIFrame_EnableDisableCommand(ByVal commandID As Long, ByVal enable As Boolean)
  On Error Resume Next
  tbMenu.CommandEnabled(commandID) = enable
  tbCommon.CommandEnabled(commandID) = enable
  tbFormat.CommandEnabled(commandID) = enable
End Sub

Private Sub IMDIFrame_OpenedMDIChild(ByVal frm As Form, windowID As Long)
  If Not frm Is Nothing Then
    windowID = nextDocumentIndex
    nextDocumentIndex = nextDocumentIndex + 1

    openChildren.Add frm, "k" & windowID

    tbMenu.Buttons(CMDID_WINDOW, btitID).Visible = True
    ReBarU.Bands(BANDID_MENUBAR, bitID).IdealWidth = tbMenu.IdealWidth
    ReBarU.Bands(BANDID_TOOLBAR_FORMAT, bitID).Visible = True
  End If
End Sub

Private Sub IMDIFrame_UpdateFontSize(ByVal fontSize As Long)
  cmbFontSize.ListIndex = (fontSize - 8) / 2
End Sub

Private Function IMenuManager_IsMenuBand(ByVal band As TBarCtlsLibUCtl.IReBarBand) As Boolean
  If Not band Is Nothing Then
    IMenuManager_IsMenuBand = (band.ID = BANDID_MENUBAR)
  End If
End Function

Private Function IMenuManager_IsSkinnedMenu(ByVal hMenu As Long) As Boolean
  Dim i As Long
  Dim ret As Boolean

  If Not menusToDestroy Is Nothing Then
    For i = 1 To menusToDestroy.Count
      If menusToDestroy.Item(i) = hMenu Then
        ret = True
        Exit For
      End If
    Next i
  End If

  If Not ret Then
    ret = (hMenu = tbMenu.hChevronMenu)
  End If

  IMenuManager_IsSkinnedMenu = ret
End Function

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


Private Sub cmbFontSize_Click()
  Dim doc As IDocument

  Set doc = Me.ActiveForm
  If Not doc Is Nothing Then
    doc.SetFontSize cmbFontSize.ItemData(cmbFontSize.ListIndex)
  End If
End Sub

Private Sub MDIForm_Initialize()
  Const ILC_COLOR24 = &H18
  Const ILC_COLOR32 = &H20
  Const ILC_MASK = &H1
  Const IMAGE_ICON = 1
  Const LR_DEFAULTSIZE = &H40
  Const LR_LOADFROMFILE = &H10
  Dim DLLVerData As DLLVERSIONINFO
  Dim fileName As String
  Dim hIcon As Long
  Dim i As Long
  Dim iconsDir As String
  Dim iconSize As Long

  InitCommonControls

  With DLLVerData
    .cbSize = LenB(DLLVerData)
    DllGetVersion_comctl32 DLLVerData
    bComctl32Version600OrNewer = (.dwMajor >= 6)
  End With

  If Not bComctl32Version600OrNewer Then
    MsgBox "This sample requires version 6.0 or newer of comctl32.dll.", VbMsgBoxStyle.vbCritical, "Error"
    End
  End If

  iconSize = 16
  hImgLst(0) = ImageList_Create(iconSize, iconSize, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 7, 0)
  If Right$(App.Path, 3) = "bin" Then
    iconsDir = App.Path & "\..\res\"
  Else
    iconsDir = App.Path & "\res\"
  End If
  iconsDir = iconsDir & iconSize & "x" & iconSize & IIf(bComctl32Version600OrNewer, "x32bpp\", "x8bpp\")
  iconsDir = iconsDir & "normal\"
  For i = 1 To 7
    fileName = Choose(i, "New", "Open", "Save", "Print", "Cut", "Copy", "Paste") & ".ico"
    hIcon = LoadImage(0, StrPtr(iconsDir & fileName), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst(0), hIcon
      DestroyIcon hIcon
    End If
  Next i

  hImgLst(1) = ImageList_Create(iconSize, iconSize, IIf(bComctl32Version600OrNewer, ILC_COLOR32, ILC_COLOR24) Or ILC_MASK, 6, 0)
  For i = 1 To 6
    fileName = Choose(i, "Bold", "Italic", "Underline", "AlignLeft", "AlignCenter", "AlignRight") & ".ico"
    hIcon = LoadImage(0, StrPtr(iconsDir & fileName), IMAGE_ICON, iconSize, iconSize, LR_LOADFROMFILE Or LR_DEFAULTSIZE)
    If hIcon Then
      ImageList_AddIcon hImgLst(1), hIcon
      DestroyIcon hIcon
    End If
  Next i

  Set openChildren = New Collection
  nextDocumentIndex = 1
End Sub

Private Sub MDIForm_Load()
  Dim fontSize As Long

  Subclass

  Set menusToDestroy = New Collection
  CreateMenus
  For fontSize = 8 To 32 Step 2
    cmbFontSize.AddItem CStr(fontSize)
    cmbFontSize.ItemData(cmbFontSize.NewIndex) = fontSize
  Next fontSize
  cmbFontSize.ListIndex = 0

  SetStyle osOffice2003

  Me.Show

  InsertButtons
  InsertBands

  ExecuteCommand CMDID_FILE_NEW
End Sub

Private Sub MDIForm_Terminate()
  If hImgLst(0) Then ImageList_Destroy hImgLst(0)
  If hImgLst(1) Then ImageList_Destroy hImgLst(1)
  Set ctlHostWindow = Nothing
  If Not visualManager Is Nothing Then
    visualManager.SetMenuManager Nothing
  End If
  Set visualManager = Nothing
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Dim hMenu As Long
  Dim i As Long

  ' it is important to clean up contained windows on destruction
  tbFormat.Buttons.Item(CMDID_FORMAT_FONTSIZE, btitID).SetContainedWindow 0

  For i = 1 To menusToDestroy.Count
    hMenu = CLng(menusToDestroy(i))
    FreeAllMenuItemTags hMenu
    DestroyMenu hMenu
  Next i
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub

Private Sub ReBarU_CustomDraw(ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal bandState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
  If Not visualManager Is Nothing Then
    visualManager.ReBarCustomDraw ReBarU, band, drawStage, bandState, hDC, drawingRectangle, furtherProcessing
  End If
End Sub

Private Sub ReBarU_RawMenuMessage(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  If Not visualManager Is Nothing Then
    visualManager.ReBarRawMenuMessage ReBarU, message, wParam, lParam, result, handledEvent
  End If
End Sub

Private Sub tbCommon_BeforeDisplayChevronPopup(ByVal hPopup As Long, ByVal x As Single, ByVal y As Single, ByVal isMenu As Boolean, cancelPopup As Boolean, commandToExecute As Long)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarBeforeDisplayChevronPopup tbCommon, hPopup, x, y, isMenu, cancelPopup, commandToExecute
  End If
End Sub

Private Sub tbCommon_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  infoTipText = buttonsCommon(toolButton.ButtonData).Text
End Sub

Private Sub tbCommon_CustomDraw(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal buttonState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, textRectangle As TBarCtlsLibUCtl.RECTANGLE, HorizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarCustomDraw tbCommon, toolButton, normalTextColor, normalButtonBackColor, normalBackgroundMode, hotTextColor, hotButtonBackColor, markedTextBackColor, markedButtonBackColor, markedBackgroundMode, drawStage, buttonState, hDC, drawingRectangle, textRectangle, HorizontalIconCaptionGap, furtherProcessing
  End If
End Sub

Private Sub tbCommon_DestroyingChevronPopup(ByVal hPopup As Long, ByVal isMenu As Boolean)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarDestroyingChevronPopup tbCommon, hPopup, isMenu
  End If
End Sub

Private Sub tbCommon_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  Const TPM_LEFTALIGN = &H0&
  Const TPM_TOPALIGN = &H0&
  Const TPM_VERPOSANIMATION = &H1000&
  Const TPM_VERTICAL = &H40&
  Dim ptMenu As POINT
  Dim tpmexParams As TPMPARAMS

  If Not visualManager Is Nothing Then
    visualManager.ToolBarDropDown tbCommon, toolButton, buttonRectangle, furtherProcessing
  End If

  If Not toolButton Is Nothing Then
    If toolButton.ID = CMDID_FILE_OPEN Then
      ptMenu.x = buttonRectangle.Left
      ptMenu.y = buttonRectangle.Bottom
      ClientToScreen tbCommon.hWnd, ptMenu
      With tpmexParams
        .cbSize = LenB(tpmexParams)
        LSet .rcExclude = buttonRectangle
      End With
      TrackPopupMenuEx hMenuRecentFiles, TPM_LEFTALIGN Or TPM_TOPALIGN Or TPM_VERTICAL Or TPM_VERPOSANIMATION, ptMenu.x, ptMenu.y, tbCommon.hWnd, tpmexParams
    End If
  End If
End Sub

Private Sub tbCommon_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  ExecuteCommand commandID
  forwardMessage = False
End Sub

Private Sub tbCommon_RawMenuMessage(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarRawMenuMessage tbCommon, message, wParam, lParam, result, handledEvent
  End If
End Sub

Private Sub tbFormat_BeforeDisplayChevronPopup(ByVal hPopup As Long, ByVal x As Single, ByVal y As Single, ByVal isMenu As Boolean, cancelPopup As Boolean, commandToExecute As Long)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarBeforeDisplayChevronPopup tbFormat, hPopup, x, y, isMenu, cancelPopup, commandToExecute
  End If
End Sub

Private Sub tbFormat_ButtonGetInfoTipText(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If toolButton.ButtonType <> btyPlaceholder Then
    infoTipText = buttonsFormat(toolButton.ButtonData).Text
  End If
End Sub

Private Sub tbFormat_CustomDraw(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal buttonState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, textRectangle As TBarCtlsLibUCtl.RECTANGLE, HorizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarCustomDraw tbFormat, toolButton, normalTextColor, normalButtonBackColor, normalBackgroundMode, hotTextColor, hotButtonBackColor, markedTextBackColor, markedButtonBackColor, markedBackgroundMode, drawStage, buttonState, hDC, drawingRectangle, textRectangle, HorizontalIconCaptionGap, furtherProcessing
  End If
End Sub

Private Sub tbFormat_DestroyingChevronPopup(ByVal hPopup As Long, ByVal isMenu As Boolean)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarDestroyingChevronPopup tbFormat, hPopup, isMenu
  End If
End Sub

Private Sub tbFormat_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarDropDown tbFormat, toolButton, buttonRectangle, furtherProcessing
  End If
End Sub

Private Sub tbFormat_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  ExecuteCommand commandID
  forwardMessage = False
End Sub

Private Sub tbMenu_BeforeDisplayChevronPopup(ByVal hPopup As Long, ByVal x As Single, ByVal y As Single, ByVal isMenu As Boolean, cancelPopup As Boolean, commandToExecute As Long)
  If isMenu Then
    MakeMenuOwnerDrawn hPopup
  End If
  If Not visualManager Is Nothing Then
    visualManager.ToolBarBeforeDisplayChevronPopup tbMenu, hPopup, x, y, isMenu, cancelPopup, commandToExecute
  End If
End Sub

Private Sub tbMenu_ButtonGetDropDownMenu(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, hMenu As Long)
  Const MF_BYCOMMAND = &H0
  Dim i As Long

  For i = LBound(menuButtons) To UBound(menuButtons)
    If menuButtons(i).ID = toolButton.ID Then
      hMenu = menuButtons(i).hSubMenu
      If menuButtons(i).ID = CMDID_WINDOW Then
        UpdateWindowMenu hMenu
      ElseIf menuButtons(i).ID = CMDID_VIEW Then
        CheckMenuRadioItem hMenu, CMDID_VIEW_FIRSTSKIN, CMDID_VIEW_LASTSKIN, CMDID_VIEW_FIRSTSKIN + activeStyle, MF_BYCOMMAND
      End If
      Exit For
    End If
  Next i
End Sub

Private Sub tbMenu_CustomDraw(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As TBarCtlsLibUCtl.StringBackgroundModeConstants, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal buttonState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, textRectangle As TBarCtlsLibUCtl.RECTANGLE, HorizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarCustomDraw tbMenu, toolButton, normalTextColor, normalButtonBackColor, normalBackgroundMode, hotTextColor, hotButtonBackColor, markedTextBackColor, markedButtonBackColor, markedBackgroundMode, drawStage, buttonState, hDC, drawingRectangle, textRectangle, HorizontalIconCaptionGap, furtherProcessing
  End If
End Sub

Private Sub tbMenu_DestroyingChevronPopup(ByVal hPopup As Long, ByVal isMenu As Boolean)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarDestroyingChevronPopup tbMenu, hPopup, isMenu
  End If
  FreeAllMenuItemTags hPopup, False
End Sub

Private Sub tbMenu_DropDown(ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarDropDown tbMenu, toolButton, buttonRectangle, furtherProcessing
  End If
End Sub

Private Sub tbMenu_ExecuteCommand(ByVal commandID As Long, ByVal toolButton As TBarCtlsLibUCtl.IToolBarButton, ByVal commandOrigin As TBarCtlsLibUCtl.CommandOriginConstants, forwardMessage As Boolean)
  If commandOrigin <> coButton Then
    ExecuteCommand commandID
    forwardMessage = False
  End If
End Sub

Private Sub tbMenu_OpenPopupMenu(ByVal hMenu As Long, ByVal parentMenuItemIndex As Long, ByVal isSystemMenu As Boolean)
  Const MF_BYCOMMAND = &H0
  Const MF_ENABLED = &H0
  Const MF_GRAYED = &H1
  Dim i As Long
  Dim commandID As Long

  For i = 0 To GetMenuItemCount(hMenu) - 1
    commandID = GetMenuItemID(hMenu, i)
    If commandID >= 0 Then
      EnableMenuItem hMenu, commandID, MF_BYCOMMAND Or IIf(tbMenu.CommandEnabled(commandID), MF_ENABLED, MF_GRAYED)
    End If
  Next i
End Sub

Private Sub tbMenu_RawMenuMessage(ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  If Not visualManager Is Nothing Then
    visualManager.ToolBarRawMenuMessage tbMenu, message, wParam, lParam, result, handledEvent
  End If
End Sub


Private Function CreateMenuTagObject(Optional Text As String = "", Optional ByVal hImgLst As Long = 0, Optional ByVal IconIndex As Long = -2, Optional ByVal HasSubMenu As Boolean = False, Optional ByVal RadioCheckMark As Boolean = False, Optional ByVal Separator As Boolean = False) As CMenuItemData
  Dim tagObject As CMenuItemData

  Set tagObject = New CMenuItemData
  With tagObject
    .Text = Text
    .hImageList = hImgLst
    .IconIndex = IconIndex
    .HasSubMenu = HasSubMenu
    .RadioCheckMark = RadioCheckMark
    .Separator = Separator
  End With
  Set CreateMenuTagObject = tagObject
End Function

Private Sub CreateMenus()
  Const MFS_CHECKED = &H8
  Const MFS_UNCHECKED = &H0
  Const MFT_OWNERDRAW = &H100
  Const MFT_SEPARATOR = &H800
  Const MIIM_DATA = &H20
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_SUBMENU = &H4
  Const MIIM_TYPE = &H10
  Dim hMenu As Long
  Dim miiData As MENUITEMINFO
  Dim tagObject As CMenuItemData

  hMenu = CreatePopupMenu
  menuButtons(0).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU Or MIIM_DATA
      .fType = MFT_OWNERDRAW

      .wID = CMDID_FILE_NEW
      Set tagObject = CreateMenuTagObject("&New" & vbTab & "Ctrl+N", hImgLst(0), 0)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 1, 1, miiData
      .wID = CMDID_FILE_OPEN
      Set tagObject = CreateMenuTagObject("&Open..." & vbTab & "Ctrl+O", hImgLst(0), 1)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 2, 1, miiData
      .wID = CMDID_FILE_SAVE
      Set tagObject = CreateMenuTagObject("&Save" & vbTab & "Ctrl+S", hImgLst(0), 2)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 3, 1, miiData
      .wID = CMDID_FILE_PRINT
      Set tagObject = CreateMenuTagObject("&Print", hImgLst(0), 3)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 4, 1, miiData
      hMenuRecentFiles = CreatePopupMenu
      If hMenuRecentFiles Then
        menusToDestroy.Add hMenuRecentFiles
        .wID = CMDID_FILE_RECENT_FIRST
        Set tagObject = CreateMenuTagObject("Test.txt")
        .dwItemData = MakeMenuTag(tagObject)
        InsertMenuItem hMenuRecentFiles, 1, 1, miiData
      End If
      .wID = CMDID_FILE_RECENT
      .hSubMenu = hMenuRecentFiles
      Set tagObject = CreateMenuTagObject("&Recent Files", HasSubMenu:=True)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 5, 1, miiData
      .wID = CMDID_FILE_EXIT
      .hSubMenu = 0
      Set tagObject = CreateMenuTagObject("&Exit" & vbTab & "Alt+F4")
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 6, 1, miiData

      .fType = MFT_SEPARATOR Or MFT_OWNERDRAW
      Set tagObject = CreateMenuTagObject(Separator:=True)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 5, 1, miiData
      Set tagObject = CreateMenuTagObject(Separator:=True)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 4, 1, miiData
      Set tagObject = CreateMenuTagObject(Separator:=True)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 3, 1, miiData
    End With
  End If

  hMenu = CreatePopupMenu
  menuButtons(1).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU Or MIIM_DATA
      .fType = MFT_OWNERDRAW

      .wID = CMDID_EDIT_CUT
      Set tagObject = CreateMenuTagObject("Cu&t" & vbTab & "Ctrl+X", hImgLst(0), 4)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 1, 1, miiData
      .wID = CMDID_EDIT_COPY
      Set tagObject = CreateMenuTagObject("&Copy" & vbTab & "Ctrl+C", hImgLst(0), 5)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 2, 1, miiData
      .wID = CMDID_EDIT_PASTE
      Set tagObject = CreateMenuTagObject("&Paste" & vbTab & "Ctrl+V", hImgLst(0), 6)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 3, 1, miiData
    End With
  End If

  hMenu = CreatePopupMenu
  menuButtons(2).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU Or MIIM_DATA
      .fType = MFT_OWNERDRAW

      .wID = CMDID_VIEW_SKINOFFICE2000
      Set tagObject = CreateMenuTagObject("Office &2000", RadioCheckMark:=True)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 1, 1, miiData
      .wID = CMDID_VIEW_SKINOFFICEXP
      Set tagObject = CreateMenuTagObject("Office &XP", RadioCheckMark:=True)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 2, 1, miiData
      .wID = CMDID_VIEW_SKINOFFICE2003
      Set tagObject = CreateMenuTagObject("Office 200&3", RadioCheckMark:=True)
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 3, 1, miiData
    End With
  End If

  hMenu = CreatePopupMenu
  menuButtons(3).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU Or MIIM_DATA
      .fType = MFT_OWNERDRAW

      .wID = CMDID_WINDOW_CASCADE
      Set tagObject = CreateMenuTagObject("&Cascade")
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 1, 1, miiData
      .wID = CMDID_WINDOW_TILE
      Set tagObject = CreateMenuTagObject("&Tile")
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 2, 1, miiData
      .wID = CMDID_WINDOW_ARRANGEICONS
      Set tagObject = CreateMenuTagObject("Arrange &Icons")
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 3, 1, miiData
    End With
  End If

  hMenu = CreatePopupMenu
  menuButtons(4).hSubMenu = hMenu
  If hMenu Then
    menusToDestroy.Add hMenu
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU Or MIIM_DATA
      .fType = MFT_OWNERDRAW

      .wID = CMDID_HELP_ABOUT
      Set tagObject = CreateMenuTagObject("&About")
      .dwItemData = MakeMenuTag(tagObject)
      InsertMenuItem hMenu, 1, 1, miiData
    End With
  End If
End Sub

Private Sub ExecuteCommand(ByVal commandID As Long)
  Const MB_OK = &H0
  Dim doc As IDocument
  Dim frm As Form

  Select Case commandID
    Case CMDID_FILE_NEW
      Set frm = New frmDocument
      Set frm.mdiFrame = Me
      frm.Show
      Set frm = Nothing
    Case CMDID_FILE_EXIT
      Unload Me
    Case CMDID_FILE_OPEN
      ' The VB6 MsgBox command doesn't work well together with custom draw. It probably disables any events.
      'MsgBox "Command not implemented"
      MessageBox Me.hWnd, StrPtr("Command not implemented"), 0, MB_OK
    Case CMDID_WINDOW_CASCADE
      Me.Arrange vbCascade
    Case CMDID_WINDOW_TILE
      Me.Arrange vbTileHorizontal
    Case CMDID_WINDOW_ARRANGEICONS
      Me.Arrange vbArrangeIcons
    Case CMDID_HELP_ABOUT
      tbMenu.About
    Case Else
      If commandID >= CMDID_VIEW_FIRSTSKIN And commandID <= CMDID_VIEW_LASTSKIN Then
        SetStyle commandID - CMDID_VIEW_FIRSTSKIN
      ElseIf commandID >= CMDID_WINDOW_LISTSTART And commandID <= CMDID_WINDOW_LISTSTART + openChildren.Count Then
        Set frm = openChildren.Item(commandID - CMDID_WINDOW_LISTSTART)
        If Not frm Is Nothing Then
          frm.ZOrder 0
          If frm.WindowState = vbMinimized Then
            frm.WindowState = vbNormal
          End If
        End If
      ElseIf commandID >= CMDID_FILE_RECENT_FIRST And commandID <= CMDID_FILE_RECENT_LAST Then
        ' The VB6 MsgBox command doesn't work well together with custom draw. It probably disables any events.
        'MsgBox "Command not implemented"
        MessageBox Me.hWnd, StrPtr("Command not implemented"), 0, MB_OK
      Else
        Set doc = Me.ActiveForm
        If Not doc Is Nothing Then
          doc.ExecuteCommand commandID
        End If
      End If
  End Select
End Sub

Private Sub FreeAllMenuItemTags(hMenu As Long, Optional ByVal recursive As Boolean = True)
  Const MIIM_DATA = &H20
  Const MIIM_ID = &H2
  Const MIIM_SUBMENU = &H4
  Const MIIM_TYPE = &H10
  Dim i As Long
  Dim miiData As MENUITEMINFO
  Dim pData As Long

  With miiData
    .cbSize = LenB(miiData)
    .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_SUBMENU Or MIIM_DATA
    For i = 0 To GetMenuItemCount(hMenu) - 1
      If GetMenuItemInfo(hMenu, i, 1, miiData) Then
        If (miiData.hSubMenu <> 0) And recursive Then
          FreeAllMenuItemTags miiData.hSubMenu
        End If
        pData = miiData.dwItemData
        miiData.dwItemData = 0
        SetMenuItemInfo hMenu, i, 1, miiData
        FreeMenuTag pData
      End If
    Next i
  End With
End Sub

Private Sub InsertBands()
  Dim leftBorder As Long
  Dim rightBorder As Long

  With ReBarU.Bands
    With .Add(, tbMenu.hWnd, , SizingGripVisibilityConstants.sgvAlways, KeepInFirstRow:=True, IdealWidth:=tbMenu.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_MENUBAR = .ID
    End With
    With .Add(, tbCommon.hWnd, True, SizingGripVisibilityConstants.sgvAlways, IdealWidth:=tbCommon.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_TOOLBAR_COMMON = .ID
    End With
    With .Add(, tbFormat.hWnd, , SizingGripVisibilityConstants.sgvAlways, IdealWidth:=tbFormat.IdealWidth, MinimumHeight:=22, MaximumHeight:=22)
      .UseChevron = True
      BANDID_TOOLBAR_FORMAT = .ID
      .Maximize
    End With

    With .Item(BANDID_TOOLBAR_COMMON, bitID)
      .GetBorderSizes leftBorder, 0, rightBorder, 0
      .CurrentWidth = tbCommon.IdealWidth + leftBorder + rightBorder
    End With
  End With
  Set ReBarU.MDIFrameMenuBand = ReBarU.Bands(BANDID_MENUBAR, bitID)
  picContainer.Visible = False
End Sub

Private Sub InsertButtons()
  Dim i As Long

  For i = LBound(menuButtons) To UBound(menuButtons)
    menuButtons(i).ID = Choose(i + 1, CMDID_FILE, CMDID_EDIT, CMDID_VIEW, CMDID_WINDOW, CMDID_HELP)
    menuButtons(i).Text = Choose(i + 1, "&File", "&Edit", "&View", "&Window", "&?")
  Next i

  With tbMenu
    With .Buttons
      .RemoveAll
      For i = LBound(menuButtons) To UBound(menuButtons)
        .Add menuButtons(i).ID, Text:=menuButtons(i).Text, DropDownStyle:=DropDownStyleConstants.ddstNormal, Visible:=(menuButtons(i).ID <> CMDID_WINDOW)
      Next i
    End With
    .RegisterHotkey mkCtrl, vbKeyN, CMDID_FILE_NEW
    .RegisterHotkey mkCtrl, vbKeyO, CMDID_FILE_OPEN
    .RegisterHotkey mkCtrl, vbKeyS, CMDID_FILE_SAVE
    .RegisterHotkey mkAlt, vbKeyF4, CMDID_FILE_EXIT
    .RegisterHotkey mkCtrl, vbKeyX, CMDID_EDIT_CUT
    .RegisterHotkey mkCtrl, vbKeyC, CMDID_EDIT_COPY
    .RegisterHotkey mkCtrl, vbKeyV, CMDID_EDIT_PASTE
    .RegisterHotkey mkCtrl, vbKeyA, CMDID_EDIT_SELECTALL
    .RegisterHotkey 0, vbKeyF1, CMDID_HELP_ABOUT
  End With

  'tbCommon.LoadDefaultImages siltSmallStandardBitmaps
  tbCommon.hImageList(ilNormalButtons) = hImgLst(0)
  buttonsCommon(0).ID = CMDID_FILE_NEW
  buttonsCommon(0).Text = "&New"
  buttonsCommon(0).IconIndex = 0 'siiStandardFileNew
  buttonsCommon(1).ID = CMDID_FILE_OPEN
  buttonsCommon(1).Text = "&Open..."
  buttonsCommon(1).IconIndex = 1 'siiStandardFileOpen
  buttonsCommon(2).ID = CMDID_FILE_SAVE
  buttonsCommon(2).Text = "&Save"
  buttonsCommon(2).IconIndex = 2 'siiStandardFileSave
  buttonsCommon(3).ID = CMDID_FILE_PRINT
  buttonsCommon(3).Text = "&Print"
  buttonsCommon(3).IconIndex = 3 'siiStandardPrint
  buttonsCommon(4).ID = CMDID_EDIT_CUT
  buttonsCommon(4).Text = "Cu&t"
  buttonsCommon(4).IconIndex = 4 'siiStandardCut
  buttonsCommon(5).ID = CMDID_EDIT_COPY
  buttonsCommon(5).Text = "&Copy"
  buttonsCommon(5).IconIndex = 5 'siiStandardCopy
  buttonsCommon(6).ID = CMDID_EDIT_PASTE
  buttonsCommon(6).Text = "&Paste"
  buttonsCommon(6).IconIndex = 6 'siiStandardPaste

  With tbCommon.Buttons
    .RemoveAll
    For i = LBound(buttonsCommon) To UBound(buttonsCommon)
      .Add buttonsCommon(i).ID, IconIndex:=buttonsCommon(i).IconIndex, DropDownStyle:=IIf(buttonsCommon(i).ID = CMDID_FILE_OPEN, DropDownStyleConstants.ddstNormal, DropDownStyleConstants.ddstNoDropDown), ButtonData:=i
    Next i
    ' Now insert a separator between Print and Cut command. NOTE: It's better to disable auto-sizing for separators
    .Add 0, 4, ButtonType:=btySeparator, AutoSize:=False
  End With

  tbFormat.hImageList(ilNormalButtons) = hImgLst(1)
  buttonsFormat(0).ID = CMDID_FORMAT_BOLD
  buttonsFormat(0).Text = "&Bold"
  buttonsFormat(0).IconIndex = 0
  buttonsFormat(1).ID = CMDID_FORMAT_ITALIC
  buttonsFormat(1).Text = "&Italic"
  buttonsFormat(1).IconIndex = 1
  buttonsFormat(2).ID = CMDID_FORMAT_UNDERLINE
  buttonsFormat(2).Text = "&Underline"
  buttonsFormat(2).IconIndex = 2
  buttonsFormat(3).ID = CMDID_FORMAT_ALIGNLEFT
  buttonsFormat(3).Text = "Align &Left"
  buttonsFormat(3).IconIndex = 3
  buttonsFormat(4).ID = CMDID_FORMAT_ALIGNCENTER
  buttonsFormat(4).Text = "C&enter"
  buttonsFormat(4).IconIndex = 4
  buttonsFormat(5).ID = CMDID_FORMAT_ALIGNRIGHT
  buttonsFormat(5).Text = "Align &Right"
  buttonsFormat(5).IconIndex = 5

  With tbFormat.Buttons
    .RemoveAll
    For i = LBound(buttonsFormat) To UBound(buttonsFormat)
      .Add buttonsFormat(i).ID, ButtonType:=btyCheckButton, PartOfGroup:=(buttonsFormat(i).ID >= CMDID_FORMAT_ALIGNLEFT And buttonsFormat(i).ID <= CMDID_FORMAT_ALIGNRIGHT), IconIndex:=buttonsFormat(i).IconIndex, ButtonData:=i
    Next i
    ' Now insert a separator between Underline and Align Left command. NOTE: It's better to disable auto-sizing for separators
    .Add 0, 3, ButtonType:=btySeparator, AutoSize:=False
    ' insert a placeholder for the font size combo box
    With .Add(CMDID_FORMAT_FONTSIZE, 4, ButtonType:=btyPlaceholder, AutoSize:=False, Width:=cmbFontSize.Width + 4)
      .SetContainedWindow cmbFontSize.hWnd
    End With
    ' also insert a separator after the placeholder
    .Add 0, 5, ButtonType:=btySeparator, AutoSize:=False
    .Item(CMDID_FORMAT_ALIGNLEFT, btitID).SelectionState = ssChecked
  End With

  IMDIFrame_EnableDisableCommand CMDID_FILE_SAVE, False
  IMDIFrame_EnableDisableCommand CMDID_EDIT_CUT, False
  IMDIFrame_EnableDisableCommand CMDID_EDIT_COPY, False
  IMDIFrame_EnableDisableCommand CMDID_EDIT_PASTE, False
End Sub

Private Sub MakeMenuOwnerDrawn(ByVal hMenu As Long)
  Const MFT_OWNERDRAW = &H100
  Const MFT_SEPARATOR = &H800
  Const MIIM_DATA = &H20
  Const MIIM_SUBMENU = &H4
  Const MIIM_TYPE = &H10
  Dim i As Long
  Dim menuText As String
  Dim miiData As MENUITEMINFO
  Dim tagObject As CMenuItemData

  With miiData
    .cbSize = LenB(miiData)
    For i = 0 To GetMenuItemCount(hMenu) - 1
      .fMask = MIIM_TYPE Or MIIM_SUBMENU
      menuText = String$(200, Chr$(0))
      .dwTypeData = StrPtr(menuText)
      .cch = 200
      GetMenuItemInfo hMenu, i, 1, miiData
      menuText = Left$(menuText, .cch)

      Set tagObject = CreateMenuTagObject(menuText, HasSubMenu:=(.hSubMenu <> 0), Separator:=(.fType And MFT_SEPARATOR))
      .fMask = MIIM_TYPE Or MIIM_DATA
      .fType = MFT_OWNERDRAW
      .dwItemData = MakeMenuTag(tagObject)
      SetMenuItemInfo hMenu, i, 1, miiData
    Next i
  End With
End Sub

Private Sub SetStyle(ByVal newStyle As OfficeStyleConstants)
  If Not visualManager Is Nothing Then
    visualManager.SetMenuManager Nothing
  End If
  activeStyle = newStyle
  Select Case newStyle
    Case osOffice2000
      Set visualManager = New CVisualManagerOffice2000
    Case osOfficeXP
      Set visualManager = New CVisualManagerOfficeXP
    Case osOffice2003
      Set visualManager = New CVisualManagerOffice2003
  End Select
  If Not visualManager Is Nothing Then
    visualManager.SetMenuManager Me
    visualManager.InitializeReBar ReBarU
    visualManager.InitializeToolBar tbMenu
    visualManager.InitializeToolBar tbCommon
    visualManager.InitializeToolBar tbFormat
  End If
  ReBarU.Refresh
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
    SendMessageAsLong tbMenu.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong tbCommon.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
    SendMessageAsLong tbFormat.hWnd, WM_NOTIFYFORMAT, Me.hWnd, NF_REQUERY
  #End If
End Sub

Private Sub UpdateWindowMenu(ByVal hWindowMenu As Long)
  Const MF_BYPOSITION = &H400
  Const MFS_CHECKED = &H8
  Const MFS_UNCHECKED = &H0
  Const MFT_OWNERDRAW = &H100
  Const MFT_SEPARATOR = &H800
  Const MIIM_DATA = &H20
  Const MIIM_ID = &H2
  Const MIIM_STATE = &H1
  Const MIIM_TYPE = &H10
  Dim endOfMenu As Long
  Dim frm As Form
  Dim i As Long
  Dim miiData As MENUITEMINFO
  Dim tagObject As CMenuItemData

  If hWindowMenu Then
    ' remove any menu items with an ID >= CMDID_WINDOW_LISTSTART
    With miiData
      .cbSize = LenB(miiData)
      .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_DATA
      For i = GetMenuItemCount(hWindowMenu) - 1 To 0 Step -1
        If GetMenuItemInfo(hWindowMenu, i, 1, miiData) Then
          If miiData.wID >= CMDID_WINDOW_LISTSTART Then
            DeleteMenu hWindowMenu, i, MF_BYPOSITION
            FreeMenuTag miiData.dwItemData
          End If
        End If
      Next i
      .dwItemData = 0
      .dwTypeData = 0
      .cch = 0
    End With

    If openChildren.Count > 0 Then
      endOfMenu = GetMenuItemCount(hWindowMenu) + 1

      ' insert a separator
      With miiData
        .cbSize = LenB(miiData)
        .fMask = MIIM_TYPE Or MIIM_ID Or MIIM_DATA
        .fType = MFT_SEPARATOR Or MFT_OWNERDRAW
        Set tagObject = CreateMenuTagObject(Separator:=True)
        .dwItemData = MakeMenuTag(tagObject)

        .wID = CMDID_WINDOW_LISTSTART
        InsertMenuItem hWindowMenu, endOfMenu, 1, miiData

        ' insert an item for each child window
        .fType = MFT_OWNERDRAW
        For i = 1 To openChildren.Count
          Set frm = openChildren.Item(i)
          If Not frm Is Nothing Then
            If frm Is Me.ActiveForm Then
              .fMask = .fMask Or MIIM_STATE
              .fState = MFS_CHECKED
            Else
              .fMask = .fMask And Not MIIM_STATE
            End If

            .wID = CMDID_WINDOW_LISTSTART + i
            Set tagObject = CreateMenuTagObject(frm.Caption)
            .dwItemData = MakeMenuTag(tagObject)
            InsertMenuItem hWindowMenu, endOfMenu + i, 1, miiData
          End If
        Next i
      End With
    End If
  End If
End Sub
