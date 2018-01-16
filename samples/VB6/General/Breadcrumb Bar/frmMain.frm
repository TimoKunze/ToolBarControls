VERSION 5.00
Object = "{5D0D0ABC-4898-4E46-AB48-291074A737A1}#1.3#0"; "TBarCtlsU.ocx"
Begin VB.Form frmMain 
   Caption         =   "ToolBarControls 1.3 - Breadcrumb Bar Sample"
   ClientHeight    =   2850
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   361
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      Height          =   1575
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   5055
   End
   Begin BreadcrumbBar.UCAddressBand AddressBand 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
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
      Left            =   1980
      TabIndex        =   3
      Top             =   2400
      Width           =   1455
   End
   Begin TBarCtlsLibUCtl.ReBar NavBar 
      Align           =   1  'Oben ausrichten
      Height          =   660
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   5415
      _cx             =   9551
      _cy             =   1164
      AllowBandReordering=   -1  'True
      Appearance      =   0
      AutoUpdateLayout=   -1  'True
      BackColor       =   -1
      BorderStyle     =   0
      DisabledEvents  =   524779
      DisplayBandSeparators=   0   'False
      DisplaySplitter =   0   'False
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      FixedBandHeight =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      RegisterForOLEDragDrop=   0
      RightToLeft     =   0
      ShadowColor     =   -1
      SupportOLEDragImages=   -1  'True
      ToggleOnDoubleClick=   0   'False
      UseSystemFont   =   -1  'True
      VerticalSizingGripsOnVerticalOrientation=   -1  'True
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  ' This sample was heavily inspired by the "Explorer breadcrumbs in Vista" sample by
  ' Bjarke Viksoe which can be found on http://www.viksoe.dk/code/breadcrumbs.htm


  Implements ISubclassedWindow


  Private Type OSVERSIONINFOEX
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion(1 To 256) As Byte
    wServicePackMajor As Integer
    wServicePackMinor As Integer
    wSuiteMask As Integer
    wProductType As Byte
    wReserved As Byte
  End Type


  Private Const CX_ADDRESS As Long = 250
  Private Const CX_MIN_ADDRESS As Long = 200
  Private Const CY_ADDRESS As Long = 24
  Private Const CY_NAVBAR As Long = 35


  Private Declare Function DwmExtendFrameIntoClientArea Lib "dwmapi.dll" (ByVal hWnd As Long, ByRef pMarInset As MARGINS) As Long
  Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
  Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExW" (lpVersionInfo As OSVERSIONINFOEX) As Long
  Private Declare Function ILClone Lib "shell32.dll" (ByVal pIDL As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long


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
  Const WM_ERASEBKGND As Long = &H14
  Const WM_PAINT As Long = &HF
  Const WM_PRINTCLIENT As Long = &H318
  Dim hDC As Long
  Dim ps As PAINTSTRUCT
  Dim lRet As Long

  On Error GoTo StdHandler_Error
  Select Case uMsg
    Case WM_ERASEBKGND
      If Not IsCompositionEnabled Then DoPaint wParam
      lRet = 1
      bCallDefProc = False
    Case WM_PAINT, WM_PRINTCLIENT
      If wParam Then
        DoPaint wParam
      Else
        hDC = BeginPaint(hWnd, ps)
        If IsCompositionEnabled Then DoPaint hDC
        EndPaint hWnd, ps
      End If
      bCallDefProc = False
  End Select

StdHandler_Ende:
  HandleMessage_Form = lRet
  Exit Function

StdHandler_Error:
  Debug.Print "Error in frmMain.HandleMessage_Form: ", Err.Number, Err.Description
  Resume StdHandler_Ende
End Function


Private Sub AddressBand_ChangeCurrentPIDL(ByVal pIDL As Long)
  AddressBand.CurrentPIDL = ILClone(pIDL)
End Sub

Private Sub cmdAbout_Click()
  NavBar.About
End Sub

Private Sub Form_Activate()
  Dim windowMargins As MARGINS

  If IsCompositionEnabled Then
    ' extend Aero Glass border into the client area so that the search box sits on the glass area
    windowMargins.cyTopHeight = CY_NAVBAR
    DwmExtendFrameIntoClientArea Me.hWnd, windowMargins
  End If
  Me.Refresh
End Sub

Private Sub Form_Initialize()
  Dim hTheme As Long
  Dim OSVerData As OSVERSIONINFOEX

  InitCommonControls

  With OSVerData
    .dwOSVersionInfoSize = LenB(OSVerData)
    GetVersionEx OSVerData
    If .dwMajorVersion < 6 Then
      MsgBox "This sample requires Windows Vista or newer.", VbMsgBoxStyle.vbCritical, "Error"
      End
    End If
  End With
  hTheme = OpenThemeData(Me.hWnd, StrPtr("BreadcrumbBarComposited::BreadcrumbBar"))
  If hTheme = 0 Then
    MsgBox "Cannot open theme class data." & vbNewLine & "Wrong version of Windows or no theming supported.", VbMsgBoxStyle.vbCritical, "Error"
    End
  End If
  CloseThemeData hTheme
End Sub

Private Sub Form_Load()
  Const CSIDL_DESKTOP As Long = &H0
  Dim band As ReBarBand
  Dim pIDLDesktop As Long

  ' change theme of the rebar to a special one
  SetWindowTheme NavBar.hWnd, StrPtr("NavbarComposited"), 0
  ' subclass the main form in order to draw the background
  If Not SubclassWindow(Me.hWnd, Me, EnumSubclassID.escidFrmMain) Then
    Debug.Print "Subclassing failed!"
  End If

  Set band = NavBar.Bands.Add(, AddressBand.hWnd, , sgvNever, False, , , CX_ADDRESS, CX_MIN_ADDRESS, CY_ADDRESS, CY_ADDRESS)
  SHGetFolderLocation Me.hWnd, CSIDL_DESKTOP, 0, 0, pIDLDesktop
  AddressBand.CurrentPIDL = pIDLDesktop
End Sub

Private Sub Form_Resize()
  NavBar.Height = CY_NAVBAR
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not UnSubclassWindow(Me.hWnd, EnumSubclassID.escidFrmMain) Then
    Debug.Print "UnSubclassing failed!"
  End If
End Sub


Private Sub DoPaint(ByVal hDC As Long)
  Const BLACK_BRUSH As Long = 4
  Const COLOR_BTNFACE As Long = 15
  Const COLOR_GRADIENTACTIVECAPTION As Long = 27
  Const COLOR_GRADIENTINACTIVECAPTION As Long = 28
  Dim rcClient As RECT
  Dim hBrush As Long

  ' at first draw the background of the rebar
  GetClientRect Me.hWnd, rcClient
  rcClient.Bottom = rcClient.Top + CY_NAVBAR
  If IsCompositionEnabled Then
    ' We extended the Glass border into the client area. To make this work entirely, we have to
    ' paint this part of the client area black.
    hBrush = GetStockObject(BLACK_BRUSH)
  Else
    ' draw the rebar's background in a color that looks a bit like Aero Glass
    If GetActiveWindow = Me.hWnd Then
      hBrush = GetSysColorBrush(COLOR_GRADIENTACTIVECAPTION)
    Else
      hBrush = GetSysColorBrush(COLOR_GRADIENTINACTIVECAPTION)
    End If
  End If
  If hBrush Then
    ' do the actual drawing
    FillRect hDC, rcClient, hBrush
  End If

  ' now draw the remaining part of the client area
  GetClientRect Me.hWnd, rcClient
  rcClient.Top = rcClient.Top + CY_NAVBAR
  FillRect hDC, rcClient, GetSysColorBrush(COLOR_BTNFACE)
End Sub
