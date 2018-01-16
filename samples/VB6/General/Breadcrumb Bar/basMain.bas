Attribute VB_Name = "basMain"
Option Explicit

  ' This sample was heavily inspired by the "Explorer breadcrumbs in Vista" sample by
  ' Bjarke Viksoe which can be found on http://www.viksoe.dk/code/breadcrumbs.htm


  Public Const MAX_PATH As Long = 260
  Public Const S_OK As Long = &H0


  Public Type MARGINS
    cxLeftWidth As Long
    cxRightWidth As Long
    cyTopHeight As Long
    cyBottomHeight As Long
  End Type

  Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type

  Public Type PAINTSTRUCT
    hDC As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved(0 To 31) As Byte
  End Type

  Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName(1 To MAX_PATH) As Integer
    szTypeName(1 To 80) As Integer
  End Type


  Public IID_IShellFolder As UUID


  Public Declare Function BeginPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
  Public Declare Function CloseThemeData Lib "uxtheme.dll" (ByVal hTheme As Long) As Long
  Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As Long, pclsid As UUID) As Long
  Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Public Declare Function DwmIsCompositionEnabled Lib "dwmapi.dll" (ByRef pfEnabled As Long) As Long
  Public Declare Function EndPaint Lib "user32.dll" (ByVal hWnd As Long, lpPaint As PAINTSTRUCT) As Long
  Public Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, lprc As RECT, ByVal hbr As Long) As Long
  Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT) As Long
  Public Declare Function GetStockObject Lib "gdi32.dll" (ByVal fnObject As Long) As Long
  Public Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Public Declare Function ILClone Lib "shell32.dll" (ByVal pIDL As Long) As Long
  Public Declare Function ILCreateFromPath Lib "shell32.dll" Alias "ILCreateFromPathW" (ByVal pszPath As Long) As Long
  Public Declare Sub ILFree Lib "shell32.dll" (ByVal pIDL As Long)
  Public Declare Function ILGetSize Lib "shell32.dll" (ByVal pIDL As Long) As Long
  Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
  Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszClassList As Long) As Long
  Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
  Public Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hWnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
  Public Declare Function SHBindToParent Lib "shell32.dll" (ByVal pIDL As Long, riid As UUID, ppv As IVBShellFolder, ByVal ppidlLast As Long) As Long
  Public Declare Function SHGetDesktopFolder Lib "shell32.dll" (ByRef ppshf As IVBShellFolder) As Long
  Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoW" (ByVal pszPath As Long, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long
  Public Declare Function SHGetImageList Lib "shell32.dll" (ByVal iImageList As Long, riid As UUID, ByVal ppv As Long) As Long
  Public Declare Function StrRetToBuf Lib "shlwapi.dll" Alias "StrRetToBufW" (ByRef pstr As STRRET, ByRef pIDL As Long, ByVal pszBuf As Long, ByVal cchBuf As Long) As Long


Public Function GetSysIconIndex(ByVal pIDLToDesktop As Long) As Long
  Const SHGFI_PIDL As Long = &H8
  Const SHGFI_SMALLICON As Long = &H1
  Const SHGFI_SYSICONINDEX As Long = &H4000&
  Dim fileInfo As SHFILEINFO
  Dim flags As Long

  flags = SHGFI_PIDL Or SHGFI_SYSICONINDEX Or SHGFI_SMALLICON
  SHGetFileInfo pIDLToDesktop, 0, fileInfo, LenB(fileInfo), flags
  GetSysIconIndex = fileInfo.iIcon
End Function

Public Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

Public Function IsCompositionEnabled() As Boolean
  Dim compositionEnabled As Long

  ' retrieves whether Aero Glass is enabled

  If DwmIsCompositionEnabled(compositionEnabled) < 0 Then
    compositionEnabled = 0
  End If
  IsCompositionEnabled = (compositionEnabled <> 0)
End Function

Public Function LoWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value), LenB(ret)

  LoWord = ret
End Function

Public Sub SplitFullyQualifiedPIDL(pIDLToDesktop As Long, IParent As IVBShellFolder, pIDLToParent As Long)
  If ILGetSize(pIDLToDesktop) = 2 Then
    SHGetDesktopFolder IParent
    pIDLToParent = pIDLToDesktop
  Else
    SHBindToParent pIDLToDesktop, IID_IShellFolder, IParent, VarPtr(pIDLToParent)
  End If
End Sub

Public Function STRRETToString(Data As STRRET, ByVal pIDL As Long) As String
  Dim ret As String

  ret = String$(MAX_PATH, Chr$(0))
  StrRetToBuf Data, pIDL, StrPtr(ret), MAX_PATH
  STRRETToString = Left$(ret, lstrlen(StrPtr(ret)))
End Function
