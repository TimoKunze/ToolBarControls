Attribute VB_Name = "basMain"
Option Explicit

  Public Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hWnd As Long
  End Type

  Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal SIZE As Long)
  Private Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontW" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeight As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpchText As Long, ByVal nCount As Long, ByVal lpRect As Long, ByVal uFormat As Long) As Long
  Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSysColorBrush Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
  Private Declare Function OffsetRect Lib "user32.dll" (lprc As RECT, ByVal dx As Long, ByVal dy As Long) As Long
  Private Declare Function PatBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal nXLeft As Long, ByVal nYLeft As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
  Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal iBkMode As Long) As Long
  Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal pDest As Long, ByVal Length As Long)


Public Function BlendColor(ByVal colorFrom As Long, ByVal colorTo As Long, Optional ByVal alpha As Long = 128) As Long
  Dim srcR As Long
  Dim srcG As Long
  Dim srcB As Long
  Dim dstR As Long
  Dim dstG As Long
  Dim dstB As Long

  srcR = colorFrom And &HFF
  srcG = (colorFrom And &HFF00&) \ &H100&
  srcB = (colorFrom And &HFF0000) \ &H10000
  dstR = colorTo And &HFF
  dstG = (colorTo And &HFF00&) \ &H100&
  dstB = (colorTo And &HFF0000) \ &H10000

  BlendColor = RGB( _
      (srcR * alpha + dstR * (255 - alpha)) / 255, _
      (srcG * alpha + dstG * (255 - alpha)) / 255, _
      (srcB * alpha + dstB * (255 - alpha)) / 255)
End Function

Public Sub DrawCheckMark(ByVal fontSize As Long, ByVal hDC As Long, rc As RECT, Optional ByVal enabled As Boolean = True, Optional ByVal radioButton As Boolean = False)
  Const COLOR_GRAYTEXT = 17
  Const COLOR_MENUTEXT = 7
  Const DT_CENTER = &H1
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const FW_NORMAL = 400
  Const SM_CXBORDER = 5
  Const SYMBOL_CHARSET = 2
  Const TRANSPARENT = 1
  Dim hFont As Long
  Dim hPreviousFont As Long
  Dim previousBkMode As Long
  Dim previousTextColor As Long

  hFont = CreateFont(fontSize, 0, 0, 0, FW_NORMAL, 0, 0, 0, SYMBOL_CHARSET, 0, 0, 0, 0, StrPtr("Marlett"))
  If hFont Then
    hPreviousFont = SelectObject(hDC, hFont)
    previousBkMode = SetBkMode(hDC, TRANSPARENT)
    previousTextColor = SetTextColor(hDC, GetSysColor(IIf(enabled, COLOR_MENUTEXT, COLOR_GRAYTEXT)))

    DrawText hDC, IIf(radioButton, StrPtr("h"), StrPtr("a")), 1, VarPtr(rc), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
    
    SetTextColor hDC, previousTextColor
    SetBkMode hDC, previousBkMode
    SelectObject hDC, hPreviousFont
    DeleteObject hFont
  End If
End Sub

Public Sub DrawChevron(ByVal hDC As Long, chevronRect As RECT, Optional ByVal pushed As Boolean = False)
  Const COLOR_BTNTEXT = 18
  Const PATCOPY = &HF00021
  Dim hPreviousBrush As Long
  Dim rc As RECT
  Dim x As Long
  Dim y As Long

  rc = chevronRect
  If pushed Then
    OffsetRect rc, 1, 1
  End If

  hPreviousBrush = SelectObject(hDC, GetSysColorBrush(COLOR_BTNTEXT))
  x = (rc.Left + (rc.Right - 8)) / 2
  For y = -2 To 2
    PatBlt hDC, x, rc.Top + 6 + y, 2, 1, PATCOPY
    PatBlt hDC, x + 4, rc.Top + 6 + y, 2, 1, PATCOPY
    x = x + IIf(y < 0, 1, -1)
  Next y
  SelectObject hDC, hPreviousBrush
End Sub

Public Sub DrawDropDownArrow(ByVal hDC As Long, rc As RECT, Optional ByVal enabled As Boolean = True)
  Const COLOR_BTNTEXT = 18
  Const COLOR_GRAYTEXT = 17
  Const DT_CENTER = &H1
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const FW_NORMAL = 400
  Const SM_CXBORDER = 5
  Const SYMBOL_CHARSET = 2
  Const TRANSPARENT = 1
  Dim fontSize As Long
  Dim hFont As Long
  Dim hPreviousFont As Long
  Dim previousBkMode As Long
  Dim previousTextColor As Long

  fontSize = rc.Right - rc.Left
  If rc.Bottom - rc.Top < fontSize Then fontSize = rc.Bottom - rc.Top
  fontSize = fontSize - 2 * GetSystemMetrics(SM_CXBORDER)
  hFont = CreateFont(fontSize, 0, 0, 0, FW_NORMAL, 0, 0, 0, SYMBOL_CHARSET, 0, 0, 0, 0, StrPtr("Marlett"))
  If hFont Then
    hPreviousFont = SelectObject(hDC, hFont)
    previousBkMode = SetBkMode(hDC, TRANSPARENT)
    previousTextColor = SetTextColor(hDC, GetSysColor(IIf(enabled, COLOR_BTNTEXT, COLOR_GRAYTEXT)))

    DrawText hDC, StrPtr("6"), 1, VarPtr(rc), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
    
    SetTextColor hDC, previousTextColor
    SetBkMode hDC, previousBkMode
    SelectObject hDC, hPreviousFont
    DeleteObject hFont
  End If
End Sub

Public Sub DrawSubMenuArrow(ByVal fontSize As Long, ByVal hDC As Long, rc As RECT, Optional ByVal enabled As Boolean = True)
  Const COLOR_GRAYTEXT = 17
  Const COLOR_MENUTEXT = 7
  Const DT_CENTER = &H1
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const FW_NORMAL = 400
  Const SM_CXBORDER = 5
  Const SYMBOL_CHARSET = 2
  Const TRANSPARENT = 1
  Dim hFont As Long
  Dim hPreviousFont As Long
  Dim previousBkMode As Long
  Dim previousTextColor As Long

  hFont = CreateFont(fontSize, 0, 0, 0, FW_NORMAL, 0, 0, 0, SYMBOL_CHARSET, 0, 0, 0, 0, StrPtr("Marlett"))
  If hFont Then
    hPreviousFont = SelectObject(hDC, hFont)
    previousBkMode = SetBkMode(hDC, TRANSPARENT)
    previousTextColor = SetTextColor(hDC, GetSysColor(IIf(enabled, COLOR_MENUTEXT, COLOR_GRAYTEXT)))

    DrawText hDC, StrPtr("8"), 1, VarPtr(rc), DT_CENTER Or DT_SINGLELINE Or DT_VCENTER
    
    SetTextColor hDC, previousTextColor
    SetBkMode hDC, previousBkMode
    SelectObject hDC, hPreviousFont
    DeleteObject hFont
  End If
End Sub

Public Function ExtractMenuTag(pMenuTag As Long) As CMenuItemData
  Dim tagObject As CMenuItemData
  
  If pMenuTag Then
    CopyMemory VarPtr(tagObject), VarPtr(pMenuTag), 4
    Set ExtractMenuTag = tagObject
    ZeroMemory VarPtr(tagObject), 4
  End If
End Function

Public Sub FreeMenuTag(pMenuTag As Long)
  Dim tagObject As CMenuItemData
  
  If pMenuTag Then
    CopyMemory VarPtr(tagObject), VarPtr(pMenuTag), 4
    pMenuTag = 0
    Set tagObject = Nothing
  End If
End Sub

Public Function MakeMenuTag(Tag As CMenuItemData) As Long
  If Not Tag Is Nothing Then
    MakeMenuTag = ObjPtr(Tag)
    ZeroMemory VarPtr(Tag), 4
  End If
End Function
