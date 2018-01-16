VERSION 5.00
Begin VB.Form frmDocument 
   Caption         =   "Document 1"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   243
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   439
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   1440
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'Kein
      Height          =   1215
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   840
      Width           =   2295
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IDocument


  Private changed As Boolean
  Private windowID As Long


  Public mdiFrame As IMDIFrame


  Private Declare Function MessageBox Lib "user32.dll" Alias "MessageBoxW" (ByVal hWnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal uType As Long) As Long


Private Sub IDocument_ExecuteCommand(ByVal commandID As Long)
  Const MB_OK = &H0

  Select Case commandID
    Case CMDID_FILE_SAVE, CMDID_FILE_PRINT
      ' The VB6 MsgBox command doesn't work well together with custom draw. It probably disables any events.
      'MsgBox "Command not implemented"
      MessageBox Me.hWnd, StrPtr("Command not implemented"), 0, MB_OK
    Case CMDID_EDIT_CUT
      Clipboard.SetText Text1.SelText
      Text1.SelText = ""
      If Not mdiFrame Is Nothing Then
        mdiFrame.EnableDisableCommand CMDID_EDIT_CUT, False
        mdiFrame.EnableDisableCommand CMDID_EDIT_COPY, False
        mdiFrame.EnableDisableCommand CMDID_EDIT_PASTE, True
      End If
    Case CMDID_EDIT_COPY
      Clipboard.SetText Text1.SelText
      If Not mdiFrame Is Nothing Then
        mdiFrame.EnableDisableCommand CMDID_EDIT_PASTE, True
      End If
    Case CMDID_EDIT_PASTE
      Text1.SelText = Clipboard.GetText
    Case CMDID_EDIT_SELECTALL
      Text1.SelStart = 0
      Text1.SelLength = Len(Text1.Text)
    Case CMDID_FORMAT_BOLD
      Text1.FontBold = Not Text1.FontBold
    Case CMDID_FORMAT_ITALIC
      Text1.FontItalic = Not Text1.FontItalic
    Case CMDID_FORMAT_UNDERLINE
      Text1.FontUnderline = Not Text1.FontUnderline
    Case CMDID_FORMAT_ALIGNLEFT
      Text1.Alignment = vbLeftJustify
    Case CMDID_FORMAT_ALIGNCENTER
      Text1.Alignment = vbCenter
    Case CMDID_FORMAT_ALIGNRIGHT
      Text1.Alignment = vbRightJustify
  End Select
End Sub

Private Sub IDocument_SetFontSize(ByVal fontSize As Long)
  Text1.fontSize = fontSize
End Sub

Private Sub Form_GotFocus()
  Text1.SetFocus
End Sub

Private Sub Form_Load()
  If Not mdiFrame Is Nothing Then
    mdiFrame.OpenedMDIChild Me, windowID
  End If
  Me.Caption = "Document " & CStr(windowID)
  Clipboard.Clear
End Sub

Private Sub Form_Resize()
  Text1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not mdiFrame Is Nothing Then
    mdiFrame.ClosingMDIChild windowID
  End If
End Sub

Private Sub Text1_Change()
  changed = True
End Sub

Private Sub Text1_GotFocus()
  Timer1.enabled = True
  If Not mdiFrame Is Nothing Then
    mdiFrame.CheckCommand CMDID_FORMAT_BOLD, Text1.FontBold
    mdiFrame.CheckCommand CMDID_FORMAT_ITALIC, Text1.FontItalic
    mdiFrame.CheckCommand CMDID_FORMAT_UNDERLINE, Text1.FontUnderline
    mdiFrame.CheckCommand CMDID_FORMAT_ALIGNLEFT, Text1.Alignment = vbLeftJustify
    mdiFrame.CheckCommand CMDID_FORMAT_ALIGNCENTER, Text1.Alignment = vbCenter
    mdiFrame.CheckCommand CMDID_FORMAT_ALIGNRIGHT, Text1.Alignment = vbRightJustify
    mdiFrame.UpdateFontSize Text1.fontSize
  End If
End Sub

Private Sub Text1_LostFocus()
  Timer1.enabled = False
End Sub

Private Sub Timer1_Timer()
  If Not mdiFrame Is Nothing Then
    mdiFrame.EnableDisableCommand CMDID_FILE_SAVE, changed
    mdiFrame.EnableDisableCommand CMDID_EDIT_PASTE, Clipboard.GetFormat(vbCFText)
    mdiFrame.EnableDisableCommand CMDID_EDIT_CUT, Text1.SelLength > 0
    mdiFrame.EnableDisableCommand CMDID_EDIT_COPY, Text1.SelLength > 0
  End If
End Sub
