VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenuItemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

  Private m_HasSubMenu As Boolean
  Private m_hImageList As Long
  Private m_IconIndex As Long
  Private m_RadioCheckMark As Boolean
  Private m_Separator As Boolean
  Private m_Text As String

Private Sub Class_Initialize()
  m_IconIndex = -2
End Sub


Public Property Get HasSubMenu() As Boolean
  HasSubMenu = m_HasSubMenu
End Property

Public Property Let HasSubMenu(ByVal newValue As Boolean)
  m_HasSubMenu = newValue
End Property

Public Property Get hImageList() As Long
  hImageList = m_hImageList
End Property

Public Property Let hImageList(ByVal newValue As Long)
  m_hImageList = newValue
End Property

Public Property Get IconIndex() As Long
  IconIndex = m_IconIndex
End Property

Public Property Let IconIndex(ByVal newValue As Long)
  m_IconIndex = newValue
End Property

Public Property Get RadioCheckMark() As Boolean
  RadioCheckMark = m_RadioCheckMark
End Property

Public Property Let RadioCheckMark(ByVal newValue As Boolean)
  m_RadioCheckMark = newValue
End Property

Public Property Get Separator() As Boolean
  Separator = m_Separator
End Property

Public Property Let Separator(ByVal newValue As Boolean)
  m_Separator = newValue
End Property

Public Property Get Text() As String
  Text = m_Text
End Property

Public Property Let Text(ByVal newValue As String)
  m_Text = newValue
End Property
