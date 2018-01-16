VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IVisualManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Sub InitializeReBar(ByVal ctl As ReBar)
  '
End Sub

Public Sub InitializeToolBar(ByVal ctl As ToolBar)
  '
End Sub

Public Sub ReBarCustomDraw(ByVal ctl As ReBar, ByVal band As TBarCtlsLibUCtl.IReBarBand, ByVal drawStage As TBarCtlsLibUCtl.CustomDrawStageConstants, ByVal bandState As TBarCtlsLibUCtl.CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.CustomDrawReturnValuesConstants)
  '
End Sub

Public Sub ReBarRawMenuMessage(ByVal ctl As ReBar, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  '
End Sub

Public Sub SetMenuManager(ByVal manager As IMenuManager)
  '
End Sub

Public Sub ToolBarBeforeDisplayChevronPopup(ByVal ctl As ToolBar, ByVal hPopup As Long, ByVal x As Single, ByVal y As Single, ByVal isMenu As Boolean, cancelPopup As Boolean, commandToExecute As Long)
  '
End Sub

Public Sub ToolBarCustomDraw(ByVal ctl As ToolBar, ByVal toolButton As ToolBarButton, normalTextColor As stdole.OLE_COLOR, normalButtonBackColor As stdole.OLE_COLOR, normalBackgroundMode As StringBackgroundModeConstants, hotTextColor As stdole.OLE_COLOR, hotButtonBackColor As stdole.OLE_COLOR, markedTextBackColor As stdole.OLE_COLOR, markedButtonBackColor As stdole.OLE_COLOR, markedBackgroundMode As StringBackgroundModeConstants, ByVal drawStage As CustomDrawStageConstants, ByVal buttonState As CustomDrawItemStateConstants, ByVal hDC As Long, drawingRectangle As RECTANGLE, textRectangle As RECTANGLE, HorizontalIconCaptionGap As stdole.OLE_XSIZE_PIXELS, furtherProcessing As CustomDrawReturnValuesConstants)
  '
End Sub

Public Sub ToolBarDestroyingChevronPopup(ByVal ctl As IToolBar, ByVal hPopup As Long, ByVal isMenu As Boolean)
  '
End Sub

Public Sub ToolBarDropDown(ByVal ctl As ToolBar, ByVal toolButton As ToolBarButton, buttonRectangle As TBarCtlsLibUCtl.RECTANGLE, furtherProcessing As TBarCtlsLibUCtl.DropDownReturnValuesConstants)
  '
End Sub

Public Sub ToolBarRawMenuMessage(ByVal ctl As ToolBar, ByVal message As Long, ByVal wParam As Long, ByVal lParam As Long, result As Long, handledEvent As Boolean)
  '
End Sub
