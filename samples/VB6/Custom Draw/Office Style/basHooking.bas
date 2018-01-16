Attribute VB_Name = "basHooking"
Option Explicit

  Private hCallWndProcHook As Long
  Private pCallWndProcHookConsumers() As Long


  Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
  Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExW" (ByVal hookType As Long, ByVal pHookProc As Long, ByVal hInst As Long, ByVal threadID As Long) As Long
  Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
  Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal pDest As Long, ByVal sz As Long)


' registers <Obj> for a CallWndProc hook
Sub InstallCallWndProcHook(ByVal Obj As IHook)
  Const WH_CALLWNDPROC = 4
  Dim exists As Boolean
  Dim i As Integer
  Dim iAvailableSlot As Integer
  Dim pObj As Long

  ' if we didn't yet, register the hook now
  If hCallWndProcHook = 0 Then hCallWndProcHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf CallWndProc, 0&, GetCurrentThreadId)
  If hCallWndProcHook Then
    ' now search the array of listeners for a free slot
    iAvailableSlot = -1
    pObj = ObjPtr(Obj)
    On Error GoTo initArray
    i = LBound(pCallWndProcHookConsumers)
    For i = LBound(pCallWndProcHookConsumers) To UBound(pCallWndProcHookConsumers)
      If pCallWndProcHookConsumers(i) = pObj Then
        ' the object that wanted this hook, already has it
        exists = True
      ElseIf pCallWndProcHookConsumers(i) = 0 Then
        ' we found a free slot
        iAvailableSlot = i
        Exit For
      End If
    Next

    If Not exists Then
      If iAvailableSlot = 0 Then
        ' no free slot was found, so resize the array
        ReDim Preserve pCallWndProcHookConsumers(UBound(pCallWndProcHookConsumers) + 1) As Long
        iAvailableSlot = UBound(pCallWndProcHookConsumers)
      End If
      ' store the new listener
      pCallWndProcHookConsumers(iAvailableSlot) = pObj
    End If
  End If
  Exit Sub

initArray:
  ReDim pCallWndProcHookConsumers(0)
  Resume Next
End Sub

' the callback method for a CallWndProc hook
Function CallWndProc(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Dim eatIt As Boolean
  Dim i As Integer
  Dim Obj As IHook

  On Error GoTo MyError
  If hookCode >= 0 Then
    ' call each listener before we forward this message to the next hook
    For i = LBound(pCallWndProcHookConsumers) To UBound(pCallWndProcHookConsumers)
      If pCallWndProcHookConsumers(i) Then
        Set Obj = GetObjectFromPtr(pCallWndProcHookConsumers(i))
        Obj.CallWndProcBefore hookCode, wParam, lParam, eatIt
      End If
    Next
    If eatIt Then
      ' bon appetit :)
      CallWndProc = 1
    Else
      ' let the next hook do his job
      CallWndProc = CallNextHookEx(hCallWndProcHook, hookCode, wParam, lParam)
      ' call each listener again
      For i = LBound(pCallWndProcHookConsumers) To UBound(pCallWndProcHookConsumers)
        If pCallWndProcHookConsumers(i) Then
          Set Obj = GetObjectFromPtr(pCallWndProcHookConsumers(i))
          Obj.CallWndProcAfter hookCode, wParam, lParam
        End If
      Next
    End If
  Else
    CallWndProc = CallNextHookEx(hCallWndProcHook, hookCode, wParam, lParam)
  End If
  Exit Function

MyError:
  CallWndProc = CallNextHookEx(hCallWndProcHook, hookCode, wParam, lParam)
End Function

' removes <Obj> from the listeners of the CallWndProc hook
Sub RemoveCallWndProcHook(ByVal Obj As IHook)
  Dim countRefs As Integer
  Dim i As Integer
  Dim pObj As Long

  ' find the object in the array of listeners
  pObj = ObjPtr(Obj)
  On Error GoTo MyError
  For i = LBound(pCallWndProcHookConsumers) To UBound(pCallWndProcHookConsumers)
    If pCallWndProcHookConsumers(i) = pObj Then
      ' found it, free its slot
      pCallWndProcHookConsumers(i) = 0
    ElseIf pCallWndProcHookConsumers(i) Then
      ' there's still another object listening
      countRefs = countRefs + 1
    End If
  Next

  If countRefs = 0 Then
    ' all listeners are gone, so unregister the hook
    UnhookWindowsHookEx hCallWndProcHook
    hCallWndProcHook = 0
  End If

MyError:
End Sub


' returns the object <Ptr> points to
Private Function GetObjectFromPtr(ByVal Ptr As Long) As Object
  Dim ret As Object

  ' get the object <Ptr> points to
  CopyMemory VarPtr(ret), VarPtr(Ptr), 4
  Set GetObjectFromPtr = ret
  ' free memory
  ZeroMemory VarPtr(ret), 4
End Function
