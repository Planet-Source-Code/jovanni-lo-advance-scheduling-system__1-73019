Attribute VB_Name = "modMsgHook"
Option Explicit

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_WNDPROC As Long = -4
Private Const WM_DESTROY As Long = &H2

Private mcolItems       As Collection

Public Sub HookWindow(ByRef pobjMsgHook As MsgHook)
Dim lnghWnd As Long
    lnghWnd = pobjMsgHook.hWnd
    pobjMsgHook.OldWndProc = SetWindowLong(lnghWnd, GWL_WNDPROC, AddressOf WndProc)
    If mcolItems Is Nothing Then
        Set mcolItems = New Collection
    End If
    mcolItems.Add ObjPtr(pobjMsgHook), lnghWnd & "K"
End Sub

Public Sub UnhookWindow(ByRef pobjMsgHook As MsgHook)
    Call SetWindowLong(pobjMsgHook.hWnd, GWL_WNDPROC, pobjMsgHook.OldWndProc)
    mcolItems.Remove pobjMsgHook.hWnd & "K"
    If mcolItems.Count = 0 Then
        Set mcolItems = Nothing
    End If
End Sub

Public Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lngPointer  As Long
Dim objMsgHook  As MsgHook
    lngPointer = mcolItems.Item(hWnd & "K")
    Set objMsgHook = PtrObj(lngPointer)
    WndProc = objMsgHook.WndProc(hWnd, uMsg, wParam, lParam)
    If uMsg = WM_DESTROY Then
        Call objMsgHook.StopSubclass
    End If
    Set objMsgHook = Nothing
End Function
