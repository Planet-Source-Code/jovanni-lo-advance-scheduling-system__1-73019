VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsgHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const CLASS_NAME As String = "MsgHook"

Public Event BeforeMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef RetValue As Long, ByRef Cancel As Boolean)
Public Event AfterMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Private mlnghWnd        As Long
Private mlngOldWndProc  As Long
    
Friend Property Get hWnd() As Long
    hWnd = mlnghWnd
End Property

Friend Property Let OldWndProc(ByVal plngValue As Long)
    mlngOldWndProc = plngValue
End Property

Friend Property Get OldWndProc() As Long
    OldWndProc = mlngOldWndProc
End Property
    
Public Sub StartSubclass(ByVal hWnd As Long)
    If mlnghWnd = 0 Then
        mlnghWnd = hWnd
        If mlnghWnd Then
            Call HookWindow(Me)
        End If
    End If
End Sub

Public Sub StopSubclass()
    If mlnghWnd > 0 Then
        Call UnhookWindow(Me)
        mlnghWnd = 0
    End If
End Sub

Friend Function WndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim blnCancel   As Boolean
Dim lngRet      As Long
    RaiseEvent BeforeMessage(hWnd, uMsg, wParam, lParam, lngRet, blnCancel)
    If Not blnCancel And lngRet = 0 Then
        lngRet = CallWindowProc(mlngOldWndProc, hWnd, uMsg, wParam, lParam)
    End If
    WndProc = lngRet
    RaiseEvent AfterMessage(hWnd, uMsg, wParam, lParam)
End Function

Private Sub Class_Terminate()
    Call StopSubclass
End Sub
