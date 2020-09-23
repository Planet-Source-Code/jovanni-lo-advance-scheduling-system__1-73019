VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MsgWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const CLASS_NAME As String = "MsgWindow"

Public Event BeforeMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByRef RetValue As Long, ByRef Cancel As Boolean)
Public Event AfterMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private mlnghWnd            As Long
Private mstrWindowName      As String

Private WithEvents mobjHook        As MsgHook
Attribute mobjHook.VB_VarHelpID = -1

Public Property Get hWnd() As Long
    hWnd = mlnghWnd
End Property

Public Property Get WindowName() As String
    WindowName = mstrWindowName
End Property

Private Sub CreateWindow()
    mstrWindowName = CreateNewUID
    If mlnghWnd = 0 Then
        mlnghWnd = CreateWindowEx(0, "STATIC", mstrWindowName, 0, 0, 0, 0, 0, 0, 0, App.hInstance, ByVal 0&)
        Set mobjHook = New MsgHook
        Call mobjHook.StartSubclass(mlnghWnd)
    End If
End Sub

Private Sub Destroy()
    If mlnghWnd > 0 Then
        Call mobjHook.StopSubclass
        Set mobjHook = Nothing
        Call DestroyWindow(mlnghWnd)
    End If
End Sub

Private Sub Class_Initialize()
    Call CreateWindow
End Sub

Private Sub Class_Terminate()
    Call Destroy
End Sub

Private Sub mobjHook_AfterMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
    RaiseEvent AfterMessage(hWnd, uMsg, wParam, lParam)
End Sub

Private Sub mobjHook_BeforeMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, RetValue As Long, Cancel As Boolean)
    RaiseEvent BeforeMessage(hWnd, uMsg, wParam, lParam, RetValue, Cancel)
End Sub

