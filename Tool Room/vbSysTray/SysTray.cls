VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const CLASS_NAME As String = "SystemTray"

Private mudtSysTray                 As NOTIFYICONDATA

Private WithEvents mobjMsgWindow    As MsgWindow
Attribute mobjMsgWindow.VB_VarHelpID = -1

Private mblnVisible                 As Boolean
Private mblnEnableMenu              As Boolean
Private milsImageList               As MSComctlLib.ImageList
Private mvarIcon                    As Variant
Private mstrToolTipText             As String

Private mobjMenuItems               As MenuItems

Private Const CONTROL_NAME As String = "SystemTray Class"

Public Event Click()
Public Event MouseDown(ByVal Button As MouseButtonConstants)
Public Event MouseUp(ByVal Button As MouseButtonConstants)
Public Event MouseMove()
Public Event DoubleClick(ByVal Button As MouseButtonConstants)
Public Event MenuClick(Item As MenuItem)

Public Property Get Menu() As MenuItems
    Set Menu = mobjMenuItems
End Property

Public Property Let Visible(ByVal Value As Boolean)
    If Value = Not mblnVisible Then
        mblnVisible = Value
        ShowInTray
    End If
End Property

Public Property Get Visible() As Boolean
    Visible = mblnVisible
End Property

Public Property Let EnableMenu(ByVal Value As Boolean)
    If Value = Not mblnEnableMenu Then
        mblnEnableMenu = Value
    End If
End Property

Public Property Get EnableMenu() As Boolean
    EnableMenu = mblnEnableMenu
End Property

Public Property Set ImageList(ByRef Value As Object)
    If TypeOf Value Is MSComctlLib.ImageList Then
        Set milsImageList = Value
    Else
        Err.Raise 35610, CONTROL_NAME, "Invalid Object"
    End If
End Property

Public Property Get ImageList() As Object
    Set ImageList = milsImageList
End Property

Public Property Let Icon(ByVal Value As Variant)
    If mvarIcon <> Value Then
        mudtSysTray.hIcon = milsImageList.ListImages.Item(Value).Picture.Handle
        mudtSysTray.uFlags = NIF_ICON
        mvarIcon = Value
        ModifySystemTray
    End If
End Property

Public Property Get Icon() As Variant
    Icon = mvarIcon
End Property

Public Property Let ToolTipText(ByVal pstrValue As String)
    If mstrToolTipText <> Trim$(pstrValue) Then
        mudtSysTray.szTip = Trim$(pstrValue) & vbNullChar
        mudtSysTray.uFlags = NIF_TIP
        mstrToolTipText = Trim$(pstrValue)
        ModifySystemTray
    End If
End Property

Public Property Get ToolTipText() As String
    ToolTipText = mstrToolTipText
End Property

Private Sub Class_Initialize()
    Set mobjMsgWindow = New MsgWindow
    Set mobjMenuItems = New MenuItems
    With mudtSysTray
        .cbSize = Len(mudtSysTray)
        .hWnd = mobjMsgWindow.hWnd
        .UID = 0&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = TRAY_CALLBACK
        .hIcon = 0
        .szTip = vbNullChar
    End With
End Sub

Private Sub Class_Terminate()
    Visible = False
    Set mobjMsgWindow = Nothing
    Set mobjMenuItems = Nothing
End Sub

Private Sub ShowInTray()
    If Visible Then
        mudtSysTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        Shell_NotifyIcon NIM_ADD, mudtSysTray
    Else
        mudtSysTray.uFlags = 0&
        Shell_NotifyIcon NIM_DELETE, mudtSysTray
    End If
End Sub

Private Sub ModifySystemTray()
    If Visible Then
        Shell_NotifyIcon NIM_MODIFY, mudtSysTray
    End If
End Sub

Private Sub ShowMenu()
Dim lnghWnd         As Long
Dim lngMenuFlags    As Long
Dim lngIndex        As Long
Dim lngReturn       As Long
Dim objMenuItem     As MenuItem
Dim strCaption      As String
Dim Pt              As POINTAPI
    If Menu.Count > 0 Then
        lnghWnd = CreatePopupMenu
        For lngIndex = 1 To Menu.Count
            Set objMenuItem = Menu.Item(lngIndex)
            If objMenuItem.Caption = "-" Then
                strCaption = vbNullString
                lngMenuFlags = MF_SEPARATOR
            Else
                strCaption = objMenuItem.Caption
                If objMenuItem.Enabled Then
                    If objMenuItem.Checked Then
                        lngMenuFlags = MF_CHECKED
                    Else
                        lngMenuFlags = MF_STRING
                    End If
                Else
                    lngMenuFlags = MF_GRAYED Or MF_DISABLED
                End If
            End If
            AppendMenu lnghWnd, lngMenuFlags, lngIndex, strCaption
        Next lngIndex
        GetCursorPos Pt
        SetForegroundWindow mobjMsgWindow.hWnd
        lngReturn = TrackPopupMenuEx(lnghWnd, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, Pt.x, Pt.y, mobjMsgWindow.hWnd, ByVal 0&)
        DestroyMenu lnghWnd
        If lngReturn > 0 Then
            RaiseEvent MenuClick(mobjMenuItems.Item(lngReturn))
        End If
    End If
End Sub

Private Sub SendEvent(ByVal MouseEvent As Long)
    Select Case MouseEvent
        Case WM_MOUSEMOVE
            RaiseEvent MouseMove
        Case WM_LBUTTONDOWN
            RaiseEvent MouseDown(vbLeftButton)
        Case WM_LBUTTONUP
            RaiseEvent MouseUp(vbLeftButton)
        Case WM_LBUTTONDBLCLK
            RaiseEvent DoubleClick(vbLeftButton)
        Case WM_RBUTTONDOWN
            RaiseEvent MouseDown(vbRightButton)
        Case WM_RBUTTONUP
            RaiseEvent MouseUp(vbRightButton)
            If EnableMenu Then
                ShowMenu
            End If
        Case WM_RBUTTONDBLCLK
            RaiseEvent DoubleClick(vbRightButton)
    End Select
End Sub

Private Sub mobjMsgWindow_BeforeMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, RetValue As Long, Cancel As Boolean)
    If uMsg = TRAY_CALLBACK Then
        SendEvent lParam
        Cancel = True
    End If
End Sub
