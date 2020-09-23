VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CLASS_NAME As String = "MenuItem"

Dim strCaption      As String
Dim strKey          As String
Dim varTag          As Variant
Dim lngItemData     As Long
Dim blnChecked      As Boolean
Dim blnEnabled      As Boolean

Public Property Let Caption(ByVal Value As String)
    strCaption = Value
End Property

Public Property Get Caption() As String
    Caption = strCaption
End Property

Public Property Let Key(ByVal Value As String)
    strKey = Value
End Property

Public Property Get Key() As String
    Key = strKey
End Property

Public Property Let Tag(ByVal Value As Variant)
    varTag = Value
End Property

Public Property Get Tag() As Variant
    Tag = varTag
End Property

Public Property Let ItemData(ByVal Value As Long)
    lngItemData = Value
End Property

Public Property Get ItemData() As Long
    ItemData = lngItemData
End Property

Public Property Let Checked(ByVal Value As Boolean)
    blnChecked = Value
End Property

Public Property Get Checked() As Boolean
    Checked = blnChecked
End Property

Public Property Let Enabled(ByVal Value As Boolean)
    blnEnabled = Value
End Property

Public Property Get Enabled() As Boolean
    Enabled = blnEnabled
End Property

Private Sub Class_Initialize()
    strCaption = vbNullString
    strKey = vbNullString
    varTag = vbNullString
    lngItemData = 0
    blnEnabled = True
End Sub

