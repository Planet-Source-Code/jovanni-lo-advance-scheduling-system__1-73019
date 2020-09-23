VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MenuItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const CLASS_NAME As String = "MenuItems"

Dim mCol    As Collection

Public Property Get Item(Index) As MenuItem
    Set Item = mCol.Item(Index)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Function Add(ByVal Caption As String, Optional ByVal Key As String, Optional ByVal Tag As Variant, Optional ByVal ItemData As Long, Optional Checked As Boolean, Optional ByVal Enabled As Boolean = True) As MenuItem
Dim objItem     As MenuItem
    Set objItem = New MenuItem
    With objItem
        .Caption = Caption
        .Key = Key
        .Tag = Tag
        .ItemData = ItemData
        .Checked = Checked
        .Enabled = Enabled
    End With
    If objItem.Key = vbNullString Then
        mCol.Add objItem
    Else
        mCol.Add objItem, objItem.Key
    End If
    Set Add = objItem
    Set objItem = Nothing
End Function

Public Function Remove(Index)
    mCol.Remove Index
End Function

Public Function Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Function

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

