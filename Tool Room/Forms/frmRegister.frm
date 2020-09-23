VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   6015
      Begin VB.TextBox txtDescription 
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   840
         Width           =   4215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Item ID:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblId 
         AutoSize        =   -1  'True
         Caption         =   "---"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   6015
      Begin VB.TextBox txtSrchStr 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         ForeColor       =   &H80000011&
         Height          =   300
         Left            =   2640
         TabIndex        =   7
         Text            =   "Search"
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmRegister.frx":038A
         Left            =   240
         List            =   "frmRegister.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   2280
         TabIndex        =   8
         Top             =   360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   5400
         Picture         =   "frmRegister.frx":038E
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.ComboBox cboFilterInven 
      Height          =   315
      ItemData        =   "frmRegister.frx":0C58
      Left            =   120
      List            =   "frmRegister.frx":0C5A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   6000
      Width           =   1935
   End
   Begin VB.TextBox txtSrchStrInven 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """Php""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   13321
         SubFormatType   =   2
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Text            =   "Search"
      Top             =   6000
      Width           =   2895
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lvwItems 
      Height          =   1575
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2778
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Item ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Category"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Location"
         Object.Width           =   2540
      EndProperty
   End
   Begin ComctlLib.ListView lvwInventory 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   6480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Item ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "QTY"
         Object.Width           =   529
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   2280
      TabIndex        =   14
      Top             =   6120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdClear 
      Height          =   375
      Left            =   1920
      TabIndex        =   19
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Clear"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmRegister.frx":0C5C
      Picture         =   "frmRegister.frx":0E36
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdView 
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Inventory <<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmRegister.frx":118A
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLoad 
      Height          =   375
      Left            =   3360
      TabIndex        =   21
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Load"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmRegister.frx":1364
      Picture         =   "frmRegister.frx":153E
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdRegister 
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   5400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Register"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmRegister.frx":1892
      Picture         =   "frmRegister.frx":1A6C
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdRemove 
      Height          =   375
      Left            =   3360
      TabIndex        =   23
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Remove"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmRegister.frx":1DC0
      Picture         =   "frmRegister.frx":1F9A
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   4800
      TabIndex        =   24
      Top             =   8280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Close"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmRegister.frx":22EE
      Picture         =   "frmRegister.frx":24C8
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Register an item to Inventory list"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   165
      Left            =   720
      TabIndex        =   18
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Left            =   720
      TabIndex        =   17
      Top             =   120
      Width           =   900
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmRegister.frx":281C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of unregistered Items on Database."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   3360
      Width           =   3360
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmRegister.frx":3460
      Top             =   8280
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Remove to unregister an item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   480
      TabIndex        =   15
      Top             =   8400
      Width           =   2490
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   5640
      Picture         =   "frmRegister.frx":3D2A
      Top             =   5880
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000016&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ClrFlds()
lblId.Caption = "---"
txtDescription.Text = Empty
ViewUnreg "item_id", "%"
viewAvailable "item_id", "%"
End Sub

Private Sub cboFilter_Click()
txtSrchStr_Change
End Sub

Private Sub cboFilterInven_Click()
txtSrchStrInven_Change
End Sub

Private Sub cmdClear_Click()
ClrFlds
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdLoad_Click()
lvwItems_DblClick
End Sub

Private Sub cmdRegister_Click()
If lblId.Caption = "---" Then MsgBox "Please load an item from the list.", vbExclamation: Exit Sub
RunSql "Select * from tblRegistered"
With Rs
    SubSql "Select * from tblItemList where item_id = '" & lblId.Caption & "'"
    .AddNew
    .Fields!item_id = SubRs.Fields!item_id
    .Fields!Description = SubRs.Fields!Description
    n = ValBox("Input the total number of copy of this item.", imgIcon, , , "Quantity")
    If n < 1 Then
        MsgBox "Cannot accept " & n & " input. Item registration is aborted.", vbExclamation
        Exit Sub
    End If
    .Fields!qty = n
    .Fields!date_added = Format(Date, "mm/dd/yyyy")
    .Update
End With
frmItems.ViewItems "item_id", "%"
frmItems.viewAvailable "item_id", "%"

If MsgBox("Item " & lblId.Caption & " has been successfully registered to Item list. You need to update the status for transactions." & vbNewLine & vbNewLine & _
        "Would you like to update it's status now?", vbInformation + vbYesNo) = vbYes Then
    With frmStatus
        .ExecSrch "item_id", lblId.Caption
        .cmdLoad_Click
        .Show 1
    End With
End If
ClrFlds
End Sub

Public Sub viewAvailable(RcrdFld As String, RcrdStr As String)
RunSql "Select item_id, description, qty from tblRegistered where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
With Rs
    lvwInventory.ListItems.Clear
    While Not .EOF = True
        Set x = lvwInventory.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Private Sub cmdRemove_Click()
ExecRemove lvwInventory.SelectedItem
End Sub

Public Sub ExecRemove(ItemId As String)
If NoRcrd(lvwInventory, "No record available on the list. Please Search for an item.") = True Then Exit Sub
If MsgBox("Your about to remove item " & ItemId & " from your registered list." & _
            " Removing this item may affect system records including reports." & vbNewLine & vbNewLine & _
            "Do you want to continue?", vbExclamation + vbOKCancel) = vbOK Then
    If RunSql("Delete * from tblRegistered where item_id = '" & ItemId & "'") = False Then
        frmItems.ViewItems "item_id", "%"
        frmItems.viewAvailable "item_id", "%"
        MsgBox "Item " & ItemId & " has been removed from inventory.", vbInformation
    Else
        MsgBox "Item cannot be unregistered because it is being used by other transaction.", vbExclamation
    End If
End If
ClrFlds
End Sub

Private Sub cmdView_Click()
If cmdView.Caption = "&Inventory >>" Then
    Me.Height = 9255
    cmdView.Caption = "&Inventory <<"
Else
    Me.Height = 6375
    cmdView.Caption = "&Inventory >>"
End If
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
LoadCboFld "tblItemList", "item_id, description, category, location", cboFilter
RunSql "Select item_id, description, qty from tblRegistered"
With Rs
    cboFilterInven.Clear
    For i = 0 To (.Fields.Count - 1)
        cboFilterInven.AddItem (.Fields(i).Name)
    Next i
End With
cboFilterInven.Text = "description"

SetLv lvwItems, True, True
SetLv lvwInventory, True, True
ViewUnreg "item_id", "%"
viewAvailable "item_id", "%"
End Sub

Public Sub ViewUnreg(RcrdFld As String, RcrdStr As String)
RunSql "Select item_id, description, category, location from tblItemList where " & RcrdFld & " LIKE '" & RcrdStr & "%' Order By description ASC"
With Rs
    lvwItems.ListItems.Clear
    While Not .EOF = True
        SubSql "Select * from tblRegistered where item_id = '" & .Fields!item_id & "'"
        If SubRs.EOF = True Then
            Set x = lvwItems.ListItems.Add(, , .Fields(0))
            For i = 1 To (.Fields.Count - 1)
                x.SubItems(i) = .Fields(i)
            Next i
        End If
        .MoveNext
    Wend
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 11
mdiMain.cmdWarning.Caption = Warnings & " W&arnings"
Screen.MousePointer = 0
End Sub

Public Sub lvwItems_DblClick()
If NoRcrd(lvwItems, "No record available on the list. Please Search for an item.") = True Then Exit Sub
RunSql "Select item_id, description, category, location from tblItemList where item_id = '" & lvwItems.SelectedItem & "'"
With Rs
    lblId.Caption = .Fields!item_id
    txtDescription.Text = .Fields!Description
End With
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ViewUnreg cboFilter.Text, txtSrchStr.Text
    End If
Else
    ClrFlds
End If
End Sub

Private Sub txtSrchStr_GotFocus()
If txtSrchStr = "Search" Then
    txtSrchStr.Text = Empty
    txtSrchStr.ForeColor = &H80000008
End If
End Sub

Private Sub txtSrchStr_LostFocus()
If Trim(txtSrchStr) = Empty Then
    txtSrchStr.Text = "Search"
    txtSrchStr.ForeColor = &H80000011
End If
End Sub

Private Sub txtSrchStrInven_Change()
If Right(txtSrchStrInven.Text, 1) = "'" Then
    txtSrchStrInven.Text = Empty
End If
If Trim(txtSrchStrInven.Text) <> Empty Then
    If txtSrchStrInven.Text <> "Search" Then
        viewAvailable cboFilterInven.Text, txtSrchStrInven.Text
    End If
Else
    ClrFlds
End If
End Sub

Private Sub txtSrchStrInven_GotFocus()
If txtSrchStrInven = "Search" Then
    txtSrchStrInven.Text = Empty
    txtSrchStrInven.ForeColor = &H80000008
End If
End Sub

Private Sub txtSrchStrInven_LostFocus()
If Trim(txtSrchStrInven) = Empty Then
    txtSrchStrInven.Text = "Search"
    txtSrchStrInven.ForeColor = &H80000011
End If
End Sub


