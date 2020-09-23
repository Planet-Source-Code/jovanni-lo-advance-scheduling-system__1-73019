VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmStatus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7350
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Date Added"
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
      Left            =   4800
      TabIndex        =   13
      Top             =   960
      Width           =   2415
      Begin VB.Label lblDateAdded 
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
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   195
      End
   End
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
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   7095
      Begin CtrlLine.ctrlLiner ctrlLiner3 
         Height          =   30
         Left            =   4080
         TabIndex        =   27
         Top             =   1440
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   53
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   6120
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cboCondition 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblStatNo 
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
         Left            =   6120
         TabIndex        =   30
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Status #:"
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
         Index           =   3
         Left            =   4800
         TabIndex        =   29
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quantity:"
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
         Index           =   1
         Left            =   4800
         TabIndex        =   26
         Top             =   1320
         Width           =   780
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
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Condition:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   990
      End
      Begin VB.Label lblDescription 
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
         Left            =   1680
         TabIndex        =   12
         Top             =   840
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
         Left            =   360
         TabIndex        =   11
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblQty 
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
         Left            =   6120
         TabIndex        =   10
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Remaining:"
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
         Index           =   2
         Left            =   4800
         TabIndex        =   9
         Top             =   840
         Width           =   945
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
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   705
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
      TabIndex        =   3
      Top             =   960
      Width           =   4575
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmStatus.frx":038A
         Left            =   240
         List            =   "frmStatus.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
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
         Left            =   1800
         TabIndex        =   5
         Text            =   "Search"
         Top             =   240
         Width           =   2295
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   1560
         TabIndex        =   4
         Top             =   360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4080
         Picture         =   "frmStatus.frx":038E
         Top             =   120
         Width           =   480
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lvwAvailable 
      Height          =   1575
      Left            =   120
      TabIndex        =   18
      Top             =   6120
      Width           =   7095
      _ExtentX        =   12515
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
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Count"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date Added"
         Object.Width           =   2540
      EndProperty
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStatus.frx":0C58
      Picture         =   "frmStatus.frx":0E32
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClear 
      Height          =   375
      Left            =   5880
      TabIndex        =   22
      Top             =   7800
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
      MouseIcon       =   "frmStatus.frx":1186
      Picture         =   "frmStatus.frx":1360
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLoad 
      Height          =   375
      Left            =   4440
      TabIndex        =   23
      Top             =   7800
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
      MouseIcon       =   "frmStatus.frx":16B4
      Picture         =   "frmStatus.frx":188E
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdView 
      Height          =   375
      Left            =   5880
      TabIndex        =   24
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&View <<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStatus.frx":1BE2
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ComctlLib.ListView lvwItemStat 
      Height          =   1215
      Left            =   120
      TabIndex        =   28
      Top             =   3840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2143
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
         Text            =   "#"
         Object.Width           =   265
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remarks"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Qty"
         Object.Width           =   529
      EndProperty
   End
   Begin ctrlButton.ThemedButton cmdNew 
      Height          =   375
      Left            =   1560
      TabIndex        =   31
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&New"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStatus.frx":1DBC
      Picture         =   "frmStatus.frx":1F96
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDelete 
      Height          =   375
      Left            =   3000
      TabIndex        =   32
      Top             =   5160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmStatus.frx":22EA
      Picture         =   "frmStatus.frx":24C4
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   4440
      TabIndex        =   33
      Top             =   5160
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
      MouseIcon       =   "frmStatus.frx":2818
      Picture         =   "frmStatus.frx":29F2
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double click an item from the list to load details."
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
      TabIndex        =   19
      Top             =   7920
      Width           =   3405
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   0
      Picture         =   "frmStatus.frx":2D46
      Top             =   7800
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List of registered items on Inventory"
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
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   5760
      Width           =   3120
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update the product status"
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
      TabIndex        =   1
      Top             =   480
      Width           =   1590
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS"
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
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmStatus.frx":3610
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   -840
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCondition_Click()
cmdSave.Caption = "&Save"
lblQty.Caption = GenQty(lblId.Caption)
End Sub

Private Sub cmdClear_Click()
ClrFlds
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If NoRcrd(lvwItemStat, "No record available on the list.") = True Then Exit Sub
If MsgBox("Are you sure you want to remove this status?", vbExclamation + vbYesNo) = vbYes Then
    RunSql "SELECT bar.*, stat.* " & _
            "FROM tblItemStatus as stat INNER JOIN tblBorrow as bar ON stat.item_id = bar.item_id " & _
            "WHERE stat.record_no = " & lvwItemStat.SelectedItem
    If Rs.EOF = False Then
        MsgBox "Cannot delete this record. It is used by other transaction.", vbExclamation
        Exit Sub
    End If
    RunSql "Delete * from tblItemStatus where record_no = " & lvwItemStat.SelectedItem
    MsgBox "Item " & lblId.Caption & "'s condition has been updated.", vbInformation
End If
cmdLoad_Click
End Sub

Public Sub cmdLoad_Click()
If NoRcrd(lvwAvailable, "No record available on your Item list. Please search for an item.") = True Then Exit Sub
RunSql "SELECT item_id, description, qty, date_added FROM tblRegistered where item_id = '" & lvwAvailable.SelectedItem & "'"
With Rs
    lblId.Caption = .Fields!item_id
    lblDescription.Caption = .Fields!Description
    lblDateAdded.Caption = Format(.Fields!date_added, "mm/dd/yyyy")
End With
cmdNew_Click
End Sub

Private Function GenQty(ID As String) As Long
RunSql "Select * from tblItemStatus where item_id = '" & ID & "'"
With Rs
    While Not .EOF = True
        SubSql "Select * from tblStatus where description = '" & .Fields!Status & "'"
        If SubRs.EOF = False Then
            GenQty = GenQty + .Fields!qty
        End If
        .MoveNext
    Wend
End With
RunSql "Select qty from tblRegistered where item_id = '" & ID & "'"
With Rs
    GenQty = .Fields!qty - GenQty
End With
End Function

Private Sub cmdNew_Click()
cboCondition.ListIndex = 0
txtQty.Text = Empty
cmdSave.Caption = "&Save"
lblStatNo.Caption = RcrdId("tblItemStatus", , "record_no")
lblQty.Caption = GenQty(lblId.Caption)
ViewItemStat
End Sub

Private Sub cmdSave_Click()
If lblId.Caption = "---" Then
    MsgBox "No espicified item. Click Load to load an item from your inventory list", vbExclamation
    Exit Sub
End If

If CboEmp(cboCondition) = True Then Exit Sub
If TxtEmp(txtQty) = True Then Exit Sub

RunSql "Select * from tblStatus where description = '" & cboCondition.Text & "'"
If Rs.Fields!System = True Then
    MsgBox "This is a system status. You are not allowed to do this process.", vbCritical
    Exit Sub
End If

If Val(txtQty.Text) > Val(lblQty.Caption) Then
    MsgBox "Quantity must be equal or below the total quantity remaining of the item.", vbExclamation
    txtQty.SetFocus
    SelAll txtQty
    Exit Sub
End If

RunSql "Select * from tblItemStatus where status = '" & cboCondition.Text & "' and item_id = '" & lblId.Caption & "'"
With Rs
    If cmdSave.Caption = "&Save" Then
        If .EOF = False Then
            .Fields!qty = .Fields!qty + Val(txtQty.Text)
        Else
            .AddNew
            .Fields!record_no = RcrdId("tblItemStatus", , "record_no")
            .Fields!qty = Val(txtQty.Text)
        End If
    Else
        .Fields!qty = Val(txtQty.Text)
    End If
    .Fields!item_id = lblId.Caption
    .Fields!Status = cboCondition.Text
    .Update
End With
cmdLoad_Click
MsgBox "Item " & lblId.Caption & "'s condition has been updated.", vbInformation
ExecSrch "item_id", "%"
frmItems.viewAvailable "item_id", "%"
End Sub

Private Sub cmdView_Click()
If cmdView.Caption = "&View <<" Then
    Me.Height = 6075
    cmdView.Caption = "&View >>"
Else
    Me.Height = 8805
    cmdView.Caption = "&View <<"
End If
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Public Sub ExecSrch(RcrdFld As String, RcrdStr As String)
RunSql "Select item_id, description, qty, date_added from tblRegistered where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
With Rs
    lvwAvailable.ListItems.Clear
    While Not .EOF = True
        Set x = lvwAvailable.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Public Sub ViewItemStat()
RunSql "Select ItemStat.record_no, ItemStat.status, Stat.remarks, ItemStat.qty " & _
        "from tblItemStatus as ItemStat INNER JOIN tblStatus as Stat ON " & _
        "ItemStat.status = stat.description where ItemStat.item_id = '" & lblId.Caption & "'"
With Rs
    lvwItemStat.ListItems.Clear
    While Not .EOF = True
        Set x = lvwItemStat.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Private Sub Form_Load()
SetLv lvwItemStat, True, True
SetLv lvwAvailable, True, True
LoadCboFld "tblRegistered", "item_id, description, qty, date_added", cboFilter
LoadCbo "tblStatus", cboCondition, "description", "Select"
ExecSrch "item_id", "%"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 11
mdiMain.cmdWarning.Caption = Warnings & " W&arnings"
Screen.MousePointer = 0
End Sub

Private Sub lvwAvailable_DblClick()
cmdLoad_Click
End Sub

Private Sub lvwItemStat_DblClick()
If NoRcrd(lvwItemStat) = True Then Exit Sub
RunSql "Select * from tblItemStatus where record_no = " & lvwItemStat.SelectedItem
With Rs
    lblStatNo.Caption = .Fields!record_no
    cboCondition.Text = .Fields!Status
    lblQty.Caption = lblQty + .Fields!qty
    txtQty.Text = .Fields!qty
    cmdSave.Caption = "&Update"
End With
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ExecSrch cboFilter.Text, txtSrchStr.Text
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

Private Sub ClrFlds()
lblId.Caption = "---"
lblDescription.Caption = "---"
cboCondition.ListIndex = 0
lblQty.Caption = "---"
lblDateAdded.Caption = "---"
ExecSrch "item_id", "%"
cmdNew_Click
End Sub
