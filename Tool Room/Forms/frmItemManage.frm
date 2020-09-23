VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmItemManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8430
   Icon            =   "frmItemManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Item Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   21
      Top             =   2520
      Width           =   8175
      Begin VB.CheckBox chkAutoCategory 
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   4
         Top             =   960
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chkAutoLocation 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   5
         Top             =   960
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.ComboBox cboLocation 
         Height          =   315
         ItemData        =   "frmItemManage.frx":038A
         Left            =   5760
         List            =   "frmItemManage.frx":038C
         TabIndex        =   3
         Text            =   "cboLocation"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtRemarks 
         Height          =   615
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1680
         Width           =   6255
      End
      Begin VB.ComboBox cboCategory 
         Height          =   315
         ItemData        =   "frmItemManage.frx":038E
         Left            =   1680
         List            =   "frmItemManage.frx":0390
         TabIndex        =   0
         Text            =   "cboCategory"
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtDescription 
         Height          =   495
         Left            =   1680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Auto Save:"
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
         Left            =   4560
         TabIndex        =   26
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Location:"
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
         Left            =   4440
         TabIndex        =   25
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Remarks:"
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
         TabIndex        =   24
         Top             =   1680
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "* Category:"
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
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "* Description:"
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
         TabIndex        =   22
         Top             =   960
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Total Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3600
      TabIndex        =   13
      Top             =   1800
      Width           =   1455
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
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
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   105
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
      TabIndex        =   9
      Top             =   960
      Width           =   4935
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmItemManage.frx":0392
         Left            =   240
         List            =   "frmItemManage.frx":0394
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   1575
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
         Left            =   2160
         TabIndex        =   11
         Text            =   "Search"
         Top             =   240
         Width           =   2175
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   4320
         Picture         =   "frmItemManage.frx":0396
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Item Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3375
      Begin VB.Label lblId 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0000"
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
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   420
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   15
      Top             =   840
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   53
   End
   Begin MSComDlg.CommonDialog dlgPic 
      Left            =   7920
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      Caption         =   "Item Image"
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
      Left            =   5160
      TabIndex        =   16
      Top             =   960
      Width           =   3135
      Begin ctrlButton.ThemedButton cmdRemove 
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
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
         MouseIcon       =   "frmItemManage.frx":0C60
         Picture         =   "frmItemManage.frx":0E3A
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdBrowse 
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Browse"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmItemManage.frx":118E
         Picture         =   "frmItemManage.frx":1368
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000011&
         FillColor       =   &H80000004&
         Height          =   1095
         Left            =   1680
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image imgProfile 
         Height          =   1095
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "NO IMAGE"
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
         Left            =   1920
         TabIndex        =   17
         Top             =   720
         Width           =   825
      End
   End
   Begin ComctlLib.ListView lvwItems 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   5880
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3413
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   5292
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
   Begin ctrlButton.ThemedButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   28
      Top             =   5280
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
      MouseIcon       =   "frmItemManage.frx":16BC
      Picture         =   "frmItemManage.frx":1896
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdNew 
      Height          =   375
      Left            =   4080
      TabIndex        =   29
      Top             =   5280
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
      MouseIcon       =   "frmItemManage.frx":1BEA
      Picture         =   "frmItemManage.frx":1DC4
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   5520
      TabIndex        =   30
      Top             =   5280
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
      MouseIcon       =   "frmItemManage.frx":2118
      Picture         =   "frmItemManage.frx":22F2
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdOptions 
      Height          =   375
      Left            =   6960
      TabIndex        =   31
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Options <<"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmItemManage.frx":2646
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDelete 
      Height          =   375
      Left            =   6960
      TabIndex        =   32
      Top             =   7920
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
      MouseIcon       =   "frmItemManage.frx":2820
      Picture         =   "frmItemManage.frx":29FA
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLoad 
      Height          =   375
      Left            =   4080
      TabIndex        =   33
      Top             =   7920
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
      MouseIcon       =   "frmItemManage.frx":2D4E
      Picture         =   "frmItemManage.frx":2F28
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdRegister 
      Height          =   375
      Left            =   5520
      TabIndex        =   34
      Top             =   7920
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
      MouseIcon       =   "frmItemManage.frx":327C
      Picture         =   "frmItemManage.frx":3456
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   0
      Picture         =   "frmItemManage.frx":37AA
      Top             =   7920
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select an Item and click Load to manage"
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
      TabIndex        =   27
      Top             =   8040
      Width           =   2880
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MANAGE ITEMS"
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
      Index           =   0
      Left            =   720
      TabIndex        =   20
      Top             =   120
      Width           =   1440
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmItemManage.frx":4074
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add, Update items on database."
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
      Index           =   1
      Left            =   720
      TabIndex        =   19
      Top             =   480
      Width           =   2010
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Indecates required field"
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
      TabIndex        =   18
      Top             =   5400
      Width           =   1845
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmItemManage.frx":4CB8
      Top             =   5280
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      FillColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmItemManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
On Error GoTo InvldPic
dlgPic.DialogTitle = "Load Item Image"
dlgPic.InitDir = "My Documents"
dlgPic.Filter = "Jepeg Image (*.jpg;*.jpeg)|*.jpg;*.jpeg|Bitmap Image (*.bmp)|*.bmp|All Files (*.*)|*.*"
dlgPic.ShowOpen
If dlgPic.FileName = "" Then
    Exit Sub
End If
imgProfile.Picture = LoadPicture(dlgPic.FileName)
ImgName = dlgPic.FileTitle
ImgSrc = dlgPic.FileName
Exit Sub
InvldPic:
    MsgBox "It is not a valid picture", vbExclamation
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Your about to delete the item, " & lvwItems.SelectedItem.SubItems(1) & ". Deleting this item may affect system reports." & vbNewLine & vbNewLine & _
    "Are you sure?", vbExclamation + vbYesNo) = vbYes Then
    RunSql "Delete * from tblItemList where item_id = '" & lvwItems.SelectedItem.SubItems(1) & "'"
    ClrFlds
    MsgBox "Successfully deleted a record.", vbInformation
    frmItems.ViewItems "item_id", "%"
    frmItems.viewAvailable "item_id", "%"
End If
End Sub

Private Sub cmdLoad_Click()
If NoRcrd(lvwItems) = True Then Exit Sub
ExecSrch "item_id", lvwItems.SelectedItem.SubItems(1)
End Sub

Private Sub cmdNew_Click()
ClrFlds
cboCategory.SetFocus
End Sub

Private Sub cmdRegister_Click()
frmRegister.ViewUnreg "item_id", lvwItems.SelectedItem.SubItems(1)
frmRegister.lvwItems_DblClick
frmRegister.Show 1
End Sub

Private Sub cmdRemove_Click()
imgProfile.Picture = LoadPicture()
ImgName = Empty
End Sub

Private Sub cmdSave_Click()
If TxtEmp(cboCategory) = True Then Exit Sub
If TxtEmp(txtDescription) = True Then Exit Sub
If TxtEmp(cboLocation) = True Then Exit Sub

If chkAutoCategory.Value = 1 Then
    RunSql "Select * from tblCategories where description = '" & Trim(cboCategory.Text) & "'"
    With Rs
        If .EOF = True Then
            .AddNew
            .Fields!Description = Trim(cboCategory.Text)
            .Fields!remarks = ""
            .Update
        End If
    End With
End If

If chkAutoLocation.Value = 1 Then
    RunSql "Select * from tblLocations where description = '" & Trim(cboLocation.Text) & "'"
    With Rs
        If .EOF = True Then
            .AddNew
            .Fields!Description = Trim(cboLocation.Text)
            .Fields!remarks = ""
            .Update
        End If
    End With
End If

RunSql "Select * from tblItemList where item_id = '" & lblId.Caption & "'"
With Rs
    If cmdSave.Caption = "&Save" Then
        .AddNew
        MSG = "Added new item on database"
        If ImgName <> Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Items\" & ImgName
        End If
    Else
        If ImgName <> Empty And .Fields!image_name = Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Items\" & ImgName
        End If
        If .Fields!image_name <> Empty And .Fields!image_name <> ImgName Then
            Kill App.Path & "\Images\Items\" & .Fields!image_name
            If ImgName <> Empty Then
                FileCopy ImgSrc, App.Path & "\Images\Items\" & ImgName
            End If
        End If
        MSG = "Item " & lblId.Caption & " has been updated"
    End If
    .Fields!item_id = lblId.Caption
    .Fields!Description = txtDescription.Text
    .Fields!category = cboCategory.Text
    .Fields!location = cboLocation.Text
    .Fields!remarks = txtRemarks.Text
    .Fields!image_name = ImgName
    .Fields!reg_date = Format(Date, "mm/dd/yyyy")
    .Update
End With
frmItems.ViewItems "item_id", "%"
frmItems.viewAvailable "item_id", "%"
MsgBox MSG, vbInformation
'-------------
RunSql "Select * from tblRegistered where item_id = '" & lblId.Caption & "'"
With Rs
    If .EOF = True Then
        x = MsgBox("This item is not yet registered on your Item List. This will not be included on scheduling transactions." & vbNewLine & vbNewLine & _
                "Would you like to register it now?", vbExclamation + vbYesNo)
        If x = vbYes Then
            frmRegister.ViewUnreg "item_id", lblId.Caption
            frmRegister.lvwItems_DblClick
            frmRegister.Show 1
        End If
    End If
End With
ClrFlds
ViewItems "item_id", "%"
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub SetUp()
LoadCbo "tblCategories", cboCategory, "description"
LoadCbo "tblLocations", cboLocation, "description"
RunSql "Select * from tblItemList"
lblCount.Caption = Rs.RecordCount

ImgName = Empty
ImgSrc = Empty
ClrFlds
End Sub

Private Sub Form_Load()
SetLv lvwItems, True, True
LoadCboFld "tblItemList", "*", cboFilter
SetUp
End Sub

Private Sub ClrFlds()
lblId.Caption = RcrdId("tblItemList", Format(Date, "yymm-dd"), "item_id")
imgProfile.Picture = LoadPicture()
cboCategory.Text = Empty
txtDescription.Text = Empty
txtRemarks.Text = Empty
cboLocation.Text = Empty
cmdSave.Caption = "&Save"
ViewItems "item_id", "%"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 11
mdiMain.cmdWarning.Caption = Warnings & " W&arnings"
Screen.MousePointer = 0
End Sub

Private Sub lvwItems_DblClick()
cmdLoad_Click
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ExecSrch cboFilter.Text, txtSrchStr.Text
        ViewItems cboFilter.Text, txtSrchStr.Text
        frmItems.ViewItems cboFilter.Text, txtSrchStr.Text
    End If
Else
    ClrFlds
    frmItems.ViewItems "item_id", "%"
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

Public Sub ExecSrch(SrchFld As String, SrchStr As String)
RunSql "Select * from tblItemList where " & SrchFld & " LIKE '" & SrchStr & "%'"
With Rs
    If .EOF = False Then
        cboCategory.Text = .Fields!category
        txtRemarks.Text = .Fields!remarks
        cboLocation.Text = .Fields!location
        lblId.Caption = .Fields!item_id
        If .Fields!image_name <> Empty Then
            imgProfile.Picture = LoadPicture(App.Path & "\Images\Items\" & .Fields!image_name)
        Else
            imgProfile.Picture = LoadPicture()
        End If
        ImgName = .Fields!image_name
        txtDescription.Text = .Fields!Description
        cmdSave.Caption = "&Update"
    Else
        ClrFlds
    End If
End With
End Sub

Public Sub ViewItems(RcrdFld As String, RcrdStr As String)
RunSql "Select * from tblItemList where " & RcrdFld & " LIKE '" & RcrdStr & "%' Order By description ASC"
With Rs
    lvwItems.ListItems.Clear
    While Not .EOF = True
        Set x = lvwItems.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 5)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Private Sub XPButton2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub
