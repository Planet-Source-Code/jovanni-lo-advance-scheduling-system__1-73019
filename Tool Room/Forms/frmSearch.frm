VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   2625
   ClientLeft      =   7290
   ClientTop       =   4935
   ClientWidth     =   5670
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5670
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optTable 
      Caption         =   "Registered Items"
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
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Tag             =   "tblRegistered"
      Top             =   2160
      Width           =   1575
   End
   Begin VB.OptionButton optTable 
      Caption         =   "Item List"
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
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Tag             =   "tblItemList"
      Top             =   2160
      Value           =   -1  'True
      Width           =   975
   End
   Begin CtrlLine.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   1800
      TabIndex        =   12
      Top             =   1200
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   53
   End
   Begin VB.ComboBox cboOrder 
      Height          =   315
      ItemData        =   "frmSearch.frx":038A
      Left            =   4320
      List            =   "frmSearch.frx":0394
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2160
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
      ForeColor       =   &H8000000B&
      Height          =   300
      Left            =   2160
      TabIndex        =   0
      Text            =   "Search"
      Top             =   1080
      Width           =   3015
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   53
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      ItemData        =   "frmSearch.frx":03A3
      Left            =   240
      List            =   "frmSearch.frx":03A5
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "OK"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmSearch.frx":03A7
      Picture         =   "frmSearch.frx":0581
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdOptions 
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Options >>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmSearch.frx":08D5
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   0
      Picture         =   "frmSearch.frx":0AAF
      Top             =   1530
      Width           =   480
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Options to specify search"
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
      TabIndex        =   13
      Top             =   1680
      Width           =   1830
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "|"
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
      Left            =   3120
      TabIndex        =   11
      Top             =   2160
      Width           =   105
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Order By:"
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
      Left            =   3345
      TabIndex        =   10
      Top             =   2160
      Width           =   780
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search for a specific item"
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
      TabIndex        =   8
      Top             =   480
      Width           =   1545
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH"
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
      TabIndex        =   7
      Top             =   120
      Width           =   765
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmSearch.frx":1379
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5040
      Picture         =   "frmSearch.frx":1FBD
      Top             =   960
      Width           =   480
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Table As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOptions_Click()
If cmdOptions.Caption = "&Options >>" Then
        Me.Height = 3090
        cmdOptions.Caption = "&Options <<"
        lblLabel.Caption = "Select an item search below"
    Else
        Me.Height = 2535
        cmdOptions.Caption = "&Options >>"
        lblLabel.Caption = "Options to specify search"
    End If
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
Me.Height = 2535
cboOrder.ListIndex = 0
optTable_Click 0
LoadCboFld Table, "*", cboFilter
cboFilter.Text = ReadINI("Preferences", "Filter")
frmItems.cboMonth.ListIndex = 0
End Sub

Private Sub optTable_Click(Index As Integer)
Select Case Index
    Case 0
        Table = "tblItemList"
    Case 1
        Table = "tblRegistered"
End Select
End Sub

Private Sub txtSrchStr_Change()
For i = 1 To Len(txtSrchStr.Text)
    If Right(Left(txtSrchStr.Text, i), 1) = "'" Then
        txtSrchStr.Text = Empty
        Exit Sub
    End If
Next i
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        If Table = "tblItemList" Then
            frmItems.ViewItems cboFilter.Text, "%" & txtSrchStr.Text, cboOrder.Text
        Else
            frmItems.viewAvailable cboFilter.Text, "%" & txtSrchStr.Text, cboOrder.Text
        End If
    End If
Else
    If Table = "tblItemList" Then
        frmItems.ViewItems cboFilter.Text, "%", cboOrder.Text
    Else
        frmItems.viewAvailable cboFilter.Text, "%", cboOrder.Text
    End If
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
    txtSrchStr.ForeColor = &H8000000B
End If
End Sub
