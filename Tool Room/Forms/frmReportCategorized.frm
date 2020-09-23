VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmRptCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reports"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   Icon            =   "frmReportCategorized.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Query Output"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   4200
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
      Begin VB.TextBox txtOutput 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Categories"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   3975
      Begin ComctlLib.TreeView tvwCategory 
         Height          =   2415
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   4260
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "lstIcons"
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
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Query Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   6
      Top             =   960
      Width           =   2175
      Begin VB.ComboBox cboYear 
         Height          =   315
         ItemData        =   "frmReportCategorized.frx":038A
         Left            =   240
         List            =   "frmReportCategorized.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdDisplay 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Display"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmReportCategorized.frx":038E
      Picture         =   "frmReportCategorized.frx":0568
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   4200
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
      MouseIcon       =   "frmReportCategorized.frx":08BC
      Picture         =   "frmReportCategorized.frx":0A96
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ComctlLib.ImageList lstIcons 
      Left            =   3120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   52
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":0DEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":113C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":148E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":17E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":1B32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":1E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":21D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":2528
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":287A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":2BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":2F1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":3270
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":35C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":3914
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":3C66
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":3FB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":430A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":465C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":49AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":4D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":5052
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":53A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":56F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":5A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":5D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":60EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":643E
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":6790
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":6AE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":6E34
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":7186
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":74D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":782A
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":7B7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":7ECE
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":8220
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":8572
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":88C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":8C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":8F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":92BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":960C
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":995E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":9CB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":A002
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":A354
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":A6A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":A9F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":AD4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":B09C
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":B3EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmReportCategorized.frx":B740
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a category and click Display"
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
      TabIndex        =   9
      Top             =   4335
      Width           =   2460
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmReportCategorized.frx":BA92
      Top             =   4200
      Width           =   480
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set system report by categories"
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
      TabIndex        =   5
      Top             =   480
      Width           =   1980
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CATEGORIZED"
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
      TabIndex        =   4
      Top             =   120
      Width           =   1290
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmReportCategorized.frx":C35C
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "frmRptCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDisplay_Click()
Dim rptReport As Variant
Dim dYear As String
Dim Parameter As String

dYear = cboYear.Text
Select Case tvwCategory.SelectedItem.Key
    Case "summary"
        CloseRs dtaGroups.rscmdSummary
        Set rptReport = rptSummary
        dtaGroups.cmdSummary dYear
    Case "category"
        CloseRs dtaGroups.rscmdCategory
        Set rptReport = rptCategory
        dtaGroups.cmdCategory dYear
    Case "location"
        CloseRs dtaGroups.rscmdLocation
        Set rptReport = rptLocation
        dtaGroups.cmdLocation dYear
    Case "status"
        CloseRs dtaGroups.rscmdStatus
        Set rptReport = rptStatus
        dtaGroups.cmdStatus dYear
    Case "user_trans"
        StrBox "Select a client to generate.", _
                imgIcon, , , "clients", 3, "tblClientProfile", _
                "fname", "%"
    Case Else
        MsgBox "On progress", vbExclamation
        Exit Sub
End Select
Unload Me
rptReport.Show
End Sub

Private Sub Form_Load()
mdiMain.tbrMenu.Buttons(3).Value = tbrPressed
For i = Format(Date, "yyyy") To 2000 Step -1
    cboYear.AddItem i
Next i
cboYear.Text = ReadINI("Preferences", "Year")
With tvwCategory.Nodes
    .Add , , "items", "Registered Items", 16
    .Add "items", tvwChild, "summary", "Year Summary", 1
    .Add "items", tvwChild, "category", "By Item Categories", 1
    .Add "items", tvwChild, "location", "Locations", 1
    .Add "items", tvwChild, "status", "Item Status", 1
    
    .Add , , "trans", "Transactions", 44
    .Add "trans", tvwChild, "user_trans", "User Transactions", 1
    .Add "trans", tvwChild, "trans_no", "Transaction Numbers", 1
    
    .Add "trans", tvwChild, "sched_trans", "Scheduling Transactions", 24
    .Add "sched_trans", tvwChild, "borrow", "Borrowed Items", 1
    .Add "sched_trans", tvwChild, "reserve", "Reserved Items", 1
    .Add "sched_trans", tvwChild, "return", "Returned Items", 1
    .Add "sched_trans", tvwChild, "cancel", "Canceled Reservations", 1
    
    .Item(1).Expanded = ReadINI("Report Nodes", "1")
    .Item(6).Expanded = ReadINI("Report Nodes", "6")
    .Item(9).Expanded = ReadINI("Report Nodes", "9")
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiMain.tbrMenu.Buttons(3).Value = tbrUnpressed
WriteINI "Report Nodes", "1", tvwCategory.Nodes(1).Expanded
WriteINI "Report Nodes", "6", tvwCategory.Nodes(6).Expanded
WriteINI "Report Nodes", "9", tvwCategory.Nodes(9).Expanded
End Sub

Private Sub tvwCategory_DblClick()
cmdDisplay_Click
End Sub

Private Sub tvwCategory_NodeClick(ByVal Node As ComctlLib.Node)
Select Case tvwCategory.SelectedItem.Index
    Case 2
        txtOutput.Text = "Display summary of registered items for the year " & cboYear.Text & "."
    Case 3
        txtOutput.Text = "Display item details by Categories. For those items on the year " & cboYear.Text & "."
    Case 4
        txtOutput.Text = "Display item details by Location. For those items on the year " & cboYear.Text & "."
    Case 5
        txtOutput.Text = "Display item details by Status/Condition. For those items on the year " & cboYear.Text & "."
    Case Else
        txtOutput.Text = Empty
End Select
End Sub
