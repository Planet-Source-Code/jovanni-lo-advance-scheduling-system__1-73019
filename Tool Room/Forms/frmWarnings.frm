VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmWarnings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   Icon            =   "frmWarnings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freDescription 
      Caption         =   "Message Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3720
      TabIndex        =   7
      Top             =   960
      Width           =   3975
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame freList 
      Caption         =   "List of detections"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3495
      Begin ComctlLib.ListView lvwList 
         Height          =   1815
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3201
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
            Text            =   "#"
            Object.Width           =   265
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Reference"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Title"
            Object.Width           =   2540
         EndProperty
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
   Begin ctrlButton.ThemedButton cmdClose 
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
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
      MouseIcon       =   "frmWarnings.frx":038A
      Picture         =   "frmWarnings.frx":0564
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdHelp 
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "&Help"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmWarnings.frx":08B8
      Picture         =   "frmWarnings.frx":0A92
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Image imgNotifications 
      Height          =   480
      Left            =   6000
      Picture         =   "frmWarnings.frx":0DE6
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Left            =   5160
      Picture         =   "frmWarnings.frx":1A2A
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmWarnings.frx":266E
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Double click list to manage or click to view."
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
      TabIndex        =   6
      Top             =   3600
      Width           =   3030
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblHeader 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HEADER"
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
      Width           =   765
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prompt"
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
      TabIndex        =   3
      Top             =   480
      Width           =   450
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmWarnings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdHelp_Click()
PathToDoc = App.Path & "\help.chm"
ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
If NoRcrd(lvwList) = True Then Exit Sub
lvwList.ListItems(1).Selected = True
lvwList_Click
End Sub

Private Sub Form_Load()
SetLv lvwList, True, True
Select Case DetectionType
    Case 1
        SetUp "Warnings", "Manage system warnings", imgWarning, "tblDetections"
    Case 2
        SetUp "Notifications", "View system notifications", imgNotifications, "tblDetections"
End Select
End Sub

Public Sub SetUp(Header As String, Prompt As String, Image As Image, Table As String)
lblHeader.Caption = StrConv(Header, vbUpperCase)
lblPrompt.Caption = Prompt
imgIcon.Picture = Image.Picture
RunSql "Select * from " & Table
With Rs
    lvwList.ListItems.Clear
    While Not .EOF = True
        Set x = lvwList.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 2)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub

Private Sub lvwList_Click()
If NoRcrd(lvwList) = True Then Exit Sub
RunSql "Select * from tblDetections where record_no = " & lvwList.SelectedItem
With Rs
    txtDescription.Text = .Fields!Description
End With
End Sub

Private Sub lvwList_DblClick()
If NoRcrd(lvwList) = True Then Exit Sub
Select Case lvwList.SelectedItem.SubItems(2)
    Case "Scheduled Task"
        frmScheduler.ViewRcrds "title", lvwList.SelectedItem.SubItems(1)
        frmScheduler.cmdLoad_Click
        frmScheduler.Show 1
    Case "Unregistered"
        frmRegister.ViewUnreg "item_id", lvwList.SelectedItem.SubItems(1)
        frmRegister.lvwItems_DblClick
        frmRegister.Show 1
    Case "Add Items"
        frmItemManage.Show 1
    Case "Item Status"
        frmStatus.ExecSrch "item_id", lvwList.SelectedItem.SubItems(1)
        frmStatus.cmdLoad_Click
        frmStatus.Show 1
    Case "Unreturned Items", "Return Items"
        frmTransactions.ExecSrch "client_no", lvwList.SelectedItem.SubItems(1)
        frmTransactions.tabTrans.Tabs(3).Selected = True
        frmTransactions.Show 1
    Case "Expired Reservation", "Reservations"
        frmTransactions.ExecSrch "client_no", lvwList.SelectedItem.SubItems(1)
        frmTransactions.tabTrans.Tabs(4).Selected = True
        frmTransactions.Show 1
End Select
If lblHeader.Caption = "WARNINGS" Then
    DetectionType = 1
Else
    DetectionType = 2
End If
Form_Load
End Sub

Private Sub lvwList_KeyDown(KeyCode As Integer, Shift As Integer)
lvwList_Click
End Sub

Private Sub lvwList_KeyUp(KeyCode As Integer, Shift As Integer)
lvwList_Click
End Sub
