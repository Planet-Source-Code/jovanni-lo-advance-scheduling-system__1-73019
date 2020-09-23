VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTransView 
   Caption         =   "Transaction"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   Icon            =   "frmTransView.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   9600
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrMenu 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Print"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Borrow"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reserve"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Manage"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Close"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ComboBox cboRecords 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmTransView.frx":038A
      Left            =   4440
      List            =   "frmTransView.frx":038C
      TabIndex        =   5
      Top             =   840
      Width           =   2055
   End
   Begin VB.ComboBox cboFilter 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmTransView.frx":038E
      Left            =   840
      List            =   "frmTransView.frx":03A1
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   840
      Width           =   2055
   End
   Begin ComctlLib.ListView lvwRecords 
      Height          =   1095
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1931
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
      NumItems        =   0
   End
   Begin ComctlLib.TabStrip tabMenu 
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Borrowed Items"
            Key             =   "tblBorrow"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reserved Items"
            Key             =   "tblReserve"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Returned Items"
            Key             =   "tblReturn"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Canceled Reservation"
            Key             =   "tblCancel"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Records:"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label10 
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
      Index           =   1
      Left            =   3120
      TabIndex        =   6
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Filter:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   480
   End
   Begin ComctlLib.ImageList lstMenu 
      Left            =   8640
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransView.frx":03EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransView.frx":073C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransView.frx":0A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransView.frx":0DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransView.frx":1132
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmTransView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Field As String, Table As String
Private Sub cboFilter_Click()
Select Case cboFilter.ListIndex
    Case 0
        Field = "trans_no"
        Table = "tblTransactions"
    Case 1
        Field = "client_no"
        Table = "tblClientProfile"
    Case 2
        Field = "status"
        Table = "tblItemStatus"
    Case 3
        Field = "item_id"
        Table = "tblRegistered"
    Case 4
        LoadTrans "Select * from " & tabMenu.SelectedItem.Key
        cboRecords.Clear
        Exit Sub
End Select
LoadCbo Table, cboRecords, Field
End Sub

Private Sub cboRecords_Change()
cboRecords_Click
End Sub

Private Sub cboRecords_Click()
'LoadTrans "SELECT distinctrow tbl.* " & _
'        "FROM " & tabMenu.SelectedItem.Key & " as tbl INNER JOIN " & Table & " as [sub] " & _
'        "ON tbl." & Field & " = sub." & Field & " " & _
'        "WHERE tbl." & Field & " LIKE '%" & cboRecords.Text & "%'"
'End Sub
LoadTrans "SELECT * from " & tabMenu.SelectedItem.Key & " where " & Field & " LIKE '%" & cboRecords.Text & "%'"
End Sub


Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
SetLv lvwRecords, True, True
mdiMain.tbrMenu.Buttons(2).Value = tbrPressed
mdiMain.sbrDetails.Panels(5).Text = "Click Search for more search options    F3 - Lock the computer"
With tbrMenu
    .ImageList = lstMenu
    For i = 1 To lstMenu.ListImages.Count
        .Buttons(i).Image = i
        n = n + 2
    Next i
End With
cboFilter.ListIndex = 0
tabMenu_Click
End Sub

Private Sub Form_Resize()
Layout
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiMain.tbrMenu.Buttons(2).Value = tbrUnpressed
mdiMain.sbrDetails.Panels(5).Text = Empty
End Sub

Public Sub tabMenu_Click()
cboRecords_Click
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As ComctlLib.Button)
Screen.MousePointer = 0
ExecButtons Button.Index
End Sub

Public Sub ExecButtons(Index As Integer)
Screen.MousePointer = 11
Select Case Index
    Case 2
        frmTransactions.tabTrans.Tabs(1).Selected = True
        frmTransactions.Show 1
    Case 3
        frmTransactions.tabTrans.Tabs(2).Selected = True
        frmTransactions.Show 1
    Case 4
        frmTransactions.tabTrans.Tabs(5).Selected = True
        frmTransactions.Show 1
    Case 5
        Unload Me
End Select
Screen.MousePointer = 0
End Sub

Private Sub Layout()
On Error Resume Next
tabMenu.Width = ScaleWidth - (tabMenu.Left * 1.2)
tabMenu.Height = ScaleHeight - (tabMenu.Top * 1.05)

lvwRecords.Height = tabMenu.Height - 410
lvwRecords.Width = tabMenu.Width - 110
lvwRecords.Top = tabMenu.Top + 350
lvwRecords.Left = tabMenu.Left + 40
lvwRecords.Move _
        lvwRecords.Left, _
        lvwRecords.Top, _
        lvwRecords.Width, _
        lvwRecords.Height
End Sub

Public Sub LoadTrans(Statement As String)
RunSql Statement
With Rs
    n = 0
    lvwRecords.ColumnHeaders.Clear
    For i = 1 To (.Fields.Count)
        lvwRecords.ColumnHeaders.Add
        If n < .Fields.Count Then
            lvwRecords.ColumnHeaders(i).Text = .Fields(n).Name
        End If
        n = n + 1
    Next i

    lvwRecords.ListItems.Clear
    While Not .EOF = True
        Set x = lvwRecords.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub
