VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmItems 
   Caption         =   "Items"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9345
   Icon            =   "frmItems.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9345
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrMenu 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Add"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Edit"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Delete"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Search"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Register"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Status"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Close"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.ComboBox cboMonth 
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
      ItemData        =   "frmItems.frx":038A
      Left            =   3480
      List            =   "frmItems.frx":03B5
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox cboYear 
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
      ItemData        =   "frmItems.frx":0437
      Left            =   720
      List            =   "frmItems.frx":0439
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
   Begin ComctlLib.ListView lvwAvailable 
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3720
      Width           =   2415
      _ExtentX        =   4260
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
      NumItems        =   6
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
         Text            =   "Count"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date Registered"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Available"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Unavailable"
         Object.Width           =   1411
      EndProperty
   End
   Begin ComctlLib.ListView lvwItems 
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   1680
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
      NumItems        =   7
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
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Location"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remarks"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Image"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date Added"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label5 
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
      Left            =   2280
      TabIndex        =   11
      Top             =   840
      Width           =   105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "List of Items on database"
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
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Width           =   1830
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item List:"
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
      TabIndex        =   8
      Top             =   3360
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Year:"
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
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   435
   End
   Begin ComctlLib.ImageList lstMenu 
      Left            =   6840
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmItems.frx":043B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmItems.frx":078D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmItems.frx":0ADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmItems.frx":0E31
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmItems.frx":1183
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmItems.frx":14D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmItems.frx":1827
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Database Items:"
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
      TabIndex        =   6
      Top             =   1320
      Width           =   1410
   End
   Begin VB.Label lblInventory 
      AutoSize        =   -1  'True
      Caption         =   "List of available items"
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
      Left            =   1200
      TabIndex        =   5
      Top             =   3360
      Width           =   1530
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Month:"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   840
      Width           =   585
   End
End
Attribute VB_Name = "frmItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboMonth_Click()
ViewItems "item_id", "%"
End Sub

Private Sub cboYear_Click()
ViewItems "item_id", "%"
End Sub

Private Sub Form_Load()
SetLv lvwAvailable, True, True
SetLv lvwItems, True, True
mdiMain.tbrMenu.Buttons(1).Value = tbrPressed
mdiMain.sbrDetails.Panels(5).Text = "Click Search for more search options    F3 - Lock the computer"
For i = Format(Date, "yyyy") To 2000 Step -1
    cboYear.AddItem i
Next i
cboYear.AddItem "(View All)"
cboYear.Text = ReadINI("Preferences", "Year")
cboMonth.ListIndex = Val(Format(Date, "mm"))
n = 1
With tbrMenu
    .ImageList = lstMenu
    For i = 1 To lstMenu.ListImages.Count
        .Buttons(i).Image = i
        n = n + 2
    Next i
End With
ViewItems "item_id", "%"
viewAvailable "item_id", "%"
End Sub

Private Sub Form_Resize()
On Error Resume Next
lvwItems.Width = ScaleWidth - (lvwItems.Left + 100)
lvwItems.Height = (ScaleHeight - lvwItems.Top) / 2.5
lblInventory.Top = lvwItems.Height + lvwItems.Top + lvwItems.Left
Label1.Top = lvwItems.Height + lvwItems.Top + lvwItems.Left
lvwAvailable.Top = lblInventory.Top + (lblInventory.Height * 2)
lvwAvailable.Width = ScaleWidth - (lvwItems.Left + 100)
lvwAvailable.Height = ScaleHeight - (lvwAvailable.Top + 100)
End Sub

Public Sub ExecButtons(Index As Integer)
Screen.MousePointer = 11
Select Case Index
    Case 1
        If UserLimit(UserLvl, "User") = True Then Exit Sub
        frmItemManage.Show 1
    Case 2
        If UserLimit(UserLvl, "User") = True Then Exit Sub
        If NoRcrd(lvwItems, "No available record on the list. Please search for a record.") = True Then Exit Sub
        frmItemManage.ExecSrch "item_id", lvwItems.SelectedItem
        frmItemManage.Show 1
    Case 3
        If UserLimit(UserLvl, "User") = True Then Exit Sub
        If NoRcrd(lvwItems, "No available record on the list. Please search for a record.") = True Then Exit Sub
        Screen.MousePointer = 0
        If MsgBox("Your about to delete the item, " & lvwItems.SelectedItem & ". Deleting this item may affect system records" & vbNewLine & vbNewLine & _
                "Are you sure?", vbCritical + vbYesNo) = vbYes Then
            RunSql "Delete * from tblItemList where item_id = '" & lvwItems.SelectedItem & "'"
            ViewItems "item_id", "%"
            viewAvailable "item_id", "%"
            MsgBox "Successfully deleted a record.", vbInformation
        End If
    Case 4
        frmSearch.Show 1
    Case 5
        If UserLimit(UserLvl, "User") = True Then Exit Sub
        If NoRcrd(lvwAvailable) = True Then frmRegister.Show 1: Exit Sub
        frmRegister.ViewUnreg "item_id", lvwItems.SelectedItem
        frmRegister.lvwItems_DblClick
        frmRegister.Show 1
    Case 6
        If UserLimit(UserLvl, "User") = True Then Exit Sub
        If NoRcrd(lvwAvailable) = True Then frmStatus.Show 1: Exit Sub
        frmStatus.ExecSrch "item_id", lvwAvailable.SelectedItem
        frmStatus.cmdLoad_Click
        frmStatus.Show 1
    Case 7
        Unload Me
End Select
Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiMain.tbrMenu.Buttons(1).Value = tbrUnpressed
mdiMain.sbrDetails.Panels(5).Text = Empty
End Sub

Private Sub lvwAvailable_DblClick()
ExecButtons 6
End Sub

Private Sub lvwItems_DblClick()
ExecButtons 2
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    PopupMenu mdiMain.mnuDbItems
End If
End Sub

Private Sub tbrMenu_ButtonClick(ByVal Button As ComctlLib.Button)
Screen.MousePointer = 0
ExecButtons Button.Index
End Sub
Public Sub ViewItems(RcrdFld As String, RcrdStr As String, Optional Order As String = "ASC")
If cboYear.Text <> "(View All)" And cboMonth.Text = "(Unspecified)" Then
    RunSql "Select * from tblItemList where " & RcrdFld & " LIKE '" & RcrdStr & "%' and format(reg_date,'yyyy') = " & cboYear.Text & " Order by " & RcrdFld & " " & Order
Else
    If cboYear.Text = "(View All)" And cboMonth.Text = "(Unspecified)" Then
        RunSql "Select * from tblItemList ORDER by " & RcrdFld & " ASC"
    Else
        If cboYear.Text = "(View All)" And cboMonth.Text <> "(Unspecified)" Then
            RunSql "Select * from tblItemList where " & RcrdFld & " LIKE '" & RcrdStr & "%' and format(reg_date, 'm') = " & _
                cboMonth.ListIndex & " Order by description ASC"
        Else
            RunSql "Select * from tblItemList where " & RcrdFld & " LIKE '" & RcrdStr & "%' and format(reg_date, 'm') = " & _
                cboMonth.ListIndex & " and format(reg_date,'yyyy') = " & cboYear.Text & " Order by description ASC"
        End If
    End If
End If
With Rs
    lvwItems.ListItems.Clear
    While Not .EOF = True
        Set X = lvwItems.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            X.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
If NoRcrd(lvwItems) = True Then Exit Sub
lvwItems.ListItems(1).Selected = True
End Sub
Public Sub viewAvailable(RcrdFld As String, RcrdStr As String, Optional Order As String = "ASC")
RunSql "SELECT reg.*, format((Select sum(stat.qty) from tblItemStatus as stat " & _
        "INNER JOIN tblStatus ON stat.status = tblStatus.description " & _
        "where stat.item_id = reg.item_id and tblStatus.include = 1), '#0') AS Available, " & _
        "format((Select sum(stat.qty) from tblItemStatus as stat " & _
        "INNER JOIN tblStatus ON stat.status = tblStatus.description " & _
        "where stat.item_id = reg.item_id and tblStatus.include = 0), '#0') AS Unavailable " & _
        "FROM tblRegistered AS reg " & _
        "WHERE " & RcrdFld & " LIKE '" & RcrdStr & "%' ORDER BY " & RcrdFld & " " & Order
With Rs
    lvwAvailable.ListItems.Clear
    While Not .EOF = True
        Set X = lvwAvailable.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 1)
            X.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
If NoRcrd(lvwAvailable) = True Then Exit Sub
lvwAvailable.ListItems(1).Selected = True
End Sub
