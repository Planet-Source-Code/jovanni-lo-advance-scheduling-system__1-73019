VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmScheduler 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   7950
   Icon            =   "frmScheduler.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2280
      TabIndex        =   14
      Text            =   "Search"
      Top             =   4560
      Width           =   2175
   End
   Begin VB.ComboBox cboFilter 
      Height          =   315
      ItemData        =   "frmScheduler.frx":038A
      Left            =   120
      List            =   "frmScheduler.frx":038C
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Frame Frame2 
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
      Height          =   2775
      Left            =   3480
      TabIndex        =   5
      Top             =   960
      Width           =   4335
      Begin VB.TextBox txtTitle 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker dtpDate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2520
         TabIndex        =   9
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47710211
         CurrentDate     =   40071
      End
      Begin MSMask.MaskEdBox txtDate 
         Height          =   285
         Left            =   1440
         TabIndex        =   10
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.TextBox txtSchedRmrks 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         Height          =   1245
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Remarks:"
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
         TabIndex        =   11
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Date:"
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
         TabIndex        =   8
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "* Title:"
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
         TabIndex        =   7
         Top             =   360
         Width           =   570
      End
   End
   Begin VB.Frame freCalendar 
      Caption         =   "Calendar - 03/04/2010"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3255
      Begin MSComCtl2.MonthView mvwSched 
         Height          =   2370
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483644
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
         ScrollRate      =   1
         ShowWeekNumbers =   -1  'True
         StartOfWeek     =   47710209
         TitleBackColor  =   16744448
         TitleForeColor  =   16777215
         TrailingForeColor=   -2147483632
         CurrentDate     =   40099
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   53
   End
   Begin ComctlLib.ListView lvwRecords 
      Height          =   1455
      Left            =   120
      TabIndex        =   15
      Top             =   5040
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2566
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
      NumItems        =   5
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
         Text            =   "Title"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remarks"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "User ID"
         Object.Width           =   2540
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   1920
      TabIndex        =   16
      Top             =   4680
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   3960
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
      MouseIcon       =   "frmScheduler.frx":038E
      Picture         =   "frmScheduler.frx":0568
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdNew 
      Height          =   375
      Left            =   3600
      TabIndex        =   22
      Top             =   3960
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
      MouseIcon       =   "frmScheduler.frx":08BC
      Picture         =   "frmScheduler.frx":0A96
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   5040
      TabIndex        =   23
      Top             =   3960
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
      MouseIcon       =   "frmScheduler.frx":0DEA
      Picture         =   "frmScheduler.frx":0FC4
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdOptions 
      Height          =   375
      Left            =   6480
      TabIndex        =   24
      Top             =   3960
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
      MouseIcon       =   "frmScheduler.frx":1318
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLoad 
      Height          =   375
      Left            =   3600
      TabIndex        =   25
      Top             =   6600
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
      MouseIcon       =   "frmScheduler.frx":14F2
      Picture         =   "frmScheduler.frx":16CC
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDelete 
      Height          =   375
      Left            =   6480
      TabIndex        =   26
      Top             =   6600
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
      MouseIcon       =   "frmScheduler.frx":1A20
      Picture         =   "frmScheduler.frx":1BFA
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdNotes 
      Height          =   375
      Left            =   5040
      TabIndex        =   27
      Top             =   6600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Notes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmScheduler.frx":1F4E
      Picture         =   "frmScheduler.frx":2128
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Load to manage a schedule"
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
      TabIndex        =   20
      Top             =   6720
      Width           =   2325
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   0
      Picture         =   "frmScheduler.frx":247C
      Top             =   6600
      Width           =   480
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   7320
      TabIndex        =   19
      Top             =   4560
      Width           =   180
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Schedules:"
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
      Left            =   5280
      TabIndex        =   18
      Top             =   4560
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
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
      Left            =   5040
      TabIndex        =   17
      Top             =   4560
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4440
      Picture         =   "frmScheduler.frx":2D46
      Top             =   4440
      Width           =   480
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set system scheduler"
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
      TabIndex        =   2
      Top             =   480
      Width           =   1320
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SCHEDULER"
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
      TabIndex        =   1
      Top             =   120
      Width           =   1080
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "frmScheduler.frx":3610
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
      Width           =   8535
   End
End
Attribute VB_Name = "frmScheduler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If NoRcrd(lvwRecords, "No record available on the list.") = True Then Exit Sub
x = MsgBox("Your about to delete this schedule, '" & lvwRecords.SelectedItem.SubItems(1) & "'." & _
            " Do you want to continue?", vbExclamation + vbYesNo)
If x = vbYes Then
    RunSql "Delete * from tblSchedules where record_no = " & lvwRecords.SelectedItem
    MsgBox "Scheduled task has been successfully deleted.", vbInformation
End If
ClrFlds
End Sub

Public Sub cmdLoad_Click()
If NoRcrd(lvwRecords) = True Then Exit Sub
RunSql "Select * from tblSchedules where record_no = " & lvwRecords.SelectedItem
With Rs
    txtTitle.Text = .Fields!Title
    txtDate.Text = Format(.Fields!sched_date, "mm/dd/yyyy")
    txtSchedRmrks.Text = .Fields!remarks
    cmdSave.Caption = "&Update"
    mvwSched.Value = .Fields!sched_date
    mvwSched_Click
End With
End Sub

Private Sub cmdNew_Click()
ClrFlds
End Sub

Private Sub cmdNotes_Click()
frmNotes.Show 1
End Sub

Private Sub cmdOptions_Click()
If cmdOptions.Caption = "&Options >>" Then
    Me.Height = 7605
    cmdOptions.Caption = "&Options <<"
Else
    Me.Height = 4950
    cmdOptions.Caption = "&Options >>"
End If
End Sub

Private Sub cmdSave_Click()
If TxtEmp(txtTitle) = True Then Exit Sub
If TxtEmp(txtDate) = True Then Exit Sub
If TxtEmp(txtSchedRmrks) = True Then Exit Sub

RunSql "Select * from tblSchedules where title = '" & txtTitle.Text & "'"
With Rs
    If cmdSave.Caption = "&Save" Then
        .AddNew
        .Fields!record_no = Val(RcrdId("tblSchedules", , "record_no"))
        s = "Add new system schedule."
    Else
        s = "Schedule has been updated."
    End If
    .Fields!Title = txtTitle.Text
    .Fields!sched_date = Format(txtDate.Text, "mm/dd/yyyy")
    .Fields!remarks = txtSchedRmrks.Text
    .Fields!user_id = UserId
    .Update
End With
MsgBox s, vbInformation
ClrFlds
End Sub

Public Sub ViewRcrds(RcrdFld As String, RcrdStr As String)
RunSql "Select * from tblSchedules where " & RcrdFld & " LIKE '" & RcrdStr & "%'"
With Rs
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

Private Sub ClrFlds()
cmdSave.Caption = "&Save"
txtTitle.Text = Empty
txtDate.Text = "  /  /    "
txtSchedRmrks.Text = Empty
ViewRcrds "record_no", "%"
RunSql "Select * from tblSchedules"
lblCount.Caption = Rs.RecordCount
mvwSched.Value = Date
mvwSched_Click
End Sub

Private Sub cmdView_Click()
If cmdOptions.Caption = "&Options >>" Then
    Me.Height = 7575
    cmdOptions.Caption = "&Options <<"
Else
    Me.Height = 4935
    cmdOptions.Caption = "&Options >>"
End If
End Sub

Private Sub dtpDate_Change()
txtDate.Text = Format(dtpDate.Value, "mm/dd/yyyy")
End Sub

Private Sub Form_Load()
SetLv lvwRecords, True, True
DtpValue dtpDate

RunSql "Select * from tblSchedules"
With Rs
    For i = 0 To (.Fields.Count - 1)
        cboFilter.AddItem (.Fields(i).Name)
    Next i
End With
cboFilter.ListIndex = 0
ClrFlds
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 0
End Sub

Private Sub lvwRecords_DblClick()
cmdLoad_Click
End Sub

Private Sub mvwSched_Click()
freCalendar.Caption = "Calendar - " & mvwSched.Value
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ViewRcrds cboFilter.Text, txtSrchStr.Text
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
    txtSrchStr.ForeColor = &H8000000B
End If
End Sub

