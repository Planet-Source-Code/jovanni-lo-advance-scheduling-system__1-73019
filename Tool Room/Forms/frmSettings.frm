VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   555
   ClientWidth     =   6855
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freSettings 
      BorderStyle     =   0  'None
      Height          =   4575
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   6375
      Begin VB.CheckBox chkJames 
         Caption         =   "Always check James on log-in form"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   2895
      End
      Begin VB.CheckBox chkSensitive 
         Caption         =   "Case sensitive username and password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox chkUser 
         Caption         =   "&Remember user after logged in"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   1560
         Width           =   2535
      End
      Begin VB.ComboBox cboForms 
         Height          =   315
         ItemData        =   "frmSettings.frx":038A
         Left            =   2640
         List            =   "frmSettings.frx":039A
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1920
         Width           =   3495
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmSettings.frx":03D4
         Left            =   2640
         List            =   "frmSettings.frx":03D6
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3960
         Width           =   3495
      End
      Begin VB.CheckBox chkAllow 
         Caption         =   "allow other users to manage records along with the Administrator"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   5055
      End
      Begin VB.CheckBox chkTask 
         Caption         =   "Remove all late scheduled task and reminders"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   2760
         Width           =   3615
      End
      Begin VB.ComboBox cboYear 
         Height          =   315
         ItemData        =   "frmSettings.frx":03D8
         Left            =   2640
         List            =   "frmSettings.frx":03DA
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3480
         Width           =   3495
      End
      Begin VB.CheckBox chkDb 
         Caption         =   "Set current database source as default (Load Database will not prompt)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   5535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Default form after logged in:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Default item search filter:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   3960
         Width           =   1830
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Log-in Settings"
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
         TabIndex        =   27
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "System Set-up"
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
         TabIndex        =   26
         Top             =   2400
         Width           =   1260
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Default system year set-up:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   3480
         Width           =   2025
      End
   End
   Begin VB.Frame freSettings 
      BorderStyle     =   0  'None
      Height          =   855
      Index           =   1
      Left            =   2400
      TabIndex        =   8
      Top             =   0
      Width           =   495
      Begin VB.TextBox txtRemarks 
         Height          =   495
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   3840
         Width           =   4815
      End
      Begin VB.TextBox txtDescription 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   3360
         Width           =   3975
      End
      Begin ComctlLib.ListView lvwRecords 
         Height          =   2655
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4683
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Stored Records"
            Object.Width           =   7056
         EndProperty
      End
      Begin ComctlLib.TreeView tvwTables 
         Height          =   2655
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4683
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
      Begin VB.Label Label1 
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
         Left            =   240
         TabIndex        =   31
         Top             =   3840
         Width           =   810
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
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
         Left            =   4680
         TabIndex        =   14
         Top             =   120
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No. of Records:"
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
         Left            =   3240
         TabIndex        =   13
         Top             =   120
         Width           =   1245
      End
      Begin VB.Label Label2 
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
         TabIndex        =   12
         Top             =   3360
         Width           =   1005
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Data Categoires"
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
         Top             =   120
         Width           =   1365
      End
   End
   Begin ComctlLib.TabStrip tabSettings 
      Height          =   5055
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8916
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Stored Data"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "System Preferences"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   6120
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
      MouseIcon       =   "frmSettings.frx":03DC
      Picture         =   "frmSettings.frx":05B6
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdNew 
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   6120
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
      MouseIcon       =   "frmSettings.frx":090A
      Picture         =   "frmSettings.frx":0AE4
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   6120
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
      MouseIcon       =   "frmSettings.frx":0E38
      Picture         =   "frmSettings.frx":1012
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDelete 
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   6120
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
      MouseIcon       =   "frmSettings.frx":1366
      Picture         =   "frmSettings.frx":1540
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ComctlLib.ImageList lstIcons 
      Left            =   4320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSettings.frx":1894
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSettings.frx":1BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSettings.frx":1F38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmSettings.frx":228A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmSettings.frx":25DC
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SETTINGS"
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
      TabIndex        =   2
      Top             =   120
      Width           =   870
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Update system settings"
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
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   -120
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Table As String

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If Table = Empty Then
    MsgBox "please select a data category to be updated", vbExclamation
    Exit Sub
End If
If NoRcrd(lvwRecords, "No record available on the list.") = True Then Exit Sub
SubSql "Select * from " & Table & " where description = '" & lvwRecords.SelectedItem & "'"
If Table = "tblStatus" Then
    If SubRs.Fields!System = True Then
        MsgBox "You are not allowed to delete this record. It is a system record.", vbCritical
        Exit Sub
    End If
End If

If RunSql("Delete * from " & Table & " where description = '" & SubRs.Fields!Description & "'") = True Then
    MsgBox "System cannot delete this record because it is beeing used by other records.", vbCritical
    Exit Sub
End If
MsgBox "Record has been deleted successfully", vbInformation
ViewRcrds Table
ClrFlds
End Sub

Private Sub cmdNew_Click()
If tabSettings.SelectedItem.Index = 2 Then
    chkDb.Value = 0
    chkSensitive.Value = 1
    chkJames.Value = 0
    chkUser.Value = 1
    cboForms.Text = "Items"
    chkAllow.Value = 0
    cboYear.ListIndex = 0
    cboFilter.Text = "item_id"
Else
    ClrFlds
End If
End Sub

Private Sub cmdSave_Click()
Dim MSG As String

If tabSettings.SelectedItem.Index = 2 Then
    WriteINI "Preferences", "Default DB", chkDb.Value
    WriteINI "Preferences", "Case Sensitive", chkSensitive.Value
    WriteINI "Preferences", "Remember User", chkUser.Value
    WriteINI "Preferences", "James", chkJames.Value
    WriteINI "Preferences", "Default Form", cboForms.Text
    WriteINI "Preferences", "Other Users", chkAllow.Value
    WriteINI "Preferences", "Year", cboYear.Text
    WriteINI "Preferences", "Filter", cboFilter.Text
    WriteINI "Preferences", "Late Task", chkTask.Value
    MsgBox "Some changes will effect after you restart the system. " & _
             "Settings were saved successfully.", vbInformation
    Exit Sub
End If

If Table = Empty Then
    MsgBox "Please select a data category to be updated", vbExclamation
    tvwTables.SetFocus
    Exit Sub
End If
If TxtEmp(txtDescription) = True Then Exit Sub

If cmdSave.Caption = "&Save" Then
    RunSql "Select * from " & Table & " where description = '" & txtDescription.Text & "'"
    With Rs
        If .EOF = False Then
            MsgBox "The record you provided is already on the list. Please input again.", vbExclamation
            txtDescription.SetFocus
            SelAll txtDescription
            Exit Sub
        End If
    End With
End If

RunSql "Select * from " & Table & " where description = '" & lvwRecords.SelectedItem & "'"
With Rs
    If Table = "tblStatus" Or Table = "tblAccountLevel" And cmdSave.Caption = "&Update" Then
        If .Fields!System = True Then
            MsgBox "You cannot update this record because it is a system record.", vbCritical
            Exit Sub
        End If
    End If
    If cmdSave.Caption <> "&Update" Then
        .AddNew
        MSG = "New record has been added successfully"
    Else
        MSG = "Record has been updated successfully"
    End If
    .Fields!Description = txtDescription.Text
    If Table = "tblStatus" Then
        n = ValBox("Input 1 to include an item of this status, otherwise 0.", imgIcon, , .Fields!include, "item status")
        If n < 0 Or n > 1 Then
            MsgBox "System cannot accept this input. Process oborted.", vbExclamation
            Exit Sub
        End If
        .Fields!include = n
    End If
    .Fields!remarks = txtRemarks.Text
    .Update
    MsgBox MSG, vbInformation
End With
ViewRcrds Table
ClrFlds
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
SetLv lvwRecords, True, True
mdiMain.tbrMenu.Buttons(5).Value = tbrPressed

For i = 2 To freSettings.UBound
    freSettings(1).Height = tabSettings.Height - 420
    freSettings(1).Width = tabSettings.Width - 120
    freSettings(1).Top = tabSettings.Top + 350
    freSettings(1).Left = tabSettings.Left + 40
    freSettings(i).Move _
        freSettings(1).Left, _
        freSettings(1).Top, _
        freSettings(1).Width, _
        freSettings(1).Height
    freSettings(i).Visible = False
Next i

With tvwTables.Nodes
    .Add , , "Accounts", "Accounts", 1
    .Add "Accounts", tvwChild, "tblAccountLevel", "Account Levels", 4
    .Add "Accounts", tvwChild, "tblAccountPosition", "Employee Positions", 4
    
    .Add , , "Items", "Items", 2
    .Add "Items", tvwChild, "tblCategories", "Item Cateogies", 4
    .Add "Items", tvwChild, "tblLocations", "Locations", 4
    .Add "Items", tvwChild, "tblStatus", "Status and Conditions", 4
    
    .Add , , "Users", "User Profiles", 3
    .Add "Users", tvwChild, "tblCourse", "Courses", 4
    .Add "Users", tvwChild, "tblDepartments", "Departments", 4
    .Item(1).Expanded = ReadINI("Settings Nodes", "1")
    .Item(4).Expanded = ReadINI("Settings Nodes", "4")
    .Item(8).Expanded = ReadINI("Settings Nodes", "8")
End With
IniSetup
End Sub

Private Sub IniSetup()
chkDb.Value = Val(ReadINI("Preferences", "Default DB"))
chkSensitive.Value = Val(ReadINI("Preferences", "Case Sensitive"))
chkUser.Value = Val(ReadINI("Preferences", "Remember User"))
cboForms.Text = ReadINI("Preferences", "Default Form")
chkAllow.Value = Val(ReadINI("Preferences", "Other Users"))
chkJames.Value = Val(ReadINI("Preferences", "James"))
chkTask.Value = Val(ReadINI("Preferences", "Late Task"))
For i = Format(Date, "yyyy") To 2000 Step -1
    cboYear.AddItem i
Next i
cboYear.Text = ReadINI("Preferences", "Year")
LoadCboFld "tblItemList", "*", cboFilter
cboFilter.Text = ReadINI("Preferences", "Filter")
End Sub

Private Sub ClrFlds()
txtDescription.Text = Empty
txtRemarks.Text = Empty
cmdSave.Caption = "&Save"
lblCount.Caption = "---"
End Sub

Private Sub ViewRcrds(Table As String)
RunSql "Select description from " & Table & " order by description ASC"
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
lblCount.Caption = lvwRecords.ListItems.Count
End Sub

Private Sub LoadRcrds(Table As String, Description As String)
RunSql "Select * from " & Table & " where description = '" & Description & "'"
txtDescription.Text = Rs.Fields!Description
txtRemarks.Text = Rs.Fields!remarks
cmdSave.Caption = "&Update"
End Sub

Private Sub Form_Unload(Cancel As Integer)
mdiMain.tbrMenu.Buttons(5).Value = tbrUnpressed
WriteINI "Settings Nodes", "1", tvwTables.Nodes(1).Expanded
WriteINI "Settings Nodes", "4", tvwTables.Nodes(4).Expanded
WriteINI "Settings Nodes", "8", tvwTables.Nodes(8).Expanded
End Sub

Private Sub lvwRecords_Click()
If NoRcrd(lvwRecords) = True Then Exit Sub
If Table = Empty Then Exit Sub
LoadRcrds Table, lvwRecords.SelectedItem
End Sub

Private Sub tvwTables_Click()
On Error Resume Next
ClrFlds
Select Case tvwTables.SelectedItem.Key
    Case "Items", "Accounts", "Users"
        lvwRecords.ListItems.Clear
        ClrFlds
        Table = Empty
    Case Else
        Table = tvwTables.SelectedItem.Key
        ViewRcrds Table
End Select
End Sub

Private Sub tabSettings_Click()
For i = 1 To tabSettings.Tabs.Count
    If freSettings(i).Index = tabSettings.SelectedItem.Index Then
        freSettings(i).Visible = True
    Else
        freSettings(i).Visible = False
    End If
Next i

Select Case tabSettings.SelectedItem.Index
    Case 1
        cmdNew.Caption = "&New"
        cmdDelete.Enabled = True
    Case 2
        cmdNew.Caption = "&Defaults"
        cmdDelete.Enabled = False
End Select
End Sub

