VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmAcntManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Security"
   ClientHeight    =   6855
   ClientLeft      =   4905
   ClientTop       =   2160
   ClientWidth     =   4590
   Icon            =   "frmAddAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   4590
   StartUpPosition =   1  'CenterOwner
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   16
      Top             =   840
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   53
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account Security"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   4335
      Begin VB.TextBox txtId 
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "l"
         TabIndex        =   1
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   840
         Width           =   2175
      End
      Begin VB.TextBox txtCode 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "l"
         TabIndex        =   2
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cboLevel 
         Height          =   315
         ItemData        =   "frmAddAccount.frx":038A
         Left            =   1920
         List            =   "frmAddAccount.frx":0391
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   3120
         Picture         =   "frmAddAccount.frx":039D
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
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
         TabIndex        =   13
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "New password:"
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
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Secret code:"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User level:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User ID:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblId 
         AutoSize        =   -1  'True
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
         TabIndex        =   8
         Top             =   360
         Width           =   45
      End
   End
   Begin ComctlLib.ListView lstAccounts 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   4335
      _ExtentX        =   7646
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "User ID"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Position"
         Object.Width           =   2540
      EndProperty
   End
   Begin ctrlButton.ThemedButton cmdClear 
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
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
      MouseIcon       =   "frmAddAccount.frx":0C67
      Picture         =   "frmAddAccount.frx":0E41
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   18
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
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
      MouseIcon       =   "frmAddAccount.frx":1195
      Picture         =   "frmAddAccount.frx":136F
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdOptions 
      Height          =   375
      Left            =   3240
      TabIndex        =   19
      Top             =   4080
      Width           =   1215
      _ExtentX        =   2143
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
      MouseIcon       =   "frmAddAccount.frx":16C3
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdEdit 
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Edit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmAddAccount.frx":189D
      Picture         =   "frmAddAccount.frx":1A77
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDelete 
      Height          =   375
      Left            =   3240
      TabIndex        =   21
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
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
      MouseIcon       =   "frmAddAccount.frx":1DCB
      Picture         =   "frmAddAccount.frx":1FA5
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmAddAccount.frx":22F9
      MousePointer    =   99  'Custom
      Picture         =   "frmAddAccount.frx":2BC3
      Stretch         =   -1  'True
      ToolTipText     =   "View warnings"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set user security and system access"
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
      TabIndex        =   15
      Top             =   480
      Width           =   2235
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USER SECURITY"
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
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Click Edit to view"
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
      Top             =   6480
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   0
      Picture         =   "frmAddAccount.frx":3807
      Top             =   6360
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   -120
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmAcntManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SecId As String
Private Sub cmdClear_Click()
'clear all
txtId.Text = Empty
txtId.SetFocus
ClrFlds
End Sub

Private Sub cmdDelete_Click()
If NoRcrd(lstAccounts, "Please add a user account") = True Then Exit Sub
x = MsgBox("Are you sure you want to delete this security account?", vbExclamation + vbYesNo)
If x = vbYes Then
    RunSql "Select * from tblAccountSecurity"
    If Rs.RecordCount = 1 Then
        MsgBox "Delete Failed: You cannot delete this account.", vbCritical
        Exit Sub
    End If
    RunSql "Select * from tblAccountSecurity where id = '" & UserId & "' and level = 'Administrator'"
    With Rs
        If .EOF = False Then
            MsgBox "You cannot delete this account", vbCritical
            Exit Sub
        End If
    End With
    RunSql "delete * from tblAccountSecurity where id = '" & lstAccounts.SelectedItem & "'"
End If
End Sub

Private Sub ExecSrch(ID As String)
If NoRcrd(lstAccounts, "No account found on the list. Please add a user account") = True Then Exit Sub
RunSql "Select * from tblAccountProfile where id LIKE '" & ID & "%'"
SecId = Rs.Fields!ID
txtId.Text = SecId
RunSql "Select * from tblAccountSecurity where id = '" & ID & "'"
With Rs
    If .EOF = False Then
        txtUsername.Text = .Fields!username
        txtPassword.Text = .Fields!password
        txtCode.Text = .Fields!code
        cboLevel.Text = .Fields!Level
        cmdSave.Caption = "&Update"
    End If
End With
End Sub

Public Sub cmdEdit_Click()
ExecSrch lstAccounts.SelectedItem
End Sub

Private Sub cmdSave_Click()
'if no account were selecte on the list or no user accounts on database
If SecId = Empty Then
    MsgBox "No specified Account found. Please add a new Account or select an Account from the list. Click Options", vbExclamation
    ClrFlds
    Exit Sub
End If
'text trappings
If TxtEmp(txtUsername) = True Then Exit Sub
If TxtEmp(txtPassword) = True Then Exit Sub
If TxtEmp(txtCode) = True Then Exit Sub
If CboEmp(cboLevel) = True Then Exit Sub

RunSql "Select * from tblAccountSecurity where id = '" & SecId & "'"
With Rs
    If cmdSave.Caption <> "&Update" Then
        If cboLevel.Text <> "Administrator" Then
            SubSql "Select * from tblAccountSecurity where level = 'Administrator' and id <> '" & SecId & "'"
            If SubRs.EOF = True Then
                MsgBox "Update failed: No any Administrator found on accounts.  You are not allowed to change the level of this account", vbCritical
                Exit Sub
            End If
        End If
        SubSql "Select * from tblAccountSecurity where level = 'Administrator'"
        If SubRs.EOF = True Then
            If cboLevel.Text <> "Administrator" Then
                MsgBox "You cannot add another level without 'Administrator'", vbInformation
                cboLevel.SetFocus
                Exit Sub
            End If
        End If
        .AddNew
        MSG = "Add security account of user " & SecId
    Else
        SubSql "Select * from tblAccountSecurity where username = '" & txtUsername & "'"
        If SubRs.EOF = False Then
            MsgBox "Username '" & txtUsername.Text & "' is already taken.", vbInformation
            txtUsername.SetFocus
            txtUsername.SelStart = 0
            txtUsername.SelLength = Len(txtUsername)
            Exit Sub
        End If
        MSG = "Record has been successfully updated"
    End If
    .Fields!ID = SecId
    .Fields!username = txtUsername
    .Fields!password = txtPassword
    .Fields!Level = cboLevel
    .Fields!code = txtCode
    .Update
End With
MsgBox MSG, vbInformation
ClrFlds
RcrdView
End Sub

Public Sub RcrdView()
RunSql "Select * from tblAccountProfile"
lstAccounts.ListItems.Clear
With Rs
    While Not .EOF
        Set x = lstAccounts.ListItems.Add(, , .Fields!ID)
        x.SubItems(1) = .Fields!Position
        .MoveNext
    Wend
End With
End Sub

Public Sub ClrFlds()
txtUsername.Text = Empty
txtPassword.Text = Empty
txtCode.Text = Empty
cboLevel.Text = cboLevel.List(0)
SecId = Empty
cmdSave.Caption = "&Save"
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdOptions_Click()
'hide and show options
    If cmdOptions.Caption = "&Options >>" Then
        Me.Height = 7335
        cmdOptions.Caption = "&Options <<"
    Else
        Me.Height = 5070
        cmdOptions.Caption = "&Options >>"
    End If
End Sub

Private Sub Form_Load()
SetLv lstAccounts, True, True
RunSql "Select * from tblAccountLevel"
While Not Rs.EOF = True
    cboLevel.AddItem (Rs.Fields!Description)
    Rs.MoveNext
Wend
RcrdView
End Sub

Private Sub Form_Unload(Cancel As Integer)
If FrstUsr = True Then
    RunSql "Select * from tblAccountSecurity"
    If Rs.EOF = True Then
        MsgBox "System has failed to initialized. Please contact your administrator" & vbNewLine _
                & vbNewLine & "Status: analysing accounts configuration" & vbNewLine & _
                "Error: No security account found", vbCritical
        End
    End If
    frmxSplash.tmrLoad.Enabled = True
End If
End Sub

Private Sub lstAccounts_DblClick()
cmdEdit_Click
End Sub

Private Sub txtCode_GotFocus()
SelAll txtCode
End Sub

Private Sub txtId_Change()
cmdEdit_Click
End Sub

Private Sub txtPassword_GotFocus()
SelAll txtPassword
End Sub

Private Sub txtUsername_GotFocus()
SelAll txtUsername
End Sub
