VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System"
   ClientHeight    =   4110
   ClientLeft      =   5505
   ClientTop       =   5280
   ClientWidth     =   4590
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkJames 
      Caption         =   "Load james after logged in"
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
      Left            =   240
      TabIndex        =   15
      Top             =   2640
      Width           =   2220
   End
   Begin VB.ComboBox cboUser 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   14
      Top             =   840
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   53
   End
   Begin VB.TextBox txtUserCode 
      Enabled         =   0   'False
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3120
      Width           =   2535
   End
   Begin VB.CheckBox chkNoAccess 
      Caption         =   "Unable to Access"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   2640
      Width           =   1500
   End
   Begin VB.TextBox txtPassword 
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "l"
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
   Begin ctrlButton.ThemedButton cmdCancel 
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmLogin.frx":038A
      Picture         =   "frmLogin.frx":0564
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLogin 
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Log-in"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmLogin.frx":08B8
      Picture         =   "frmLogin.frx":0A92
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdOptions 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2143
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
      MouseIcon       =   "frmLogin.frx":0DE6
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdHelp 
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
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
      MouseIcon       =   "frmLogin.frx":0FC0
      Picture         =   "frmLogin.frx":119A
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdRetrieve 
      Height          =   375
      Left            =   1920
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Retrieve"
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmLogin.frx":14EE
      Picture         =   "frmLogin.frx":16C8
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   4200
      Picture         =   "frmLogin.frx":1A1C
      Top             =   1125
      Width           =   240
   End
   Begin AgentObjectsCtl.Agent ctlAgent 
      Left            =   120
      Top             =   3480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      Index           =   1
      Left            =   2520
      TabIndex        =   16
      Top             =   2640
      Width           =   105
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   4080
      Picture         =   "frmLogin.frx":1FA6
      Top             =   1485
      Width           =   480
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log-in to start the system"
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
      TabIndex        =   13
      Top             =   480
      Width           =   1575
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ACCESS"
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
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmLogin.frx":2870
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin.frx":313A
      Stretch         =   -1  'True
      ToolTipText     =   "View warnings"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblcode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   600
      TabIndex        =   11
      Top             =   3120
      Width           =   1050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
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
      Left            =   720
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   915
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   -240
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim James As Variant

Private Sub chkNoAccess_Click()
'if forgot password?
If chkNoAccess.Value = 1 Then
    txtUserCode.Enabled = True
    cmdRetrieve.Enabled = True
    cmdRetrieve.Default = True
    txtUserCode.SetFocus
Else
    txtUserCode.Enabled = False
    cmdRetrieve.Enabled = False
    cmdLogin.Default = True
    cboUser.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
UserLvl = ""
UserNme = ""
UserId = ""
Unload Me
End Sub

Private Sub cmdHelp_Click()
PathToDoc = App.Path & "\help.chm"
ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5
End Sub

Private Sub UserFound(uName As String, uLvl As String, UID As String)
'set user values
UserNme = uName
UserLvl = uLvl
UserId = UID
LoadJames
mdiMain.cmdLogout.Caption = "Log-out"
    'set status bar on main form
With mdiMain.sbrDetails
    .Panels(1).Text = "User: " & UserNme & "  "
    .Panels(2).Text = "Level: " & UserLvl & "  "
    .Panels(3).Text = "User ID: " & UserId & "  "
    .Panels(4).Text = "PC Name: " & PcId & "  "
End With
WriteINI "Last User", "UserName", UserNme
Screen.MousePointer = 0
Me.Hide
FrmShow UserLvl
End Sub

Public Sub LoadJames()
On Error Resume Next

If chkJames.Value = 1 Then
    ctlAgent.Characters.Load "James", App.Path & "\James.acs"
    Set James = ctlAgent.Characters("James")
    James.Show
    James.MoveTo 900, 550
    James.Play "Greet"
    James.Speak "Hello " & UserNme & ", welcome to ACSAT Tool Room Scheduling System."
    James.Play "GestureDown"
    James.Speak "I am James, your virtual assistant for today."
    James.Play "Think"
    James.Play "Read"
    James.Speak "If you want a short tour of this system, right click me and select 'Tour'"
    James.Play "ReadReturn"
    James.Play "GestureDown"
    James.Commands.Add "Tour", "Tour", "...Tour...", True, True
    James.Commands.Add "About", "About", "...About...", True, True
    James.Commands.Add "Help", "Help", "...Help...", True, True
End If

Me.MousePointer = 0
End Sub

Private Sub ctlAgent_Command(ByVal UserInput As Object)
With mdiMain
    'agent james command languages
    Select Case UserInput.Name
        Case "Tour"
            Set James = ctlAgent.Characters("James")
            James.Show
            James.Play "Announce"
            James.Speak "Ok! Let's start the tour."
            James.MoveTo 0, 0
            James.Speak "You are currently the " & UserLvl & " of this system."
            James.Play "Think"
            James.MoveTo 470, 280
            James.Speak "Your Access privilege will depend on your user level"
            James.Speak "and only the Administrator can access all the task of this system."
            James.Play "Explain"
            James.Speak "This system includes transactions, reports, user individaul records, and some other stuffs."
            James.MoveTo 400, 0
            James.Play "GestureRight"
            James.Speak "This section is the Menu Toolbar section, it acts as the guide for the basic operations of this system."
            James.MoveTo 0, 650
            James.Play "Read"
            James.Play "ReadReturn"
            James.Play "GestureUp"
            James.Speak "This section is the sidebar section, it contains categorized task for you to manage and choose from."
            James.MoveTo 150, 150
            James.Play "Announce"
            James.Play "GestureRight"
            James.Speak "It includes System Task, Scheduling which is the main process of this system. Manage item status and transactions."
            James.MoveTo 500, 300
            James.Play "Announce"
            James.Speak "Some of the functions of this system will be explained by the programmers."
            James.Play "GestureDown"
            James.Play "Announce"
            James.Play "Explain"
            James.Speak "Ok, that's the coverage of our tour. Hit F1 for Help or right click me and click Help."
            James.Play "Alert"
            James.Speak "Have a good day " & UserNme & "."
            James.Hide
        Case "About"
            frmAbout.Show 1
        Case "Help"
            PathToDoc = App.Path & "\help.chm"
            ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5
    End Select
End With
End Sub

Private Sub cmdLogin_Click()
Screen.MousePointer = 11
If TxtEmp(cboUser) = True Then Exit Sub
If TxtEmp(txtPassword) = True Then Exit Sub

If ReadINI("Preferences", "Case Sensitive") = "1" Then
    RunSql "Select * from tblAccountSecurity"
    With Rs
        While Not .EOF = True
            If .Fields!username = cboUser.Text And .Fields!password = txtPassword.Text Then
                UserFound .Fields!username, .Fields!Level, .Fields!ID
                Exit Sub
            End If
            .MoveNext
    Wend
    End With
Else
    RunSql "SELECT * FROM tblAccountSecurity WHERE username = '" & cboUser & "' and password = '" & txtPassword & "'"
    With Rs
        If .EOF = False Then
            UserFound .Fields!username, .Fields!Level, .Fields!ID
            Exit Sub
        End If
    End With
End If
cboUser.SelStart = 0
Screen.MousePointer = 0
MsgBox "Invalid 'Username' or 'Password'", vbExclamation
cboUser.SetFocus
cboUser.SelLength = Len(cboUser)
End Sub

Private Sub cmdOptions_Click()
    'show the options below
    If cmdOptions.Caption = "&Options >>" Then
        Me.Height = 4590
        cmdOptions.Caption = "&Options <<"
    Else
        Me.Height = 3000
        cmdOptions.Caption = "&Options >>"
        cboUser.SetFocus
    End If
End Sub

Private Sub cmdRetrieve_Click()
If TxtEmp(txtUserCode) = True Then Exit Sub

'scan security accounts that mathches the security code
RunSql "Select * from tblAccountSecurity Where code = '" & txtUserCode & "'"
With Rs
    If .EOF = True Then
        MsgBox "No user account found", vbExclamation
        txtUserCode.SetFocus
        txtUserCode.SelStart = 0
        txtUserCode.SelLength = Len(txtUserCode)
        Exit Sub
    End If
    'load the user security info if code mathches
    MsgBox "Your Username is '" & .Fields!username & "'; Password is '" & .Fields!password & "'", vbInformation
    cboUser.SetFocus
    cmdLogin.Default = True
    chkNoAccess.Value = 0
End With
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
txtPassword.Text = Empty
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
PcId = Environ("ComputerName")
Me.Caption = Me.Caption & " [" & PcId & "]"
Me.Height = 3000
LoadCbo "tblAccountSecurity", cboUser, "username"
If ReadINI("Preferences", "Remember User") = "1" Then
    cboUser.Text = ReadINI("Last User", "UserName")
End If
chkJames.Value = Val(ReadINI("Preferences", "James"))
Form_Unload 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If UserId = Empty Then
    mdiMain.cmdLogout.Caption = "&Log-in"
    With mdiMain.sbrDetails
        .Panels(1).Text = "User: ---"
        .Panels(2).Text = "Level: ---"
        .Panels(3).Text = "User ID: ---"
        .Panels(4).Text = "PC Name: ---"
    End With
End If
Screen.MousePointer = 0
End Sub

Private Sub txtPassword_GotFocus()
txtPassword.SelStart = 0
txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub
Private Sub txtUserCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub cbouser_GotFocus()
cboUser.SelStart = 0
cboUser.SelLength = Len(cboUser)
End Sub

Private Sub cbouser_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub
