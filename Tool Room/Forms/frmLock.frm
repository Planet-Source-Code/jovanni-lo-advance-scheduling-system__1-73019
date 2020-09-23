VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "System Locked"
   ClientHeight    =   3960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7530
   Icon            =   "frmLock.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7530
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.Frame frePass 
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1815
      Left            =   1560
      TabIndex        =   2
      Top             =   1800
      Width           =   3015
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   380
         Left            =   120
         TabIndex        =   5
         Text            =   "Password"
         Top             =   595
         Width           =   2775
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         Caption         =   "Invalid Password!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         Top             =   1080
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Access failed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Image imgWarning 
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmLock.frx":038A
      MousePointer    =   99  'Custom
      Picture         =   "frmLock.frx":0C54
      Stretch         =   -1  'True
      ToolTipText     =   "View warnings"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SYSTEM LOCKED"
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
      Width           =   1530
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This System is locked by the Administrator"
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
      TabIndex        =   0
      Top             =   480
      Width           =   2670
   End
   Begin VB.Shape shpHeader 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
txtPassword.Text = "Password"
End Sub

Private Sub Form_Resize()
On Error Resume Next
shpHeader.Width = ScaleWidth - shpHeader.Left
frePass.Left = (ScaleWidth / 2) - (frePass.Width / 2)
frePass.Top = (ScaleHeight / 2) - (frePass.Height / 2)
End Sub
Private Sub txtPassword_Click()
With txtPassword
    If .Text = "Password" And .SelStart = 8 Then
        .Text = Empty
    End If
End With
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    RunSql "Select * from tblAccountSecurity"
    With Rs
        While Not .EOF = True
            If .Fields!ID = UserId And .Fields!password = txtPassword.Text Then
                Unload Me
                mdiMain.Show
                Exit Sub
            End If
            .MoveNext
        Wend
        txtPassword.Text = Empty
        lblPrompt.Visible = True
        lblMsg.Visible = True
    End With
End If
With txtPassword
    If Trim(.Text) = "Password" Or Trim(.Text) = Empty Then
        If KeyAscii = 8 Then Exit Sub
        .FontName = "Wingdings"
        .PasswordChar = "l"
        .Text = Empty
        .ForeColor = &H80000012
        .Text = Left(Str(KeyAscii), 1)
        .Text = Trim(.Text)
    ElseIf Len(.Text) = 1 And KeyAscii = 8 Then
        .PasswordChar = ""
        .Text = "Password"
        .ForeColor = &H8000000C
        .FontName = "Tahoma"
        lblPrompt.Visible = False
        lblMsg.Visible = False
    End If
End With
End Sub
