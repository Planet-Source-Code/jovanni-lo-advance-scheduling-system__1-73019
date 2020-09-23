VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmDb 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5550
   Icon            =   "frmDb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgPath 
      Left            =   3120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtPath 
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
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   4575
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdBrowse 
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "&Browse"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmDb.frx":038A
      Picture         =   "frmDb.frx":0564
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLoad 
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2143
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
      MouseIcon       =   "frmDb.frx":08B8
      Picture         =   "frmDb.frx":0A92
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Browse to locate Database"
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
      TabIndex        =   5
      Top             =   1695
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmDb.frx":0DE6
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Path:"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   435
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   120
      Picture         =   "frmDb.frx":16B0
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LOAD DATABASE"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1620
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Browse Database location"
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
      Width           =   1605
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000004&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdBrowse_Click()
dlgPath.DialogTitle = "Database Path"
dlgPath.InitDir = App.Path & "\Database\"
dlgPath.Filter = "MS Access (*.mdb)|*.mdb|All Files (*.*)|*.*"
dlgPath.ShowOpen
If dlgPath.FileName = "" Then
    Exit Sub
End If
txtPath.Text = dlgPath.FileName
End Sub

Private Sub cmdLoad_Click()
On Error GoTo ErrMsg
Screen.MousePointer = 11
If Trim(txtPath.Text) = Empty Then
    MsgBox "Click Browse to locate database.", vbExclamation
    Exit Sub
End If
s = txtPath.Text
WriteINI "Last Path", "Path", s
Me.Hide
Screen.MousePointer = 0
Exit Sub
ErrMsg:
    MsgBox "Invalid database path.", vbCritical
    txtPath.SetFocus
    SelAll txtPath
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
txtPath.Text = ReadINI("Last Path", "Path")
End Sub

Private Sub Form_Unload(Cancel As Integer)
s = Empty
End Sub

Private Sub txtPath_GotFocus()
SelAll txtPath
End Sub

Private Sub txtPath_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Then
    Unload Me
End If
End Sub
