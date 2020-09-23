VERSION 5.00
Begin VB.Form frmxSplash 
   BackColor       =   &H00000007&
   BorderStyle     =   0  'None
   ClientHeight    =   2175
   ClientLeft      =   5250
   ClientTop       =   3270
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   Picture         =   "frmSplash.frx":4ACE3
   ScaleHeight     =   2175
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrLoad 
      Interval        =   18
      Left            =   2400
      Top             =   840
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   120
      Picture         =   "frmSplash.frx":81467
      Top             =   120
      Width           =   555
   End
   Begin VB.Shape pgbLoad 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   90
      Index           =   2
      Left            =   5280
      Top             =   1980
      Width           =   90
   End
   Begin VB.Shape pgbLoad 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   90
      Index           =   1
      Left            =   5160
      Top             =   1980
      Width           =   90
   End
   Begin VB.Shape pgbLoad 
      BackColor       =   &H80000003&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808000&
      Height          =   90
      Index           =   0
      Left            =   5040
      Top             =   1980
      Width           =   90
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "initializing system..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   165
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "frmxSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cntr As Integer
Dim DbShow As Boolean, UserTest As Boolean

Private Sub Form_Load()
DbShow = False
UserTest = False
End Sub

Private Sub tmrLoad_Timer()

'set the progress bar codes
For i = 0 To 2
If pgbLoad(i).Left = 6000 Then
    pgbLoad(i).Visible = False
    pgbLoad(i).Left = 4920
    If pgbLoad(2).Left = 4920 Then
        Cntr = Cntr + 1
    End If
End If
Next i
For i = 0 To 2
    If pgbLoad(i).Left = 5040 Then
        pgbLoad(i).Visible = True
    End If
Next i

For i = 0 To 2
    pgbLoad(i).Left = pgbLoad(i).Left + 30
Next i
Load Cntr
End Sub

Private Sub Load(Cntr As Integer)
On Error GoTo ErrorMsg
'conditions on loading processes.
If Cntr = 1 Then
    tmrLoad.Enabled = False
    lblStatus.Caption = "setting up connection..."
    If DbShow = False Then
        If ReadINI("Preferences", "Default DB") = "0" Then
            frmDb.Show 1
            DbShow = True
        Else
            s = ReadINI("Last Path", "Path")
        End If
    End If
    OpenCon s
    tmrLoad.Enabled = True
ElseIf Cntr = 2 Then
    lblStatus.Caption = "reading system preferences..."
    If ReadINI("Preferences", "Late Task") = "1" Then
        RunSql "Select * from tblSchedules"
        With Rs
            While Not .EOF = True
                If DateDiff("d", .Fields!sched_date, Date) > 0 Then
                    .Delete
                End If
                .MoveNext
            Wend
        End With
    End If
ElseIf Cntr = 3 Then
    lblStatus.Caption = "scanning on user accounts..."
    'searching for recent accounts
    If UserTest = False Then
        RunSql "SELECT * FROM tblAccountProfile"
        With Rs
            If .RecordCount = 0 Then
                tmrLoad.Enabled = False
                MsgBox "No registered account found. Please add an account", vbExclamation
                frmAccountProfile.Show 1
                tmrLoad.Enabled = True
            End If
        End With
   
        RunSql "Select * from tblAccountSecurity"
        With Rs
            If .RecordCount = 0 Then
                tmrLoad.Enabled = False
                MsgBox "No Security Account found. Please set a security account", vbExclamation
                frmAcntManage.Show 1
                tmrLoad.Enabled = True
            End If
        End With
        UserTest = True
    End If
ElseIf Cntr = 4 Then
    tmrLoad.Enabled = False
    lblStatus.Caption = "scanning for system detections..."
    mdiMain.cmdWarning.Caption = Warnings & " W&arnings"
    mdiMain.cmdReminders.Caption = Notifications & " N&otification"
    tmrLoad.Enabled = True
ElseIf Cntr = 5 Then
    lblStatus.Caption = "executing system startup..."
ElseIf Cntr = 6 Then
    Screen.MousePointer = 11
    mdiMain.Show
    Unload Me
    frmLogin.Show 1
End If
Exit Sub
'if unexpected error will occur
ErrorMsg:
    MsgBox "The system has failed to initialized.  Please contact the developer for more information. " & vbNewLine & vbNewLine & _
            "Error Status: " & lblStatus & vbNewLine & "System Error: " & Err.Description, vbCritical
    End
End Sub

