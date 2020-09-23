VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmNotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame freNotes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   495
      Begin VB.TextBox txtUpdates 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   120
         Width           =   6375
      End
   End
   Begin VB.Frame freNotes 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3960
      TabIndex        =   9
      Top             =   120
      Width           =   495
      Begin VB.TextBox txtMemo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2835
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   120
         Width           =   6375
      End
   End
   Begin ComctlLib.TabStrip tabNotes 
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6165
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   2
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Memo"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Updates"
            Key             =   ""
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
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdViewLog 
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&View Log"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmNotes.frx":038A
      Picture         =   "frmNotes.frx":0564
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   4560
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
      MouseIcon       =   "frmNotes.frx":08B8
      Picture         =   "frmNotes.frx":0A92
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   4560
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
      MouseIcon       =   "frmNotes.frx":0DE6
      Picture         =   "frmNotes.frx":0FC0
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click Save to update notes."
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
      Left            =   540
      TabIndex        =   4
      Top             =   4680
      Width           =   1980
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   0
      Picture         =   "frmNotes.frx":1314
      Top             =   4560
      Width           =   480
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View and manage system Logs"
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
      Width           =   1950
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTES"
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
      Width           =   585
   End
   Begin VB.Image img 
      Height          =   480
      Left            =   120
      Picture         =   "frmNotes.frx":1BDE
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Text As String
Dim Output As String
Dim MemoSvd As Boolean
Dim UpdateSvd As Boolean
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Select Case tabNotes.SelectedItem.Index
    Case 1
        SaveMemo
    Case 2
        SaveUpdate
End Select
MsgBox "Notes has been successfully saved", vbInformation
End Sub

Private Sub SaveMemo()
Screen.MousePointer = 11
Open App.Path & "\Memo.log" For Output As #1
Print #1, txtMemo.Text
Close #1
MemoSvd = True
Screen.MousePointer = 0
End Sub

Private Sub SaveUpdate()
Screen.MousePointer = 11
Open App.Path & "\Updates.log" For Output As #1
Print #1, txtUpdates.Text
Close #1
UpdateSvd = True
Screen.MousePointer = 0
End Sub

Private Sub cmdViewLog_Click()
Select Case tabNotes.SelectedItem.Index
    Case 1
        PathToDoc = App.Path & "\Memo.log"
    Case 2
        PathToDoc = App.Path & "\Updates.log"
End Select
ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 8
End Sub

Private Sub Form_Activate()
MemoSvd = True
UpdateSvd = True
End Sub

Private Sub Form_Load()

For i = 2 To freNotes.UBound
    freNotes(1).Height = tabNotes.Height - 420
    freNotes(1).Width = tabNotes.Width - 120
    freNotes(1).Top = tabNotes.Top + 350
    freNotes(1).Left = tabNotes.Left + 40
    freNotes(i).Move _
        freNotes(1).Left, _
        freNotes(1).Top, _
        freNotes(1).Width, _
        freNotes(1).Height
    freNotes(i).Visible = False
Next i

Output = Empty
Text = Empty

Open App.Path & "\Memo.log" For Input As #1
While Not EOF(1) = True
    Line Input #1, Text
    Output = Output & Text & vbCrLf
Wend
txtMemo.Text = Output
Close #1
        
Output = Empty
Text = Empty

Open App.Path & "\Updates.log" For Input As #1
While Not EOF(1) = True
    Line Input #1, Text
    Output = Output & Text & vbCrLf
Wend
txtUpdates.Text = Output
Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MemoSvd = False Then
    x = MsgBox("Would you like to save changes on your memo?", vbExclamation + vbYesNoCancel)
    If x = vbYes Then
        SaveMemo
    ElseIf x = vbCancel Then
        Cancel = 1
    End If
End If

If UpdateSvd = False Then
    x = MsgBox("Would you like to save changes on system updates?", vbExclamation + vbYesNoCancel)
    If x = vbYes Then
        SaveUpdate
    ElseIf x = vbCancel Then
        Cancel = 1
    End If
End If

End Sub

Private Sub tabNotes_Click()

For i = 1 To tabNotes.Tabs.Count
    If freNotes(i).Index = tabNotes.SelectedItem.Index Then
        freNotes(i).Visible = True
    Else
        freNotes(i).Visible = False
    End If
Next i
End Sub

Private Sub txtMemo_Change()
MemoSvd = False
End Sub

Private Sub txtMemo_GotFocus()
txtMemo.SelStart = Len(txtMemo)
End Sub

Private Sub txtUpdates_GotFocus()
txtUpdates.SelStart = Len(txtUpdates)
End Sub

Private Sub txtUpdates_Change()
UpdateSvd = False
End Sub
