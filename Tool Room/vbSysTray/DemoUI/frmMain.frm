VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Systray Demo"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDisplay 
      Caption         =   "Hide Systray Icon"
      Height          =   465
      Left            =   1410
      TabIndex        =   1
      Top             =   1740
      Width           =   1725
   End
   Begin VB.CommandButton cmdAnimate 
      Caption         =   "Start Animation"
      Height          =   465
      Left            =   1410
      TabIndex        =   0
      Top             =   2310
      Width           =   1725
   End
   Begin VB.Timer tmrAnim 
      Left            =   2790
      Top             =   1260
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1620
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0326
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":095A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C74
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":222A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2544
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":285E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B78
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E92
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjSysTray As vbSysTray.SysTray
Attribute mobjSysTray.VB_VarHelpID = -1

Private Sub cmdAnimate_Click()
    Annimate
End Sub

Private Sub cmdDisplay_Click()
    mobjSysTray.Visible = Not mobjSysTray.Visible
    If mobjSysTray.Visible Then
        cmdDisplay.Caption = "Hide Systray Icon"
        cmdAnimate.Enabled = True
    Else
        cmdDisplay.Caption = "Show Systray Icon"
        cmdAnimate.Enabled = False
    End If
End Sub

Private Sub Form_Load()
    tmrAnim.Enabled = False
    tmrAnim.Interval = 200
    Set mobjSysTray = New vbSysTray.SysTray
    Set mobjSysTray.ImageList = ImageList1
    With mobjSysTray
        .Icon = 1
        .Visible = True
        .Menu.Add "Show", "SHOW"
        .Menu.Add "Hide", "HIDE"
        .Menu.Add "-"
        .Menu.Add "Annimate", "ANNIMATE", , , False
        .Menu.Add "-"
        .Menu.Add "Exit Application", "EXIT"
        .EnableMenu = True
    End With
    ShowForm
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu And mobjSysTray.Visible Then
        HideForm
        Cancel = True
    Else
        mobjSysTray.Visible = False
        Set mobjSysTray = Nothing
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        HideForm
    End If
End Sub

Private Sub tmrAnim_Timer()
Dim lngIcon As Long
    lngIcon = mobjSysTray.Icon + 1
    If lngIcon = 17 Then
        lngIcon = 1
    End If
    mobjSysTray.Icon = lngIcon
End Sub

Private Sub mobjSysTray_DoubleClick(ByVal Button As MouseButtonConstants)
    ShowForm
End Sub

Private Sub mobjSysTray_MenuClick(Item As vbSysTray.MenuItem)
    Select Case Item.Key
        Case "SHOW"
            ShowForm
        Case "HIDE"
            HideForm
        Case "ANNIMATE"
            Annimate
        Case "EXIT"
            Unload Me
    End Select
End Sub

Private Sub ShowForm()
    With mobjSysTray.Menu
        .Item("SHOW").Enabled = False
        .Item("HIDE").Enabled = True
    End With
    Me.WindowState = vbNormal
    Me.Show
End Sub

Private Sub HideForm()
    With mobjSysTray.Menu
        .Item("SHOW").Enabled = True
        .Item("HIDE").Enabled = False
    End With
    Me.Hide
End Sub

Private Sub Annimate()
    tmrAnim.Enabled = Not tmrAnim.Enabled
    mobjSysTray.Menu.Item("ANNIMATE").Checked = tmrAnim.Enabled
    If tmrAnim.Enabled Then
        cmdAnimate.Caption = "Stop Animation"
    Else
        cmdAnimate.Caption = "Start Animation"
    End If
End Sub
