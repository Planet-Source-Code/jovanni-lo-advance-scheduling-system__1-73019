VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3D800911-77E3-43DE-82EA-7FC87C713180}#1.1#0"; "cPopMenu6.ocx"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   BackColor       =   &H8000000C&
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   9795
   Icon            =   "mdiMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "mdiMain.frx":038A
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbrMenu 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   1111
      ButtonWidth     =   1667
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Items"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Transaction"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reports"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Accounts"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Settings"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Database"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Help"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exit"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList lstSbr 
      Left            =   3120
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiMain.frx":1A1C3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrDate 
      Interval        =   1000
      Left            =   2280
      Top             =   1440
   End
   Begin VB.PictureBox picSideBar 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8355
      Left            =   0
      ScaleHeight     =   8355
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   630
      Width           =   1920
      Begin VB.CommandButton cmdHide 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   13
         Top             =   6840
         Width           =   255
      End
      Begin CtrlLine.ctrlLiner ctrlLiner5 
         Height          =   30
         Left            =   0
         TabIndex        =   2
         Top             =   3480
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   53
      End
      Begin CtrlLine.ctrlLiner ctrlLiner3 
         Height          =   30
         Left            =   0
         TabIndex        =   3
         Top             =   4320
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   53
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   0
         TabIndex        =   4
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   53
      End
      Begin CtrlLine.ctrlLiner ctrlLiner1 
         Height          =   30
         Left            =   0
         TabIndex        =   5
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   53
      End
      Begin ctrlButton.ThemedButton cmdScheduler 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   5040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Scheduler"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1A55D
         Picture         =   "mdiMain.frx":1A737
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdNotes 
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   4560
         Width           =   1455
         _ExtentX        =   2566
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
         MouseIcon       =   "mdiMain.frx":1AA8B
         Picture         =   "mdiMain.frx":1AC65
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdLogout 
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Log-out"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1AFB9
         Picture         =   "mdiMain.frx":1B193
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdExplore 
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   5520
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "E&xplorer"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1B4E7
         Picture         =   "mdiMain.frx":1B6C1
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin CtrlLine.ctrlLiner ctrlLiner4 
         Height          =   30
         Left            =   0
         TabIndex        =   18
         Top             =   6960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   53
      End
      Begin CtrlLine.ctrlLiner ctrlLiner6 
         Height          =   30
         Left            =   0
         TabIndex        =   19
         Top             =   6120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   53
      End
      Begin ctrlButton.ThemedButton cmdWarning 
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   7200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1BA15
         Picture         =   "mdiMain.frx":1BBEF
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdStatus 
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Stat&us"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1BF43
         Picture         =   "mdiMain.frx":1C11D
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdRegister 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Register"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1C471
         Picture         =   "mdiMain.frx":1C64B
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdTransactions 
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "&Transactions"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1C99F
         Picture         =   "mdiMain.frx":1CB79
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdReminders 
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   7680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "mdiMain.frx":1CECD
         Picture         =   "mdiMain.frx":1D0A7
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LOGS"
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
         Index           =   5
         Left            =   720
         TabIndex        =   21
         Top             =   6240
         Width           =   480
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   2
         Left            =   120
         Picture         =   "mdiMain.frx":1D3FB
         Top             =   6240
         Width           =   480
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System Logs"
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
         Index           =   4
         Left            =   720
         TabIndex        =   20
         Top             =   6600
         Width           =   795
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Features"
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
         Index           =   3
         Left            =   720
         TabIndex        =   9
         Top             =   3960
         Width           =   540
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   1
         Left            =   120
         Picture         =   "mdiMain.frx":1E03F
         Top             =   3600
         Width           =   480
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SYSTEM"
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
         Index           =   2
         Left            =   720
         TabIndex        =   8
         Top             =   3600
         Width           =   750
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TASK"
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
         Index           =   0
         Left            =   720
         TabIndex        =   7
         Top             =   120
         Width           =   495
      End
      Begin VB.Image imgIcon 
         Height          =   480
         Index           =   0
         Left            =   120
         Picture         =   "mdiMain.frx":1EC83
         Top             =   120
         Width           =   480
      End
      Begin VB.Label lblGwave 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quick access"
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
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Width           =   795
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         FillColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   0
         Top             =   0
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         FillColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   -240
         Top             =   3480
         Width           =   2295
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000004&
         FillColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   0
         Top             =   6120
         Width           =   2295
      End
   End
   Begin ComctlLib.StatusBar sbrDetails 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   8985
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin cPopMenu6.PopMenu cPopMnu 
      Left            =   3120
      Top             =   1440
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
   End
   Begin ComctlLib.ImageList lstIcons 
      Left            =   2760
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   52
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":1F8C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":1FC19
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":1FF6B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":202BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2060F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":20961
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":20CB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":21005
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":21357
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":216A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":219FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":21D4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2209F
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":223F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":22743
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":22A95
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":22DE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":23139
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2348B
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":237DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":23B2F
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":23E81
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":241D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":24525
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":24877
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":24BC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":24F1B
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2526D
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":255BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":25911
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":25C63
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":25FB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":26307
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":26659
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":269AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":26CFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2704F
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":273A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":276F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":27A45
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":27D97
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":280E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2843B
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2878D
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":28ADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":28E31
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":29183
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":294D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":29827
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":29B79
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":29ECB
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2A21D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList lstMenu 
      Left            =   4080
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2A56F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2B1C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2BE13
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2CA65
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2D6B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2E309
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2EF5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdiMain.frx":2FBAD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSettings 
         Caption         =   "System Settings"
         Shortcut        =   {F11}
      End
      Begin VB.Menu lSettings 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReports 
         Caption         =   "Manage &Reports"
         Begin VB.Menu mnuRpt 
            Caption         =   "&Categorized"
            Index           =   0
            Shortcut        =   {F6}
         End
         Begin VB.Menu mnuRpt 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu mnuRpt 
            Caption         =   "A&dvance Query"
            Index           =   2
         End
      End
      Begin VB.Menu lineManage 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccess 
         Caption         =   "Security &Access"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuLock 
         Caption         =   "&Lock System"
         Shortcut        =   {F3}
      End
      Begin VB.Menu lTrans 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuTransactions 
      Caption         =   "&Transactions"
      Begin VB.Menu mnuTrans 
         Caption         =   "&Borrow Items"
         Index           =   0
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "&Reservation"
         Index           =   1
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "R&eturn Items"
         Index           =   3
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "&Cancel Reservation"
         Index           =   4
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuTrans 
         Caption         =   "Manage Transactions"
         Index           =   6
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuDb 
         Caption         =   "&Load"
         Index           =   0
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuDb 
         Caption         =   "&Refresh"
         Index           =   1
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDb 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuDb 
         Caption         =   "&Back up"
         Index           =   3
      End
      Begin VB.Menu mnuDb 
         Caption         =   "&View"
         Index           =   4
      End
      Begin VB.Menu mnuDb 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuDb 
         Caption         =   "&Clear Data"
         Index           =   6
      End
   End
   Begin VB.Menu mnuItem 
      Caption         =   "Items"
      Begin VB.Menu mnuDbItems 
         Caption         =   "Manage Items"
         Begin VB.Menu mnuItems 
            Caption         =   "&Add"
            Index           =   0
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuItems 
            Caption         =   "&Edit"
            Index           =   1
         End
         Begin VB.Menu mnuItems 
            Caption         =   "&Delete"
            Index           =   2
         End
         Begin VB.Menu mnuItems 
            Caption         =   "-"
            Index           =   3
         End
         Begin VB.Menu mnuItems 
            Caption         =   "&Search"
            Index           =   4
            Shortcut        =   ^F
         End
      End
      Begin VB.Menu lneDbItems 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "Condition and &Status"
      End
      Begin VB.Menu mnuReg 
         Caption         =   "Item Registration"
      End
   End
   Begin VB.Menu mnuAccounts 
      Caption         =   "Accounts"
      Begin VB.Menu mnuProfiles 
         Caption         =   "Account Profiles"
      End
      Begin VB.Menu mnuSecurity 
         Caption         =   "Security Manager"
      End
      Begin VB.Menu lSecu 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUsers 
         Caption         =   "Client Profiles"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuManual 
         Caption         =   "&User Manual"
         Shortcut        =   {F1}
      End
      Begin VB.Menu lineUser 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About The system..."
      End
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mobjSysTray As SysTray
Attribute mobjSysTray.VB_VarHelpID = -1

Private Sub cmdExplore_Click()
Shell "explorer", vbNormalFocus
End Sub

Private Sub cmdNotes_Click()
frmNotes.Show 1
End Sub

Private Sub cmdRegister_Click()
mnuReg_Click
End Sub

Private Sub cmdReminders_Click()
Screen.MousePointer = 11
Notifications 2
frmWarnings.Show 1
End Sub

Private Sub cmdScheduler_Click()
If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
frmScheduler.Show 1
End Sub

Private Sub cmdStatus_Click()
mnuStatus_Click
End Sub

Private Sub cmdTransactions_Click()
If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
frmTransactions.Show 1
End Sub

Private Sub cmdWarning_Click()
Screen.MousePointer = 11
Warnings 1
frmWarnings.Show 1
End Sub

Private Sub MDIForm_Activate()
With mobjSysTray
    .Menu.Item("log").Caption = cmdLogout.Caption
    .Menu.Item("log").Enabled = True
    .ToolTipText = "User name: " & UserNme & vbNewLine & _
                    "Level: " & UserLvl & vbNewLine & _
                    "ID: " & UserId & vbNewLine & _
                    Format(Date, "mm/dd/yyyy")
End With
ExecLogs
End Sub

Private Sub ShowForm()
    Me.WindowState = vbMaximized
    Me.Show
    With mobjSysTray.Menu
        .Item("show").Enabled = False
        .Item("hide").Enabled = True
    End With
End Sub

Private Sub HideForm()
    Me.Hide
    With mobjSysTray.Menu
        .Item("show").Enabled = True
        .Item("hide").Enabled = False
    End With
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState = vbMinimized Then
    HideForm
End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Me.WindowState = vbMinimized
Cancel = 1
End Sub

Private Sub mnuReg_Click()
If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
frmRegister.Show 1
End Sub

Private Sub mnuRpt_Click(Index As Integer)
Select Case Index
    Case 0
        frmRptCategories.Show 1
End Select
End Sub

Private Sub mnuStatus_Click()
If UserLimit(UserLvl, "User") = True Then Exit Sub
frmStatus.Show 1
End Sub

Private Sub mobjSysTray_DoubleClick(ByVal Button As MouseButtonConstants)
ShowForm
End Sub

Private Sub mobjSysTray_MenuClick(Item As MenuItem)
On Error Resume Next
    Select Case Item.Key
        Case "show"
            ShowForm
        Case "manual"
            PathToDoc = App.Path & "\help.chm"
            ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5
        Case "about"
            mnuAbout_Click
        Case "log"
            cmdLogOut_Click
        Case "hide"
            HideForm
    End Select
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuAccess_Click()
cmdLogOut_Click
End Sub

Private Sub mnuClear_Click()

End Sub

Private Sub mnuDb_Click(Index As Integer)
On Error GoTo ExecErr
DbMgt = True
If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
Select Case Index
    Case 0
        frmDb.Show 1
        If s <> Empty Then
            Con.Close
            OpenCon s
            MsgBox "Loaded new database. Please log-in to start.", vbInformation
            cmdLogOut_Click
        End If
    Case 1
        i = MsgBox("Compacting your database may take several minutes to finish. Would you like to continue?", vbExclamation + vbYesNo)
        If i = vbYes Then
            Con.Close
            Screen.MousePointer = 11
            s = ReadINI("Last Path", "Path")
            CompactDB s
            Screen.MousePointer = 0
            MsgBox "Database compacting was successfully completed.", vbInformation
            OpenCon s
        End If
    Case 3
        Dim Folder, DbName As String
        If MsgBox("Your about to backup your current database. This may take a while to complete. " & vbNewLine & vbNewLine & _
                    "Do you want to proceed?", vbExclamation + vbYesNo) = vbNo Then
            Exit Sub
        End If
        s = StrBox("Please input the folder name", imgIcon(1), , Format(Now, "mm-dd-yyyy(hns)"), "back up")
        If s = Empty Then Exit Sub
        Screen.MousePointer = 11
        Folder = App.Path & "\Database\" & s & "\"
        MkDir Folder
        Con.Close
        Screen.MousePointer = 0
        FileCopy ReadINI("Last Path", "Path"), Folder & "Tool.mdb"
        Screen.MousePointer = 0
        MsgBox "Successfully created a backup of your database. " & vbNewLine & vbNewLine & _
                "Path: " & Folder & "Tool.mdb", vbInformation
        OpenCon ReadINI("Last Path", "Path")
    Case 4
        frmDb.Show 1
        PathToDoc = s
        ShellExecute 0, "open", PathToDoc, vbNullString, vbNullString, 5
    Case 6
        x = MsgBox("Your about to clear all records from your database. Would you like to backup it first?", vbExclamation + vbYesNoCancel)
        If x <> vbCancel Then
            If x = vbYes Then
                mnuDb_Click 3
            End If
            If MsgBox("The system will now clear all records from your database. Click Yes to continue.", vbYesNo + vbExclamation) = vbYes Then
                n = 0
                SubSql "select * from tblItemList"
                With SubRs
                    While Not .EOF = True
                        If RunSql("Delete * from tblItemList where item_id = '" & SubRs.Fields!item_id & "'") = False Then
                            n = n + 1
                        End If
                        .MoveNext
                    Wend
                End With
                MsgBox n & " record(s) were delete on items.", vbInformation
                cmdLogOut_Click
            End If
        End If
End Select
DbMgt = False
Exit Sub

ExecErr:
    MsgBox "Process cannot continue because an error has occured." & vbNewLine & vbNewLine & _
            "System Error: " & Err.Description, vbCritical
    OpenCon ReadINI("Last Path", "Path")
DbMgt = False
End Sub

Public Sub mnuExit_Click()
mobjSysTray.Visible = False
Set mobjSysTray = Nothing
End
End Sub

Private Sub mnuItems_Click(Index As Integer)
With frmItems
    Select Case Index
        Case 0
            .ExecButtons 1
        Case 1
            .ExecButtons 2
        Case 2
            .ExecButtons 3
        Case 4
            .ExecButtons 4
    End Select
End With
End Sub

Private Sub mnuList_Click(Index As Integer)

End Sub

Private Sub mnuLock_Click()
If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
Me.Hide
frmLock.Show
End Sub

Private Sub mnuProfiles_Click()
frmAccountProfile.Show 1
End Sub

Private Sub mnuSecurity_Click()
frmAcntManage.Show 1
End Sub

Private Sub mnuSettings_Click()
If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
frmSettings.Show 1
End Sub

Private Sub mnuTrans_Click(Index As Integer)
Select Case Index
    Case 0
        frmTransView.ExecButtons 2
    Case 1
        frmTransView.ExecButtons 3
    Case 3
        frmTransactions.tabTrans.Tabs(3).Selected = True
        frmTransactions.Show 1
    Case 4
        frmTransactions.tabTrans.Tabs(4).Selected = True
        frmTransactions.Show 1
    Case 6
        frmTransView.ExecButtons 4
End Select
End Sub

Private Sub mnuUsers_Click()
frmClientProfiles.Show 1
End Sub

Private Sub tmrDate_Timer()
sbrDetails.Panels(6).Text = "Today is: " & Format(Now, "dddd, mmm dd, yyyy") & "  "
sbrDetails.Panels(7).Text = "Time: " & Format(Time, "hh:mm:ss")
End Sub
Private Sub cmdHide_Click()
If cmdHide.Caption = "<<" Then
    picSideBar.Width = 240
    cmdHide.Caption = ">>"
    imgIcon(0).Visible = False
    imgIcon(1).Visible = False
    imgIcon(2).Visible = False
Else
    picSideBar.Width = 1920
    cmdHide.Caption = "<<"
    imgIcon(0).Visible = True
    imgIcon(1).Visible = True
    imgIcon(2).Visible = True
End If
End Sub

Private Sub MDIForm_Load()
Me.Caption = App.ProductName & " v" & App.Major & "." & App.Minor & "." & App.Revision

Set mobjSysTray = New SysTray
Set mobjSysTray.ImageList = lstSbr
With mobjSysTray
    .Icon = 1
    .Visible = True
    .Menu.Add "Restore Window", "show"
    .Menu.Add "Hide Window", "hide"
    .Menu.Add "-"
    .Menu.Add "Log-in", "log"
    .Menu.Add "-"
    .Menu.Add "About the System", "about"
    .Menu.Add "User's Manual", "manual"
    .EnableMenu = True
    .ToolTipText = Me.Caption
    .Menu.Item("show").Enabled = False
End With

With tbrMenu
    .ImageList = lstMenu
    For i = 1 To lstMenu.ListImages.Count
        .Buttons(i).Image = i
        n = n + 2
    Next i
End With

With cPopMnu
    .MenuBackgroundColor = RGB(255, 255, 255)
    .ImageList = lstIcons
    .SubClassMenu Me
    .ItemIcon("mnuReports") = lstIcons.ListImages(49).Index - 1
    .ItemIcon("mnuSettings") = lstIcons.ListImages(18).Index - 1
    .ItemIcon("mnuDatabase") = lstIcons.ListImages(30).Index - 1
    .ItemIcon("mnuAccess") = lstIcons.ListImages(32).Index - 1
    .ItemIcon("mnuExit") = lstIcons.ListImages(17).Index - 1
    .ItemIcon("mnuTrans(0)") = lstIcons.ListImages(16).Index - 1
    .ItemIcon("mnuTrans(1)") = lstIcons.ListImages(24).Index - 1
    .ItemIcon("mnuTrans(3)") = lstIcons.ListImages(8).Index - 1
    .ItemIcon("mnuTrans(4)") = lstIcons.ListImages(41).Index - 1
    .ItemIcon("mnuTrans(6)") = lstIcons.ListImages(43).Index - 1
    .ItemIcon("mnuDb(0)") = lstIcons.ListImages(3).Index - 1
    .ItemIcon("mnuDb(1)") = lstIcons.ListImages(7).Index - 1
    .ItemIcon("mnuDb(3)") = lstIcons.ListImages(8).Index - 1
    .ItemIcon("mnuDb(4)") = lstIcons.ListImages(42).Index - 1
    .ItemIcon("mnuDb(6)") = lstIcons.ListImages(11).Index - 1
    .ItemIcon("mnuExit") = lstIcons.ListImages(17).Index - 1
    .ItemIcon("mnuItems(0)") = lstIcons.ListImages(2).Index - 1
    .ItemIcon("mnuItems(1)") = lstIcons.ListImages(44).Index - 1
    .ItemIcon("mnuItems(2)") = lstIcons.ListImages(9).Index - 1
    .ItemIcon("mnuItems(4)") = lstIcons.ListImages(27).Index - 1
    .ItemIcon("mnuStatus") = lstIcons.ListImages(1).Index - 1
    .ItemIcon("mnuReg") = lstIcons.ListImages(52).Index - 1
    .ItemIcon("mnuProfiles") = lstIcons.ListImages(50).Index - 1
    .ItemIcon("mnuSecurity") = lstIcons.ListImages(34).Index - 1
    .ItemIcon("mnuUsers") = lstIcons.ListImages(15).Index - 1
    .ItemIcon("mnuManual") = lstIcons.ListImages(29).Index - 1
    .ItemIcon("mnuAbout") = lstIcons.ListImages(25).Index - 1
    .ItemIcon("mnuLock") = lstIcons.ListImages(5).Index - 1
    .ItemIcon("mnuRpt(0)") = lstIcons.ListImages(38).Index - 1
    .ItemIcon("mnuRpt(2)") = lstIcons.ListImages(13).Index - 1
End With

sbrDetails.Panels(6).Text = "Today is: " & Format(Now, "dddd, mmm dd, yyyy")
sbrDetails.Panels(7).Text = "Time: " & Format(Time, "hh:mm:ss")
End Sub

Private Sub picSideBar_Resize()
cmdHide.Top = picSideBar.Height - cmdHide.Height
cmdHide.Left = picSideBar.Width - cmdHide.Width
End Sub
Private Sub cmdLogOut_Click()
ClrAccess
frmLogin.Show 1
End Sub

Private Sub ClrAccess()
UserLvl = Empty
UserNme = Empty
UserId = Empty
Unload frmItems
cmdLogout.Caption = "&Log-in"
For i = 1 To 7
    sbrDetails.Panels(i).Text = Empty
Next i
mobjSysTray.Menu.Item("log").Enabled = False
End Sub
Private Sub tbrMenu_ButtonClick(ByVal Button As ComctlLib.Button)
Screen.MousePointer = 11
With tbrMenu
    Select Case Button.Index
        Case 1
            If UserLimit(UserLvl, "User") = True Then Exit Sub
            Unload frmTransView
            frmItems.Show
        Case 2
            If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
            Unload frmItems
            frmTransView.Show
        Case 3
            If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
            Screen.MousePointer = 0
            PopupMenu mnuReports, x:=Button.Left, y:=.Top + .Height
        Case 4
            If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
            Screen.MousePointer = 0
            PopupMenu mnuAccounts, x:=Button.Left, y:=.Top + .Height
        Case 5
            mnuSettings_Click
        Case 6
            If UserLimit(UserLvl, "Administrator") = True Then Exit Sub
            Screen.MousePointer = 0
            PopupMenu mnuDatabase, x:=Button.Left, y:=.Top + .Height
        Case 7
            Screen.MousePointer = 0
            PopupMenu mnuHelp, x:=Button.Left, y:=.Top + .Height
        Case 8
            mnuExit_Click
    End Select
End With
Screen.MousePointer = 0
End Sub
