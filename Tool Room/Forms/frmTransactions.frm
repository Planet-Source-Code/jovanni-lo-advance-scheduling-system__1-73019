VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmTransactions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   750
   ClientWidth     =   7950
   Icon            =   "frmTransactions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Image"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2880
      TabIndex        =   128
      Top             =   1800
      Width           =   1455
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000011&
         FillColor       =   &H80000004&
         Height          =   975
         Left            =   120
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image imgProfile 
         Height          =   975
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NO IMAGE"
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
         Left            =   360
         TabIndex        =   131
         Top             =   600
         Width           =   825
      End
   End
   Begin VB.Frame freTrans 
      BorderStyle     =   0  'None
      Height          =   615
      Index           =   3
      Left            =   2640
      TabIndex        =   90
      Top             =   120
      Width           =   375
      Begin VB.CheckBox chkSvReturn 
         Caption         =   "Save this transaction"
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
         Left            =   2640
         TabIndex        =   111
         Top             =   1800
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         Height          =   1215
         Left            =   240
         TabIndex        =   92
         Top             =   2160
         Width           =   7215
         Begin VB.Label lblRDate 
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
            Left            =   5160
            TabIndex        =   105
            Top             =   240
            Width           =   180
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Return Date:"
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
            Left            =   3840
            TabIndex        =   104
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lblBDate 
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
            Left            =   1680
            TabIndex        =   103
            Top             =   240
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Borrowed Date:"
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
            Index           =   6
            Left            =   240
            TabIndex        =   102
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Index           =   5
            Left            =   3600
            TabIndex        =   101
            Top             =   240
            Width           =   105
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Qty:"
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
            Index           =   7
            Left            =   6240
            TabIndex        =   100
            Top             =   720
            Width           =   345
         End
         Begin VB.Label lblQty 
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
            Left            =   6840
            TabIndex        =   99
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblBrange 
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
            Left            =   1680
            TabIndex        =   98
            Top             =   720
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Time Interval:"
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
            Index           =   8
            Left            =   240
            TabIndex        =   97
            Top             =   720
            Width           =   1200
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Index           =   3
            Left            =   6000
            TabIndex        =   96
            Top             =   720
            Width           =   105
         End
         Begin VB.Label lblBTrans 
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
            Left            =   5160
            TabIndex        =   95
            Top             =   720
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Transaction #:"
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
            Left            =   3840
            TabIndex        =   94
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Index           =   9
            Left            =   3600
            TabIndex        =   93
            Top             =   720
            Width           =   105
         End
      End
      Begin VB.TextBox txtBarSrch 
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
         Left            =   240
         TabIndex        =   91
         Text            =   "Search"
         Top             =   1800
         Width           =   1695
      End
      Begin ctrlButton.ThemedButton cmdReturn 
         Height          =   375
         Left            =   4320
         TabIndex        =   106
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTransactions.frx":038A
      End
      Begin ComctlLib.ListView lvwBorrowed 
         Height          =   1455
         Left            =   240
         TabIndex        =   107
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         SmallIcons      =   "imgBullet"
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
            Text            =   "Item ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Qty"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin ctrlButton.ThemedButton cmdCReturn 
         Height          =   375
         Left            =   4320
         TabIndex        =   108
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTransactions.frx":0564
      End
      Begin ComctlLib.ListView lvwReturned 
         Height          =   1455
         Left            =   4800
         TabIndex        =   109
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
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
         NumItems        =   4
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
            Text            =   "Item ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Qty"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Index           =   20
         Left            =   4560
         TabIndex        =   113
         Top             =   1800
         Width           =   105
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Index           =   10
         Left            =   2400
         TabIndex        =   112
         Top             =   1800
         Width           =   105
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   1920
         Picture         =   "frmTransactions.frx":073E
         Top             =   1680
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   4680
         Picture         =   "frmTransactions.frx":1008
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click '>' to Return an Item."
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
         Left            =   5160
         TabIndex        =   110
         Top             =   1800
         Width           =   1935
      End
   End
   Begin VB.Frame freTrans 
      BorderStyle     =   0  'None
      Height          =   3255
      Index           =   5
      Left            =   240
      TabIndex        =   66
      Top             =   3840
      Width           =   7455
      Begin ctrlButton.ThemedButton cmdDelSelected 
         Height          =   375
         Left            =   5760
         TabIndex        =   126
         Top             =   1560
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Caption         =   "Delete Selected"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTransactions.frx":18D2
         Picture         =   "frmTransactions.frx":1AAC
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin VB.Frame Frame8 
         Height          =   1335
         Left            =   3600
         TabIndex        =   118
         Top             =   120
         Width           =   3855
         Begin VB.Label lblTransCategory 
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
            Left            =   1560
            TabIndex        =   125
            Top             =   960
            Width           =   180
         End
         Begin VB.Label lblTransDate 
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
            Left            =   1560
            TabIndex        =   124
            Top             =   600
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "User's Name:"
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
            Index           =   26
            Left            =   240
            TabIndex        =   123
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Date:"
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
            Index           =   25
            Left            =   240
            TabIndex        =   122
            Top             =   600
            Width           =   450
         End
         Begin VB.Label Label10 
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
            Index           =   24
            Left            =   240
            TabIndex        =   121
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label lblTransName 
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
            Left            =   1560
            TabIndex        =   120
            Top             =   240
            Width           =   180
         End
      End
      Begin VB.ListBox lstTrans 
         Height          =   1410
         ItemData        =   "frmTransactions.frx":1E00
         Left            =   240
         List            =   "frmTransactions.frx":1E02
         Style           =   1  'Checkbox
         TabIndex        =   117
         Top             =   180
         Width           =   3135
      End
      Begin VB.CheckBox chkTransAll 
         Caption         =   "Show unused transactions only"
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
         Left            =   3000
         TabIndex        =   116
         Top             =   1680
         Width           =   2535
      End
      Begin VB.CheckBox chkTransUsed 
         Caption         =   "Check all unused transactions"
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
         TabIndex        =   114
         Top             =   1680
         Width           =   2415
      End
      Begin ComctlLib.ListView lvwTransView 
         Height          =   1335
         Left            =   240
         TabIndex        =   119
         Top             =   2040
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         SmallIcons      =   "imgBullet"
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
         NumItems        =   0
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Index           =   22
         Left            =   5520
         TabIndex        =   127
         Top             =   1680
         Width           =   105
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Index           =   21
         Left            =   2760
         TabIndex        =   115
         Top             =   1680
         Width           =   105
      End
   End
   Begin VB.Frame freTrans 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   4
      Left            =   3840
      TabIndex        =   13
      Top             =   120
      Width           =   495
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   240
         TabIndex        =   76
         Top             =   2160
         Width           =   7215
         Begin VB.TextBox txtCRemarks 
            Height          =   315
            Left            =   5040
            TabIndex        =   89
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label27 
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
            Left            =   3960
            TabIndex        =   88
            Top             =   240
            Width           =   810
         End
         Begin VB.Label lblTDate 
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
            Left            =   1920
            TabIndex        =   87
            Top             =   240
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Transaction Date:"
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
            Index           =   19
            Left            =   240
            TabIndex        =   86
            Top             =   240
            Width           =   1500
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Index           =   18
            Left            =   3720
            TabIndex        =   85
            Top             =   240
            Width           =   105
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Qty:"
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
            Index           =   17
            Left            =   3960
            TabIndex        =   84
            Top             =   720
            Width           =   345
         End
         Begin VB.Label lblCQty 
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
            Left            =   4440
            TabIndex        =   83
            Top             =   720
            Width           =   180
         End
         Begin VB.Label lblReserveDate 
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
            Left            =   1920
            TabIndex        =   82
            Top             =   720
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Reserved Date:"
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
            Index           =   16
            Left            =   240
            TabIndex        =   81
            Top             =   720
            Width           =   1305
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Index           =   15
            Left            =   4800
            TabIndex        =   80
            Top             =   720
            Width           =   105
         End
         Begin VB.Label lblCTrans 
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
            Left            =   6360
            TabIndex        =   79
            Top             =   720
            Width           =   180
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Transaction #:"
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
            Index           =   13
            Left            =   5040
            TabIndex        =   78
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
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
            Index           =   12
            Left            =   3720
            TabIndex        =   77
            Top             =   720
            Width           =   105
         End
      End
      Begin VB.CheckBox chkSvCancel 
         Caption         =   "Save this transaction"
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
         Left            =   2640
         TabIndex        =   73
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtResSrch 
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
         Left            =   240
         TabIndex        =   70
         Text            =   "Search"
         Top             =   1800
         Width           =   1695
      End
      Begin ComctlLib.ListView lvwReserved 
         Height          =   1455
         Left            =   240
         TabIndex        =   67
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         SmallIcons      =   "imgBullet"
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
            Text            =   "Item ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Qty"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin ctrlButton.ThemedButton cmdCancel 
         Height          =   375
         Left            =   4320
         TabIndex        =   68
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTransactions.frx":1E04
      End
      Begin ctrlButton.ThemedButton cmdCCancel 
         Height          =   375
         Left            =   4320
         TabIndex        =   69
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTransactions.frx":1FDE
      End
      Begin ComctlLib.ListView lvwCanceled 
         Height          =   1455
         Left            =   4800
         TabIndex        =   72
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
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
            Text            =   "Item ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Qty"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Remarks"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   4680
         Picture         =   "frmTransactions.frx":21B8
         Top             =   1680
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Click '>' to Cancel Reservation"
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
         Left            =   5160
         TabIndex        =   75
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Index           =   14
         Left            =   4560
         TabIndex        =   74
         Top             =   1800
         Width           =   105
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         Index           =   11
         Left            =   2400
         TabIndex        =   71
         Top             =   1800
         Width           =   105
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   1920
         Picture         =   "frmTransactions.frx":2A82
         Top             =   1680
         Width           =   480
      End
   End
   Begin VB.Frame freTrans 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   2
      Left            =   6000
      TabIndex        =   12
      Top             =   120
      Width           =   375
      Begin VB.TextBox txtRItemSrch 
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
         Left            =   240
         TabIndex        =   65
         Text            =   "Search"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CheckBox chkTest 
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   60
         Top             =   2520
         Width           =   255
      End
      Begin VB.CheckBox chkTest 
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   59
         Top             =   3000
         Width           =   255
      End
      Begin VB.Frame Frame6 
         Height          =   1335
         Left            =   3240
         TabIndex        =   48
         Top             =   2040
         Width           =   4215
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Item Details"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   0
            Width           =   1050
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Available Qty:"
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
            Left            =   2520
            TabIndex        =   56
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label lblRLocation 
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
            Left            =   1200
            TabIndex        =   55
            Top             =   360
            Width           =   180
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Location:"
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
            TabIndex        =   54
            Top             =   360
            Width           =   765
         End
         Begin VB.Label lblRQty 
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
            Left            =   3840
            TabIndex        =   53
            Top             =   360
            Width           =   180
         End
         Begin VB.Label lblRStatus 
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
            Left            =   1200
            TabIndex        =   52
            Top             =   840
            Width           =   180
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            TabIndex        =   51
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total Count:"
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
            Left            =   2520
            TabIndex        =   50
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lblRCount 
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
            Left            =   3840
            TabIndex        =   49
            Top             =   840
            Width           =   180
         End
      End
      Begin ctrlButton.ThemedButton cmdRAdd 
         Height          =   375
         Left            =   2760
         TabIndex        =   45
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTransactions.frx":334C
      End
      Begin ctrlButton.ThemedButton cmdRRemove 
         Height          =   375
         Left            =   2760
         TabIndex        =   46
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTransactions.frx":3526
      End
      Begin ComctlLib.ListView lvwRItems 
         Height          =   1695
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         SmallIcons      =   "imgBullet"
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Item ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Available"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
      End
      Begin ComctlLib.ListView lvwRView 
         Height          =   1695
         Left            =   3240
         TabIndex        =   58
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2990
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Item ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Qty"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Reserve_date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpRDate 
         Height          =   315
         Left            =   1080
         TabIndex        =   61
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20578307
         CurrentDate     =   40252
      End
      Begin MSComCtl2.DTPicker dtpRTime 
         Height          =   315
         Left            =   1080
         TabIndex        =   62
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   20578306
         CurrentDate     =   40252
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   2280
         Picture         =   "frmTransactions.frx":3700
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Time:"
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
         TabIndex        =   64
         Top             =   3000
         Width           =   465
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Date:"
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
         TabIndex        =   63
         Top             =   2520
         Width           =   450
      End
   End
   Begin VB.Frame freTrans 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   1
      Left            =   7320
      TabIndex        =   11
      Top             =   120
      Width           =   735
      Begin CtrlLine.ctrlLiner ctrlLiner3 
         Height          =   30
         Left            =   1560
         TabIndex        =   39
         Top             =   2640
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin VB.ComboBox cboBGap 
         Height          =   315
         ItemData        =   "frmTransactions.frx":3FCA
         Left            =   1800
         List            =   "frmTransactions.frx":3FDD
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtBTime 
         Height          =   285
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   33
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtBItemSrch 
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
         Left            =   240
         TabIndex        =   27
         Text            =   "Search"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   3240
         TabIndex        =   19
         Top             =   2040
         Width           =   4215
         Begin VB.Label lblBCount 
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
            Left            =   3840
            TabIndex        =   38
            Top             =   840
            Width           =   180
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Total Count:"
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
            Left            =   2520
            TabIndex        =   37
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Status:"
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
            Top             =   840
            Width           =   600
         End
         Begin VB.Label lblBStatus 
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
            Left            =   1200
            TabIndex        =   25
            Top             =   840
            Width           =   180
         End
         Begin VB.Label lblBQty 
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
            Left            =   3840
            TabIndex        =   24
            Top             =   360
            Width           =   180
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Location:"
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
            TabIndex        =   23
            Top             =   360
            Width           =   765
         End
         Begin VB.Label lblBLocation 
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
            Left            =   1200
            TabIndex        =   22
            Top             =   360
            Width           =   180
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Available Qty:"
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
            Left            =   2520
            TabIndex        =   21
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Item Details"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   0
            Width           =   1050
         End
      End
      Begin ComctlLib.ListView lvwBView 
         Height          =   1695
         Left            =   3240
         TabIndex        =   18
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   2990
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
            Text            =   "Item ID"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Qty"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Time Value"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Interval"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin ComctlLib.ListView lvwBItems 
         Height          =   1695
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2990
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         SmallIcons      =   "imgBullet"
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Item ID"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Available"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
      End
      Begin ctrlButton.ThemedButton cmdBAdd 
         Height          =   375
         Left            =   2760
         TabIndex        =   30
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   ">"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTransactions.frx":4000
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdBRemove 
         Height          =   375
         Left            =   2760
         TabIndex        =   43
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Caption         =   "<"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmTransactions.frx":41DA
         PictureSize     =   0
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2280
         Picture         =   "frmTransactions.frx":43B4
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label lblBReturn 
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
         Left            =   1080
         TabIndex        =   36
         Top             =   3000
         Width           =   180
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Return:"
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
         TabIndex        =   35
         Top             =   3000
         Width           =   630
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Time:"
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
         TabIndex        =   32
         Top             =   2520
         Width           =   465
      End
   End
   Begin ComctlLib.TabStrip tabTrans 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   7011
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Borrow Items"
            Key             =   "Borrow"
            Object.Tag             =   "Borrowed"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Reservation"
            Key             =   "Reserve"
            Object.Tag             =   "Reserved"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Return Items"
            Key             =   "Return"
            Object.Tag             =   "Returned"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cancel Reservation"
            Key             =   "Cancel"
            Object.Tag             =   "Canceled"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Manage Transactions"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Client History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   4440
      TabIndex        =   8
      Top             =   960
      Width           =   3375
      Begin VB.TextBox txtHistory 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblTrans 
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
         Left            =   1920
         TabIndex        =   130
         Top             =   1800
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "New Transaction #:"
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
         Left            =   120
         TabIndex        =   129
         Top             =   1800
         Width           =   1620
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Client Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   2655
      Begin VB.Label lblName 
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
         Left            =   1080
         TabIndex        =   17
         Top             =   840
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
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
         TabIndex        =   16
         Top             =   840
         Width           =   525
      End
      Begin VB.Label lblId 
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
         Left            =   1080
         TabIndex        =   15
         Top             =   360
         Width           =   180
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Client #:"
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
         TabIndex        =   14
         Top             =   360
         Width           =   705
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   53
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   6480
      TabIndex        =   29
      Top             =   7440
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
      MouseIcon       =   "frmTransactions.frx":4C7E
      Picture         =   "frmTransactions.frx":4E58
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Height          =   375
      Left            =   3600
      TabIndex        =   40
      Top             =   7440
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
      MouseIcon       =   "frmTransactions.frx":51AC
      Picture         =   "frmTransactions.frx":5386
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdNew 
      Height          =   375
      Left            =   5040
      TabIndex        =   41
      Top             =   7440
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
      MouseIcon       =   "frmTransactions.frx":56DA
      Picture         =   "frmTransactions.frx":58B4
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4215
      Begin ctrlButton.ThemedButton cmdSearch 
         Default         =   -1  'True
         Height          =   375
         Left            =   3240
         TabIndex        =   42
         Top             =   240
         Width           =   375
         _ExtentX        =   661
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
         MouseIcon       =   "frmTransactions.frx":5C08
         Picture         =   "frmTransactions.frx":5DE2
         PictureAlign    =   2
         PictureSize     =   0
      End
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
         Left            =   1800
         TabIndex        =   0
         Text            =   "Search"
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmTransactions.frx":6136
         Left            =   120
         List            =   "frmTransactions.frx":6138
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin ctrlButton.ThemedButton cmdClear 
         Height          =   375
         Left            =   3720
         TabIndex        =   44
         Top             =   240
         Width           =   375
         _ExtentX        =   661
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
         MouseIcon       =   "frmTransactions.frx":613A
         Picture         =   "frmTransactions.frx":6314
         PictureAlign    =   2
         PictureSize     =   0
      End
   End
   Begin ComctlLib.ImageList imgBullet 
      Left            =   4920
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransactions.frx":6668
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTransactions.frx":69BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmTransactions.frx":6D0C
      Top             =   7440
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search for a user profile to process"
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
      TabIndex        =   10
      Top             =   7560
      Width           =   2535
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   120
      Picture         =   "frmTransactions.frx":75D6
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSACTIONS"
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
      Width           =   1455
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tool room main transactions"
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
      Width           =   1740
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000016&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IntSame As Integer, Condition As Boolean, Table As String

Private Sub cboBGap_Click()
txtBTime_Change
End Sub

Private Sub chkTransAll_Click()
cmdNew_Click
End Sub

Private Sub chkTransUsed_Click()
Dim lstIndex As Integer
If chkTransUsed.Value = 1 Then
    For lstIndex = 0 To lstTrans.ListCount - 1
        RunSql "Select * from tblTransactions where trans_no = " & lstTrans.List(lstIndex)
        Select Case Rs.Fields!Transaction
            Case "Borrow"
                s = "Item Barrowing"
                Table = "tblBorrow"
            Case "Reserve"
                s = "Item Reservation"
                Table = "tblReserve"
            Case "Return"
                s = "Returned Items"
                Table = "tblReturn"
            Case "Cancel"
                s = "Canceled Reservations"
                Table = "tblCancel"
        End Select
        
        If Table = "tblReturn" Or Table = "tblCancel" Then
            lstTrans.Selected(lstIndex) = True
        Else
            RunSql "SELECT tbl.* " & _
                    "FROM " & Table & " as tbl INNER JOIN tblTransactions as trans ON tbl.trans_no = trans.trans_no " & _
                    "WHERE trans.trans_no = " & lstTrans.List(lstIndex)
            With Rs
                If .EOF = True Then
                    lstTrans.Selected(lstIndex) = True
                End If
            End With
        End If
    Next lstIndex
Else
    For lstIndex = 0 To lstTrans.ListCount - 1
        lstTrans.Selected(lstIndex) = False
    Next lstIndex
End If
End Sub

Private Sub cmdBRemove_Click()
If NoRcrd(lvwBView) = True Then Exit Sub
lvwBView.ListItems.Remove lvwBView.SelectedItem.Index
End Sub

Private Sub cmdBAdd_Click()
If lblId.Caption = "---" Then
    MsgBox "Please search for a client profile to continue transaction", vbExclamation
    txtSrchStr.SetFocus
    Exit Sub
End If
If NoRcrd(lvwBItems, "No available items on the list. Please search for an items.") = True Then Exit Sub
If TxtEmp(txtBTime) = True Then Exit Sub

n = ValBox("Input quantity to Borrow.", imgIcon, , lvwBItems.SelectedItem.SubItems(2), "Borrow")
If n = 0 Or n > Val(lvwBItems.SelectedItem.SubItems(2)) Then
    MsgBox "Cannot accept 0 or greater than the total available quantity of the item.", vbExclamation
    Exit Sub
End If

With lvwBView
    IntSame = 0
    For i = 1 To .ListItems.Count
        If .ListItems(i).Text = lvwBItems.SelectedItem And .ListItems(i).SubItems(4) = lvwBItems.SelectedItem.SubItems(3) Then
            IntSame = i
        End If
    Next i
    If IntSame = 0 Then
        Set x = .ListItems.Add(, , lvwBItems.SelectedItem)
    Else
        Set x = .ListItems(IntSame)
    End If
    x.SubItems(1) = n
    x.SubItems(2) = txtBTime.Text
    x.SubItems(3) = cboBGap.Text
    x.SubItems(4) = lvwBItems.SelectedItem.SubItems(3)
End With
ClearBItems
End Sub

Private Sub cmdCancel_Click()
If lblId.Caption = "---" Then
    MsgBox "Please search for a profile to begin transaction.", vbExclamation
    txtSrchStr.SetFocus
    Exit Sub
End If
If NoRcrd(lvwReserved, "This client does not have any reserved items.") = True Then Exit Sub
With lvwCanceled
    IntSame = 0
    For i = 1 To .ListItems.Count
        If .ListItems(i).Text = lvwReserved.SelectedItem Then
            IntSame = i
        End If
    Next i
    If IntSame = 0 Then
        Set x = .ListItems.Add(, , lvwReserved.SelectedItem)
    Else
        Set x = .ListItems(IntSame)
    End If
    x.SubItems(1) = lvwReserved.SelectedItem.SubItems(1)
    x.SubItems(2) = lvwReserved.SelectedItem.SubItems(3)
    x.SubItems(3) = lvwReserved.SelectedItem.SubItems(4)
    x.SubItems(4) = txtCRemarks.Text
End With
End Sub

Private Sub cmdCCancel_Click()
If NoRcrd(lvwCanceled) = True Then Exit Sub
lvwCanceled.ListItems.Remove lvwCanceled.SelectedItem.Index
End Sub

Private Sub cmdCReturn_Click()
If NoRcrd(lvwReturned) = True Then Exit Sub
lvwReturned.ListItems.Remove lvwReturned.SelectedItem.Index
End Sub

Private Sub cmdClear_Click()
lblId.Caption = "---"
lblName.Caption = "---"
lblTrans.Caption = "---"
cmdNew_Click
txtSrchStr.SetFocus
imgProfile.Picture = LoadPicture()
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelSelected_Click()
Dim DelCount As Long
If lstTrans.ListCount = 0 Then
    MsgBox "No available transactions on the list.", vbExclamation
    Exit Sub
End If
DelCount = 0
If MsgBox("You cannot view thier history description if you delete this transactions. Click OK to continue.", vbExclamation + vbOKCancel) = vbCancel Then Exit Sub
For n = 0 To lstTrans.ListCount - 1
    If lstTrans.Selected(n) = True Then
        If RunSql("Delete * from tblTransactions where trans_no = " & lstTrans.List(n)) = True Then
            MsgBox "The system failed to delete Transaction # " & lstTrans.List(n) _
                    & ". Other records maybe using this transaction.", vbCritical
        Else
            DelCount = DelCount + 1
        End If
    End If
Next n
MsgBox DelCount & " transaction(s) were deleted.", vbInformation
cmdNew_Click
frmTransView.tabMenu_Click
End Sub

Private Sub cmdNew_Click()
If lblId.Caption = "---" Then Exit Sub
Select Case tabTrans.SelectedItem.Index
    Case 1
        ClearBItems
        lvwBView.ListItems.Clear
    Case 2
        ClearRItems
        lvwRView.ListItems.Clear
    Case 3
        ClearBorrowed
        lvwReturned.ListItems.Clear
    Case 4
        ClearReserved
        lvwCanceled.ListItems.Clear
    Case 5
        ViewTransactions chkTransAll.Value, lstTrans, Val(lblId.Caption)
        lblTransName.Caption = "---"
        lblTransDate.Caption = "---"
        lblTransCategory.Caption = "---"
        lvwTransView.ListItems.Clear
        chkTransUsed.Value = 0
        lstTrans_Click
End Select
txtHistory.Text = ViewHistory(Val(lblId.Caption), tabTrans.SelectedItem.Key)
End Sub

Private Sub cmdRAdd_Click()
If lblId.Caption = "---" Then
    MsgBox "Please search for a client profile to continue transaction", vbExclamation
    txtSrchStr.SetFocus
    Exit Sub
End If
If NoRcrd(lvwRItems, "No available items on the list. Please search for an items.") = True Then Exit Sub
For i = 0 To 1
    If chkTest(i).Value <> 1 Then
        MsgBox "Please confirm your date settings. Check to confirm.", vbExclamation
        Exit Sub
    End If
Next i

n = DateDiff("d", dtpRDate.Value, Now)
If n > 0 Then
    MsgBox "Invalid Date: reserve date must be beyond the current date.", vbExclamation
    chkTest(0).Value = 0
    Exit Sub
ElseIf n = 0 And DateDiff("h", dtpRTime.Value, Now) > 0 Then
    MsgBox "Invalid Time: The inputed date is the current date, but your time value is later than the current time.", vbExclamation
    chkTest(0).Value = 0
    Exit Sub
End If

n = ValBox("Input quantity to Reserve.", imgIcon, , lvwRItems.SelectedItem.SubItems(2), "Borrow")
If n = 0 Or n > Val(lvwRItems.SelectedItem.SubItems(2)) Then
    MsgBox "Cannot accept 0 or greater than the total available quantity of the item.", vbExclamation
    Exit Sub
End If

With lvwRView
    IntSame = 0
    For i = 1 To .ListItems.Count
        If .ListItems(i).Text = lvwRItems.SelectedItem And .ListItems(i).SubItems(3) = lvwRItems.SelectedItem.SubItems(3) Then
            IntSame = i
        End If
    Next i
    If IntSame = 0 Then
        Set x = .ListItems.Add(, , lvwRItems.SelectedItem)
    Else
        Set x = .ListItems(IntSame)
    End If
    x.SubItems(1) = n
    x.SubItems(2) = Format(dtpRDate.Value, "mm/dd/yyyy") & " " & Format(dtpRTime.Value, "hh:nn ampm")
    x.SubItems(3) = lvwRItems.SelectedItem.SubItems(3)
End With
ClearRItems
End Sub

Private Sub cmdReturn_Click()
If lblId.Caption = "---" Then
    MsgBox "Please search for a profile to begin transaction.", vbExclamation
    txtSrchStr.SetFocus
    Exit Sub
End If
If NoRcrd(lvwBorrowed, "This client does not have any Borrowed items.") = True Then Exit Sub
With lvwReturned
    IntSame = 0
    For i = 1 To .ListItems.Count
        If .ListItems(i).Text = lvwBorrowed.SelectedItem Then
            IntSame = i
        End If
    Next i
    If IntSame = 0 Then
        Set x = .ListItems.Add(, , lvwBorrowed.SelectedItem)
    Else
        Set x = .ListItems(IntSame)
    End If
    x.SubItems(1) = lvwBorrowed.SelectedItem.SubItems(1)
    x.SubItems(2) = lvwBorrowed.SelectedItem.SubItems(3)
    x.SubItems(3) = lvwBorrowed.SelectedItem.SubItems(4)
End With
End Sub

Private Sub cmdRRemove_Click()
If NoRcrd(lvwRView) = True Then Exit Sub
lvwRView.ListItems.Remove lvwRView.SelectedItem.Index
End Sub

Private Sub cmdSave_Click()
Select Case tabTrans.SelectedItem.Index
    Case 1
        If NoRcrd(lvwBView, "No transaction to save.") = True Then Exit Sub
        If MsgBox("Transaction " & lblTrans.Caption & " will be saved. Click OK to continue.", vbExclamation + vbOKCancel) = vbOK Then
            MSG = SaveBTrans
            ClearBItems
        End If
    Case 2
        If NoRcrd(lvwRView, "No transaction to save.") = True Then Exit Sub
        If MsgBox("Transaction " & lblTrans.Caption & " will be saved. Click OK to continue.", vbExclamation + vbOKCancel) = vbOK Then
            MSG = SaveRTrans
            ClearRItems
        End If
    Case 3
        If NoRcrd(lvwReturned, "No returned items yet.") = True Then Exit Sub
        If chkSvReturn.Value = 1 Then
            If MsgBox("Transaction " & lblTrans.Caption & " will be saved. Click OK to continue.", vbExclamation + vbOKCancel) = vbOK Then
                MSG = SaveReturn
                ClearBorrowed
            End If
        Else
            For i = 1 To lvwReturned.ListItems.Count
                UpdateStat lvwReturned, "Returned", lvwReturned.ListItems(i).SubItems(3), i
                MSG = "Transaction has been successfully processed."
                ClearBorrowed
            Next i
        End If
    Case 4
        If NoRcrd(lvwCanceled, "No returned items yet.") = True Then Exit Sub
        If chkSvCancel.Value = 1 Then
            If MsgBox("Transaction " & lblTrans.Caption & " will be saved. Click OK to continue.", vbExclamation + vbOKCancel) = vbOK Then
                MSG = SaveCancel
                ClearReserved
            End If
        Else
            For i = 1 To lvwCanceled.ListItems.Count
                UpdateStat lvwCanceled, "Canceled", lvwCanceled.ListItems(i).SubItems(3), i
                ClearReserved
                MSG = "Transaction has been successfully processed."
            Next i
        End If
    Case Else
        MSG = "No need to save for this section."
End Select
If MSG <> Empty Then MsgBox MSG, vbInformation
cmdNew_Click
frmTransView.tabMenu_Click
End Sub

Private Sub cmdSearch_Click()
For i = 1 To Len(txtSrchStr.Text)
    If Right(Left(txtSrchStr.Text, i), 1) = "'" Then
        txtSrchStr.Text = Empty
        MsgBox "The input contains an invalid character.", vbCritical
        txtSrchStr.SetFocus
        Exit Sub
    End If
Next i

If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ExecSrch cboFilter.Text, txtSrchStr.Text
    End If
Else
    cmdClear_Click
End If
End Sub

Private Sub dtpRDate_Change()
chkTest(0).Value = 0
End Sub

Private Sub dtpRTime_Change()
chkTest(1).Value = 0
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Public Sub ExecSrch(FldStr As String, RcrdStr As String)
RunSql "Select * from tblClientProfile where " & FldStr & " LIKE '" & RcrdStr & "%'"
If Rs.RecordCount = 0 Then
    cmdClear_Click
    MsgBox "No record found on search.", vbExclamation
    txtSrchStr.SetFocus
    SelAll txtSrchStr
    Exit Sub
End If
s = StrBox(Rs.RecordCount & " record(s) found on search.", imgIcon, , , "select client", 3, "tblClientProfile", FldStr, RcrdStr)
If s = Empty Then Exit Sub
SubSql "Select * from tblClientProfile where " & FldStr & " LIKE '" & s & "'"
With SubRs
    If .EOF = False Then
        lblId.Caption = .Fields!client_no
        lblName.Caption = StrConv(.Fields!fname & " " & Left(.Fields!mname, 1) & ". " & .Fields!lname, vbProperCase)
        If .Fields!image_name <> Empty Then
            imgProfile.Picture = LoadPicture(App.Path & "\Images\Accounts\" & .Fields!image_name)
        Else
            imgProfile.Picture = LoadPicture()
        End If
        cmdNew_Click
    End If
End With
End Sub

Private Sub Form_Load()
SetLv lvwBView, True, True
SetLv lvwBItems, True, True
SetLv lvwRView, True, True
SetLv lvwRItems, True, True
SetLv lvwBorrowed, True, True
SetLv lvwReturned, True, True
SetLv lvwReserved, True, True
SetLv lvwCanceled, True, True
SetLv lvwTransView, True, True
DtpValue dtpRTime
DtpValue dtpRDate
LoadCboFld "tblClientProfile", "*", cboFilter
For i = 2 To freTrans.UBound
    freTrans(1).Height = tabTrans.Height - 420
    freTrans(1).Width = tabTrans.Width - 110
    freTrans(1).Top = tabTrans.Top + 340
    freTrans(1).Left = tabTrans.Left + 30
    freTrans(i).Move _
        freTrans(1).Left, _
        freTrans(1).Top, _
        freTrans(1).Width, _
        freTrans(1).Height
    freTrans(i).Visible = False
Next i
tabTrans_Click
End Sub

Private Sub lstTrans_Click()
If lstTrans.Text = Empty Then Exit Sub
RunSql "SELECT trans.*, client.* FROM tblTransactions as trans INNER JOIN tblClientProfile as [client] ON trans.client_no = client.client_no " & _
        "WHERE trans.trans_no = " & lstTrans.Text
With Rs
    Select Case .Fields!Transaction
        Case "Borrow"
            s = "Item Barrowing"
            Table = "tblBorrow"
        Case "Reserve"
            s = "Item Reservation"
            Table = "tblReserve"
        Case "Return"
            s = "Returned Items"
            Table = "tblReturn"
        Case "Cancel"
            s = "Canceled Reservations"
            Table = "tblCancel"
    End Select
    lblTransName.Caption = StrConv(.Fields!fname & " " & Left(.Fields!mname, 1) & ". " & .Fields!lname, vbProperCase)
    lblTransDate.Caption = Format(.Fields!trans_date, "mm/dd/yyyy hh:nn ampm")
    lblTransCategory.Caption = s
    ViewTrans Table, lstTrans.Text, lblId.Caption, lvwTransView
End With
End Sub

Private Sub lvwBorrowed_Click()
If NoRcrd(lvwBorrowed) = True Then Exit Sub
RunSql "SELECT bar.*, trans.*, client.* " & _
        "FROM (tblBorrow as bar INNER JOIN tblTransactions as trans ON bar.trans_no = trans.trans_no) " & _
        "INNER JOIN tblClientProfile as [client] ON client.client_no = trans.client_no " & _
        "WHERE bar.record_no = " & lvwBorrowed.SelectedItem & " and client.client_no = " & lblId.Caption
With Rs
    lblBDate.Caption = Format(.Fields("trans_date"), "mm/dd/yyyy hh:nn ampm")
    lblRDate.Caption = Scheduler(.Fields("trans_date"), .Fields("gap_val"), .Fields("interval"))
    lblBTrans.Caption = .Fields("trans.trans_no")
    lblBrange.Caption = .Fields!gap_val & " " & .Fields!Interval & "(s)"
    lblQty.Caption = .Fields("qty")
End With
End Sub

Private Sub lvwBorrowed_DblClick()
cmdReturn_Click
End Sub

Private Sub lvwBItems_Click()
If NoRcrd(lvwBItems) = True Then Exit Sub
RunSql "SELECT list.location, reg.qty " & _
        "FROM tblItemList as list INNER JOIN tblRegistered as reg ON list.item_id = reg.item_id " & _
        "WHERE list.item_id = '" & lvwBItems.SelectedItem & "'"
With Rs
    lblBLocation.Caption = .Fields!location
    lblBCount.Caption = .Fields!qty
    lblBStatus.Caption = lvwBItems.SelectedItem.SubItems(3)
    lblBQty.Caption = lvwBItems.SelectedItem.SubItems(2)
End With
End Sub

Private Sub lvwBItems_DblClick()
cmdBAdd_Click
End Sub

Private Sub lvwBItems_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBAdd_Click
End If
End Sub

Private Sub lvwBView_DblClick()
If NoRcrd(lvwBView) = True Then Exit Sub
With lvwBView
    ViewItemList "tblBorrow", "item_id", .SelectedItem, lvwBItems, lblId.Caption
    txtBTime.Text = .SelectedItem.SubItems(2)
    cboBGap.Text = .SelectedItem.SubItems(3)
End With
End Sub

Private Sub lvwReserved_Click()
If NoRcrd(lvwReserved) = True Then Exit Sub
RunSql "SELECT res.*, trans.*, client.* " & _
        "FROM (tblReserve as res INNER JOIN tblTransactions as trans ON res.trans_no = trans.trans_no) " & _
        "INNER JOIN tblClientProfile as [client] ON client.client_no = trans.client_no " & _
        "WHERE res.record_no = " & lvwReserved.SelectedItem & " and client.client_no = " & lblId.Caption
With Rs
    lblTDate.Caption = Format(.Fields!trans_date, "mm/dd/yyyy hh:nn ampm")
    lblReserveDate.Caption = Format(.Fields!reserve_date, "mm/dd/yyyy hh:nn ampm")
    lblCQty.Caption = .Fields!qty
    lblCTrans.Caption = .Fields("trans.trans_no")
End With
End Sub

Private Sub lvwReserved_DblClick()
cmdCancel_Click
End Sub

Private Sub lvwRItems_Click()
If NoRcrd(lvwRItems) = True Then Exit Sub
RunSql "SELECT list.location, reg.qty " & _
        "FROM tblItemList as list INNER JOIN tblRegistered as reg ON list.item_id = reg.item_id " & _
        "WHERE list.item_id = '" & lvwRItems.SelectedItem & "'"
With Rs
    lblRLocation.Caption = .Fields!location
    lblRCount.Caption = .Fields!qty
    lblRStatus.Caption = lvwRItems.SelectedItem.SubItems(3)
    lblRQty.Caption = lvwRItems.SelectedItem.SubItems(2)
End With
End Sub

Private Sub lvwRItems_DblClick()
cmdRAdd_Click
End Sub

Private Sub lvwRView_DblClick()
If NoRcrd(lvwRView) = True Then Exit Sub
With lvwRView
    ViewItemList "tblReserve", "item_id", .SelectedItem, lvwRItems, lblId.Caption
    dtpRDate.Value = Format(.SelectedItem.SubItems(2), "mm/dd/yyyy")
    dtpRTime.Value = Format(.SelectedItem.SubItems(2), "hh:nn ampm")
End With
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub tabTrans_Click()
For i = 1 To tabTrans.Tabs.Count
    If freTrans(i).Index = tabTrans.SelectedItem.Index Then
        freTrans(i).Visible = True
    Else
        freTrans(i).Visible = False
    End If
Next i
cmdNew_Click
End Sub

Private Sub txtBarSrch_Change()
If Right(txtBarSrch.Text, 1) = "'" Then
    txtBarSrch.Text = Empty
End If
If Trim(txtBarSrch.Text) <> Empty Then
    If txtBarSrch.Text <> "Search" Then
        ViewBorrowed "description", txtBarSrch.Text, lvwBorrowed, lblId.Caption
    End If
Else
    ViewBorrowed "description", "%", lvwBorrowed, lblId.Caption
End If
End Sub

Private Sub txtBTime_Change()
If Trim(txtBTime.Text) = Empty Or IsNumeric(txtBTime) = False Then
    lblBReturn.Caption = "---"
    Exit Sub
End If
lblBReturn.Caption = Scheduler(Now, Val(Trim(txtBTime.Text)), cboBGap.Text)
End Sub

Private Sub txtBItemSrch_Change()
If Right(txtBItemSrch.Text, 1) = "'" Then
    txtBItemSrch.Text = Empty
End If
If Trim(txtBItemSrch.Text) <> Empty Then
    If txtBItemSrch.Text <> "Search" Then
        ViewItemList "tblBorrow", "description", txtBItemSrch.Text, lvwBItems, lblId.Caption
    End If
Else
    ViewItemList "tblBorrow", "description", "%", lvwBItems, lblId.Caption
End If
End Sub

Private Sub txtBItemSrch_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
    lvwBItems.SetFocus
End If
End Sub

Private Sub txtBItemSrch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBAdd_Click
End If
End Sub

Private Sub txtRDate_GotFocus()
SelAll txtRDate
End Sub

Private Sub txtRTime_GotFocus()
SelAll txtRTime
End Sub

Private Sub txtResSrch_Change()
If Right(txtResSrch.Text, 1) = "'" Then
    txtResSrch.Text = Empty
End If
If Trim(txtResSrch.Text) <> Empty Then
    If txtResSrch.Text <> "Search" Then
        ViewReserved "description", txtResSrch.Text, lvwReserved, lblId.Caption
    End If
Else
    ViewReserved "description", "%", lvwReserved, lblId.Caption
End If
End Sub

Private Sub txtRItemSrch_Change()
If Right(txtRItemSrch.Text, 1) = "'" Then
    txtRItemSrch.Text = Empty
End If
If Trim(txtRItemSrch.Text) <> Empty Then
    If txtRItemSrch.Text <> "Search" Then
        ViewItemList "tblReserve", "description", txtRItemSrch.Text, lvwRItems, lblId.Caption
    End If
Else
    ViewItemList "tblReserve", "description", "%", lvwRItems, lblId.Caption
End If
End Sub

Private Sub txtRItemSrch_GotFocus()
If txtRItemSrch = "Search" Then
    txtRItemSrch.Text = Empty
    txtRItemSrch.ForeColor = &H80000008
End If
End Sub

Private Sub txtRItemSrch_LostFocus()
If Trim(txtRItemSrch) = Empty Then
    txtRItemSrch.Text = "Search"
    txtRItemSrch.ForeColor = &H80000011
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
    txtSrchStr.ForeColor = &H80000011
End If
End Sub

Private Sub txtBItemSrch_GotFocus()
If txtBItemSrch = "Search" Then
    txtBItemSrch.Text = Empty
    txtBItemSrch.ForeColor = &H80000008
End If
End Sub

Private Sub txtBItemSrch_LostFocus()
If Trim(txtBItemSrch) = Empty Then
    txtBItemSrch.Text = "Search"
    txtBItemSrch.ForeColor = &H80000011
End If
End Sub

Private Sub txtBarSrch_GotFocus()
If txtBarSrch = "Search" Then
    txtBarSrch.Text = Empty
    txtBarSrch.ForeColor = &H80000008
End If
End Sub

Private Sub txtBarSRch_LostFocus()
If Trim(txtBarSrch) = Empty Then
    txtBarSrch.Text = "Search"
    txtBarSrch.ForeColor = &H80000011
End If
End Sub

Private Sub txtresSrch_GotFocus()
If txtResSrch = "Search" Then
    txtResSrch.Text = Empty
    txtResSrch.ForeColor = &H80000008
End If
End Sub

Private Sub txtresSRch_LostFocus()
If Trim(txtResSrch) = Empty Then
    txtResSrch.Text = "Search"
    txtResSrch.ForeColor = &H80000011
End If
End Sub
