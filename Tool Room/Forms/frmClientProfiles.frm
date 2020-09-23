VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmClientProfiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profiles"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   Icon            =   "frmClientProfiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   35
      Top             =   960
      Width           =   4095
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
         Left            =   1920
         TabIndex        =   37
         Text            =   "Search"
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox cboFilter 
         Height          =   315
         ItemData        =   "frmClientProfiles.frx":038A
         Left            =   240
         List            =   "frmClientProfiles.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   240
         Width           =   1335
      End
      Begin CtrlLine.ctrlLiner ctrlLiner2 
         Height          =   30
         Left            =   1680
         TabIndex        =   38
         Top             =   360
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   3600
         Picture         =   "frmClientProfiles.frx":038E
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Personal Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   25
      Top             =   2520
      Width           =   7095
      Begin VB.TextBox txtIdNum 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   9
         Top             =   2400
         Width           =   1695
      End
      Begin VB.CheckBox chkCourse 
         Caption         =   "Course"
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
         Left            =   1680
         TabIndex        =   45
         Top             =   3240
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "Department"
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
         Left            =   1680
         TabIndex        =   44
         Top             =   2880
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   645
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2880
         Width           =   1695
      End
      Begin VB.ComboBox cboDept 
         Height          =   315
         ItemData        =   "frmClientProfiles.frx":0C58
         Left            =   5160
         List            =   "frmClientProfiles.frx":0C65
         TabIndex        =   5
         Text            =   "cboDept"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtStat 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   8
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cboCourse 
         Height          =   315
         ItemData        =   "frmClientProfiles.frx":0C7F
         Left            =   5160
         List            =   "frmClientProfiles.frx":0C8C
         TabIndex        =   7
         Text            =   "cboCourse"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtMname 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtFname 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtLname 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   2
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtAge 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1920
         Width           =   375
      End
      Begin VB.ComboBox cboGender 
         Height          =   315
         ItemData        =   "frmClientProfiles.frx":0CA6
         Left            =   1680
         List            =   "frmClientProfiles.frx":0CB3
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtContact 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   """Php""#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "ID Number:"
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
         TabIndex        =   47
         Top             =   2400
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Auto Save:"
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
         TabIndex        =   46
         Top             =   2880
         Width           =   915
      End
      Begin VB.Label Label9 
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
         Left            =   3840
         TabIndex        =   43
         Top             =   2880
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "* Department:"
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
         Left            =   3720
         TabIndex        =   42
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label7 
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
         Left            =   3840
         TabIndex        =   41
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Course:"
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
         TabIndex        =   40
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Contact #:"
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
         TabIndex        =   31
         Top             =   960
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "* First Name:"
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
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Middle Name:"
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
         TabIndex        =   29
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "* Surname:"
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
         TabIndex        =   28
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "* Age:"
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
         Top             =   1920
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "* Gender:"
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
         Width           =   810
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Client #"
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
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   2535
      Begin VB.Label lblId 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0000"
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
         TabIndex        =   24
         Top             =   240
         Width           =   420
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "User Count"
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
      Left            =   2760
      TabIndex        =   20
      Top             =   1800
      Width           =   1455
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0000"
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
         Left            =   1620
         TabIndex        =   22
         Top             =   280
         Width           =   420
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   210
         Left            =   600
         TabIndex        =   21
         Top             =   240
         Width           =   120
      End
   End
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
   End
   Begin MSComDlg.CommonDialog dlgPic 
      Left            =   4920
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ListView lvwProfiles 
      Height          =   1575
      Left            =   120
      TabIndex        =   34
      Top             =   6960
      Width           =   7095
      _ExtentX        =   12515
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Client ID"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ID Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "First Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Middle Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Last Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Age"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Gender"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Contact Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Course"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   10
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Department"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   11
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Remarks"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.Frame Frame5 
      Caption         =   "Profile Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4320
      TabIndex        =   32
      Top             =   960
      Width           =   2895
      Begin ctrlButton.ThemedButton cmdRemove 
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Caption         =   "&Remove"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmClientProfiles.frx":0CCD
         Picture         =   "frmClientProfiles.frx":0EA7
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdBrowse 
         Height          =   375
         Left            =   120
         TabIndex        =   49
         Top             =   840
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
         MouseIcon       =   "frmClientProfiles.frx":11FB
         Picture         =   "frmClientProfiles.frx":13D5
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin VB.Image imgProfile 
         Height          =   1095
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000011&
         FillColor       =   &H80000004&
         Height          =   1095
         Left            =   1440
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label16 
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
         Left            =   1680
         TabIndex        =   33
         Top             =   720
         Width           =   825
      End
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   6480
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
      MouseIcon       =   "frmClientProfiles.frx":1729
      Picture         =   "frmClientProfiles.frx":1903
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdNew 
      Height          =   375
      Left            =   3000
      TabIndex        =   13
      Top             =   6480
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
      MouseIcon       =   "frmClientProfiles.frx":1C57
      Picture         =   "frmClientProfiles.frx":1E31
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdDelete 
      Height          =   375
      Left            =   5880
      TabIndex        =   17
      Top             =   8640
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
      MouseIcon       =   "frmClientProfiles.frx":2185
      Picture         =   "frmClientProfiles.frx":235F
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdClose 
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   6480
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
      MouseIcon       =   "frmClientProfiles.frx":26B3
      Picture         =   "frmClientProfiles.frx":288D
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdLoad 
      Height          =   375
      Left            =   4440
      TabIndex        =   16
      Top             =   8640
      Width           =   1335
      _ExtentX        =   2355
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
      MouseIcon       =   "frmClientProfiles.frx":2BE1
      Picture         =   "frmClientProfiles.frx":2DBB
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdOptions 
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   6480
      Width           =   1335
      _ExtentX        =   2355
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
      MouseIcon       =   "frmClientProfiles.frx":310F
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select a User to manage"
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
      TabIndex        =   39
      Top             =   8760
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmClientProfiles.frx":32E9
      Top             =   8640
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmClientProfiles.frx":3BB3
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLIENTS"
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
      TabIndex        =   19
      Top             =   120
      Width           =   750
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Manage client profiles"
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
      TabIndex        =   18
      Top             =   480
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000000&
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "frmClientProfiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()
On Error GoTo InvldPic
dlgPic.DialogTitle = "Load Item Image"
dlgPic.InitDir = "My Documents"
dlgPic.Filter = "Jepeg Image (*.jpg;*.jpeg)|*.jpg;*.jpeg|Bitmap Image (*.bmp)|*.bmp|All Files (*.*)|*.*"
dlgPic.ShowOpen
If dlgPic.FileName = "" Then
    Exit Sub
End If
imgProfile.Picture = LoadPicture(dlgPic.FileName)
ImgName = dlgPic.FileTitle
ImgSrc = dlgPic.FileName
Exit Sub
InvldPic:
    MsgBox "It is not a valid picture", vbExclamation
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If MsgBox("Your about to delete the item, " & lvwProfiles.SelectedItem & ". Deleting this item may affect system reports." & vbNewLine & vbNewLine & _
    "Are you sure?", vbExclamation + vbYesNo) = vbYes Then
    If RunSql("Delete * from tblClientProfile where client_no = " & lvwProfiles.SelectedItem) = True Then
        MsgBox "Cannot delete this client because it is used by other transactions.", vbCritical
        Exit Sub
    End If
    ClrFlds
    MsgBox "Successfully deleted a record.", vbInformation
    viewProfiles "client_no", "%"
End If
End Sub

Private Sub cmdLoad_Click()
If NoRcrd(lvwProfiles) = True Then Exit Sub
ExecSrch "client_no", lvwProfiles.SelectedItem
End Sub

Private Sub cmdNew_Click()
ClrFlds
End Sub

Private Sub cmdRemove_Click()
imgProfile.Picture = LoadPicture()
ImgName = Empty
End Sub

Private Sub cmdSave_Click()

If chkDept.Value = 1 Then
    RunSql "Select * from tblDepartments where description = '" & Trim(cboDept.Text) & "'"
    With Rs
        If .EOF = True Then
            .AddNew
            .Fields!Description = Trim(cboDept.Text)
            .Fields!remarks = ""
            .Update
        End If
    End With
End If

If chkCourse.Value = 1 Then
    RunSql "Select * from tblCourse where description = '" & Trim(cboCourse.Text) & "'"
    With Rs
        If .EOF = True Then
            .AddNew
            .Fields!Description = Trim(cboCourse.Text)
            .Fields!remarks = ""
            .Update
        End If
    End With
End If

RunSql "Select * from tblClientProfile where client_no = " & lblId.Caption
With Rs
    If cmdSave.Caption = "&Save" Then
        .AddNew
        MSG = "Added new client on database"
        If ImgName <> Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Accounts\" & ImgName
        End If
    Else
        If ImgName <> Empty And .Fields!image_name = Empty Then
            FileCopy ImgSrc, App.Path & "\Images\Accounts\" & ImgName
        End If
        If .Fields!image_name <> Empty And .Fields!image_name <> ImgName Then
            Kill App.Path & "\Images\Accounts\" & .Fields!image_name
            If ImgName <> Empty Then
                FileCopy ImgSrc, App.Path & "\Images\Accounts\" & ImgName
            End If
        End If
        MSG = "Client " & lblId.Caption & " has been updated"
    End If
    .Fields!client_no = Val(lblId.Caption)
    .Fields!id_number = txtIdNum.Text
    .Fields!fname = txtFname.Text
    .Fields!mname = txtMname.Text
    .Fields!lname = txtLname.Text
    .Fields!age = Val(txtAge.Text)
    .Fields!gender = cboGender.Text
    .Fields!contact_num = txtContact.Text
    .Fields!course = cboCourse.Text
    .Fields!Status = txtStat.Text
    .Fields!department = cboDept.Text
    .Fields!remarks = txtRemarks.Text
    .Fields!image_name = ImgName
    .Fields!date_reg = Format(Date, "mm/dd/yyyy")
    .Update
End With
MsgBox MSG, vbInformation
ClrFlds
viewProfiles "client_no", "%"
End Sub

Private Sub Form_Activate()
Screen.MousePointer = 0
End Sub

Private Sub SetUp()
LoadCbo "tblCourse", cboCourse, "description"
LoadCbo "tblDepartments", cboDept, "description"
RunSql "Select * from tblClientProfile"
lblCount.Caption = Rs.RecordCount

ImgName = Empty
ImgSrc = Empty
ClrFlds
End Sub

Private Sub Form_Load()
SetLv lvwProfiles, True, True
LoadCboFld "tblClientProfile", "*", cboFilter
SetUp
End Sub

Private Sub ClrFlds()
lblId.Caption = RcrdId("tblClientProfile", Format(Date, "yymm"), "client_no")
cboDept.Text = Empty
cboCourse.Text = Empty
txtFname.Text = Empty
txtIdNum.Text = Empty
txtMname.Text = Empty
txtLname.Text = Empty
txtAge.Text = Empty
cboGender.ListIndex = 0
txtStat.Text = Empty
txtRemarks.Text = Empty
cmdSave.Caption = "&Save"
viewProfiles "client_no", "%"
End Sub

Private Sub lvwProfiles_DblClick()
cmdLoad_Click
End Sub

Private Sub txtSrchStr_Change()
If Right(txtSrchStr.Text, 1) = "'" Then
    txtSrchStr.Text = Empty
End If
If Trim(txtSrchStr.Text) <> Empty Then
    If txtSrchStr.Text <> "Search" Then
        ExecSrch cboFilter.Text, txtSrchStr.Text
        viewProfiles cboFilter.Text, txtSrchStr.Text
    End If
Else
    ClrFlds
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

Public Sub ExecSrch(SrchFld As String, SrchStr As String)
RunSql "Select * from tblClientProfile where " & SrchFld & " LIKE '" & SrchStr & "%'"
With Rs
    If .EOF = False Then
        txtFname.Text = .Fields!fname
        txtMname.Text = .Fields!mname
        txtLname.Text = .Fields!lname
        txtAge.Text = .Fields!age
        cboGender.Text = .Fields!gender
        cboDept.Text = .Fields!department
        txtContact.Text = .Fields!contact_num
        cboCourse.Text = .Fields!course
        txtRemarks.Text = .Fields!remarks
        cboCourse.Text = .Fields!course
        txtIdNum.Text = .Fields!id_number
        txtStat.Text = .Fields!Status
        lblId.Caption = .Fields!client_no
        If .Fields!image_name <> Empty Then
            imgProfile.Picture = LoadPicture(App.Path & "\Images\Accounts\" & .Fields!image_name)
        Else
            imgProfile.Picture = LoadPicture()
        End If
        ImgName = .Fields!image_name
        cmdSave.Caption = "&Update"
    Else
        ClrFlds
    End If
End With
End Sub

Public Sub viewProfiles(RcrdFld As String, RcrdStr As String)
RunSql "Select * from tblClientProfile where " & RcrdFld & " LIKE '" & RcrdStr & "%' Order By fname ASC"
With Rs
    lvwProfiles.ListItems.Clear
    While Not .EOF = True
        Set x = lvwProfiles.ListItems.Add(, , .Fields(0))
        For i = 1 To (.Fields.Count - 3)
            x.SubItems(i) = .Fields(i)
        Next i
        .MoveNext
    Wend
End With
End Sub


