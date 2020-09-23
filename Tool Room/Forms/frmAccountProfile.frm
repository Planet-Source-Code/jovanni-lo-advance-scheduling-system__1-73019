VERSION 5.00
Object = "{31E6A7F3-C63A-434F-97FB-33491A4E7C95}#1.0#0"; "CtrlLine.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FFB3BC8A-E4B0-40B1-93E5-84F95251C328}#1.0#0"; "ctrlButton.ocx"
Begin VB.Form frmAccountProfile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounts"
   ClientHeight    =   7950
   ClientLeft      =   3930
   ClientTop       =   1965
   ClientWidth     =   7350
   Icon            =   "frmAccountProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   7950
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin CtrlLine.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   0
      TabIndex        =   51
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   53
   End
   Begin VB.Frame Frame7 
      Caption         =   "ID Count"
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
      Left            =   3120
      TabIndex        =   46
      Top             =   1800
      Width           =   1095
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
         Left            =   480
         TabIndex        =   48
         Top             =   240
         Width           =   120
      End
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
         TabIndex        =   47
         Top             =   280
         Width           =   420
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "PC NAME"
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
      TabIndex        =   42
      Top             =   960
      Width           =   1815
      Begin VB.Label lblPcId 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PC ID"
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
         Left            =   1050
         TabIndex        =   44
         Top             =   285
         Width           =   450
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0000"
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
         Left            =   3240
         TabIndex        =   43
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "User ID"
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
      Left            =   2040
      TabIndex        =   37
      Top             =   960
      Width           =   2175
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
         Left            =   1620
         TabIndex        =   38
         Top             =   280
         Width           =   420
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Employee Profile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Width           =   7095
      Begin VB.ComboBox cboEmpStatus 
         Height          =   315
         ItemData        =   "frmAccountProfile.frx":038A
         Left            =   1680
         List            =   "frmAccountProfile.frx":0397
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cboPosition 
         Height          =   315
         ItemData        =   "frmAccountProfile.frx":03B9
         Left            =   5040
         List            =   "frmAccountProfile.frx":03C0
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox cboRemarks 
         Height          =   315
         ItemData        =   "frmAccountProfile.frx":03CC
         Left            =   5040
         List            =   "frmAccountProfile.frx":03D6
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   960
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtpDHired 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   20
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47710211
         CurrentDate     =   40071
      End
      Begin MSMask.MaskEdBox txtDhired 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label lblEmpStatus 
         AutoSize        =   -1  'True
         Caption         =   "* Emp. Stat.:"
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
         TabIndex        =   36
         Top             =   960
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "* Position:"
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
         TabIndex        =   33
         Top             =   480
         Width           =   870
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "* Date Hired:"
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
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "* Remarks:"
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
         TabIndex        =   31
         Top             =   960
         Width           =   960
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
      Height          =   3015
      Left            =   120
      TabIndex        =   22
      Top             =   2520
      Width           =   7095
      Begin MSMask.MaskEdBox txtBdate 
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   1920
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   " "
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
         Left            =   5040
         TabIndex        =   8
         Top             =   2400
         Width           =   1575
      End
      Begin VB.ComboBox cboStatus 
         Height          =   315
         ItemData        =   "frmAccountProfile.frx":03EE
         Left            =   5040
         List            =   "frmAccountProfile.frx":03F8
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   1575
      End
      Begin VB.ComboBox cboGender 
         Height          =   315
         ItemData        =   "frmAccountProfile.frx":040D
         Left            =   5040
         List            =   "frmAccountProfile.frx":041A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
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
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   4
         Top             =   2400
         Width           =   375
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
      Begin VB.TextBox txtAddress 
         Height          =   735
         Left            =   5040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker dtpBdate 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "dd-mmm-yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   19
         Top             =   1920
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   556
         _Version        =   393216
         Format          =   47710211
         CurrentDate     =   40071
      End
      Begin VB.Label Label11 
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
         Left            =   3720
         TabIndex        =   35
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "* Birth Date:"
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
         Top             =   1920
         Width           =   1050
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
         Left            =   3600
         TabIndex        =   29
         Top             =   1440
         Width           =   810
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
         TabIndex        =   28
         Top             =   2400
         Width           =   525
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
         TabIndex        =   27
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "* Middle Name:"
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
         Top             =   960
         Width           =   1275
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
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "* Address:"
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
         Left            =   3600
         TabIndex        =   24
         Top             =   480
         Width           =   885
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
         Left            =   3720
         TabIndex        =   23
         Top             =   2400
         Width           =   885
      End
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
      Height          =   615
      Left            =   120
      TabIndex        =   39
      Top             =   1800
      Width           =   2895
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
         ForeColor       =   &H8000000B&
         Height          =   300
         Left            =   720
         TabIndex        =   16
         Text            =   "Search"
         Top             =   200
         Width           =   1695
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "ID:"
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
         TabIndex        =   40
         Top             =   240
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2400
         Picture         =   "frmAccountProfile.frx":0434
         Top             =   100
         Width           =   480
      End
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
      TabIndex        =   41
      Top             =   960
      Width           =   2895
      Begin ctrlButton.ThemedButton cmdRemove 
         Height          =   375
         Left            =   120
         TabIndex        =   17
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
         MouseIcon       =   "frmAccountProfile.frx":0CFE
         Picture         =   "frmAccountProfile.frx":0ED8
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin ctrlButton.ThemedButton cmdBrowse 
         Height          =   375
         Left            =   120
         TabIndex        =   18
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
         MouseIcon       =   "frmAccountProfile.frx":122C
         Picture         =   "frmAccountProfile.frx":1406
         PictureAlign    =   1
         PictureSize     =   0
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000011&
         FillColor       =   &H80000004&
         Height          =   1095
         Left            =   1440
         Top             =   240
         Width           =   1335
      End
      Begin VB.Image imgProfile 
         Height          =   1095
         Left            =   1440
         Stretch         =   -1  'True
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
         TabIndex        =   45
         Top             =   720
         Width           =   825
      End
   End
   Begin MSComDlg.CommonDialog dlgPic 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ctrlButton.ThemedButton cmdDelete 
      Height          =   375
      Left            =   4440
      TabIndex        =   14
      Top             =   7440
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
      MouseIcon       =   "frmAccountProfile.frx":175A
      Picture         =   "frmAccountProfile.frx":1934
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdSecurity 
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "&Security"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmAccountProfile.frx":1C88
      Picture         =   "frmAccountProfile.frx":1E62
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin ctrlButton.ThemedButton cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   13
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
      MouseIcon       =   "frmAccountProfile.frx":21B6
      Picture         =   "frmAccountProfile.frx":2390
      PictureAlign    =   1
      PictureSize     =   0
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Add, update and delete user accounts"
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
      TabIndex        =   50
      Top             =   480
      Width           =   2325
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROFILE"
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
      TabIndex        =   49
      Top             =   120
      Width           =   765
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   120
      Picture         =   "frmAccountProfile.frx":26E4
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Indecates a required field"
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
      TabIndex        =   34
      Top             =   7560
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   0
      Picture         =   "frmAccountProfile.frx":3328
      Top             =   7425
      Width           =   480
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000016&
      Height          =   855
      Left            =   -120
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "frmAccountProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'declare essencial variables
Option Explicit
Private Sub cmdBrowse_Click()
On Error GoTo InvldPic
dlgPic.DialogTitle = "Load Profile Image"
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

Private Sub cmdDelete_Click()
'locate a user id and delete if found
RunSql "Select * from tblAccountProfile where id = '" & lblId.Caption & "'"
With Rs
    'if no record...
    If .EOF = True Then
        MsgBox "No specified account to delete. Please search for an account", vbExclamation
        ClrFlds
        txtSrchStr.SetFocus
    Else
    If MsgBox("Are you sure you want to delete this user account? Click Yes to continue.", vbExclamation + vbYesNo) = vbNo Then
        Exit Sub
    End If
    'else if a record found
        'test if the image is not empty then delete from accounts images
        If imgProfile.Picture <> LoadPicture() Then
            Me.MousePointer = 11
            Kill App.Path & "\Images\accounts\" & .Fields!image_name
        End If
        'delete the account
        .Delete
        'clear all
        ClrFlds
        MsgBox "Account id " & lblId.Caption & " has been deleted", vbInformation
    End If
End With
'test if there is no account left in the database, then exit system
RunSql "Select * from tblAccountProfile"
If Rs.RecordCount = 0 Then
    MsgBox "All accounts has been deleted. The system will now exit", vbCritical
    'exit
    mdiMain.mnuExit_Click
End If
End Sub


Private Sub cmdRemove_Click()
imgProfile.Picture = LoadPicture()
ImgName = Empty
End Sub

Private Sub cmdSave_Click()
Dim MSG As String

'set functions in trappings
If TxtEmp(txtFname) = True Then Exit Sub
If TxtEmp(txtMname) = True Then Exit Sub
If TxtEmp(txtLname) = True Then Exit Sub
If CboEmp(cboGender) = True Then Exit Sub
If TxtEmp(txtBdate) = True Then Exit Sub: SelAll (txtBdate)
If TxtEmp(txtAddress) = True Then Exit Sub
If TxtEmp(txtDhired) = True Then Exit Sub
If CboEmp(cboEmpStatus) = True Then Exit Sub
If CboEmp(cboPosition) = True Then Exit Sub

'select an specified account
RunSql "Select * from tblAccountProfile where id = '" & lblId.Caption & "'"
With Rs
    'if a command button is Save then add new record
    'else update the record
    If cmdSave.Caption <> "&Update" Then
        MSG = "Account id " & lblId.Caption & " has been successfully added."
        .AddNew
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
        MSG = "Account id " & lblId.Caption & " has been successfully updated."
    End If
    Me.MousePointer = 11
    .Fields!ID = lblId.Caption
    .Fields!fname = txtFname.Text
    .Fields!mname = txtMname.Text
    .Fields!lname = txtLname.Text
    .Fields!age = Val(txtAge)
    .Fields!gender = cboGender.Text
    .Fields!bdate = txtBdate.Text
    .Fields!address = txtAddress.Text
    .Fields!Status = cboStatus.Text
    .Fields!contact = Val(txtContact)
    .Fields!date_hired = txtDhired.Text
    .Fields!Position = cboPosition.Text
    .Fields!remarks = cboRemarks.Text
    .Fields!emp_status = cboEmpStatus.Text
    .Fields!image_name = ImgName
    .Fields!date_reg = Format(Date, "mm/dd/yyyy")
    .Update
    i = MsgBox(MSG & " Do you want to set the security of this Account?", vbQuestion + vbYesNo)
    If i = vbYes Then
        frmAcntManage.cmdEdit_Click
        frmAcntManage.Show 1
    End If
    ClrFlds
End With
If FrstUsr = True Then
    Unload Me
End If
End Sub

Private Sub cmdSecurity_Click()
'show security manager and set security accounts
RunSql "Select * from tblAccountProfile where id = '" & lblId.Caption & "'"
If Rs.EOF = False Then
    frmAcntManage.txtId = lblId.Caption
    frmAcntManage.Show 1
Else
    frmAcntManage.Show 1
End If
End Sub

Private Sub dtpBdate_Change()
'format txtbdate as short date
txtBdate.Text = Format(dtpBdate.Value, "mm/dd/yyyy")
End Sub

Private Sub dtpDHired_Change()
'change format to short date
txtDhired.Text = Format(dtpDHired.Value, "mm/dd/yyyy")
End Sub

Private Sub Form_Load()
'pc name
lblPcId.Caption = PcId
ClrFlds

'add the saved postions from database
RunSql "Select * from tblAccountPosition"
While Not Rs.EOF = True
    cboPosition.AddItem (Rs.Fields!Description)
    Rs.MoveNext
Wend

RunSql "Select* from tblAccountProfile"
lblCount.Caption = Rs.RecordCount
lblId.Caption = RcrdId("tblAccountProfile", "NEW-", "id")
End Sub
Public Sub ClrFlds()
'clear all
Me.MousePointer = 0
dlgPic.FileName = Empty
ImgName = Empty
ImgSrc = Empty
imgProfile.Picture = LoadPicture()
txtFname.Text = Empty
txtMname.Text = Empty
txtLname.Text = Empty
txtBdate.Text = "  /  /    "
txtAge.Text = Empty
txtAddress.Text = Empty
txtContact.Text = Empty
txtDhired.Text = "  /  /    "
cboGender.ListIndex = 0
cboStatus.ListIndex = 0
cboPosition.ListIndex = 0
cboRemarks.ListIndex = 0
cboEmpStatus.ListIndex = 0
cmdSave.Caption = "&Save"
lblId.Caption = RcrdId("tblAccountProfile", "NEW-", "id")
End Sub

Private Sub txtAddress_LostFocus()
'the function that converts text to proper cases
txtAddress.Text = StrConv(txtAddress, vbProperCase)
End Sub

Private Sub txtBdate_Change()
'birth date conditions and trappings
txtAge.Text = Val(Format(Date, "YYYY")) - Val(Format(txtBdate, "YYYY"))

If Format(txtBdate, "MM") = Format(Date, "MM") Then
    If Format(txtBdate, "DD") > Format(Date, "DD") Then
        txtAge.Text = txtAge.Text - 1
    End If
End If

If Format(txtBdate, "MM") > Format(Date, "MM") Then
    txtAge.Text = txtAge.Text - 1
End If
End Sub

'---------------------highlight if got focus-------------
Private Sub txtBdate_GotFocus()
txtBdate.SelStart = 0
If txtBdate.Text <> "  /  /    " Then
    txtBdate.SelLength = Len(txtBdate)
End If
End Sub

Private Sub txtBdate_LostFocus()
If Val(txtAge.Text) <= 10 Then
    MsgBox "Invalid date: Birth date must be less than current date", vbExclamation
    txtBdate.Text = "  /  /    "
    txtAge.Text = Empty
End If
End Sub

Private Sub txtDhired_GotFocus()
txtDhired.SelStart = 0
If txtDhired.Text = "  /  /    " Then
    txtDhired.SelLength = Len(txtDhired)
End If
End Sub

Private Sub txtFname_LostFocus()
txtFname.Text = StrConv(txtFname, vbProperCase)
End Sub

Private Sub txtLname_Change()
lblId.Caption = RcrdId("tblAccountProfile", StrConv(Left(txtLname.Text, 3), vbUpperCase) & "-", "id")
End Sub

Private Sub txtLname_LostFocus()
txtLname.Text = StrConv(txtLname, vbProperCase)
End Sub

Private Sub txtMname_LostFocus()
txtMname.Text = StrConv(txtMname, vbProperCase)
End Sub
'------------------end hghlights----------
Private Sub txtSrchStr_Change()
If Trim(txtSrchStr.Text) <> Empty Then
    ExecSrch (txtSrchStr.Text)
Else
    ClrFlds
End If
End Sub
Public Sub ExecSrch(ByVal ID As String)
'execute the search for users
RunSql "Select * From tblAccountProfile where id Like '" & ID & "%'"
With Rs
'the system has found something
If .EOF = False Then
    txtFname.Text = .Fields!fname
    txtMname.Text = .Fields!mname
    txtLname.Text = .Fields!lname
    txtAge.Text = .Fields!age
    cboGender.Text = .Fields!gender
    txtAddress.Text = .Fields!address
    txtBdate.Text = Format(.Fields!bdate, "mm/dd/yyyy")
    cboStatus.Text = .Fields!Status
    txtContact.Text = .Fields!contact
    txtDhired.Text = Format(.Fields!date_hired, "mm/dd/yyyy")
    cboEmpStatus.Text = .Fields!emp_status
    cboPosition.Text = .Fields!Position
    cboRemarks.Text = .Fields!remarks
    If .Fields!image_name <> Empty Then
        imgProfile.Picture = LoadPicture(App.Path & "\Images\accounts\" & .Fields!image_name)
    Else
        imgProfile.Picture = LoadPicture()
    End If
    ImgName = .Fields!image_name
    cmdSave.Caption = "&Update"
    lblId.Caption = .Fields!ID
Else
    ClrFlds
End If
End With
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
    txtSrchStr.ForeColor = &H8000000B
End If
End Sub
