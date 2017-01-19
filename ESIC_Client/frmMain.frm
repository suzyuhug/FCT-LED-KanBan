VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "EStation - ESIC"
   ClientHeight    =   10710
   ClientLeft      =   60
   ClientTop       =   -4350
   ClientWidth     =   16080
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   16080
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   11640
      TabIndex        =   59
      Top             =   840
      Width           =   255
   End
   Begin VB.Frame frmCertifiedFail 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   120
      TabIndex        =   57
      Top             =   480
      Visible         =   0   'False
      Width           =   11055
      Begin VB.Label Label8 
         Caption         =   "Certified Fail (认证失败)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   3720
         TabIndex        =   58
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame fraOthers 
      Caption         =   "Information Center(信息中心)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12120
      TabIndex        =   49
      Top             =   10680
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton cmdInfoCenterRefresh 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdInfoCenterEnlarge 
         Caption         =   "<>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   375
      End
      Begin TabDlg.SSTab tabInfoCenter 
         Height          =   975
         Left            =   600
         TabIndex        =   52
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1720
         _Version        =   393216
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Tab 0"
         TabPicture(0)   =   "frmMain.frx":6852
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lstTab0"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Tab 1"
         TabPicture(1)   =   "frmMain.frx":686E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         TabCaption(2)   =   "Tab 2"
         TabPicture(2)   =   "frmMain.frx":688A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).ControlCount=   0
         TabCaption(3)   =   "Tab 3"
         TabPicture(3)   =   "frmMain.frx":68A6
         Tab(3).ControlEnabled=   0   'False
         Tab(3).ControlCount=   0
         Begin VB.ListBox lstTab0 
            BackColor       =   &H000000FF&
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   120
            TabIndex        =   53
            Top             =   480
            Width           =   3255
         End
      End
   End
   Begin VB.CommandButton cmdMin 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   11520
      TabIndex        =   48
      Top             =   5040
      Width           =   495
   End
   Begin VB.Frame FrameMain 
      Height          =   10575
      Left            =   12120
      TabIndex        =   21
      Top             =   0
      Width           =   4095
      Begin VB.Frame Frainfo 
         Caption         =   "Infomation Center(信息中心)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   54
         Top             =   7560
         Width           =   3735
         Begin VB.CommandButton Command1 
            Caption         =   "<>"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3240
            TabIndex        =   55
            Top             =   120
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label2 
            Height          =   135
            Left            =   1320
            TabIndex        =   56
            Top             =   120
            Visible         =   0   'False
            Width           =   3015
         End
      End
      Begin VB.Frame frmStationInfo 
         Height          =   2575
         Left            =   120
         TabIndex        =   41
         Top             =   7920
         Width           =   3735
         Begin VB.CommandButton cmdUnlockScreen 
            Height          =   735
            Left            =   120
            Picture         =   "frmMain.frx":68C2
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   240
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.CommandButton cmdQA 
            Height          =   735
            Left            =   960
            Picture         =   "frmMain.frx":8B9C
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   240
            Width           =   765
         End
         Begin VB.CommandButton cmdPC 
            Height          =   735
            Left            =   1800
            Picture         =   "frmMain.frx":8F11
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   240
            Width           =   765
         End
         Begin VB.CommandButton cmdInfoCenter 
            Height          =   735
            Left            =   2640
            Picture         =   "frmMain.frx":9353
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   240
            Width           =   765
         End
         Begin VB.CommandButton cmdLockScreen 
            Height          =   735
            Left            =   120
            Picture         =   "frmMain.frx":9515
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   240
            Width           =   765
         End
         Begin VB.Image Image1 
            Height          =   1435
            Left            =   2280
            Stretch         =   -1  'True
            Top             =   1000
            Width           =   1350
         End
         Begin VB.Label lblClient 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   46
            Top             =   1080
            Width           =   2775
         End
      End
      Begin VB.Timer Timer2 
         Interval        =   200
         Left            =   240
         Top             =   3120
      End
      Begin VB.Frame frmCertOperList 
         Caption         =   "Certified Operators List(已认证员工)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   39
         Top             =   5880
         Visible         =   0   'False
         Width           =   3735
         Begin VB.ListBox lstCertedOpsList 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   120
            TabIndex        =   32
            Top             =   235
            Width           =   3495
         End
      End
      Begin VB.Frame frmCertList 
         Caption         =   "Certification List(认证列表)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   120
         TabIndex        =   34
         Top             =   2760
         Width           =   3735
         Begin VB.CommandButton cmdCertEnd 
            Caption         =   "End Certification"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            TabIndex        =   27
            Top             =   360
            Width           =   1695
         End
         Begin VB.ListBox lstCertOpsList 
            Height          =   960
            ItemData        =   "frmMain.frx":B7EF
            Left            =   120
            List            =   "frmMain.frx":B7F1
            TabIndex        =   28
            Top             =   1320
            Width           =   3540
         End
         Begin VB.CommandButton cmdCertStart 
            Caption         =   "Start Certification"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtOPEmpNo 
            Height          =   375
            Left            =   80
            TabIndex        =   29
            Top             =   2520
            Width           =   1695
         End
         Begin VB.CommandButton cmdCertOpAdd 
            Caption         =   "Add"
            Height          =   375
            Left            =   1920
            TabIndex        =   30
            Top             =   2520
            Width           =   735
         End
         Begin VB.CommandButton cmdCertOpClear 
            Caption         =   "Clear"
            Height          =   375
            Left            =   2280
            TabIndex        =   35
            Top             =   2520
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.CommandButton cmdCertOpDelete 
            Caption         =   "Delete"
            Height          =   375
            Left            =   2760
            TabIndex        =   31
            Top             =   2520
            Width           =   735
         End
         Begin VB.Label lblCertStatus 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   3000
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Label Label3 
            Caption         =   "Certified Operators List(员工列表):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label lblCertResult 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2520
            TabIndex        =   36
            Top             =   960
            Width           =   1095
         End
      End
      Begin VB.Frame frmSICList 
         Caption         =   "SIC List(SIC列表)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2535
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   3735
         Begin VB.ListBox lstSIC 
            Height          =   1140
            ItemData        =   "frmMain.frx":B7F3
            Left            =   75
            List            =   "frmMain.frx":B7F5
            TabIndex        =   25
            Top             =   1200
            Width           =   3615
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "Close SIC"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            TabIndex        =   24
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdOpen 
            Caption         =   "Open SIC "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "SIC Open List(已经打开的SIC):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   840
            Width           =   3015
         End
      End
   End
   Begin VB.Frame FrmQAAudit 
      Caption         =   "QA Audit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   3240
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   4575
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtempid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtRemark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         TabIndex        =   16
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton CmdPass 
         Caption         =   "&Pass"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   15
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton CmdFail 
         Caption         =   "&Fail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   14
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Emp ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Remark:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Frame FrmNewVersion 
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   1920
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   7935
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   12
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Label lblRefreshMsg 
         Caption         =   "There are new version of the SIC released!  Please Reload it!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1815
         Left            =   1080
         TabIndex        =   11
         Top             =   1680
         Width           =   5775
      End
   End
   Begin VB.Frame frmCertified 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   11055
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   10080
         Top             =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Certified(认证完成)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   4440
         TabIndex        =   9
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame frmCertifing 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   10335
      Begin VB.Label Label4 
         Caption         =   "Certifing(认证中)..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   3720
         TabIndex        =   7
         Top             =   120
         Width           =   4095
      End
   End
   Begin VB.Frame frmCertNot 
      BorderStyle     =   0  'None
      Height          =   3175
      Left            =   3360
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Label lblCertLabel 
         Caption         =   "Data from SP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1380
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   3300
      End
   End
   Begin VB.Timer timerSck 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   6360
      Top             =   600
   End
   Begin VB.Timer TimerCert 
      Left            =   5760
      Top             =   600
   End
   Begin VB.Timer TimerMs 
      Left            =   5160
      Top             =   600
   End
   Begin VB.Timer TimerM1 
      Interval        =   60000
      Left            =   4560
      Top             =   600
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   10695
      Left            =   -840
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11175
      Begin VB.Label lbltemp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   9960
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblSIC 
         Caption         =   "Current SIC(当前SIC):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11055
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   10695
      Left            =   -240
      TabIndex        =   0
      Top             =   -120
      Width           =   11775
      ExtentX         =   20770
      ExtentY         =   18865
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu Schedule 
         Caption         =   "Schedule"
         Shortcut        =   ^S
      End
      Begin VB.Menu OpenSIC 
         Caption         =   "Open SIC"
         Shortcut        =   ^O
      End
      Begin VB.Menu CloseSIC 
         Caption         =   "Close SIC"
         Shortcut        =   ^C
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim gdSckStart As Date

Dim gdCertStart As Date
Dim gbCertStarted As Boolean
Dim psRefreshAction As String
  Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long

  Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
 
  
Private Type SHITEMID

          cb   As Long

          abID   As Byte

  End Type

  Private Type ITEMIDLIST

          mkid   As SHITEMID

  End Type


'Array to store OP certification list
'Column 1 -- SICID, default = "NOSIC"
'Column 2 -- Certified Status, default = "0", 0-Not certified, 1-Certified
'Column 3 - 20, Certified OP lists, default = ""
'Max 10 SICs
'Max 12 OPs can be certified for same SIC one time
Const iMaxSIC = 10
Const iMaxOPs = 10


Private Sub About_Click()
frmVersion.Show vbModal
End Sub

Private Sub CloseSIC_Click()
   Call cmdClose_Click
End Sub

'Public pSICCertList(10, 12) As String

Private Sub cmdCancel_Click()
     TxtEMPID.Text = ""
    txtRemark.Text = ""
    FrmQAAudit.Visible = False
End Sub

Private Sub cmdCertEnd_Click()
On Error GoTo Local_Error

    Dim i As Integer
    Dim lsSICID As String
    
    
    If LstSIC.ListIndex < 0 Then
        MsgBox "Please select SIC to be certified!" & vbCr & "请选择要认证的 SIC!", vbExclamation
        
        Exit Sub
    End If
    
    If lstCertOpsList.ListCount < 1 Then
        MsgBox "Please input operators to be certified!" & vbCr & "请选择要认证的员工!", vbExclamation
        
        Exit Sub
    End If
    
    If DateDiff("s", gdCertStart, Now()) < (gCertificateTime * 60) Then
        MsgBox "Certification time must larger than " & CStr(gCertificateTime) & " minutes!" & vbCr & "认证时间没有达到 " & CStr(gCertificateTime) & " 分钟!", vbExclamation
        
        Exit Sub
    End If
    
    If gbCertStarted = False Then
        MsgBox "Certification NOT started!" & vbCr & "没有开始认证!", vbExclamation
        
        Exit Sub
    End If
    
    lsSICID = CStr(LstSIC.ItemData(LstSIC.ListIndex))
    
    txtOPEmpNo.Text = ""
    cmdCertOpAdd.Enabled = True
    frmSICList.Enabled = True
    

' only some medical project can show record form'
Dim grstTemp As New ADODB.Recordset
gSQL = "exec SUZ_usp_CheckProjectCertify '" & gProject & "'"

If DB_Master_RecordSet(gSQL, grstTemp) Then
    If grstTemp(0) = "0" Then
            For i = 0 To lstCertOpsList.ListCount - 1
                'steven 2013 01 17, get all info from the list, and then show form '
                ESICID = gSICID
                ClientName = gClientName
                EmpID = Trim(Left(lstCertOpsList.List(i), InStr(lstCertOpsList.List(i), "-") - 1))
                CertifyBy = gEmpID
                frmCertifyRecord.Show
                
                
                gSQL = "exec usp_Certification_Add " & CStr(LstSIC.ItemData(LstSIC.ListIndex)) & ",'" & gClientName & "',"
                'lstCertOpsList.List (i)
                gSQL = gSQL & "'" & Trim(Left(lstCertOpsList.List(i), InStr(lstCertOpsList.List(i), "-") - 1)) & "',"
                gSQL = gSQL & "'" & Format(gdCertStart, "yyyy-mm-dd hh:mm") & "','" & Format(Now(), "yyyy-mm-dd hh:mm") & "','" & gEmpID & "'"
                If Not DB_Project_ExecSQL(gSQL) Then
                    MsgBox "Error Code: 3011, Connect to project database failed, usp_Certification_Add!", vbCritical
                End If
                
'                frmCertNot.Visible = False
'                frmCertified.Visible = True
'                frmCertifing.Visible = False
                Call SICCertList_Update(lsSICID, "1")
            
                lblCertStatus = "Certified at " + Format(Now(), "yyyy/mm/dd hh:mm")
                lblCertStatus.ForeColor = vbGreen
            
                gbCertStarted = False
                
                'steven 2013 01 17, get all info from the list, and then show form '
           Next i
    End If
End If
' only some medical project can show record form'



If grstTemp(0) <> "0" Then
        For i = 0 To lstCertOpsList.ListCount - 1
                gSQL = "exec usp_Certification_Add " & CStr(LstSIC.ItemData(LstSIC.ListIndex)) & ",'" & gClientName & "',"
                'lstCertOpsList.List (i)
                gSQL = gSQL & "'" & Trim(Left(lstCertOpsList.List(i), InStr(lstCertOpsList.List(i), "-") - 1)) & "',"
                gSQL = gSQL & "'" & Format(gdCertStart, "yyyy-mm-dd hh:mm") & "','" & Format(Now(), "yyyy-mm-dd hh:mm") & "','" & gEmpID & "'"
                If Not DB_Project_ExecSQL(gSQL) Then
                    MsgBox "Error Code: 3011, Connect to project database failed, usp_Certification_Add!", vbCritical
                End If
        Next i
    
        frmCertNot.Visible = False
        frmCertified.Visible = True
        frmCertifing.Visible = False
        
        Call SICCertList_Update(lsSICID, "1")
    
        lblCertStatus = "Certified at " + Format(Now(), "yyyy/mm/dd hh:mm")
        lblCertStatus.ForeColor = vbGreen
    
        gbCertStarted = False
End If

    Exit Sub

Local_Error:
    MsgBox "Error Code: " & CStr(Err.Number) & " , " & Err.Description, vbCritical
End Sub

Private Sub OriCertify()

    Dim i As Integer
    Dim lsSICID As String
    
    
    If LstSIC.ListIndex < 0 Then
        MsgBox "Please select SIC to be certified!" & vbCr & "请选择要认证的 SIC!", vbExclamation
        
        Exit Sub
    End If
    
    If lstCertOpsList.ListCount < 1 Then
        MsgBox "Please input operators to be certified!" & vbCr & "请选择要认证的员工!", vbExclamation
        
        Exit Sub
    End If
    
    If DateDiff("s", gdCertStart, Now()) < (gCertificateTime * 60) Then
        MsgBox "Certification time must larger than " & CStr(gCertificateTime) & " minutes!" & vbCr & "认证时间没有达到 " & CStr(gCertificateTime) & " 分钟!", vbExclamation
        
        Exit Sub
    End If
    
    If gbCertStarted = False Then
        MsgBox "Certification NOT started!" & vbCr & "没有开始认证!", vbExclamation
        
        Exit Sub
    End If
    
    lsSICID = CStr(LstSIC.ItemData(LstSIC.ListIndex))
    
    txtOPEmpNo.Text = ""
    cmdCertOpAdd.Enabled = True
    frmSICList.Enabled = True
    
        For i = 0 To lstCertOpsList.ListCount - 1
                gSQL = "exec usp_Certification_Add " & CStr(LstSIC.ItemData(LstSIC.ListIndex)) & ",'" & gClientName & "',"
                'lstCertOpsList.List (i)
                gSQL = gSQL & "'" & Trim(Left(lstCertOpsList.List(i), InStr(lstCertOpsList.List(i), "-") - 1)) & "',"
                gSQL = gSQL & "'" & Format(gdCertStart, "yyyy-mm-dd hh:mm") & "','" & Format(Now(), "yyyy-mm-dd hh:mm") & "','" & gEmpID & "'"
                If Not DB_Project_ExecSQL(gSQL) Then
                    MsgBox "Error Code: 3011, Connect to project database failed, usp_Certification_Add!", vbCritical
                End If
        Next i
    
        frmCertNot.Visible = False
        frmCertified.Visible = True
        frmCertifing.Visible = False
        
        Call SICCertList_Update(lsSICID, "1")
    
        lblCertStatus = "Certified at " + Format(Now(), "yyyy/mm/dd hh:mm")
        lblCertStatus.ForeColor = vbGreen
    
        gbCertStarted = False
End Sub

Private Sub cmdCertOpAdd_Click()
    Dim strOPEmpNo, strTemp As String
    Dim MaxOfCertificater As Integer
    If LstSIC.ListIndex < 0 Then
        MsgBox "Please open esic first!"
        Exit Sub
    End If
    txtOPEmpNo.Text = Trim(txtOPEmpNo.Text)
    strOPEmpNo = Trim(txtOPEmpNo.Text)
    If IsNull(strOPEmpNo) Or strOPEmpNo = "" Then
        MsgBox "Please input Employee No.!", vbExclamation
        
        Exit Sub
    End If
    Dim rstTemp As ADODB.Recordset
    Set rstTemp = New ADODB.Recordset
    Dim Project As String
    Dim ProjectID As Integer
    Project = LstSIC.Text
    
    Project = Trim(Mid(Project, 1, InStr(1, Project, "-") - 1))
    gSQL = "exec Project_GetByName " & Project
    If DB_Master_RecordSet(gSQL, rstTemp) Then
        If Not rstTemp.EOF Then
             ProjectID = rstTemp(0)
        Else
             MsgBox "Error Code 3009: Connect to Master database failed, Project_GetByName!!!"
        End If
    Else
     MsgBox "Error Code 3009: Connect to Master database failed, Project_GetByName!!!"
    End If
    'Check right
    Set rstTemp = New ADODB.Recordset
    gSQL = "exec usp_Check_CertRight_Station " & CStr(ProjectID) & ",'" & gEmpID & "','" & CStr(LstSIC.ItemData(LstSIC.ListIndex)) & "'"
    If DB_Project_RecordSet(gSQL, rstTemp) Then
        If rstTemp.Fields("result") = 1 Then
            MsgBox "Error Code 4002: You have not the right to certificate!!!"
            Exit Sub
        End If
    End If
    'check user added by terry
     Set rstTemp = New ADODB.Recordset
    gSQL = "exec usp_Check_CertEmpid " & CStr(ProjectID) & ",'" & gEmpID & "','" & txtOPEmpNo.Text & "'"
    If DB_Project_RecordSet(gSQL, rstTemp) Then
        If rstTemp.Fields("result") = 1 Then
            MsgBox "Error Code 4003: The user can not be certified!!!"
            Exit Sub
        End If
    End If
    
    Set rstTemp = New ADODB.Recordset
    gSQL = "exec SystemRule_Get " & CStr(ProjectID) & ",4"
        
    If DB_Project_RecordSet(gSQL, rstTemp) Then
        If rstTemp.EOF Then
            MsgBox "Error Code 4004: Project rule MaxOfCertificater NOT configured in system!!!"
        Else
            MaxOfCertificater = CInt(rstTemp.Fields("ProjectValue"))
        End If
            
        If rstTemp.State = adStateOpen Then rstTemp.Close
    Else
        MsgBox "Error Code 3010: Connect to project database failed, SystemRule_Get MaxOfCertificater!!!"
    End If
    
    If lstCertOpsList.ListCount = MaxOfCertificater Then
        MsgBox "Can not add certificater again", vbInformation
        
        Exit Sub
    End If
    
    Set rstTemp = New ADODB.Recordset
    gSQL = "exec usp_LCD_GetHREmployeeInfo '" & strOPEmpNo & "'"

    If DB_Master_RecordSet(gSQL, rstTemp) Then
        If rstTemp.EOF Then
            If rstTemp.State = adStateOpen Then
                rstTemp.Close
            End If
            
            MsgBox "Invalid employee no!", vbExclamation
            
            Exit Sub
        Else
            strTemp = strOPEmpNo + " - " + rstTemp.Fields("English_Name") + " - " + rstTemp.Fields("Chinese_Name") + " - " + rstTemp.Fields("Admin_Dept")
            lstCertOpsList.AddItem strTemp
        End If
        
        If rstTemp.State = adStateOpen Then
            rstTemp.Close
        End If
        
        txtOPEmpNo.Text = ""
    Else
        MsgBox "Error Code 3009: Connect to Master database failed, usp_LCD_GetHREmployeeInfo!!!"
    End If
    
End Sub

Private Sub cmdCertOpClear_Click()
    lstCertOpsList.Clear
End Sub

Private Sub cmdCertOpDelete_Click()
    If lstCertOpsList.ListIndex >= 0 Then
        lstCertOpsList.RemoveItem (lstCertOpsList.ListIndex)
    End If
End Sub

Private Sub cmdCertStart_Click()
    If LstSIC.ListIndex < 0 Then
        MsgBox "Please select SIC to be certified!" & vbCr & "请选择要认证的 SIC!", vbExclamation
        
        Exit Sub
    End If
    
    If lstCertOpsList.ListCount < 1 Then
        MsgBox "Please input operators to be certified!" & vbCr & "请选择要认证的员工!", vbExclamation
        
        Exit Sub
    End If

    If gbCertStarted = True Then
        If MsgBox("Certification already started, Restart it?" & vbCr & "认证已开始. 要重新开始吗?", vbYesNo) = vbNo Then
            Exit Sub
        Else
            gbCertStarted = False
            
            
            
                '2016 04 11, Steven remark the "Not Certify"
                Dim rstNotCertify As New ADODB.Recordset
                If DB_Master_RecordSet("exec SUZ_NotCertify '" & gProject & "','" & gEmpID & "'", rstNotCertify) Then
                    If rstNotCertify.Fields("Result") = 1 Then
                        frmCertNot.Visible = True
                        lblCertLabel = rstNotCertify.Fields("Des")
                    End If
                End If
                
                
            'frmCertNot.Visible = True
            frmCertified.Visible = False
            frmCertifing.Visible = False
    
            txtOPEmpNo.Text = ""
            cmdCertOpAdd.Enabled = True
            cmdCertOpClear.Enabled = True
            cmdCertOpDelete.Enabled = True
            
            Exit Sub
        End If
    Else
    
    End If
    
    frmCertNot.Visible = False
    frmCertified.Visible = False
    frmCertifing.Visible = True
    
    gdCertStart = Now()
    'Format(Now(), "yyyy/mm/dd hh:mm:ss")
    frmSICList.Enabled = False
    gbCertStarted = True
    
    txtOPEmpNo.Text = ""
    cmdCertOpAdd.Enabled = False
    cmdCertOpClear.Enabled = False
    cmdCertOpDelete.Enabled = False
    
    lblCertStatus = "Starting at " + Format(Now(), "yyyy/mm/dd hh:mm")
    lblCertStatus.ForeColor = vbBlue
    
End Sub

Private Sub cmdClose_Click()


'        ' This is for test only
'        'steven 2013 01 17, get all info from the list, and then show form '
'        ESICID = gSICID
'        ClientName = gClientName
'        EmpID = "32456"
'        CertifyBy = gEmpID
'        frmCertifyRecord.Show
'        'steven 2013 01 17, get all info from the list, and then show form '

        
        
    Call Do_Action("SIC_CLOSE,N/A")
    
    
End Sub

Private Sub CmdFail_Click()
    Dim gSQL As String
    If Trim(txtRemark.Text) = "" Then
        MsgBox "Please input the remark !"
        Exit Sub
    End If
    If Trim(TxtEMPID.Text) = "" Then
        MsgBox "Please input the employee no!"
        Exit Sub
    Else
        Dim rstTemp As New ADODB.Recordset
        gSQL = "exec usp_LCD_GetHREmployeeInfo '" & Trim(TxtEMPID.Text) & "'"

        If DB_Master_RecordSet(gSQL, rstTemp) Then
        If rstTemp.EOF Then
            If rstTemp.State = adStateOpen Then
                rstTemp.Close
            End If
            
            MsgBox "Invalid employee no!", vbExclamation
            
            Exit Sub
        End If
        End If
    End If
    
    gSQL = "exec usp_QAAudit_Add '" & CStr(Trim(TxtEMPID.Text)) & "','" & CStr(txtRemark.Text) & "','" & CStr(Now()) & "',1"
        If Not DB_Project_ExecSQL(gSQL) Then
            MsgBox "Error Code: 3011, Connect to project database failed, usp_QAAudit_Add!", vbCritical
        End If
        
    TxtEMPID.Text = ""
    txtRemark.Text = ""
    FrmQAAudit.Visible = False
    
End Sub

Private Sub cmdInfoCenter_Click()
    frmQuery.Show
End Sub

Private Sub cmdInfoCenterEnlarge_Click()
    If cmdInfoCenterEnlarge.Caption = "<>" Then
        cmdInfoCenterEnlarge.Caption = "><"
        fraOthers.Height = fraOthers.Height * 2
        fraOthers.Width = fraOthers.Width * 2
        fraOthers.Left = Me.Width - fraOthers.Width - 100
        tabInfoCenter.Width = tabInfoCenter.Width * 2
        tabInfoCenter.Height = tabInfoCenter.Height * 2
        lstTab0.Width = lstTab0.Width * 2
        lstTab0.Height = lstTab0.Height * 2
        tabInfoCenter.Left = (fraOthers.Width - tabInfoCenter.Width) / 2
        tabInfoCenter.Top = (fraOthers.Height - tabInfoCenter.Height) / 2
    Else
        fraOthers.Height = fraOthers.Height / 2
        cmdInfoCenterEnlarge.Caption = "<>"
        fraOthers.Width = fraOthers.Width / 2
        fraOthers.Left = FrameMain.Left + frmSICList.Left
        tabInfoCenter.Width = tabInfoCenter.Width / 2
        tabInfoCenter.Height = tabInfoCenter.Height / 2
        lstTab0.Width = lstTab0.Width / 2
        lstTab0.Height = lstTab0.Height / 2
        tabInfoCenter.Left = (fraOthers.Width - tabInfoCenter.Width) / 2 + 150
        tabInfoCenter.Top = (fraOthers.Height - tabInfoCenter.Height) / 2
    End If
End Sub

Private Sub cmdLockScreen_Click()
    frmCertList.Enabled = False
    cmdOpen.Enabled = False
    cmdClose.Enabled = False
    
    cmdCertStart.Enabled = False
    cmdCertEnd.Enabled = False
    
    cmdCertOpAdd.Enabled = False
    cmdCertOpClear.Enabled = False
    cmdCertOpDelete.Enabled = False
    
    If gClientQueryButtonEnable = "YES" Then
        cmdPC.Enabled = True
        cmdQA.Enabled = True
        cmdInfoCenter.Enabled = True
    Else
        cmdPC.Enabled = False
        cmdQA.Enabled = False
        cmdInfoCenter.Enabled = False
    End If
    
    frmMain.Caption = "EStation - ESIC (Locked)"
    
    cmdLockScreen.Visible = False
    cmdUnlockScreen.Visible = True
    cmdUnlockScreen.SetFocus
End Sub

Private Sub cmdMin_Click()
    If cmdMin.Caption = ">>" Then

'        FrameMain.Width = 370
'         WebBrowser1.Width = Me.Width - FrameMain.Width - 100
'        frmCertifing.Width = Me.Width - FrameMain.Width - 100
'        Frame1.Width = Me.Width - FrameMain.Width - 100
'        frmCertified.Width = Me.Width - FrameMain.Width - 100
'        FrameMain.Left = WebBrowser1.Width + 50
'
'        frmCertified.Width = 15015 * (Me.Width / 15360)
'        cmdMin.Caption = "<<"
'            fraOthers.Left = FrameMain.Left + frmSICList.Left
        FrameMain.Width = 350
        WebBrowser1.Width = frmMain.Width - FrameMain.Width
        Frame1.Width = WebBrowser1.Width
        frmCertified.Width = WebBrowser1.Width
        FrmNewVersion.Left = (WebBrowser1.Width - FrmNewVersion.Width) / 2
        FrmQAAudit.Left = (WebBrowser1.Width - FrmQAAudit.Width) / 2
        frmCertNot.Left = (WebBrowser1.Width - frmCertNot.Width) / 2
        FrameMain.Left = WebBrowser1.Width + 270
        cmdMin.Caption = "<<"
        cmdMin.Left = WebBrowser1.Width - 200
    Else

'        FrameMain.Width = 4335
'        WebBrowser1.Width = Me.Width - FrameMain.Width - 100
'        frmCertifing.Width = Me.Width - FrameMain.Width - 100
'        Frame1.Width = Me.Width - FrameMain.Width - 100
'        frmCertified.Width = Me.Width - FrameMain.Width - 100
'         FrmNewVersion.Left = 1560 * (Me.Width / 15360)
'        frmCertNot.Left = 3000 * (Me.Width / 15360)
'
'        FrameMain.Left = WebBrowser1.Width + 50
'            fraOthers.Left = FrameMain.Left + frmSICList.Left
'
'        cmdMin.Caption = ">>"
         FrameMain.Width = 4335
        WebBrowser1.Width = frmMain.Width - FrameMain.Width
        Frame1.Width = WebBrowser1.Width
        frmCertified.Width = WebBrowser1.Width
        FrmNewVersion.Left = (WebBrowser1.Width - FrmNewVersion.Width) / 2
        FrmQAAudit.Left = (WebBrowser1.Width - FrmQAAudit.Width) / 2
        frmCertNot.Left = (WebBrowser1.Width - frmCertNot.Width) / 2
        FrameMain.Left = WebBrowser1.Width + 270
        cmdMin.Caption = ">>"
        cmdMin.Left = WebBrowser1.Width - 200
    End If
    
    
End Sub

Private Sub cmdOpen_Click()
     Dim rstTemp As New ADODB.Recordset
     gSQL = "exec usp_GetClient_DefaultWindow '" & gLineName & "','" & gClientName & "','" & gEmpID & "'"

        If DB_Master_RecordSet(gSQL, rstTemp) Then
        
        If rstTemp(0) = 0 Then
            frmESICSelect.Show vbModal
        Else
            frmESICGet.Show vbModal
        End If
        
        End If
        
      
End Sub

Private Sub CmdReload_Click()
'empty lisT
End Sub

Private Sub CmdPass_Click()
    Dim gSQL As String
    
    If Trim(TxtEMPID.Text) = "" Then
        MsgBox "Please input the employee no!"
        Exit Sub
    Else
        Dim rstTemp As New ADODB.Recordset
        gSQL = "exec usp_LCD_GetHREmployeeInfo '" & Trim(TxtEMPID.Text) & "'"

        If DB_Master_RecordSet(gSQL, rstTemp) Then
        If rstTemp.EOF Then
            If rstTemp.State = adStateOpen Then
                rstTemp.Close
            End If
            
            MsgBox "Invalid employee no!", vbExclamation
            
            Exit Sub
        End If
        End If
    End If
    
    gSQL = "exec usp_QAAudit_Add '" & CStr(Trim(TxtEMPID.Text)) & "','" & CStr(txtRemark.Text) & "','" & CStr(Now()) & "',0"
        If Not DB_Project_ExecSQL(gSQL) Then
            MsgBox "Error Code: 3011, Connect to project database failed, usp_QAAudit_Add!", vbCritical
        End If
        
    TxtEMPID.Text = ""
    txtRemark.Text = ""
    FrmQAAudit.Visible = False
    
End Sub

Private Sub cmdPC_Click()
    frmInfoCenter.Show
End Sub

Private Sub cmdQA_Click()
    If LstSIC.ListCount > 0 Then
    
        FrmQAAudit.Visible = True
    End If
    
End Sub

Private Sub CmdRefresh_Click()
    
    FrmNewVersion.Visible = False
    frmCertList.Enabled = True
    cmdOpen.Enabled = True
    cmdClose.Enabled = True
    
    cmdCertStart.Enabled = True
    cmdCertEnd.Enabled = True
            
    cmdCertOpAdd.Enabled = True
    cmdCertOpClear.Enabled = True
    cmdCertOpDelete.Enabled = True
    
    frmSICList.Enabled = True
    
    Call Do_Action(psRefreshAction)
    
    psRefreshAction = "NA,NA"
    
End Sub

Private Sub cmdSendData_Click()
'    Winsock1.SendData txtSockDataOut.Text
End Sub

Private Sub cmdUnlockScreen_Click()
    Dim rstUser As ADODB.Recordset
    
    gInputBoxTitle = "Please input password:"
    gInputBoxPwdChar = "*"
    
    frmInputBox.Show vbModal

    If gInputBoxResult = "INPUTBOX_CANCEL" Then
        Exit Sub
    End If
    
    Unload frmInputBox
    
    Set rstUser = New ADODB.Recordset
    
    If DB_Master_RecordSet("exec usp_Login_Validate '" & gEmpID & "','" & gInputBoxResult & "'", rstUser) Then
        If rstUser.EOF Then
            rstUser.Close
            MsgBox "Error Code: 2001, Invalid user/password!", vbExclamation
        Else
            rstUser.Close
            
            frmCertList.Enabled = True
            cmdOpen.Enabled = True
            cmdClose.Enabled = True
           
            
            cmdCertStart.Enabled = True
            cmdCertEnd.Enabled = True
            
            
            cmdCertOpAdd.Enabled = True
            cmdCertOpClear.Enabled = True
            cmdCertOpDelete.Enabled = True
            
            frmMain.Caption = "EStation - ESIC"
            
            cmdUnlockScreen.Visible = False
            cmdLockScreen.Visible = True
            
            cmdLockScreen.SetFocus
             cmdQA.Enabled = True
            cmdPC.Enabled = True
            cmdInfoCenter.Enabled = True
        End If
    Else
        MsgBox "Error Code: 3001, Connect to Master database failed, usp_Login_Validate!!!", vbExclamation
    End If


End Sub



Private Sub Command2_Click()
If Not WebBrowser1.Busy Then
    If InStr(WebBrowser1.Document.documentelement.innerText, "Flextronics") Then
        MsgBox = "找到了"
    Else
        MsgBox = "未找到"
    End If
End If

End Sub

Private Sub Form_Activate()
On Error GoTo Local_Error
'    If (WebBrowser1.Width <> Me.Width - FrameMain.Width - 100) Then
'        WebBrowser1.Height = Me.Height - 500
'
'
'        WebBrowser1.Width = Me.Width - FrameMain.Width - 100
'        lblSIC.Width = WebBrowser1.Width
'
'
'        Frame1.Width = WebBrowser1.Width
'        frmCertified.Width = WebBrowser1.Width
'
'        frmCertifing.Width = WebBrowser1.Width
'        FrameMain.Height = WebBrowser1.Height
'        FrameMain.Left = WebBrowser1.Width + 50
'      '  FrmNewVersion.Left = (WebBrowser1.Width - FrmNewVersion.Width) / 2
'      '   FrmQAAudit.Left = (WebBrowser1.Width - FrmQAAudit.Width) / 2
'      '    frmCertNot.Left = (WebBrowser1.Width - frmCertNot.Width) / 2
'          cmdMin.Left = WebBrowser1.Width + 2
'
'    End If
    
    WebBrowser1.Height = frmMain.Height - 1100
    FrameMain.Height = WebBrowser1.Height
    'FrameMain.Height = Me.Height - 1500
    WebBrowser1.Width = frmMain.Width - FrameMain.Width - 20
    Frame1.Width = WebBrowser1.Width
    frmCertified.Width = WebBrowser1.Width
    FrmNewVersion.Left = (WebBrowser1.Width - FrmNewVersion.Width) / 2
    FrmQAAudit.Left = (WebBrowser1.Width - FrmQAAudit.Width) / 2
    frmCertNot.Left = (WebBrowser1.Width - frmCertNot.Width) / 2
    cmdMin.Left = WebBrowser1.Width - 200
    cmdMin.Top = WebBrowser1.Height / 2 - cmdMin.Height / 2
    FrameMain.Left = WebBrowser1.Width + 270
'    If FrameMain.Height < (frmStationInfo.Top + frmStationInfo.Height + 20) Then
'       Frainfo.Top = FrameMain.Height - (Frainfo.Height + frmStationInfo.Height + 30)
'       frmStationInfo.Top = Frainfo.Top + Frainfo.Height
        frmStationInfo.Top = FrameMain.Height - frmStationInfo.Height - 30
        Frainfo.Top = frmStationInfo.Top - Frainfo.Height - 20
'    End If
    
    
    Dim i As Integer
    Dim bSICOpened As Boolean
    
    If IsNull(gURL) Then
        gURL = ""
    End If
    
    Dim test As String
    
    test = "Open SIC" & vbCr & "打开SIC"
    cmdOpen.Caption = test
    test = "Close SIC" & vbCr & "关闭SIC"
    cmdClose.Caption = test
    test = "Start Certification" & vbCr & "开始认证"
    cmdCertStart.Caption = test
    test = "End Certification" & vbCr & "结束认证"
    cmdCertEnd.Caption = test
    test = "Add" & vbCr & "添加"
    cmdCertOpAdd.Caption = test
    test = "Delete" & vbCr & "删除"
    cmdCertOpDelete.Caption = test
    
    With tabInfoCenter
        tabInfoCenter.Tab = 0
        tabInfoCenter.ForeColor = vbRed
        tabInfoCenter.Tab = 1
        tabInfoCenter.ForeColor = vbBlack
        '.Tab(0).TabVisible = False
        
    End With
    
    Dim IDL     As ITEMIDLIST
    Dim pathname As String
     
    pathname = Space(512)
    SHGetSpecialFolderLocation 100, &H0, IDL
    SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal pathname
    pathname = Split(pathname, Chr$(0))(0)
    
     Dim rs As New ADODB.Recordset
        
        If DB_Master_RecordSet("exec usp_Client_LoadPic '" & gClientName & "','" & gEmpID & "'", rs) Then
        If Not rs.EOF Then
            If (rs(0) <> "0") Then
                 Image1.Visible = True
                Dim StmPic As New ADODB.Stream
                StmPic.Type = adTypeBinary
                StmPic.Open
                StmPic.Write rs(1)              '写入数据库中的数据至Stream中
                StmPic.SaveToFile pathname & "\TEMP.JPG", adSaveCreateOverWrite
                StmPic.Close
              
                Image1.Picture = LoadPicture(pathname & "\TEMP.JPG")
               Else
            Image1.Visible = False
            End If
          
        End If
        End If
'    If gDebug = True Then
'        frmSocket.Visible = True
'    Else
'        frmSocket.Visible = False
'    End If
    'update refresh time
    'If gRefreshTime > 0 Then
        'If Timer1.Interval > gRefreshTime * 60 * 1000 Then
            'Timer1.Interval = gRefreshTime * 60 * 1000
        'End If
    'End If
    
    bSICOpened = False
    
    If Not IsNull(gURL) And gURL <> "" And gURL <> "NO_CHANGE" Then
        For i = 0 To LstSIC.ListCount - 1
            If gSICID = LstSIC.ItemData(i) Then
                bSICOpened = True
                
                Exit For
            End If
        Next i
        
        If bSICOpened = False Then
            
            LstSIC.AddItem gSIC
            LstSIC.ItemData(LstSIC.ListCount - 1) = CLng(gSICID)
            
            LstSIC.Text = gSIC
            lblSIC.Caption = "Current SIC: " + gSIC
            
            SICCertList_Add (gSICID)
            
            WebBrowser1.Stop
            WebBrowser1.Navigate gURL
            
            'if the user is operator ,not certficated do not display
            Dim gSQL As String
            Dim rstTemp As New ADODB.Recordset

            gSQL = "exec usp_Check_Operator '" & gEmpID & "','" & gSICID & "'"
            If DB_Project_RecordSet(gSQL, rstTemp) Then
                    If Not rstTemp.EOF Then
                        frmCertNot.Visible = False
                        frmCertified.Visible = False
                        frmCertifing.Visible = False
                    Else
                    
                            '2016 04 11, Steven remark the "Not Certify"
                            Dim rstNotCertify As New ADODB.Recordset
                            If DB_Master_RecordSet("exec SUZ_NotCertify '" & gProject & "','" & gEmpID & "'", rstNotCertify) Then
                                If rstNotCertify.Fields("Result") = 1 Then
                                    frmCertNot.Visible = True
                                    lblCertLabel = rstNotCertify.Fields("Des")
                                End If
                            End If
                            
                            
                        'frmCertNot.Visible = True
                        frmCertified.Visible = False
                        frmCertifing.Visible = False
                    End If
                    
            End If

            
            'frmCertNot.Visible = True
            'frmCertified.Visible = False
            'frmCertifing.Visible = False
        End If
    End If
    
    gURL = ""
    
    Exit Sub

Local_Error:
    If InStr(Err.Description, "MSWINSCK.OCX") > 0 Then
        MsgBox "Error Code 1003: MSWINSCK.OCX Not Installed!!!", vbExclamation
    Else
        MsgBox "Error Code: " & CStr(Err.Number) & " , " & Err.Description, vbCritical
    End If

End Sub

Private Sub Form_Load()
On Error GoTo Local_Error

    gbCertStarted = False
    lblCertStatus.Caption = "NOT Certified"
    lblCertStatus.ForeColor = vbRed
    
    lblClient.Caption = "Client: " & gClientName & vbCr & "Server: " & gMasterServer & vbCr & "User: " & gEmpID & vbCr & "Line: " & gLineName & vbCr & "&Help"

    Call SICCertList_Initiate
    
    Timer1.Interval = 60000
    Timer1.Tag = "0"
    
    WebBrowser1.Width = 11415
    frmCertifing.Width = 11415
    Frame1.Width = 11415
    frmCertified.Width = 11415
    frmCertList.Visible = False
'    Winsock1.Protocol = sckTCPProtocol
'    Winsock1.LocalPort = 50
'    Winsock1.Listen
    
    LstSIC.Clear
    WebBrowser1.Navigate (gMasterDefaultURL)
    lstCertOpsList.Clear
    Dim rs As New ADODB.Recordset
        
'    If DB_Master_RecordSet("exec usp_GetClient_AutoOpen '" & gEmpID & "','" & gClientName & "'", rs) Then
'    If rs(0) = 0 Then
'           cmdOpen_Click
'    End If
'    End If
        

    
    Exit Sub
    

Local_Error:
    If InStr(Err.Description, "MSWINSCK.OCX") > 0 Then
        MsgBox "Error Code 1003: MSWINSCK.OCX Not Installed!!!", vbExclamation
    Else
        MsgBox "Error Code: " & CStr(Err.Number) & " , " & Err.Description, vbCritical
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    DoEvents
End Sub


Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Local_Error

'    Winsock1.Close
'
'    WebBrowser1.Stop
    Dim rstCheckStatus As New ADODB.Recordset
    
    If Not DB_Master_RecordSet("exec usp_Login_CheckStatus '" & gEmpID & "','" & gClientName & "',1", rstCheckStatus) Then
         MsgBox "Error Code: 3001, Connect to Master database failed, usp_Login_CheckStatus!!!", vbExclamation
    End If
    
    End
    
Local_Error:
    End
End Sub

Private Sub lblClient_DblClick()
    frmVersion.Show vbModal
    
End Sub


Public Sub lstSIC_Click()
On Error GoTo Local_Error
    Dim rstTemp As New ADODB.Recordset
    
    Set rstTemp = New ADODB.Recordset
    Dim strURL As String
    
    
    If LstSIC.ListCount > 0 Then
        If Right(LstSIC.Text, 5) = "(OLD)" Then
                Exit Sub
            End If
            Set rstTemp = New ADODB.Recordset
            gSQL = "exec usp_CheckESICUpdate_ByID " & CStr(LstSIC.ItemData(LstSIC.ListIndex)) & "," & gLineID & ",'" & gClientName & "','" & gEmpID & "'"
            If DB_Project_RecordSet(gSQL, rstTemp) Then
                    If Not rstTemp.EOF Then
                        If rstTemp(0) <> "0" Then
                        
                            FrmNewVersion.Visible = True
                            lblSIC.Caption = "Current SIC: "
                            frmCertList.Enabled = False
                            cmdOpen.Enabled = False
                            cmdClose.Enabled = False
                
                            cmdCertStart.Enabled = False
                            cmdCertEnd.Enabled = False
                        
                            cmdCertOpAdd.Enabled = False
                            cmdCertOpClear.Enabled = False
                            cmdCertOpDelete.Enabled = False
                            frmSICList.Enabled = False

                            CmdRefresh.Caption = IIf(IsNull(rstTemp("Button")), "Refresh", rstTemp("Button"))
                            lblRefreshMsg.Caption = IIf(IsNull(rstTemp("Msg")), "Need Refreh SIC", rstTemp("Msg"))
                            psRefreshAction = IIf(IsNull(rstTemp("Action_Code")), "NA,NA", rstTemp("Action_Code"))
                            'Call Do_Action(IIf(IsNull(rstTemp("Action_Code")), "NA,NA", rstTemp("Action_Code")))
                            
                            'Have new rev
  
                        End If
                        
                        'lstSic.
                    End If
            Else
                MsgBox "Error Code 3012: Connect to Project database failed, usp_CheckESICUpdate_ByID!!!", vbCritical
            End If
            
    
'        If Right(lstSIC.Text, 5) <> "(OLD)" Then
'
'            gSQL = "exec usp_CheckNewRevESIC_ByID " & CStr(lstSIC.ItemData(lstSIC.ListIndex))
'            If DB_Project_RecordSet(gSQL, rsttemp) Then
'                If Not rsttemp.EOF Then
'                    'Have new rev
'                    If rsttemp(0) <> "0" Then
'
'                        FrmNewVersion.Visible = True
'                        lblSIC.Caption = "Current SIC: "
'                        frmCertList.Enabled = False
'                        cmdOpen.Enabled = False
'                        cmdClose.Enabled = False
'
'                        cmdCertStart.Enabled = False
'                        cmdCertEnd.Enabled = False
'
'                        cmdCertOpAdd.Enabled = False
'                        cmdCertOpClear.Enabled = False
'                        cmdCertOpDelete.Enabled = False
'                        frmSICList.Enabled = True
'
'                    End If
'                    'lstSic.
'                 End If
'            End If
'        End If
    End If
    
    If LstSIC.ListIndex >= 0 Then
        Set rstTemp = New ADODB.Recordset
        
        If Replace(lblSIC.Caption, "Current SIC: ", "") = LstSIC.Text Then
            Exit Sub
        Else
            lblSIC.Caption = "Current SIC: " + LstSIC.Text
        End If
        
        gSQL = "exec usp_GetESIC_ByID " & CStr(LstSIC.ItemData(LstSIC.ListIndex))
        If DB_Project_RecordSet(gSQL, rstTemp) Then
            If rstTemp.EOF Then
                rstTemp.Close
                
                MsgBox "Error Code 5001: Can NOT find SIC in system!", vbExclamation
                
                Exit Sub
            End If
            
            If IsNull(rstTemp.Fields("FileUrl")) Or rstTemp.Fields("FileUrl") = "" Then
                MsgBox "Error Code 5002: SIC file NOT configured in system!", vbExclamation
                    
                Exit Sub
            End If
            

          gProject = LstSIC.Text
    
          gProject = Trim(Mid(gProject, 1, InStr(1, gProject, "-") - 1))
        Dim rs As New ADODB.Recordset
        
        If DB_Master_RecordSet("exec usp_GetProject_Info '" & gProject & "'", rs) Then
        If Not rs.EOF Then
            gProjectWebURL = IIf(IsNull(rs.Fields("ProjectWebURL")), "", Trim(rs.Fields("ProjectWebURL")))
        End If
        End If
        frmCertNot.Visible = False
        frmCertified.Visible = False
        frmCertifing.Visible = False
       
        Set rs = New ADODB.Recordset
        gSQL = "exec usp_Check_CertRight_ByEmpid  '" & gEmpID & "','" & CStr(LstSIC.ItemData(LstSIC.ListIndex)) & "'"
        If DB_Project_RecordSet(gSQL, rs) Then
            If rs.Fields("result") = 1 Then
               frmCertList.Visible = False
               frmCertOperList.Top = 3120
               frmCertOperList.Height = Frainfo.Top - frmCertOperList.Top - 3
               lstCertedOpsList.Height = Frainfo.Top - frmCertOperList.Top - 100
            Else
                frmCertList.Visible = True
                'frmCertOperList.Top = 6600
                frmCertOperList.Top = frmCertList.Top + frmCertList.Height + 100
                frmCertOperList.Height = 1815
                frmCertOperList.Height = Frainfo.Top - frmCertList.Top - frmCertList.Height - 100
                'lstCertedOpsList.Height = 1230
                lstCertedOpsList.Height = frmCertOperList.Height - 500
                
                Call CertOpsList_Refresh(CStr(LstSIC.ItemData(LstSIC.ListIndex)))
            End If
        Else
            frmCertList.Visible = True
            frmCertOperList.Top = 6600
            frmCertOperList.Height = 1815
            lstCertedOpsList.Height = 1230
            Call CertOpsList_Refresh(CStr(LstSIC.ItemData(LstSIC.ListIndex)))
        
        End If
        If LstSIC.ItemData(LstSIC.ListIndex) > 1000000 Then
            strURL = Replace(gProjectWebURL, LCase(Trim(gProject)), "generalsic") + rstTemp.Fields("FileUrl")
        Else
            strURL = gProjectWebURL + rstTemp.Fields("FileUrl")
        End If
            
            
            Call CertedOpsList_Refresh(CStr(LstSIC.ItemData(LstSIC.ListIndex)))
            Call Do_Action("SIC_OPEN," & strURL)
        Else
            MsgBox "Error Code 3005: Connect to project database failed, usp_GetESIC_ByID", vbExclamation
        End If
    End If
    
    Exit Sub
Local_Error:
    MsgBox "Error Code: " & CStr(Err.Number) & " , " & Err.Description, vbCritical
End Sub


Private Sub OpenSIC_Click()
  Call cmdOpen_Click
End Sub

Private Sub Schedule_Click()
  FrmArrangeSchedule.Show
End Sub

Private Sub Timer1_Timer()
    'exec SP to check Current rev
   Dim i As Integer
   
    If gRefreshTime = 0 Then
        gRefreshTime = 5
    End If
    
    'psRefreshAction = "NA,NA"
    
   If Val(Timer1.Tag) > gRefreshTime Then
        Timer1.Tag = "0"
        
        Dim rstTemp As ADODB.Recordset
        'insert current esic data
  
        For i = 0 To LstSIC.ListCount - 1
            If i = LstSIC.ListIndex Then
                gSQL = "exec usp_InsertCurrentESIC_ByClient " & CStr(LstSIC.ItemData(i)) & ",'" & gClientName & "',1,'" & gEmpID & "'"
            Else
                gSQL = "exec usp_InsertCurrentESIC_ByClient " & CStr(LstSIC.ItemData(i)) & ",'" & gClientName & "',0,'" & gEmpID & "'"
            End If
            If Not DB_Project_ExecSQL(gSQL) Then
                MsgBox "Error Code 1005: Insert Client Fail usp_InsertCurrentESIC_ByClient!!!", vbInformation
                Exit Sub
            End If
        Next i
        
        
        Set rstTemp = New ADODB.Recordset
        If LstSIC.ListCount > 0 Then
            If Right(LstSIC.Text, 5) = "(OLD)" Then
                Exit Sub
            End If
            
            gSQL = "exec usp_CheckESICUpdate_ByID " & CStr(LstSIC.ItemData(LstSIC.ListIndex)) & "," & gLineID & ",'" & gClientName & "','" & gEmpID & "'"
            If DB_Project_RecordSet(gSQL, rstTemp) Then
                    If Not rstTemp.EOF Then
                        If rstTemp(0) <> "0" Then
                        
                            FrmNewVersion.Visible = True
                            lblSIC.Caption = "Current SIC: "
                            frmCertList.Enabled = False
                            cmdOpen.Enabled = False
                            cmdClose.Enabled = False
                
                            cmdCertStart.Enabled = False
                            cmdCertEnd.Enabled = False
                        
                            cmdCertOpAdd.Enabled = False
                            cmdCertOpClear.Enabled = False
                            cmdCertOpDelete.Enabled = False
                            frmSICList.Enabled = False
                            
                            CmdRefresh.Caption = IIf(IsNull(rstTemp("Button")), "Refresh", rstTemp("Button"))
                            lblRefreshMsg.Caption = IIf(IsNull(rstTemp("Msg")), "Need Refreh SIC", rstTemp("Msg"))
                            psRefreshAction = IIf(IsNull(rstTemp("Action_Code")), "NA,NA", rstTemp("Action_Code"))
                            'Call Do_Action(IIf(IsNull(rstTemp("Action_Code")), "NA,NA", rstTemp("Action_Code")))
                            
                            'Have new rev
  
                        End If
                        
                        'lstSic.
                    End If
            Else
                MsgBox "Error Code 3012: Connect to Project database failed, usp_CheckESICUpdate_ByID!!!", vbCritical
            End If
        End If
    Else
        Timer1.Tag = Trim(str(Val(Timer1.Tag) + 1))
    End If
    
End Sub

Private Sub Timer2_Timer()
If Dir(App.Path & "\SICStart.txt") <> "" Then
Kill App.Path & "\SICStart.txt"
gstartid = True
Me.WindowState = 2
frmMain.Show
Me.SetFocus
Dim inifile As String
Dim str As String
inifile = App.Path & "\SICSetting.ini"
str = GetPrivateStringValue("setting", "Project", inifile) + " - " + GetPrivateStringValue("setting", "Station", inifile) + " - " + GetPrivateStringValue("setting", "Model", inifile) + " - " + GetPrivateStringValue("setting", "Side", inifile) + " - " + GetPrivateStringValue("setting", "Category", inifile)
Dim bo As Boolean
Dim i As Integer
For i = 0 To LstSIC.ListCount - 1
If LstSIC.List(i) = str Then
LstSIC.ListIndex = i
bo = True
Exit For
End If
Next
If bo = False Then
cmdOpen_Click
End If
End If
End Sub

Private Sub timerSck_Timer()
'    If Winsock1.State = sckConnected Then
'        If DateDiff("s", gdSckStart, Now()) > 30 Then
'            If Winsock1.State <> sckClosed Then
'                Winsock1.Close
'                Winsock1.Listen
'                lblSckState = "Socket State:" + CStr(Winsock1.State)
'            End If
'        End If
'    End If
End Sub


Private Sub txtempid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdPass_Click
    End If
End Sub

Private Sub txtOPEmpNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdCertOpAdd_Click
    End If
End Sub

Private Sub txtRemark_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdPass_Click
    End If
End Sub

Private Sub Winsock1_Close()
'    Winsock1.Close
'    Winsock1.Listen
'    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
'
'    lblSckState = "Socket State:" + CStr(Winsock1.State)
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
'    If Winsock1.State <> sckClosed Then
'        Winsock1.Close
'    End If
'
'    Winsock1.Accept requestID
'    DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents: DoEvents
'
'    gdSckStart = Now()
'
'    lblSckState = "Socket State:" + CStr(Winsock1.State)
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'    Dim s As String
'
'    Sleep 1000
'
'    Winsock1.GetData s
'    txtSockDataIn.Text = s
'
'    Call Do_Action(s)
'
'    lblSckState = "Socket State:" + CStr(Winsock1.State)
End Sub

Private Sub Do_Action(ByVal strActionString As String)
    Dim strAction, strParValue As String
    Dim i As Integer
    
    i = InStr(strActionString, ",")
    If i > 0 Then
        strAction = Mid(strActionString, 1, i - 1)
        strParValue = Mid(strActionString, i + 1)
    Else
        strAction = strActionString
    End If
    
    Select Case strAction
    Case "SIC_CLEAR":
        Call Do_Action_SIC_Clear
    Case "SIC_OPEN":
        Call Do_Action_SIC_Open(strParValue)
    Case "SIC_RELOAD":
        Call Do_Action_SIC_Reload
    Case "SIC_CLOSE":
        Call Do_Action_SIC_Close
'    Case "Service_Complete":
'        Winsock1.Close
'
'        Winsock1.LocalPort = 50
'        Winsock1.Listen
    End Select
End Sub

Private Sub Do_Action_SIC_Clear()
    LstSIC.Clear
    frmCertNot.Visible = False
    frmCertified.Visible = False
    frmCertifing.Visible = False
    WebBrowser1.Stop
    WebBrowser1.Navigate (gMasterDefaultURL)
End Sub

Private Sub Do_Action_SIC_Open(ByVal strURL As String)
    WebBrowser1.Stop
    WebBrowser1.Navigate (strURL)
End Sub

Private Sub Do_Action_SIC_Reload()
    Dim rstTemp As ADODB.Recordset
    Dim OldSicid As String
    Dim NewSicid As String
    Set rstTemp = New ADODB.Recordset
    
    Dim i, j As Integer
    
    OldSicid = LstSIC.ItemData(LstSIC.ListIndex)
     
    gSQL = "exec usp_GetNewESICID_ByID " & CStr(OldSicid)
    If DB_Project_RecordSet(gSQL, rstTemp) Then
    If Not rstTemp.EOF Then
        NewSicid = rstTemp.Fields("ESicID")
    End If
    
    LstSIC.ItemData(LstSIC.ListIndex) = NewSicid
    
    For i = 0 To 9
        If pSICCertList(i, 0) = OldSicid Then
            pSICCertList(i, 0) = NewSicid
            pSICCertList(i, 1) = "0"
            For j = 2 To 11
                pSICCertList(i, j) = ""
            Next j
        End If
        
    Next i
    lblSIC.Caption = "Current SIC: "
    lstSIC_Click
    End If
End Sub

Private Sub Do_Action_SIC_Close()
     Dim i, j As Integer
     
     If LstSIC.ListIndex < 0 Then
        Exit Sub
    End If
    
    For i = 0 To 9
        If pSICCertList(i, 0) = CStr(LstSIC.ItemData(LstSIC.ListIndex)) Then
            pSICCertList(i, 0) = "NOSIC"
            For j = 1 To 11
                pSICCertList(i, j) = ""
            Next j
            
        End If
        
    Next i
    LstSIC.RemoveItem (LstSIC.ListIndex)
   
    If LstSIC.ListCount = 0 Then
        LstSIC.Clear
        
        frmCertNot.Visible = False
        frmCertifing.Visible = False
        frmCertified.Visible = False
        frmCertList.Visible = True
        lstCertedOpsList.Clear
        frmCertOperList.Visible = False
        frmCertList.Visible = False
        
        
        lblSIC.Caption = "Current SIC: "
        
        WebBrowser1.Navigate (gMasterDefaultURL)
        lstCertOpsList.Clear
        
        
        
    Else
        LstSIC.ListIndex = 0
        lstSIC_Click
    End If

End Sub

Private Sub SICCertList_Initiate()
    Dim i, j As Integer
    
    For i = 0 To iMaxSIC - 1
        pSICCertList(i, 0) = "NOSIC"
        pSICCertList(i, 1) = "0"
        For j = 2 To 11
            pSICCertList(i, j) = ""
        Next j
    Next i
End Sub

Private Sub SICCertList_Add(ByVal strSICID As String)
    Dim i, j As Integer
    
    For i = 0 To iMaxSIC - 1
        If (pSICCertList(i, 0) = "NOSIC") Then
            pSICCertList(i, 0) = strSICID
            pSICCertList(i, 1) = "0"
            
            Exit For
        End If
    Next i
End Sub

Private Sub SICCertList_Update(ByVal strSICID As String, ByVal strSICStatus As String)
    Dim i, j As Integer
    Dim jMax As Integer
    
    For i = 0 To iMaxSIC - 1
        If (pSICCertList(i, 0) = strSICID) Then
            pSICCertList(i, 1) = strSICStatus
            
            If strSICStatus = "1" Then
                jMax = iMaxOPs
                If lstCertOpsList.ListCount < jMax Then
                    jMax = lstCertOpsList.ListCount
                End If
                
                If lstCertOpsList.ListCount > 0 Then
                    For j = 0 To jMax - 1
                        pSICCertList(i, j + 2) = lstCertOpsList.List(j)
                    Next j
                    
                    For j = jMax To iMaxOPs - 1
                        pSICCertList(i, j + 2) = ""
                    Next j
                End If
            End If
            
            Exit For
        End If
    Next i
End Sub

Private Sub CertOpsList_Refresh(ByVal strSICID As String)
    Dim i, j As Integer
    lstCertOpsList.Clear
    
    For i = 0 To iMaxSIC - 1
        If (pSICCertList(i, 0) = strSICID) Then
            'Certified
            If pSICCertList(i, 1) = "1" Then
                    frmCertNot.Visible = False
                    frmCertified.Visible = True
                    frmCertifing.Visible = False
                    
                    
                    For j = 0 To iMaxOPs - 1
                        If pSICCertList(i, 2 + j) = "" Or IsNull(pSICCertList(i, 2 + j)) Then
                            Exit For
                        End If
                        
                        lstCertOpsList.AddItem (pSICCertList(i, 2 + j))
                    Next j
            Else
                        '2016 04 11, Steven remark the "Not Certify"
                        Dim rstNotCertify As New ADODB.Recordset
                        If DB_Master_RecordSet("exec SUZ_NotCertify '" & gProject & "','" & gEmpID & "'", rstNotCertify) Then
                            If rstNotCertify.Fields("Result") = 1 Then
                                frmCertNot.Visible = True
                                lblCertLabel = rstNotCertify.Fields("Des")
                            End If
                        End If
                        'frmCertNot.Visible = True
                    frmCertified.Visible = False
                    frmCertifing.Visible = False
            End If
            
            Exit For
        End If
    Next i
End Sub

Private Sub CertedOpsList_Refresh(ByVal strSICID As String)
    Dim i, j As Integer
    'lstCertedOpsList.Visible = False
    frmCertOperList.Visible = False
    
    lstCertedOpsList.Clear
    'diplay
    Dim rstTemp As New ADODB.Recordset
    Dim gSQL As String
    
    'gSQL = "exec usp_CheckTempSICListDisplay " & CStr(strSICID)
   ' If DB_Project_RecordSet(gSQL, rstTemp) Then
     '   If rstTemp.EOF Then
      '      lbltemp.Caption = ""
      '  Else
      '      lbltemp.Caption = rstTemp.Fields("Content")
      '  End If
        
   ' End If
    Set rstTemp = New ADODB.Recordset
    
    gSQL = "exec usp_CheckCertedOpsListDisplay " & CStr(strSICID)
    If DB_Project_RecordSet(gSQL, rstTemp) Then
            If Not rstTemp.EOF Then
                If rstTemp.Fields("rulevalue") = "0" Then
                     frmCertOperList.Visible = False
                Else
                      
                    frmCertOperList.Visible = True
                    'update lstCertedOpsList
                   
                    Set rstTemp = New ADODB.Recordset
                     gSQL = "exec usp_Get_CertedOpsList " & CStr(strSICID)
                     
                     
                     If DB_Project_RecordSet(gSQL, rstTemp) Then
                        lstCertedOpsList.Clear
                        
                        While Not rstTemp.EOF
            
                            lstCertedOpsList.AddItem (rstTemp.Fields("employee"))
                            
                
                            rstTemp.MoveNext
                            
                        Wend
                        
                     End If
                End If
                
                  
               
                
            End If
    Else
            MsgBox "Error Code 3012: Connect to Project database failed, usp_CheckCertedOpsListDisplay!!!", vbCritical
    End If
    
    
End Sub





