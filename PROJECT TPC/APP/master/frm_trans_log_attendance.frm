VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Begin VB.Form frm_trans_log_attendance 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LOG ATTENDANCE (BASE IP)"
   ClientHeight    =   9090
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   14685
   Icon            =   "frm_trans_log_attendance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_downloading 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   1980
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   7935
      Begin VB.Label lbl_downloading 
         AutoSize        =   -1  'True
         Caption         =   "Downloading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   1890
      End
   End
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   465
      Left            =   10470
      OleObjectBlob   =   "frm_trans_log_attendance.frx":000C
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   2685
      Left            =   3360
      TabIndex        =   10
      Top             =   2430
      Visible         =   0   'False
      Width           =   5655
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   270
         TabIndex        =   11
         Top             =   2100
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Processing, Please Wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   960
         TabIndex        =   3
         Top             =   300
         Width           =   2730
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   345
         Left            =   300
         TabIndex        =   12
         Top             =   1740
         Width           =   2475
      End
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   7
      Top             =   240
      Width           =   3855
   End
   Begin VB.Frame fra_button_control 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   14175
      Begin VB.CommandButton cmd_delete_log 
         Caption         =   "D&elete Log"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   2220
         Picture         =   "frm_trans_log_attendance.frx":0030
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   3240
         Picture         =   "frm_trans_log_attendance.frx":0459
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_download 
         Caption         =   "&Download"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1200
         Picture         =   "frm_trans_log_attendance.frx":09E3
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   120
         Top             =   360
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6735
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11880
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "COMP. CODE"
      Columns(0).DataField=   "company_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "COMP. NAME"
      Columns(1).DataField=   "company_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "DEPT. CODE"
      Columns(2).DataField=   "department_code"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DEPT. NAME"
      Columns(3).DataField=   "department_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DIV. CODE"
      Columns(4).DataField=   "division_code"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DIV. NAME"
      Columns(5).DataField=   "division_name"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "EMP. CODE"
      Columns(6).DataField=   "employee_code"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "EMP. NAME"
      Columns(7).DataField=   "employee_name"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "NICK NAME"
      Columns(8).DataField=   "employee_nick_name"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "TITLE CODE"
      Columns(9).DataField=   "title_code"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "TITLE NAME"
      Columns(10).DataField=   "title_name"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "DATE"
      Columns(11).DataField=   "att_date"
      Columns(11).NumberFormat=   "FormatText Event"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "CHECK IN"
      Columns(12).DataField=   "time_in"
      Columns(12).NumberFormat=   "FormatText Event"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "CHECK OUT"
      Columns(13).DataField=   "time_out"
      Columns(13).NumberFormat=   "FormatText Event"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   4
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "LATE"
      Columns(14).DataField=   "flag_late"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   4
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "EARLY"
      Columns(15).DataField=   "flag_early"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   1
      Splits(0).Size  =   4004.788
      Splits(0).Size.vt=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1958"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1879"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3916"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3836"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2064"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1984"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=3545"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3466"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1508"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1429"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1508"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1429"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=1588"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1508"
      Splits(0)._ColumnProps(34)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(37)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=1588"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=1508"
      Splits(0)._ColumnProps(42)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(44)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(45)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(50)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(52)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(53)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(55)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(58)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(59)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(60)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(61)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(62)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(63)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(64)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(65)=   "Column(10)._MinWidth=79702332"
      Splits(0)._ColumnProps(66)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(67)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(68)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(69)=   "Column(11).AllowSizing=0"
      Splits(0)._ColumnProps(70)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(71)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(72)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(73)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(74)=   "Column(11)._MinWidth=-1"
      Splits(0)._ColumnProps(75)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(76)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(77)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(78)=   "Column(12).AllowSizing=0"
      Splits(0)._ColumnProps(79)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(80)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(81)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(82)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(83)=   "Column(12)._MinWidth=110292176"
      Splits(0)._ColumnProps(84)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(85)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(87)=   "Column(13).AllowSizing=0"
      Splits(0)._ColumnProps(88)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(89)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(90)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(91)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(92)=   "Column(13)._MinWidth=110264768"
      Splits(0)._ColumnProps(93)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(94)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(96)=   "Column(14).AllowSizing=0"
      Splits(0)._ColumnProps(97)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(98)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(99)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(100)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(101)=   "Column(14)._MinWidth=110296448"
      Splits(0)._ColumnProps(102)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(103)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(104)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(105)=   "Column(15).AllowSizing=0"
      Splits(0)._ColumnProps(106)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(107)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(108)=   "Column(15).AllowFocus=0"
      Splits(0)._ColumnProps(109)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(110)=   "Column(15)._MinWidth=110287984"
      Splits(1)._UserFlags=   0
      Splits(1).Size  =   2
      Splits(1).Size.vt=   2
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).ScrollBars=   3
      Splits(1).DividerColor=   13160660
      Splits(1).FilterBar=   -1  'True
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=16"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1826"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=1826"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=1746"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=516"
      Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(17)=   "Column(2).Width=1720"
      Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=1640"
      Splits(1)._ColumnProps(20)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._ColStyle=516"
      Splits(1)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(1)._ColumnProps(23)=   "Column(2).AllowFocus=0"
      Splits(1)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(25)=   "Column(3).Width=1720"
      Splits(1)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(27)=   "Column(3)._WidthInPix=1640"
      Splits(1)._ColumnProps(28)=   "Column(3).AllowSizing=0"
      Splits(1)._ColumnProps(29)=   "Column(3)._ColStyle=516"
      Splits(1)._ColumnProps(30)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(31)=   "Column(3).AllowFocus=0"
      Splits(1)._ColumnProps(32)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(33)=   "Column(4).Width=1508"
      Splits(1)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(4)._WidthInPix=1429"
      Splits(1)._ColumnProps(36)=   "Column(4).AllowSizing=0"
      Splits(1)._ColumnProps(37)=   "Column(4)._ColStyle=516"
      Splits(1)._ColumnProps(38)=   "Column(4).Visible=0"
      Splits(1)._ColumnProps(39)=   "Column(4).AllowFocus=0"
      Splits(1)._ColumnProps(40)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(41)=   "Column(4)._MinWidth=80002672"
      Splits(1)._ColumnProps(42)=   "Column(5).Width=1508"
      Splits(1)._ColumnProps(43)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(44)=   "Column(5)._WidthInPix=1429"
      Splits(1)._ColumnProps(45)=   "Column(5).AllowSizing=0"
      Splits(1)._ColumnProps(46)=   "Column(5)._ColStyle=516"
      Splits(1)._ColumnProps(47)=   "Column(5).Visible=0"
      Splits(1)._ColumnProps(48)=   "Column(5).AllowFocus=0"
      Splits(1)._ColumnProps(49)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(50)=   "Column(5)._MinWidth=80001968"
      Splits(1)._ColumnProps(51)=   "Column(6).Width=1826"
      Splits(1)._ColumnProps(52)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(53)=   "Column(6)._WidthInPix=1746"
      Splits(1)._ColumnProps(54)=   "Column(6)._ColStyle=516"
      Splits(1)._ColumnProps(55)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(56)=   "Column(6)._MinWidth=80000960"
      Splits(1)._ColumnProps(57)=   "Column(7).Width=3678"
      Splits(1)._ColumnProps(58)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(7)._WidthInPix=3598"
      Splits(1)._ColumnProps(60)=   "Column(7)._ColStyle=516"
      Splits(1)._ColumnProps(61)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(62)=   "Column(7)._MinWidth=79999936"
      Splits(1)._ColumnProps(63)=   "Column(8).Width=2725"
      Splits(1)._ColumnProps(64)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(65)=   "Column(8)._WidthInPix=2646"
      Splits(1)._ColumnProps(66)=   "Column(8).AllowSizing=0"
      Splits(1)._ColumnProps(67)=   "Column(8)._ColStyle=516"
      Splits(1)._ColumnProps(68)=   "Column(8).Visible=0"
      Splits(1)._ColumnProps(69)=   "Column(8).AllowFocus=0"
      Splits(1)._ColumnProps(70)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(71)=   "Column(8)._MinWidth=80007280"
      Splits(1)._ColumnProps(72)=   "Column(9).Width=2725"
      Splits(1)._ColumnProps(73)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(74)=   "Column(9)._WidthInPix=2646"
      Splits(1)._ColumnProps(75)=   "Column(9).AllowSizing=0"
      Splits(1)._ColumnProps(76)=   "Column(9)._ColStyle=516"
      Splits(1)._ColumnProps(77)=   "Column(9).Visible=0"
      Splits(1)._ColumnProps(78)=   "Column(9).AllowFocus=0"
      Splits(1)._ColumnProps(79)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(80)=   "Column(10).Width=2725"
      Splits(1)._ColumnProps(81)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(82)=   "Column(10)._WidthInPix=2646"
      Splits(1)._ColumnProps(83)=   "Column(10)._ColStyle=516"
      Splits(1)._ColumnProps(84)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(85)=   "Column(11).Width=2170"
      Splits(1)._ColumnProps(86)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(87)=   "Column(11)._WidthInPix=2090"
      Splits(1)._ColumnProps(88)=   "Column(11)._ColStyle=513"
      Splits(1)._ColumnProps(89)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(90)=   "Column(11)._MinWidth=110189680"
      Splits(1)._ColumnProps(91)=   "Column(12).Width=1879"
      Splits(1)._ColumnProps(92)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(93)=   "Column(12)._WidthInPix=1799"
      Splits(1)._ColumnProps(94)=   "Column(12)._ColStyle=513"
      Splits(1)._ColumnProps(95)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(96)=   "Column(13).Width=1773"
      Splits(1)._ColumnProps(97)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(98)=   "Column(13)._WidthInPix=1693"
      Splits(1)._ColumnProps(99)=   "Column(13)._ColStyle=513"
      Splits(1)._ColumnProps(100)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(101)=   "Column(14).Width=1482"
      Splits(1)._ColumnProps(102)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(103)=   "Column(14)._WidthInPix=1402"
      Splits(1)._ColumnProps(104)=   "Column(14)._ColStyle=513"
      Splits(1)._ColumnProps(105)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(106)=   "Column(14)._MinWidth=14"
      Splits(1)._ColumnProps(107)=   "Column(15).Width=1455"
      Splits(1)._ColumnProps(108)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(109)=   "Column(15)._WidthInPix=1376"
      Splits(1)._ColumnProps(110)=   "Column(15)._ColStyle=513"
      Splits(1)._ColumnProps(111)=   "Column(15).Order=16"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "LOG OF ATTENDANCE"
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.alignment=0,.bgcolor=&H80000002&"
      _StyleDefs(8)   =   ":id=4,.fgcolor=&H80000009&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(9)   =   ":id=4,.strikethrough=0,.charset=0"
      _StyleDefs(10)  =   ":id=4,.fontname=Tahoma"
      _StyleDefs(11)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(12)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(13)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(14)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(15)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(16)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(17)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&H80000002&,.fgcolor=&H80000009&"
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(24)  =   ":id=14,.fgcolor=&H80000002&"
      _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=86,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=82,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=78,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=74,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=70,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=98,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=94,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=91,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=92,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=93,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=106,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=103,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=104,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=105,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=234,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=231,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=232,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=233,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=242,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=239,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=240,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=241,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=250,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=247,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=248,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=249,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=258,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=255,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=256,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=257,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=266,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=263,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=264,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=265,.parent=17"
      _StyleDefs(98)  =   "Splits(1).Style:id=123,.parent=1"
      _StyleDefs(99)  =   "Splits(1).CaptionStyle:id=132,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(100) =   ":id=132,.fgcolor=&H80000009&"
      _StyleDefs(101) =   "Splits(1).HeadingStyle:id=124,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(102) =   ":id=124,.fgcolor=&H80000002&"
      _StyleDefs(103) =   "Splits(1).FooterStyle:id=125,.parent=3"
      _StyleDefs(104) =   "Splits(1).InactiveStyle:id=126,.parent=5"
      _StyleDefs(105) =   "Splits(1).SelectedStyle:id=128,.parent=6"
      _StyleDefs(106) =   "Splits(1).EditorStyle:id=127,.parent=7"
      _StyleDefs(107) =   "Splits(1).HighlightRowStyle:id=129,.parent=8"
      _StyleDefs(108) =   "Splits(1).EvenRowStyle:id=130,.parent=9"
      _StyleDefs(109) =   "Splits(1).OddRowStyle:id=131,.parent=10"
      _StyleDefs(110) =   "Splits(1).RecordSelectorStyle:id=133,.parent=11"
      _StyleDefs(111) =   "Splits(1).FilterBarStyle:id=134,.parent=12"
      _StyleDefs(112) =   "Splits(1).Columns(0).Style:id=138,.parent=123"
      _StyleDefs(113) =   "Splits(1).Columns(0).HeadingStyle:id=135,.parent=124"
      _StyleDefs(114) =   "Splits(1).Columns(0).FooterStyle:id=136,.parent=125"
      _StyleDefs(115) =   "Splits(1).Columns(0).EditorStyle:id=137,.parent=127"
      _StyleDefs(116) =   "Splits(1).Columns(1).Style:id=142,.parent=123"
      _StyleDefs(117) =   "Splits(1).Columns(1).HeadingStyle:id=139,.parent=124"
      _StyleDefs(118) =   "Splits(1).Columns(1).FooterStyle:id=140,.parent=125"
      _StyleDefs(119) =   "Splits(1).Columns(1).EditorStyle:id=141,.parent=127"
      _StyleDefs(120) =   "Splits(1).Columns(2).Style:id=146,.parent=123"
      _StyleDefs(121) =   "Splits(1).Columns(2).HeadingStyle:id=143,.parent=124"
      _StyleDefs(122) =   "Splits(1).Columns(2).FooterStyle:id=144,.parent=125"
      _StyleDefs(123) =   "Splits(1).Columns(2).EditorStyle:id=145,.parent=127"
      _StyleDefs(124) =   "Splits(1).Columns(3).Style:id=150,.parent=123"
      _StyleDefs(125) =   "Splits(1).Columns(3).HeadingStyle:id=147,.parent=124"
      _StyleDefs(126) =   "Splits(1).Columns(3).FooterStyle:id=148,.parent=125"
      _StyleDefs(127) =   "Splits(1).Columns(3).EditorStyle:id=149,.parent=127"
      _StyleDefs(128) =   "Splits(1).Columns(4).Style:id=154,.parent=123"
      _StyleDefs(129) =   "Splits(1).Columns(4).HeadingStyle:id=151,.parent=124"
      _StyleDefs(130) =   "Splits(1).Columns(4).FooterStyle:id=152,.parent=125"
      _StyleDefs(131) =   "Splits(1).Columns(4).EditorStyle:id=153,.parent=127"
      _StyleDefs(132) =   "Splits(1).Columns(5).Style:id=158,.parent=123"
      _StyleDefs(133) =   "Splits(1).Columns(5).HeadingStyle:id=155,.parent=124"
      _StyleDefs(134) =   "Splits(1).Columns(5).FooterStyle:id=156,.parent=125"
      _StyleDefs(135) =   "Splits(1).Columns(5).EditorStyle:id=157,.parent=127"
      _StyleDefs(136) =   "Splits(1).Columns(6).Style:id=162,.parent=123"
      _StyleDefs(137) =   "Splits(1).Columns(6).HeadingStyle:id=159,.parent=124"
      _StyleDefs(138) =   "Splits(1).Columns(6).FooterStyle:id=160,.parent=125"
      _StyleDefs(139) =   "Splits(1).Columns(6).EditorStyle:id=161,.parent=127"
      _StyleDefs(140) =   "Splits(1).Columns(7).Style:id=166,.parent=123"
      _StyleDefs(141) =   "Splits(1).Columns(7).HeadingStyle:id=163,.parent=124"
      _StyleDefs(142) =   "Splits(1).Columns(7).FooterStyle:id=164,.parent=125"
      _StyleDefs(143) =   "Splits(1).Columns(7).EditorStyle:id=165,.parent=127"
      _StyleDefs(144) =   "Splits(1).Columns(8).Style:id=230,.parent=123"
      _StyleDefs(145) =   "Splits(1).Columns(8).HeadingStyle:id=227,.parent=124"
      _StyleDefs(146) =   "Splits(1).Columns(8).FooterStyle:id=228,.parent=125"
      _StyleDefs(147) =   "Splits(1).Columns(8).EditorStyle:id=229,.parent=127"
      _StyleDefs(148) =   "Splits(1).Columns(9).Style:id=202,.parent=123"
      _StyleDefs(149) =   "Splits(1).Columns(9).HeadingStyle:id=199,.parent=124"
      _StyleDefs(150) =   "Splits(1).Columns(9).FooterStyle:id=200,.parent=125"
      _StyleDefs(151) =   "Splits(1).Columns(9).EditorStyle:id=201,.parent=127"
      _StyleDefs(152) =   "Splits(1).Columns(10).Style:id=206,.parent=123"
      _StyleDefs(153) =   "Splits(1).Columns(10).HeadingStyle:id=203,.parent=124"
      _StyleDefs(154) =   "Splits(1).Columns(10).FooterStyle:id=204,.parent=125"
      _StyleDefs(155) =   "Splits(1).Columns(10).EditorStyle:id=205,.parent=127"
      _StyleDefs(156) =   "Splits(1).Columns(11).Style:id=238,.parent=123,.alignment=2"
      _StyleDefs(157) =   "Splits(1).Columns(11).HeadingStyle:id=235,.parent=124"
      _StyleDefs(158) =   "Splits(1).Columns(11).FooterStyle:id=236,.parent=125"
      _StyleDefs(159) =   "Splits(1).Columns(11).EditorStyle:id=237,.parent=127"
      _StyleDefs(160) =   "Splits(1).Columns(12).Style:id=246,.parent=123,.alignment=2"
      _StyleDefs(161) =   "Splits(1).Columns(12).HeadingStyle:id=243,.parent=124"
      _StyleDefs(162) =   "Splits(1).Columns(12).FooterStyle:id=244,.parent=125"
      _StyleDefs(163) =   "Splits(1).Columns(12).EditorStyle:id=245,.parent=127"
      _StyleDefs(164) =   "Splits(1).Columns(13).Style:id=254,.parent=123,.alignment=2"
      _StyleDefs(165) =   "Splits(1).Columns(13).HeadingStyle:id=251,.parent=124"
      _StyleDefs(166) =   "Splits(1).Columns(13).FooterStyle:id=252,.parent=125"
      _StyleDefs(167) =   "Splits(1).Columns(13).EditorStyle:id=253,.parent=127"
      _StyleDefs(168) =   "Splits(1).Columns(14).Style:id=262,.parent=123,.alignment=2"
      _StyleDefs(169) =   "Splits(1).Columns(14).HeadingStyle:id=259,.parent=124"
      _StyleDefs(170) =   "Splits(1).Columns(14).FooterStyle:id=260,.parent=125"
      _StyleDefs(171) =   "Splits(1).Columns(14).EditorStyle:id=261,.parent=127"
      _StyleDefs(172) =   "Splits(1).Columns(15).Style:id=270,.parent=123,.alignment=2"
      _StyleDefs(173) =   "Splits(1).Columns(15).HeadingStyle:id=267,.parent=124"
      _StyleDefs(174) =   "Splits(1).Columns(15).FooterStyle:id=268,.parent=125"
      _StyleDefs(175) =   "Splits(1).Columns(15).EditorStyle:id=269,.parent=127"
      _StyleDefs(176) =   "Named:id=33:Normal"
      _StyleDefs(177) =   ":id=33,.parent=0"
      _StyleDefs(178) =   "Named:id=34:Heading"
      _StyleDefs(179) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(180) =   ":id=34,.wraptext=-1"
      _StyleDefs(181) =   "Named:id=35:Footing"
      _StyleDefs(182) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(183) =   "Named:id=36:Selected"
      _StyleDefs(184) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(185) =   "Named:id=37:Caption"
      _StyleDefs(186) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(187) =   "Named:id=38:HighlightRow"
      _StyleDefs(188) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(189) =   "Named:id=39:EvenRow"
      _StyleDefs(190) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(191) =   "Named:id=40:OddRow"
      _StyleDefs(192) =   ":id=40,.parent=33"
      _StyleDefs(193) =   "Named:id=41:RecordSelector"
      _StyleDefs(194) =   ":id=41,.parent=34"
      _StyleDefs(195) =   "Named:id=42:FilterBar"
      _StyleDefs(196) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_trans_log_attendance.frx":0F6D
      TabIndex        =   8
      Top             =   240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1320
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frm_trans_log_attendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn_fp As Boolean
Dim rs_bound As New ADODB.Recordset
Public public_int_caller As Integer

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks



Private Function clear_all_log() As Boolean
Dim vRet As Boolean
Dim vErrorCode As Long

vRet = CZKEM1.ClearGLog(vMachineNumber)
If vRet Then
    clear_all_log = True
Else
    clear_all_log = False
    'CZKEM1.GetLastError vErrorCode
    'lblMessage.Caption = ErrorPrint(vErrorCode)
End If
End Function


Public Function download_data_log_recover() As Boolean
On Error GoTo capErr

Dim rs As New ADODB.Recordset

Dim dwEnrollNumber As Long
Dim dwVerifyMode As Long
Dim dwInOutMode As Long
Dim timeStr As String
Dim i, j As Long


rs.Open "select * from h_log_attendance_recover where att_number = -77", CnG, adOpenKeyset, adLockOptimistic
i = 0
j = get_inc_number("h_log_attendance_recover", "att_number", " ") - 1

If CZKEM1.ReadGeneralLogData(vMachineNumber) Then
    While CZKEM1.GetGeneralLogDataStr _
                (vMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)
        
        j = j + 1
        
        timeStr = Trim(timeStr)
        str_date = Left(timeStr, Len(timeStr) - 2) & "00"
        With rs
            .AddNew
            
            .Fields("att_number").Value = j 'must have tried by "no auto increment"
            .Fields("att_date").Value = str_date
            .Fields("ip_address").Value = FG_IP_ADDRESS
            
            .Fields("enrollnumber").Value = dwEnrollNumber
            .Fields("verifymode").Value = dwVerifyMode
            .Fields("flag_io").Value = dwInOutMode
            .Fields("entry_date").Value = Now
            
'            '--.Fields("att_number").Value = 1 'must have tried by "no auto increment"
'            .Fields("att_date").Value = str_date
'            .Fields("ip_address").Value = FG_IP_ADDRESS
'            .Fields("enrollnumber").Value = dwEnrollNumber
'            .Fields("verifymode").Value = dwVerifyMode
'            .Fields("flag_io").Value = dwInOutMode
'            .Fields("entry_date").Value = Now
            
            .Update
        End With
        
        i = i + 1
        lbl_downloading.Caption = "Downloading...(" & i & ")"
    Wend

Else
    download_data_log_recover = False
End If

MousePointer = vbDefault
'If public_int_caller = 0 Then _
    MsgBox i & " are successfully downloaded...", vbInformation, headerMSG
lbl_downloading.Caption = "Downloading..."

download_data_log_recover = True
If i = 0 Then download_data_log_recover = False
Exit Function

capErr:
MsgBox "Error downloading data..." & vbCr & _
    err.Description, vbCritical, headerMSG
MousePointer = vbDefault
download_data_log_recover = False
End Function

Public Function download_data_log() As Boolean
'On Error GoTo capErr

Dim dwEnrollNumber As Long
Dim dwVerifyMode As Long
Dim dwInOutMode As Long
Dim timeStr As String
Dim i As Long
Dim lng_year As Long
Dim lng_month As Long
Dim lng_day As Long
Dim lng_hour As Long
Dim lng_minute As Long
Dim lng_second As Long
Dim dw_work As Long
Dim strSQL As String

Dim rsabsen As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs7 As New ADODB.Recordset

Dim str1, v_company_code, v_employee_code, v_shift_code As String
Dim v_shift_number, v_flag_batch As Integer
Dim str_date As String

'If Not download_data_log_recover Then
'    download_data_log = False
'    Exit Function
'End If

dw_work = 0
i = 0
If CZKEM1.ReadGeneralLogData(vMachineNumber) Then
'    fra_downloading.Visible = True
    While CZKEM1.GetGeneralLogDataStr(vMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)
        
'    While CZKEM1.SSR_GetGeneralLogData(vMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, _
'    lng_year, lng_month, lng_day, lng_hour, lng_minute, lng_second, dw_work)
                
        timeStr = Trim(timeStr)
        str_date = Left(timeStr, Len(timeStr) - 2) & "00"
              
'        str_date = _
'        Trim(str(lng_year)) & "-" & Right("00" & Trim(str(lng_month)), 2) _
'                            & "-" & Right("00" & Trim(str(lng_day)), 2) _
'        & " " & Right("00" & Trim(str(lng_hour)), 2) & ":" & Right("00" & Trim(str(lng_minute)), 2) _
'                                                    & ":" & Right("00" & Trim(str(lng_second)), 2)
        
        If rs7.State = 1 Then rs7.Close
        rs7.Open "select count(*) as recs from h_log_attendance where att_date='" _
        & str_date & "' and ip_address='" & FG_IP_ADDRESS & "' and enrollnumber=" _
        & dwEnrollNumber, CnG, adOpenStatic, adLockReadOnly
        
        'Set-up the progress bar's properties
        ProgressBar1.Min = 0
        ProgressBar1.Max = rs7.RecordCount
        ProgressBar1.Value = 0
        
        '*************2011-11-11 RUBAH DOWNLOAD DARI TRIGER KE VB**********************
        If rs7.Fields("recs").Value <= 0 Then

            With rs_bound

                .AddNew

                '.Fields("att_number").Value = 1 'must have tried by "no auto increment"
                .Fields("att_date").Value = str_date
                .Fields("ip_address").Value = FG_IP_ADDRESS
                .Fields("enrollnumber").Value = Val(dwEnrollNumber)
                '
                .Fields("employee_code").Value = get_employee_code(Val("" & dwEnrollNumber))
                .Fields("verifymode").Value = dwVerifyMode
                .Fields("flag_io").Value = dwInOutMode
                .Fields("flag_attendance").Value = 0
                .Fields("entry_date").Value = Now

                .Update
            End With
            i = i + 1
            lbl_downloading.Caption = "Downloading...(" & i & ")"
        End If
        '*********************************************************
    Wend
'        fra_downloading.Visible = False
Else
    download_data_log = False
End If

    strSQL = "SELECT DATE(a.att_date) tgl_att,a.employee_code,MIN(att_date) check_in,MAX(att_date) check_out,a.enrollnumber,a.ip_address," & _
        "((HOUR( TIMEDIFF(MAX(att_date),MIN(att_date)))*60) + MINUTE(TIMEDIFF(MAX(att_date),MIN(att_date)) )) selisih " & _
    "FROM h_log_attendance a LEFT JOIN m_employee b ON a.employee_code = b.employee_code " & _
    "WHERE IFNULL(a.sdhupdate,0) = 0 GROUP BY 1,2"
    
    rsabsen.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
    If rsabsen.RecordCount > 0 Then
        Dim insertAtt As New clsInsertAtt_New
'        Frame1.Visible = True
'        ProgressBar1.Value = 0
'        ProgressBar1.Max = rsabsen.RecordCount
'        rsabsen.MoveFirst
        
        While Not rsabsen.EOF
            DoEvents
                    
'            ProgressBar1.Value = ProgressBar1.Value + 1
            
'            Label5.Caption = "Download Data, Please Wait......"
'            Label1.Caption = "Data Ke " & ProgressBar1.Value & " dari " & rsabsen.RecordCount
        
            Call insertAtt.Insert_h_attendance(Format(rsabsen!tgl_att, "yyyy-MM-dd"), rsabsen!employee_code, _
                Format(rsabsen!check_in, "yyyy-MM-dd hh:mm:ss"), Format(rsabsen!check_out, "yyyy-MM-dd hh:mm:ss"), _
                rsabsen!ip_address, rsabsen!enrollnumber)
            rsabsen.MoveNext
        Wend
    End If
'    Frame1.Visible = False
    
    '++++++++++++++++
    strSQL = "UPDATE h_log_attendance SET sdhupdate = 1 WHERE sdhupdate = 0"
    CnG.Execute strSQL
    '++++++++++++++++
    '++++++++++++++++++++++++++++++++++++++++++
    rsabsen.Close

MousePointer = vbDefault
If public_int_caller = 0 Then _
    MsgBox i & " are successfully downloaded...", vbInformation, headerMSG
lbl_downloading.Caption = "Downloading..."

download_data_log = True
If i = 0 Then MsgBox "Tidak ada data !", vbCritical, headerMSG
Exit Function

capErr:
'MsgBox "Error downloading data..." & vbCr & _
'    err.Description, vbCritical, headerMSG
MousePointer = vbDefault
download_data_log = False
End Function

Private Function get_employee_code(ByVal int_number As Long) As String
Dim rs1 As New ADODB.Recordset
Dim str1, str_employee_code As String

str1 = "SELECT company_code, employee_code  " _
    & "FROM m_enroll_link WHERE ip_address = '" & FG_IP_ADDRESS & _
    "' AND enrollnumber = " & int_number

If rs1.State = 1 Then rs1.Close
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then
    str_employee_code = rs1.Fields("employee_code").Value
Else
    str_employee_code = ""
End If

get_employee_code = str_employee_code
End Function

'Public Function download_data_log() As Boolean
'On Error GoTo capErr
'Dim rsabsen As New ADODB.Recordset
'Dim strSQL As String
'
'Dim dwEnrollNumber As Long
'Dim dwVerifyMode As Long
'Dim dwInOutMode As Long
'Dim timeStr As String
'Dim i As Long
'
''Update 28-11-2011++++++++++++++++++++
''If Not download_data_log_recover Then
''    download_data_log = False
''    Exit Function
''End If
''+++++++++++++++++++++++++++++++++++++
'
'i = 0
'If CZKEM1.ReadGeneralLogData(vMachineNumber) Then
'    While CZKEM1.GetGeneralLogDataStr(vMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)
'
'        timeStr = Trim(timeStr)
'        str_date = Left(timeStr, Len(timeStr) - 2) & "00"
'        With rs_bound
'            .AddNew
'
'            '.Fields("att_number").Value = 1 'must have tried by "no auto increment"
'            .Fields("att_date").Value = str_date
'            .Fields("ip_address").Value = FG_IP_ADDRESS
'
'            .Fields("enrollnumber").Value = dwEnrollNumber
'            .Fields("verifymode").Value = dwVerifyMode
'            .Fields("flag_io").Value = dwInOutMode
'            .Fields("entry_date").Value = Now
'
'            .Update
'        End With
'
'        'i = i + 1
'        'lbl_downloading.Caption = "Downloading...(" & i & ")"
'    Wend
'
'    'Update 22-10-2011+++++++++++++++++++++++++++++++++
'    strSQL = "SELECT * FROM ( " & _
'        "SELECT DATE(a.att_date) tgl_att,a.employee_code,MIN(att_date) check_in,MAX(att_date) check_out,a.enrollnumber,a.ip_address," & _
'            "((HOUR( TIMEDIFF(MAX(att_date),MIN(att_date)))*60) + MINUTE(TIMEDIFF(MAX(att_date),MIN(att_date)) )) selisih " & _
'        "FROM h_log_attendance a LEFT JOIN m_employee b ON a.employee_code = b.employee_code " & _
'        "WHERE a.sdhupdate = 0 GROUP BY 1,2)xx " & _
'        "WHERE selisih > 30"
'
'    rsabsen.Open strSQL, CnG, adOpenForwardOnly, adLockReadOnly
'    If rsabsen.RecordCount > 0 Then
'        Dim insertAtt As New clsInsertAtt_New
'        Frame1.Visible = True
'        ProgressBar1.Value = 0
'
'        rsabsen.MoveFirst
'        While Not rsabsen.EOF
'            ProgressBar1.Max = rsabsen.RecordCount
'            ProgressBar1.Value = ProgressBar1.Value + 1
'
'            Label5.Caption = "Download Data, Please Wait......"
'            Label1.Caption = "Data Ke " & ProgressBar1.Value & " dari " & rsabsen.RecordCount
'
'            Call insertAtt.Insert_h_attendance(Format(rsabsen!tgl_att, "yyyy-MM-dd"), rsabsen!employee_code, _
'                Format(rsabsen!check_in, "yyyy-MM-dd hh:mm:ss"), Format(rsabsen!check_out, "yyyy-MM-dd hh:mm:ss"), _
'                rsabsen!ip_address, rsabsen!enrollnumber)
'            rsabsen.MoveNext
'        Wend
'    End If
'
'    '++++++++++++++++
'    strSQL = "UPDATE h_log_attendance SET sdhupdate = 1 WHERE sdhupdate = 0"
'    CnG.Execute strSQL
'    '++++++++++++++++
'    '++++++++++++++++++++++++++++++++++++++++++
'Else
'    download_data_log = False
'End If
'
'    Frame1.Visible = False
'
'MousePointer = vbDefault
'If public_int_caller = 0 Then _
'    MsgBox i & " are successfully downloaded...", vbInformation, headerMSG
'lbl_downloading.Caption = "Downloading..."
'
'download_data_log = True
'If i = 0 Then download_data_log = False
'Exit Function
'
'capErr:
'MsgBox "Error downloading data..." & vbCr & _
'    err.Description, vbCritical, headerMSG
'MousePointer = vbDefault
'download_data_log = False
'End Function

Public Function download_data_log_bug1() As Boolean
On Error GoTo capErr

Dim vTMachineNumber, vSMachineNumber, vSEnrollNumber As Long
Dim vVerifyMode, vInOutMode As Long
Dim vYear, vMonth, vDay As Long
Dim vHour, vMinute As Long
Dim vErrorCode As Long
Dim vRet As Boolean
Dim i, n As Long
Dim str_date As String
Dim rs As New ADODB.Recordset

i = 0
DoEvents
MousePointer = vbHourglass

Do
    vRet = CZKEM1.GetAllGLogData(vMachineNumber, vTMachineNumber, _
                                vSEnrollNumber, vSMachineNumber, vVerifyMode, _
                                vInOutMode, vYear, vMonth, vDay, vHour, vMinute)
    
    If (vRet = False) Then Exit Do
    str_date = CStr(vYear) & "-" & Format(vMonth, "0#") & "-" & Format(vDay, "0#") _
                & " " & Format(vHour, "0#") & ":" & Format(vMinute, "0#") & ":00"
    
    With rs_bound
        .AddNew
        
        '.Fields("att_number").Value = 1
        .Fields("att_date").Value = str_date
        .Fields("enrollnumber").Value = vSEnrollNumber
        .Fields("verifymode").Value = vVerifyMode
        .Fields("flag_io").Value = vInOutMode
        .Fields("entry_date").Value = Format(Now, "yyyy-MM-dd HH:mm:ss")
        
        .Update
    End With
    
    i = i + 1
    lbl_downloading.Caption = "Downloading...(" & i & ")"
Loop

MousePointer = vbDefault
If public_int_caller = 0 Then _
    MsgBox i & " are successfully downloaded...", vbInformation, headerMSG
lbl_downloading.Caption = "Downloading..."

download_data_log_bug1 = True
If i = 0 Then download_data_log_bug1 = False
Exit Function

capErr:
MsgBox "Error downloading data...", vbCritical, headerMSG
MousePointer = vbDefault
download_data_log_bug1 = False
End Function

Private Function connect() As Boolean
If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.disconnect
End If

cn_fp = CZKEM1.Connect_Net(FG_IP_ADDRESS, CLng(FG_PORT_NUMBER))

If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, False
    connect = True
Else
    connect = False
    Exit Function
End If
End Function

Private Function disconnect() As Boolean
If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.disconnect
End If
End Function

Private Sub cmd_delete_log_Click()
frm_lookup_mst_device.public_int_mode = 2
frm_lookup_mst_device.Show 1
End Sub

Public Sub cmd_download_Click()
frm_lookup_mst_device.public_int_mode = 1
frm_lookup_mst_device.Show 1

End Sub

Public Sub download_action()
If rs_bound.State = 1 Then rs_bound.Close
rs_bound.Open "select * from h_log_attendance where att_number = -1", CnG, adOpenKeyset, adLockOptimistic

fra_downloading.Visible = True: TDBGrid1.Enabled = False: fra_button_control.Enabled = False

If Not connect Then
    If Not public_int_caller = 1 Then MsgBox "Error connecting to source device...", vbCritical, headerMSG
Else
    If download_data_log Then
        MsgBox "Downloading is Done !", vbInformation, headerMSG
    Else
        MsgBox "Error downloading !", vbCritical, headerMSG
    End If
    Call disconnect
End If

Call disconnect
fra_downloading.Visible = False: TDBGrid1.Enabled = True: fra_button_control.Enabled = True

Call load_data_log_att
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn
Adodc_company.ConnectionString = strConn

Frame1.Visible = False

Call load_data_company

public_int_caller = 0
cn_fp = False
vMachineNumber = 1
vEMachineNumber = 1

Call load_data_user_access(Me)
cmd_download.Enabled = blnUser_Add And blnUser_Edit
Timer1.Enabled = True
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub load_data_log_att()
Adodc1.RecordSource = "select * from v_h_attendance where company_code='" _
& TDBCombo_company.Columns("company_code").Value _
& "' order by company_code, department_code, division_code, employee_code, att_date asc limit 0,50"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    Call load_data_log_att
End If
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo ErrHandler

Dim i As Integer

Set Cols = TDBGrid1.Columns
i = TDBGrid1.Col
TDBGrid1.HoldFields

Adodc1.Recordset.Filter = getFilter()
TDBGrid1.Col = i
TDBGrid1.EditActive = True

TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
If TDBGrid1.ApproxCount < 1 Then
    Call clear_filter
    TDBGrid1.Col = i
End If

Exit Sub
ErrHandler:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub clear_filter()
For Each Col In TDBGrid1.Columns
    Col.FilterText = ""
Next Col
Adodc1.Recordset.Filter = adFilterNone
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer

For Each Col In Cols
    If Trim(Col.FilterText) <> "" Then
        n = n + 1
        If n > 1 Then
            tmp = tmp & " AND "
        End If
        
        tmp = tmp & Col.DataField & " LIKE '" & Col.FilterText & "*'"
    End If
Next Col
getFilter = tmp
End Function

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "DATE" Then
    Value = Format(Value, "yyyy-mm-dd")
ElseIf TDBGrid1.Columns(ColIndex).Caption = "CHECK IN" _
Or TDBGrid1.Columns(ColIndex).Caption = "CHECK OUT" Then
    Value = Format(Value, "HH:mm")
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Public Sub delete_log_action()
If rs_bound.State = 1 Then rs_bound.Close
rs_bound.Open "select * from h_log_attendance where att_number = -1", CnG, adOpenKeyset, adLockOptimistic

lbl_downloading.Caption = "Delete log data..."
fra_downloading.Visible = True: TDBGrid1.Enabled = False: fra_button_control.Enabled = False

'Label5.Caption = "Delete log data..."
'Frame1.Visible = True: TDBGrid1.Enabled = False: fra_button_control.Enabled = False

If Not connect Then
    If Not public_int_caller = 1 Then MsgBox "Error connecting to source device...", vbCritical, headerMSG
Else
    If clear_all_log = False Then
        If Not public_int_caller = 1 Then MsgBox "Error deleting log...", vbCritical, headerMSG
    End If
    Call disconnect
End If

fra_downloading.Visible = False: TDBGrid1.Enabled = True: fra_button_control.Enabled = True

MsgBox "Deleting is Done !", vbInformation, headerMSG

End Sub