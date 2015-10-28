VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{66A5AC41-25A9-11D2-9BBF-00A024695830}#1.0#0"; "titime6.ocx"
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_log_attendance_manual 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MANUAL LOG ATTENDANCE"
   ClientHeight    =   9090
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   14655
   Icon            =   "frm_trans_log_attendance_manual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14655
   ShowInTaskbar   =   0   'False
   Begin TDBTime6Ctl.TDBTime TDBTime1 
      Height          =   375
      Left            =   7920
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
      _Version        =   65536
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "frm_trans_log_attendance_manual.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Keys            =   "frm_trans_log_attendance_manual.frx":0078
      Spin            =   "frm_trans_log_attendance_manual.frx":00C8
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "hh:nn"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "hh:nn"
      HighlightText   =   0
      Hour12Mode      =   1
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxTime         =   0.99999
      MidnightMode    =   0
      MinTime         =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "06:37"
      ValidateMode    =   0
      ValueVT         =   1768816647
      Value           =   0.275833333333333
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4560
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   7800
      TabIndex        =   18
      Top             =   240
      Width           =   5295
      Begin VB.OptionButton opt_break_out 
         Caption         =   "BREAK OUT"
         Height          =   195
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton opt_break_in 
         Caption         =   "BREAK IN"
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton opt_out 
         Caption         =   "OUT"
         Height          =   195
         Left            =   4320
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton opt_in 
         Caption         =   "IN"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   661
      Calendar        =   "frm_trans_log_attendance_manual.frx":00F0
      Caption         =   "frm_trans_log_attendance_manual.frx":0208
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_trans_log_attendance_manual.frx":0274
      Keys            =   "frm_trans_log_attendance_manual.frx":0292
      Spin            =   "frm_trans_log_attendance_manual.frx":02F0
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd-mm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd-mm-yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "15-12-2009"
      ValidateMode    =   0
      ValueVT         =   1768816647
      Value           =   40162
      CenturyMode     =   0
   End
   Begin VB.CheckBox chk_company 
      Caption         =   "PER COMPANY"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   15
      Top             =   720
      Width           =   3855
   End
   Begin VB.Frame fra_downloading 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   2520
      TabIndex        =   12
      Top             =   3120
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
         TabIndex        =   13
         Top             =   480
         Width           =   1890
      End
   End
   Begin VB.Frame fra_button_control 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   7560
      Width           =   14175
      Begin VB.CommandButton cmd_post 
         Caption         =   "&Post"
         Height          =   645
         Left            =   4920
         Picture         =   "frm_trans_log_attendance_manual.frx":0318
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   3840
         Picture         =   "frm_trans_log_attendance_manual.frx":08A2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   12240
         Picture         =   "frm_trans_log_attendance_manual.frx":0E2C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_load 
         Caption         =   "&Load"
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
         Left            =   2760
         Picture         =   "frm_trans_log_attendance_manual.frx":13B6
         Style           =   1  'Graphical
         TabIndex        =   6
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
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   375
      Left            =   0
      OleObjectBlob   =   "frm_trans_log_attendance_manual.frx":1940
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6135
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   10821
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "COMPANY NAME"
      Columns(0).DataField=   "company_name"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EMPLOYEE CODE"
      Columns(1).DataField=   "employee_code"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "EMPLOYEE NAME"
      Columns(2).DataField=   "employee_name"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "CHECK TIME"
      Columns(3).DataField=   "att_date"
      Columns(3).NumberFormat=   "External Editor"
      Columns(3).ExternalEditor=   "TDBTime1"
      Columns(3).ExternalEditor.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   4
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "POSTED"
      Columns(4).DataField=   "flag_posted"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=7752"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=7673"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8708"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3810"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3731"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(1)._MinWidth=80000960"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=7329"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=7250"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=8708"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(2)._MinWidth=79999936"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=3096"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=3016"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=1879"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1799"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=8705"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(4)._MinWidth=138900112"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "LIST OF LOG ATTENDANCE"
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
      _StyleDefs(21)  =   "Splits(0).Style:id=123,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=132,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(23)  =   ":id=132,.fgcolor=&H80000009&"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=124,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(25)  =   ":id=124,.fgcolor=&H80000002&"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=125,.parent=3"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=126,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=128,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=127,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=129,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=130,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=131,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=133,.parent=11"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=134,.parent=12"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=142,.parent=123,.locked=-1"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=139,.parent=124"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=140,.parent=125"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=141,.parent=127"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=162,.parent=123,.locked=-1"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=159,.parent=124"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=160,.parent=125"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=161,.parent=127"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=166,.parent=123,.locked=-1"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=163,.parent=124"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=164,.parent=125"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=165,.parent=127"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=246,.parent=123,.alignment=2"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=243,.parent=124"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=244,.parent=125"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=245,.parent=127"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=16,.parent=123,.alignment=2,.locked=-1"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=13,.parent=124"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=14,.parent=125"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=15,.parent=127"
      _StyleDefs(55)  =   "Named:id=33:Normal"
      _StyleDefs(56)  =   ":id=33,.parent=0"
      _StyleDefs(57)  =   "Named:id=34:Heading"
      _StyleDefs(58)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   ":id=34,.wraptext=-1"
      _StyleDefs(60)  =   "Named:id=35:Footing"
      _StyleDefs(61)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   "Named:id=36:Selected"
      _StyleDefs(63)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(64)  =   "Named:id=37:Caption"
      _StyleDefs(65)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(66)  =   "Named:id=38:HighlightRow"
      _StyleDefs(67)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(68)  =   "Named:id=39:EvenRow"
      _StyleDefs(69)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(70)  =   "Named:id=40:OddRow"
      _StyleDefs(71)  =   ":id=40,.parent=33"
      _StyleDefs(72)  =   "Named:id=41:RecordSelector"
      _StyleDefs(73)  =   ":id=41,.parent=34"
      _StyleDefs(74)  =   "Named:id=42:FilterBar"
      _StyleDefs(75)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1800
      OleObjectBlob   =   "frm_trans_log_attendance_manual.frx":1964
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1800
      Top             =   840
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
   Begin VB.Label Label1 
      Caption         =   "DATE"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frm_trans_log_attendance_manual"
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
            .Fields("enrollnumber").Value = dwEnrollNumber
            .Fields("verifymode").Value = dwVerifyMode
            .Fields("flag_io").Value = dwInOutMode
            .Fields("entry_date").Value = Now
            
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
MsgBox "Error downloading data...", vbCritical, headerMSG
MousePointer = vbDefault
download_data_log_recover = False
End Function

Public Function download_data_log() As Boolean
On Error GoTo capErr

Dim dwEnrollNumber As Long
Dim dwVerifyMode As Long
Dim dwInOutMode As Long
Dim timeStr As String
Dim i As Long

If Not download_data_log_recover Then
    download_data_log = False
    Exit Function
End If

i = 0
If CZKEM1.ReadGeneralLogData(vMachineNumber) Then
    While CZKEM1.GetGeneralLogDataStr _
                (vMachineNumber, dwEnrollNumber, dwVerifyMode, dwInOutMode, timeStr)
        
        timeStr = Trim(timeStr)
        str_date = Left(timeStr, Len(timeStr) - 2) & "00"
        With rs_bound
            .AddNew
            
            '.Fields("att_number").Value = 1 'must have tried by "no auto increment"
            .Fields("att_date").Value = str_date
            .Fields("ip_address").Value = FG_IP_ADDRESS
            .Fields("enrollnumber").Value = dwEnrollNumber
            .Fields("verifymode").Value = dwVerifyMode
            .Fields("flag_io").Value = dwInOutMode
            .Fields("entry_date").Value = Now
            
            .Update
        End With
        
        i = i + 1
        lbl_downloading.Caption = "Downloading...(" & i & ")"
    Wend

Else
    download_data_log = False
End If

MousePointer = vbDefault
If public_int_caller = 0 Then _
    MsgBox i & " are successfully downloaded...", vbInformation, headerMSG
lbl_downloading.Caption = "Downloading..."

download_data_log = True
If i = 0 Then download_data_log = False
Exit Function

capErr:
MsgBox "Error downloading data...", vbCritical, headerMSG
MousePointer = vbDefault
download_data_log = False
End Function

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

Public Sub cmd_download_Click()
frm_lookup_mst_device.public_int_mode = 1
frm_lookup_mst_device.Show 1

End Sub

Public Sub download_action()
If rs_bound.State = 1 Then rs_bound.Close
rs_bound.Open "select * from h_log_attendance where att_number = -1", CnG, adOpenKeyset, adLockOptimistic

fra_downloading.Visible = True: TDBGrid1.Enabled = False: fra_button_control.Enabled = False

If Not connect Then
    If Not public_int_caller = 1 Then MsgBox "Gagal Menghubungkan Fingerprint...", vbCritical, headerMSG
    Exit Sub
End If

If download_data_log Then
    If clear_all_log = False Then
        If Not public_int_caller = 1 Then MsgBox "Error deleting log...", vbCritical, headerMSG
    End If
End If

Call disconnect
fra_downloading.Visible = False: TDBGrid1.Enabled = True: fra_button_control.Enabled = True

Call load_data_log_att
End Sub

Private Sub chk_company_Click()
If Not chk_company.Value = vbChecked Then
    TDBCombo_company.Text = ""
    txt_company_name = ""
    TDBCombo_company.Enabled = False
Else
    TDBCombo_company.Enabled = True
    Call TDBCombo_company_ItemChange
End If

Call load_data_manual_log
End Sub

Private Sub cmd_load_Click()
Dim strx As String
Dim rs1 As New ADODB.Recordset
Dim rsx As New ADODB.Recordset
Dim str1, str2, str3, stry As String
Dim i, int1 As Integer

If chk_company.Value = vbChecked Then
    str1 = " and company_code = '" & TDBCombo_company.Columns("company_code").Value & "' "
Else
    str1 = " "
End If

If opt_in Then
    str2 = " 0 "
    str3 = " 08:00:00"
    int1 = 0
ElseIf opt_out Then
    str2 = " 1 "
    str3 = " 17:00:00"
    int1 = 1
ElseIf opt_break_in Then
    str2 = " 2 "
    str3 = " 12:00:00"
    int1 = 2
ElseIf opt_break_out Then
    str2 = " 3 "
    str3 = " 13:00:00"
    int1 = 3
End If

stry = "select employee_code from t_log_attendance where left(att_date,10) = '" _
    & Format(TDBDate1, "yyyy-mm-dd") & "' and flag_io = " & str2 & str1

strx = "select * from m_employee where employee_code not in (" _
        & stry & ")"

'Text1 = strx
'Exit Sub

rs1.Open strx, CnG, adOpenStatic, adLockReadOnly
rsx.Open "select * from t_log_attendance where employee_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic

If Not rs1.RecordCount > 0 Then
    Exit Sub
End If

CnG.BeginTrans
i = 0
rs1.MoveFirst
While Not rs1.EOF
    rsx.AddNew
    
    rsx.Fields("att_date").Value = Format(TDBDate1, "yyyy-mm-dd") & str3
    rsx.Fields("ip_address").Value = "" 'rs1.Fields("ip_address").Value
    'rsx.Fields("enrollnumber").Value = "" 'rs1.Fields("enrollnumber").Value
    rsx.Fields("company_code").Value = rs1.Fields("company_code").Value
    rsx.Fields("company_name").Value = rs1.Fields("company_name").Value
    rsx.Fields("employee_code").Value = rs1.Fields("employee_code").Value
    rsx.Fields("employee_name").Value = rs1.Fields("employee_name").Value
    'rsx.Fields("verifymode").Value = "" 'rs1.Fields("verifymode").Value
    rsx.Fields("flag_io").Value = int1 'rs1.Fields("flag_io").Value
    rsx.Fields("flag_posted").Value = 0 'rs1.Fields("flag_posted").Value
    rsx.Fields("entry_date").Value = Now
    
    rsx.Update
    i = i + 1
    rs1.MoveNext
Wend

CnG.CommitTrans

Call load_data_manual_log
End Sub

Private Sub cmd_post_Click()
Dim i, j As Integer
Dim rs1 As New ADODB.Recordset
Dim str1 As String

If Not Adodc1.Recordset.RecordCount > 0 Then
    Exit Sub
End If

If opt_in Then
    str1 = "IN"
ElseIf opt_out Then
    str1 = "OUT"
ElseIf opt_break_in Then
    str1 = "BREAk IN"
ElseIf opt_break_out Then
    str1 = "BREAK OUT"
End If

j = MsgBox("Are you sure want to post attendance data (" & str1 & ") on '" _
    & Format(TDBDate1, "dd-mm-yyyy") & "'?", vbYesNo + vbQuestion, headerMSG)
If Not j = vbYes Then Exit Sub

i = 0
rs1.Open "select * from h_log_attendance where employee_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic
CnG.BeginTrans

If Adodc1.Recordset.RecordCount > 0 Then Adodc1.Recordset.MoveFirst
While Not Adodc1.Recordset.EOF
    If Not Adodc1.Recordset.Fields("flag_posted").Value = 1 Then
        rs1.AddNew
        'rs1.Fields("att_number").Value = Adodc1.Recordset.Fields("att_number").Value
        rs1.Fields("att_date").Value = Adodc1.Recordset.Fields("att_date").Value
        'rs1.Fields("ip_address").Value = Adodc1.Recordset.Fields("ip_address").Value
        'rs1.Fields("enrollnumber").Value = Adodc1.Recordset.Fields("enrollnumber").Value
        rs1.Fields("employee_code").Value = Adodc1.Recordset.Fields("employee_code").Value
        'rs1.Fields("verifymode").Value = Adodc1.Recordset.Fields("verifymode").Value
        rs1.Fields("flag_io").Value = Adodc1.Recordset.Fields("flag_io").Value
        'rs1.Fields("flag_attendance").Value = Adodc1.Recordset.Fields("flag_attendance").Value
        rs1.Fields("entry_date").Value = Adodc1.Recordset.Fields("entry_date").Value
        rs1.Update
        
        i = i + 1
        Adodc1.Recordset.Fields("flag_posted").Value = 1
    End If
    Adodc1.Recordset.MoveNext
Wend
TDBGrid1.Refresh

CnG.CommitTrans

MsgBox i & " data was successfully posted!", vbInformation, headerMSG
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
TDBGrid1.Update
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn
Adodc_company.ConnectionString = strConn

Call load_data_company

'public_int_caller = 0
'cn_fp = False
'vMachineNumber = 1
'vEMachineNumber = 1

Call load_data_user_access(Me)
TDBDate1 = Now
Call chk_company_Click
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company

If Not Adodc_company.Recordset.RecordCount > 0 Then
    MsgBox "No data company to select!", vbInformation, headerMSG
    'Unload Me
End If
End Sub

Private Sub load_data_log_att()
Adodc1.RecordSource = "select * from v_h_attendance where company_code='" _
& TDBCombo_company.Columns("company_code").Value _
& "' order by company_code, department_code, division_code, employee_code, att_date asc limit 0,50"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub load_data_manual_log()
Dim str1, str2 As String

If chk_company Then
    str1 = " and company_code = '" & TDBCombo_company.Columns("company_code").Value & "' "
Else
    str1 = " "
End If
If opt_in Then
    str2 = " 0 "
ElseIf opt_out Then
    str2 = " 1 "
ElseIf opt_break_in Then
    str2 = " 2 "
ElseIf opt_break_out Then
    str2 = " 3 "
End If

Adodc1.RecordSource = "select * from t_log_attendance where left(att_date,10) = '" _
    & Format(TDBDate1, "yyyy-mm-dd") & "' and flag_io = " & str2 & str1 _
    & " order by company_code asc, employee_code asc"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub opt_in_Click()
Call load_data_manual_log
End Sub
Private Sub opt_out_Click()
Call load_data_manual_log
End Sub
Private Sub opt_break_in_Click()
Call load_data_manual_log
End Sub
Private Sub opt_break_out_Click()
Call load_data_manual_log
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
'  And TDBCombo_company.Bookmark > 0
'  And Not Trim(TDBCombo_company.Text) = "" Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    Call load_data_manual_log
End If
End Sub

Private Sub TDBDate1_Change()
Call load_data_manual_log
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
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
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
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub
