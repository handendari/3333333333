VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_working_day 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER WORKING DAY"
   ClientHeight    =   5655
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   12615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mst_wd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   12615
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7560
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox txt_working_time_name 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3360
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   19
      Top             =   240
      Width           =   3855
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   12135
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3240
         Picture         =   "frm_mst_wd.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4320
         Picture         =   "frm_mst_wd.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   0
         Picture         =   "frm_mst_wd.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   1080
         Picture         =   "frm_mst_wd.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   10200
         Picture         =   "frm_mst_wd.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5400
         Picture         =   "frm_mst_wd.frx":1BBE
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2160
         Picture         =   "frm_mst_wd.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_refresh 
         Caption         =   "&Load"
         Height          =   645
         Left            =   7920
         Picture         =   "frm_mst_wd.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
   End
   Begin VB.Frame fra_entry 
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   12135
      Begin VB.TextBox txt_break_interval_minute 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9960
         MaxLength       =   10
         TabIndex        =   23
         Top             =   1440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox cbo_working_day 
         Height          =   315
         ItemData        =   "frm_mst_wd.frx":2C5C
         Left            =   2040
         List            =   "frm_mst_wd.frx":2C75
         TabIndex        =   22
         Text            =   "..."
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cbo_active 
         Height          =   315
         ItemData        =   "frm_mst_wd.frx":2CB9
         Left            =   2040
         List            =   "frm_mst_wd.frx":2CC3
         TabIndex        =   17
         Text            =   "..."
         Top             =   1080
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker DTPicker_start_time 
         Height          =   315
         Left            =   6120
         TabIndex        =   6
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "HH:mm"
         Format          =   143130627
         UpDown          =   -1  'True
         CurrentDate     =   39270.5
      End
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin MSComCtl2.DTPicker DTPicker_end_time 
         Height          =   315
         Left            =   6120
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "HH:mm"
         Format          =   143130627
         UpDown          =   -1  'True
         CurrentDate     =   39270
      End
      Begin MSComCtl2.DTPicker DTPicker_min_break_in 
         Height          =   315
         Left            =   9960
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "HH:mm"
         Format          =   143130627
         UpDown          =   -1  'True
         CurrentDate     =   39270.5
      End
      Begin MSComCtl2.DTPicker DTPicker_max_break_out 
         Height          =   315
         Left            =   9960
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "HH:mm"
         Format          =   143130627
         UpDown          =   -1  'True
         CurrentDate     =   39270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "MIN BREAK-IN"
         Height          =   195
         Left            =   8520
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "MAX BREAK-OUT"
         Height          =   195
         Left            =   8520
         TabIndex        =   27
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "INTERVAL"
         Height          =   195
         Left            =   8520
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ACTIVE"
         Height          =   195
         Left            =   720
         TabIndex        =   18
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "IN"
         Height          =   195
         Left            =   5520
         TabIndex        =   5
         Top             =   720
         Width           =   165
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "WORKING DAY"
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "OUT"
         Height          =   195
         Left            =   5520
         TabIndex        =   3
         Top             =   1080
         Width           =   315
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   5741
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "SHIFT CODE"
      Columns(0).DataField=   "shift_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "W. DAY CODE"
      Columns(1).DataField=   "day_code"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "W. DAY NAME"
      Columns(2).DataField=   "day_name"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "IN"
      Columns(3).DataField=   "start_time"
      Columns(3).NumberFormat=   "FormatText Event"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "OUT"
      Columns(4).DataField=   "end_time"
      Columns(4).NumberFormat=   "FormatText Event"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DAY OVER"
      Columns(5).DataField=   "flag_day_over"
      Columns(5).NumberFormat=   "FormatText Event"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "MIN BREAK IN"
      Columns(6).DataField=   "min_break_in"
      Columns(6).NumberFormat=   "FormatText Event"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "MAX BREAK OUT"
      Columns(7).DataField=   "max_break_out"
      Columns(7).NumberFormat=   "FormatText Event"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "INTERVAL (M)"
      Columns(8).DataField=   "break_interval_minute"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   4
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "ACTIVE"
      Columns(9).DataField=   "flag_active"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=3731"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3651"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=6297"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6218"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2699"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2619"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=2381"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2302"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(5)._MinWidth=10"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=2223"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2143"
      Splits(0)._ColumnProps(38)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(40)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(41)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(43)=   "Column(6)._MinWidth=54215968"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=2381"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2302"
      Splits(0)._ColumnProps(47)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(49)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(50)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(51)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(52)=   "Column(7)._MinWidth=54215968"
      Splits(0)._ColumnProps(53)=   "Column(8).Width=2117"
      Splits(0)._ColumnProps(54)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(8)._WidthInPix=2037"
      Splits(0)._ColumnProps(56)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(57)=   "Column(8)._ColStyle=513"
      Splits(0)._ColumnProps(58)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(59)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(60)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(61)=   "Column(8)._MinWidth=54215968"
      Splits(0)._ColumnProps(62)=   "Column(9).Width=2355"
      Splits(0)._ColumnProps(63)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(9)._WidthInPix=2275"
      Splits(0)._ColumnProps(65)=   "Column(9)._ColStyle=513"
      Splits(0)._ColumnProps(66)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(67)=   "Column(9)._MinWidth=54215968"
      Splits.Count    =   1
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
      Caption         =   "LIST OF WORKING DAY (GENERAL)"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=98,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=102,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
      _StyleDefs(74)  =   "Named:id=33:Normal"
      _StyleDefs(75)  =   ":id=33,.parent=0"
      _StyleDefs(76)  =   "Named:id=34:Heading"
      _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   ":id=34,.wraptext=-1"
      _StyleDefs(79)  =   "Named:id=35:Footing"
      _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=36:Selected"
      _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=37:Caption"
      _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(85)  =   "Named:id=38:HighlightRow"
      _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=39:EvenRow"
      _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(89)  =   "Named:id=40:OddRow"
      _StyleDefs(90)  =   ":id=40,.parent=33"
      _StyleDefs(91)  =   "Named:id=41:RecordSelector"
      _StyleDefs(92)  =   ":id=41,.parent=34"
      _StyleDefs(93)  =   "Named:id=42:FilterBar"
      _StyleDefs(94)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_working_time 
      Height          =   375
      Left            =   1560
      OleObjectBlob   =   "frm_mst_wd.frx":2CD0
      TabIndex        =   20
      Top             =   240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_working_time 
      Height          =   375
      Left            =   1680
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
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "WORKING TIME"
      Height          =   195
      Left            =   240
      TabIndex        =   21
      Top             =   240
      Width           =   1140
   End
End
Attribute VB_Name = "frm_mst_working_day"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FlagNew As Boolean
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim str_kode_rekening As String
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns



Private Sub check_invalid()
MsgBox "Data Sudah Ada...", vbCritical, headerMSG
cbo_working_day.Text = ""
If cbo_working_day.Enabled = True Then cbo_working_day.SetFocus
End Sub

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(day_code) as rec_count from m_working_day where shift_code = '" _
& TDBCombo_working_time.Columns("shift_code").Value & "' and day_code=" _
& cbo_working_day.ListIndex
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function


Private Function check_validate_new() As Boolean
check_validate_new = True

'validasi day code
If cbo_working_day.ListIndex < 0 Then
    MsgBox "Day code Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
    cbo_working_day.SetFocus
    check_validate_new = False
    Exit Function
End If

End Function

Private Sub load_data()
timer1.Enabled = True
End Sub

Private Sub cmd_refresh_Click()
If TDBCombo_working_time.ApproxCount > 0 And TDBCombo_working_time.Bookmark > 0 Then
    Call generate_all_day
    Call load_data_working_day
End If
End Sub

Private Function check_day_code(ByVal i As Integer) As Boolean
Dim rs1 As New ADODB.Recordset

rs1.Open "select count(*) as recs from m_working_day where shift_code = '" _
& TDBCombo_working_time.Columns("shift_code").Value & "' and day_code = " & i, CnG, adOpenStatic, adLockReadOnly

If rs1!recs > 0 Then
    check_day_code = True
Else
    check_day_code = False
End If
End Function

Private Sub generate_all_day()
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim str_sql As String, i As Integer

rs.Open "select *, ifnull(break_interval_minute,0) as break_interval_minute_, " _
& "ifnull(min_break_in,sysdate()) as min_break_in_, ifnull(max_break_out,sysdate()) as max_break_out_ " _
& "from m_shift where shift_code = '" _
& TDBCombo_working_time.Columns("shift_code").Value & "'", CnG, adOpenStatic, adLockReadOnly

rs1.Open "select * from m_working_day where day_code = 77", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans
For i = 0 To 5
    cbo_working_day.ListIndex = i
    With rs1
    
        If Not check_day_code(i) Then
            .AddNew
            .Fields("shift_code").Value = TDBCombo_working_time.Columns("shift_code").Value
            .Fields("day_code").Value = cbo_working_day.ListIndex
            .Fields("day_name").Value = cbo_working_day.Text
            .Fields("start_time").Value = Format(Now, "yyyy-mm-dd ") & Format(rs!start_time, "hh:nn:ss")
            .Fields("end_time").Value = Format(Now, "yyyy-mm-dd ") & Format(rs!end_time, "hh:nn:ss")
            .Fields("flag_day_over").Value = rs!flag_day_over
            .Fields("flag_active").Value = 1
            
            .Fields("flag_moving").Value = rs!flag_moving
            .Fields("moving_number").Value = rs!moving_number
            
            .Fields("min_break_in").Value = rs!min_break_in_
            .Fields("max_break_out").Value = rs!max_break_out_
            .Fields("break_interval_minute").Value = rs!break_interval_minute_
            .Update
        
        End If
    End With
Next i
CnG.CommitTrans

End Sub

Private Sub cmdCancel_Click()
int_mode = 0
Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
    & TDBGrid1.Columns("day_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_working_day where shift_code = '" _
    & TDBCombo_working_time.Columns("shift_code").Value & "' and day_code=" _
    & TDBGrid1.Columns("day_code").Value
CnG.CommitTrans

Call load_data_working_day
int_mode = 0
Call load_mode
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_working_day where shift_code = '" _
& TDBCombo_working_time.Columns("shift_code").Value & "' and day_code=" _
& TDBGrid1.Columns("day_code").Value, CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_working_day where day_code = 'άφ'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub cmdPrint_Click()
TDBGrid1.PrintInfo.PageSetup
If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
End If
End Sub

Private Sub insert_new_data()
CnG.BeginTrans
With rsBound
    .AddNew
    
    .Fields("shift_code").Value = TDBCombo_working_time.Columns("shift_code").Value
    .Fields("day_code").Value = cbo_working_day.ListIndex
    '-----------------------------------------------------------------------------
    .Fields("day_name").Value = cbo_working_day.Text
    .Fields("start_time").Value = Format(DTPicker_start_time.Value, "yyyy-MM-dd HH:mm:00")
    .Fields("end_time").Value = Format(DTPicker_end_time.Value, "yyyy-MM-dd HH:mm:00")
    .Fields("flag_active").Value = cbo_active.ListIndex
    
    '-- additional 4 BPP
    .Fields("min_break_in").Value = Format(DTPicker_min_break_in, "yyyy-MM-dd HH:mm:00")
    .Fields("max_break_out").Value = Format(DTPicker_max_break_out, "yyyy-MM-dd HH:mm:00")
    .Fields("break_interval_minute").Value = Val(DropAllComma(txt_break_interval_minute))
    
    .Update
End With
CnG.CommitTrans
End Sub

Private Sub edit_old_data()
On Error GoTo err_capture

CnG.BeginTrans
With rsBound
    
    .Fields("shift_code").Value = TDBCombo_working_time.Columns("shift_code").Value
    .Fields("day_code").Value = cbo_working_day.ListIndex
    '-----------------------------------------------------------------------------
    .Fields("day_name").Value = cbo_working_day.Text
    .Fields("start_time").Value = Format(DTPicker_start_time.Value, "yyyy-MM-dd HH:mm:00")
    .Fields("end_time").Value = Format(DTPicker_end_time.Value, "yyyy-MM-dd HH:mm:00")
    .Fields("flag_active").Value = cbo_active.ListIndex
    
    '-- additional 4 BPP
    .Fields("min_break_in").Value = Format(DTPicker_min_break_in, "yyyy-MM-dd HH:mm:00")
    .Fields("max_break_out").Value = Format(DTPicker_max_break_out, "yyyy-MM-dd HH:mm:00")
    .Fields("break_interval_minute").Value = Val(DropAllComma(txt_break_interval_minute))
    
    .Update
End With
CnG.CommitTrans

Exit Sub
err_capture:
rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
End Sub

Private Sub cmdSave_Click()
If int_mode = 1 Then
    If Not check_validate_new Then Exit Sub
    If check_validate_exist_new Then
        Call check_invalid
        Exit Sub
    End If
    Call insert_new_data
ElseIf int_mode = 2 Then
    If Not check_validate_new Then Exit Sub
    Call edit_old_data
End If

Call load_data_working_day
int_mode = 0
Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
cmdNew.Enabled = a And blnUser_Add
cmdSave.Enabled = b
cmdEdit.Enabled = c And blnUser_Edit
cmdDelete.Enabled = d And blnUser_Delete
cmdCancel.Enabled = e

CmdPrint.Enabled = f
cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        If Not LCase(Ctr.name) = "txt_working_time_name" Then Ctr.Text = ""
    ElseIf TypeOf Ctr Is TDBCombo Then
        If Not LCase(Ctr.name) = "tdbcombo_working_time" Then Ctr.Text = ""
    ElseIf TypeOf Ctr Is DTPicker Then
        Ctr.Value = Now
    End If
Next
End Sub

Private Sub set_enabled_control(ByVal i As Boolean)
Dim Ctr As CONTROL
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        Ctr.Enabled = i
    ElseIf TypeOf Ctr Is TDBCombo Then
        Ctr.Enabled = i
    ElseIf TypeOf Ctr Is DTPicker Then
        Ctr.Value = Now
        Ctr.Enabled = i
    End If
Next
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    
    '.Fields("shift_code").Value = TDBCombo_working_time.Columns("shift_code").Value
    cbo_working_day.ListIndex = .Fields("day_code").Value
    '-----------------------------------------------------------------------------
    DTPicker_start_time.Value = .Fields("start_time").Value
    DTPicker_end_time.Value = .Fields("end_time").Value
    cbo_active.ListIndex = .Fields("flag_active").Value
    
    '-- additional 4 BPP
    DTPicker_min_break_in = IIf(IsNull(.Fields("min_break_in").Value) = True, Now, .Fields("min_break_in").Value)
    DTPicker_max_break_out = IIf(IsNull(.Fields("max_break_out").Value) = True, Now, .Fields("max_break_out").Value)
    txt_break_interval_minute = Val("" & .Fields("break_interval_minute").Value)
    
End With
End Sub

Private Sub set_new_data()
DTPicker_start_time.Value = Format(Now, "yyyy-MM-dd ") & "08:30:00"
DTPicker_end_time.Value = Format(Now, "yyyy-MM-dd ") & "17:00:00"

cbo_working_day.ListIndex = 0
cbo_active.ListIndex = 1
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    cbo_working_day.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If cbo_working_day.Enabled = True Then
        cbo_working_day.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    cbo_working_day.Enabled = False
    fra_entry.Visible = True
    TDBGrid1.Enabled = False
End If
End Sub

Private Sub load_mode()
If int_mode = 1 Then        ' new
    Call set_buttons_enable(False, True, False, False, True, False, False)
ElseIf int_mode = 0 Then    ' view
    Call set_buttons_enable(True, False, True, True, False, True, True)
ElseIf int_mode = 2 Then    ' edit/revise
    Call set_buttons_enable(False, True, False, False, True, False, False)
End If

Call set_data_mode
End Sub

Private Sub Command1_Click()
MsgBox Chr$(40)
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn
Adodc_working_time.ConnectionString = strConn

Call load_data

Call load_data_user_access(Me)
int_mode = 0
Call load_mode
End Sub

Private Function get_setting_rekening(ByVal int_no As Integer) As Boolean
Dim rs As New ADODB.Recordset

rs.Open "exec sp_get_setting_rekening " & int_no, CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    str_kode_rekening = rs.Fields("kode_rekening").Value
    get_setting_rekening = True
Else
    str_kode_rekening = ""
    get_setting_rekening = False
End If
End Function

Private Sub txtFax_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select

End Sub

Private Sub txtTelp1_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txtTelp2_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Function get_inc_kode() As String
Dim rs As New ADODB.Recordset
Dim str_inc_kode As String

rs.Open "select max(kode_supplier) as curr_kode from m_supplier", CnG, adOpenStatic, adLockReadOnly
If rs.RecordCount > 0 Then
    If IsNull(rs.Fields("curr_kode").Value) = True Then
        str_inc_kode = "00001"
    Else
        str_inc_kode = rs.Fields("curr_kode").Value
        str_inc_kode = Right("0000" & Trim(str(CLng(str_inc_kode) + 1)), 5)
    End If
End If

get_inc_kode = str_inc_kode
End Function


Private Sub txtTelp3_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 45, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
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

Private Sub TDBCombo_working_time_ItemChange()
If TDBCombo_working_time.ApproxCount > 0 Then
    TDBCombo_working_time.Text = TDBCombo_working_time.Columns("shift_code").Value
    txt_working_time_name = TDBCombo_working_time.Columns("shift_name").Value
    
    Call load_data_working_day
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
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub TDBGrid1_FormatText _
(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "IN" _
Or TDBGrid1.Columns(ColIndex).Caption = "OUT" _
Or TDBGrid1.Columns(ColIndex).Caption = "MIN BREAK IN" _
Or TDBGrid1.Columns(ColIndex).Caption = "MAX BREAK OUT" Then
    Value = Format(Value, "hh:nn")
End If
End Sub

Private Sub Timer1_Timer()
Call load_data_shift
timer1.Enabled = False
End Sub

Private Sub load_data_shift()
Adodc_working_time.RecordSource = "select * from m_shift where flag_shift=0 order by shift_code"
Adodc_working_time.Refresh

TDBCombo_working_time.RowSource = Adodc_working_time
End Sub

Private Sub load_data_working_day()
Adodc1.RecordSource = "select * from m_working_day where shift_code='" _
& TDBCombo_working_time.Columns("shift_code").Value & "' order by day_code"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub
