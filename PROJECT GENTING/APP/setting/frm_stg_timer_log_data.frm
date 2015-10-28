VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_stg_all 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SETTING"
   ClientHeight    =   5925
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   10995
   Icon            =   "frm_stg_timer_log_data.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   675
      Left            =   9900
      OleObjectBlob   =   "frm_stg_timer_log_data.frx":000C
      TabIndex        =   29
      Top             =   180
      Visible         =   0   'False
      Width           =   555
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "DOWNLOAD MODE"
      TabPicture(0)   =   "frm_stg_timer_log_data.frx":0030
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "fra_auto_log"
      Tab(0).Control(2)=   "cmdEdit"
      Tab(0).Control(3)=   "CmdCancel"
      Tab(0).Control(4)=   "CmdNew"
      Tab(0).Control(5)=   "cmdDelete"
      Tab(0).Control(6)=   "cmdSave"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "DEVICE TYPE"
      TabPicture(1)   =   "frm_stg_timer_log_data.frx":004C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_entry_wt"
      Tab(1).Control(1)=   "cmd_update_device_type"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "DEVICE TIME"
      TabPicture(2)   =   "frm_stg_timer_log_data.frx":0068
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmd_dt_refresh"
      Tab(2).Control(1)=   "cmd_dt_set"
      Tab(2).Control(2)=   "DTPicker_device_time"
      Tab(2).Control(3)=   "TDBCombo_device"
      Tab(2).Control(4)=   "Label8"
      Tab(2).Control(5)=   "Line2"
      Tab(2).Control(6)=   "Label7"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "MAIL"
      TabPicture(3)   =   "frm_stg_timer_log_data.frx":0084
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "cmd_update_mail"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame2"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "SUPER ADMIN"
      TabPicture(4)   =   "frm_stg_timer_log_data.frx":00A0
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Line3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label9"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label10"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Command1"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "txt_password"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "txt_username"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).ControlCount=   6
      Begin VB.Frame Frame2 
         Caption         =   "General"
         Height          =   1365
         Left            =   1770
         TabIndex        =   44
         Top             =   690
         Width           =   5685
         Begin VB.TextBox txt_sender_name 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1770
            TabIndex        =   46
            Top             =   330
            Width           =   3615
         End
         Begin VB.TextBox txt_sender_email 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1770
            TabIndex        =   45
            Top             =   780
            Width           =   3615
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "SENDER NAME*"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   390
            Width           =   1245
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "SENDER EMAIL*"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   870
            Width           =   1260
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Server"
         Height          =   1905
         Left            =   1770
         TabIndex        =   35
         Top             =   2190
         Width           =   5685
         Begin VB.TextBox txt_smtp_port 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1770
            TabIndex        =   39
            Text            =   "25"
            Top             =   1440
            Width           =   1605
         End
         Begin VB.TextBox txt_sender_pwd 
            Appearance      =   0  'Flat
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1770
            PasswordChar    =   "*"
            TabIndex        =   38
            Top             =   600
            Width           =   1635
         End
         Begin VB.TextBox txt_smtp 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1770
            TabIndex        =   37
            Top             =   1020
            Width           =   3615
         End
         Begin VB.TextBox txt_mail_username 
            Appearance      =   0  'Flat
            Height          =   375
            Left            =   1770
            TabIndex        =   36
            Top             =   180
            Width           =   3615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "USERNAME*"
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "PORT*"
            Height          =   195
            Left            =   240
            TabIndex        =   42
            Top             =   1500
            Width           =   510
         End
         Begin VB.Label lblPwd 
            AutoSize        =   -1  'True
            Caption         =   "PASSWORD*"
            Height          =   195
            Left            =   240
            TabIndex        =   41
            Top             =   690
            Width           =   1005
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "SMTP*"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   1110
            Width           =   510
         End
      End
      Begin VB.TextBox txt_username 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   -71220
         TabIndex        =   32
         Top             =   1650
         Width           =   3615
      End
      Begin VB.TextBox txt_password 
         Appearance      =   0  'Flat
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   -71220
         PasswordChar    =   "*"
         TabIndex        =   31
         Top             =   2130
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Update"
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
         Left            =   -68580
         Picture         =   "frm_stg_timer_log_data.frx":00BC
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3480
         Width           =   975
      End
      Begin VB.CommandButton cmd_update_device_type 
         Caption         =   "&Update"
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
         Left            =   -68160
         Picture         =   "frm_stg_timer_log_data.frx":0646
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton cmd_update_mail 
         Caption         =   "&Update"
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
         Left            =   6450
         Picture         =   "frm_stg_timer_log_data.frx":0BD0
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   4140
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   -71400
         Picture         =   "frm_stg_timer_log_data.frx":115A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   -69240
         Picture         =   "frm_stg_timer_log_data.frx":16E4
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   -72480
         Picture         =   "frm_stg_timer_log_data.frx":1C6E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   -68160
         Picture         =   "frm_stg_timer_log_data.frx":21F8
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   -70320
         Picture         =   "frm_stg_timer_log_data.frx":2782
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmd_dt_refresh 
         Caption         =   "&Refresh"
         Height          =   645
         Left            =   -69480
         Picture         =   "frm_stg_timer_log_data.frx":2D0C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3120
         Width           =   975
      End
      Begin VB.CommandButton cmd_dt_set 
         Caption         =   "&Set"
         Height          =   645
         Left            =   -68400
         Picture         =   "frm_stg_timer_log_data.frx":3296
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3120
         Width           =   975
      End
      Begin VB.Frame fra_entry_wt 
         Height          =   1215
         Left            =   -73800
         TabIndex        =   17
         Top             =   1680
         Width           =   6615
         Begin VB.OptionButton opt_base_no_device 
            Caption         =   "NO DEVICE"
            Height          =   255
            Left            =   5040
            TabIndex        =   28
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton opt_base_ip 
            Caption         =   "STAND ALONE (BASE IP)"
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   480
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton opt_base_sensor 
            Caption         =   "BASE SENSOR"
            Height          =   255
            Left            =   3120
            TabIndex        =   18
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame fra_auto_log 
         Caption         =   "AUTO"
         Height          =   2895
         Left            =   -72480
         TabIndex        =   10
         Top             =   720
         Width           =   5295
         Begin VB.Frame fra_entry 
            Caption         =   "Entry"
            Height          =   1335
            Left            =   480
            TabIndex        =   12
            Top             =   1320
            Width           =   4455
            Begin VB.ComboBox cbo_auto_log_enable 
               Height          =   315
               ItemData        =   "frm_stg_timer_log_data.frx":3820
               Left            =   1800
               List            =   "frm_stg_timer_log_data.frx":382A
               TabIndex        =   15
               Text            =   "..."
               Top             =   720
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker_download_time 
               Height          =   315
               Left            =   1800
               TabIndex        =   13
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               MousePointer    =   99
               CustomFormat    =   "HH:mm"
               Format          =   118030339
               UpDown          =   -1  'True
               CurrentDate     =   39270.5
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "ENABLE"
               Height          =   195
               Left            =   600
               TabIndex        =   16
               Top             =   720
               Width           =   630
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "TIME"
               Height          =   195
               Left            =   600
               TabIndex        =   14
               Top             =   360
               Width           =   390
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   2295
            Left            =   480
            TabIndex        =   11
            Top             =   360
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4048
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "NUMBER"
            Columns(0).DataField=   "s_number"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "DOWNLOAD TIME"
            Columns(1).DataField=   "s_time"
            Columns(1).NumberFormat=   "FormatText Event"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   4
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "ENABLED"
            Columns(2).DataField=   "s_enable"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=370"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=291"
            Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
            Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(9)=   "Column(1).Width=3784"
            Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3704"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=2805"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=2725"
            Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
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
            Caption         =   "LIST OF DOWNLOAD TIME"
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
            _StyleDefs(46)  =   "Named:id=33:Normal"
            _StyleDefs(47)  =   ":id=33,.parent=0"
            _StyleDefs(48)  =   "Named:id=34:Heading"
            _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   ":id=34,.wraptext=-1"
            _StyleDefs(51)  =   "Named:id=35:Footing"
            _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   "Named:id=36:Selected"
            _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(55)  =   "Named:id=37:Caption"
            _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(57)  =   "Named:id=38:HighlightRow"
            _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(59)  =   "Named:id=39:EvenRow"
            _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(61)  =   "Named:id=40:OddRow"
            _StyleDefs(62)  =   ":id=40,.parent=33"
            _StyleDefs(63)  =   "Named:id=41:RecordSelector"
            _StyleDefs(64)  =   ":id=41,.parent=34"
            _StyleDefs(65)  =   "Named:id=42:FilterBar"
            _StyleDefs(66)  =   ":id=42,.parent=33"
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   120
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
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   -73920
         TabIndex        =   7
         Top             =   600
         Width           =   1695
         Begin VB.OptionButton opt_manual_download 
            Caption         =   "MANUAL"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opt_auto_download 
            Caption         =   "AUTO"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   120
            Width           =   1095
         End
      End
      Begin MSComCtl2.DTPicker DTPicker_device_time 
         Height          =   315
         Left            =   -71160
         TabIndex        =   20
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   43581443
         CurrentDate     =   39270.5
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_device 
         Height          =   375
         Left            =   -71160
         OleObjectBlob   =   "frm_stg_timer_log_data.frx":3837
         TabIndex        =   26
         Top             =   1920
         Width           =   3735
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "USERNAME"
         Height          =   195
         Left            =   -72780
         TabIndex        =   34
         Top             =   1650
         Width           =   915
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "PASSWORD"
         Height          =   195
         Left            =   -72780
         TabIndex        =   33
         Top             =   2130
         Width           =   945
      End
      Begin VB.Line Line3 
         X1              =   -72780
         X2              =   -67620
         Y1              =   3090
         Y2              =   3090
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "DEVICE IP"
         Height          =   195
         Left            =   -72600
         TabIndex        =   25
         Top             =   1920
         Width           =   780
      End
      Begin VB.Line Line2 
         X1              =   -72600
         X2              =   -67440
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "SYSTEM TIME"
         Height          =   195
         Left            =   -72600
         TabIndex        =   24
         Top             =   2280
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_exit 
      Caption         =   "E&xit"
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
      Left            =   9600
      Picture         =   "frm_stg_timer_log_data.frx":5CB4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4890
      Width           =   975
   End
End
Attribute VB_Name = "frm_stg_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim cn_fp As Boolean

Private Sub cmd_dt_refresh_Click()
DTPicker_device_time.Value = Now
End Sub

Private Sub cmd_dt_set_Click()
If check_validate_tdbcombo(TDBCombo_device) = False Then
    MsgBox "No valid data!", vbInformation, headerMSG
    Exit Sub
End If

FG_IP_ADDRESS = TDBCombo_device.Columns("ip_address").Value
FG_PORT_NUMBER = TDBCombo_device.Columns("port_number").Value

If Not connect Then
    MsgBox "Error connecting to source device...", vbCritical, headerMSG
    Exit Sub
End If

Call set_device_time

Call disconnect
End Sub

Private Sub set_device_time()
'Dim iYear As Long
'Dim iMonth As Long
'Dim iDay  As Long
'Dim iHour As Long
'Dim iMinute  As Long
'Dim iSecond As Long
Dim iYear As Integer
Dim iMonth As Integer
Dim iDay  As Integer
Dim iHour As Integer
Dim iMinute  As Integer
Dim iSecond As Integer

'iYear = CLng(Year(DTPicker_device_time.Value))
'iMonth = CLng(Month(DTPicker_device_time.Value))
'iDay = CLng(Day(DTPicker_device_time.Value))
'iHour = CLng(Hour(DTPicker_device_time.Value))
'iMinute = CLng(Minute(DTPicker_device_time.Value))
'iSecond = CLng(Second(DTPicker_device_time.Value))

iYear = CInt(year(DTPicker_device_time.Value))
iMonth = CInt(month(DTPicker_device_time.Value))
iDay = CInt(Day(DTPicker_device_time.Value))
iHour = CInt(Hour(DTPicker_device_time.Value))
iMinute = CInt(Minute(DTPicker_device_time.Value))
iSecond = CInt(Second(DTPicker_device_time.Value))

If Not CZKEM1.SetDeviceTime2(vMachineNumber, iYear, iMonth, iDay, iHour, iMinute, iSecond) Then
'If Not CZKEM1.SetDeviceTime2(vMachineNumber, iYear, 11, 11, 11, 30, 35) Then
    CZKEM1.ClearLCD
    CZKEM1.EnableClock True
    CZKEM1.EnableDevice 1, True
    MsgBox "Lets check device time!" & vbCr _
        & "Please try again if no change...", vbInformation, headerMSG
Else
    CZKEM1.ClearLCD
    CZKEM1.EnableClock True
    CZKEM1.EnableDevice 1, True
    MsgBox "Device time was successfully changed", vbInformation, headerMSG
End If
End Sub

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

Private Sub cmd_exit_Click()
Unload Me
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

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
CmdNew.Enabled = a And blnUser_Add
cmdSave.Enabled = b
cmdEdit.Enabled = c And blnUser_Edit
cmdDelete.Enabled = d And blnUser_Delete
CmdCancel.Enabled = e

'CmdPrint.Enabled = f
'cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data()
Dim Ctr As Control
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        If Not LCase(Ctr.name) = "txt_company_name" And Not LCase(Ctr.name) = "txt_smtp" _
        And Not LCase(Ctr.name) = "txt_sender_email" And Not LCase(Ctr.name) = "txt_sender_name" _
        And Not LCase(Ctr.name) = "txt_username" And Not LCase(Ctr.name) = "txt_sender_pwd" _
        And Not LCase(Ctr.name) = "txt_smtp_port" And Not LCase(Ctr.name) = "txt_mail_username" Then
            Ctr.Text = ""
        End If
    ElseIf TypeOf Ctr Is TDBCombo Then
        If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
    ElseIf TypeOf Ctr Is DTPicker Then
        Ctr.Value = Now
    End If
Next
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    DTPicker_download_time.Value = .Fields("s_time").Value
    cbo_auto_log_enable.ListIndex = .Fields("s_enable").Value
End With
End Sub

Private Sub set_new_data()
DTPicker_download_time.Value = Now
cbo_auto_log_enable.ListIndex = 1
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    DTPicker_download_time.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If DTPicker_download_time.Enabled = True Then DTPicker_download_time.SetFocus
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    'DTPicker_download_time.Enabled = False
    fra_entry.Visible = True
    TDBGrid1.Enabled = False
End If
End Sub

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(s_time) as rec_count from s_auto_log where left(right(s_time,8),5) = '" _
& Format(DTPicker_download_time.Value, "HH:mm") & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
DTPicker_download_time.Value = Now
If DTPicker_download_time.Enabled = True Then DTPicker_download_time.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If (Not Format(DTPicker_download_time.Value, "HH:mm") = Format(Adodc1.Recordset.Fields("s_time").Value, "hh:nn")) _
And check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Sub insert_new_data()
CnG.BeginTrans

With rsBound
    .AddNew
    
    .Fields("s_time").Value = Format(DTPicker_download_time.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("s_enable").Value = cbo_auto_log_enable.ListIndex
    
    .Update
End With

CnG.CommitTrans
End Sub

Private Sub edit_old_data()
On Error GoTo err_capture

CnG.BeginTrans
With rsBound

    .Fields("s_time").Value = Format(DTPicker_download_time.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("s_enable").Value = cbo_auto_log_enable.ListIndex
    
    .Update
End With
CnG.CommitTrans

Exit Sub
err_capture:
rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
End Sub

Private Sub cmd_update_device_type_Click()
Call update_data_device_type
End Sub

Private Sub cmd_update_mail_Click()
Call update_data_mail
End Sub

Private Sub CmdCancel_Click()
int_mode = 0
Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "No Data selected!", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Are you sure want to delete data '" _
    & Format(TDBGrid1.Columns("s_time").Value, "hh:nn") & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from s_auto_log where s_number = " _
    & TDBGrid1.Columns("s_number").Value
CnG.CommitTrans

Call load_data_auto_log
int_mode = 0
Call load_mode
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from s_auto_log where s_number = " _
& Adodc1.Recordset.Fields("s_number").Value, CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from s_auto_log where s_number = -1", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub CmdSave_Click()
If SSTab1.Tab = 0 Then

    If int_mode = 1 Then
        'If Not check_validate_new Then Exit Sub
        If check_validate_exist_new Then Call check_invalid: Exit Sub
        Call insert_new_data
    ElseIf int_mode = 2 Then
        'If Not check_validate_new Then Exit Sub
        If check_validate_exist_edit Then Call check_invalid: Exit Sub
        Call edit_old_data
    End If
    
    Call load_data_auto_log
    int_mode = 0
    Call load_mode

End If
End Sub

Private Sub Command1_Click()
Call update_data_superadmin
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn

Call load_data_auto_log
Call load_data_download_mode
Call set_download_mode
Call load_data_mail
Call load_data_device
Call load_data_device_type
Call load_data_superadmin

DTPicker_device_time.Value = Now
vMachineNumber = 1
cn_fp = False

Call load_data_user_access(Me)
SSTab1.Tab = 0
Call SSTab1_Click(0)
'Timer1.Enabled = True
End Sub

Private Sub load_data_device()
Dim rs1 As New ADODB.Recordset

rs1.Open "select *, concat(name,' (',ip_address,')') as device_ip from m_device order by ip_address", CnG, adOpenStatic, adLockReadOnly
TDBCombo_device.RowSource = rs1
End Sub

Private Sub load_data_device_type()
Dim rs1 As New ADODB.Recordset

rs1.Open "select * from s_device where s_number = 1", CnG, adOpenStatic, adLockReadOnly
If rs1.RecordCount = 0 Then
    opt_base_ip = True
Else
    If rs1.Fields("s_device").Value = 1 Then
        opt_base_ip = True
    ElseIf rs1.Fields("s_device").Value = 2 Then
        opt_base_sensor = True
    Else
        opt_base_no_device = True
    End If
End If
End Sub

Private Sub update_data_device_type()
Dim rs1 As New ADODB.Recordset
Dim i As Integer

i = IIf(opt_base_ip, 1, IIf(opt_base_sensor, 2, 3))
'If opt_base_ip Then
'    i = 1
'ElseIf opt_base_sensor Then
'    i = 2
'ElseIf opt_base_no_device Then
'    i = 3
'End If

CnG.BeginTrans

CnG.Execute "delete from s_device where s_number=1"
rs1.Open "select * from s_device where s_number=-77", CnG, adOpenKeyset, adLockOptimistic

rs1.AddNew
    rs1.Fields("s_number").Value = 1
    rs1.Fields("s_device").Value = i
    rs1.Fields("description").Value = "setting device type"
rs1.Update

CnG.CommitTrans
End Sub

Private Sub update_data_superadmin()
Dim rs1 As New ADODB.Recordset
Dim strsql As String

CnG.BeginTrans

CnG.Execute "delete from m_user where user_level = 100"

strsql = "insert into m_user (user_code, user_name, user_pass, user_pass_key, user_level) " _
        & "values('adminSSD', '" & Replace$(Trim$(txt_username), Chr(39), Chr(96)) _
        & "','" & RC4DeCryptASC(Replace(Trim(txt_password.Text), Chr(39), Chr(96)), pEncryptionPassword) _
        & "','" & pEncryptionPassword & "',100)"
CnG.Execute strsql

strsql = "Insert into t_user_access_level (level_code,access_level_code,level_name,allow_access) " _
& "Select 'adminSSD', code, name, 1 " _
& "from m_akses_level_group where code not in " _
& "(select access_level_code from t_user_access_level where level_code = 'adminSSD')"
CnG.Execute strsql

CnG.CommitTrans

MsgBox "Update Successfully", vbInformation, headerMSG
End Sub

Private Sub update_data_mail()
Dim rs1 As New ADODB.Recordset
Dim strsql As String

If txt_sender_email.Text = "" Then
    MsgBox "Please Fill Sender Mail!", vbExclamation, headerMSG
    Exit Sub
End If

If txt_sender_name.Text = "" Then
    MsgBox "Please Fill Sender Name!", vbExclamation, headerMSG
    Exit Sub
End If

If txt_mail_username.Text = "" Then
    MsgBox "Please Fill Username!", vbExclamation, headerMSG
    Exit Sub
End If

If txt_sender_pwd.Text = "" Then
    MsgBox "Please Fill Sender Password!", vbExclamation, headerMSG
    Exit Sub
End If

If txt_smtp.Text = "" Then
    MsgBox "Please Fill SMTP!", vbExclamation, headerMSG
    Exit Sub
End If

If txt_smtp_port.Text = "" Then
    MsgBox "Please Fill SMTP Port!", vbExclamation, headerMSG
    Exit Sub
End If

CnG.BeginTrans

strsql = "UPDATE s_mail SET " _
        & "s_sender_email = '" & txt_sender_email.Text & "'," _
        & "s_sender_name = '" & txt_sender_name.Text & "'," _
        & "username = '" & txt_mail_username.Text & "'," _
        & "password = '" & Replace(RC4DeCryptASC(Trim(txt_sender_pwd.Text), pEncryptionPassword), "'", "") & "'," _
        & "smtp = '" & txt_smtp.Text & "'," _
        & "port = '" & txt_smtp_port.Text & "' " _
        & "WHERE s_number = 1"
CnG.Execute (strsql)

CnG.CommitTrans

MsgBox "Save Succesfully!", vbInformation, headerMSG
End Sub

Private Sub load_data_superadmin()
Dim rs1 As New ADODB.Recordset

rs1.Open "select * from m_user where user_level = 100", CnG, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then
    txt_username = rs1.Fields("user_name").Value
    txt_password = RC4DeCryptASC(Replace(Trim(rs1.Fields("user_pass").Value), Chr(39), Chr(96)), pEncryptionPassword)
Else
    txt_username = ""
    txt_password = ""
End If
End Sub

Private Sub load_data_mail()
Dim rs1 As New ADODB.Recordset
Dim strsql As String

strsql = "select * from s_mail where s_number = 1"
rs1.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly

'rs1.Open "select * from s_mail where s_number = 1", CnG, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then
    txt_sender_email.Text = rs1!s_sender_email
    txt_sender_name.Text = rs1!s_sender_name
    txt_smtp.Text = rs1!smtp
    txt_smtp_port.Text = rs1!PORT
    txt_mail_username.Text = rs1!Username
    txt_sender_pwd.Text = RC4DeCryptASC(rs1!Password, pEncryptionPassword)
Else
    txt_smtp.Text = ""
    txt_sender_email.Text = ""
    txt_sender_name.Text = ""
    txt_sender_pwd.Text = ""
    txt_smtp_port.Text = ""
    txt_mail_username.Text = ""
End If
End Sub

Private Sub load_data_auto_log()
Adodc1.RecordSource = "select * from s_auto_log order by right(s_time,8) asc"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub load_data_download_mode()
opt_auto_download.Value = IIf(BLN_AUTO_LOG, True, False)
End Sub

Private Sub set_download_mode()
If opt_auto_download Then
    fra_auto_log.Enabled = True
    Call set_buttons_enable(True, False, True, True, False, True, True)
ElseIf opt_manual_download Then
    fra_auto_log.Enabled = False
    Call set_buttons_enable(False, False, False, False, False, False, False)
End If
End Sub

Private Sub update_download_mode()
CnG.Execute "delete from s_download_mode where s_number = 1"
CnG.Execute "insert into s_download_mode(s_number, s_auto_download) " _
                                & "values (1," & IIf(opt_auto_download, 1, 0) & ")"
BLN_AUTO_LOG = IIf(opt_auto_download, True, False)
mdi_absensi.timer_get_log_data.Enabled = BLN_AUTO_LOG
End Sub

Private Sub opt_auto_download_Click()
Call set_download_mode
Call update_download_mode
End Sub

Private Sub opt_manual_download_Click()
Call set_download_mode
Call update_download_mode
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Tab = 0 Then
    int_mode = 0
    Call load_mode
ElseIf SSTab1.Tab = 1 Then
    Call set_buttons_enable(False, True, False, False, False, False, False)
End If
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "DOWNLOAD TIME" Then Value = Format(Value, "hh:nn")
End Sub


