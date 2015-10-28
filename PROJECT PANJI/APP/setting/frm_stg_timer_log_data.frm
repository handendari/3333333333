VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_stg_all 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ALAT"
   ClientHeight    =   5820
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10950
   Icon            =   "frm_stg_timer_log_data.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   675
      Left            =   9900
      OleObjectBlob   =   "frm_stg_timer_log_data.frx":058A
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   555
   End
   Begin prj_panji.vbButton cmd_exit 
      Height          =   705
      Left            =   9810
      TabIndex        =   1
      Top             =   4800
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1244
      BTYPE           =   14
      TX              =   "&Keluar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_stg_timer_log_data.frx":05AE
      PICN            =   "frm_stg_timer_log_data.frx":05CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   360
      TabIndex        =   2
      Top             =   210
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "TIPE FINGERPRINT"
      TabPicture(0)   =   "frm_stg_timer_log_data.frx":165C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmd_update_device_type"
      Tab(0).Control(1)=   "fra_entry_wt"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "TAMBAH FINGERPRINT"
      TabPicture(1)   =   "frm_stg_timer_log_data.frx":1678
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TDBGrid2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraEntry_Device"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "frmTombol"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "SETTING WAKTU"
      TabPicture(2)   =   "frm_stg_timer_log_data.frx":1694
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(1)=   "Line2"
      Tab(2).Control(2)=   "Label7"
      Tab(2).Control(3)=   "TDBCombo_device"
      Tab(2).Control(4)=   "DTPicker_device_time"
      Tab(2).Control(5)=   "cmd_dt_refresh"
      Tab(2).Control(6)=   "cmd_dt_set"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "MODE DOWNLOAD"
      TabPicture(3)   =   "frm_stg_timer_log_data.frx":16B0
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fra_auto_log"
      Tab(3).Control(1)=   "Frame4"
      Tab(3).Control(2)=   "Frame1"
      Tab(3).ControlCount=   3
      Begin VB.Frame fra_auto_log 
         Caption         =   "AUTO"
         Height          =   2895
         Left            =   -72480
         TabIndex        =   34
         Top             =   750
         Width           =   5295
         Begin VB.Frame fra_entry 
            Caption         =   "Entry"
            Height          =   1335
            Left            =   480
            TabIndex        =   35
            Top             =   1320
            Visible         =   0   'False
            Width           =   4455
            Begin VB.ComboBox cbo_auto_log_enable 
               Height          =   315
               ItemData        =   "frm_stg_timer_log_data.frx":16CC
               Left            =   1800
               List            =   "frm_stg_timer_log_data.frx":16D6
               TabIndex        =   36
               Text            =   "..."
               Top             =   720
               Width           =   1335
            End
            Begin MSComCtl2.DTPicker DTPicker_download_time 
               Height          =   315
               Left            =   1800
               TabIndex        =   37
               Top             =   360
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   556
               _Version        =   393216
               MousePointer    =   99
               CustomFormat    =   "HH:mm"
               Format          =   167837699
               UpDown          =   -1  'True
               CurrentDate     =   39270.5
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "WAKTU"
               Height          =   195
               Left            =   600
               TabIndex        =   39
               Top             =   360
               Width           =   600
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "ENABLE"
               Height          =   195
               Left            =   600
               TabIndex        =   38
               Top             =   720
               Width           =   630
            End
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   2295
            Left            =   480
            TabIndex        =   40
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
            Columns(1).Caption=   "WAKTU DOWNLOAD"
            Columns(1).DataField=   "s_time"
            Columns(1).NumberFormat=   "HH:mm"
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
            Caption         =   "WAKTU DOWNLOAD"
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
      End
      Begin VB.CommandButton cmd_dt_set 
         Caption         =   "&Set"
         Height          =   645
         Left            =   -68400
         Picture         =   "frm_stg_timer_log_data.frx":16E3
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   2820
         Width           =   975
      End
      Begin VB.CommandButton cmd_dt_refresh 
         Caption         =   "&Refresh"
         Height          =   645
         Left            =   -69480
         Picture         =   "frm_stg_timer_log_data.frx":1C6D
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   2820
         Width           =   975
      End
      Begin VB.Frame fra_entry_wt 
         Height          =   1215
         Left            =   -73740
         TabIndex        =   24
         Top             =   1440
         Width           =   6615
         Begin VB.OptionButton opt_base_no_device 
            Caption         =   "NO DEVICE"
            Height          =   255
            Left            =   5040
            TabIndex        =   27
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton opt_base_ip 
            Caption         =   "STAND ALONE (BASE IP)"
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   480
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.OptionButton opt_base_sensor 
            Caption         =   "BASE SENSOR"
            Height          =   255
            Left            =   3120
            TabIndex        =   25
            Top             =   480
            Width           =   1815
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1185
         Left            =   330
         TabIndex        =   17
         Top             =   3870
         Width           =   8715
         Begin VB.Timer timer1 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew_Device 
            Height          =   705
            Left            =   150
            TabIndex        =   18
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Tambah"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":21F7
            PICN            =   "frm_stg_timer_log_data.frx":2213
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Device 
            Height          =   705
            Left            =   1200
            TabIndex        =   19
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Simpan"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":32A5
            PICN            =   "frm_stg_timer_log_data.frx":32C1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Device 
            Height          =   705
            Left            =   2250
            TabIndex        =   20
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Ubah"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":4353
            PICN            =   "frm_stg_timer_log_data.frx":436F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Device 
            Height          =   705
            Left            =   3300
            TabIndex        =   21
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Hapus"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":5401
            PICN            =   "frm_stg_timer_log_data.frx":541D
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Device 
            Height          =   705
            Left            =   4350
            TabIndex        =   22
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Batal"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":64AF
            PICN            =   "frm_stg_timer_log_data.frx":64CB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmd_test 
            Height          =   705
            Left            =   7620
            TabIndex        =   23
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Test"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":755D
            PICN            =   "frm_stg_timer_log_data.frx":7579
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame fraEntry_Device 
         Height          =   2055
         Left            =   330
         TabIndex        =   10
         Top             =   1770
         Visible         =   0   'False
         Width           =   8715
         Begin VB.ComboBox cboType 
            Height          =   315
            ItemData        =   "frm_stg_timer_log_data.frx":860B
            Left            =   3660
            List            =   "frm_stg_timer_log_data.frx":8615
            TabIndex        =   46
            Text            =   "TIPE A"
            Top             =   1140
            Width           =   1605
         End
         Begin VB.TextBox txt_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3660
            MaxLength       =   50
            TabIndex        =   13
            Top             =   1500
            Width           =   3135
         End
         Begin VB.TextBox txt_port 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3660
            MaxLength       =   50
            TabIndex        =   12
            Top             =   780
            Width           =   1575
         End
         Begin VB.TextBox txt_ip_address 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3660
            MaxLength       =   50
            TabIndex        =   11
            Top             =   420
            Width           =   3135
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "TYPE"
            Height          =   195
            Left            =   2220
            TabIndex        =   47
            Top             =   1170
            Width           =   420
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   2220
            TabIndex        =   16
            Top             =   1530
            Width           =   1110
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "PORT"
            Height          =   195
            Left            =   2220
            TabIndex        =   15
            Top             =   780
            Width           =   450
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "IP ADDRESS"
            Height          =   195
            Left            =   2220
            TabIndex        =   14
            Top             =   420
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data Control Button"
         Height          =   1185
         Left            =   -72480
         TabIndex        =   3
         Top             =   3630
         Width           =   5295
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Left            =   120
            TabIndex        =   4
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Tambah"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":8629
            PICN            =   "frm_stg_timer_log_data.frx":8645
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave 
            Height          =   705
            Left            =   1140
            TabIndex        =   5
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Simpan"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":96D7
            PICN            =   "frm_stg_timer_log_data.frx":96F3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit 
            Height          =   705
            Left            =   2160
            TabIndex        =   6
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Ubah"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":A785
            PICN            =   "frm_stg_timer_log_data.frx":A7A1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete 
            Height          =   705
            Left            =   3180
            TabIndex        =   7
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Hapus"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":B833
            PICN            =   "frm_stg_timer_log_data.frx":B84F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel 
            Height          =   705
            Left            =   4200
            TabIndex        =   8
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Batal"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":C8E1
            PICN            =   "frm_stg_timer_log_data.frx":C8FD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton vbButton6 
            Height          =   705
            Left            =   7620
            TabIndex        =   9
            Top             =   330
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Test"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_stg_timer_log_data.frx":D98F
            PICN            =   "frm_stg_timer_log_data.frx":D9AB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin MSComCtl2.DTPicker DTPicker_device_time 
         Height          =   315
         Left            =   -71160
         TabIndex        =   30
         Top             =   1980
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   167772163
         CurrentDate     =   39270.5
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_device 
         Height          =   375
         Left            =   -71160
         OleObjectBlob   =   "frm_stg_timer_log_data.frx":EA3D
         TabIndex        =   31
         Top             =   1620
         Width           =   3735
      End
      Begin prj_panji.vbButton cmd_update_device_type 
         Height          =   450
         Left            =   -68460
         TabIndex        =   32
         Top             =   2970
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   794
         BTYPE           =   14
         TX              =   "&Update"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_stg_timer_log_data.frx":10EBA
         PICN            =   "frm_stg_timer_log_data.frx":10ED6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
         Height          =   3255
         Left            =   330
         TabIndex        =   33
         Top             =   570
         Width           =   8715
         _ExtentX        =   15372
         _ExtentY        =   5741
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "IP ADDRESS"
         Columns(0).DataField=   "ip_address"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "PORT"
         Columns(1).DataField=   "port_number"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "FLAG TYPE"
         Columns(2).DataField=   "flag_type"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "TYPE"
         Columns(3).DataField=   "type"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "KETERANGAN"
         Columns(4).DataField=   "name"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4604"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4524"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2275"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2196"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2090"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2011"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=5371"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=5292"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=512"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
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
         Caption         =   "DAFTAR FINGERPRINT"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(54)  =   "Named:id=33:Normal"
         _StyleDefs(55)  =   ":id=33,.parent=0"
         _StyleDefs(56)  =   "Named:id=34:Heading"
         _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(58)  =   ":id=34,.wraptext=-1"
         _StyleDefs(59)  =   "Named:id=35:Footing"
         _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=36:Selected"
         _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=37:Caption"
         _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(65)  =   "Named:id=38:HighlightRow"
         _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(67)  =   "Named:id=39:EvenRow"
         _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(69)  =   "Named:id=40:OddRow"
         _StyleDefs(70)  =   ":id=40,.parent=33"
         _StyleDefs(71)  =   "Named:id=41:RecordSelector"
         _StyleDefs(72)  =   ":id=41,.parent=34"
         _StyleDefs(73)  =   "Named:id=42:FilterBar"
         _StyleDefs(74)  =   ":id=42,.parent=33"
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1455
         Left            =   -73920
         TabIndex        =   41
         Top             =   750
         Width           =   1695
         Begin VB.OptionButton opt_auto_download 
            Caption         =   "AUTO"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton opt_manual_download 
            Caption         =   "MANUAL"
            Height          =   255
            Left            =   240
            TabIndex        =   42
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "WAKTU SISTEM"
         Height          =   195
         Left            =   -72600
         TabIndex        =   45
         Top             =   1980
         Width           =   1245
      End
      Begin VB.Line Line2 
         X1              =   -72600
         X2              =   -67440
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ALAMAT IP"
         Height          =   195
         Left            =   -72600
         TabIndex        =   44
         Top             =   1620
         Width           =   840
      End
   End
End
Attribute VB_Name = "frm_stg_all"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsDevice As New ADODB.Recordset
Dim rsDevice_Tdb As New ADODB.Recordset
Dim rsDownload As New ADODB.Recordset
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim int_mode As Integer
Dim cn_fp As Boolean
Dim v_s_number As Integer

Dim vIPAddress As String

Private Sub cmd_dt_refresh_Click()
    DTPicker_device_time.Value = Now
End Sub

Private Sub cmd_dt_set_Click()
    If check_validate_tdbcombo(TDBCombo_device) = False Then
        MsgBox "Data Tidak Valid...", vbInformation, headerMSG
        Exit Sub
    End If
    
    FG_IP_ADDRESS = TDBCombo_device.Columns("ip_address").Value
    FG_PORT_NUMBER = TDBCombo_device.Columns("port_number").Value
    
    If Not connect Then
        MsgBox "Gagal Menghubungkan Fingerprint...", vbCritical, headerMSG
        Exit Sub
    End If
    
    Call set_device_time
    Call disconnect
End Sub

Private Sub set_device_time()
Dim iYear As Integer
Dim iMonth As Integer
Dim iDay  As Integer
Dim iHour As Integer
Dim iMinute  As Integer
Dim iSecond As Integer
    
    iYear = CInt(year(DTPicker_device_time.Value))
    iMonth = CInt(month(DTPicker_device_time.Value))
    iDay = CInt(Day(DTPicker_device_time.Value))
    iHour = CInt(hour(DTPicker_device_time.Value))
    iMinute = CInt(Minute(DTPicker_device_time.Value))
    iSecond = CInt(Second(DTPicker_device_time.Value))
    
    If Not CZKEM1.SetDeviceTime2(vMachineNumber, iYear, iMonth, iDay, iHour, iMinute, iSecond) Then
        CZKEM1.ClearLCD
        CZKEM1.EnableClock True
        CZKEM1.EnableDevice 1, True
        MsgBox "Cek Waktu Di Fingerprint" & vbCr _
            & "Coba Ulangi Jika Tidak Ada Perubahan...", vbInformation, headerMSG
    Else
        CZKEM1.ClearLCD
        CZKEM1.EnableClock True
        CZKEM1.EnableDevice 1, True
        MsgBox "Waktu Fingerprint Berhasil Diubah...", vbInformation, headerMSG
    End If
End Sub

Private Function connect() As Boolean
    If cn_fp Then
        CZKEM1.EnableDevice vMachineNumber, True
        CZKEM1.disconnect
    End If
    
    FG_IP_ADDRESS = rsDevice.Fields("ip_address").Value
    FG_PORT_NUMBER = rsDevice.Fields("port_number").Value
    
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
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
    cmdNew.Enabled = a And blnUser_Add
    cmdSave.Enabled = b
    cmdEdit.Enabled = c And blnUser_Edit
    cmdDelete.Enabled = d And blnUser_Delete
    cmdCancel.Enabled = e
    
    cmdNew_Device.Enabled = a And blnUser_Add
    cmdSave_Device.Enabled = b
    cmdEdit_Device.Enabled = c And blnUser_Edit
    cmdDelete_Device.Enabled = d And blnUser_Delete
    cmdCancel_Device.Enabled = e
    
    'CmdPrint.Enabled = f
    'cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
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
    With rsDownload
        DTPicker_download_time.Value = .Fields("s_time").Value
        cbo_auto_log_enable.ListIndex = .Fields("s_enable").Value
    End With
End Sub

Public Sub set_edit_data_device()
    With rsDevice
        txt_ip_address = .Fields("ip_address").Value
        txt_port = .Fields("port_number").Value
        cboType.ListIndex = IIf(IsNull(.Fields("flag_type").Value), 0, .Fields("flag_type").Value)
        txt_name = .Fields("name").Value
        
        vIPAddress = .Fields("ip_address").Value
    End With
End Sub

Private Sub set_new_data_device()
    DTPicker_download_time.Value = Now
    cbo_auto_log_enable.ListIndex = 1
End Sub

Private Sub set_data_mode()
    
    If SSTab1.Tab = 1 Then
        If int_mode = 1 Then        'NEW
            Call clear_view_data
            fraEntry_Device.Visible = True
            txt_ip_address.Enabled = True
            TDBGrid2.Enabled = False
    '        Call set_new_data
            
            If txt_ip_address.Enabled = True Then
                txt_ip_address.SetFocus
            End If
        ElseIf int_mode = 0 Then    'VIEW
             Call clear_view_data
            fraEntry_Device.Visible = False
            TDBGrid2.Enabled = True
        
        ElseIf int_mode = 2 Then    'EDIT
            Call set_edit_data_device
            'txt_ip_address.Enabled = False
            fraEntry_Device.Visible = True
            TDBGrid1.Enabled = False
        End If
    ElseIf SSTab1.Tab = 3 Then
        If int_mode = 1 Then        'NEW
            Call clear_view_data
            fra_entry.Visible = True
            DTPicker_download_time.Enabled = True
            TDBGrid1.Enabled = False
            Call set_new_data_device
            
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
    End If
End Sub

Private Function check_validate_exist_new_download() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_validate_exist_new_download = False
    
    str_sql = "select count(s_time) as rec_count from s_auto_log where left(right(s_time,8),5) = '" _
    & Format(DTPicker_download_time.Value, "HH:mm") & "'"
    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new_download = True
        Exit Function
    End If
End Function

Private Function check_validate_exist_new_device() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_validate_exist_new_device = False
    
    str_sql = "select count(ip_address) as rec_count from m_device where ip_address = '" _
    & Trim(txt_ip_address) & "'"
    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new_device = True
        Exit Function
    End If
End Function

Private Sub check_invalid()
    MsgBox "Data Sudah Ada...", vbCritical, headerMSG
    If SSTab1.Tab = 1 Then
        txt_ip_address.Text = ""
        txt_port.Text = ""
        txt_name.Text = ""
        txt_ip_address.SetFocus
    ElseIf SSTab1.Tab = 3 Then
        DTPicker_download_time.Value = Now
        If DTPicker_download_time.Enabled = True Then DTPicker_download_time.SetFocus
    End If
End Sub

Private Function check_validate_exist_edit_download() As Boolean
    check_validate_exist_edit_download = False
    
    If (Not Format(DTPicker_download_time.Value, "HH:mm") = Format(rsDownload.Fields("s_time").Value, "hh:ii")) _
    And check_validate_exist_new_download Then
        check_validate_exist_edit_download = True
        Exit Function
    End If
End Function

Private Function check_validate_exist_edit_device() As Boolean
    check_validate_exist_edit_device = False
    
    If Not (txt_ip_address = rsDevice.Fields("ip_address").Value) And _
    check_validate_exist_new_device Then
        check_validate_exist_edit_device = True
        Exit Function
    End If
End Function

Private Sub insert_new_data_download()
On Error GoTo Err
    CnG.BeginTrans

    SQL = "INSERT INTO s_auto_log(s_time,s_enable) " & _
            "VALUES('" & Format(DTPicker_download_time.Value, "yyyy-MM-dd HH:mm:ss") & "', " & _
            "'" & cbo_auto_log_enable.ListIndex & "')"
    CnG.Execute SQL
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub insert_new_data_device()
On Error GoTo Err
    CnG.BeginTrans

    SQL = "INSERT INTO m_device(ip_address,port_number,name,flag_type) " & _
            "VALUES('" & Trim(txt_ip_address) & "', " & _
            "'" & Val(DropAllComma(txt_port)) & "', " & _
            "'" & Trim(txt_name) & "', " & _
            "'" & cboType.ListIndex & "')"
    CnG.Execute SQL
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub edit_old_data_download()
On Error GoTo Err

    CnG.BeginTrans
    SQL = "UPDATE s_auto_log SET s_time = '" & Format(DTPicker_download_time.Value, "yyyy-MM-dd HH:mm:ss") & "', " & _
            "s_enable = '" & cbo_auto_log_enable.ListIndex & "' " & _
          "WHERE s_number = " & rsDownload.Fields("s_number").Value & ""
    CnG.Execute SQL
    CnG.CommitTrans
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub edit_old_data_device()
On Error GoTo Err

    CnG.BeginTrans
    SQL = "UPDATE m_device SET ip_address = '" & txt_ip_address.Text & "', name = '" & Trim(txt_name) & "', " & _
            "port_number = '" & Val(DropAllComma(txt_port)) & "', " & _
            "flag_type = '" & cboType.ListIndex & "' " & _
          "WHERE ip_address = '" & vIPAddress & "'"
    CnG.Execute SQL
    CnG.CommitTrans
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub cmd_update_device_type_Click()
    Call update_data_device_type
End Sub

'Private Sub cmd_update_mail_Click()
'    Call update_data_mail
'End Sub

Private Sub cmdCancel_Click()
    int_mode = 0
    Call load_mode
End Sub

Private Sub CmdCancel_Device_Click()
    int_mode = 0
    Call load_mode
End Sub

Private Sub delete_data()
'On Error GoTo Err
Dim i As Integer

    If SSTab1.Tab = 1 Then
        If Not (TDBGrid2.ApproxCount > 0 And TDBGrid2.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid2.Columns("ip_address").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_device where ip_address = '" _
            & TDBGrid2.Columns("ip_address").Value & "'"
        CnG.CommitTrans
        
        Call load_data_device
        int_mode = 0
        Call load_mode
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & Format(TDBGrid1.Columns("s_time").Value, "hh:nn") & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from s_auto_log where s_number = " _
            & TDBGrid1.Columns("s_number").Value
        CnG.CommitTrans
        
        Call load_data_auto_log
        int_mode = 0
        Call load_mode
    End If
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub new_data()
    int_mode = 1
    Call load_mode
End Sub

Private Sub simpan_data()
    If SSTab1.Tab = 1 Then
        If int_mode = 1 Then
'            If Not check_validate_exist_edit_device Then Exit Sub
            If check_validate_exist_new_device Then
                Call check_invalid: Exit Sub
            End If
            Call insert_new_data_device
        ElseIf int_mode = 2 Then
'            If Not check_validate_exist_edit_device Then Exit Sub
            If check_validate_exist_edit_device Then
                Call check_invalid: Exit Sub
            End If
            Call edit_old_data_device
        End If
        
        Call load_data_device
        int_mode = 0
        Call load_mode
    ElseIf SSTab1.Tab = 3 Then
        If int_mode = 1 Then
'            If Not check_validate_exist_new_download Then Exit Sub
            If check_validate_exist_new_download Then
                Call check_invalid: Exit Sub
            End If
            Call insert_new_data_download
        ElseIf int_mode = 2 Then
'            If Not check_validate_exist_new_download Then Exit Sub
            If check_validate_exist_edit_download Then
                Call check_invalid: Exit Sub
            End If
            Call edit_old_data_download
        End If
        
        Call load_data_auto_log
        int_mode = 0
        Call load_mode
    End If
End Sub

Private Sub cmdNew_Click()
    Call new_data
End Sub

Private Sub CmdNew_Device_Click()
    Call new_data
End Sub

Private Sub cmdSave_Click()
    Call simpan_data
End Sub

Private Sub cmdSave_Device_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Click()
    Call edit_data
End Sub

Private Sub cmdEdit_Device_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Click()
    Call delete_data
End Sub

Private Sub cmdDelete_Device_Click()
    Call delete_data
End Sub

'Private Sub Command1_Click()
'    Call update_data_superadmin
'End Sub

Private Sub Form_Load()
    Call load_data_auto_log
    Call load_data_download_mode
    Call set_download_mode
    Call load_data_device
    Call load_data_device_type
    
    DTPicker_device_time.Value = Now
    vMachineNumber = 1
    cn_fp = False
    
    Call load_data_user_access(Me)
    SSTab1.Tab = 0
    Call SSTab1_Click(0)
    'Timer1.Enabled = True
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
On Error GoTo Err
Dim rs1 As New ADODB.Recordset
Dim i As Integer

    i = IIf(opt_base_ip, 1, IIf(opt_base_sensor, 2, 3))
    
    CnG.BeginTrans
    
'    CnG.Execute "delete from s_device where s_number=1"
    SQL = "UPDATE s_device SET s_device = " & i & ", " & _
            "description = 'setting device type' " & _
          "WHERE s_number = 1"
    CnG.Execute SQL
    
    MsgBox "Penyimpanan Berhasil...", vbInformation, headerMSG
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

'Private Sub update_data_superadmin()
'Dim rs1 As New ADODB.Recordset
'Dim strsql As String
'
'    CnG.BeginTrans
'
'    CnG.Execute "delete from m_user where user_level = 100"
'
'    strsql = "insert into m_user (user_code, user_name, user_pass, user_pass_key, user_level) " _
'            & "values('adminSSD', '" & Replace$(Trim$(txt_username), Chr(39), Chr(96)) _
'            & "','" & RC4DeCryptASC(Replace(Trim(txt_password.Text), Chr(39), Chr(96)), pEncryptionPassword) _
'            & "','" & pEncryptionPassword & "',100)"
'    CnG.Execute strsql
'
'    strsql = "Insert into t_user_access_level (level_code,access_level_code,level_name,allow_access) " _
'    & "Select 'adminSSD', code, name, 1 " _
'    & "from m_akses_level_group where code not in " _
'    & "(select access_level_code from t_user_access_level where level_code = 'adminSSD')"
'    CnG.Execute strsql
'
'    CnG.CommitTrans
'
'    MsgBox "Update Successfully", vbInformation, headerMSG
'End Sub

'Private Sub update_data_mail()
'Dim rs1 As New ADODB.Recordset
'Dim strsql As String
'
'    If txt_sender_email.Text = "" Then
'        MsgBox "Please Fill Sender Mail!", vbExclamation, headerMSG
'        Exit Sub
'    End If
'
'    If txt_sender_name.Text = "" Then
'        MsgBox "Please Fill Sender Name!", vbExclamation, headerMSG
'        Exit Sub
'    End If
'
'    If txt_mail_username.Text = "" Then
'        MsgBox "Please Fill Username!", vbExclamation, headerMSG
'        Exit Sub
'    End If
'
'    If txt_sender_pwd.Text = "" Then
'        MsgBox "Please Fill Sender Password!", vbExclamation, headerMSG
'        Exit Sub
'    End If
'
'    If txt_smtp.Text = "" Then
'        MsgBox "Please Fill SMTP!", vbExclamation, headerMSG
'        Exit Sub
'    End If
'
'    If txt_smtp_port.Text = "" Then
'        MsgBox "Please Fill SMTP Port!", vbExclamation, headerMSG
'        Exit Sub
'    End If
'
'    CnG.BeginTrans
'
'    strsql = "UPDATE s_mail SET " _
'            & "s_sender_email = '" & txt_sender_email.Text & "'," _
'            & "s_sender_name = '" & txt_sender_name.Text & "'," _
'            & "username = '" & txt_mail_username.Text & "'," _
'            & "password = '" & Replace(RC4DeCryptASC(Trim(txt_sender_pwd.Text), pEncryptionPassword), "'", "") & "'," _
'            & "smtp = '" & txt_smtp.Text & "'," _
'            & "port = '" & txt_smtp_port.Text & "' " _
'            & "WHERE s_number = 1"
'    CnG.Execute (strsql)
'
'    CnG.CommitTrans
'
'    MsgBox "Save Succesfully!", vbInformation, headerMSG
'End Sub

'Private Sub load_data_superadmin()
'Dim rs1 As New ADODB.Recordset
'
'    rs1.Open "select * from m_user where user_level = 100", CnG, adOpenStatic, adLockReadOnly
'    If rs1.RecordCount > 0 Then
'        txt_username = rs1.Fields("user_name").Value
'        txt_password = RC4DeCryptASC(Replace(Trim(rs1.Fields("user_pass").Value), Chr(39), Chr(96)), pEncryptionPassword)
'    Else
'        txt_username = ""
'        txt_password = ""
'    End If
'End Sub

'Private Sub load_data_mail()
'Dim rs1 As New ADODB.Recordset
'Dim strsql As String
'
'    strsql = "select * from s_mail where s_number = 1"
'    rs1.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
'
'    'rs1.Open "select * from s_mail where s_number = 1", CnG, adOpenStatic, adLockReadOnly
'    If rs1.RecordCount > 0 Then
'        txt_sender_email.Text = rs1!s_sender_email
'        txt_sender_name.Text = rs1!s_sender_name
'        txt_smtp.Text = rs1!smtp
'        txt_smtp_port.Text = rs1!PORT
'        txt_mail_username.Text = rs1!Username
'        txt_sender_pwd.Text = RC4DeCryptASC(rs1!Password, pEncryptionPassword)
'    Else
'        txt_smtp.Text = ""
'        txt_sender_email.Text = ""
'        txt_sender_name.Text = ""
'        txt_sender_pwd.Text = ""
'        txt_smtp_port.Text = ""
'        txt_mail_username.Text = ""
'    End If
'End Sub

Private Sub load_data_auto_log()
    If rsDownload.State = 1 Then rsDownload.Close
    SQL = "select * from s_auto_log order by right(s_time,8) asc"
    rsDownload.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsDownload
End Sub

Private Sub load_data_download_mode()
    opt_auto_download.Value = IIf(BLN_AUTO_LOG, True, False)
End Sub

Private Sub set_download_mode()
    If opt_auto_download Then
        fra_auto_log.Visible = True
        Frame4.Visible = True
        Call set_buttons_enable(True, False, True, True, False, True, True)
    ElseIf opt_manual_download Then
        fra_auto_log.Visible = False
        Frame4.Visible = False
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

Private Sub Form_Unload(Cancel As Integer)
    Set frm_stg_all = Nothing
End Sub

Private Sub opt_auto_download_Click()
    fra_auto_log.Visible = True
    Frame4.Visible = True
    
    Call set_download_mode
    Call update_download_mode
End Sub

Private Sub opt_manual_download_Click()
    fra_auto_log.Visible = False
    Frame4.Visible = False
    
    Call set_download_mode
    Call update_download_mode
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 1 Or SSTab1.Tab = 3 Then
        int_mode = 0
        Call load_mode
    ElseIf SSTab1.Tab = 2 Then
        Call load_data_device_tdb
    ElseIf SSTab1.Tab = 0 Then
        Call set_buttons_enable(False, True, False, False, False, False, False)
    End If
End Sub

Private Sub TDBCombo_device_ItemChange()
    If TDBCombo_device.ApproxCount > 0 Then
        TDBCombo_device.Text = TDBCombo_device.Columns("ip_address").Value
    End If
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If TDBGrid1.Columns(ColIndex).Caption = "DOWNLOAD TIME" Then Value = Format(Value, "hh:nn")
End Sub

Private Sub load_data_device()
    If rsDevice.State = 1 Then rsDevice.Close
    SQL = "select *,case when flag_type = 0 then 'TIPE A' else 'TIPE B' end type from m_device order by ip_address"
    rsDevice.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid2.DataSource = rsDevice
End Sub

Private Sub load_data_device_tdb()
    If rsDevice_Tdb.State = 1 Then rsDevice_Tdb.Close
    SQL = "select * from m_device order by ip_address"
    rsDevice_Tdb.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_device.RowSource = rsDevice_Tdb
End Sub

Private Sub cmd_test_Click()
Call test_device_connection
End Sub

Private Sub test_device_connection()
On Error GoTo Err

    If Not connect Then
        MsgBox "Error connecting to device...", vbCritical, headerMSG
        Exit Sub
    End If
    
    MsgBox "Connecting to device successfully tested!", vbInformation, headerMSG
    
    Call disconnect
    
    Exit Sub

Err:
MsgBox "Error Connecting to Device!", vbCritical, headerMSG
Call disconnect
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsDevice.Filter = adFilterNone
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

Private Sub TDBGrid2_FilterChange()
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid2.Columns
    i = TDBGrid2.Col
    TDBGrid2.HoldFields
    
    rsDevice.Filter = getFilter()
    TDBGrid2.Col = i
    TDBGrid2.EditActive = True
    
    TDBGrid2.SelStart = Len(TDBGrid2.Columns(i).FilterText)
    If TDBGrid2.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid2.Col = i
    End If

    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

