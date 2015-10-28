VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_status 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WORKING STATUS"
   ClientHeight    =   6690
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   11760
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mst_status.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_entry 
      Height          =   2775
      Left            =   210
      TabIndex        =   4
      Top             =   2280
      Width           =   11295
      Begin VB.ComboBox cboPeriode 
         Height          =   315
         ItemData        =   "frm_mst_status.frx":058A
         Left            =   5280
         List            =   "frm_mst_status.frx":0594
         TabIndex        =   16
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txt_remark 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1860
         Width           =   5145
      End
      Begin VB.TextBox txt_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1020
         Width           =   5145
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
         TabIndex        =   6
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5280
         MaxLength       =   10
         TabIndex        =   1
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "PERIOD STATUS"
         Height          =   195
         Left            =   3600
         TabIndex        =   17
         Top             =   1470
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   1890
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CODE*"
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NAME*"
         Height          =   195
         Left            =   3600
         TabIndex        =   7
         Top             =   1050
         Width           =   510
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   11295
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin prj_tpc.vbButton cmdNew 
         Height          =   705
         Left            =   540
         TabIndex        =   10
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&New"
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
         MICON           =   "frm_mst_status.frx":05AC
         PICN            =   "frm_mst_status.frx":05C8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdSave 
         Height          =   705
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Save"
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
         MICON           =   "frm_mst_status.frx":165A
         PICN            =   "frm_mst_status.frx":1676
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdEdit 
         Height          =   705
         Left            =   2580
         TabIndex        =   12
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Edit"
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
         MICON           =   "frm_mst_status.frx":2708
         PICN            =   "frm_mst_status.frx":2724
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdDelete 
         Height          =   705
         Left            =   3600
         TabIndex        =   13
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Delete"
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
         MICON           =   "frm_mst_status.frx":37B6
         PICN            =   "frm_mst_status.frx":37D2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdCancel 
         Height          =   705
         Left            =   4620
         TabIndex        =   14
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Cancel"
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
         MICON           =   "frm_mst_status.frx":4864
         PICN            =   "frm_mst_status.frx":4880
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdExit 
         Height          =   705
         Left            =   10080
         TabIndex        =   15
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Exit"
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
         MICON           =   "frm_mst_status.frx":5912
         PICN            =   "frm_mst_status.frx":592E
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7435
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "CODE"
      Columns(0).DataField=   "status_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NAME"
      Columns(1).DataField=   "status_name"
      Columns(1).NumberFormat=   "FormatText Event"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "PERIOD STATUS"
      Columns(2).DataField=   "period_desc"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DESCRIPTION"
      Columns(3).DataField=   "description"
      Columns(3).NumberFormat=   "FormatText Event"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=6218"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6138"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2646"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2566"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=7276"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=7197"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
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
      Caption         =   "LIST OF WORKING STATUS"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.alignment=0,.bgcolor=&H80000002&"
      _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000009&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(11)  =   ":id=4,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=4,.fontname=Tahoma"
      _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(18)  =   ":id=3,.fontname=Tahoma"
      _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&H80000002&,.fgcolor=&H80000009&"
      _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(30)  =   ":id=14,.fgcolor=&H80000002&"
      _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(34)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(53)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER WORKING STATUS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   18
      Top             =   150
      Width           =   4365
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_status.frx":69C0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11850
   End
End
Attribute VB_Name = "frm_mst_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsWorkingStatus As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
    check_validate_exist_new = False
    
    SQL = "select count(status_code) as rec_count from m_emp_status where status_code = '" & Trim(txt_code) & "'"
    rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        Exit Function
    End If
End Function

Private Sub check_invalid()
    MsgBox "Data found!", vbCritical, headerMSG
    txt_code = ""
    If txt_code.Enabled = True Then txt_code.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If Not txt_code.Text = rsWorkingStatus.Fields("status_code").Value And _
    check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    If Trim(txt_code.Text) = "" Then
        MsgBox "Code is empty!", vbOKOnly + vbInformation, headerMSG
        txt_code.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    If Trim(txt_name) = "" Then
        MsgBox "Name is empty!", vbOKOnly + vbInformation, headerMSG
        txt_name.SetFocus
        check_validate_new = False
        Exit Function
    End If
End Function

Private Sub cmdCancel_Click()
    int_mode = 0
    Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

On Error GoTo Err
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        Exit Sub
    End If
    
    i = MsgBox("Are you sure want to delete data '" _
        & TDBGrid1.Columns("status_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    CnG.BeginTrans
    CnG.Execute "delete from m_emp_status where status_code = '" _
        & TDBGrid1.Columns("status_code").Value & "'"
    CnG.CommitTrans
    
    Call load_data
    int_mode = 0
    Call load_mode
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
    vSetData = 1
    
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    With rsWorkingStatus
        txt_code.Text = .Fields("status_code").Value
        txt_name.Text = .Fields("status_name").Value
        cboPeriode.ListIndex = .Fields("flag_period").Value
        txt_remark.Text = .Fields("description").Value
    End With
End Sub

Private Sub cmdEdit_Click()
    int_mode = 2
    Call load_mode
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNew_Click()
    int_mode = 1
    Call load_mode
    cboPeriode.ListIndex = 0
End Sub

Private Sub insert_new_data()
    On Error GoTo Err
    CnG.BeginTrans
    
    SQL = "INSERT INTO m_emp_status(status_code,status_name,flag_period,description) " & _
            "VALUES " & _
            "('" & Trim(txt_code.Text) & "','" & Trim(txt_name.Text) & "'," & _
            "'" & cboPeriode.ListIndex & "','" & Trim(txt_remark.Text) & "')"
    CnG.Execute SQL
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err

    CnG.BeginTrans
    SQL = "UPDATE m_emp_status SET status_code = '" & txt_code.Text & "'," & _
            "status_name = '" & txt_name.Text & "'," & _
            "flag_period = '" & cboPeriode.ListIndex & "'," & _
            "description = '" & txt_remark.Text & "' " & _
          "WHERE status_code = '" & txt_code.Text & "'"
    CnG.Execute SQL
    CnG.CommitTrans
    
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdSave_Click()
    If int_mode = 1 Then
        If Not check_validate_new Then Exit Sub
        If check_validate_exist_new Then
            Call check_invalid: Exit Sub
        End If
        Call insert_new_data
    ElseIf int_mode = 2 Then
        If Not check_validate_new Then Exit Sub
        If check_validate_exist_edit Then
            Call check_invalid: Exit Sub
        End If
        Call edit_old_data
    End If
    
    Call load_data
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
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
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

Private Sub set_new_data()
    'cbo_jns_kelamin.ListIndex = 1
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        fra_entry.Visible = True
        txt_code.Enabled = True
        TDBGrid1.Enabled = False
        Call set_new_data
        
        If txt_code.Enabled = True Then
            txt_code.SetFocus
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        fra_entry.Visible = False
        TDBGrid1.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        txt_code.Enabled = False
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

Private Sub Form_Load()
    Call load_data
    oClause = ""
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    timer1.Enabled = True
    
    cboPeriode.ListIndex = 0
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsWorkingStatus.Filter = adFilterNone
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

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_status = Nothing
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo Err
    
    Dim i As Integer
    
    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields
    
    rsWorkingStatus.Filter = getFilter()
    TDBGrid1.Col = i
    TDBGrid1.EditActive = True
    
    TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
    If TDBGrid1.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid1.Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data()
    If rsWorkingStatus.State Then rsWorkingStatus.Close
    SQL = "select status_code,status_name," & _
            "case when flag_period = 0 then 'NON PERIOD' else 'PERIOD' end period_desc," & _
            "flag_period,description " & _
          "from m_emp_status " & oClause
    rsWorkingStatus.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsWorkingStatus
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If TDBGrid1.Columns(ColIndex).Caption = "ENTER FROM" Or _
    TDBGrid1.Columns(ColIndex).Caption = "ENTER TO" Then
        Value = Format(Value, "dd-mm")
    End If
End Sub

Private Sub Timer1_Timer()
'timer1.Enabled = False
'Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub txt_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid1.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid1.Columns(ColIndex).DataField
    Call load_data

End Sub
