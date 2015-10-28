VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_pph21 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER PPh PASAL 21"
   ClientHeight    =   7125
   ClientLeft      =   -15
   ClientTop       =   240
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
   Icon            =   "frm_mst_pph21.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_entry1 
      Height          =   1815
      Left            =   240
      TabIndex        =   27
      Top             =   3660
      Width           =   11295
      Begin VB.TextBox txt_pph_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         MaxLength       =   10
         TabIndex        =   29
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
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
         TabIndex        =   28
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_pph_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   31
         Top             =   960
         Width           =   3495
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "NAMA PPH"
         Height          =   195
         Left            =   3480
         TabIndex        =   32
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "KODE PPH"
         Height          =   195
         Left            =   3480
         TabIndex        =   30
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.TextBox txt_pph21_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3030
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   24
      Top             =   780
      Width           =   3855
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   5580
      Width           =   11295
      Begin VB.CommandButton cmdDelete_All 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   8880
         Picture         =   "frm_mst_pph21.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew_Master 
         Caption         =   "&New"
         Height          =   645
         Left            =   7800
         Picture         =   "frm_mst_pph21.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton cmd_refresh 
         Caption         =   "&Refresh"
         Height          =   645
         Left            =   1800
         Picture         =   "frm_mst_pph21.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1020
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save Detail"
         Height          =   645
         Left            =   1800
         Picture         =   "frm_mst_pph21.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5040
         Picture         =   "frm_mst_pph21.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   10080
         Picture         =   "frm_mst_pph21.frx":1BBE
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New Detail"
         Height          =   645
         Left            =   720
         Picture         =   "frm_mst_pph21.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   0
         Picture         =   "frm_mst_pph21.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Detail"
         Height          =   645
         Left            =   3960
         Picture         =   "frm_mst_pph21.frx":2C5C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit Detail"
         Height          =   645
         Left            =   2880
         Picture         =   "frm_mst_pph21.frx":31E6
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_pph 
      Height          =   375
      Left            =   1230
      OleObjectBlob   =   "frm_mst_pph21.frx":3770
      TabIndex        =   25
      Top             =   780
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_pph 
      Height          =   375
      Left            =   1350
      Top             =   900
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
   Begin VB.Frame fra_entry 
      Height          =   2775
      Left            =   240
      TabIndex        =   14
      Top             =   2700
      Width           =   11295
      Begin VB.CheckBox chk_flag_over 
         Height          =   255
         Left            =   6120
         TabIndex        =   4
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txt_pph21_percentage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txt_pph21_upper 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   3
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txt_pph21_under 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   2
         Top             =   480
         Width           =   1695
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
         TabIndex        =   17
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_pph21_number 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "PROSENTASE"
         Height          =   195
         Left            =   4800
         TabIndex        =   23
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "UP"
         Height          =   195
         Left            =   4800
         TabIndex        =   22
         Top             =   1200
         Width           =   195
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "HINGGA"
         Height          =   195
         Left            =   4800
         TabIndex        =   21
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DESKRIPSI"
         Height          =   195
         Left            =   4800
         TabIndex        =   20
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NO."
         Height          =   195
         Left            =   600
         TabIndex        =   19
         Top             =   480
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "DARI"
         Height          =   195
         Left            =   4800
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4335
      Left            =   240
      TabIndex        =   0
      Top             =   1140
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NUMBER"
      Columns(0).DataField=   "pph21_number"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "FROM"
      Columns(1).DataField=   "pph21_under"
      Columns(1).NumberFormat=   "Standard"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TO"
      Columns(2).DataField=   "pph21_upper"
      Columns(2).NumberFormat=   "Standard"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   4
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "UP"
      Columns(3).DataField=   "flag_over"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "PERCENTAGE"
      Columns(4).DataField=   "pph21_percentage"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DESCRIPTION"
      Columns(5).DataField=   "description"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2381"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3784"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3704"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=514"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3916"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3836"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1429"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1349"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2355"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2275"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=4948"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=4868"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=512"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
      Caption         =   "LIST OF PPh PASAL 21"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=1"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=0"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(58)  =   "Named:id=33:Normal"
      _StyleDefs(59)  =   ":id=33,.parent=0"
      _StyleDefs(60)  =   "Named:id=34:Heading"
      _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   ":id=34,.wraptext=-1"
      _StyleDefs(63)  =   "Named:id=35:Footing"
      _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(65)  =   "Named:id=36:Selected"
      _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(67)  =   "Named:id=37:Caption"
      _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(69)  =   "Named:id=38:HighlightRow"
      _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(71)  =   "Named:id=39:EvenRow"
      _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(73)  =   "Named:id=40:OddRow"
      _StyleDefs(74)  =   ":id=40,.parent=33"
      _StyleDefs(75)  =   "Named:id=41:RecordSelector"
      _StyleDefs(76)  =   ":id=41,.parent=34"
      _StyleDefs(77)  =   "Named:id=42:FilterBar"
      _StyleDefs(78)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   0
      Top             =   1620
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
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "MASTER PPH PASAL 21"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5100
      TabIndex        =   35
      Top             =   0
      Width           =   3615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "PPH21 TYPE"
      Height          =   195
      Left            =   240
      TabIndex        =   26
      Top             =   810
      Width           =   870
   End
End
Attribute VB_Name = "frm_mst_pph21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim v_value_under, v_value_upper, v_value_percent As Double
Dim strsql As String

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(pph21_number) as rec_count from m_pph21_detail where pph21_number = " _
& "'" & Int(txt_pph21_number) & "' AND pph21_code = '" & TDBCombo_pph.Text & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
txt_pph21_number = ""
If txt_pph21_number.Enabled = True Then txt_pph21_number.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not txt_pph21_number = Adodc1.Recordset.Fields("pph21_number").Value And _
check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

If Trim(txt_pph21_number) = "" Then
    MsgBox "Number is empty!", vbOKOnly + vbInformation, headerMSG
    txt_pph21_number.SetFocus
    check_validate_new = False
    Exit Function
End If

If Trim(txt_pph21_under) = "" Then
    MsgBox "Start from value is empty!", vbOKOnly + vbInformation, headerMSG
    txt_pph21_under.SetFocus
    check_validate_new = False
    Exit Function
End If

'If Trim(txt_pph21_upper) = "" Then
'    MsgBox "End to value is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_pph21_upper.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
End Function

Private Sub cmd_refresh_Click()
Call load_data
End Sub

Private Sub CmdCancel_Click()
int_mode = 0
Call load_mode
CmdNew_Master.Caption = "&New"
End Sub

Private Sub cmdDelete_All_Click()
Dim i As Integer

i = MsgBox("Are you sure want to delete data '" _
    & txt_pph21_name.Text & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_pph21_detail where " _
& "pph21_code = '" & TDBCombo_pph.Text & "'"
CnG.Execute "delete from m_pph21 where pph21_code = " _
    & "'" & TDBCombo_pph.Text & "'"

'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
strsql = "Update temp_sal_proses set salary_proses = 0"
CnG.Execute strsql
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
CnG.CommitTrans

Call load_data_pph21
Call load_data
int_mode = 0
Call load_mode

TDBCombo_pph.Text = ""
txt_pph21_name = ""
Set TDBGrid1.DataSource = Nothing

End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "No Data selected!", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Are you sure want to delete data '" _
    & TDBGrid1.Columns("pph21_number").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_pph21_detail where pph21_number = " _
    & "'" & TDBGrid1.Columns("pph21_number").Value & "' AND pph21_code = '" & TDBCombo_pph.Text & "'"

'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
strsql = "Update temp_sal_proses set salary_proses = 0"
CnG.Execute strsql
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
CnG.CommitTrans

Call load_data
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    txt_pph21_number = .Fields("pph21_number").Value
    txt_pph21_under = FormatNumber(.Fields("pph21_under").Value)
    txt_pph21_upper = FormatNumber(.Fields("pph21_upper").Value)
    chk_flag_over.Value = .Fields("flag_over").Value
    txt_pph21_percentage = FormatNumber(.Fields("pph21_percentage").Value)
    txt_description = .Fields("description").Value
End With

    v_value_under = txt_pph21_under
    v_value_upper = txt_pph21_upper
    v_value_percent = txt_pph21_percentage
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_pph21_detail where pph21_number = " _
& "'" & Adodc1.Recordset.Fields("pph21_number").Value & "' AND " _
& "pph21_code = '" & TDBCombo_pph.Text & "'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_pph21_detail where pph21_code = 'άφ' AND pph21_number = 'άφ'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub CmdNew_Master_Click()
Dim strsql As String

If CmdNew_Master.Caption = "&New" Then
    CmdNew_Master.Caption = "&Save"
    Call set_buttons_enable(False, False, False, False, True, False, False)
    fra_entry1.Visible = True
    fra_entry.Visible = False
    
    txt_pph_code.Text = ""
    txt_pph_name.Text = ""
    txt_pph_code.SetFocus
    
Else
    strsql = "INSERT INTO m_pph21(pph21_code,pph21_name) " _
            & "VALUES ('" & txt_pph_code & "','" & txt_pph_name & "')"
    CnG.Execute strsql
    
    Call set_buttons_enable(True, False, True, True, False, True, True)
    CmdNew_Master.Caption = "&New"
        
    fra_entry1.Visible = False
    
    txt_pph_code.Text = ""
    txt_pph_name.Text = ""
'    txt_pph_code.SetFocus

    Call load_data_pph21
'    Call load_data
End If
End Sub

Private Sub CmdPrint_Click()
TDBGrid1.PrintInfo.PageSetup
If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
End If
End Sub

Private Sub insert_new_data()
CnG.BeginTrans

'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
strsql = "Update temp_sal_proses set salary_proses = 0"
CnG.Execute strsql
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

With rsBound
    .AddNew
    
    .Fields("pph21_code").Value = Trim(TDBCombo_pph.Text)
    .Fields("pph21_number").Value = Trim(txt_pph21_number)
    '-----------------------------------------------------------------------------
    .Fields("pph21_under").Value = Val(DropAllComma(txt_pph21_under))
    .Fields("pph21_upper").Value = Val(DropAllComma(txt_pph21_upper))
    .Fields("flag_over").Value = IIf(chk_flag_over, 1, 0)
    .Fields("pph21_percentage").Value = Val(DropAllComma(txt_pph21_percentage))
    .Fields("description").Value = Trim(txt_description)
    
    .Update
End With

CnG.CommitTrans
End Sub

Private Sub edit_old_data()
On Error GoTo err_capture

CnG.BeginTrans
'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
If v_value_under <> txt_pph21_under.Text Or v_value_upper <> txt_pph21_upper.Text Or v_value_percent <> txt_pph21_percentage.Text Then
    strsql = "Update temp_sal_proses set salary_proses = 0"
    CnG.Execute strsql
End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

With rsBound

    .Fields("pph21_code").Value = Trim(TDBCombo_pph)
    .Fields("pph21_number").Value = Trim(txt_pph21_number)
    '-----------------------------------------------------------------------------
    .Fields("pph21_under").Value = Val(DropAllComma(txt_pph21_under))
    .Fields("pph21_upper").Value = Val(DropAllComma(txt_pph21_upper))
    .Fields("flag_over").Value = IIf(chk_flag_over, 1, 0)
    .Fields("pph21_percentage").Value = Val(DropAllComma(txt_pph21_percentage))
    .Fields("description").Value = Trim(txt_description)
    
    .Update
End With
CnG.CommitTrans

Exit Sub
err_capture:
rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
End Sub

Private Sub CmdSave_Click()
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
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
CmdNew.Enabled = a And blnUser_Add
CmdSave.Enabled = b
cmdEdit.Enabled = c And blnUser_Edit
cmdDelete.Enabled = d And blnUser_Delete
CmdCancel.Enabled = e

CmdPrint.Enabled = F
cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data()
Dim Ctr As Control
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        If Not LCase(Ctr.name) = "txt_pph21_name" Then Ctr.Text = ""
'    ElseIf TypeOf Ctr Is TDBCombo Then
'        If Not LCase(Ctr.name) = "TDBCombo_pph" Then Ctr.Text = ""
    ElseIf TypeOf Ctr Is DTPicker Then
        Ctr.Value = Now
    End If
Next
End Sub

Private Sub set_enabled_control(ByVal i As Boolean)
Dim Ctr As Control
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
If int_mode = 1 Then        'NEW Rule
    Call clear_view_data
    fra_entry.Visible = True
    txt_pph21_number.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_pph21_number.Enabled = True Then
        txt_pph21_number.SetFocus
    End If

'ElseIf int_mode = 3 Then        'NEW Master
'    Call clear_view_data
'    fra_entry1.Visible = True
'    txt_pph_code.Enabled = True
''    Call set_new_data
'
'    If txt_pph_code.Enabled = True Then
'        txt_pph21_code.SetFocus
'    End If
    
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    fra_entry1.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    txt_pph21_number.Enabled = False
    fra_entry.Visible = True
    TDBGrid1.Enabled = False
End If
End Sub

Private Sub load_mode()
If int_mode = 1 Then        ' new Rule
    Call set_buttons_enable(False, True, False, False, True, False, False)
'ElseIf int_mode = 3 Then        ' new Master
'    Call set_buttons_enable(False, True, False, False, True, False, False)
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
Adodc_pph.ConnectionString = strConn

Call load_data_pph21
    
Call load_data_user_access(Me)
int_mode = 0
Call load_mode

'timer1.Enabled = True
End Sub

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

Private Sub load_data()
Adodc1.RecordSource = "select * from m_pph21_detail where pph21_code = '" _
& TDBCombo_pph.Columns("pph21_code").Value & "' order by pph21_code,pph21_number"
Adodc1.Refresh

'cmdEdit.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'cmdDelete.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub load_data_pph21()
Adodc_pph.RecordSource = "select * from m_pph21 order by pph21_code"
Adodc_pph.Refresh

TDBCombo_pph.RowSource = Adodc_pph
End Sub

Private Sub TDBCombo_pph_ItemChange()
If TDBCombo_pph.ApproxCount > 0 Then
    TDBCombo_pph.Text = TDBCombo_pph.Columns("pph21_code").Value
    txt_pph21_name = TDBCombo_pph.Columns("pph21_name").Value
    
    Call load_data
End If
End Sub

Private Sub txt_pph_code_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Timer1_Timer()
'timer1.Enabled = False
'Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub


