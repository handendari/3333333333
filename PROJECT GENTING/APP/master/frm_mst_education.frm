VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_education 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER PENDIDIKAN"
   ClientHeight    =   9090
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   14640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mst_education.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fra_entry 
      Height          =   3375
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   14175
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   4
         Top             =   2280
         Width           =   4815
      End
      Begin VB.TextBox txt_address 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   4800
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1560
         Width           =   4815
      End
      Begin VB.ComboBox cbo_education_code 
         Height          =   315
         ItemData        =   "frm_mst_education.frx":000C
         Left            =   4800
         List            =   "frm_mst_education.frx":0031
         TabIndex        =   1
         Text            =   "..."
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txt_education_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4800
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "KETERANGAN"
         Height          =   195
         Left            =   3120
         TabIndex        =   17
         Top             =   2280
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ALAMAT"
         Height          =   195
         Left            =   3120
         TabIndex        =   16
         Top             =   1560
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NAMA PENDIDIKAN"
         Height          =   195
         Left            =   3120
         TabIndex        =   15
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "LEVEL PENDIDIKAN"
         Height          =   195
         Left            =   3120
         TabIndex        =   14
         Top             =   840
         Width           =   1380
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   13
      Top             =   7560
      Width           =   14175
      Begin VB.CommandButton cmd_select 
         Caption         =   "&Select"
         Height          =   645
         Left            =   8640
         Picture         =   "frm_mst_education.frx":0066
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2160
         Picture         =   "frm_mst_education.frx":05F0
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5400
         Picture         =   "frm_mst_education.frx":0B7A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   11520
         Picture         =   "frm_mst_education.frx":1104
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   1080
         Picture         =   "frm_mst_education.frx":168E
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4320
         Picture         =   "frm_mst_education.frx":1C18
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3240
         Picture         =   "frm_mst_education.frx":21A2
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   7215
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   12726
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   16
      Columns(0)._MaxComboItems=   5
      Columns(0).ValueItems(0)._DefaultItem=   0
      Columns(0).ValueItems(0).Value=   "0"
      Columns(0).ValueItems(0).Value.vt=   8
      Columns(0).ValueItems(0).DisplayValue=   "OTHER"
      Columns(0).ValueItems(0).DisplayValue.vt=   8
      Columns(0).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(1)._DefaultItem=   0
      Columns(0).ValueItems(1).Value=   "1"
      Columns(0).ValueItems(1).Value.vt=   8
      Columns(0).ValueItems(1).DisplayValue=   "SD"
      Columns(0).ValueItems(1).DisplayValue.vt=   8
      Columns(0).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(2)._DefaultItem=   0
      Columns(0).ValueItems(2).Value=   "2"
      Columns(0).ValueItems(2).Value.vt=   8
      Columns(0).ValueItems(2).DisplayValue=   "SMP"
      Columns(0).ValueItems(2).DisplayValue.vt=   8
      Columns(0).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(3)._DefaultItem=   0
      Columns(0).ValueItems(3).Value=   "3"
      Columns(0).ValueItems(3).Value.vt=   8
      Columns(0).ValueItems(3).DisplayValue=   "SMA"
      Columns(0).ValueItems(3).DisplayValue.vt=   8
      Columns(0).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(4)._DefaultItem=   0
      Columns(0).ValueItems(4).Value=   "4"
      Columns(0).ValueItems(4).Value.vt=   8
      Columns(0).ValueItems(4).DisplayValue=   "D1"
      Columns(0).ValueItems(4).DisplayValue.vt=   8
      Columns(0).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(5)._DefaultItem=   0
      Columns(0).ValueItems(5).Value=   "5"
      Columns(0).ValueItems(5).Value.vt=   8
      Columns(0).ValueItems(5).DisplayValue=   "D2"
      Columns(0).ValueItems(5).DisplayValue.vt=   8
      Columns(0).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(6)._DefaultItem=   0
      Columns(0).ValueItems(6).Value=   "6"
      Columns(0).ValueItems(6).Value.vt=   8
      Columns(0).ValueItems(6).DisplayValue=   "D3"
      Columns(0).ValueItems(6).DisplayValue.vt=   8
      Columns(0).ValueItems(6)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(7)._DefaultItem=   0
      Columns(0).ValueItems(7).Value=   "7"
      Columns(0).ValueItems(7).Value.vt=   8
      Columns(0).ValueItems(7).DisplayValue=   "D4"
      Columns(0).ValueItems(7).DisplayValue.vt=   8
      Columns(0).ValueItems(7)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(8)._DefaultItem=   0
      Columns(0).ValueItems(8).Value=   "8"
      Columns(0).ValueItems(8).Value.vt=   8
      Columns(0).ValueItems(8).DisplayValue=   "S1"
      Columns(0).ValueItems(8).DisplayValue.vt=   8
      Columns(0).ValueItems(8)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(9)._DefaultItem=   0
      Columns(0).ValueItems(9).Value=   "9"
      Columns(0).ValueItems(9).Value.vt=   8
      Columns(0).ValueItems(9).DisplayValue=   "S2"
      Columns(0).ValueItems(9).DisplayValue.vt=   8
      Columns(0).ValueItems(9)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems(10)._DefaultItem=   0
      Columns(0).ValueItems(10).Value=   "10"
      Columns(0).ValueItems(10).Value.vt=   8
      Columns(0).ValueItems(10).DisplayValue=   "S3"
      Columns(0).ValueItems(10).DisplayValue.vt=   8
      Columns(0).ValueItems(10)._PropDict=   "_DefaultItem,517,2"
      Columns(0).ValueItems.Count=   11
      Columns(0).Caption=   "EDUCATION LEVEL"
      Columns(0).DataField=   "education_code_name"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EDUCATION NAME"
      Columns(1).DataField=   "education_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ADDRESS"
      Columns(2).DataField=   "address"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DESCRIPTION"
      Columns(3).DataField=   "description"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   1
      Splits(0).Size  =   3000.189
      Splits(0).Size.vt=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4128"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4048"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2328"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2249"
      Splits(0)._ColumnProps(9)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=6191"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6112"
      Splits(0)._ColumnProps(17)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(19)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(20)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=1429"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=1349"
      Splits(0)._ColumnProps(25)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(30)=   "Column(3)._MinWidth=55133872"
      Splits(1)._UserFlags=   0
      Splits(1).Size  =   2
      Splits(1).Size.vt=   2
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).ScrollBars=   1
      Splits(1).DividerColor=   13160660
      Splits(1).FilterBar=   -1  'True
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=4"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=3493"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=3413"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(0)._MinWidth=2883584"
      Splits(1)._ColumnProps(10)=   "Column(1).Width=6747"
      Splits(1)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(12)=   "Column(1)._WidthInPix=6668"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=513"
      Splits(1)._ColumnProps(14)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(15)=   "Column(1)._MinWidth=54620744"
      Splits(1)._ColumnProps(16)=   "Column(2).Width=8837"
      Splits(1)._ColumnProps(17)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(18)=   "Column(2)._WidthInPix=8758"
      Splits(1)._ColumnProps(19)=   "Column(2)._ColStyle=516"
      Splits(1)._ColumnProps(20)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(21)=   "Column(2)._MinWidth=1195525204"
      Splits(1)._ColumnProps(22)=   "Column(3).Width=3942"
      Splits(1)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(24)=   "Column(3)._WidthInPix=3863"
      Splits(1)._ColumnProps(25)=   "Column(3)._ColStyle=516"
      Splits(1)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(27)=   "Column(3)._MinWidth=55175648"
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
      Caption         =   "DAFTAR PENDIDIKAN"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Splits(1).Style:id=63,.parent=1"
      _StyleDefs(51)  =   "Splits(1).CaptionStyle:id=76,.parent=4,.bgcolor=&H80000002&,.fgcolor=&H80000009&"
      _StyleDefs(52)  =   "Splits(1).HeadingStyle:id=64,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(53)  =   ":id=64,.fgcolor=&H80000002&"
      _StyleDefs(54)  =   "Splits(1).FooterStyle:id=65,.parent=3"
      _StyleDefs(55)  =   "Splits(1).InactiveStyle:id=66,.parent=5"
      _StyleDefs(56)  =   "Splits(1).SelectedStyle:id=72,.parent=6"
      _StyleDefs(57)  =   "Splits(1).EditorStyle:id=71,.parent=7"
      _StyleDefs(58)  =   "Splits(1).HighlightRowStyle:id=73,.parent=8"
      _StyleDefs(59)  =   "Splits(1).EvenRowStyle:id=74,.parent=9"
      _StyleDefs(60)  =   "Splits(1).OddRowStyle:id=75,.parent=10"
      _StyleDefs(61)  =   "Splits(1).RecordSelectorStyle:id=77,.parent=11"
      _StyleDefs(62)  =   "Splits(1).FilterBarStyle:id=78,.parent=12"
      _StyleDefs(63)  =   "Splits(1).Columns(0).Style:id=86,.parent=63"
      _StyleDefs(64)  =   "Splits(1).Columns(0).HeadingStyle:id=83,.parent=64"
      _StyleDefs(65)  =   "Splits(1).Columns(0).FooterStyle:id=84,.parent=65"
      _StyleDefs(66)  =   "Splits(1).Columns(0).EditorStyle:id=85,.parent=71"
      _StyleDefs(67)  =   "Splits(1).Columns(1).Style:id=90,.parent=63,.alignment=2"
      _StyleDefs(68)  =   "Splits(1).Columns(1).HeadingStyle:id=87,.parent=64"
      _StyleDefs(69)  =   "Splits(1).Columns(1).FooterStyle:id=88,.parent=65"
      _StyleDefs(70)  =   "Splits(1).Columns(1).EditorStyle:id=89,.parent=71"
      _StyleDefs(71)  =   "Splits(1).Columns(2).Style:id=94,.parent=63"
      _StyleDefs(72)  =   "Splits(1).Columns(2).HeadingStyle:id=91,.parent=64"
      _StyleDefs(73)  =   "Splits(1).Columns(2).FooterStyle:id=92,.parent=65"
      _StyleDefs(74)  =   "Splits(1).Columns(2).EditorStyle:id=93,.parent=71"
      _StyleDefs(75)  =   "Splits(1).Columns(3).Style:id=98,.parent=63,.alignment=3"
      _StyleDefs(76)  =   "Splits(1).Columns(3).HeadingStyle:id=95,.parent=64"
      _StyleDefs(77)  =   "Splits(1).Columns(3).FooterStyle:id=96,.parent=65"
      _StyleDefs(78)  =   "Splits(1).Columns(3).EditorStyle:id=97,.parent=71"
      _StyleDefs(79)  =   "Named:id=33:Normal"
      _StyleDefs(80)  =   ":id=33,.parent=0"
      _StyleDefs(81)  =   "Named:id=34:Heading"
      _StyleDefs(82)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   ":id=34,.wraptext=-1"
      _StyleDefs(84)  =   "Named:id=35:Footing"
      _StyleDefs(85)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   "Named:id=36:Selected"
      _StyleDefs(87)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(88)  =   "Named:id=37:Caption"
      _StyleDefs(89)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(90)  =   "Named:id=38:HighlightRow"
      _StyleDefs(91)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(92)  =   "Named:id=39:EvenRow"
      _StyleDefs(93)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(94)  =   "Named:id=40:OddRow"
      _StyleDefs(95)  =   ":id=40,.parent=33"
      _StyleDefs(96)  =   "Named:id=41:RecordSelector"
      _StyleDefs(97)  =   ":id=41,.parent=34"
      _StyleDefs(98)  =   "Named:id=42:FilterBar"
      _StyleDefs(99)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   0
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
Attribute VB_Name = "frm_mst_education"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Public public_int_mode As Integer



Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(*) as rec_count from m_education where upper(education_name) = '" _
& UCase(Trim(txt_education_name)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockBatchOptimistic

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
txt_education_name = ""
If txt_education_name.Enabled = True Then txt_education_name.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not (Trim(UCase(txt_education_name)) = UCase(Adodc1.Recordset.Fields("education_name").Value)) And _
    check_validate_exist_new Then

    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

If cbo_education_code.ListIndex < 0 Then
    MsgBox "Education level is not selected!", vbInformation, headerMSG
    cbo_education_code.SetFocus
    check_validate_new = False
    Exit Function
End If

If Trim(txt_education_name) = "" Then
    MsgBox "Education name is empty!", vbInformation, headerMSG
    txt_education_name.SetFocus
    check_validate_new = False
    Exit Function
End If

End Function

Private Sub load_data_grid()
Adodc1.RecordSource = "select a.* " _
& "from m_education a " _
& "order by a.education_code, a.education_name"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub cmd_refresh_Click()
Call load_data_grid
End Sub

Private Sub cmd_browse_study_Click()
'frm_lookup_mst_study.public_int_mode = 0
'frm_lookup_mst_study.Show 1
End Sub

Private Sub cmd_select_Click()
Dim i As Integer
Dim item

If Not TDBGrid1.ApproxCount > 0 Then
    Exit Sub
End If
    
'Set SelBks = TDBGrid1.SelBookmarks
'i = 0

If public_int_mode = 1 Then
    With frm_mst_family_from
        .txt_education_code = Adodc1.Recordset.Fields("education_code").Value
        .txt_education_name = Adodc1.Recordset.Fields("education_name").Value
    End With
    
ElseIf public_int_mode = 2 Then
    With frm_mst_education_history
        .txt_education_code = Adodc1.Recordset.Fields("education_code").Value
        .txt_education_name = Adodc1.Recordset.Fields("education_name").Value
        .txt_address = Adodc1.Recordset.Fields("address").Value
    End With
    
End If

Call CmdExit_Click

Exit Sub
err_capture:
MsgBox err.Description
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
    & TDBGrid1.Columns("education_name").Value & " ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_education where sequence_number = " _
    & Adodc1.Recordset.Fields("sequence_number").Value
CnG.CommitTrans

Call load_data_grid
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    
    'txt_sequence_number = .Fields("sequence_number").Value
    cbo_education_code.ListIndex = .Fields("education_code").Value
    txt_education_name = .Fields("education_name").Value
    txt_address = .Fields("address").Value
    txt_description = .Fields("description").Value
    'txt_entry_date = .Fields("entry_date").Value
    'txt_deleted = .Fields("deleted").Value
    'txt_delete_date = .Fields("delete_date").Value

End With
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_education where sequence_number = " _
& Adodc1.Recordset.Fields("sequence_number").Value, CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_education where education_name = 'άφ'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub CmdPosting_Click()
Dim rs As New ADODB.Recordset
Dim i As Integer

rs.Open "select count(kode_supplier) as jml_rec from m_supplier " _
& "where isnull(flag_posting,0)=0", CnG, adOpenStatic, adLockBatchOptimistic

If rs.Fields("jml_rec").Value < 1 Then
    MsgBox "Semua Data Supplier sudah terposting", vbInformation, headerMSG
    Exit Sub
End If

i = MsgBox("Anda yakin ingin melakukan posting " & rs.Fields("jml_rec").Value _
& " data supplier ?", vbOKCancel, headerMSG)
If Not i = vbOK Then Exit Sub

If rs.State = 1 Then rs.Close
rs.Open "select * from m_supplier " _
& "where isnull(flag_posting,0)=0 order by kode_supplier", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans
With rs
    While Not .EOF
        .Fields("flag_posting").Value = 1
        .Update
        
        .MoveNext
    Wend
End With
CnG.CommitTrans

MsgBox rs.RecordCount & " data Supplier sudah terposting", vbInformation, headerMSG
Call load_data_grid
int_mode = 0
Call load_mode
End Sub

Private Sub insert_new_data()
Dim rs1 As New ADODB.Recordset

CnG.BeginTrans

With rsBound
    .AddNew
    
    .Fields("sequence_number").Value = get_inc_number("m_education", "sequence_number", "")
    .Fields("education_code").Value = cbo_education_code.ListIndex
    .Fields("education_code_name").Value = Trim(cbo_education_code.Text)
    
    .Fields("education_name").Value = Trim(txt_education_name)
    .Fields("address").Value = Trim(txt_address)
    .Fields("description").Value = Trim(txt_description)
    .Fields("entry_date").Value = Now
    '.Fields("deleted").Value = Trim(txt_deleted)
    '.Fields("delete_date").Value = Trim(txt_delete_date)

    .Update
End With

CnG.CommitTrans
End Sub

Private Sub edit_old_data()
On Error GoTo err_capture
Dim rs1 As New ADODB.Recordset

CnG.BeginTrans
With rsBound

    '.Fields("sequence_number").value = get_inc_number("m_education", "sequence_number", "")
    .Fields("education_code").Value = cbo_education_code.ListIndex
    .Fields("education_code_name").Value = Trim(cbo_education_code.Text)
    
    .Fields("education_name").Value = Trim(txt_education_name)
    .Fields("address").Value = Trim(txt_address)
    .Fields("description").Value = Trim(txt_description)
    .Fields("entry_date").Value = Now
    '.Fields("deleted").Value = Trim(txt_deleted)
    '.Fields("delete_date").Value = Trim(txt_delete_date)
    
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

Call load_data_grid
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

'CmdPrint.Enabled = f
'cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data()
Dim Ctr As Control
For Each Ctr In Me
    If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
        Ctr.Text = ""
    ElseIf TypeOf Ctr Is TDBCombo Then
        Ctr.Text = ""
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
cbo_education_code.ListIndex = 3
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    txt_education_name.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_education_name.Enabled = True Then
        txt_education_name.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    'txt_education_name.Enabled = False
    fra_entry.Visible = True
    TDBGrid1.Enabled = False
End If
End Sub

Private Sub load_mode()
If int_mode = 1 Then        ' new
    Call set_buttons_enable(False, True, False, False, True, False, False)
ElseIf int_mode = 0 Then    ' view
    Call set_buttons_enable(True, False, True, True, False, True, True)
'    cmdEdit.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'    cmdDelete.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
ElseIf int_mode = 2 Then    ' edit/revise
    Call set_buttons_enable(False, True, False, False, True, False, False)
End If

Call set_data_mode
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = strConn

Call load_data_grid

Call load_data_user_access(Me)
int_mode = 0
Call load_mode
End Sub

Private Sub TDBGrid1_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid1, Adodc1)
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
Adodc1.Recordset.Sort = TDBGrid1.Columns(ColIndex).name
End Sub
