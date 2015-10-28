VERSION 5.00
Object = "{E684D8A3-716C-4E59-AA94-7144C04B0074}#1.1#0"; "GridEX20.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_user_level 
   Appearance      =   0  'Flat
   Caption         =   "MASTER USER LEVEL"
   ClientHeight    =   9285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mst_user_level.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   13770
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   645
      Left            =   12300
      Picture         =   "frm_mst_user_level.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8460
      Width           =   975
   End
   Begin VB.Frame fraAdmin 
      Height          =   8205
      Left            =   360
      TabIndex        =   11
      Top             =   120
      Width           =   12975
      Begin VB.CommandButton cmd_delete_dtl 
         Caption         =   "&Delete Dtl"
         Height          =   555
         Left            =   8400
         Picture         =   "frm_mst_user_level.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   7500
         Width           =   855
      End
      Begin VB.CommandButton cmd_save_dtl 
         Caption         =   "&Save Dtl"
         Height          =   555
         Left            =   7440
         Picture         =   "frm_mst_user_level.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   7500
         Width           =   855
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_form 
         Height          =   4245
         Left            =   720
         TabIndex        =   21
         Top             =   3120
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "MENU"
         Columns(0).DataField=   "menu_name"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "FORM TITLE"
         Columns(1).DataField=   "form_title"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   4
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "READ"
         Columns(2).DataField=   "allow_read"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "ADD"
         Columns(3).DataField=   "allow_add"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   4
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "EDIT"
         Columns(4).DataField=   "allow_edit"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   4
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DELETE"
         Columns(5).DataField=   "allow_delete"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   4
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "POST"
         Columns(6).DataField=   "allow_post"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "PRINT"
         Columns(7).DataField=   "allow_print"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2699"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2619"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8708"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=6482"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6403"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1667"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1588"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(2)._MinWidth=3342336"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=1640"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1561"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=1693"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=1614"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(4)._MinWidth=182755968"
         Splits(0)._ColumnProps(28)=   "Column(5).Width=1693"
         Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=1614"
         Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(33)=   "Column(5)._MinWidth=182748752"
         Splits(0)._ColumnProps(34)=   "Column(6).Width=1667"
         Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=1588"
         Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=513"
         Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(39)=   "Column(6)._MinWidth=182794208"
         Splits(0)._ColumnProps(40)=   "Column(7).Width=1746"
         Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=1667"
         Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=513"
         Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(45)=   "Column(7)._MinWidth=182756384"
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
         Caption         =   "LIST OF PRIVILEGES"
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
         _StyleDefs(8)   =   ":id=4,.fgcolor=&H80000009&,.bold=0,.fontsize=900,.italic=0,.underline=0"
         _StyleDefs(9)   =   ":id=4,.strikethrough=0,.charset=0"
         _StyleDefs(10)  =   ":id=4,.fontname=Microsoft Sans Serif"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
         _StyleDefs(66)  =   "Named:id=33:Normal"
         _StyleDefs(67)  =   ":id=33,.parent=0"
         _StyleDefs(68)  =   "Named:id=34:Heading"
         _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(70)  =   ":id=34,.wraptext=-1"
         _StyleDefs(71)  =   "Named:id=35:Footing"
         _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(73)  =   "Named:id=36:Selected"
         _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(75)  =   "Named:id=37:Caption"
         _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(77)  =   "Named:id=38:HighlightRow"
         _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(79)  =   "Named:id=39:EvenRow"
         _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(81)  =   "Named:id=40:OddRow"
         _StyleDefs(82)  =   ":id=40,.parent=33"
         _StyleDefs(83)  =   "Named:id=41:RecordSelector"
         _StyleDefs(84)  =   ":id=41,.parent=34"
         _StyleDefs(85)  =   "Named:id=42:FilterBar"
         _StyleDefs(86)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   300
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   4815
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   120
         Top             =   5880
      End
      Begin VB.CommandButton cmdGenerate 
         Cancel          =   -1  'True
         Caption         =   "&Load Dtl"
         Height          =   555
         Left            =   6480
         Picture         =   "frm_mst_user_level.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   7500
         Width           =   855
      End
      Begin VB.CommandButton cmdEditUser 
         Caption         =   "&Edit"
         Height          =   555
         Left            =   1680
         Picture         =   "frm_mst_user_level.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7500
         Width           =   855
      End
      Begin VB.CommandButton cmdNewUser 
         Caption         =   "&New"
         Height          =   555
         Left            =   720
         Picture         =   "frm_mst_user_level.frx":1BBE
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7500
         Width           =   855
      End
      Begin VB.CommandButton cmdSimpanUser 
         Caption         =   "&Save"
         Height          =   555
         Left            =   3600
         Picture         =   "frm_mst_user_level.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7500
         Width           =   855
      End
      Begin VB.CommandButton cmdDeleteUser 
         Caption         =   "&Delete"
         Height          =   555
         Left            =   2640
         Picture         =   "frm_mst_user_level.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7500
         Width           =   855
      End
      Begin GridEX20.GridEX GridEXForm 
         Height          =   975
         Left            =   11520
         TabIndex        =   12
         Top             =   5640
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1720
         Version         =   "2.0"
         CursorLocation  =   3
         HeaderStyle     =   3
         MethodHoldFields=   -1  'True
         LockType        =   4
         Options         =   8
         RecordsetType   =   3
         ForeColorInfoText=   13759468
         ForeColorHeader =   14415859
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         BackColorHeader =   12632256
         DataMode        =   1
         HeaderFontName  =   "MS Sans Serif"
         HeaderFontBold  =   -1  'True
         HeaderFontWeight=   700
         BackColorBkg    =   16777215
         ColumnHeaderHeight=   285
         ColumnsCount    =   2
         Column(1)       =   "frm_mst_user_level.frx":2C5C
         Column(2)       =   "frm_mst_user_level.frx":2D24
         FormatStylesCount=   6
         FormatStyle(1)  =   "frm_mst_user_level.frx":2DC8
         FormatStyle(2)  =   "frm_mst_user_level.frx":2EF0
         FormatStyle(3)  =   "frm_mst_user_level.frx":2FA0
         FormatStyle(4)  =   "frm_mst_user_level.frx":3054
         FormatStyle(5)  =   "frm_mst_user_level.frx":312C
         FormatStyle(6)  =   "frm_mst_user_level.frx":31E4
         ImageCount      =   0
         PrinterProperties=   "frm_mst_user_level.frx":32C4
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   1680
         OleObjectBlob   =   "frm_mst_user_level.frx":349C
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc_company 
         Height          =   375
         Left            =   1680
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin MSAdodcLib.Adodc Adodc_form 
         Height          =   375
         Left            =   0
         Top             =   3360
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
      Begin VB.Frame fra_EntryUser 
         Caption         =   "Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   720
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   11535
         Begin VB.TextBox txt_user_code 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   300
            Left            =   2760
            MaxLength       =   40
            TabIndex        =   1
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox txt_NamaUser 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   300
            Left            =   2760
            MaxLength       =   40
            TabIndex        =   2
            Top             =   840
            Width           =   2895
         End
         Begin VB.TextBox txt_PasswordUser 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   2760
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1200
            Width           =   7185
         End
         Begin VB.TextBox txt_KodeUser 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   8160
            MaxLength       =   40
            TabIndex        =   16
            Top             =   1440
            Visible         =   0   'False
            Width           =   1750
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "LEVEL CODE"
            Height          =   195
            Left            =   1560
            TabIndex        =   20
            Top             =   480
            Width           =   885
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "LEVEL NAME"
            Height          =   195
            Left            =   1560
            TabIndex        =   18
            Top             =   840
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1560
            TabIndex        =   17
            Top             =   1230
            Width           =   1020
         End
      End
      Begin GridEX20.GridEX GridEXUser 
         Height          =   1935
         Left            =   720
         TabIndex        =   19
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   3413
         Version         =   "2.0"
         HeaderStyle     =   3
         MethodHoldFields=   -1  'True
         ForeColorHeader =   -2147483639
         AllowEdit       =   0   'False
         BorderStyle     =   3
         GroupByBoxVisible=   0   'False
         BackColorHeader =   -2147483646
         HeaderFontName  =   "Microsoft Sans Serif"
         HeaderFontSize  =   9
         ColumnHeaderHeight=   315
         IntProp1        =   0
         ColumnsCount    =   1
         Column(1)       =   "frm_mst_user_level.frx":545A
         FormatStylesCount=   6
         FormatStyle(1)  =   "frm_mst_user_level.frx":54FE
         FormatStyle(2)  =   "frm_mst_user_level.frx":5626
         FormatStyle(3)  =   "frm_mst_user_level.frx":56D6
         FormatStyle(4)  =   "frm_mst_user_level.frx":578A
         FormatStyle(5)  =   "frm_mst_user_level.frx":5862
         FormatStyle(6)  =   "frm_mst_user_level.frx":591A
         ImageCount      =   0
         PrinterProperties=   "frm_mst_user_level.frx":59FA
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "COMPANY"
         Height          =   195
         Left            =   720
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frm_mst_user_level"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Sub EnableButtonEntryForm _
(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, ByVal d As Boolean)
'cmdNewForm.Enabled = a And blnUser_Add
'cmdEditForm.Enabled = b And blnUser_Edit
'cmdDeleteForm.Enabled = c And blnUser_Delete
'cmdSaveForm.Enabled = d
End Sub

Private Sub EnableButtonEntryUser _
(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, ByVal d As Boolean)
cmdNewUser.Enabled = a And blnUser_Add
cmdEditUser.Enabled = b And blnUser_Edit
cmdDeleteUser.Enabled = c And blnUser_Delete
cmdSimpanUser.Enabled = d

TDBGrid_form.Enabled = Not d
cmdGenerate.Enabled = Not d
cmd_save_dtl.Enabled = Not d
End Sub

Private Sub fill_grid_user()
Dim rs1 As New ADODB.Recordset
Dim cmd1 As New ADODB.Command
Set cmd1.ActiveConnection = CnG

cmd1.CommandText = _
"select code,name,remark " _
& "from m_akses_level_group order by code"
rs1.CursorLocation = adUseClient
rs1.Open cmd1, , adOpenStatic, adLockBatchOptimistic

With GridEXUser
    Set .ADORecordset = rs1
    .AllowAddNew = False
    .AllowEdit = False
    .AllowDelete = False
    
    With .Columns("code")
        .Caption = "LEVEL CODE"
        .HeaderAlignment = jgexAlignCenter
        .AllowSizing = False
        .TextAlignment = jgexAlignLeft
        .Width = 2000
    End With
    With .Columns("name")
        .Caption = "LEVEL NAME"
        .HeaderAlignment = jgexAlignCenter
        .AllowSizing = False
        .TextAlignment = jgexAlignLeft
        .Width = 3000
    End With
    With .Columns("remark")
        .Caption = "REMARK"
        .HeaderAlignment = jgexAlignCenter
        .AllowSizing = False
        .TextAlignment = jgexAlignLeft
        .Width = 3000
    End With
    
'    .Columns("user_level").Visible = False
'    .Columns("user_pass_key").Visible = False
'    .Columns("company_code").Visible = False
'    .Columns("flag_company_access").Visible = False
End With
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub
'
'Private Sub ShowWindowEntryTop(ByVal i As Boolean)
'If i = True Then ' Max
'    fraEntryTop.Visible = True
'    GridEXTop.Height = 1435
'ElseIf i = False Then ' Min
'    fraEntryTop.Visible = False
'    GridEXTop.Height = 2655
'End If
'End Sub


Private Sub ShowWindowEntryUser(ByVal i As Boolean)
If i = True Then ' Max
    fra_EntryUser.Visible = True
    'GridEXUser.Height = 960
ElseIf i = False Then ' Min
    fra_EntryUser.Visible = False
    'GridEXUser.Height = 1800
End If
End Sub

Private Sub cmd_delete_dtl_Click()
TDBGrid_form.Delete
End Sub

Private Sub cmd_save_dtl_Click()
TDBGrid_form.Update
End Sub

Private Sub cmdDeleteUser_Click()
If GridEXUser.RowCount < 1 Or GridEXUser.Row < 1 Then Exit Sub

Dim i As Integer
i = MsgBox("Are you sure want to delete data '" _
    & GridEXUser.Value(GridEXUser.Columns("name").ColPosition) & "' ?", vbOKCancel, headerMSG)

If i = vbOK Then
    CnG.Execute "delete from m_akses_level_group where code = '" _
    & GridEXUser.Value(GridEXUser.Columns("code").ColPosition) & "'"
    
    CnG.Execute "delete from t_user where level_code = '" _
    & GridEXUser.Value(GridEXUser.Columns("code").ColPosition) & "'"
    
    Call fill_grid_user
End If
End Sub

Private Sub cmdEditForm_Click()
'Dim lintIndex As Integer
'
'If cmdEditForm.Caption = "&Edit" Then
'    cmdEditForm.Caption = "&Cancel"
'    Call EnableButtonEntryForm(False, True, False, True)
'    Call ShowWindowEntryForm(True)
'    fraEntryForm.Caption = "Edit"
'    txtKodeForm = GridEXF.Value(GridEXF.Columns("form_code").ColPosition)
'    txtNamaForm = GridEXF.Value(GridEXF.Columns("form_name").ColPosition)
'    txt_form_title = GridEXF.Value(GridEXF.Columns("form_title").ColPosition)
'
'    If GridEXF.Value(GridEXF.Columns("user_level").ColPosition) = 100 Then
'        lintIndex = 0
'    ElseIf GridEXF.Value(GridEXF.Columns("user_level").ColPosition) = 1 Then
'        lintIndex = 2
'    ElseIf GridEXF.Value(GridEXF.Columns("user_level").ColPosition) = 2 Then
'        lintIndex = 1
'    End If
'
'    cboLevelForm.ListIndex = lintIndex
'    txtNamaForm.SelStart = 0
'    txtNamaForm.SelLength = Len(Trim(txtNamaForm))
'    txtNamaForm.SetFocus
'    Call EnabledOptionForm(False)
'Else
'    cmdEditForm.Caption = "&Edit"
'    Call EnableButtonEntryForm(True, True, True, False)
'    Call ShowWindowEntryForm(False)
'    Call EnabledOptionForm(True)
'End If
End Sub

Private Sub cmdEditUser_Click()
If GridEXUser.RowCount < 1 Or GridEXUser.Row < 1 Then Exit Sub

If cmdEditUser.Caption = "&Edit" Then
    cmdEditUser.Caption = "&Cancel"
    Call EnableButtonEntryUser(False, True, False, True)
    Call ShowWindowEntryUser(True)
    fra_EntryUser.Caption = "Edit"
    
    txt_user_code = GridEXUser.Value(GridEXUser.Columns("code").ColPosition)
    txt_NamaUser = GridEXUser.Value(GridEXUser.Columns("name").ColPosition)
    txt_PasswordUser = GridEXUser.Value(GridEXUser.Columns("remark").ColPosition)
'    txt_PasswordUser = RC4DeCryptASC(GridEXUser.Value(GridEXUser.Columns("user_pass").ColPosition), pEncryptionPassword)
    'cbo_user_level.ListIndex = IIf(GridEXUser.Value(GridEXUser.Columns("user_level").ColPosition) = 2, 0, 1)
    
    txt_NamaUser.SelStart = 0
    txt_NamaUser.SelLength = Len(Trim(txt_NamaUser))
    txt_NamaUser.SetFocus
    Call EnabledOptionUser(False)
Else
    cmdEditUser.Caption = "&Edit"
    Call EnableButtonEntryUser(True, True, True, False)
    Call ShowWindowEntryUser(False)
    Call EnabledOptionUser(True)
End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub


Private Sub cmdGenerate_Click()
Dim str_user_code As String, int_level As Integer

    If GridEXUser.RowCount < 1 Or GridEXUser.Row < 1 Then Exit Sub
    str_user_code = GridEXUser.Value(GridEXUser.Columns("code").ColPosition)
    int_level = GridEXUser.Value(GridEXUser.Columns("code").ColPosition)
    Call generate_form(str_user_code, int_level)
    Call fill_grid_user_form
End Sub

Private Sub generate_form(ByVal str_user_code As String, ByVal int_level As Integer)
Dim str1 As String

'
'str1 = "Insert into t_user (level_code, sub_menu_code, sub_menu_name, menu_code, menu_name, " _
'& "form_name, form_title, allow_read, allow_add, allow_edit, allow_delete, allow_post, allow_print) " _
'& "Select '" & str_user_code & "', sub_menu_code, sub_menu_name, menu_code, menu_name, form_name, form_title, " _
'& "case when user_level<=" & int_level & " then 1 else 0 end, " _
'& "case when user_level<=" & int_level & " then 1 else 0 end, " _
'& "case when user_level<=" & int_level & " then 1 else 0 end, " _
'& "case when user_level<=" & int_level & " then 1 else 0 end, " _
'& "case when user_level<=" & int_level & " then 1 else 0 end, " _
'& "case when user_level<=" & int_level & " then 1 else 0 end " _
'& "from m_sub_menu where (sub_menu_code <> 'M02-01' AND sub_menu_code <> 'M02-02') AND ifnull(form_name,'')<>'' and sub_menu_code not in " _
'& "(select sub_menu_code from t_user where level_code = '" & str_user_code & "')"

str1 = "Insert into t_user (level_code, sub_menu_code, sub_menu_name, menu_code, menu_name, " _
& "form_name, form_title, allow_read, allow_add, allow_edit, allow_delete, allow_post, allow_print) " _
& "Select '" & str_user_code & "', sub_menu_code, sub_menu_name, menu_code, menu_name, form_name, form_title, " _
& "1,1,1,1,1,1 " _
& "from m_sub_menu where (sub_menu_code <> 'M02-01' AND sub_menu_code <> 'M02-02' AND sub_menu_code <> 'M01-01') AND ifnull(form_name,'')<>'' and sub_menu_code not in " _
& "(select sub_menu_code from t_user where level_code = '" & str_user_code & "')"
'MsgBox str1

CnG.Execute str1
End Sub

Private Sub generate_form_bak(ByVal str_user_code As Integer, ByVal int_level As Integer)
CnG.Execute "Insert into t_user (user_code, form_code, form_name, form_title, company_code, " _
& "allow_read, allow_add, allow_edit, allow_delete, allow_post, allow_print) " _
& "Select " & str_user_code & ", form_code, form_name, form_title, '" _
& TDBCombo_company.Columns("company_code").Value & "', " _
& "case when user_level<=" & int_level & " then 1 else 0 end, " _
& "case when user_level<=" & int_level & " then 1 else 0 end, " _
& "case when user_level<=" & int_level & " then 1 else 0 end, " _
& "case when user_level<=" & int_level & " then 1 else 0 end, " _
& "case when user_level<=" & int_level & " then 1 else 0 end, " _
& "case when user_level<=" & int_level & " then 1 else 0 end " _
& "from m_form where form_code not in " _
& "(select form_code from t_user where user_code = " & str_user_code & ")"
End Sub

Private Sub cmdNewForm_Click()
'If cmdNewForm.Caption = "&New" Then
'    cmdNewForm.Caption = "&Cancel"
'    Call EnableButtonEntryForm(True, False, False, True)
'    Call ShowWindowEntryForm(True)
'    fraEntryForm.Caption = "Entry"
'    txtNamaForm = ""
'    txt_form_title = ""
'    cboLevelForm.ListIndex = 1
'    txtNamaForm.SetFocus
'    Call EnabledOptionForm(False)
'Else
'    cmdNewForm.Caption = "&New"
'    Call EnableButtonEntryForm(True, True, True, False)
'    Call ShowWindowEntryForm(False)
'    Call EnabledOptionForm(True)
'End If
End Sub

Private Function CekValidateDataUser() As Boolean
If Not Trim(txt_NamaUser) = "" And Not Trim(txt_PasswordUser) = "" Then
    CekValidateDataUser = True
Else
    CekValidateDataUser = False
End If
End Function

Private Function CekDuplicateNameUser() As Boolean
Dim cmd1 As New ADODB.Command
Dim rs1 As New ADODB.Recordset
Set cmd1.ActiveConnection = CnG
cmd1.CommandText = "select count(user_name) as JumlahRec from m_user " & _
                   "where user_name = '" & Replace$(Trim$(txt_NamaUser), Chr(39), Chr(96)) & _
                   "'"
rs1.CursorLocation = adUseClient
rs1.Open cmd1, , adOpenStatic, adLockBatchOptimistic

If rs1!JumlahRec > 0 Then
    CekDuplicateNameUser = True
Else
    CekDuplicateNameUser = False
End If
End Function

Private Sub cmdNewUser_Click()
If cmdNewUser.Caption = "&New" Then
    cmdNewUser.Caption = "&Cancel"
    Call EnableButtonEntryUser(True, False, False, True)
    Call ShowWindowEntryUser(True)
    fra_EntryUser.Caption = "Entry"
    
    txt_user_code = ""
    txt_NamaUser = ""
    txt_PasswordUser = ""
    txt_user_code.SetFocus
    
    Call EnabledOptionUser(False)
Else
    cmdNewUser.Caption = "&New"
    Call EnableButtonEntryUser(True, True, True, False)
    Call ShowWindowEntryUser(False)
    Call EnabledOptionUser(True)
End If
End Sub

Private Sub cmdSaveForm_Click()
'Dim rs As New ADODB.Recordset
'Dim lintLevel As Integer
'
'
'If cboLevelForm.ListIndex = 0 Then
'    lintLevel = 100
'ElseIf cboLevelForm.ListIndex = 1 Then
'    lintLevel = 2
'ElseIf cboLevelForm.ListIndex = 2 Then
'    lintLevel = 1
'End If
'
'If fraEntryForm.Caption = "Entry" Then
'    If CekValidateDataForm = False Then
'        MsgBox "Data is not valid", vbCritical, headerMsg
'        Exit Sub
'    End If
'
'    rs.Open "select ifnull(max(form_code),0)+1 as form_code from m_form", cng, adOpenStatic, adLockBatchOptimistic
'
'    cng.Execute _
'    "insert into m_form (form_code, form_name, form_title, user_level) " _
'    & "values(" & rs!form_code _
'    & ",'" & Replace$(Trim$(txtNamaForm), Chr(39), Chr(96)) _
'    & "','" & Replace$(Trim$(txt_form_title), Chr(39), Chr(96)) _
'    & "'," & lintLevel & ")"
'
'    rs.Close
'
'    Call fill_grid_form
'    Call EnableButtonEntryForm(True, True, True, False)
'    Call ShowWindowEntryForm(False)
'    cmdNewForm.Caption = "&New"
'    Call EnabledOptionForm(True)
'
'ElseIf fraEntryForm.Caption = "Edit" Then
'    If CekValidateDataForm = False Then
'        MsgBox "Data is not valid", vbCritical, headerMsg
'        Exit Sub
'    End If
'
'    cng.Execute _
'    "update m_form set form_name = '" & Replace$(Trim$(txtNamaForm), Chr(39), Chr(96)) _
'    & "', form_title = '" & Replace$(Trim$(txt_form_title), Chr(39), Chr(96)) _
'    & "', user_level = " & lintLevel _
'    & " where form_code = " & GridEXF.Value(GridEXF.Columns("form_code").ColPosition)
'
'    Call fill_grid_form
'    Call EnableButtonEntryForm(True, True, True, False)
'    Call ShowWindowEntryForm(False)
'    cmdEditForm.Caption = "&Edit"
'    Call EnabledOptionForm(True)
'End If
End Sub

Private Sub cmdSimpanUser_Click()
Dim rs As New ADODB.Recordset
Dim int_level As Integer
Dim clsFunc As New clsFunction

If fra_EntryUser.Caption = "Entry" Then
    If CekValidateDataUser = False Then
        MsgBox "Data is not valid", vbCritical, headerMSG
        Exit Sub
    End If
    If CekDuplicateNameUser = True Then
        MsgBox "Data found!", vbCritical, headerMSG
        Exit Sub
    End If
    
    rs.Open "select * from m_akses_level_group where code='uOu'", CnG, adOpenKeyset, adLockOptimistic
    
    CnG.BeginTrans
    With rs
        .AddNew
        
        .Fields("code").Value = Trim(txt_user_code)
        .Fields("name").Value = Trim(txt_NamaUser)
        .Fields("remark").Value = Trim(txt_PasswordUser.Text)
        .Fields("user_entry").Value = LOGIN_CODE
        .Fields("date_entry").Value = Now
        '.Fields("company_code").Value = TDBCombo_company.Columns("company_code").Value
        
        .Update
    End With
    clsFunc.InsertLog ("Insert hakakses level : " & txt_user_code.Text)
    
    CnG.CommitTrans
    
    Call fill_grid_user
    Call EnableButtonEntryUser(True, True, True, False)
    Call ShowWindowEntryUser(False)
    cmdNewUser.Caption = "&New"
    Call EnabledOptionUser(True)
    
ElseIf fra_EntryUser.Caption = "Edit" Then
    If CekValidateDataUser = False Then
        MsgBox "Editing Data Not Valid", vbCritical, "Request validate data"
        Exit Sub
    End If
    If Not Trim(txt_NamaUser) = GridEXUser.Value(GridEXUser.Columns("name").ColPosition) _
    And CekDuplicateNameUser = True Then
        MsgBox "Data found!", vbCritical, headerMSG
        Exit Sub
    End If
    
    rs.Open "select * from m_akses_level_group where code='" _
    & GridEXUser.Value(GridEXUser.Columns("code").ColPosition) & "'", CnG, adOpenKeyset, adLockOptimistic
    
    CnG.BeginTrans
    With rs
        
        .Fields("code").Value = Trim(txt_user_code)
        .Fields("name").Value = Trim(txt_NamaUser)
        .Fields("remark").Value = Trim(txt_PasswordUser.Text)
        .Fields("date_edit").Value = Now
        '.Fields("user_level").Value = IIf(cbo_user_level.ListIndex = 0, 2, 1)
        .Fields("user_edit").Value = LOGIN_CODE
        
        .Update
    End With
    
    clsFunc.InsertLog ("Edit hakakses level : " & txt_user_code.Text)
    CnG.CommitTrans
    
    Call fill_grid_user
    Call EnableButtonEntryUser(True, True, True, False)
    Call ShowWindowEntryUser(False)
    cmdEditUser.Caption = "&Edit"
    Call EnabledOptionUser(True)
End If
End Sub

Private Sub Form_Load()
Adodc_company.ConnectionString = strConn
Adodc_form.ConnectionString = strConn

Call load_data_company
Call load_data_user_access(Me)

Call EnableButtonEntryForm(True, True, True, False)
Call EnableButtonEntryUser(True, True, True, False)
Call ShowWindowEntryUser(False)

'Call fill_grid_form

GridEXForm.AllowEdit = blnUser_Edit
Timer1.Enabled = True
End Sub

Public Sub load_data_company()
Adodc_company.RecordSource = "select company_code, company_name from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub GridEXForm_Change()
Dim rs As ADODB.Recordset
GridEXForm.Update
Set rs = New ADODB.Recordset
rs.CursorLocation = adUseClient
Set rs = GridEXForm.ADORecordset
rs.UpdateBatch

Set rs = Nothing
End Sub

Private Sub GridEXUser_RowColChange(ByVal LastRow As Long, ByVal LastCol As Integer)
If GridEXUser.Row < 1 Then
    'Set GridEXForm.ADORecordset = Nothing
    Set TDBGrid_form.DataSource = Nothing
    Exit Sub
End If
Call fill_grid_user_form
End Sub

Private Sub GridEXUser_DblClick()
Call cmdEditUser_Click
End Sub

Private Sub GridEXUser_RowChange()
    Call fill_grid_user_form
End Sub

Private Sub GridEXF_DblClick()
Call cmdEditForm_Click
End Sub

Private Sub EnabledOptionUser(ByVal i As Boolean)
'fraOption.Enabled = i
GridEXUser.Enabled = i
End Sub

Private Sub fill_grid_user_form()
Adodc_form.RecordSource = "select * from t_user where level_code = '" _
    & GridEXUser.Value(GridEXUser.Columns("code").ColPosition) _
    & "' order by menu_code asc, sub_menu_code asc"
Adodc_form.Refresh
Set TDBGrid_form.DataSource = Adodc_form

'Dim rs As New ADODB.Recordset
'
'rs.Open "select * from t_user where user_code = '" _
'& GridEXUser.Value(GridEXUser.Columns("user_code").ColPosition) & _
'"' order by menu_code asc, sub_menu_code asc", cng, adOpenKeyset, adLockOptimistic
'
'With GridEXForm
'    Set .ADORecordset = rs
'    .AllowAddNew = False
'    .AllowEdit = True
'    .AllowDelete = False
'
'    With .Columns("form_title")
'        .Caption = "FORM TITLE"
'        .HeaderAlignment = jgexAlignCenter
'        .AllowSizing = False
'        .TextAlignment = jgexAlignLeft
'        .Width = 4600
'        .EditType = jgexEditNone
'    End With
'    With .Columns("allow_read")
'        .Caption = "READ"
'        .HeaderAlignment = jgexAlignCenter
'        .AllowSizing = False
'        .TextAlignment = jgexAlignCenter
'        .Width = 1000
'        .HasValueList = True
'        .EditType = jgexEditDropDown
'        .ValueList.Add 1, "Yes"
'        .ValueList.Add 0, "No"
'    End With
'    With .Columns("allow_add")
'        .Caption = "ADD"
'        .HeaderAlignment = jgexAlignCenter
'        .AllowSizing = False
'        .TextAlignment = jgexAlignCenter
'        .Width = 1000
'        .HasValueList = True
'        .EditType = jgexEditDropDown
'        .ValueList.Add 1, "Yes"
'        .ValueList.Add 0, "No"
'    End With
'    With .Columns("allow_edit")
'        .Caption = "EDIT"
'        .HeaderAlignment = jgexAlignCenter
'        .AllowSizing = False
'        .TextAlignment = jgexAlignCenter
'        .Width = 1000
'        .HasValueList = True
'        .EditType = jgexEditDropDown
'        .ValueList.Add 1, "Yes"
'        .ValueList.Add 0, "No"
'    End With
'    With .Columns("allow_delete")
'        .Caption = "DELETE"
'        .HeaderAlignment = jgexAlignCenter
'        .AllowSizing = False
'        .TextAlignment = jgexAlignCenter
'        .Width = 1000
'        .HasValueList = True
'        .EditType = jgexEditDropDown
'        .ValueList.Add 1, "Yes"
'        .ValueList.Add 0, "No"
'    End With
'    With .Columns("allow_post")
'        .Caption = "POSTING"
'        .HeaderAlignment = jgexAlignCenter
'        .AllowSizing = False
'        .TextAlignment = jgexAlignCenter
'        .Width = 1200
'        .HasValueList = True
'        .EditType = jgexEditDropDown
'        .ValueList.Add 1, "Yes"
'        .ValueList.Add 0, "No"
'    End With
'    With .Columns("allow_print")
'        .Caption = "PRINTING"
'        .HeaderAlignment = jgexAlignCenter
'        .AllowSizing = False
'        .TextAlignment = jgexAlignCenter
'        .Width = 1200
'        .HasValueList = True
'        .EditType = jgexEditDropDown
'        .ValueList.Add 1, "Yes"
'        .ValueList.Add 0, "No"
'    End With
'
'    .Columns("sub_menu_code").Visible = False
'    .Columns("form_name").Visible = False
'    .Columns("user_code").Visible = False
'    .Columns("menu_code").Visible = False
'End With
End Sub


Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    Set GridEXForm.ADORecordset = Nothing
    Call fill_grid_user
End If
End Sub

Private Sub TDBGrid_form_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid_form, Adodc_form)
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
TDBCombo_company.Locked = IIf(LOGIN_LEVEL = 100, False, True)
End Sub
