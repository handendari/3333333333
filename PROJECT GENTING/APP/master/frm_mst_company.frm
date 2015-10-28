VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_company 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER COMPANY"
   ClientHeight    =   6855
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   14685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mst_company.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   22
      Top             =   5400
      Width           =   14175
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton cmd_refresh 
         Caption         =   "&Refresh"
         Height          =   645
         Left            =   7440
         Picture         =   "frm_mst_company.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2160
         Picture         =   "frm_mst_company.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5400
         Picture         =   "frm_mst_company.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   11520
         Picture         =   "frm_mst_company.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   1080
         Picture         =   "frm_mst_company.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   8520
         Picture         =   "frm_mst_company.frx":1BBE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4320
         Picture         =   "frm_mst_company.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3240
         Picture         =   "frm_mst_company.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
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
   Begin VB.Frame fra_entry 
      Height          =   3495
      Left            =   240
      TabIndex        =   21
      Top             =   1800
      Width           =   14175
      Begin VB.TextBox txt_hari_kerja 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9240
         MaxLength       =   30
         TabIndex        =   39
         Top             =   2970
         Width           =   1005
      End
      Begin VB.TextBox txtNPP 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9240
         MaxLength       =   30
         TabIndex        =   36
         Top             =   2610
         Width           =   3495
      End
      Begin VB.TextBox msk_npwp 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9240
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1530
         Width           =   3495
      End
      Begin VB.TextBox msk_npwp_pimpinan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9240
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2250
         Width           =   3495
      End
      Begin VB.TextBox txt_pimpinan 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9240
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1890
         Width           =   3495
      End
      Begin VB.TextBox txt_city_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1860
         Width           =   3495
      End
      Begin VB.TextBox txt_email_address 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9240
         MaxLength       =   50
         TabIndex        =   9
         Top             =   1170
         Width           =   3495
      End
      Begin VB.TextBox txt_web_address 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9240
         MaxLength       =   50
         TabIndex        =   8
         Top             =   810
         Width           =   3495
      End
      Begin VB.TextBox txt_fax_number 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9240
         MaxLength       =   50
         TabIndex        =   7
         Top             =   450
         Width           =   3495
      End
      Begin VB.TextBox txt_phone_number 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2580
         Width           =   3495
      End
      Begin VB.TextBox txt_postal_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2220
         Width           =   3495
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   2
         Top             =   780
         Width           =   3495
      End
      Begin VB.TextBox txt_address 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   2520
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1140
         Width           =   3495
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
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_company_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   1
         Top             =   420
         Width           =   1695
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "HARI KERJA"
         Height          =   195
         Left            =   7560
         TabIndex        =   40
         Top             =   3030
         Width           =   885
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "N P P"
         Height          =   195
         Left            =   7560
         TabIndex        =   37
         Top             =   2670
         Width           =   375
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "NPWP PEMOTONG PJK"
         Height          =   195
         Left            =   7560
         TabIndex        =   35
         Top             =   2310
         Width           =   1620
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "PEMOTONG PAJAK"
         Height          =   195
         Left            =   7560
         TabIndex        =   34
         Top             =   1950
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "KOTA"
         Height          =   195
         Left            =   1200
         TabIndex        =   33
         Top             =   1860
         Width           =   405
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "N P W P"
         Height          =   195
         Left            =   7560
         TabIndex        =   32
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "EMAIL"
         Height          =   195
         Left            =   7560
         TabIndex        =   31
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "WEBSITE"
         Height          =   195
         Left            =   7560
         TabIndex        =   30
         Top             =   840
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "NO. FAX"
         Height          =   195
         Left            =   7560
         TabIndex        =   29
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "NO. TELP"
         Height          =   195
         Left            =   1200
         TabIndex        =   28
         Top             =   2610
         Width           =   675
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "KODE POS"
         Height          =   195
         Left            =   1200
         TabIndex        =   27
         Top             =   2220
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ALAMAT"
         Height          =   195
         Left            =   1200
         TabIndex        =   26
         Top             =   1140
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "KODE PRSH."
         Height          =   195
         Left            =   1200
         TabIndex        =   25
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NAMA PRSH."
         Height          =   195
         Left            =   1200
         TabIndex        =   24
         Top             =   780
         Width           =   930
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   7223
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
      Columns(2).Caption=   "ADDRESS"
      Columns(2).DataField=   "address"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "POSTAL CODE"
      Columns(3).DataField=   "postal_code"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "PHONE NUMBER"
      Columns(4).DataField=   "phone_number"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "FAX NUMBER"
      Columns(5).DataField=   "fax_number"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "WEB ADDRESS"
      Columns(6).DataField=   "web_address"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "EMAIL ADDRESS"
      Columns(7).DataField=   "email_address"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=6773"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6694"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=4630"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4551"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2355"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2540"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2461"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2566"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2487"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(5)._MinWidth=10"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2566"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2487"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(6)._MinWidth=54215968"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(43)=   "Column(7)._MinWidth=54215968"
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
      Caption         =   "LIST OF COMPANY"
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
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=102,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=110,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=107,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=108,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=109,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
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
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "MASTER PERUSAHAAN"
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
      Left            =   5160
      TabIndex        =   38
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frm_mst_company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim strsql As String



Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(company_code) as rec_count from m_company where company_code = '" _
& Replace$(Trim$(txt_company_code), Chr$(39), Chr$(96)) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
txt_company_code = ""
If txt_company_code.Enabled = True Then txt_company_code.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not txt_company_code = Adodc1.Recordset.Fields("company_code").Value And _
check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

'validasi company code
If Trim(txt_company_code) = "" Then
    MsgBox "Company Code is empty!", vbOKOnly + vbInformation, headerMSG
    txt_company_code.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi company name
If Trim(txt_company_name) = "" Then
    MsgBox "Company Name is empty!", vbOKOnly + vbInformation, headerMSG
    txt_company_name.SetFocus
    check_validate_new = False
    Exit Function
End If

''validasi address
'If Trim(txt_address) = "" Then
'    MsgBox "Address is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_address.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'
''validasi postal code
'If Trim(txt_postal_code) = "" Then
'    MsgBox "Postal Code is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_postal_code.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'
''validasi phone number
'If Trim(txt_phone_number) = "" Then
'    MsgBox "Phone Number is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_phone_number.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'
''validasi fax number
'If Trim(txt_fax_number) = "" Then
'    MsgBox "Fax. Number is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_fax_number.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'
''validasi web address
'If Trim(txt_web_address) = "" Then
'    MsgBox "Web Address is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_web_address.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
'
''validasi email address
'If Trim(txt_email_address) = "" Then
'    MsgBox "Email Address is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_email_address.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
End Function

Private Sub load_data_grid()
Adodc1.RecordSource = "select * from m_company order by company_code"
Adodc1.Refresh

'cmdEdit.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'cmdDelete.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub cmd_refresh_Click()
Call load_data_grid
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
    & TDBGrid1.Columns("company_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from m_company where company_code = '" & TDBGrid1.Columns("company_code").Value & "'"
'+++++++++++++++++++++++++++++++ Delete tem salary Proses +++++++++++++++++++++++++++++++++++++++++++++++
CnG.Execute "delete from temp_sal_proses where company_code = '" & TDBGrid1.Columns("company_code").Value & "'"
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
CnG.CommitTrans

Call load_data_grid
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
On Error Resume Next
With Adodc1.Recordset
    txt_company_code = .Fields("company_code").Value
    txt_company_name = .Fields("company_name").Value
    txt_address = .Fields("address").Value
    txt_city_name = "" & .Fields("city_name").Value
    
    txt_postal_code = .Fields("postal_code").Value
    txt_phone_number = .Fields("phone_number").Value
    txt_fax_number = .Fields("fax_number").Value
    txt_web_address = .Fields("web_address").Value
    txt_email_address = .Fields("email_address").Value
    msk_npwp = "" & .Fields("npwp").Value
    txt_pimpinan = .Fields("pimpinan").Value
    msk_npwp_pimpinan = .Fields("pimpinan_npwp").Value
    txtNPP = .Fields("npp").Value
    txt_hari_kerja = .Fields("hari_kerja").Value
End With
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_company where company_code = '" _
& Adodc1.Recordset.Fields("company_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_company where company_code = '��'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 1
Call load_mode
End Sub

Private Sub CmdPosting_Click()
Dim rs As New ADODB.Recordset
Dim i As Integer

rs.Open "select count(kode_supplier) as jml_rec from m_supplier " _
& "where isnull(flag_posting,0)=0", CnG, adOpenStatic, adLockReadOnly

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
    
    .Fields("company_code").Value = Trim(txt_company_code)      ' key
    .Fields("company_name").Value = Trim(txt_company_name)
    .Fields("address").Value = Trim(txt_address)
    .Fields("city_name").Value = Trim(txt_city_name)
    
    .Fields("postal_code").Value = Trim(txt_postal_code)
    .Fields("phone_number").Value = Trim(txt_phone_number)
    .Fields("fax_number").Value = Trim(txt_fax_number)
    .Fields("web_address").Value = Trim(txt_web_address)
    .Fields("email_address").Value = Trim(txt_email_address)
    .Fields("npwp").Value = Trim(msk_npwp)
    .Fields("pimpinan").Value = Trim(txt_pimpinan)
    .Fields("pimpinan_npwp").Value = Trim(msk_npwp_pimpinan)
    .Fields("npp").Value = Trim(txtNPP.Text)
    .Fields("hari_kerja").Value = Val(txt_hari_kerja.Text)
    
    .Update
End With

    '+++++++++++ Insert into temp salary Proses +++++++++++++++++++++++++++++++++++++
    strsql = "INSERT into temp_sal_proses VALUES('" & Trim(txt_company_code) & "',0)"
    CnG.Execute strsql
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

CnG.CommitTrans
End Sub

Private Sub edit_old_data()
On Error GoTo err_capture

CnG.BeginTrans
With rsBound

    .Fields("company_code").Value = Trim(txt_company_code)     ' key
    .Fields("company_name").Value = Trim(txt_company_name)
    .Fields("address").Value = Trim(txt_address)
    .Fields("city_name").Value = Trim(txt_city_name)
    
    .Fields("postal_code").Value = Trim(txt_postal_code)
    .Fields("phone_number").Value = Trim(txt_phone_number)
    .Fields("fax_number").Value = Trim(txt_fax_number)
    .Fields("web_address").Value = Trim(txt_web_address)
    .Fields("email_address").Value = Trim(txt_email_address)
    .Fields("npwp").Value = Trim(msk_npwp)
    .Fields("pimpinan").Value = Trim(txt_pimpinan)
    .Fields("pimpinan_npwp").Value = Trim(msk_npwp_pimpinan)
    .Fields("npp").Value = Trim(txtNPP.Text)
    .Fields("hari_kerja").Value = Val(txt_hari_kerja.Text)
    
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

CmdPrint.Enabled = F
cmd_refresh.Enabled = g
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
'cbo_jns_kelamin.ListIndex = 1
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    txt_company_code.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_company_code.Enabled = True Then
        txt_company_code.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    txt_company_code.Enabled = False
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

Call load_data_grid

Call load_data_user_access(Me)
int_mode = 0
Call load_mode
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