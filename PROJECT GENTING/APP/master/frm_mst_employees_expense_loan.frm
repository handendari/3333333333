VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_salary_exp_employee_installment 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EMPLOYEE LOAN"
   ClientHeight    =   9090
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
   Icon            =   "frm_mst_employees_expense_loan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5160
      Top             =   420
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   582
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
      Caption         =   "Adodc2"
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Height          =   2415
      Left            =   240
      TabIndex        =   34
      Top             =   5040
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   4260
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO."
      Columns(0).DataField=   "sequence_number"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "MONTH"
      Columns(1).DataField=   "installment_month"
      Columns(1).NumberFormat=   "External Editor"
      Columns(1).ExternalEditor=   "TDBDate1"
      Columns(1).ExternalEditor.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "AMOUNT"
      Columns(2).DataField=   "installment_amount"
      Columns(2).NumberFormat=   "Standard"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "TO PAY"
      Columns(3).DataField=   "installment_pay"
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   4
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "PAID"
      Columns(4).DataField=   "flag_paid"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3122"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3043"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8705"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4048"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3969"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=4260"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4180"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=4180"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=4101"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2090"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2011"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
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
      Caption         =   "LIST OF LOAN DETAIL"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=1,.locked=0"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=78,.parent=13,.alignment=1,.locked=0"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7200
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   660
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3240
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   23
      Top             =   660
      Width           =   3855
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   7560
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
         Left            =   8880
         Picture         =   "frm_mst_employees_expense_loan.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2310
         Picture         =   "frm_mst_employees_expense_loan.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5520
         Picture         =   "frm_mst_employees_expense_loan.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   11880
         Picture         =   "frm_mst_employees_expense_loan.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   1200
         Picture         =   "frm_mst_employees_expense_loan.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   0
         Picture         =   "frm_mst_employees_expense_loan.frx":1BBE
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4440
         Picture         =   "frm_mst_employees_expense_loan.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3360
         Picture         =   "frm_mst_employees_expense_loan.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3060
      Top             =   420
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1500
      OleObjectBlob   =   "frm_mst_employees_expense_loan.frx":2C5C
      TabIndex        =   1
      Top             =   660
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1260
      Top             =   420
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
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   255
      Left            =   0
      TabIndex        =   43
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
      _Version        =   65536
      _ExtentX        =   3201
      _ExtentY        =   450
      Calendar        =   "frm_mst_employees_expense_loan.frx":4C1A
      Caption         =   "frm_mst_employees_expense_loan.frx":4D1D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frm_mst_employees_expense_loan.frx":4D82
      Keys            =   "frm_mst_employees_expense_loan.frx":4DA0
      Spin            =   "frm_mst_employees_expense_loan.frx":4DFE
      AlignHorizontal =   2
      AlignVertical   =   2
      Appearance      =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "mm-yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "mm-yyyy"
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
      Text            =   "04-2010"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   40296
      CenturyMode     =   0
   End
   Begin VB.Frame fra_entry 
      Height          =   3825
      Left            =   240
      TabIndex        =   27
      Top             =   1110
      Width           =   14175
      Begin VB.TextBox txt_employee_code 
         Height          =   315
         Left            =   4650
         TabIndex        =   44
         Top             =   390
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.TextBox txt_loan_value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10560
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txt_interest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10560
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txt_instalment_times 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10560
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txt_instalment_value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txt_loan_total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Franklin Gothic Book"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10560
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CommandButton cmd_browse_item 
         Caption         =   "..."
         Height          =   300
         Left            =   3360
         TabIndex        =   3
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txt_item_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   33
         Top             =   2760
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox txt_item_code 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
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
         TabIndex        =   29
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2010
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1890
         Width           =   3495
      End
      Begin VB.CommandButton cmd_browse_employee 
         Caption         =   "..."
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   28
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txt_nik 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   19
         Top             =   840
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker_date 
         Height          =   300
         Left            =   2040
         TabIndex        =   18
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   110297091
         CurrentDate     =   39278
      End
      Begin MSComCtl2.DTPicker DTPicker_instalment_start 
         Height          =   315
         Left            =   10560
         TabIndex        =   9
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   110297091
         CurrentDate     =   39270
      End
      Begin MSComCtl2.DTPicker DTPicker_instalment_end 
         Height          =   315
         Left            =   10560
         TabIndex        =   10
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   110297091
         CurrentDate     =   39270
      End
      Begin VB.Label Label5 
         Caption         =   "* yyyy-MM-dd"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   12030
         TabIndex        =   46
         Top             =   2910
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "* yyyy-MM-dd"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   12030
         TabIndex        =   45
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "DESKRIPSI"
         Height          =   195
         Left            =   690
         TabIndex        =   42
         Top             =   1890
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "NILAI PINJJAMAN (IDR)"
         Height          =   195
         Left            =   8280
         TabIndex        =   41
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "NILAI CICILAN (IDR)"
         Height          =   195
         Left            =   8280
         TabIndex        =   40
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "CICILAN AWAL"
         Height          =   195
         Left            =   8280
         TabIndex        =   39
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "BUNGA (%)"
         Height          =   195
         Left            =   8280
         TabIndex        =   38
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "WAKTU CICILAN"
         Height          =   195
         Left            =   8280
         TabIndex        =   37
         Top             =   1440
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "CICILAN AKHIR"
         Height          =   195
         Left            =   8280
         TabIndex        =   36
         Top             =   2880
         Width           =   1125
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL PINJJAMAN (IDR)"
         Height          =   195
         Left            =   8280
         TabIndex        =   35
         Top             =   1800
         Width           =   1800
      End
      Begin VB.Line Line1 
         X1              =   720
         X2              =   7200
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "EXPENSE TYPE"
         Height          =   195
         Left            =   720
         TabIndex        =   32
         Top             =   2760
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "KARYAWAN"
         Height          =   195
         Left            =   720
         TabIndex        =   31
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "TANGGAL"
         Height          =   195
         Left            =   720
         TabIndex        =   30
         Top             =   480
         Width           =   690
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3765
      Left            =   240
      TabIndex        =   26
      Top             =   1170
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   6641
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "DATE"
      Columns(0).DataField=   "date"
      Columns(0).NumberFormat=   "FormatText Event"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EMP. CODE"
      Columns(1).DataField=   "nik"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "EMP. CODE"
      Columns(2).DataField=   "employee_code"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "EMP. NAME"
      Columns(3).DataField=   "employee_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "LOAN (IDR)"
      Columns(4).DataField=   "loan_value"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "INTEREST (%)"
      Columns(5).DataField=   "loan_interest"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "LOAN TOTAL (IDR)"
      Columns(6).DataField=   "loan_total"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "INST. TIMES"
      Columns(7).DataField=   "installment_time"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "DESCRIPTION"
      Columns(8).DataField=   "description"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   9
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=9"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2566"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2487"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2355"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2275"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=4154"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=4075"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2646"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2566"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2117"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2037"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2646"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2566"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(7).Width=1746"
      Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1667"
      Splits(0)._ColumnProps(40)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(42)=   "Column(8).Width=5794"
      Splits(0)._ColumnProps(43)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(8)._WidthInPix=5715"
      Splits(0)._ColumnProps(45)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(46)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(47)=   "Column(8)._MinWidth=131022288"
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
      Caption         =   "LIST OF LOAN"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=82,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=74,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=71,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=72,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=73,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.alignment=1"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=98,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=17"
      _StyleDefs(70)  =   "Named:id=33:Normal"
      _StyleDefs(71)  =   ":id=33,.parent=0"
      _StyleDefs(72)  =   "Named:id=34:Heading"
      _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(74)  =   ":id=34,.wraptext=-1"
      _StyleDefs(75)  =   "Named:id=35:Footing"
      _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(77)  =   "Named:id=36:Selected"
      _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=37:Caption"
      _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(81)  =   "Named:id=38:HighlightRow"
      _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=39:EvenRow"
      _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(85)  =   "Named:id=40:OddRow"
      _StyleDefs(86)  =   ":id=40,.parent=33"
      _StyleDefs(87)  =   "Named:id=41:RecordSelector"
      _StyleDefs(88)  =   ":id=41,.parent=34"
      _StyleDefs(89)  =   "Named:id=42:FilterBar"
      _StyleDefs(90)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "PINJAMAN KARYAWAN"
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
      Left            =   5550
      TabIndex        =   47
      Top             =   0
      Width           =   3675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Perusahaan"
      Height          =   195
      Left            =   240
      TabIndex        =   24
      Top             =   750
      Width           =   855
   End
End
Attribute VB_Name = "frm_mst_salary_exp_employee_installment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim v_value As Double
Dim v_installment_time As Double
Dim v_interest As Double
Dim strsql As String

Private Function check_validate_new() As Boolean
check_validate_new = True

If Trim(txt_employee_code) = "" Then
    MsgBox "Employee Code is empty!", vbOKOnly + vbInformation, headerMSG
    txt_employee_code.SetFocus
    check_validate_new = False
    Exit Function
End If

'If Trim(txt_item_code) = "" Then
'    MsgBox "Item Type is empty!", vbOKOnly + vbInformation, headerMSG
'    txt_item_code.SetFocus
'    check_validate_new = False
'    Exit Function
'End If

End Function

Private Sub load_data()
timer1.Enabled = True
End Sub

Private Sub cmd_browse_Click()
frm_lookup_mst_employee.public_int_mode = 77
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_browse_employee_Click()
frm_lookup_mst_employee.public_int_mode = 162
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_browse_item_Click()
frm_mst_salary_item.public_int_mode = 13
frm_mst_salary_item.cmd_select.Visible = True
frm_mst_salary_item.Show 1
End Sub

Private Sub cmd_refresh_Click()
Call load_data_salary
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
    & Adodc1.Recordset.Fields("employee_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans

With Adodc1.Recordset
    CnG.Execute "delete from td_loan where left(date,10) = '" _
        & Format(.Fields("date").Value, "yyyy-mm-dd") & "' and employee_code = '" _
        & .Fields("employee_code").Value & "'"
    
    CnG.Execute "delete from tm_loan where left(date,10) = '" _
        & Format(.Fields("date").Value, "yyyy-mm-dd") & "' and employee_code = '" _
        & .Fields("employee_code").Value & "'"
End With

'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
CnG.Execute strsql
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

CnG.CommitTrans

Call load_data_salary
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    DTPicker_date = .Fields("date").Value
'    txt_company_code = .Fields("company_code").Value
    txt_employee_code = .Fields("employee_code").Value
    txt_nik = .Fields("nik").Value
'    txt_item_code = .Fields("salary_item_code").Value
    txt_employee_name = .Fields("employee_name").Value
'    txt_item_name = .Fields("salary_item_name").Value
    
'    txt_flag_main_salary = .Fields("flag_main_salary").Value
'    txt_flag_sign = .Fields("flag_sign").Value
'    txt_flag_detail = .Fields("flag_detail").Value
'    txt_flag_per_presence = .Fields("flag_per_presence").Value
'    txt_default_value = .Fields("default_value").Value
'    txt_item_code = .Fields("item_code").Value
'    txt_item_name = .Fields("item_name").Value
'    txt_item_type = .Fields("item_type").Value
'    txt_item_type_name = .Fields("item_type_name").Value
    txt_description = .Fields("description").Value
    
    txt_loan_value = .Fields("loan_value").Value
    txt_interest = .Fields("loan_interest").Value
    '.Fields("loan_total").Value = Val(DropAllComma(txt_loan_total))
    txt_instalment_times = .Fields("installment_time").Value
    '.Fields("installment_value").Value = Val(DropAllComma(txt_instalment_value))
'    DTPicker_instalment_start = .Fields("installment_start").Value
    '.Fields("installment_end").Value = DTPicker_instalment_end
    
    v_value = txt_loan_value.Text
    v_interest = txt_interest.Text
    v_installment_time = txt_instalment_times
End With
End Sub

Private Sub cmdEdit_Click()
If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    MsgBox "No Data selected!", vbInformation, headerMSG
    Exit Sub
End If

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
int_mode = 1

Set TDBGrid2.DataSource = Nothing

Call load_mode
End Sub

Private Sub cmdPrint_Click()
TDBGrid1.PrintInfo.PageSetup
If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
End If
End Sub

Private Sub insert_new_data()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim i As Integer

If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from tm_loan where employee_code = '-77'", CnG, adOpenKeyset, adLockOptimistic

'rs1.Open "select * from m_salary_item where salary_item_code = '" _
'    & Trim(txt_item_code) & "'", CnG, adOpenStatic, adLockReadOnly

CnG.BeginTrans

'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
CnG.Execute strsql
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

With rsBound
    .AddNew
    
    .Fields("date").Value = DTPicker_date
    .Fields("company_code").Value = TDBCombo_company.Columns("company_code").Value
    .Fields("employee_code").Value = Trim(txt_employee_code)
'    .Fields("salary_item_code").Value = Trim(txt_item_code)
'    -------------------------------------------------------------------------
    .Fields("employee_name").Value = Trim(txt_employee_name)
'    .Fields("salary_item_name").Value = Trim(txt_item_name)
    
'    .Fields("flag_main_salary").Value = rs1.Fields("flag_main_salary").Value
'    .Fields("flag_sign").Value = rs1.Fields("flag_sign").Value
'    .Fields("flag_detail").Value = rs1.Fields("flag_detail").Value
'    .Fields("flag_per_presence").Value = rs1.Fields("flag_per_presence").Value
'    .Fields("flag_installment").Value = rs1.Fields("flag_installment").Value
'    .Fields("default_value").Value = rs1.Fields("default_value").Value
'
'    .Fields("item_code").Value = rs1.Fields("item_code").Value
'    .Fields("item_name").Value = rs1.Fields("item_name").Value
'    .Fields("item_type").Value = rs1.Fields("item_type").Value
'    .Fields("item_type_name").Value = rs1.Fields("item_type_name").Value
    .Fields("description").Value = Trim(txt_description)
    
    .Fields("loan_value").Value = Val(DropAllComma(txt_loan_value))
    .Fields("loan_interest").Value = Val(DropAllComma(txt_interest))
    .Fields("loan_total").Value = Val(DropAllComma(txt_loan_total))
    .Fields("installment_time").Value = Val(DropAllComma(txt_instalment_times))
    .Fields("installment_value").Value = Val(DropAllComma(txt_instalment_value))
    .Fields("installment_start").Value = DTPicker_instalment_start
    .Fields("installment_end").Value = DTPicker_instalment_end
    
    .Update
End With

rs2.Open "select * from td_loan where salary_item_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic
For i = 1 To Val(txt_instalment_times)
    rs2.AddNew
    rs2.Fields("date").Value = DTPicker_date
    rs2.Fields("employee_code").Value = Trim(txt_employee_code)
'    rs2.Fields("salary_item_code").Value = Trim(txt_item_code)
    rs2.Fields("sequence_number").Value = i
    '------------------------------------
    rs2.Fields("employee_name").Value = Trim(txt_employee_name)
'    rs2.Fields("salary_item_name").Value = Trim(txt_item_name)
    
    rs2.Fields("installment_month").Value = DateAdd("m", i - 1, DTPicker_instalment_start)
    
    rs2.Fields("installment_amount").Value = Val(DropAllComma(txt_instalment_value))
    rs2.Fields("installment_pay").Value = Val(DropAllComma(txt_instalment_value))
    rs2.Fields("installment_pay_date").Value = DateAdd("m", i - 1, DTPicker_instalment_start)
    rs2.Fields("flag_paid").Value = 0
    rs2.Update
Next i

CnG.CommitTrans
End Sub

Private Sub edit_old_data()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim i As Integer
'On Error GoTo err_capture

If rsBound.State = 1 Then rsBound.Close
'rsBound.Open "select * from tm_loan where employee_code = '" _
'& Adodc1.Recordset.Fields("employee_code").Value & "' and date = '" _
'& Format(Adodc1.Recordset.Fields("date").Value, "yyyy-mm-dd hh:nn:ss") & "' and company_code='" _
'& TDBCombo_company.Columns("company_code").Value & "' and salary_item_code='" _
'& Adodc1.Recordset.Fields("salary_item_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

rsBound.Open "select * from tm_loan where employee_code = '" _
& Adodc1.Recordset.Fields("employee_code").Value & "' and date(date) = '" _
& Format(Adodc1.Recordset.Fields("date").Value, "yyyy-mm-dd") & "' and company_code='" _
& TDBCombo_company.Columns("company_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

rs1.Open "select * from m_salary_item where salary_item_code = '" _
    & Trim(txt_item_code) & "'", CnG, adOpenStatic, adLockReadOnly

CnG.BeginTrans
'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
If v_value <> txt_loan_value.Text Or v_interest <> txt_interest.Text Or v_installment_time <> txt_instalment_times.Text Then
    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute strsql
End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

With rsBound

    .Fields("date").Value = Format(DTPicker_date, "yyyy-MM-dd")
    .Fields("company_code").Value = TDBCombo_company.Columns("company_code").Value
    .Fields("employee_code").Value = Trim(txt_employee_code)
'    .Fields("salary_item_code").Value = Trim(txt_item_code)
    '-------------------------------------------------------------------------
    .Fields("employee_name").Value = Trim(txt_employee_name)
'    .Fields("salary_item_name").Value = Trim(txt_item_name)
    
'    .Fields("flag_main_salary").Value = rs1.Fields("flag_main_salary").Value
'    .Fields("flag_sign").Value = rs1.Fields("flag_sign").Value
'    .Fields("flag_detail").Value = rs1.Fields("flag_detail").Value
'    .Fields("flag_per_presence").Value = rs1.Fields("flag_per_presence").Value
'    .Fields("flag_installment").Value = rs1.Fields("flag_installment").Value
'    .Fields("default_value").Value = rs1.Fields("default_value").Value
'
'    .Fields("item_code").Value = rs1.Fields("item_code").Value
'    .Fields("item_name").Value = rs1.Fields("item_name").Value
'    .Fields("item_type").Value = rs1.Fields("item_type").Value
'    .Fields("item_type_name").Value = rs1.Fields("item_type_name").Value
    .Fields("description").Value = Trim(txt_description)
    
    .Fields("loan_value").Value = Val(DropAllComma(txt_loan_value))
    .Fields("loan_interest").Value = Val(DropAllComma(txt_interest))
    .Fields("loan_total").Value = Val(DropAllComma(txt_loan_total))
    .Fields("installment_time").Value = Int(DropAllComma(txt_instalment_times))
    .Fields("installment_value").Value = Val(DropAllComma(txt_instalment_value))
    .Fields("installment_start").Value = DTPicker_instalment_start
    .Fields("installment_end").Value = DTPicker_instalment_end

    .Update
End With

With Adodc1.Recordset
    CnG.Execute "delete from td_loan where left(date,10) = '" _
        & Format(.Fields("date").Value, "yyyy-mm-dd") & "' and employee_code = '" _
        & .Fields("employee_code").Value & "'"
End With

rs2.Open "select * from td_loan where salary_item_code = 'uOu'", CnG, adOpenKeyset, adLockOptimistic
For i = 1 To Int(txt_instalment_times)
    rs2.AddNew
    rs2.Fields("date").Value = DTPicker_date
    rs2.Fields("employee_code").Value = Trim(txt_employee_code)
'    rs2.Fields("salary_item_code").Value = Trim(txt_item_code)
    rs2.Fields("sequence_number").Value = i
    '------------------------------------
    rs2.Fields("employee_name").Value = Trim(txt_employee_name)
'    rs2.Fields("salary_item_name").Value = Trim(txt_item_name)
    
    rs2.Fields("installment_month").Value = DateAdd("m", i - 1, DTPicker_instalment_start)
    
    rs2.Fields("installment_amount").Value = Val(DropAllComma(txt_instalment_value))
    rs2.Fields("installment_pay").Value = Val(DropAllComma(txt_instalment_value))
    rs2.Fields("installment_pay_date").Value = DateAdd("m", i - 1, DTPicker_instalment_start)
    rs2.Fields("flag_paid").Value = 0
    rs2.Update
Next i


CnG.CommitTrans

Exit Sub
err_capture:
MsgBox Err.Description
rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
End Sub

Private Sub CmdSave_Click()
If int_mode = 1 Then
    If Not check_validate_new Then Exit Sub
'    If check_validate_exist_new Then
'        Call check_invalid: Exit Sub
'    End If
    Call insert_new_data
ElseIf int_mode = 2 Then
    If Not check_validate_new Then Exit Sub
'    If check_validate_exist_edit Then
'        Call check_invalid: Exit Sub
'    End If
    Call edit_old_data
End If

Call load_data_salary
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
        If Not LCase(Ctr.name) = "txt_company_name" Then Ctr.Text = ""
    ElseIf TypeOf Ctr Is TDBCombo Then
        If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
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
txt_employee_code = ""
txt_employee_name = ""
DTPicker_date.Value = Now
txt_description = ""
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    txt_employee_code.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_employee_code.Enabled = True Then
        txt_nik.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
    txt_employee_code.Enabled = False
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
Adodc2.ConnectionString = strConn
Adodc_company.ConnectionString = strConn

Call load_data_company

Call load_data_user_access(Me)
int_mode = 0
Call load_mode
timer1.Enabled = True

'cmdEdit.Enabled = False
'cmdDelete.Enabled = False
'CmdNew.Enabled = False

End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    Call load_data_salary
End If
End Sub

Private Sub TDBGrid1_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid1, Adodc1)
If TDBGrid1.ApproxCount > 0 Then load_data_detail
End Sub

Private Sub load_data_salary()
'Adodc1.RecordSource = "select * from tm_loan where company_code = '" _
'    & TDBCombo_company.Columns("company_code").Value & "' and flag_sign=-1 " _
'    & "order by employee_code asc, salary_item_code asc, date desc"
'Adodc1.Refresh
Adodc1.RecordSource = "select date(date) date,b.employee_code,b.employee_name, " _
    & "loan_value,loan_interest,loan_total,installment_time,b.description,b.nik " _
    & "from tm_loan a join m_employee b on a.employee_code = b.employee_code where a.company_code = '" _
    & TDBCombo_company.Columns("company_code").Value & "' " _
    & "order by employee_code asc, date desc"
Adodc1.Refresh

'cmdEdit.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'cmdDelete.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'CmdNew.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)

TDBGrid1.DataSource = Adodc1
Call load_data_detail
End Sub

Private Sub load_data_detail()
Dim str1 As String

With Adodc1.Recordset
    If Adodc1.Recordset.RecordCount > 0 Then
'        str1 = "select * from td_loan where left(date,10) = '" _
'            & Format(.Fields("date").Value, "yyyy-mm-dd") & "' and employee_code ='" _
'            & .Fields("employee_code").Value & "' and salary_item_code = '" _
'            & .Fields("salary_item_code").Value & "' order by sequence_number asc"
        str1 = "select * from td_loan where Date(date) = '" _
            & Format(.Fields("date").Value, "yyyy-mm-dd") & "' and employee_code ='" _
            & .Fields("employee_code").Value & "' order by sequence_number asc"
    Else
        str1 = "select * from td_loan where employee_code='uOu'"
    End If
End With

Adodc2.RecordSource = str1
Adodc2.Refresh
TDBGrid2.DataSource = Adodc2
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "DATE" Then
    Value = Format(Value, "dd-mm-yyyy")
End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If (TDBGrid1.Row + 1) > 0 And (TDBGrid1.Row + 1) <> LastRow Then
    'MsgBox "LETS..."
    Call load_data_detail
End If
End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub txt_department_code_Change()

End Sub

'Private Sub txt_salary_Validate(Cancel As Boolean)
'If Not Trim(txt_salary) = "" Then
'    txt_salary = FormatNumber(DropAllComma(txt_salary))
'End If
'End Sub




' --


Private Sub txt_loan_value_Validate(Cancel As Boolean)
If Not Trim(txt_loan_value) = "" Then
    txt_loan_value = FormatNumber(DropAllComma(txt_loan_value))
End If
End Sub

Private Sub txt_instalment_times_Change()
Call calc_time_interval

Call calc_loan_total
Call calc_instalment_value
End Sub

Private Sub DTPicker_instalment_start_Change()
Call calc_time_interval
End Sub

Private Sub txt_loan_value_Change()
Call calc_loan_total
Call calc_instalment_value
End Sub

Private Sub txt_interest_Change()
Call calc_loan_total
Call calc_instalment_value
End Sub

'================================================================

Private Sub calc_time_interval()
Dim dt1, dt2 As Date

If Trim(txt_instalment_times) = "" Then Exit Sub
If Not Int(txt_instalment_times) > 0 Then Exit Sub

dt1 = DTPicker_instalment_start.Value
dt2 = DateAdd("m", Int(txt_instalment_times) - 1, dt1)

DTPicker_instalment_end = dt2
End Sub

Private Sub calc_instalment_value()
Dim dt1, dt2 As Date
Dim j As Double

If Trim(txt_instalment_times) = "" Then Exit Sub
If Not Int(txt_instalment_times) > 0 Then Exit Sub

j = Val(DropAllComma(txt_loan_total)) / Val(DropAllComma(txt_instalment_times))
txt_instalment_value = FormatNumber(j)
End Sub

Private Sub calc_loan_total()
Dim i, t, j As Double

i = ((Val(DropAllComma(txt_interest)) / 100) * Val(DropAllComma(txt_loan_value))) / 12
t = Val(DropAllComma(txt_instalment_times))
j = Val(DropAllComma(txt_loan_value))

txt_loan_total = FormatNumber((i * t) + j)
End Sub


