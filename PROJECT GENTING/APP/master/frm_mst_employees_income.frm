VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_salary_inc_employee 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ADDITIONAL INCOME"
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
   Icon            =   "frm_mst_employees_income.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_entry 
      Height          =   3495
      Left            =   240
      TabIndex        =   19
      Top             =   3960
      Width           =   14175
      Begin VB.CommandButton cmd_browse_item 
         Caption         =   "..."
         Height          =   300
         Left            =   4920
         TabIndex        =   4
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txt_item_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   29
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox txt_item_code 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   28
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txt_salary_value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2160
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
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3600
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2520
         Width           =   3495
      End
      Begin VB.CommandButton cmd_browse_employee 
         Caption         =   "..."
         Height          =   300
         Left            =   4920
         TabIndex        =   3
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   5400
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txt_employee_code 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3600
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker_date 
         Height          =   300
         Left            =   3600
         TabIndex        =   2
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   186449923
         CurrentDate     =   39278
      End
      Begin VB.Line Line1 
         X1              =   2280
         X2              =   8760
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "VALUE"
         Height          =   195
         Left            =   2280
         TabIndex        =   27
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   2280
         TabIndex        =   26
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "INCOME TYPE"
         Height          =   195
         Left            =   2280
         TabIndex        =   25
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "EMPLOYEE"
         Height          =   195
         Left            =   2280
         TabIndex        =   24
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DATE"
         Height          =   195
         Left            =   2280
         TabIndex        =   23
         Top             =   480
         Width           =   390
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   7200
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   15
      Top             =   240
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
         Left            =   8640
         Picture         =   "frm_mst_employees_income.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   1800
         Picture         =   "frm_mst_employees_income.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5040
         Picture         =   "frm_mst_employees_income.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   12000
         Picture         =   "frm_mst_employees_income.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   720
         Picture         =   "frm_mst_employees_income.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   0
         Picture         =   "frm_mst_employees_income.frx":1BBE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   3960
         Picture         =   "frm_mst_employees_income.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   2880
         Picture         =   "frm_mst_employees_income.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   3000
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_mst_employees_income.frx":2C5C
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1200
      Top             =   0
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6735
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11880
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
      Columns(1).DataField=   "employee_code"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "EMP. NAME"
      Columns(2).DataField=   "employee_name"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "ITEM CODE"
      Columns(3).DataField=   "salary_item_code"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "ITEM NAME"
      Columns(4).DataField=   "salary_item_name"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "TYPE"
      Columns(5).DataField=   "item_name"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "STATUS"
      Columns(6).DataField=   "item_type_name"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   4
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "MAIN"
      Columns(7).DataField=   "flag_main_salary"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   4
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "DETAIL"
      Columns(8).DataField=   "flag_detail"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   4
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "PRESENCE"
      Columns(9).DataField=   "flag_per_presence"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   16
      Columns(10)._MaxComboItems=   5
      Columns(10).ValueItems(0)._DefaultItem=   0
      Columns(10).ValueItems(0).Value=   "1"
      Columns(10).ValueItems(0).Value.vt=   8
      Columns(10).ValueItems(0).DisplayValue=   "+"
      Columns(10).ValueItems(0).DisplayValue.vt=   8
      Columns(10).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(10).ValueItems(1)._DefaultItem=   0
      Columns(10).ValueItems(1).Value=   "-1"
      Columns(10).ValueItems(1).Value.vt=   8
      Columns(10).ValueItems(1).DisplayValue=   "-"
      Columns(10).ValueItems(1).DisplayValue.vt=   8
      Columns(10).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(10).ValueItems.Count=   2
      Columns(10).Caption=   "SIGN"
      Columns(10).DataField=   "flag_sign"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "AMOUNT"
      Columns(11).DataField=   "salary_value"
      Columns(11).NumberFormat=   "Standard"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
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
      Splits(0)._ColumnProps(11)=   "Column(2).Width=4154"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4075"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2170"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2090"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=5212"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=5133"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2963"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2884"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=1746"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=1667"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(37)=   "Column(7).Width=1323"
      Splits(0)._ColumnProps(38)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(7)._WidthInPix=1244"
      Splits(0)._ColumnProps(40)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(42)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(43)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(45)=   "Column(8).Width=1270"
      Splits(0)._ColumnProps(46)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(8)._WidthInPix=1191"
      Splits(0)._ColumnProps(48)=   "Column(8)._ColStyle=513"
      Splits(0)._ColumnProps(49)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(51)=   "Column(9).Width=1614"
      Splits(0)._ColumnProps(52)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(53)=   "Column(9)._WidthInPix=1535"
      Splits(0)._ColumnProps(54)=   "Column(9)._ColStyle=513"
      Splits(0)._ColumnProps(55)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(56)=   "Column(10).Width=979"
      Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=900"
      Splits(0)._ColumnProps(59)=   "Column(10)._ColStyle=513"
      Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(61)=   "Column(11).Width=2302"
      Splits(0)._ColumnProps(62)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(11)._WidthInPix=2223"
      Splits(0)._ColumnProps(64)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(65)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(66)=   "Column(11)._MinWidth=131022288"
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
      Caption         =   "LIST OF ADDITIONAL INCOME"
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
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=74,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=70,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=43,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=44,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=45,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
      _StyleDefs(82)  =   "Named:id=33:Normal"
      _StyleDefs(83)  =   ":id=33,.parent=0"
      _StyleDefs(84)  =   "Named:id=34:Heading"
      _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(86)  =   ":id=34,.wraptext=-1"
      _StyleDefs(87)  =   "Named:id=35:Footing"
      _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(89)  =   "Named:id=36:Selected"
      _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(91)  =   "Named:id=37:Caption"
      _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(93)  =   "Named:id=38:HighlightRow"
      _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(95)  =   "Named:id=39:EvenRow"
      _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(97)  =   "Named:id=40:OddRow"
      _StyleDefs(98)  =   ":id=40,.parent=33"
      _StyleDefs(99)  =   "Named:id=41:RecordSelector"
      _StyleDefs(100) =   ":id=41,.parent=34"
      _StyleDefs(101) =   "Named:id=42:FilterBar"
      _StyleDefs(102) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frm_mst_salary_inc_employee"
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
Dim strsql As String

Private Function check_validate_new() As Boolean
check_validate_new = True

If Trim(txt_employee_code) = "" Then
    MsgBox "Employee Code is empty!", vbOKOnly + vbInformation, headerMSG
    txt_employee_code.SetFocus
    check_validate_new = False
    Exit Function
End If

If Trim(txt_item_code) = "" Then
    MsgBox "Item Type is empty!", vbOKOnly + vbInformation, headerMSG
    txt_item_code.SetFocus
    check_validate_new = False
    Exit Function
End If

End Function

Private Sub load_data()
timer1.Enabled = True
End Sub

Private Sub cmd_browse_Click()
frm_lookup_mst_employee.public_int_mode = 77
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_browse_employee_Click()
frm_lookup_mst_employee.public_int_mode = 101
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_browse_item_Click()
frm_mst_salary_item.public_int_mode = 11
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
CnG.Execute "delete from t_salary_item where employee_code = '" _
    & Adodc1.Recordset.Fields("employee_code").Value & "' and date = '" _
    & Format(Adodc1.Recordset.Fields("date").Value, "yyyy-mm-dd hh:nn:ss") & "' and company_code='" _
    & TDBCombo_company.Columns("company_code").Value & "' and salary_item_code='" _
    & Adodc1.Recordset.Fields("salary_item_code").Value & "'"

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
    txt_item_code = .Fields("salary_item_code").Value
    txt_employee_name = .Fields("employee_name").Value
    txt_item_name = .Fields("salary_item_name").Value
    
'    txt_flag_main_salary = .Fields("flag_main_salary").Value
'    txt_flag_sign = .Fields("flag_sign").Value
'    txt_flag_detail = .Fields("flag_detail").Value
'    txt_flag_per_presence = .Fields("flag_per_presence").Value
'    txt_default_value = .Fields("default_value").Value
    txt_salary_value = .Fields("salary_value").Value
'    txt_item_code = .Fields("item_code").Value
'    txt_item_name = .Fields("item_name").Value
'    txt_item_type = .Fields("item_type").Value
'    txt_item_type_name = .Fields("item_type_name").Value
    txt_description = .Fields("description").Value
End With

    v_value = txt_salary_value
    
End Sub

Private Sub cmdEdit_Click()
int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
int_mode = 1
Call load_mode
End Sub

Private Sub CmdPrint_Click()
TDBGrid1.PrintInfo.PageSetup
If Not TDBGrid1.PrintInfo.PageSetupCancelled = True Then
    TDBGrid1.PrintInfo.PrintPreview dbgAllRows
End If
End Sub

Private Sub insert_new_data()
Dim rs1 As New ADODB.Recordset

If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from t_salary_item where employee_code = '-77'", CnG, adOpenKeyset, adLockOptimistic

rs1.Open "select * from m_salary_item where salary_item_code = '" _
    & Trim(txt_item_code) & "'", CnG, adOpenStatic, adLockReadOnly

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
    .Fields("salary_item_code").Value = Trim(txt_item_code)
    '-------------------------------------------------------------------------
    .Fields("employee_name").Value = Trim(txt_employee_name)
    .Fields("salary_item_name").Value = Trim(txt_item_name)
    
    .Fields("flag_main_salary").Value = rs1.Fields("flag_main_salary").Value
    .Fields("flag_sign").Value = rs1.Fields("flag_sign").Value
    .Fields("flag_detail").Value = rs1.Fields("flag_detail").Value
    .Fields("flag_per_presence").Value = rs1.Fields("flag_per_presence").Value
    .Fields("default_value").Value = rs1.Fields("default_value").Value
    .Fields("salary_value").Value = Val(DropAllComma(txt_salary_value))
    
    .Fields("item_code").Value = rs1.Fields("item_code").Value
    .Fields("item_name").Value = rs1.Fields("item_name").Value
    .Fields("item_type").Value = rs1.Fields("item_type").Value
    .Fields("item_type_name").Value = rs1.Fields("item_type_name").Value
    .Fields("description").Value = Trim(txt_description)
    
    .Update
End With
CnG.CommitTrans
End Sub

Private Sub edit_old_data()
Dim rs1 As New ADODB.Recordset
'On Error GoTo err_capture

If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from t_salary_item where employee_code = '" _
& Adodc1.Recordset.Fields("employee_code").Value & "' and date = '" _
& Format(Adodc1.Recordset.Fields("date").Value, "yyyy-mm-dd hh:nn:ss") & "' and company_code='" _
& TDBCombo_company.Columns("company_code").Value & "' and salary_item_code='" _
& Adodc1.Recordset.Fields("salary_item_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

rs1.Open "select * from m_salary_item where salary_item_code = '" _
    & Trim(txt_item_code) & "'", CnG, adOpenStatic, adLockReadOnly

CnG.BeginTrans
'+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
If v_value <> txt_salary_value.Text Then
    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute strsql
End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

With rsBound

    .Fields("date").Value = DTPicker_date
    .Fields("company_code").Value = TDBCombo_company.Columns("company_code").Value
    .Fields("employee_code").Value = Trim(txt_employee_code)
    .Fields("salary_item_code").Value = Trim(txt_item_code)
    '-------------------------------------------------------------------------
    .Fields("employee_name").Value = Trim(txt_employee_name)
    .Fields("salary_item_name").Value = Trim(txt_item_name)
    
    .Fields("flag_main_salary").Value = rs1.Fields("flag_main_salary").Value
    .Fields("flag_sign").Value = rs1.Fields("flag_sign").Value
    .Fields("flag_detail").Value = rs1.Fields("flag_detail").Value
    .Fields("flag_per_presence").Value = rs1.Fields("flag_per_presence").Value
    .Fields("default_value").Value = rs1.Fields("default_value").Value
    .Fields("salary_value").Value = Val(DropAllComma(txt_salary_value))
    
    .Fields("item_code").Value = rs1.Fields("item_code").Value
    .Fields("item_name").Value = rs1.Fields("item_name").Value
    .Fields("item_type").Value = rs1.Fields("item_type").Value
    .Fields("item_type_name").Value = rs1.Fields("item_type_name").Value
    .Fields("description").Value = Trim(txt_description)

    .Update
End With
CnG.CommitTrans

Exit Sub
err_capture:
MsgBox err.Description
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
        txt_employee_code.SetFocus
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
End Sub

Private Sub load_data_salary()
Adodc1.RecordSource = "select * from t_salary_item where company_code = '" _
    & TDBCombo_company.Columns("company_code").Value & "' and flag_sign=1 " _
    & "order by employee_code asc, salary_item_code asc, date desc"
Adodc1.Refresh

'cmdEdit.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'cmdDelete.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'CmdNew.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)

TDBGrid1.DataSource = Adodc1
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

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub txt_department_code_Change()

End Sub

Private Sub txt_salary_Validate(Cancel As Boolean)
'If Not Trim(txt_salary) = "" Then
'    txt_salary = FormatNumber(DropAllComma(txt_salary))
'End If
End Sub
