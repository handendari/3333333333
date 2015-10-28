VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_loan 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "EMPLOYEE INSTALLMENT"
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
   Icon            =   "frm_trans_loan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7200
      TabIndex        =   31
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
      TabIndex        =   27
      Top             =   240
      Width           =   3855
   End
   Begin VB.Frame fra_entry 
      Height          =   4215
      Left            =   240
      TabIndex        =   0
      Top             =   3240
      Width           =   14175
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
         Left            =   10320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1800
         Width           =   3015
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
         Left            =   10320
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         Top             =   3240
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker DTPicker_instalment_start 
         Height          =   315
         Left            =   10320
         TabIndex        =   9
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   151453699
         CurrentDate     =   39270
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Left            =   9960
         TabIndex        =   38
         Top             =   2760
         Width           =   2775
         Begin MSComCtl2.DTPicker DTPicker_instalment_end 
            Height          =   315
            Left            =   360
            TabIndex        =   10
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   151453699
            CurrentDate     =   39270
         End
      End
      Begin VB.TextBox txt_instalment_times 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10320
         MaxLength       =   10
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txt_interest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10320
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
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
         Left            =   10320
         MaxLength       =   50
         TabIndex        =   5
         Top             =   600
         Width           =   3015
      End
      Begin VB.CommandButton cmd_browse 
         Caption         =   "..."
         Height          =   320
         Left            =   3720
         TabIndex        =   2
         ToolTipText     =   "Browse employee data..."
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   1635
         Left            =   2280
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         Top             =   960
         Width           =   4335
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
      Begin VB.TextBox txt_employee_code 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2280
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   19
         Top             =   600
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker_loan_date 
         Height          =   315
         Left            =   2280
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   101842947
         CurrentDate     =   39270
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "LOAN TOTAL (IDR)"
         Height          =   195
         Left            =   8040
         TabIndex        =   39
         Top             =   1800
         Width           =   1365
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "INSTALMENT END"
         Height          =   195
         Left            =   8040
         TabIndex        =   37
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "INSTALMENT TIMES"
         Height          =   195
         Left            =   8040
         TabIndex        =   36
         Top             =   1440
         Width           =   1425
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FLAT INTEREST/MONTH (%)"
         Height          =   195
         Left            =   8040
         TabIndex        =   35
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "INSTALMENT START"
         Height          =   195
         Left            =   8040
         TabIndex        =   34
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "INSTALMENT VALUE (IDR)"
         Height          =   195
         Left            =   8040
         TabIndex        =   33
         Top             =   3240
         Width           =   1875
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "LOAN VALUE (IDR)"
         Height          =   195
         Left            =   8040
         TabIndex        =   30
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DATE"
         Height          =   195
         Left            =   840
         TabIndex        =   29
         Top             =   1320
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   840
         TabIndex        =   26
         Top             =   1680
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EMP. CODE"
         Height          =   195
         Left            =   840
         TabIndex        =   25
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "EMP. NAME"
         Height          =   195
         Left            =   840
         TabIndex        =   24
         Top             =   960
         Width           =   825
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   21
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
         Left            =   7440
         Picture         =   "frm_trans_loan.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   1800
         Picture         =   "frm_trans_loan.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5040
         Picture         =   "frm_trans_loan.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   9600
         Picture         =   "frm_trans_loan.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   720
         Picture         =   "frm_trans_loan.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   0
         Picture         =   "frm_trans_loan.frx":1BBE
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   3960
         Picture         =   "frm_trans_loan.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   2880
         Picture         =   "frm_trans_loan.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_trans_loan.frx":2C5C
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1200
      Top             =   360
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6735
      Left            =   240
      TabIndex        =   32
      Top             =   720
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11880
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "EMP. CODE"
      Columns(0).DataField=   "employee_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EMP. NAME"
      Columns(1).DataField=   "employee_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "DATE"
      Columns(2).DataField=   "subsidy_date"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "LOAN (IDR)"
      Columns(3).DataField=   "loan_value"
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "INTEREST (%)"
      Columns(4).DataField=   "loan_interest"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "LOAN TOTAL (IDR)"
      Columns(5).DataField=   "loan_total"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "INSTALMENT TIMES"
      Columns(6).DataField=   "instalment_time"
      Columns(6).NumberFormat=   "General Number"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "INS. START"
      Columns(7).DataField=   "instalment_start"
      Columns(7).NumberFormat=   "FormatText Event"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "INS. END"
      Columns(8).DataField=   "instalment_end"
      Columns(8).NumberFormat=   "FormatText Event"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   16
      Columns(9)._MaxComboItems=   5
      Columns(9).ValueItems(0)._DefaultItem=   0
      Columns(9).ValueItems(0).Value=   "0"
      Columns(9).ValueItems(0).Value.vt=   8
      Columns(9).ValueItems(0).DisplayValue=   "Yearly"
      Columns(9).ValueItems(0).DisplayValue.vt=   8
      Columns(9).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(9).ValueItems.Count=   1
      Columns(9).Caption=   "INSTALMENT (IDR)"
      Columns(9).DataField=   "instalment_value"
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   1
      Splits(0).Size  =   4004.788
      Splits(0).Size.vt=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=4260"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4180"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1958"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1879"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2143"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2064"
      Splits(0)._ColumnProps(19)=   "Column(3).AllowSizing=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(3).AllowFocus=0"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2275"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2196"
      Splits(0)._ColumnProps(27)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(29)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(30)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(4)._MinWidth=10"
      Splits(0)._ColumnProps(33)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(5)._MinWidth=54215968"
      Splits(0)._ColumnProps(39)=   "Column(6).Width=2275"
      Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=2196"
      Splits(0)._ColumnProps(42)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(44)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(45)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(47)=   "Column(6)._MinWidth=54215968"
      Splits(0)._ColumnProps(48)=   "Column(7).Width=2461"
      Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=2381"
      Splits(0)._ColumnProps(51)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(53)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(54)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(55)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(56)=   "Column(8).Width=2275"
      Splits(0)._ColumnProps(57)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(8)._WidthInPix=2196"
      Splits(0)._ColumnProps(59)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(60)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(61)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(62)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(63)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(64)=   "Column(9).Width=1588"
      Splits(0)._ColumnProps(65)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(9)._WidthInPix=1508"
      Splits(0)._ColumnProps(67)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(68)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(69)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(70)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(71)=   "Column(9).Order=10"
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
      Splits(1)._ColumnProps(0)=   "Columns.Count=10"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=2461"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2381"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(0)._MinWidth=83786356"
      Splits(1)._ColumnProps(10)=   "Column(1).Width=3810"
      Splits(1)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(12)=   "Column(1)._WidthInPix=3731"
      Splits(1)._ColumnProps(13)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(14)=   "Column(1)._ColStyle=516"
      Splits(1)._ColumnProps(15)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(16)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(17)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(18)=   "Column(1)._MinWidth=1967288933"
      Splits(1)._ColumnProps(19)=   "Column(2).Width=1958"
      Splits(1)._ColumnProps(20)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._WidthInPix=1879"
      Splits(1)._ColumnProps(22)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(23)=   "Column(2)._ColStyle=516"
      Splits(1)._ColumnProps(24)=   "Column(2).Visible=0"
      Splits(1)._ColumnProps(25)=   "Column(2).AllowFocus=0"
      Splits(1)._ColumnProps(26)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(27)=   "Column(3).Width=2831"
      Splits(1)._ColumnProps(28)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(29)=   "Column(3)._WidthInPix=2752"
      Splits(1)._ColumnProps(30)=   "Column(3)._ColStyle=514"
      Splits(1)._ColumnProps(31)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(32)=   "Column(3)._MinWidth=25972"
      Splits(1)._ColumnProps(33)=   "Column(4).Width=2090"
      Splits(1)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(4)._WidthInPix=2011"
      Splits(1)._ColumnProps(36)=   "Column(4)._ColStyle=513"
      Splits(1)._ColumnProps(37)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(38)=   "Column(4)._MinWidth=71305344"
      Splits(1)._ColumnProps(39)=   "Column(5).Width=2725"
      Splits(1)._ColumnProps(40)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(41)=   "Column(5)._WidthInPix=2646"
      Splits(1)._ColumnProps(42)=   "Column(5)._ColStyle=514"
      Splits(1)._ColumnProps(43)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(44)=   "Column(5)._MinWidth=54215968"
      Splits(1)._ColumnProps(45)=   "Column(6).Width=2805"
      Splits(1)._ColumnProps(46)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(47)=   "Column(6)._WidthInPix=2725"
      Splits(1)._ColumnProps(48)=   "Column(6)._ColStyle=513"
      Splits(1)._ColumnProps(49)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(50)=   "Column(6)._MinWidth=54215968"
      Splits(1)._ColumnProps(51)=   "Column(7).Width=2117"
      Splits(1)._ColumnProps(52)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(53)=   "Column(7)._WidthInPix=2037"
      Splits(1)._ColumnProps(54)=   "Column(7)._ColStyle=513"
      Splits(1)._ColumnProps(55)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(56)=   "Column(8).Width=2037"
      Splits(1)._ColumnProps(57)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(58)=   "Column(8)._WidthInPix=1958"
      Splits(1)._ColumnProps(59)=   "Column(8)._ColStyle=513"
      Splits(1)._ColumnProps(60)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(61)=   "Column(9).Width=2884"
      Splits(1)._ColumnProps(62)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(63)=   "Column(9)._WidthInPix=2805"
      Splits(1)._ColumnProps(64)=   "Column(9)._ColStyle=514"
      Splits(1)._ColumnProps(65)=   "Column(9).Order=10"
      Splits.Count    =   2
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
      Caption         =   "LIST OF EMPLOYEE LOAN"
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
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=102,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=99,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=100,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=101,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=154,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=151,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=152,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=153,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=78,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=17"
      _StyleDefs(74)  =   "Splits(1).Style:id=79,.parent=1"
      _StyleDefs(75)  =   "Splits(1).CaptionStyle:id=88,.parent=4,.bgcolor=&H80000002&,.fgcolor=&H80000009&"
      _StyleDefs(76)  =   "Splits(1).HeadingStyle:id=80,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(77)  =   ":id=80,.fgcolor=&H80000002&"
      _StyleDefs(78)  =   "Splits(1).FooterStyle:id=81,.parent=3"
      _StyleDefs(79)  =   "Splits(1).InactiveStyle:id=82,.parent=5"
      _StyleDefs(80)  =   "Splits(1).SelectedStyle:id=84,.parent=6"
      _StyleDefs(81)  =   "Splits(1).EditorStyle:id=83,.parent=7"
      _StyleDefs(82)  =   "Splits(1).HighlightRowStyle:id=85,.parent=8"
      _StyleDefs(83)  =   "Splits(1).EvenRowStyle:id=86,.parent=9"
      _StyleDefs(84)  =   "Splits(1).OddRowStyle:id=87,.parent=10"
      _StyleDefs(85)  =   "Splits(1).RecordSelectorStyle:id=89,.parent=11"
      _StyleDefs(86)  =   "Splits(1).FilterBarStyle:id=90,.parent=12"
      _StyleDefs(87)  =   "Splits(1).Columns(0).Style:id=94,.parent=79"
      _StyleDefs(88)  =   "Splits(1).Columns(0).HeadingStyle:id=91,.parent=80"
      _StyleDefs(89)  =   "Splits(1).Columns(0).FooterStyle:id=92,.parent=81"
      _StyleDefs(90)  =   "Splits(1).Columns(0).EditorStyle:id=93,.parent=83"
      _StyleDefs(91)  =   "Splits(1).Columns(1).Style:id=98,.parent=79"
      _StyleDefs(92)  =   "Splits(1).Columns(1).HeadingStyle:id=95,.parent=80"
      _StyleDefs(93)  =   "Splits(1).Columns(1).FooterStyle:id=96,.parent=81"
      _StyleDefs(94)  =   "Splits(1).Columns(1).EditorStyle:id=97,.parent=83"
      _StyleDefs(95)  =   "Splits(1).Columns(2).Style:id=106,.parent=79"
      _StyleDefs(96)  =   "Splits(1).Columns(2).HeadingStyle:id=103,.parent=80"
      _StyleDefs(97)  =   "Splits(1).Columns(2).FooterStyle:id=104,.parent=81"
      _StyleDefs(98)  =   "Splits(1).Columns(2).EditorStyle:id=105,.parent=83"
      _StyleDefs(99)  =   "Splits(1).Columns(3).Style:id=114,.parent=79,.alignment=1"
      _StyleDefs(100) =   "Splits(1).Columns(3).HeadingStyle:id=111,.parent=80"
      _StyleDefs(101) =   "Splits(1).Columns(3).FooterStyle:id=112,.parent=81"
      _StyleDefs(102) =   "Splits(1).Columns(3).EditorStyle:id=113,.parent=83"
      _StyleDefs(103) =   "Splits(1).Columns(4).Style:id=122,.parent=79,.alignment=2"
      _StyleDefs(104) =   "Splits(1).Columns(4).HeadingStyle:id=119,.parent=80"
      _StyleDefs(105) =   "Splits(1).Columns(4).FooterStyle:id=120,.parent=81"
      _StyleDefs(106) =   "Splits(1).Columns(4).EditorStyle:id=121,.parent=83"
      _StyleDefs(107) =   "Splits(1).Columns(5).Style:id=158,.parent=79,.alignment=1"
      _StyleDefs(108) =   "Splits(1).Columns(5).HeadingStyle:id=155,.parent=80"
      _StyleDefs(109) =   "Splits(1).Columns(5).FooterStyle:id=156,.parent=81"
      _StyleDefs(110) =   "Splits(1).Columns(5).EditorStyle:id=157,.parent=83"
      _StyleDefs(111) =   "Splits(1).Columns(6).Style:id=130,.parent=79,.alignment=2"
      _StyleDefs(112) =   "Splits(1).Columns(6).HeadingStyle:id=127,.parent=80"
      _StyleDefs(113) =   "Splits(1).Columns(6).FooterStyle:id=128,.parent=81"
      _StyleDefs(114) =   "Splits(1).Columns(6).EditorStyle:id=129,.parent=83"
      _StyleDefs(115) =   "Splits(1).Columns(7).Style:id=138,.parent=79,.alignment=2"
      _StyleDefs(116) =   "Splits(1).Columns(7).HeadingStyle:id=135,.parent=80"
      _StyleDefs(117) =   "Splits(1).Columns(7).FooterStyle:id=136,.parent=81"
      _StyleDefs(118) =   "Splits(1).Columns(7).EditorStyle:id=137,.parent=83"
      _StyleDefs(119) =   "Splits(1).Columns(8).Style:id=146,.parent=79,.alignment=2"
      _StyleDefs(120) =   "Splits(1).Columns(8).HeadingStyle:id=143,.parent=80"
      _StyleDefs(121) =   "Splits(1).Columns(8).FooterStyle:id=144,.parent=81"
      _StyleDefs(122) =   "Splits(1).Columns(8).EditorStyle:id=145,.parent=83"
      _StyleDefs(123) =   "Splits(1).Columns(9).Style:id=150,.parent=79,.alignment=1"
      _StyleDefs(124) =   "Splits(1).Columns(9).HeadingStyle:id=147,.parent=80"
      _StyleDefs(125) =   "Splits(1).Columns(9).FooterStyle:id=148,.parent=81"
      _StyleDefs(126) =   "Splits(1).Columns(9).EditorStyle:id=149,.parent=83"
      _StyleDefs(127) =   "Named:id=33:Normal"
      _StyleDefs(128) =   ":id=33,.parent=0"
      _StyleDefs(129) =   "Named:id=34:Heading"
      _StyleDefs(130) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(131) =   ":id=34,.wraptext=-1"
      _StyleDefs(132) =   "Named:id=35:Footing"
      _StyleDefs(133) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(134) =   "Named:id=36:Selected"
      _StyleDefs(135) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(136) =   "Named:id=37:Caption"
      _StyleDefs(137) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(138) =   "Named:id=38:HighlightRow"
      _StyleDefs(139) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(140) =   "Named:id=39:EvenRow"
      _StyleDefs(141) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(142) =   "Named:id=40:OddRow"
      _StyleDefs(143) =   ":id=40,.parent=33"
      _StyleDefs(144) =   "Named:id=41:RecordSelector"
      _StyleDefs(145) =   ":id=41,.parent=34"
      _StyleDefs(146) =   "Named:id=42:FilterBar"
      _StyleDefs(147) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   28
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frm_trans_loan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Function check_validate_new() As Boolean
check_validate_new = True

'validasi employee code
If Trim(txt_employee_code) = "" Then
    MsgBox "Kode Karyawan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_employee_code.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi employee name
If Trim(txt_employee_name) = "" Then
    MsgBox "Nama Karyawan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_employee_name.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi loan value
If Trim(txt_loan_value) = "" Then
    MsgBox "Loan value Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_loan_value.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi interest
If Trim(txt_interest) = "" Then
    MsgBox "Interest Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_interest.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi instalment times
If Trim(txt_instalment_times) = "" Then
    MsgBox "Instalment times Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_instalment_times.SetFocus
    check_validate_new = False
    Exit Function
End If

'validasi description
If Trim(txt_description) = "" Then
    MsgBox "Description Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_description.SetFocus
    check_validate_new = False
    Exit Function
End If
End Function

Private Sub load_data()
timer1.Enabled = True
End Sub

'Private Sub date_event()
'If cbo_date_to.ListIndex = 0 Then
'    DTPicker_date_to.Visible = False
'Else
'    DTPicker_date_to.Visible = True
'    DTPicker_date_to.Value = DTPicker_date_from.Value
'End If
'End Sub
'
'Private Sub cbo_date_to_Click()
'Call date_event
'End Sub

Private Sub cmd_browse_Click()
frm_lookup_mst_employee.public_int_mode = 80
frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_refresh_Click()
'CnG.Execute "call spg_subsidy('" & TDBCombo_company.Columns("company_code").Value _
            & "')"
Call load_data_loan
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
    & TDBGrid1.Columns("employee_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from t_subsidy where employee_code = '" _
    & TDBGrid1.Columns("employee_code").Value & "'"
CnG.CommitTrans

Call load_data_loan
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    
    txt_employee_code = .Fields("employee_code").Value
    '-----------------------------------------------------------------------------
    txt_employee_name = .Fields("employee_name").Value
    
    DTPicker_loan_date = .Fields("loan_date").Value
    txt_loan_value = FormatNumber(.Fields("loan_value").Value)
    txt_interest = FormatNumber(.Fields("loan_interest").Value)
    txt_loan_total = FormatNumber(.Fields("loan_total").Value)
    
    txt_instalment_times = .Fields("instalment_time").Value
    DTPicker_instalment_start = .Fields("instalment_start").Value
    DTPicker_instalment_end = .Fields("instalment_end").Value
    txt_instalment_value = FormatNumber(.Fields("instalment_value").Value)
    
    txt_description = .Fields("description").Value
    
End With
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from t_loan where employee_code = '" _
& Adodc1.Recordset.Fields("employee_code").Value & "' and loan_date = '" _
& Format(Adodc1.Recordset.Fields("loan_date").Value, "yyyy-mm-dd hh:nn:ss") & "'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from t_loan where employee_code = '-77'", CnG, adOpenKeyset, adLockOptimistic

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
    
    .Fields("employee_code").Value = txt_employee_code
    '-----------------------------------------------------------------------------
    .Fields("loan_date").Value = Format(DTPicker_loan_date.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("loan_value").Value = Val(DropAllComma(txt_loan_value))
    .Fields("loan_interest").Value = Val(DropAllComma(txt_interest))
    .Fields("loan_total").Value = Val(DropAllComma(txt_loan_total))
    
    .Fields("instalment_time").Value = Int(txt_instalment_times)
    .Fields("instalment_start").Value = Format(DTPicker_instalment_start.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("instalment_end").Value = Format(DTPicker_instalment_end.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("instalment_value").Value = Val(DropAllComma(txt_instalment_value))
    
    .Fields("description").Value = Trim(txt_description)
    
    .Update
End With
CnG.CommitTrans
End Sub

Private Sub edit_old_data()
On Error GoTo err_capture

CnG.BeginTrans
With rsBound

    '.Fields("employee_code").Value = txt_employee_code
    '-----------------------------------------------------------------------------
    .Fields("loan_date").Value = Format(DTPicker_loan_date.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("loan_value").Value = Val(DropAllComma(txt_loan_value))
    .Fields("loan_interest").Value = Val(DropAllComma(txt_interest))
    .Fields("loan_total").Value = Val(DropAllComma(txt_loan_total))
    
    .Fields("instalment_time").Value = Int(txt_instalment_times)
    .Fields("instalment_start").Value = Format(DTPicker_instalment_start.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("instalment_end").Value = Format(DTPicker_instalment_end.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("instalment_value").Value = Val(DropAllComma(txt_instalment_value))
    
    .Fields("description").Value = Trim(txt_description)

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

Call load_data_loan
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
txt_employee_code = ""
txt_employee_name = ""
DTPicker_loan_date.Value = Now
txt_description = ""

txt_loan_value = ""
txt_loan_total = ""
txt_interest = ""
txt_instalment_times = ""
DTPicker_instalment_start = Now
DTPicker_instalment_end = Now
txt_instalment_value = ""
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
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
MsgBox KeyAscii
End Sub

Private Sub txt_salary_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 44, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
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

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    Call load_data_loan
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

Private Sub load_data_loan()
Adodc1.RecordSource = "select * from v_t_loan where company_code = '" _
& TDBCombo_company.Columns("company_code").Value & "' " _
& "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
& "order by employee_code"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid1.Columns(ColIndex).Caption = "INS. START" Or _
TDBGrid1.Columns(ColIndex).Caption = "INS. END" Then
    Value = Format(Value, "yyyy-mm-dd")
End If
End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub







Private Sub txt_loan_value_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 44, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

Private Sub txt_interest_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
    Case 8, 40, 41, 43, 44, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
        Exit Sub
    Case Else
        KeyAscii = 0
End Select
End Sub

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
dt2 = DateAdd("m", Int(txt_instalment_times), dt1)

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
Dim i, j As Double

i = (Val(DropAllComma(txt_interest)) / 100) * Val(DropAllComma(txt_instalment_times)) _
                                            * Val(DropAllComma(txt_loan_value))
j = Val(DropAllComma(txt_loan_value))

txt_loan_total = FormatNumber(i + j)
End Sub

