VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_salary_slip 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER SALARY SLIP"
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
   Icon            =   "frm_mst_salary_slip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_entry 
      Height          =   4335
      Left            =   240
      TabIndex        =   28
      Top             =   1080
      Width           =   14175
      Begin VB.TextBox txt_tax_salary_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10080
         MaxLength       =   10
         TabIndex        =   13
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CheckBox chk_flag_tax 
         Height          =   255
         Left            =   9720
         TabIndex        =   12
         Top             =   3120
         Width           =   495
      End
      Begin VB.CheckBox chk_flag_visible 
         Height          =   255
         Left            =   9720
         TabIndex        =   14
         Top             =   3480
         Width           =   495
      End
      Begin VB.CheckBox chk_flag_pph21 
         Height          =   255
         Left            =   9720
         TabIndex        =   10
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txt_pph21_number 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9720
         MaxLength       =   10
         TabIndex        =   11
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CheckBox chk_flag_ptkp 
         Height          =   255
         Left            =   9720
         TabIndex        =   8
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmd_browse_formula 
         Caption         =   "..."
         Height          =   255
         Left            =   2760
         TabIndex        =   6
         ToolTipText     =   "Browse Formula..."
         Top             =   2760
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chk_flag_use_formula 
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox txt_salary_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   10
         TabIndex        =   0
         Top             =   600
         Width           =   1575
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
      Begin VB.TextBox txt_salary_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin VB.CheckBox chk_flag_main 
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   1680
         Width           =   495
      End
      Begin VB.CheckBox chk_flag_detail 
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   2040
         Width           =   495
      End
      Begin VB.OptionButton opt_plus 
         Caption         =   "+"
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   2400
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton opt_minus 
         Caption         =   "-"
         Height          =   195
         Left            =   3240
         TabIndex        =   5
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txt_ptkp_salary_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9720
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "TAX"
         Height          =   195
         Left            =   8160
         TabIndex        =   41
         Top             =   3120
         Width           =   285
      End
      Begin VB.Line Line1 
         X1              =   960
         X2              =   11400
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "PPH21"
         Height          =   195
         Left            =   8160
         TabIndex        =   40
         Top             =   2400
         Width           =   465
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "PPH21 NUMBER"
         Height          =   195
         Left            =   8160
         TabIndex        =   39
         Top             =   2760
         Width           =   1125
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "P T K P"
         Height          =   195
         Left            =   8160
         TabIndex        =   38
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "FORMULA"
         Height          =   195
         Left            =   960
         TabIndex        =   37
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "SALARY NAME"
         Height          =   195
         Left            =   960
         TabIndex        =   36
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "SALARY CODE"
         Height          =   195
         Left            =   960
         TabIndex        =   35
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SIGN"
         Height          =   195
         Left            =   960
         TabIndex        =   34
         Top             =   2400
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "MAIN SALARY"
         Height          =   195
         Left            =   960
         TabIndex        =   33
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DETAIL"
         Height          =   195
         Left            =   960
         TabIndex        =   32
         Top             =   2040
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "VISIBLE"
         Height          =   195
         Left            =   8160
         TabIndex        =   31
         Top             =   3480
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "PTKP SALARY CODE"
         Height          =   195
         Left            =   8160
         TabIndex        =   30
         Top             =   2040
         Width           =   1440
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   26
      Top             =   7560
      Width           =   14175
      Begin VB.CommandButton cmd_select 
         Caption         =   "&Select"
         Height          =   645
         Left            =   11040
         Picture         =   "frm_mst_salary_slip.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmd_update 
         Caption         =   "&Update"
         Height          =   645
         Left            =   9240
         Picture         =   "frm_mst_salary_slip.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_delete_detail 
         Caption         =   "&Del Dtl"
         Height          =   645
         Left            =   8160
         Picture         =   "frm_mst_salary_slip.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmd_add_detail 
         Caption         =   "&Add Dtl"
         Height          =   645
         Left            =   7080
         Picture         =   "frm_mst_salary_slip.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   1560
         Picture         =   "frm_mst_salary_slip.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   4800
         Picture         =   "frm_mst_salary_slip.frx":1BBE
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   12600
         Picture         =   "frm_mst_salary_slip.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   480
         Picture         =   "frm_mst_salary_slip.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   3720
         Picture         =   "frm_mst_salary_slip.frx":2C5C
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   2640
         Picture         =   "frm_mst_salary_slip.frx":31E6
         Style           =   1  'Graphical
         TabIndex        =   17
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2280
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   5055
      Left            =   240
      TabIndex        =   27
      Top             =   360
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   8916
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "SALARY CODE"
      Columns(0).DataField=   "salary_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "SALARY NAME"
      Columns(1).DataField=   "salary_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   4
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "MAIN"
      Columns(2).DataField=   "flag_main_salary"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   4
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DETAIL"
      Columns(3).DataField=   "flag_detail"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   16
      Columns(4)._MaxComboItems=   5
      Columns(4).ValueItems(0)._DefaultItem=   0
      Columns(4).ValueItems(0).Value=   "1"
      Columns(4).ValueItems(0).Value.vt=   8
      Columns(4).ValueItems(0).DisplayValue=   "+"
      Columns(4).ValueItems(0).DisplayValue.vt=   8
      Columns(4).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems(1)._DefaultItem=   0
      Columns(4).ValueItems(1).Value=   "-1"
      Columns(4).ValueItems(1).Value.vt=   8
      Columns(4).ValueItems(1).DisplayValue=   "-"
      Columns(4).ValueItems(1).DisplayValue.vt=   8
      Columns(4).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(4).ValueItems.Count=   2
      Columns(4).Caption=   "SIGN"
      Columns(4).DataField=   "flag_sign"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "FORMULA"
      Columns(5).DataField=   "flag_use_formula"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "FN"
      Columns(6).DataField=   "f_detail"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   4
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "PTKP"
      Columns(7).DataField=   "flag_ptkp"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "PTKP SALARY"
      Columns(8).DataField=   "ptkp_salary_code"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   4
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "PPH21"
      Columns(9).DataField=   "flag_pph21"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "PPH21 NO."
      Columns(10).DataField=   "pph21_number"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   4
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "VISIBLE"
      Columns(11).DataField=   "flag_visible"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2619"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2540"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8708"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=7011"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6932"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8708"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1323"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1244"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8705"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=1270"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1191"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8705"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1217"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1138"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8705"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1588"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1508"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(5)._MinWidth=131022288"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=688"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=609"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(6).Button=1"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(6).ButtonAlways=1"
      Splits(0)._ColumnProps(39)=   "Column(6)._MinWidth=131059520"
      Splits(0)._ColumnProps(40)=   "Column(7).Width=1111"
      Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=1032"
      Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=8705"
      Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(45)=   "Column(7)._MinWidth=131068464"
      Splits(0)._ColumnProps(46)=   "Column(8).Width=2434"
      Splits(0)._ColumnProps(47)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(8)._WidthInPix=2355"
      Splits(0)._ColumnProps(49)=   "Column(8)._ColStyle=8708"
      Splits(0)._ColumnProps(50)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(51)=   "Column(8)._MinWidth=131053248"
      Splits(0)._ColumnProps(52)=   "Column(9).Width=1191"
      Splits(0)._ColumnProps(53)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(9)._WidthInPix=1111"
      Splits(0)._ColumnProps(55)=   "Column(9)._ColStyle=8705"
      Splits(0)._ColumnProps(56)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(57)=   "Column(9)._MinWidth=131052256"
      Splits(0)._ColumnProps(58)=   "Column(10).Width=1852"
      Splits(0)._ColumnProps(59)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(10)._WidthInPix=1773"
      Splits(0)._ColumnProps(61)=   "Column(10)._ColStyle=8705"
      Splits(0)._ColumnProps(62)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(63)=   "Column(10)._MinWidth=131061680"
      Splits(0)._ColumnProps(64)=   "Column(11).Width=1614"
      Splits(0)._ColumnProps(65)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(11)._WidthInPix=1535"
      Splits(0)._ColumnProps(67)=   "Column(11)._ColStyle=8705"
      Splits(0)._ColumnProps(68)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(69)=   "Column(11)._MinWidth=131062720"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowDelete     =   -1  'True
      AllowUpdate     =   0   'False
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "LIST OF SALARY SLIP"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.locked=-1"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=70,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=74,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=78,.parent=13,.locked=-1"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=82,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=86,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=83,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=84,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=85,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=90,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=87,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=88,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=89,.parent=17"
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Height          =   1815
      Left            =   240
      TabIndex        =   25
      Top             =   5640
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   3201
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ITEM CODE"
      Columns(0).DataField=   "salary_item_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "ITEM NAME"
      Columns(1).DataField=   "salary_item_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TYPE"
      Columns(2).DataField=   "item_name"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "STATUS"
      Columns(3).DataField=   "item_type_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   4
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "MAIN"
      Columns(4).DataField=   "flag_main_salary"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DETAIL"
      Columns(5).DataField=   "flag_detail"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   4
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "PRESENCE"
      Columns(6).DataField=   "flag_per_presence"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   16
      Columns(7)._MaxComboItems=   5
      Columns(7).ValueItems(0)._DefaultItem=   0
      Columns(7).ValueItems(0).Value=   "1"
      Columns(7).ValueItems(0).Value.vt=   8
      Columns(7).ValueItems(0).DisplayValue=   "+"
      Columns(7).ValueItems(0).DisplayValue.vt=   8
      Columns(7).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems(1)._DefaultItem=   0
      Columns(7).ValueItems(1).Value=   "-1"
      Columns(7).ValueItems(1).Value.vt=   8
      Columns(7).ValueItems(1).DisplayValue=   "-"
      Columns(7).ValueItems(1).DisplayValue.vt=   8
      Columns(7).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(7).ValueItems.Count=   2
      Columns(7).Caption=   "SIGN"
      Columns(7).DataField=   "flag_sign"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "DEFAULT VALUE"
      Columns(8).DataField=   "default_value"
      Columns(8).NumberFormat=   "Standard"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3201"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3122"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=7303"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=7223"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3228"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3149"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2275"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2196"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=1323"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1244"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=1270"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1191"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=1614"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=1535"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=1217"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=1138"
      Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=513"
      Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(41)=   "Column(8).Width=2514"
      Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=2434"
      Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(46)=   "Column(8)._MinWidth=131022288"
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
      Caption         =   "LIST OF SALARY ITEM"
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
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=66,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
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
End
Attribute VB_Name = "frm_mst_salary_slip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBound As New ADODB.Recordset
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns



Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(salary_code) as rec_count from m_salary_slip where salary_code = '" _
& Trim(txt_salary_code) & "'"
rs.Open str_sql, CnG, adOpenStatic, adLockBatchOptimistic

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data Sudah Ada...", vbCritical, headerMSG
txt_salary_code = ""
If txt_salary_code.Enabled = True Then txt_salary_code.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not (txt_salary_code = Adodc1.Recordset.Fields("salary_code").Value) And _
    check_validate_exist_new Then
    
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

If Trim(txt_salary_code) = "" Then
    MsgBox "Salary Masih Kosong...", vbInformation, headerMSG
    txt_salary_code.SetFocus
    check_validate_new = False
    Exit Function
End If
End Function

Private Sub load_data_header()
Adodc1.RecordSource = "select * from m_salary_slip order by salary_code"
Adodc1.Refresh

TDBGrid1.DataSource = Adodc1
End Sub

Private Sub load_data_detail()
Adodc2.RecordSource = "select * from d_salary_slip " _
& "where salary_code = '" & Adodc1.Recordset.Fields("salary_code").Value & "' " _
& "order by sequence_number"
Adodc2.Refresh

TDBGrid2.DataSource = Adodc2
End Sub

Private Sub cmd_browse_formula_Click()
'frm_lookup_mst_book.public_int_mode = 0
'frm_lookup_mst_book.Show 1
End Sub

Private Sub cmd_add_detail_Click()
frm_mst_salary_item.public_int_mode = 0
frm_mst_salary_item.cmd_Select.Visible = True
frm_mst_salary_item.Show 1
End Sub

Private Sub cmd_delete_detail_Click()
Adodc2.Recordset.Delete adAffectCurrent
End Sub

Private Sub cmd_Select_Click()
On Error GoTo err_capture

Dim i As Integer
Dim item

If Not TDBGrid1.ApproxCount > 0 Then
    Exit Sub
End If


'If public_int_mode = 0 Then
'    With frm_mst_salary_slip_formula
'        .txt_formula_salary_code = Adodc1.Recordset("salary_code").Value
'        .txt_formula_salary_name = Adodc1.Recordset("salary_name").Value
'    End With
'
'End If

Call CmdExit_Click

Exit Sub
err_capture:
MsgBox "Gagal Memilih Data...", vbInformation, headerMSG
End Sub

Private Sub cmd_update_Click()
TDBGrid1.Update
TDBGrid2.Update
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
    & Adodc1.Recordset.Fields("salary_name").Value & " ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from d_salary_slip where salary_code = '" _
    & Adodc1.Recordset.Fields("salary_code").Value & "'"
CnG.Execute "delete from m_salary_slip where salary_code = '" _
    & Adodc1.Recordset.Fields("salary_code").Value & "'"
CnG.CommitTrans

Call load_data_header
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    
    txt_salary_code = .Fields("salary_code").Value
    '-------------------------------------------------------------------
    txt_salary_name = .Fields("salary_name").Value
    chk_flag_main = .Fields("flag_main_salary").Value
    If .Fields("flag_sign").Value = 1 Then
        opt_plus = True
    Else
        opt_minus = True
    End If
    chk_flag_detail = .Fields("flag_detail").Value
    chk_flag_use_formula = .Fields("flag_use_formula").Value
'    txt_formula_salary_code = .Fields("formula_salary_code").Value
    chk_flag_ptkp = .Fields("flag_ptkp").Value
    txt_ptkp_salary_code = .Fields("ptkp_salary_code").Value
    chk_flag_pph21 = .Fields("flag_pph21").Value
    txt_pph21_number = .Fields("pph21_number").Value
    
    chk_flag_tax = Val("" & .Fields("flag_tax").Value)
    txt_tax_salary_code = Trim("" & .Fields("tax_salary_code").Value)
    
    chk_flag_visible = .Fields("flag_visible").Value
'    txt_description = .Fields("description").Value

End With
End Sub

Private Sub cmdEdit_Click()
int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
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
Call load_data_header
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
Dim rs1 As New ADODB.Recordset

If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_salary_slip where salary_code = 'άφ'", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans

With rsBound
    .AddNew
    
    .Fields("salary_code").Value = Trim(txt_salary_code)
    '--------------------------------------------------------------------------
    .Fields("salary_name").Value = Trim(txt_salary_name)
    .Fields("flag_main_salary").Value = IIf(chk_flag_main.Value = vbChecked, 1, 0)
    .Fields("flag_sign").Value = IIf(opt_plus, 1, -1)
    .Fields("flag_detail").Value = IIf(chk_flag_detail.Value = vbChecked, 1, 0)
    
    .Fields("flag_use_formula").Value = IIf(chk_flag_use_formula.Value = vbChecked, 1, 0)
'    .Fields("formula_salary_code").Value = Trim(txt_formula_salary_code)
    .Fields("flag_ptkp").Value = IIf(chk_flag_ptkp.Value = vbChecked, 1, 0)
    .Fields("ptkp_salary_code").Value = Trim(txt_ptkp_salary_code)
    .Fields("flag_pph21").Value = IIf(chk_flag_pph21.Value = vbChecked, 1, 0)
    .Fields("pph21_number").Value = Val(txt_pph21_number)
    
    .Fields("flag_tax").Value = IIf(chk_flag_tax.Value = vbChecked, 1, 0)
    .Fields("tax_salary_code").Value = Trim(txt_tax_salary_code)
    
    .Fields("flag_visible").Value = IIf(chk_flag_visible.Value = vbChecked, 1, 0)
'    .Fields("description").Value = Trim(txt_description)

    .Update
End With

CnG.CommitTrans
End Sub

Private Sub edit_old_data()
On Error GoTo err_capture
Dim rs1 As New ADODB.Recordset

If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_salary_slip where salary_code = '" _
& Adodc1.Recordset.Fields("salary_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans
With rsBound

   .Fields("salary_code").Value = Trim(txt_salary_code)
    '--------------------------------------------------------------------------
    .Fields("salary_name").Value = Trim(txt_salary_name)
    .Fields("flag_main_salary").Value = IIf(chk_flag_main.Value = vbChecked, 1, 0)
    .Fields("flag_sign").Value = IIf(opt_plus, 1, -1)
    .Fields("flag_detail").Value = IIf(chk_flag_detail.Value = vbChecked, 1, 0)
    
    .Fields("flag_use_formula").Value = IIf(chk_flag_use_formula.Value = vbChecked, 1, 0)
'    .Fields("formula_salary_code").Value = Trim(txt_formula_salary_code)
    .Fields("flag_ptkp").Value = IIf(chk_flag_ptkp.Value = vbChecked, 1, 0)
    .Fields("ptkp_salary_code").Value = Trim(txt_ptkp_salary_code)
    .Fields("flag_pph21").Value = IIf(chk_flag_pph21.Value = vbChecked, 1, 0)
    .Fields("pph21_number").Value = Trim(txt_pph21_number)
    
    .Fields("flag_tax").Value = IIf(chk_flag_tax.Value = vbChecked, 1, 0)
    .Fields("tax_salary_code").Value = Trim(txt_tax_salary_code)
    
    .Fields("flag_visible").Value = IIf(chk_flag_visible.Value = vbChecked, 1, 0)
'    .Fields("description").Value = Trim(txt_description)
    
    .Update
End With

CnG.Execute "update m_salary_slip_formula set salary_code = '" & Trim(txt_salary_code) & "' " _
    & "where salary_code = '" & Adodc1.Recordset.Fields("salary_code").Value & "'"

CnG.Execute "update d_salary_slip set salary_code = '" & Trim(txt_salary_code) & "' " _
    & "where salary_code = '" & Adodc1.Recordset.Fields("salary_code").Value & "'"

CnG.CommitTrans

Exit Sub
err_capture:
rsBound.CancelBatch adAffectCurrent: rsBound.Close: CnG.RollbackTrans
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

Call load_data_header
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

'CmdPrint.Enabled = f
'cmd_refresh.Enabled = g
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
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
'cbo_requirement.ListIndex = 0
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
    txt_salary_code.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
    If txt_salary_code.Enabled = True Then
        txt_salary_code.SetFocus
    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
'    txt_salary_code.Enabled = False
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
Adodc1.ConnectionString = strConn
Adodc2.ConnectionString = strConn

Call load_data_header

'Call LoadData_UserAccess(Me.Caption)
int_mode = 0
Call load_mode
End Sub

Private Sub TDBGrid1_ButtonClick(ByVal ColIndex As Integer)
If TDBGrid1.Columns(ColIndex).Caption = "FN" And TDBGrid1.Columns("flag_use_formula").Value = 1 Then
    frm_mst_salary_slip_formula.Show 1
End If
End Sub

Private Sub TDBGrid1_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid1, Adodc1)
If TDBGrid1.ApproxCount > 0 Then
    Call load_data_detail
End If
End Sub
Private Sub TDBGrid2_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid2, Adodc2)
End Sub

Private Sub txt_sks_value_KeyPress(KeyAscii As Integer)
'Call NumericFilter(KeyAscii)
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If (TDBGrid1.Row + 1) > 0 And (TDBGrid1.Row + 1) <> LastRow Then
    'MsgBox "LETS..."
    Call load_data_detail
End If
End Sub


