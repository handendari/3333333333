VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_salary 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER SALARY"
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
   Icon            =   "frm_mst_salary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   7200
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   0
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
      TabIndex        =   22
      Top             =   240
      Width           =   3855
   End
   Begin VB.Frame fra_entry 
      Height          =   5475
      Left            =   240
      TabIndex        =   13
      Top             =   1980
      Width           =   14175
      Begin VB.TextBox txt_jstk_accident 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         TabIndex        =   51
         Top             =   2970
         Width           =   1575
      End
      Begin VB.TextBox txt_special 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         TabIndex        =   50
         Top             =   3690
         Width           =   1575
      End
      Begin VB.TextBox txt_jstk_life 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         TabIndex        =   49
         Top             =   3330
         Width           =   1575
      End
      Begin VB.TextBox txt_add_title 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   38
         Top             =   1380
         Width           =   1575
      End
      Begin VB.TextBox txt_add_work 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   37
         Top             =   1740
         Width           =   1575
      End
      Begin VB.TextBox txt_add_other 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   36
         Top             =   2100
         Width           =   1575
      End
      Begin VB.TextBox txt_add_hp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   35
         Top             =   2460
         Width           =   1575
      End
      Begin VB.TextBox txt_add_transport 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   34
         Top             =   2820
         Width           =   1575
      End
      Begin VB.TextBox txt_add_presence 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   33
         Top             =   3180
         Width           =   1575
      End
      Begin VB.TextBox txt_add_spl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   32
         Top             =   3540
         Width           =   1575
      End
      Begin VB.TextBox txt_add_food_daily 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   31
         Top             =   4260
         Width           =   1575
      End
      Begin VB.TextBox txt_main_salary_daily 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   30
         Top             =   3900
         Width           =   1575
      End
      Begin VB.TextBox txtdriver 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   8520
         MaxLength       =   10
         TabIndex        =   29
         Top             =   4620
         Width           =   1575
      End
      Begin VB.TextBox txt_over_time 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3300
         MaxLength       =   50
         TabIndex        =   27
         Top             =   4050
         Width           =   1575
      End
      Begin VB.TextBox txt_salary 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8520
         MaxLength       =   50
         TabIndex        =   4
         Top             =   930
         Width           =   3015
      End
      Begin VB.CommandButton cmd_browse 
         Caption         =   "..."
         Height          =   320
         Left            =   3480
         TabIndex        =   2
         ToolTipText     =   "Browse employee data..."
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   8400
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   5250
         Width           =   4695
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1440
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
         TabIndex        =   18
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_employee_code 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DTPicker_date 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   152109059
         CurrentDate     =   39270
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "JSTK LIFE ALLOWANCE (%)"
         Height          =   195
         Left            =   450
         TabIndex        =   54
         Top             =   3360
         Width           =   2010
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "JSTK ACCIDENT ALLOWANCE (%)"
         Height          =   195
         Left            =   450
         TabIndex        =   53
         Top             =   3000
         Width           =   2460
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "SPECIAL ALLOWANCE"
         Height          =   195
         Left            =   450
         TabIndex        =   52
         Top             =   3720
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "STAFF ALLOWANCE"
         Height          =   195
         Left            =   6480
         TabIndex        =   48
         Top             =   1380
         Width           =   1440
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FUNCTIONAL ALLOWANCE"
         Height          =   195
         Left            =   6480
         TabIndex        =   47
         Top             =   1740
         Width           =   1935
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "ADDTIONAL OTHERS"
         Height          =   195
         Left            =   6480
         TabIndex        =   46
         Top             =   2100
         Width           =   1515
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   " PHONE ALLOWANCE"
         Height          =   195
         Left            =   6480
         TabIndex        =   45
         Top             =   2460
         Width           =   1530
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "TRANSPORT ALLOWANCE"
         Height          =   195
         Left            =   6480
         TabIndex        =   44
         Top             =   2880
         Width           =   1875
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "PRESENCE ALLOWANCE"
         Height          =   195
         Left            =   6480
         TabIndex        =   43
         Top             =   3210
         Width           =   1740
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "OVERTIME ALLOWANCE"
         Height          =   195
         Left            =   6480
         TabIndex        =   42
         Top             =   3570
         Width           =   1740
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "MEAL ALLOWANCE"
         Height          =   195
         Left            =   6480
         TabIndex        =   41
         Top             =   4260
         Width           =   1365
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "MAIN SALARY DAILY"
         Height          =   195
         Left            =   6480
         TabIndex        =   40
         Top             =   3930
         Width           =   1485
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "DRIVER ALLOWANCE"
         Height          =   195
         Left            =   6480
         TabIndex        =   39
         Top             =   4620
         Width           =   1530
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "OT (Rp/HOUR)"
         Height          =   195
         Left            =   480
         TabIndex        =   28
         Top             =   4080
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "MAIN SALARY (Rp)"
         Height          =   195
         Left            =   6480
         TabIndex        =   25
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "DATE"
         Height          =   195
         Left            =   720
         TabIndex        =   24
         Top             =   660
         Width           =   390
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   6930
         TabIndex        =   21
         Top             =   5250
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "EMP. CODE"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "EMP. NAME"
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   1440
         Width           =   825
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   16
      Top             =   7560
      Width           =   14175
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin VB.CommandButton cmd_refresh 
         Caption         =   "&Load"
         Height          =   645
         Left            =   8520
         Picture         =   "frm_mst_salary.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   2040
         Picture         =   "frm_mst_salary.frx":0596
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   5280
         Picture         =   "frm_mst_salary.frx":0B20
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   11400
         Picture         =   "frm_mst_salary.frx":10AA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   960
         Picture         =   "frm_mst_salary.frx":1634
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "Re&port"
         Height          =   645
         Left            =   0
         Picture         =   "frm_mst_salary.frx":1BBE
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   4200
         Picture         =   "frm_mst_salary.frx":2148
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   3120
         Picture         =   "frm_mst_salary.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6735
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   11880
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
      Columns(2).Caption=   "DEPT. CODE"
      Columns(2).DataField=   "department_code"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "DEPT. NAME"
      Columns(3).DataField=   "department_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DIV. CODE"
      Columns(4).DataField=   "division_code"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DIV. NAME"
      Columns(5).DataField=   "division_name"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "TITLE CODE"
      Columns(6).DataField=   "title_code"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "TITLE NAME"
      Columns(7).DataField=   "title_name"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "EMP. CODE"
      Columns(8).DataField=   "employee_code"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "EMP. NAME"
      Columns(9).DataField=   "employee_name"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "DATE"
      Columns(10).DataField=   "salary_date"
      Columns(10).NumberFormat=   "FormatText Event"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "SALARY"
      Columns(11).DataField=   "salary"
      Columns(11).NumberFormat=   "Standard"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "DESCRIPTION"
      Columns(12).DataField=   "desc_salary"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   1
      Splits(0).Size  =   4995.213
      Splits(0).Size.vt=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=2223"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2143"
      Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(22)=   "Column(3).Width=5530"
      Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=5450"
      Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(27)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(37)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(41)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(42)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(43)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(45)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(46)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(47)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(48)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(50)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(52)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(53)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(55)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(58)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(59)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(60)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(61)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(62)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(63)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(64)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(66)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(67)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(68)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(69)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(70)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(71)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(72)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(73)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(74)=   "Column(11).AllowSizing=0"
      Splits(0)._ColumnProps(75)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(76)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(77)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(78)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(79)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(80)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(82)=   "Column(12).AllowSizing=0"
      Splits(0)._ColumnProps(83)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(84)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(85)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(86)=   "Column(12).Order=13"
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
      Splits(1)._ColumnProps(0)=   "Columns.Count=13"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(1)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(1)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Width=2725"
      Splits(1)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(1)._WidthInPix=2646"
      Splits(1)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(13)=   "Column(1)._ColStyle=516"
      Splits(1)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(17)=   "Column(2).Width=3942"
      Splits(1)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(19)=   "Column(2)._WidthInPix=3863"
      Splits(1)._ColumnProps(20)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._ColStyle=516"
      Splits(1)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(1)._ColumnProps(23)=   "Column(2).AllowFocus=0"
      Splits(1)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(25)=   "Column(3).Width=7408"
      Splits(1)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(27)=   "Column(3)._WidthInPix=7329"
      Splits(1)._ColumnProps(28)=   "Column(3).AllowSizing=0"
      Splits(1)._ColumnProps(29)=   "Column(3)._ColStyle=516"
      Splits(1)._ColumnProps(30)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(31)=   "Column(3).AllowFocus=0"
      Splits(1)._ColumnProps(32)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(33)=   "Column(4).Width=2725"
      Splits(1)._ColumnProps(34)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(4)._WidthInPix=2646"
      Splits(1)._ColumnProps(36)=   "Column(4).AllowSizing=0"
      Splits(1)._ColumnProps(37)=   "Column(4)._ColStyle=516"
      Splits(1)._ColumnProps(38)=   "Column(4).Visible=0"
      Splits(1)._ColumnProps(39)=   "Column(4).AllowFocus=0"
      Splits(1)._ColumnProps(40)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(41)=   "Column(5).Width=2725"
      Splits(1)._ColumnProps(42)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(43)=   "Column(5)._WidthInPix=2646"
      Splits(1)._ColumnProps(44)=   "Column(5).AllowSizing=0"
      Splits(1)._ColumnProps(45)=   "Column(5)._ColStyle=516"
      Splits(1)._ColumnProps(46)=   "Column(5).Visible=0"
      Splits(1)._ColumnProps(47)=   "Column(5).AllowFocus=0"
      Splits(1)._ColumnProps(48)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(49)=   "Column(6).Width=2725"
      Splits(1)._ColumnProps(50)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(51)=   "Column(6)._WidthInPix=2646"
      Splits(1)._ColumnProps(52)=   "Column(6).AllowSizing=0"
      Splits(1)._ColumnProps(53)=   "Column(6)._ColStyle=516"
      Splits(1)._ColumnProps(54)=   "Column(6).Visible=0"
      Splits(1)._ColumnProps(55)=   "Column(6).AllowFocus=0"
      Splits(1)._ColumnProps(56)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(57)=   "Column(7).Width=2725"
      Splits(1)._ColumnProps(58)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(7)._WidthInPix=2646"
      Splits(1)._ColumnProps(60)=   "Column(7).AllowSizing=0"
      Splits(1)._ColumnProps(61)=   "Column(7)._ColStyle=516"
      Splits(1)._ColumnProps(62)=   "Column(7).Visible=0"
      Splits(1)._ColumnProps(63)=   "Column(7).AllowFocus=0"
      Splits(1)._ColumnProps(64)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(65)=   "Column(8).Width=2090"
      Splits(1)._ColumnProps(66)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(67)=   "Column(8)._WidthInPix=2011"
      Splits(1)._ColumnProps(68)=   "Column(8)._ColStyle=516"
      Splits(1)._ColumnProps(69)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(70)=   "Column(9).Width=5106"
      Splits(1)._ColumnProps(71)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(72)=   "Column(9)._WidthInPix=5027"
      Splits(1)._ColumnProps(73)=   "Column(9)._ColStyle=516"
      Splits(1)._ColumnProps(74)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(75)=   "Column(10).Width=2143"
      Splits(1)._ColumnProps(76)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(77)=   "Column(10)._WidthInPix=2064"
      Splits(1)._ColumnProps(78)=   "Column(10)._ColStyle=513"
      Splits(1)._ColumnProps(79)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(80)=   "Column(11).Width=3175"
      Splits(1)._ColumnProps(81)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(82)=   "Column(11)._WidthInPix=3096"
      Splits(1)._ColumnProps(83)=   "Column(11)._ColStyle=514"
      Splits(1)._ColumnProps(84)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(85)=   "Column(12).Width=3360"
      Splits(1)._ColumnProps(86)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(87)=   "Column(12)._WidthInPix=3281"
      Splits(1)._ColumnProps(88)=   "Column(12)._ColStyle=516"
      Splits(1)._ColumnProps(89)=   "Column(12).Order=13"
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
      Caption         =   "LIST OF SALARY"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=78,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=82,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=86,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=83,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=84,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=85,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=94,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=91,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=92,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=93,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=98,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=95,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=96,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=97,.parent=17"
      _StyleDefs(86)  =   "Splits(1).Style:id=99,.parent=1"
      _StyleDefs(87)  =   "Splits(1).CaptionStyle:id=116,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(88)  =   ":id=116,.fgcolor=&H80000009&"
      _StyleDefs(89)  =   "Splits(1).HeadingStyle:id=100,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(90)  =   ":id=100,.fgcolor=&H80000002&"
      _StyleDefs(91)  =   "Splits(1).FooterStyle:id=101,.parent=3"
      _StyleDefs(92)  =   "Splits(1).InactiveStyle:id=102,.parent=5"
      _StyleDefs(93)  =   "Splits(1).SelectedStyle:id=104,.parent=6"
      _StyleDefs(94)  =   "Splits(1).EditorStyle:id=103,.parent=7"
      _StyleDefs(95)  =   "Splits(1).HighlightRowStyle:id=105,.parent=8"
      _StyleDefs(96)  =   "Splits(1).EvenRowStyle:id=106,.parent=9"
      _StyleDefs(97)  =   "Splits(1).OddRowStyle:id=115,.parent=10"
      _StyleDefs(98)  =   "Splits(1).RecordSelectorStyle:id=117,.parent=11"
      _StyleDefs(99)  =   "Splits(1).FilterBarStyle:id=118,.parent=12"
      _StyleDefs(100) =   "Splits(1).Columns(0).Style:id=122,.parent=99"
      _StyleDefs(101) =   "Splits(1).Columns(0).HeadingStyle:id=119,.parent=100"
      _StyleDefs(102) =   "Splits(1).Columns(0).FooterStyle:id=120,.parent=101"
      _StyleDefs(103) =   "Splits(1).Columns(0).EditorStyle:id=121,.parent=103"
      _StyleDefs(104) =   "Splits(1).Columns(1).Style:id=126,.parent=99"
      _StyleDefs(105) =   "Splits(1).Columns(1).HeadingStyle:id=123,.parent=100"
      _StyleDefs(106) =   "Splits(1).Columns(1).FooterStyle:id=124,.parent=101"
      _StyleDefs(107) =   "Splits(1).Columns(1).EditorStyle:id=125,.parent=103"
      _StyleDefs(108) =   "Splits(1).Columns(2).Style:id=130,.parent=99"
      _StyleDefs(109) =   "Splits(1).Columns(2).HeadingStyle:id=127,.parent=100"
      _StyleDefs(110) =   "Splits(1).Columns(2).FooterStyle:id=128,.parent=101"
      _StyleDefs(111) =   "Splits(1).Columns(2).EditorStyle:id=129,.parent=103"
      _StyleDefs(112) =   "Splits(1).Columns(3).Style:id=134,.parent=99"
      _StyleDefs(113) =   "Splits(1).Columns(3).HeadingStyle:id=131,.parent=100"
      _StyleDefs(114) =   "Splits(1).Columns(3).FooterStyle:id=132,.parent=101"
      _StyleDefs(115) =   "Splits(1).Columns(3).EditorStyle:id=133,.parent=103"
      _StyleDefs(116) =   "Splits(1).Columns(4).Style:id=146,.parent=99"
      _StyleDefs(117) =   "Splits(1).Columns(4).HeadingStyle:id=143,.parent=100"
      _StyleDefs(118) =   "Splits(1).Columns(4).FooterStyle:id=144,.parent=101"
      _StyleDefs(119) =   "Splits(1).Columns(4).EditorStyle:id=145,.parent=103"
      _StyleDefs(120) =   "Splits(1).Columns(5).Style:id=150,.parent=99"
      _StyleDefs(121) =   "Splits(1).Columns(5).HeadingStyle:id=147,.parent=100"
      _StyleDefs(122) =   "Splits(1).Columns(5).FooterStyle:id=148,.parent=101"
      _StyleDefs(123) =   "Splits(1).Columns(5).EditorStyle:id=149,.parent=103"
      _StyleDefs(124) =   "Splits(1).Columns(6).Style:id=154,.parent=99"
      _StyleDefs(125) =   "Splits(1).Columns(6).HeadingStyle:id=151,.parent=100"
      _StyleDefs(126) =   "Splits(1).Columns(6).FooterStyle:id=152,.parent=101"
      _StyleDefs(127) =   "Splits(1).Columns(6).EditorStyle:id=153,.parent=103"
      _StyleDefs(128) =   "Splits(1).Columns(7).Style:id=158,.parent=99"
      _StyleDefs(129) =   "Splits(1).Columns(7).HeadingStyle:id=155,.parent=100"
      _StyleDefs(130) =   "Splits(1).Columns(7).FooterStyle:id=156,.parent=101"
      _StyleDefs(131) =   "Splits(1).Columns(7).EditorStyle:id=157,.parent=103"
      _StyleDefs(132) =   "Splits(1).Columns(8).Style:id=170,.parent=99"
      _StyleDefs(133) =   "Splits(1).Columns(8).HeadingStyle:id=167,.parent=100"
      _StyleDefs(134) =   "Splits(1).Columns(8).FooterStyle:id=168,.parent=101"
      _StyleDefs(135) =   "Splits(1).Columns(8).EditorStyle:id=169,.parent=103"
      _StyleDefs(136) =   "Splits(1).Columns(9).Style:id=174,.parent=99"
      _StyleDefs(137) =   "Splits(1).Columns(9).HeadingStyle:id=171,.parent=100"
      _StyleDefs(138) =   "Splits(1).Columns(9).FooterStyle:id=172,.parent=101"
      _StyleDefs(139) =   "Splits(1).Columns(9).EditorStyle:id=173,.parent=103"
      _StyleDefs(140) =   "Splits(1).Columns(10).Style:id=178,.parent=99,.alignment=2"
      _StyleDefs(141) =   "Splits(1).Columns(10).HeadingStyle:id=175,.parent=100"
      _StyleDefs(142) =   "Splits(1).Columns(10).FooterStyle:id=176,.parent=101"
      _StyleDefs(143) =   "Splits(1).Columns(10).EditorStyle:id=177,.parent=103"
      _StyleDefs(144) =   "Splits(1).Columns(11).Style:id=186,.parent=99,.alignment=1"
      _StyleDefs(145) =   "Splits(1).Columns(11).HeadingStyle:id=183,.parent=100"
      _StyleDefs(146) =   "Splits(1).Columns(11).FooterStyle:id=184,.parent=101"
      _StyleDefs(147) =   "Splits(1).Columns(11).EditorStyle:id=185,.parent=103"
      _StyleDefs(148) =   "Splits(1).Columns(12).Style:id=190,.parent=99"
      _StyleDefs(149) =   "Splits(1).Columns(12).HeadingStyle:id=187,.parent=100"
      _StyleDefs(150) =   "Splits(1).Columns(12).FooterStyle:id=188,.parent=101"
      _StyleDefs(151) =   "Splits(1).Columns(12).EditorStyle:id=189,.parent=103"
      _StyleDefs(152) =   "Named:id=33:Normal"
      _StyleDefs(153) =   ":id=33,.parent=0"
      _StyleDefs(154) =   "Named:id=34:Heading"
      _StyleDefs(155) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(156) =   ":id=34,.wraptext=-1"
      _StyleDefs(157) =   "Named:id=35:Footing"
      _StyleDefs(158) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(159) =   "Named:id=36:Selected"
      _StyleDefs(160) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(161) =   "Named:id=37:Caption"
      _StyleDefs(162) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(163) =   "Named:id=38:HighlightRow"
      _StyleDefs(164) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(165) =   "Named:id=39:EvenRow"
      _StyleDefs(166) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(167) =   "Named:id=40:OddRow"
      _StyleDefs(168) =   ":id=40,.parent=33"
      _StyleDefs(169) =   "Named:id=41:RecordSelector"
      _StyleDefs(170) =   ":id=41,.parent=34"
      _StyleDefs(171) =   "Named:id=42:FilterBar"
      _StyleDefs(172) =   ":id=42,.parent=33"
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
      OleObjectBlob   =   "frm_mst_salary.frx":2C5C
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   1320
      Top             =   360
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   23
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frm_mst_salary"
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

'validasi salary
If Trim(txt_salary) = "" Then
    MsgBox "Salary Masih Kosong...", vbOKOnly + vbInformation, headerMSG
    txt_salary.SetFocus
    check_validate_new = False
    Exit Function
End If

''validasi description
'If Trim(txt_description) = "" Then
'    MsgBox "Description Masih Kosong...", vbOKOnly + vbInformation, headerMSG
'    txt_description.SetFocus
'    check_validate_new = False
'    Exit Function
'End If
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
frm_lookup_mst_employee.public_int_mode = 77
frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_refresh_Click()
Call generate_data_salary
Call load_data_salary
End Sub


Private Sub generate_data_salary()
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim str1 As String
Dim i As Integer



i = MsgBox("Are you sure want to generate data main salary?", vbOKCancel, headerMSG)
If Not i = vbOK Then Exit Sub


str1 = "select * from m_employee where company_code='" & TDBCombo_company.Columns("company_code").Value _
& "' and employee_code not in (select distinct employee_code from m_salary)"
rs1.Open str1, CnG, adOpenStatic, adLockReadOnly

rs2.Open "select * from m_salary where employee_code='uOu'", CnG, adOpenKeyset, adLockOptimistic

While Not rs1.EOF
    rs2.AddNew
    
    rs2.Fields("employee_code").Value = rs1.Fields("employee_code").Value
    rs2.Fields("salary_date").Value = Now 'rs1.Fields("salary_date").Value
    rs2.Fields("salary").Value = 0 'rs1.Fields("salary").Value
    rs2.Fields("over_time").Value = 0 'rs1.Fields("over_time").Value
    rs2.Fields("description").Value = "" 'rs1.Fields("description").Value
    
    rs2.Update
    
    rs1.MoveNext
Wend
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
CnG.Execute "delete from m_salary where employee_code = '" _
    & TDBGrid1.Columns("employee_code").Value & "' and salary_date = '" _
    & Format(TDBGrid1.Columns("salary_date").Value, "yyyy-mm-dd hh:nn:ss") & "'"
CnG.CommitTrans

Call load_data_salary
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    txt_employee_code = .Fields("employee_code").Value
    txt_employee_name = .Fields("employee_name").Value
    DTPicker_date.Value = .Fields("salary_date").Value
    txt_salary = FormatNumber(.Fields("salary").Value)
    txt_over_time = FormatNumber(.Fields("over_time").Value)
    
    txt_description = .Fields("desc_salary").Value
End With
End Sub

Private Sub cmdEdit_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_salary where employee_code = '" _
& Adodc1.Recordset.Fields("employee_code").Value & "' and salary_date = '" _
& Format(Adodc1.Recordset.Fields("salary_date").Value, "yyyy-mm-dd hh:nn:ss") & "'", CnG, adOpenKeyset, adLockOptimistic

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdNew_Click()
If rsBound.State = 1 Then rsBound.Close
rsBound.Open "select * from m_salary where employee_code = '-77'", CnG, adOpenKeyset, adLockOptimistic

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
    .Fields("salary_date").Value = Format(DTPicker_date.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("salary").Value = Val(DropAllComma(txt_salary))
    .Fields("over_time").Value = Val(DropAllComma(txt_over_time))
    
    .Fields("description").Value = Trim(txt_description)
    '.Fields("entry_date").Value = Format(Now, "yyyy-MM-dd HH:mm:ss")
    
    .Update
End With
CnG.CommitTrans
End Sub

Private Sub edit_old_data()
'On Error GoTo err_capture

CnG.BeginTrans
With rsBound

    '.Fields("employee_code").Value = TDBGrid1.Columns("employee_code").Value
    '-----------------------------------------------------------------------------
    '.Fields("salary_date").Value = Format(DTPicker_date.Value, "yyyy-MM-dd HH:mm:ss")
    .Fields("salary").Value = Val(DropAllComma(txt_salary))
    .Fields("over_time").Value = Val(DropAllComma(txt_over_time))
    
    .Fields("description").Value = Trim(txt_description)
    '.Fields("entry_date").Value = Format(Now, "yyyy-MM-dd HH:mm:ss")

    .Update
End With
CnG.CommitTrans

Exit Sub
err_capture:
MsgBox Err.Description
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

Call load_data_salary
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
    
    Call load_data_salary
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

Private Sub load_data_salary()
Adodc1.RecordSource = "select * from v_m_salary where company_code = '" _
& TDBCombo_company.Columns("company_code").Value & "' order by salary_date desc"
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
    Value = Format(Value, "yyyy-mm-dd")
End If
End Sub

Private Sub Timer1_Timer()
timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub txt_department_code_Change()

End Sub

Private Sub txt_salary_Validate(Cancel As Boolean)
If Not Trim(txt_salary) = "" Then
    txt_salary = FormatNumber(DropAllComma(txt_salary))
End If
End Sub
