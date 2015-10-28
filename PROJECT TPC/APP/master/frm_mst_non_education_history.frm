VERSION 5.00
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_education_non_history 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MASTER NON FORMAL EDUCATION HISTORY"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   14715
   Icon            =   "frm_mst_non_education_history.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   14715
   ShowInTaskbar   =   0   'False
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   345
      Left            =   7710
      OleObjectBlob   =   "frm_mst_non_education_history.frx":000C
      TabIndex        =   28
      Top             =   210
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame fra_entry 
      Height          =   3615
      Left            =   240
      TabIndex        =   21
      Top             =   3840
      Width           =   14175
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1560
         Width           =   3495
      End
      Begin VB.ComboBox cbo_status 
         Height          =   315
         ItemData        =   "frm_mst_non_education_history.frx":0030
         Left            =   6960
         List            =   "frm_mst_non_education_history.frx":003A
         TabIndex        =   8
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txt_sequence_number 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt_education_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1200
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
         TabIndex        =   22
         Top             =   120
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.TextBox txt_address 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6960
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1920
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker DTPicker_start 
         Height          =   315
         Left            =   6960
         TabIndex        =   6
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   232652803
         CurrentDate     =   39270
      End
      Begin MSComCtl2.DTPicker DTPicker_finish 
         Height          =   315
         Left            =   8880
         TabIndex        =   7
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   232652803
         CurrentDate     =   39270
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "STATUS"
         Height          =   195
         Left            =   5640
         TabIndex        =   16
         Top             =   2760
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION"
         Height          =   195
         Left            =   5640
         TabIndex        =   27
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "NO"
         Height          =   195
         Left            =   5640
         TabIndex        =   26
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "EDUCATION"
         Height          =   195
         Left            =   5640
         TabIndex        =   25
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ADDRESS"
         Height          =   195
         Left            =   5640
         TabIndex        =   24
         Top             =   1920
         Width           =   780
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "START / FINISH"
         Height          =   195
         Left            =   5640
         TabIndex        =   23
         Top             =   2280
         Width           =   1230
      End
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   19
      Top             =   240
      Width           =   3855
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Height          =   6735
      Left            =   240
      TabIndex        =   18
      Top             =   720
      Width           =   6975
      _ExtentX        =   12303
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
      Columns(6).Caption=   "EMP. CODE"
      Columns(6).DataField=   "employee_code"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "EMP. NAME"
      Columns(7).DataField=   "employee_name"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "NICK NAME"
      Columns(8).DataField=   "employee_nick_name"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "BIRTH DATE"
      Columns(9).DataField=   "date_of_birth"
      Columns(9).NumberFormat=   "FormatText Event"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "PLACE OF BIRTH"
      Columns(10).DataField=   "place_of_birth"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   16
      Columns(11)._MaxComboItems=   5
      Columns(11).ValueItems(0)._DefaultItem=   0
      Columns(11).ValueItems(0).Value=   "0"
      Columns(11).ValueItems(0).Value.vt=   8
      Columns(11).ValueItems(0).DisplayValue=   "Female"
      Columns(11).ValueItems(0).DisplayValue.vt=   8
      Columns(11).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(11).ValueItems(1)._DefaultItem=   0
      Columns(11).ValueItems(1).Value=   "1"
      Columns(11).ValueItems(1).Value.vt=   8
      Columns(11).ValueItems(1).DisplayValue=   "Male"
      Columns(11).ValueItems(1).DisplayValue.vt=   8
      Columns(11).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(11).ValueItems.Count=   2
      Columns(11).Caption=   "SEX"
      Columns(11).DataField=   "sex"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   16
      Columns(12)._MaxComboItems=   5
      Columns(12).ValueItems(0)._DefaultItem=   0
      Columns(12).ValueItems(0).Value=   "0"
      Columns(12).ValueItems(0).Value.vt=   8
      Columns(12).ValueItems(0).DisplayValue=   "Single"
      Columns(12).ValueItems(0).DisplayValue.vt=   8
      Columns(12).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(12).ValueItems(1)._DefaultItem=   0
      Columns(12).ValueItems(1).Value=   "1"
      Columns(12).ValueItems(1).Value.vt=   8
      Columns(12).ValueItems(1).DisplayValue=   "Married"
      Columns(12).ValueItems(1).DisplayValue.vt=   8
      Columns(12).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(12).ValueItems(2)._DefaultItem=   0
      Columns(12).ValueItems(2).Value=   "2"
      Columns(12).ValueItems(2).Value.vt=   8
      Columns(12).ValueItems(2).DisplayValue=   "Widow"
      Columns(12).ValueItems(2).DisplayValue.vt=   8
      Columns(12).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
      Columns(12).ValueItems(3)._DefaultItem=   0
      Columns(12).ValueItems(3).Value=   "3"
      Columns(12).ValueItems(3).Value.vt=   8
      Columns(12).ValueItems(3).DisplayValue=   "Widower"
      Columns(12).ValueItems(3).DisplayValue.vt=   8
      Columns(12).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
      Columns(12).ValueItems.Count=   4
      Columns(12).Caption=   "STATUS"
      Columns(12).DataField=   "marital_status"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "ADDRESS"
      Columns(13).DataField=   "address"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "PHONE NUMBER"
      Columns(14).DataField=   "phone_number"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "BANK ACCOUNT"
      Columns(15).DataField=   "bank_account"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "START WORKING"
      Columns(16).DataField=   "start_working"
      Columns(16).NumberFormat=   "FormatText Event"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "TITLE CODE"
      Columns(17).DataField=   "title_code"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "TITLE NAME"
      Columns(18).DataField=   "title_name"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   4
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "SHIFTABLE"
      Columns(19).DataField=   "flag_shiftable"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   4
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "ACTIVE"
      Columns(20).DataField=   "flag_active"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "END WORKING"
      Columns(21).DataField=   "end_working"
      Columns(21).NumberFormat=   "FormatText Event"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "REASON"
      Columns(22).DataField=   "reason"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   23
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   1
      Splits(0).Size  =   2505.26
      Splits(0).Size.vt=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=23"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1958"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1879"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=3916"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3836"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(17)=   "Column(2).Width=2064"
      Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=1984"
      Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
      Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
      Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(25)=   "Column(3).Width=3545"
      Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=3466"
      Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(29)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(30)=   "Column(4).Width=1508"
      Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=1429"
      Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(34)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(35)=   "Column(5).Width=1508"
      Splits(0)._ColumnProps(36)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(5)._WidthInPix=1429"
      Splits(0)._ColumnProps(38)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(39)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(40)=   "Column(6).Width=1588"
      Splits(0)._ColumnProps(41)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._WidthInPix=1508"
      Splits(0)._ColumnProps(43)=   "Column(6).AllowSizing=0"
      Splits(0)._ColumnProps(44)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(45)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(46)=   "Column(6).AllowFocus=0"
      Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(48)=   "Column(7).Width=1588"
      Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=1508"
      Splits(0)._ColumnProps(51)=   "Column(7).AllowSizing=0"
      Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(53)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(54)=   "Column(7).AllowFocus=0"
      Splits(0)._ColumnProps(55)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(56)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(57)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(59)=   "Column(8).AllowSizing=0"
      Splits(0)._ColumnProps(60)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(61)=   "Column(8).Visible=0"
      Splits(0)._ColumnProps(62)=   "Column(8).AllowFocus=0"
      Splits(0)._ColumnProps(63)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(64)=   "Column(9).Width=2064"
      Splits(0)._ColumnProps(65)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(9)._WidthInPix=1984"
      Splits(0)._ColumnProps(67)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(68)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(69)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(70)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(71)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(72)=   "Column(10).Width=3016"
      Splits(0)._ColumnProps(73)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(74)=   "Column(10)._WidthInPix=2937"
      Splits(0)._ColumnProps(75)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(76)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(77)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(78)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(79)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(80)=   "Column(11).Width=2037"
      Splits(0)._ColumnProps(81)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(82)=   "Column(11)._WidthInPix=1958"
      Splits(0)._ColumnProps(83)=   "Column(11).AllowSizing=0"
      Splits(0)._ColumnProps(84)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(85)=   "Column(11).Visible=0"
      Splits(0)._ColumnProps(86)=   "Column(11).AllowFocus=0"
      Splits(0)._ColumnProps(87)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(88)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(89)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(90)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(91)=   "Column(12).AllowSizing=0"
      Splits(0)._ColumnProps(92)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(93)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(94)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(95)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(96)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(97)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(98)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(99)=   "Column(13).AllowSizing=0"
      Splits(0)._ColumnProps(100)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(101)=   "Column(13).Visible=0"
      Splits(0)._ColumnProps(102)=   "Column(13).AllowFocus=0"
      Splits(0)._ColumnProps(103)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(104)=   "Column(13)._MinWidth=10"
      Splits(0)._ColumnProps(105)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(106)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(108)=   "Column(14).AllowSizing=0"
      Splits(0)._ColumnProps(109)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(110)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(111)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(112)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(113)=   "Column(14)._MinWidth=54215968"
      Splits(0)._ColumnProps(114)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(115)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(116)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(117)=   "Column(15).AllowSizing=0"
      Splits(0)._ColumnProps(118)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(119)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(120)=   "Column(15).AllowFocus=0"
      Splits(0)._ColumnProps(121)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(122)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(123)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(124)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(125)=   "Column(16).AllowSizing=0"
      Splits(0)._ColumnProps(126)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(127)=   "Column(16).Visible=0"
      Splits(0)._ColumnProps(128)=   "Column(16).AllowFocus=0"
      Splits(0)._ColumnProps(129)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(130)=   "Column(16)._MinWidth=60129312"
      Splits(0)._ColumnProps(131)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(132)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(133)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(134)=   "Column(17).AllowSizing=0"
      Splits(0)._ColumnProps(135)=   "Column(17)._ColStyle=516"
      Splits(0)._ColumnProps(136)=   "Column(17).Visible=0"
      Splits(0)._ColumnProps(137)=   "Column(17).AllowFocus=0"
      Splits(0)._ColumnProps(138)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(139)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(140)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(141)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(142)=   "Column(18).AllowSizing=0"
      Splits(0)._ColumnProps(143)=   "Column(18)._ColStyle=516"
      Splits(0)._ColumnProps(144)=   "Column(18).Visible=0"
      Splits(0)._ColumnProps(145)=   "Column(18).AllowFocus=0"
      Splits(0)._ColumnProps(146)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(147)=   "Column(18)._MinWidth=79702332"
      Splits(0)._ColumnProps(148)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(149)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(150)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(151)=   "Column(19).AllowSizing=0"
      Splits(0)._ColumnProps(152)=   "Column(19)._ColStyle=516"
      Splits(0)._ColumnProps(153)=   "Column(19).Visible=0"
      Splits(0)._ColumnProps(154)=   "Column(19).AllowFocus=0"
      Splits(0)._ColumnProps(155)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(156)=   "Column(19)._MinWidth=79897920"
      Splits(0)._ColumnProps(157)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(158)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(159)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(160)=   "Column(20).AllowSizing=0"
      Splits(0)._ColumnProps(161)=   "Column(20)._ColStyle=516"
      Splits(0)._ColumnProps(162)=   "Column(20).Visible=0"
      Splits(0)._ColumnProps(163)=   "Column(20).AllowFocus=0"
      Splits(0)._ColumnProps(164)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(165)=   "Column(20)._MinWidth=79914544"
      Splits(0)._ColumnProps(166)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(167)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(168)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(169)=   "Column(21).AllowSizing=0"
      Splits(0)._ColumnProps(170)=   "Column(21)._ColStyle=516"
      Splits(0)._ColumnProps(171)=   "Column(21).Visible=0"
      Splits(0)._ColumnProps(172)=   "Column(21).AllowFocus=0"
      Splits(0)._ColumnProps(173)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(174)=   "Column(21)._MinWidth=79914544"
      Splits(0)._ColumnProps(175)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(176)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(177)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(178)=   "Column(22).AllowSizing=0"
      Splits(0)._ColumnProps(179)=   "Column(22)._ColStyle=516"
      Splits(0)._ColumnProps(180)=   "Column(22).Visible=0"
      Splits(0)._ColumnProps(181)=   "Column(22).AllowFocus=0"
      Splits(0)._ColumnProps(182)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(183)=   "Column(22)._MinWidth=79789632"
      Splits(1)._UserFlags=   0
      Splits(1).MarqueeStyle=   4
      Splits(1).Size  =   2
      Splits(1).Size.vt=   2
      Splits(1).RecordSelectors=   0   'False
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).ScrollBars=   3
      Splits(1).DividerColor=   13160660
      Splits(1).FilterBar=   -1  'True
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=23"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=1826"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
      Splits(1)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(1)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(1)._ColumnProps(6)=   "Column(0)._ColStyle=516"
      Splits(1)._ColumnProps(7)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(8)=   "Column(0).AllowFocus=0"
      Splits(1)._ColumnProps(9)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(10)=   "Column(1).Width=1826"
      Splits(1)._ColumnProps(11)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(12)=   "Column(1)._WidthInPix=1746"
      Splits(1)._ColumnProps(13)=   "Column(1)._EditAlways=0"
      Splits(1)._ColumnProps(14)=   "Column(1).AllowSizing=0"
      Splits(1)._ColumnProps(15)=   "Column(1)._ColStyle=516"
      Splits(1)._ColumnProps(16)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(17)=   "Column(1).AllowFocus=0"
      Splits(1)._ColumnProps(18)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(19)=   "Column(2).Width=1720"
      Splits(1)._ColumnProps(20)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(21)=   "Column(2)._WidthInPix=1640"
      Splits(1)._ColumnProps(22)=   "Column(2)._EditAlways=0"
      Splits(1)._ColumnProps(23)=   "Column(2).AllowSizing=0"
      Splits(1)._ColumnProps(24)=   "Column(2)._ColStyle=516"
      Splits(1)._ColumnProps(25)=   "Column(2).Visible=0"
      Splits(1)._ColumnProps(26)=   "Column(2).AllowFocus=0"
      Splits(1)._ColumnProps(27)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(28)=   "Column(3).Width=1720"
      Splits(1)._ColumnProps(29)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(30)=   "Column(3)._WidthInPix=1640"
      Splits(1)._ColumnProps(31)=   "Column(3)._EditAlways=0"
      Splits(1)._ColumnProps(32)=   "Column(3).AllowSizing=0"
      Splits(1)._ColumnProps(33)=   "Column(3)._ColStyle=516"
      Splits(1)._ColumnProps(34)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(35)=   "Column(3).AllowFocus=0"
      Splits(1)._ColumnProps(36)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(37)=   "Column(4).Width=1508"
      Splits(1)._ColumnProps(38)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(39)=   "Column(4)._WidthInPix=1429"
      Splits(1)._ColumnProps(40)=   "Column(4)._EditAlways=0"
      Splits(1)._ColumnProps(41)=   "Column(4).AllowSizing=0"
      Splits(1)._ColumnProps(42)=   "Column(4)._ColStyle=516"
      Splits(1)._ColumnProps(43)=   "Column(4).Visible=0"
      Splits(1)._ColumnProps(44)=   "Column(4).AllowFocus=0"
      Splits(1)._ColumnProps(45)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(46)=   "Column(4)._MinWidth=80002672"
      Splits(1)._ColumnProps(47)=   "Column(5).Width=1508"
      Splits(1)._ColumnProps(48)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(49)=   "Column(5)._WidthInPix=1429"
      Splits(1)._ColumnProps(50)=   "Column(5)._EditAlways=0"
      Splits(1)._ColumnProps(51)=   "Column(5).AllowSizing=0"
      Splits(1)._ColumnProps(52)=   "Column(5)._ColStyle=516"
      Splits(1)._ColumnProps(53)=   "Column(5).Visible=0"
      Splits(1)._ColumnProps(54)=   "Column(5).AllowFocus=0"
      Splits(1)._ColumnProps(55)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(56)=   "Column(5)._MinWidth=80001968"
      Splits(1)._ColumnProps(57)=   "Column(6).Width=2011"
      Splits(1)._ColumnProps(58)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(6)._WidthInPix=1931"
      Splits(1)._ColumnProps(60)=   "Column(6)._EditAlways=0"
      Splits(1)._ColumnProps(61)=   "Column(6)._ColStyle=516"
      Splits(1)._ColumnProps(62)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(63)=   "Column(6)._MinWidth=80000960"
      Splits(1)._ColumnProps(64)=   "Column(7).Width=4551"
      Splits(1)._ColumnProps(65)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(66)=   "Column(7)._WidthInPix=4471"
      Splits(1)._ColumnProps(67)=   "Column(7)._EditAlways=0"
      Splits(1)._ColumnProps(68)=   "Column(7)._ColStyle=516"
      Splits(1)._ColumnProps(69)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(70)=   "Column(7)._MinWidth=79999936"
      Splits(1)._ColumnProps(71)=   "Column(8).Width=2725"
      Splits(1)._ColumnProps(72)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(73)=   "Column(8)._WidthInPix=2646"
      Splits(1)._ColumnProps(74)=   "Column(8)._EditAlways=0"
      Splits(1)._ColumnProps(75)=   "Column(8)._ColStyle=516"
      Splits(1)._ColumnProps(76)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(77)=   "Column(8)._MinWidth=80007280"
      Splits(1)._ColumnProps(78)=   "Column(9).Width=2064"
      Splits(1)._ColumnProps(79)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(80)=   "Column(9)._WidthInPix=1984"
      Splits(1)._ColumnProps(81)=   "Column(9)._EditAlways=0"
      Splits(1)._ColumnProps(82)=   "Column(9)._ColStyle=513"
      Splits(1)._ColumnProps(83)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(84)=   "Column(9)._MinWidth=80007280"
      Splits(1)._ColumnProps(85)=   "Column(10).Width=3016"
      Splits(1)._ColumnProps(86)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(87)=   "Column(10)._WidthInPix=2937"
      Splits(1)._ColumnProps(88)=   "Column(10)._EditAlways=0"
      Splits(1)._ColumnProps(89)=   "Column(10)._ColStyle=516"
      Splits(1)._ColumnProps(90)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(91)=   "Column(10)._MinWidth=80010048"
      Splits(1)._ColumnProps(92)=   "Column(11).Width=2037"
      Splits(1)._ColumnProps(93)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(94)=   "Column(11)._WidthInPix=1958"
      Splits(1)._ColumnProps(95)=   "Column(11)._EditAlways=0"
      Splits(1)._ColumnProps(96)=   "Column(11)._ColStyle=513"
      Splits(1)._ColumnProps(97)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(98)=   "Column(12).Width=2725"
      Splits(1)._ColumnProps(99)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(100)=   "Column(12)._WidthInPix=2646"
      Splits(1)._ColumnProps(101)=   "Column(12)._EditAlways=0"
      Splits(1)._ColumnProps(102)=   "Column(12)._ColStyle=513"
      Splits(1)._ColumnProps(103)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(104)=   "Column(13).Width=2725"
      Splits(1)._ColumnProps(105)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(106)=   "Column(13)._WidthInPix=2646"
      Splits(1)._ColumnProps(107)=   "Column(13)._EditAlways=0"
      Splits(1)._ColumnProps(108)=   "Column(13)._ColStyle=516"
      Splits(1)._ColumnProps(109)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(110)=   "Column(14).Width=2725"
      Splits(1)._ColumnProps(111)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(112)=   "Column(14)._WidthInPix=2646"
      Splits(1)._ColumnProps(113)=   "Column(14)._EditAlways=0"
      Splits(1)._ColumnProps(114)=   "Column(14)._ColStyle=516"
      Splits(1)._ColumnProps(115)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(116)=   "Column(15).Width=2725"
      Splits(1)._ColumnProps(117)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(118)=   "Column(15)._WidthInPix=2646"
      Splits(1)._ColumnProps(119)=   "Column(15)._EditAlways=0"
      Splits(1)._ColumnProps(120)=   "Column(15)._ColStyle=516"
      Splits(1)._ColumnProps(121)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(122)=   "Column(16).Width=2725"
      Splits(1)._ColumnProps(123)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(124)=   "Column(16)._WidthInPix=2646"
      Splits(1)._ColumnProps(125)=   "Column(16)._EditAlways=0"
      Splits(1)._ColumnProps(126)=   "Column(16)._ColStyle=513"
      Splits(1)._ColumnProps(127)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(128)=   "Column(17).Width=2725"
      Splits(1)._ColumnProps(129)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(130)=   "Column(17)._WidthInPix=2646"
      Splits(1)._ColumnProps(131)=   "Column(17)._EditAlways=0"
      Splits(1)._ColumnProps(132)=   "Column(17)._ColStyle=516"
      Splits(1)._ColumnProps(133)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(134)=   "Column(18).Width=2725"
      Splits(1)._ColumnProps(135)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(136)=   "Column(18)._WidthInPix=2646"
      Splits(1)._ColumnProps(137)=   "Column(18)._EditAlways=0"
      Splits(1)._ColumnProps(138)=   "Column(18)._ColStyle=516"
      Splits(1)._ColumnProps(139)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(140)=   "Column(19).Width=1720"
      Splits(1)._ColumnProps(141)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(142)=   "Column(19)._WidthInPix=1640"
      Splits(1)._ColumnProps(143)=   "Column(19)._EditAlways=0"
      Splits(1)._ColumnProps(144)=   "Column(19)._ColStyle=513"
      Splits(1)._ColumnProps(145)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(146)=   "Column(20).Width=1588"
      Splits(1)._ColumnProps(147)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(148)=   "Column(20)._WidthInPix=1508"
      Splits(1)._ColumnProps(149)=   "Column(20)._EditAlways=0"
      Splits(1)._ColumnProps(150)=   "Column(20)._ColStyle=513"
      Splits(1)._ColumnProps(151)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(152)=   "Column(21).Width=2725"
      Splits(1)._ColumnProps(153)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(154)=   "Column(21)._WidthInPix=2646"
      Splits(1)._ColumnProps(155)=   "Column(21)._EditAlways=0"
      Splits(1)._ColumnProps(156)=   "Column(21)._ColStyle=513"
      Splits(1)._ColumnProps(157)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(158)=   "Column(22).Width=2725"
      Splits(1)._ColumnProps(159)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(160)=   "Column(22)._WidthInPix=2646"
      Splits(1)._ColumnProps(161)=   "Column(22)._EditAlways=0"
      Splits(1)._ColumnProps(162)=   "Column(22)._ColStyle=516"
      Splits(1)._ColumnProps(163)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(164)=   "Column(22)._MinWidth=80015760"
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
      Caption         =   "LIST OF EMPLOYEE"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=86,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=82,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=78,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=74,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=70,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=98,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=50,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=47,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=48,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=49,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=54,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=62,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=66,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=63,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=64,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=65,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=102,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=99,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=100,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=101,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=110,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=107,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=108,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=109,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=46,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=43,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=44,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=45,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=58,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=55,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=56,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=57,.parent=17"
      _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=94,.parent=13"
      _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=91,.parent=14"
      _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=92,.parent=15"
      _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=93,.parent=17"
      _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=106,.parent=13"
      _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=103,.parent=14"
      _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=104,.parent=15"
      _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=105,.parent=17"
      _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=114,.parent=13"
      _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=14"
      _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=15"
      _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=17"
      _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=222,.parent=13"
      _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=219,.parent=14"
      _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=220,.parent=15"
      _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=221,.parent=17"
      _StyleDefs(118) =   "Splits(0).Columns(21).Style:id=118,.parent=13"
      _StyleDefs(119) =   "Splits(0).Columns(21).HeadingStyle:id=115,.parent=14"
      _StyleDefs(120) =   "Splits(0).Columns(21).FooterStyle:id=116,.parent=15"
      _StyleDefs(121) =   "Splits(0).Columns(21).EditorStyle:id=117,.parent=17"
      _StyleDefs(122) =   "Splits(0).Columns(22).Style:id=122,.parent=13"
      _StyleDefs(123) =   "Splits(0).Columns(22).HeadingStyle:id=119,.parent=14"
      _StyleDefs(124) =   "Splits(0).Columns(22).FooterStyle:id=120,.parent=15"
      _StyleDefs(125) =   "Splits(0).Columns(22).EditorStyle:id=121,.parent=17"
      _StyleDefs(126) =   "Splits(1).Style:id=123,.parent=1"
      _StyleDefs(127) =   "Splits(1).CaptionStyle:id=132,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(128) =   ":id=132,.fgcolor=&H80000009&"
      _StyleDefs(129) =   "Splits(1).HeadingStyle:id=124,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(130) =   ":id=124,.fgcolor=&H80000002&"
      _StyleDefs(131) =   "Splits(1).FooterStyle:id=125,.parent=3"
      _StyleDefs(132) =   "Splits(1).InactiveStyle:id=126,.parent=5"
      _StyleDefs(133) =   "Splits(1).SelectedStyle:id=128,.parent=6"
      _StyleDefs(134) =   "Splits(1).EditorStyle:id=127,.parent=7"
      _StyleDefs(135) =   "Splits(1).HighlightRowStyle:id=129,.parent=8"
      _StyleDefs(136) =   "Splits(1).EvenRowStyle:id=130,.parent=9"
      _StyleDefs(137) =   "Splits(1).OddRowStyle:id=131,.parent=10"
      _StyleDefs(138) =   "Splits(1).RecordSelectorStyle:id=133,.parent=11"
      _StyleDefs(139) =   "Splits(1).FilterBarStyle:id=134,.parent=12"
      _StyleDefs(140) =   "Splits(1).Columns(0).Style:id=138,.parent=123"
      _StyleDefs(141) =   "Splits(1).Columns(0).HeadingStyle:id=135,.parent=124"
      _StyleDefs(142) =   "Splits(1).Columns(0).FooterStyle:id=136,.parent=125"
      _StyleDefs(143) =   "Splits(1).Columns(0).EditorStyle:id=137,.parent=127"
      _StyleDefs(144) =   "Splits(1).Columns(1).Style:id=142,.parent=123"
      _StyleDefs(145) =   "Splits(1).Columns(1).HeadingStyle:id=139,.parent=124"
      _StyleDefs(146) =   "Splits(1).Columns(1).FooterStyle:id=140,.parent=125"
      _StyleDefs(147) =   "Splits(1).Columns(1).EditorStyle:id=141,.parent=127"
      _StyleDefs(148) =   "Splits(1).Columns(2).Style:id=146,.parent=123"
      _StyleDefs(149) =   "Splits(1).Columns(2).HeadingStyle:id=143,.parent=124"
      _StyleDefs(150) =   "Splits(1).Columns(2).FooterStyle:id=144,.parent=125"
      _StyleDefs(151) =   "Splits(1).Columns(2).EditorStyle:id=145,.parent=127"
      _StyleDefs(152) =   "Splits(1).Columns(3).Style:id=150,.parent=123"
      _StyleDefs(153) =   "Splits(1).Columns(3).HeadingStyle:id=147,.parent=124"
      _StyleDefs(154) =   "Splits(1).Columns(3).FooterStyle:id=148,.parent=125"
      _StyleDefs(155) =   "Splits(1).Columns(3).EditorStyle:id=149,.parent=127"
      _StyleDefs(156) =   "Splits(1).Columns(4).Style:id=154,.parent=123"
      _StyleDefs(157) =   "Splits(1).Columns(4).HeadingStyle:id=151,.parent=124"
      _StyleDefs(158) =   "Splits(1).Columns(4).FooterStyle:id=152,.parent=125"
      _StyleDefs(159) =   "Splits(1).Columns(4).EditorStyle:id=153,.parent=127"
      _StyleDefs(160) =   "Splits(1).Columns(5).Style:id=158,.parent=123"
      _StyleDefs(161) =   "Splits(1).Columns(5).HeadingStyle:id=155,.parent=124"
      _StyleDefs(162) =   "Splits(1).Columns(5).FooterStyle:id=156,.parent=125"
      _StyleDefs(163) =   "Splits(1).Columns(5).EditorStyle:id=157,.parent=127"
      _StyleDefs(164) =   "Splits(1).Columns(6).Style:id=162,.parent=123"
      _StyleDefs(165) =   "Splits(1).Columns(6).HeadingStyle:id=159,.parent=124"
      _StyleDefs(166) =   "Splits(1).Columns(6).FooterStyle:id=160,.parent=125"
      _StyleDefs(167) =   "Splits(1).Columns(6).EditorStyle:id=161,.parent=127"
      _StyleDefs(168) =   "Splits(1).Columns(7).Style:id=166,.parent=123"
      _StyleDefs(169) =   "Splits(1).Columns(7).HeadingStyle:id=163,.parent=124"
      _StyleDefs(170) =   "Splits(1).Columns(7).FooterStyle:id=164,.parent=125"
      _StyleDefs(171) =   "Splits(1).Columns(7).EditorStyle:id=165,.parent=127"
      _StyleDefs(172) =   "Splits(1).Columns(8).Style:id=230,.parent=123"
      _StyleDefs(173) =   "Splits(1).Columns(8).HeadingStyle:id=227,.parent=124"
      _StyleDefs(174) =   "Splits(1).Columns(8).FooterStyle:id=228,.parent=125"
      _StyleDefs(175) =   "Splits(1).Columns(8).EditorStyle:id=229,.parent=127"
      _StyleDefs(176) =   "Splits(1).Columns(9).Style:id=170,.parent=123,.alignment=2"
      _StyleDefs(177) =   "Splits(1).Columns(9).HeadingStyle:id=167,.parent=124"
      _StyleDefs(178) =   "Splits(1).Columns(9).FooterStyle:id=168,.parent=125"
      _StyleDefs(179) =   "Splits(1).Columns(9).EditorStyle:id=169,.parent=127"
      _StyleDefs(180) =   "Splits(1).Columns(10).Style:id=174,.parent=123"
      _StyleDefs(181) =   "Splits(1).Columns(10).HeadingStyle:id=171,.parent=124"
      _StyleDefs(182) =   "Splits(1).Columns(10).FooterStyle:id=172,.parent=125"
      _StyleDefs(183) =   "Splits(1).Columns(10).EditorStyle:id=173,.parent=127"
      _StyleDefs(184) =   "Splits(1).Columns(11).Style:id=178,.parent=123,.alignment=2"
      _StyleDefs(185) =   "Splits(1).Columns(11).HeadingStyle:id=175,.parent=124"
      _StyleDefs(186) =   "Splits(1).Columns(11).FooterStyle:id=176,.parent=125"
      _StyleDefs(187) =   "Splits(1).Columns(11).EditorStyle:id=177,.parent=127"
      _StyleDefs(188) =   "Splits(1).Columns(12).Style:id=182,.parent=123,.alignment=2"
      _StyleDefs(189) =   "Splits(1).Columns(12).HeadingStyle:id=179,.parent=124"
      _StyleDefs(190) =   "Splits(1).Columns(12).FooterStyle:id=180,.parent=125"
      _StyleDefs(191) =   "Splits(1).Columns(12).EditorStyle:id=181,.parent=127"
      _StyleDefs(192) =   "Splits(1).Columns(13).Style:id=186,.parent=123"
      _StyleDefs(193) =   "Splits(1).Columns(13).HeadingStyle:id=183,.parent=124"
      _StyleDefs(194) =   "Splits(1).Columns(13).FooterStyle:id=184,.parent=125"
      _StyleDefs(195) =   "Splits(1).Columns(13).EditorStyle:id=185,.parent=127"
      _StyleDefs(196) =   "Splits(1).Columns(14).Style:id=190,.parent=123"
      _StyleDefs(197) =   "Splits(1).Columns(14).HeadingStyle:id=187,.parent=124"
      _StyleDefs(198) =   "Splits(1).Columns(14).FooterStyle:id=188,.parent=125"
      _StyleDefs(199) =   "Splits(1).Columns(14).EditorStyle:id=189,.parent=127"
      _StyleDefs(200) =   "Splits(1).Columns(15).Style:id=194,.parent=123"
      _StyleDefs(201) =   "Splits(1).Columns(15).HeadingStyle:id=191,.parent=124"
      _StyleDefs(202) =   "Splits(1).Columns(15).FooterStyle:id=192,.parent=125"
      _StyleDefs(203) =   "Splits(1).Columns(15).EditorStyle:id=193,.parent=127"
      _StyleDefs(204) =   "Splits(1).Columns(16).Style:id=198,.parent=123,.alignment=2"
      _StyleDefs(205) =   "Splits(1).Columns(16).HeadingStyle:id=195,.parent=124"
      _StyleDefs(206) =   "Splits(1).Columns(16).FooterStyle:id=196,.parent=125"
      _StyleDefs(207) =   "Splits(1).Columns(16).EditorStyle:id=197,.parent=127"
      _StyleDefs(208) =   "Splits(1).Columns(17).Style:id=202,.parent=123"
      _StyleDefs(209) =   "Splits(1).Columns(17).HeadingStyle:id=199,.parent=124"
      _StyleDefs(210) =   "Splits(1).Columns(17).FooterStyle:id=200,.parent=125"
      _StyleDefs(211) =   "Splits(1).Columns(17).EditorStyle:id=201,.parent=127"
      _StyleDefs(212) =   "Splits(1).Columns(18).Style:id=206,.parent=123"
      _StyleDefs(213) =   "Splits(1).Columns(18).HeadingStyle:id=203,.parent=124"
      _StyleDefs(214) =   "Splits(1).Columns(18).FooterStyle:id=204,.parent=125"
      _StyleDefs(215) =   "Splits(1).Columns(18).EditorStyle:id=205,.parent=127"
      _StyleDefs(216) =   "Splits(1).Columns(19).Style:id=210,.parent=123,.alignment=2"
      _StyleDefs(217) =   "Splits(1).Columns(19).HeadingStyle:id=207,.parent=124"
      _StyleDefs(218) =   "Splits(1).Columns(19).FooterStyle:id=208,.parent=125"
      _StyleDefs(219) =   "Splits(1).Columns(19).EditorStyle:id=209,.parent=127"
      _StyleDefs(220) =   "Splits(1).Columns(20).Style:id=226,.parent=123,.alignment=2"
      _StyleDefs(221) =   "Splits(1).Columns(20).HeadingStyle:id=223,.parent=124"
      _StyleDefs(222) =   "Splits(1).Columns(20).FooterStyle:id=224,.parent=125"
      _StyleDefs(223) =   "Splits(1).Columns(20).EditorStyle:id=225,.parent=127"
      _StyleDefs(224) =   "Splits(1).Columns(21).Style:id=214,.parent=123,.alignment=2"
      _StyleDefs(225) =   "Splits(1).Columns(21).HeadingStyle:id=211,.parent=124"
      _StyleDefs(226) =   "Splits(1).Columns(21).FooterStyle:id=212,.parent=125"
      _StyleDefs(227) =   "Splits(1).Columns(21).EditorStyle:id=213,.parent=127"
      _StyleDefs(228) =   "Splits(1).Columns(22).Style:id=218,.parent=123"
      _StyleDefs(229) =   "Splits(1).Columns(22).HeadingStyle:id=215,.parent=124"
      _StyleDefs(230) =   "Splits(1).Columns(22).FooterStyle:id=216,.parent=125"
      _StyleDefs(231) =   "Splits(1).Columns(22).EditorStyle:id=217,.parent=127"
      _StyleDefs(232) =   "Named:id=33:Normal"
      _StyleDefs(233) =   ":id=33,.parent=0"
      _StyleDefs(234) =   "Named:id=34:Heading"
      _StyleDefs(235) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(236) =   ":id=34,.wraptext=-1"
      _StyleDefs(237) =   "Named:id=35:Footing"
      _StyleDefs(238) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(239) =   "Named:id=36:Selected"
      _StyleDefs(240) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(241) =   "Named:id=37:Caption"
      _StyleDefs(242) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(243) =   "Named:id=38:HighlightRow"
      _StyleDefs(244) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(245) =   "Named:id=39:EvenRow"
      _StyleDefs(246) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(247) =   "Named:id=40:OddRow"
      _StyleDefs(248) =   ":id=40,.parent=33"
      _StyleDefs(249) =   "Named:id=41:RecordSelector"
      _StyleDefs(250) =   ":id=41,.parent=34"
      _StyleDefs(251) =   "Named:id=42:FilterBar"
      _StyleDefs(252) =   ":id=42,.parent=33"
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   15
      Top             =   7560
      Width           =   14175
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   645
         Left            =   4680
         Picture         =   "frm_mst_non_education_history.frx":004E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   645
         Left            =   5760
         Picture         =   "frm_mst_non_education_history.frx":05D8
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "&New"
         Height          =   645
         Left            =   2520
         Picture         =   "frm_mst_non_education_history.frx":0B62
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   645
         Left            =   6840
         Picture         =   "frm_mst_non_education_history.frx":10EC
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         Height          =   645
         Left            =   3600
         Picture         =   "frm_mst_non_education_history.frx":1676
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   10680
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmd_exit 
         Caption         =   "E&xit"
         Height          =   645
         Left            =   11520
         Picture         =   "frm_mst_non_education_history.frx":1C00
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   600
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6735
      Left            =   7320
      TabIndex        =   0
      Top             =   720
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   11880
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
      Columns(1).Caption=   "EDUCATION"
      Columns(1).DataField=   "education_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ADDRESS"
      Columns(2).DataField=   "address"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "START"
      Columns(3).DataField=   "start_"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "FINISH"
      Columns(4).DataField=   "finish_"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "STATUS"
      Columns(5).DataField=   "status_"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=794"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=714"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._MinWidth=64"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=4101"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4022"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=512"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(1)._MinWidth=75697772"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=2884"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=2805"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=1826"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=1746"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=1826"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=1746"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=512"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(33)=   "Column(5)._MinWidth=121772248"
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
      Caption         =   "LIST OF NON FORMAL EDUCATION"
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=46,.parent=13,.alignment=0"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=28,.parent=13,.alignment=0"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
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
      Left            =   12600
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Appearance      =   0
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
         Name            =   "MS Sans Serif"
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
      Left            =   10680
      Top             =   0
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      Appearance      =   0
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc_company 
      Height          =   375
      Left            =   2880
      Top             =   240
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
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "frm_mst_non_education_history.frx":218A
      TabIndex        =   1
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   795
   End
End
Attribute VB_Name = "frm_mst_education_non_history"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn_fp As Boolean
Dim glngEnrollData(DataSize) As Long
Dim glngEnrollPData As Integer
Dim gbytEnrollData(DataSize * 5) As Byte
Dim vEMachineNumber As Integer
Public start_rec, end_rec, running_rec, db_rec As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks
Dim int_mode As Integer




Public Sub download_enroll(ByVal int_number As Integer)
Dim vEnrollNumber 'As Integer
Dim vFingerNumber 'As Integer
Dim vPrivilege As Integer
Dim vEnable As Integer
Dim vFlag As Boolean
Dim vRet As Boolean
Dim vErrorCode As Long
Dim vStr As String
Dim i As Long
Dim j
Dim rs As New ADODB.Recordset

'---- Get Enroll data and save it into database -------------
vEnrollNumber = int_number
vRet = CZKEM1.GetEnrollData(vMachineNumber, vEnrollNumber, vEMachineNumber, _
        vFingerNumber, vPrivilege, glngEnrollData(0), glngEnrollPData)
   
If vRet Then
    rs.Open "select * from m_enroll where enrollnumber = " & vEnrollNumber & " and ip_address = '" _
        & FG_IP_ADDRESS & "'", CnG, adOpenKeyset, adLockOptimistic
    With rs
        '!EMachineNumber = vMachineNumber
        '!EnrollNumber = vEnrollNumber
        !FingerNumber = vFingerNumber
        !Privilige = vPrivilege
        
        'If vFingerNumber = 10 Then
        '    !passWord = glngEnrollPData
        'Else
            For i = 0 To DataSize - 1
                gbytEnrollData(i * 5) = 1
                If glngEnrollData(i) < 0 Then
                    gbytEnrollData(i * 5) = 0
                    glngEnrollData(i) = Abs(glngEnrollData(i))
                End If
                gbytEnrollData(i * 5 + 1) = (glngEnrollData(i) \ 256 \ 256 \ 256)
                gbytEnrollData(i * 5 + 2) = (glngEnrollData(i) \ 256 \ 256) Mod 256
                gbytEnrollData(i * 5 + 3) = (glngEnrollData(i) \ 256) Mod 256
                gbytEnrollData(i * 5 + 4) = glngEnrollData(i) Mod 256
            Next
            !FPdata = gbytEnrollData
        'End If
        .Update
    End With
    rs.Close
End If

End Sub

Private Function set_bookmark_ado(ByVal lng_no As Long) As Boolean
Adodc1.Recordset.MoveFirst

Adodc1.Recordset.Find ("enrollnumber=" & lng_no)   ', 0, adSearchForward, 1)
If Not (Adodc1.Recordset.EOF = True Or Adodc1.Recordset.BOF = True) Then
    set_bookmark_ado = True
Else
    set_bookmark_ado = False
End If
End Function

Private Sub cmd_browse_relation_Click()
frm_mst_family_level.public_int_mode = 1
frm_mst_family_level.cmd_select.Visible = True
frm_mst_family_level.Show 1
End Sub

Private Sub cmd_browse_education_Click()
frm_mst_education.public_int_mode = 2
frm_mst_education.cmd_select.Visible = True
frm_mst_education.Show 1
End Sub

Private Sub cmd_exit_Click()
Unload Me
End Sub


Public Sub cmd_refresh_Click()
Timer1.Enabled = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Adodc1.Caption = "Record " & Adodc1.Recordset.AbsolutePosition & " of " & Adodc1.Recordset.RecordCount
End Sub




Private Sub Form_Load()
Adodc1.ConnectionString = strConn
Adodc2.ConnectionString = strConn
Adodc_company.ConnectionString = strConn

cn_fp = False
start_rec = 0
end_rec = 100
running_rec = 0
db_rec = 0

vMachineNumber = 1
vEMachineNumber = 1

Call load_data_company
Call load_data_user_access(Me)

Timer1.Enabled = True

int_mode = 0
Call load_mode

'cmdEdit.Enabled = False
'cmdDelete.Enabled = False
'CmdNew.Enabled = False

End Sub

Private Sub load_data_employee()
Adodc2.RecordSource = "select * from m_employee where company_code = '" _
& TDBCombo_company.Columns("company_code").Value & "' order by employee_code"
Adodc2.Refresh

TDBGrid2.DataSource = Adodc2
If Adodc2.Recordset.RecordCount > 0 Then Call load_data_education_history
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    Call load_data_employee
End If
End Sub



Private Sub TDBGrid2_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid2, Adodc2)
If TDBGrid2.ApproxCount > 0 Then Call load_data_education_history
End Sub

Private Sub TDBGrid1_FilterChange()
Call tdbgrid_filter(Cols, Col, TDBGrid1, Adodc1)

End Sub


Private Sub TDBGrid2_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
If TDBGrid2.Columns(ColIndex).Caption = "BIRTH DATE" Or _
TDBGrid2.Columns(ColIndex).Caption = "START WORKING" Or _
TDBGrid2.Columns(ColIndex).Caption = "END WORKING" Then
    Value = Format(Value, "dd-mm-yyyy")
End If
End Sub

Private Sub TDBGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If (TDBGrid2.Row + 1) > 0 And (TDBGrid2.Row + 1) <> LastRow Then
    Call load_data_education_history
End If
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Call set_company_mode(Adodc_company, TDBCombo_company, txt_company_name)
End Sub

Private Sub load_data_company()
Adodc_company.RecordSource = "select * from m_company order by company_code"
Adodc_company.Refresh

TDBCombo_company.RowSource = Adodc_company
End Sub

Public Sub load_data_education_history()
Adodc1.RecordSource = "select *, CAST(DATE_FORMAT(start,'%d-%m-%Y') AS CHAR) AS start_, " _
& "CAST(DATE_FORMAT(finish,'%d-%m-%Y') AS CHAR) AS finish_, case when status=1 then 'PASS' " _
& "else 'NOT PASS' end as status_ from t_education_non where employee_code='" _
& Adodc2.Recordset.Fields("employee_code").Value & "' order by sequence_number asc"
Adodc1.Refresh

'cmdEdit.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'cmdDelete.Enabled = IIf(Adodc1.Recordset.RecordCount = 0, False, True)
'CmdNew.Enabled = IIf(TDBCombo_company.Columns("company_code").Text = "", False, True)

TDBGrid1.DataSource = Adodc1
End Sub

Private Function connect() As Boolean
If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.disconnect
End If

cn_fp = CZKEM1.Connect_Net(FG_IP_ADDRESS, CLng(FG_PORT_NUMBER))

If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, False
    connect = True
Else
    connect = False
    Exit Function
End If
End Function

Private Function disconnect() As Boolean
If cn_fp Then
    CZKEM1.EnableDevice vMachineNumber, True
    CZKEM1.disconnect
End If
End Function


Public Function get_all_user_info() As Integer
On Error GoTo err_capture

Dim dwEnrollNmber As Long
Dim dwMachineNumber As Long
Dim name As String
Dim Password As String
Dim privilege As Integer
Dim Enabled As Boolean
Dim blni As Boolean
Dim i As Long

dwMachineNumber = 1
i = 0
blni = CZKEM1.ReadAllUserID(1)
If blni Then
    CnG.BeginTrans
    While CZKEM1.GetAllUserInfo(CLng(dwMachineNumber), dwEnrollNmber, _
            name, Password, privilege, Enabled)
        
        CnG.Execute "call spi_m_enroll ('" & FG_IP_ADDRESS & "'," & dwEnrollNmber & ",'" _
        & Replace(name, "'", "\'") & "'," & dwMachineNumber & ", 0, " & privilege & ",'" & Password & "')"
        i = i + 1
    Wend
    CnG.CommitTrans
End If
get_all_user_info = i


Exit Function
err_capture:
CnG.RollbackTrans
MsgBox "Error reading device...", vbCritical, headerMSG
Call disconnect
get_all_user_info = 0
End Function







'===

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
check_validate_exist_new = False

str_sql = "select count(sequence_number) as rec_count from t_education_non where employee_code = '" _
& Adodc2.Recordset.Fields("employee_code").Value & "' and sequence_number = " & Val(txt_sequence_number)
rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly

If rs.Fields("rec_count").Value > 0 Then
    check_validate_exist_new = True
    Exit Function
End If
End Function

Private Sub check_invalid()
MsgBox "Data found!", vbCritical, headerMSG
txt_sequence_number = ""
If txt_sequence_number.Enabled = True Then txt_sequence_number.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
check_validate_exist_edit = False

If Not Val(txt_sequence_number) = Val(Adodc1.Recordset.Fields("sequence_number").Value) And _
check_validate_exist_new Then
    check_validate_exist_edit = True
    Exit Function
End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

If Trim(txt_sequence_number) = "" Then
    MsgBox "No. is empty!", vbOKOnly + vbInformation, headerMSG
    txt_sequence_number.SetFocus
    check_validate_new = False
    Exit Function
End If

If Trim(txt_education_name) = "" Then
    MsgBox "Education is not selected!", vbOKOnly + vbInformation, headerMSG
    txt_education_name.SetFocus
    check_validate_new = False
    Exit Function
End If

If Trim(txt_address) = "" Then
    MsgBox "Address is empty!", vbOKOnly + vbInformation, headerMSG
    txt_address.SetFocus
    check_validate_new = False
    Exit Function
End If

End Function


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
    & Adodc1.Recordset.Fields("education_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
If Not i = vbYes Then Exit Sub

CnG.BeginTrans
CnG.Execute "delete from t_education_non where employee_code = '" _
    & Adodc1.Recordset.Fields("employee_code").Value & "' and sequence_number = " _
    & Adodc1.Recordset.Fields("sequence_number").Value
CnG.CommitTrans

Call load_data_education_history
int_mode = 0
Call load_mode
End Sub

Public Sub set_edit_data()
With Adodc1.Recordset
    txt_sequence_number = .Fields("sequence_number").Value
'    txt_education_code = .Fields("education_code").Value
    txt_education_name = .Fields("education_name").Value
    txt_description = .Fields("description").Value
    txt_address = .Fields("address").Value
    DTPicker_start = .Fields("start").Value
    DTPicker_finish = .Fields("finish").Value
    cbo_status.ListIndex = Val("" & .Fields("status").Value)
End With
End Sub

Private Sub cmdEdit_Click()
If Not TDBGrid2.ApproxCount > 0 Then
    MsgBox "No data selected!", vbInformation, headerMSG
    Exit Sub
End If

int_mode = 2
Call load_mode
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdNew_Click()
If Not TDBGrid2.ApproxCount > 0 Then
    MsgBox "No data selected!", vbInformation, headerMSG
    Exit Sub
End If

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
Dim rs As New ADODB.Recordset

If rs.State = 1 Then rs.Close
rs.Open "select * from t_education_non where sequence_number = -77", CnG, adOpenKeyset, adLockOptimistic


CnG.BeginTrans

With rs
    .AddNew
    
    .Fields("employee_code").Value = Adodc2.Recordset.Fields("employee_code").Value
    .Fields("sequence_number").Value = Val(txt_sequence_number)
    '---------------------------------
    .Fields("employee_name").Value = Adodc2.Recordset.Fields("employee_name").Value
    '.Fields("education_code").Value = Trim(txt_education_code)
    .Fields("education_name").Value = Trim(txt_education_name)
    .Fields("description").Value = Trim(txt_description)
    .Fields("address").Value = Trim(txt_address)
    .Fields("start").Value = DTPicker_start
    .Fields("finish").Value = DTPicker_finish
    
    .Fields("status").Value = cbo_status.ListIndex
    
    .Update
End With

CnG.CommitTrans
End Sub

Private Sub edit_old_data()
On Error GoTo err_capture

Dim rs As New ADODB.Recordset

If rs.State = 1 Then rs.Close
rs.Open "select * from t_education_non where sequence_number = " _
    & Adodc1.Recordset.Fields("sequence_number").Value & " and employee_code ='" _
    & Adodc1.Recordset.Fields("employee_code").Value & "'", CnG, adOpenKeyset, adLockOptimistic

CnG.BeginTrans
With rs

    .Fields("employee_code").Value = Adodc2.Recordset.Fields("employee_code").Value
    .Fields("sequence_number").Value = Val(txt_sequence_number)
    '---------------------------------
    .Fields("employee_name").Value = Adodc2.Recordset.Fields("employee_name").Value
    '.Fields("education_code").Value = Trim(txt_education_code)
    .Fields("education_name").Value = Trim(txt_education_name)
    .Fields("description").Value = Trim(txt_description)
    .Fields("address").Value = Trim(txt_address)
    .Fields("start").Value = DTPicker_start
    .Fields("finish").Value = DTPicker_finish
    .Fields("status").Value = cbo_status.ListIndex
    
    .Update
End With
CnG.CommitTrans

Exit Sub
err_capture:
rs.CancelBatch adAffectCurrent: rs.Close: CnG.RollbackTrans
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

Call load_data_education_history
int_mode = 0
Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal C As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal F As Boolean, ByVal g As Boolean)
CmdNew.Enabled = a And blnUser_Add
CmdSave.Enabled = b
cmdEdit.Enabled = C And blnUser_Edit
cmdDelete.Enabled = d And blnUser_Delete
CmdCancel.Enabled = e

'CmdPrint.Enabled = F
'cmd_refresh.Enabled = g
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
cbo_status.ListIndex = 1
End Sub

Private Sub set_data_mode()
If int_mode = 1 Then        'NEW
    Call clear_view_data
    fra_entry.Visible = True
'    txt_education_code.Enabled = True
    TDBGrid1.Enabled = False
    Call set_new_data
    
'    If txt_relation_code.Enabled = True Then
'        txt_relation_code.SetFocus
'    End If
    
ElseIf int_mode = 0 Then    'VIEW
    Call clear_view_data
    fra_entry.Visible = False
    TDBGrid1.Enabled = True

ElseIf int_mode = 2 Then    'EDIT
    Call set_edit_data
'    txt_education_code.Enabled = False
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
