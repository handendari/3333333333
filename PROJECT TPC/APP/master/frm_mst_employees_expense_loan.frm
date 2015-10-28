VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_salary_exp_employee_installment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "EMPLOYEE LOAN"
   ClientHeight    =   9630
   ClientLeft      =   -15
   ClientTop       =   300
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
   ScaleHeight     =   9630
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraCloseLoan 
      Caption         =   "Close Loan"
      Height          =   3585
      Left            =   4590
      TabIndex        =   55
      Top             =   2880
      Visible         =   0   'False
      Width           =   5445
      Begin VB.TextBox txtInterest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         TabIndex        =   61
         Top             =   1230
         Width           =   2715
      End
      Begin VB.Frame Frame2 
         Height          =   555
         Left            =   1920
         TabIndex        =   62
         Top             =   1500
         Width           =   2745
         Begin VB.OptionButton optInterestFrom 
            Caption         =   "Loan Remains"
            Height          =   255
            Index           =   1
            Left            =   1260
            TabIndex        =   64
            Top             =   180
            Width           =   1335
         End
         Begin VB.OptionButton optInterestFrom 
            Caption         =   "Loan Basic"
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   63
            Top             =   180
            Value           =   -1  'True
            Width           =   1245
         End
      End
      Begin prj_tpc.vbButton cmdSave_Close 
         Height          =   705
         Left            =   2730
         TabIndex        =   66
         Top             =   2640
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":058A
         PICN            =   "frm_mst_employees_expense_loan.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdCancel_Loan 
         Height          =   705
         Left            =   3720
         TabIndex        =   67
         Top             =   2640
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":1638
         PICN            =   "frm_mst_employees_expense_loan.frx":1654
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker_Month 
         Height          =   315
         Left            =   1920
         TabIndex        =   71
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MM-yyyy"
         Format          =   95289347
         CurrentDate     =   40794
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "MONTH"
         Height          =   285
         Left            =   600
         TabIndex        =   70
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblPaid 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   69
         Top             =   2100
         Width           =   2715
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "TO BE PAID"
         Height          =   285
         Left            =   570
         TabIndex        =   68
         Top             =   2130
         Width           =   1245
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "FROM"
         Height          =   285
         Left            =   570
         TabIndex        =   65
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label lblLoanRemains 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   60
         Top             =   870
         Width           =   2715
      End
      Begin VB.Label lblLoanBasic 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1920
         TabIndex        =   59
         Top             =   540
         Width           =   2715
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "INTEREST (%)"
         Height          =   285
         Left            =   570
         TabIndex        =   58
         Top             =   1260
         Width           =   1245
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "LOAN REMAINS"
         Height          =   285
         Left            =   600
         TabIndex        =   57
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "LOAN BASIC"
         Height          =   285
         Left            =   600
         TabIndex        =   56
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame fra_status_emp 
      Caption         =   "TYPE"
      Height          =   585
      Left            =   11010
      TabIndex        =   44
      Top             =   930
      Width           =   3405
      Begin VB.OptionButton optCompany 
         Caption         =   "COMPANY"
         Height          =   225
         Left            =   240
         TabIndex        =   46
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton optKoperasi 
         Caption         =   "KOPERASI"
         Height          =   225
         Left            =   1890
         TabIndex        =   45
         Top             =   270
         Width           =   1245
      End
   End
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   2925
      Left            =   2280
      TabIndex        =   43
      Top             =   3150
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   5159
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHeader {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColorSel    =   12937777
      ForeColorSel    =   16777215
      CustomColorFrom =   16572875
      CustomColorTo   =   14722429
      GridColor       =   16367254
      FocusRectColor  =   9895934
      GridLines       =   2
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
      ColumnSort      =   -1  'True
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid2 
      Height          =   2175
      Left            =   240
      TabIndex        =   23
      Top             =   5850
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   3836
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
      Columns(1).NumberFormat=   "MM-yyyy"
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
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
      Splits(0)._ColumnProps(26)=   "Column(5).Width=6271"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=6191"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=512"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   14
      Top             =   750
      Width           =   3855
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   8070
      Width           =   14175
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin prj_tpc.vbButton cmdNew 
         Height          =   705
         Left            =   930
         TabIndex        =   36
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&New"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":26E6
         PICN            =   "frm_mst_employees_expense_loan.frx":2702
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdSave 
         Height          =   705
         Left            =   1950
         TabIndex        =   37
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":3794
         PICN            =   "frm_mst_employees_expense_loan.frx":37B0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdEdit 
         Height          =   705
         Left            =   2970
         TabIndex        =   38
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Edit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":4842
         PICN            =   "frm_mst_employees_expense_loan.frx":485E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdDelete 
         Height          =   705
         Left            =   3990
         TabIndex        =   39
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Delete"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":58F0
         PICN            =   "frm_mst_employees_expense_loan.frx":590C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdCancel 
         Height          =   705
         Left            =   5010
         TabIndex        =   40
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Cancel"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":699E
         PICN            =   "frm_mst_employees_expense_loan.frx":69BA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdExit 
         Height          =   705
         Left            =   11550
         TabIndex        =   41
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Exit"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":7A4C
         PICN            =   "frm_mst_employees_expense_loan.frx":7A68
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdCloseLoan 
         Height          =   705
         Left            =   10140
         TabIndex        =   54
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "Close &Loan"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":8AFA
         PICN            =   "frm_mst_employees_expense_loan.frx":8B16
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1260
      OleObjectBlob   =   "frm_mst_employees_expense_loan.frx":9BA8
      TabIndex        =   1
      Top             =   750
      Width           =   1695
   End
   Begin VB.Frame fra_entry 
      Height          =   3825
      Left            =   240
      TabIndex        =   17
      Top             =   1980
      Width           =   14175
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
         TabIndex        =   4
         Top             =   630
         Width           =   3015
      End
      Begin VB.TextBox txt_interest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10560
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txt_instalment_times 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10560
         MaxLength       =   10
         TabIndex        =   6
         Text            =   "1"
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
         TabIndex        =   10
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
         TabIndex        =   7
         Top             =   1800
         Width           =   3015
      End
      Begin VB.CommandButton cmd_browse_item 
         Caption         =   "..."
         Height          =   300
         Left            =   3360
         TabIndex        =   2
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
         TabIndex        =   22
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
         TabIndex        =   13
         Top             =   2760
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1890
         Width           =   5175
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3840
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   18
         Top             =   840
         Width           =   3375
      End
      Begin VB.TextBox txt_nik 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker_date 
         Height          =   300
         Left            =   2040
         TabIndex        =   11
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   95289347
         CurrentDate     =   39278
      End
      Begin MSComCtl2.DTPicker DTPicker_instalment_start 
         Height          =   315
         Left            =   10560
         TabIndex        =   8
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "MM-yyyy"
         Format          =   95289347
         CurrentDate     =   39270
      End
      Begin MSComCtl2.DTPicker DTPicker_instalment_end 
         Height          =   315
         Left            =   10560
         TabIndex        =   9
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "MM-yyyy"
         Format          =   95289347
         CurrentDate     =   39270
      End
      Begin prj_tpc.vbButton cmdBrowse 
         Height          =   315
         Left            =   3420
         TabIndex        =   42
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_mst_employees_expense_loan.frx":BB0E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txt_employee_code 
         Height          =   315
         Left            =   3480
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label5 
         Caption         =   "* yyyy-MM-dd"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   12030
         TabIndex        =   34
         Top             =   2910
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "* yyyy-MM-dd"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   12030
         TabIndex        =   33
         Top             =   2580
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION*"
         Height          =   195
         Left            =   690
         TabIndex        =   31
         Top             =   1890
         Width           =   1110
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "LOAN VALUE* (IDR)"
         Height          =   195
         Left            =   8280
         TabIndex        =   30
         Top             =   600
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "INSTALMENT VALUE (IDR)"
         Height          =   195
         Left            =   8280
         TabIndex        =   29
         Top             =   3240
         Width           =   1875
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "INSTALMENT START"
         Height          =   195
         Left            =   8280
         TabIndex        =   28
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "FLAT INTEREST/MONTH (%)"
         Height          =   195
         Left            =   8280
         TabIndex        =   27
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "INSTALMENT TIMES*"
         Height          =   195
         Left            =   8280
         TabIndex        =   26
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "INSTALMENT END"
         Height          =   195
         Left            =   8280
         TabIndex        =   25
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "LOAN TOTAL (IDR)"
         Height          =   195
         Left            =   8280
         TabIndex        =   24
         Top             =   1800
         Width           =   1365
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
         TabIndex        =   21
         Top             =   2760
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "EMPLOYEE*"
         Height          =   195
         Left            =   720
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DATE*"
         Height          =   195
         Left            =   720
         TabIndex        =   19
         Top             =   480
         Width           =   480
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3765
      Left            =   240
      TabIndex        =   16
      Top             =   2040
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
      Columns(0).NumberFormat=   "yyyy-MM-dd"
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
      Columns(4).NumberFormat=   "Standard"
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
      Columns(6).NumberFormat=   "Standard"
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
   Begin MSComCtl2.DTPicker DTPicker_from 
      Height          =   315
      Left            =   1260
      TabIndex        =   47
      Top             =   1470
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   95289347
      CurrentDate     =   41332
   End
   Begin MSComCtl2.DTPicker DTPicker_to 
      Height          =   315
      Left            =   2970
      TabIndex        =   48
      Top             =   1470
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   95289347
      CurrentDate     =   41332
   End
   Begin prj_tpc.vbButton cmdSearch 
      Height          =   465
      Left            =   4380
      TabIndex        =   49
      Top             =   1410
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&View"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_mst_employees_expense_loan.frx":BB2A
      PICN            =   "frm_mst_employees_expense_loan.frx":BB46
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.DTPicker DTPicker_Periode 
      Height          =   315
      Left            =   1260
      TabIndex        =   52
      Top             =   1110
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MM-yyyy"
      Format          =   95289347
      CurrentDate     =   40794
   End
   Begin VB.Label lbl_remains 
      Alignment       =   1  'Right Justify
      Caption         =   "LOAN REMAINS = 0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   8730
      TabIndex        =   72
      Top             =   1740
      Width           =   5685
   End
   Begin VB.Label Label18 
      Caption         =   "PERIODE"
      Height          =   195
      Left            =   450
      TabIndex        =   53
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "DATE"
      Height          =   195
      Left            =   450
      TabIndex        =   51
      Top             =   1500
      Width           =   390
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   50
      Top             =   1500
      Width           =   285
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE LOAN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   35
      Top             =   150
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY"
      Height          =   195
      Left            =   450
      TabIndex        =   15
      Top             =   810
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_employees_expense_loan.frx":CBD8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14790
   End
End
Attribute VB_Name = "frm_mst_salary_exp_employee_installment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCompany As New ADODB.Recordset
Dim rsLoan As New ADODB.Recordset
Dim rsLoanDetail As New ADODB.Recordset
Dim int_mode As Integer

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Dim v_value As Double
Dim v_installment_time As Double
Dim v_interest As Double
Dim strsql As String

Dim vParam As String

Private Function check_validate_exist_new() As Boolean
Dim str_sql As String
    check_validate_exist_new = False
    
    If rsLoan.State Then rsLoan.Close
    str_sql = "select count(employee_code) as rec_count from tm_loan where date(date) = '" _
                & Format(DTPicker_date.Value, "yyyy-MM-dd") & "' and employee_code = '" & txt_employee_code.Text & "' " & _
                "and flag_loan = '" & IIf(optCompany.Value, 0, 1) & "'"
    rsLoan.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rsLoan.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        rsLoan.Close
        Exit Function
    End If
    
    rsLoan.Close
End Function

Private Sub check_invalid()
    MsgBox "Data found!", vbCritical, headerMSG
    txt_nik = ""
    If txt_nik.Enabled = True Then txt_nik.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If Not txt_employee_code = rsLoan.Fields("employee_code").Value And _
    check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True

    If Trim(txt_nik) = "" Then
        MsgBox "Employee Code is empty!", vbOKOnly + vbInformation, headerMSG
        txt_nik.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    If Trim(txt_description) = "" Then
        MsgBox "Description is empty!", vbOKOnly + vbInformation, headerMSG
        txt_description.SetFocus
        check_validate_new = False
        Exit Function
    End If

    If Trim(txt_loan_value) = "" Then
        MsgBox "Loan Value is empty!", vbOKOnly + vbInformation, headerMSG
        txt_loan_value.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    If Trim(txt_instalment_times) = "" Then
        MsgBox "Instalment Time is empty!", vbOKOnly + vbInformation, headerMSG
        txt_instalment_times.SetFocus
        check_validate_new = False
        Exit Function
    End If

End Function

Private Sub load_data()
    timer1.Enabled = True
End Sub

'Private Sub cmd_browse_Click()
'    frm_lookup_mst_employee.public_int_mode = 77
'    frm_lookup_mst_employee.Show 1
'End Sub

Private Sub cmd_browse_employee_Click()
    frm_lookup_mst_employee.public_int_mode = 162
    frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
    frm_lookup_mst_employee.Show 1
End Sub

'Private Sub cmd_browse_item_Click()
'    frm_mst_salary_item.public_int_mode = 13
'    frm_mst_salary_item.cmd_Select.Visible = True
'    frm_mst_salary_item.Show 1
'End Sub

Private Sub cmdCancel_Click()
    int_mode = 0
    Call load_mode
End Sub

Private Sub cmdCancel_Loan_Click()
    fraCloseLoan.Visible = False
End Sub

Private Sub cmdCloseLoan_Click()
Dim vLoanRemains As Double
    If Not TDBGrid1.ApproxCount > 0 Then
        Exit Sub
    End If
    
    fraCloseLoan.Visible = True
    optInterestFrom(0).Value = True
    Call tutupKoperasi
    
    txtInterest.SetFocus
    
    lblLoanBasic.Caption = FormatNumber(TDBGrid1.Columns("loan_total").Value)
    lblPaid.Caption = FormatNumber(lblLoanBasic.Caption)
    txtInterest.Text = 0
    
    SQL = "SELECT installment_month FROM td_loan " & _
          "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND DATE(date) = '" & Format(TDBGrid1.Columns("date").Value, "yyyy-MM-dd") & "' " & _
            "AND flag_loan = " & IIf(optCompany.Value, 0, 1) & " " & _
            "AND flag_paid <> 1 " & _
          "ORDER BY installment_month ASC LIMIT 1"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        DTPicker_month.Value = rscari!installment_month
    End If
    rscari.Close
    
    SQL = "SELECT SUM(installment_amount) sisa_loan FROM td_loan " & _
          "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND DATE(date) = '" & Format(TDBGrid1.Columns("date").Value, "yyyy-MM-dd") & "' " & _
            "AND flag_loan = " & IIf(optCompany.Value, 0, 1) & " " & _
            "AND flag_paid <> 1"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vLoanRemains = IIf(IsNull(rscari!sisa_loan), 0, rscari!sisa_loan)
    End If
    rscari.Close
    
    lblLoanRemains = FormatNumber(vLoanRemains)
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer
Dim item

On Error GoTo Err
    If Not TDBGrid1.ApproxCount > 0 Then
        Exit Sub
    End If
        
    Set SelBks = TDBGrid1.SelBookmarks
    i = MsgBox("Are you sure want to delete " _
        & SelBks.Count & " loan's data ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
                
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        SQL = "delete from td_loan where left(date,10) = '" _
            & TDBGrid1.Columns("date").CellText(item) & "' and employee_code = '" _
            & TDBGrid1.Columns("employee_code").CellText(item) & "' and flag_loan = '" & IIf(optCompany.Value, 0, 1) & "'"
        CnG.Execute SQL
        
        SQL = "delete from tm_loan where left(date,10) = '" _
            & TDBGrid1.Columns("date").CellText(item) & "' and employee_code = '" _
            & TDBGrid1.Columns("employee_code").CellText(item) & "' and flag_loan = '" & IIf(optCompany.Value, 0, 1) & "'"
        CnG.Execute SQL
                
    Next
    CnG.CommitTrans
    TDBGrid2.DataSource = Nothing
    Call load_data_salary
    
    MsgBox i & " loan's data are successfully deleted", vbInformation, headerMSG
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
    vSetData = 1
    
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "No Data selected!", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    With rsLoan
        DTPicker_date = .Fields("date").Value
        txt_employee_code = .Fields("employee_code").Value
        
        txt_nik = .Fields("nik").Value
        txt_employee_name = .Fields("employee_name").Value
        txt_description = .Fields("description").Value
        txt_loan_value = .Fields("loan_value").Value
        txt_interest = .Fields("loan_interest").Value
        txt_instalment_times = .Fields("installment_time").Value
        DTPicker_instalment_start = .Fields("installment_start").Value
        DTPicker_instalment_end = .Fields("installment_end").Value
        
        
'        txt_company_code = .Fields("company_code").Value
'        txt_item_code = .Fields("salary_item_code").Value
'        txt_item_name = .Fields("salary_item_name").Value
'        txt_flag_main_salary = .Fields("flag_main_salary").Value
'        txt_flag_sign = .Fields("flag_sign").Value
'        txt_flag_detail = .Fields("flag_detail").Value
'        txt_flag_per_presence = .Fields("flag_per_presence").Value
'        txt_default_value = .Fields("default_value").Value
'        txt_item_code = .Fields("item_code").Value
'        txt_item_name = .Fields("item_name").Value
'        txt_item_type = .Fields("item_type").Value
'        txt_item_type_name = .Fields("item_type_name").Value
'        .Fields("loan_total").Value = Val(DropAllComma(txt_loan_total))
'        .Fields("installment_value").Value = Val(DropAllComma(txt_instalment_value))
'        DTPicker_instalment_start = .Fields("installment_start").Value
'        .Fields("installment_end").Value = DTPicker_instalment_end
        
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

Private Sub insert_new_data()
Dim i As Integer
Dim vDesc As String

On Error GoTo Err

    CnG.BeginTrans

    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute strsql
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    SQL = "INSERT INTO tm_loan(date,company_code,flag_loan,employee_code,description," & _
            "loan_value,loan_interest,loan_total,installment_time,installment_value," & _
            "installment_start,installment_end,entry_date,entry_user) " & _
          "VALUES (" & _
            "'" & Format(DTPicker_date.Value, "yyyy-MM-dd") & "','" & TDBCombo_company.Text & "','" & IIf(optCompany.Value, 0, 1) & "','" & txt_employee_code.Text & "'," & _
            "'" & txt_description.Text & "','" & Val(DropAllComma(txt_loan_value)) & "'," & _
            "'" & Val(DropAllComma(txt_interest)) & "','" & Val(DropAllComma(txt_loan_total)) & "'," & _
            "'" & Val(DropAllComma(txt_instalment_times)) & "','" & Val(DropAllComma(txt_instalment_value)) & "'," & _
            "'" & Format(DTPicker_instalment_start.Value, "yyyy-MM-01") & "','" & Format(DTPicker_instalment_end.Value, "yyyy-MM-01") & "'," & _
            "now(),'" & LOGIN_NAME & "')"
    CnG.Execute SQL

    For i = 1 To Val(txt_instalment_times.Text)
        If optCompany.Value Then
            vDesc = "Angsuran Bulan " & Format(DateAdd("m", i - 1, DTPicker_instalment_start), "MMMM yyyy")
        Else
            vDesc = "Angsuran Koperasi Bulan " & Format(DateAdd("m", i - 1, DTPicker_instalment_start), "MMMM yyyy")
        End If
        SQL = "INSERT INTO td_loan(date,flag_loan,employee_code,sequence_number,installment_month," & _
                "installment_amount,installment_pay,installment_pay_date,flag_paid,flag_close,description) " & _
              "VALUES (" & _
                "'" & Format(DTPicker_date.Value, "yyyy-MM-dd") & "','" & IIf(optCompany.Value, 0, 1) & "','" & Trim(txt_employee_code.Text) & "'," & _
                "'" & i & "','" & Format(DateAdd("m", i - 1, DTPicker_instalment_start), "yyyy-MM-dd") & "'," & _
                "'" & Val(DropAllComma(txt_instalment_value)) & "','" & Val(DropAllComma(txt_instalment_value)) & "'," & _
                "'" & Format(DateAdd("m", i - 1, DTPicker_instalment_start), "yyyy-MM-01") & "',0,0,'" & vDesc & "')"
        CnG.Execute SQL
    Next i

    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
Dim i As Integer
Dim vDesc As String

'On Error GoTo Err

    CnG.BeginTrans
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    If v_value <> txt_loan_value.Text Or v_interest <> txt_interest.Text Or v_installment_time <> txt_instalment_times.Text Then
        strsql = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute strsql
    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    SQL = "UPDATE tm_loan SET date = '" & Format(DTPicker_date.Value, "yyyy-MM-dd") & "'," & _
            "company_code = '" & TDBCombo_company.Text & "'," & _
            "flag_loan = '" & IIf(optCompany.Value, 0, 1) & "'," & _
            "employee_code = '" & Trim(txt_employee_code.Text) & "'," & _
            "description = '" & Trim(txt_description.Text) & "'," & _
            "loan_value = '" & Val(DropAllComma(txt_loan_value.Text)) & "'," & _
            "loan_interest = '" & Val(DropAllComma(txt_interest)) & "'," & _
            "loan_total = '" & Val(DropAllComma(txt_loan_total)) & "'," & _
            "installment_time = '" & Int(DropAllComma(txt_instalment_times)) & "'," & _
            "installment_value = '" & Val(DropAllComma(txt_instalment_value)) & "'," & _
            "installment_start = '" & Format(DTPicker_instalment_start.Value, "yyyy-MM-01") & "'," & _
            "installment_end = '" & Format(DTPicker_instalment_end.Value, "yyyy-MM-01") & "'," & _
            "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
          "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
            "AND date(date) = '" & Format(DTPicker_date.Value, "yyyy-mm-dd") & "' " & _
            "and company_code= '" & TDBCombo_company.Columns("company_code").Value & "'"
    CnG.Execute SQL
    
    CnG.Execute "delete from td_loan where left(date,10) = '" _
        & Format(DTPicker_date.Value, "yyyy-mm-dd") & "' and employee_code = '" _
        & txt_employee_code.Text & "'"
    
    For i = 1 To Val(txt_instalment_times.Text)
        If optCompany.Value Then
            vDesc = "Angsuran Bulan " & Format(DateAdd("m", i - 1, DTPicker_instalment_start), "MMMM yyyy")
        Else
            vDesc = "Angsuran Koperasi Bulan " & Format(DateAdd("m", i - 1, DTPicker_instalment_start), "MMMM yyyy")
        End If
        SQL = "INSERT INTO td_loan(date,flag_loan,employee_code,sequence_number,installment_month," & _
                "installment_amount,installment_pay,installment_pay_date,flag_paid,flag_close,description) " & _
              "VALUES (" & _
                "'" & Format(DTPicker_date.Value, "yyyy-MM-dd") & "','" & IIf(optCompany.Value, 0, 1) & "','" & Trim(txt_employee_code.Text) & "'," & _
                "'" & i & "','" & Format(DateAdd("m", i - 1, DTPicker_instalment_start), "yyyy-MM-dd") & "'," & _
                "'" & Val(DropAllComma(txt_instalment_value)) & "','" & Val(DropAllComma(txt_instalment_value)) & "'," & _
                "'" & Format(DateAdd("m", i - 1, DTPicker_instalment_start), "yyyy-MM-01") & "',0,0,'" & vDesc & "')"
        CnG.Execute SQL
    Next i
    
    CnG.CommitTrans
    
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
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
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is DTPicker Then
            If Not LCase(Ctr.name) = "dtpicker_periode" And _
               Not LCase(Ctr.name) = "dtpicker_from" And _
               Not LCase(Ctr.name) = "dtpicker_to" Then Ctr.Value = Now
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
    txt_nik = ""
    txt_employee_code = ""
    txt_employee_name = ""
    DTPicker_date.Enabled = True
    DTPicker_date.Value = Now
    txt_description = ""
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        If TDBCombo_company.Text = "" Then
            MsgBox "Company Is Not Selected!", vbExclamation, headerMSG
            TDBCombo_company.SetFocus
            
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        Call clear_view_data
        fra_entry.Visible = True
        txt_nik.Enabled = True
        TDBGrid1.Enabled = False
        Call set_new_data
        
        If txt_nik.Enabled = True Then
            txt_nik.SetFocus
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        fra_entry.Visible = False
        fraCloseLoan.Visible = False
        TDBGrid1.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        txt_nik.Enabled = False
        DTPicker_date.Enabled = False
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

Private Sub cmdSave_Close_Click()
Dim i As Integer
Dim vDesc As String
     
    vDesc = "Close Loan On " & Format(TDBGrid2.Columns("installment_month").Value, "MM-yyyy")
    i = MsgBox("Are you sure want to close this loan from " _
        & Format(TDBGrid2.Columns("installment_month").Value, "MM-yyyy") & " ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    If TDBGrid2.Columns("flag_paid").Value <> 0 Then
        MsgBox "Installment Already Paid...", vbExclamation, headerMSG
        Exit Sub
    End If
    
    SQL = "UPDATE td_loan SET installment_amount = '" & Val(DropAllComma(lblPaid.Caption)) & "', " & _
            "installment_pay = '" & Val(DropAllComma(lblPaid.Caption)) & "', " & _
            "flag_paid = 1,flag_close = 1,description = '" & vDesc & "' " & _
          "WHERE YEAR(installment_month) = '" & Format(TDBGrid2.Columns("installment_month").Value, "yyyy") & "' " & _
            "AND MONTH(installment_month) = '" & Format(TDBGrid2.Columns("installment_month").Value, "MM") & "' " & _
            "AND employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND DATE(date) = '" & Format(TDBGrid1.Columns("date").Value, "yyyy-MM-dd") & "'"
    CnG.Execute SQL
    
    SQL = "DELETE FROM td_loan " & _
          "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND DATE(date) = '" & Format(TDBGrid1.Columns("date").Value, "yyyy-MM-dd") & "' " & _
            "AND flag_paid = 0"
    CnG.Execute SQL
    
    MsgBox "Save Successfully...", vbInformation, headerMSG
    fraCloseLoan.Visible = False
    
    Call load_data_salary
End Sub

Private Sub cmdSearch_Click()
    Call load_data_salary
End Sub

Private Sub DTPicker_Periode_Change()
    Call getPeriode(DTPicker_Periode.Value, DTPicker_from, DTPicker_to)
End Sub

Private Sub Form_Load()
    Call createGridKar
    Call load_data_company
    oClause = ""
    
    DTPicker_from.Value = Now
    DTPicker_to.Value = Now
    DTPicker_Periode.Value = Now
    optCompany.Value = True
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    timer1.Enabled = True
    
    optCompany.Value = True
End Sub

Private Sub optCompany_Click()
    load_data_salary
            
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        TDBGrid2.DataSource = Nothing
    End If
    
    If rscari.State Then rscari.Close
    SQL = "SELECT SUM(installment_amount) FROM td_loan " & _
          "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND DATE(date) = '" & Format(TDBGrid1.Columns("date").Value, "yyyy-MM-dd") & "' " & _
            "AND flag_loan = '" & IIf(optCompany.Value, 0, 1) & "' " & _
            "AND flag_paid <> 1"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        lbl_remains.Caption = "LOAN REMAINS = " & FormatNumber(rscari.Fields(0).Value)
    Else
        lbl_remains.Caption = "LOAN REMAINS = 0"
    End If
    rscari.Close
End Sub

Private Sub optKoperasi_Click()
    load_data_salary
    
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        TDBGrid2.DataSource = Nothing
    End If
    
    If rscari.State Then rscari.Close
    SQL = "SELECT SUM(installment_amount) FROM td_loan " & _
          "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND DATE(date) = '" & Format(TDBGrid1.Columns("date").Value, "yyyy-MM-dd") & "' " & _
            "AND flag_loan = '" & IIf(optCompany.Value, 0, 1) & "' " & _
            "AND flag_paid <> 1"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        lbl_remains.Caption = "LOAN REMAINS = " & FormatNumber(rscari.Fields(0).Value)
    Else
        lbl_remains.Caption = "LOAN REMAINS = 0"
    End If
    rscari.Close
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
        
        Call load_data_salary
    End If
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsLoan.Filter = adFilterNone
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

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_title = Nothing
End Sub

Private Sub TDBGrid1_FilterChange()
On Error GoTo Err

    Dim i As Integer
    
    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields
    
    rsLoan.Filter = getFilter()
    TDBGrid1.Col = i
    TDBGrid1.EditActive = True
    
    TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
    If TDBGrid1.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid1.Col = i
    End If

    Exit Sub

Err:
MsgBox "No Data found in this column " & vbCr _
& "or invalid data filter", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub load_data_salary()
    If rsLoan.State Then rsLoan.Close
    SQL = "select date(date) date,b.employee_code,b.employee_name, " _
        & "loan_value,loan_interest,loan_total,installment_time,a.description,b.nik," _
        & "a.installment_start,a.installment_end " _
        & "from tm_loan a join m_employee b on a.employee_code = b.employee_code " _
        & "where a.company_code = '" _
            & TDBCombo_company.Columns("company_code").Value & "' " _
            & "and flag_loan = '" & IIf(optCompany.Value, 0, 1) & "' " _
            & "and (date(date) between '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to.Value, "yyyy-MM-dd") & "') " & oClause
    rsLoan.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid1.DataSource = rsLoan
    
    If rsLoan.RecordCount > 0 Then
        Call load_data_detail
    Else
        lbl_remains.Caption = "LOAN REMAINS = 0"
    End If
End Sub

Private Sub load_data_detail()
    If rsLoanDetail.State Then rsLoanDetail.Close
    SQL = "select * from td_loan where Date(date) = '" _
        & Format(rsLoan.Fields("date").Value, "yyyy-mm-dd") & "' and employee_code ='" _
        & rsLoan.Fields("employee_code").Value & "' and flag_loan = '" & IIf(optCompany.Value, 0, 1) & "' order by sequence_number desc"
    rsLoanDetail.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid2.DataSource = rsLoanDetail
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
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
        
        If rscari.State Then rscari.Close
        SQL = "SELECT SUM(installment_amount) FROM td_loan " & _
          "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND DATE(date) = '" & Format(TDBGrid1.Columns("date").Value, "yyyy-MM-dd") & "' " & _
            "AND flag_loan = '" & IIf(optCompany.Value, 0, 1) & "' " & _
            "AND flag_paid <> 1"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            lbl_remains.Caption = "LOAN REMAINS = " & FormatNumber(rscari.Fields(0).Value)
        Else
            lbl_remains.Caption = "LOAN REMAINS = 0"
        End If
        rscari.Close
    End If
End Sub

Private Sub timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
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

    i = ((Val(DropAllComma(txt_interest)) / 100) * Val(DropAllComma(txt_loan_value)))
    t = Val(DropAllComma(txt_instalment_times))
    j = Val(DropAllComma(txt_loan_value))
    
    txt_loan_total = FormatNumber((t * i) + j)
End Sub

Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "Employee Code", 1500, lgAlignCenterCenter, , , , , , , True
      .AddColumn "Name", 3000, , , , , , , , , True
      .AddColumn "employee_code", 2000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        
        vParam = IIf(DEPARTMENT_CODE <> "" And DIVISION_CODE = "", "a.department_code = '" & DEPARTMENT_CODE & "'", IIf(DEPARTMENT_CODE = "" And DIVISION_CODE = "", "a.company_code = '" & COMPANY_CODE & "'", "a.department_code = '" & DEPARTMENT_CODE & "' AND a.division_code = '" & DIVISION_CODE & "'"))
        
        If LOGIN_LEVEL = 100 Then
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%')"
        Else
            SQL = "select nik,employee_name,employee_code " & _
                     "from m_employee a " & _
                     "WHERE flag_active <> 0 AND company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND " & vParam & " " & _
                        "AND (nik LIKE '%" & txt_nik.Text & "%' " & _
                            "OR employee_name LIKE '%" & txt_nik.Text & "%') " & _
                        "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0))"
        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME & vbTab & rs!employee_code
                rs.MoveNext
            Wend
            LynxGrid2.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code.Text = rs!employee_code
                txt_employee_name.Text = rs!EMPLOYEE_NAME
                txt_nik.Text = rs!nik
'                TDBCombo1.SetFocus
            Else
                LynxGrid2.Visible = True
                LynxGrid2.SetFocus
            End If
        Else
            
        End If
        rs.Close
    Else
        If LynxGrid2.Rows > 0 Then
            txt_nik.Text = LynxGrid2.CellText(LynxGrid2.Row, 0)
            txt_employee_name.Text = LynxGrid2.CellText(LynxGrid2.Row, 1)
            txt_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 2)
        End If
        LynxGrid2.Visible = False
    End If
End Sub

Private Sub LynxGrid2_DblClick()
    isiGridKar (2)
End Sub

Private Sub LynxGrid2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid2.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
    End If
End Sub

Private Sub LynxGrid2_LostFocus()
    LynxGrid2.Visible = False
End Sub

Private Sub txt_nik_Change()
    If txt_nik.Text = "" Then
        txt_employee_code.Text = ""
        txt_employee_name.Text = ""
    End If
End Sub

Private Sub txt_nik_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmdBrowse_Click()
    isiGridKar (1)
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid1.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid1.Columns(ColIndex).DataField
    Call load_data_salary

End Sub

Private Sub tutupKoperasi()
    If optInterestFrom(0).Value Then
        lblPaid.Caption = ((Val(txtInterest.Text) / 100) * Val(DropAllComma(lblLoanBasic.Caption))) + Val(DropAllComma(lblLoanRemains.Caption))
        lblPaid.Caption = FormatNumber(lblPaid.Caption)
    ElseIf optInterestFrom(1).Value Then
        lblPaid.Caption = ((Val(txtInterest.Text) / 100) * Val(DropAllComma(lblLoanRemains.Caption))) + Val(DropAllComma(lblLoanRemains.Caption))
        lblPaid.Caption = FormatNumber(lblPaid.Caption)
    End If
End Sub

Private Sub txtInterest_Change()
    Call tutupKoperasi
End Sub

Private Sub optInterestFrom_Click(Index As Integer)
    Call tutupKoperasi
End Sub