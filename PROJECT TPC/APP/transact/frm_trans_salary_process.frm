VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_trans_salary_process 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PPH 21 SALARY PROCESS"
   ClientHeight    =   6690
   ClientLeft      =   -15
   ClientTop       =   240
   ClientWidth     =   10815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_trans_salary_process.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prj_tpc.LynxGrid LynxGrid2 
      Height          =   2355
      Left            =   1590
      TabIndex        =   26
      Top             =   4140
      Visible         =   0   'False
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   4154
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
      Appearance      =   0
      ColumnHeaderSmall=   0   'False
      TotalsLineShow  =   0   'False
      FocusRowHighlightKeepTextForecolor=   0   'False
      ShowRowNumbers  =   0   'False
      ShowRowNumbersVary=   0   'False
      AllowColumnResizing=   -1  'True
   End
   Begin VB.Frame fra_entry 
      Height          =   3015
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   10335
      Begin VB.CheckBox chkManual 
         Caption         =   "RESET MANUAL ENTRY"
         Height          =   285
         Left            =   6660
         TabIndex        =   27
         Top             =   810
         Width           =   2625
      End
      Begin VB.TextBox txt_nik 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         Height          =   315
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   24
         Top             =   1770
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3270
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   23
         Top             =   1770
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_employee_code 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   5580
         TabIndex        =   22
         Top             =   1770
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.CheckBox chkEmployee 
         Caption         =   "PER EMPLOYEE"
         Height          =   255
         Left            =   1350
         TabIndex        =   20
         Top             =   1440
         Width           =   1665
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   3030
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         Top             =   390
         Width           =   6195
      End
      Begin MSComCtl2.DTPicker DTPicker_month 
         Height          =   300
         Left            =   1320
         TabIndex        =   1
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "MM-yyyy"
         Format          =   39845891
         CurrentDate     =   39278
      End
      Begin MSComCtl2.DTPicker DTPicker_periode_from 
         Height          =   300
         Left            =   4920
         TabIndex        =   2
         Top             =   810
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   39845891
         CurrentDate     =   39278
      End
      Begin MSComCtl2.DTPicker DTPicker_periode_to 
         Height          =   300
         Left            =   4920
         TabIndex        =   3
         Top             =   1170
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   39845891
         CurrentDate     =   39278
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   1320
         OleObjectBlob   =   "frm_trans_salary_process.frx":000C
         TabIndex        =   11
         Top             =   390
         Width           =   1695
      End
      Begin prj_tpc.vbButton CmdBrowse 
         Height          =   315
         Left            =   2760
         TabIndex        =   21
         Top             =   1770
         Visible         =   0   'False
         Width           =   405
         _ExtentX        =   714
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
         MICON           =   "frm_trans_salary_process.frx":1F72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblEmployee 
         AutoSize        =   -1  'True
         Caption         =   "EMPLOYEE"
         Height          =   195
         Left            =   420
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label lbl_company 
         AutoSize        =   -1  'True
         Caption         =   "COMPANY"
         Height          =   195
         Left            =   450
         TabIndex        =   12
         Top             =   450
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PERIODE TO"
         Height          =   195
         Left            =   3600
         TabIndex        =   9
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PERIODE FROM"
         Height          =   195
         Left            =   3600
         TabIndex        =   8
         Top             =   810
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "MONTH"
         Height          =   195
         Left            =   630
         TabIndex        =   7
         Top             =   840
         Width           =   540
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   5
      Top             =   5160
      Width           =   10335
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   0
         Top             =   0
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   5760
         TabIndex        =   13
         Top             =   750
         Visible         =   0   'False
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Load Data"
         Height          =   645
         Left            =   9300
         Picture         =   "frm_trans_salary_process.frx":1F8E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   -390
         Visible         =   0   'False
         Width           =   975
      End
      Begin prj_tpc.vbButton cmd_new 
         Height          =   705
         Left            =   450
         TabIndex        =   16
         Top             =   330
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
         MICON           =   "frm_trans_salary_process.frx":2298
         PICN            =   "frm_trans_salary_process.frx":22B4
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
         Left            =   4770
         TabIndex        =   17
         Top             =   330
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
         MICON           =   "frm_trans_salary_process.frx":3346
         PICN            =   "frm_trans_salary_process.frx":3362
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmd_process 
         Height          =   705
         Left            =   1440
         TabIndex        =   18
         Top             =   330
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Calculate"
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
         MICON           =   "frm_trans_salary_process.frx":43F4
         PICN            =   "frm_trans_salary_process.frx":4410
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdSearch 
         Height          =   705
         Left            =   2790
         TabIndex        =   19
         Top             =   330
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
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
         MICON           =   "frm_trans_salary_process.frx":54A2
         PICN            =   "frm_trans_salary_process.frx":54BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdImport 
         Height          =   705
         Left            =   3780
         TabIndex        =   28
         Top             =   330
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Import"
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
         MICON           =   "frm_trans_salary_process.frx":6550
         PICN            =   "frm_trans_salary_process.frx":656C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblKet 
         Height          =   255
         Left            =   5790
         TabIndex        =   15
         Top             =   180
         Visible         =   0   'False
         Width           =   5145
      End
      Begin VB.Label Label3 
         Caption         =   "Please Click New"
         Height          =   255
         Left            =   5790
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   4125
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8281
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "kd company"
      Columns(0).DataField=   "company_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "COMPANY"
      Columns(1).DataField=   "company_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "MONTH"
      Columns(2).DataField=   "month_"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "PERIODE FROM"
      Columns(3).DataField=   "periode_from_"
      Columns(3).NumberFormat=   "dd-MM-yyyy"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "PERIODE TO"
      Columns(4).DataField=   "periode_to_"
      Columns(4).NumberFormat=   "dd-MM-yyyy"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=8811"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=8731"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2646"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2566"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2646"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2566"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2646"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2566"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
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
      Caption         =   "LIST OF SALARY PROCESSED"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
End
Attribute VB_Name = "frm_trans_salary_process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsCompany As New ADODB.Recordset
Dim rsSalProses As New ADODB.Recordset

Dim str1 As String
Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Dim v_15 As Double, v_20 As Double
Dim v_30 As Double, v_40 As Double
Dim v_60 As Double
Dim vTotOT As Double
Dim vFlagHoliday As Integer

Dim vOTMax_Meal As Double
Dim vFlagMeal As Integer
Dim vOTMax_Trans As Double
Dim vFlagTrans As Integer

Dim vTotMeal As Double
Dim vTotTransport As Double

Dim vParam As String

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_validate_exist_new = False
    
    'str_sql = "select count(income_code) as rec_count from m_other_income where income_code = '" _
    '& Replace$(Trim$(txt_income_code), Chr$(39), Chr$(96)) & "'"
    'rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    '
    'If rs.Fields("rec_count").Value > 0 Then
    '    check_validate_exist_new = True
    '    Exit Function
    'End If
End Function

Private Sub check_invalid()
    MsgBox "Data found!", vbCritical, headerMSG
    DTPicker_month.Value = Null
    If DTPicker_month.Enabled = True Then DTPicker_month.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False

    If Not DTPicker_month.Value = rsSalProses.Fields("month").Value And _
    check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True

    'If Trim(txt_income_code) = "" Then
    '    MsgBox "Department Code is empty!", vbOKOnly + vbInformation, headerMSG
    '    txt_income_code.SetFocus
    '    check_validate_new = False
    '    Exit Function
    'End If
    '
    'If Trim(txt_income_name) = "" Then
    '    MsgBox "Department Name is empty!", vbOKOnly + vbInformation, headerMSG
    '    txt_income_name.SetFocus
    '    check_validate_new = False
    '    Exit Function
    'End If
End Function

Private Sub cmd_refresh_Click()
    Call load_data
End Sub

Private Sub cmdCancel_Click()
    int_mode = 0
    'Call load_mode
End Sub

Private Sub chkEmployee_Click()
    If chkEmployee.Value = 0 Then
        lblEmployee.Visible = False
        
        txt_nik.Visible = False
        txt_employee_name.Visible = False
        cmdBrowse.Visible = False
    Else
        lblEmployee.Visible = True
        
        txt_nik.Visible = True
        txt_employee_name.Visible = True
        cmdBrowse.Visible = True
    End If
End Sub

Private Sub cmdDelete_Click()
    '    Dim i As Integer
    '
    '    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
    '        MsgBox "No Data selected!", vbInformation, headerMSG
    '        Exit Sub
    '    End If
    '
    '    i = MsgBox("Are you sure want to delete data '" _
    '        & Format(TDBGrid1.Columns("month").Value, "mm-yyyy") & "' ?", vbYesNo + vbQuestion, headerMSG)
    '    If Not i = vbYes Then Exit Sub
    '
    '    CnG.Execute "delete from h_d_salary where left(month,7) = '" & Format(rsSalProses.Fields("month").Value, "yyyy-mm") & "'"
    '    CnG.Execute "delete from h_m_salary where left(month,7) = '" & Format(rsSalProses.Fields("month").Value, "yyyy-mm") & "'"
    Call load_data
End Sub

Private Sub cmd_new_Click()
    fra_entry.Visible = True
    DTPicker_month = Now
    DTPicker_periode_from = Now
    DTPicker_periode_to = Now
    chkEmployee.Value = 0
    cmd_new.Enabled = False
    cmd_process.Enabled = True
    
    timer1.Enabled = True
End Sub

Private Sub cmd_process_Click()
Dim rsEmployee As New ADODB.Recordset
Dim d1, d2, dx As Date
Dim i As Integer

Dim bulan As String
Dim tgl As String
Dim v_tgl_akhir As Date
Dim v_tgl_mc As Date
Dim v_end_mc As Date
Dim int_month As Integer
Dim int_year As Integer
Dim bln_awal, bln_akhir, thn_awal, thn_akhir As String
Dim v_pph21_type As String
Dim v_jstk_type As String
Dim rs As New ADODB.Recordset

    ''+++++++++++++++++++++++++++++ Cek Tanggal Periode OT To +++++++++++++++++++++++++++++++
    'If Format(DTPicker_periode_to_ot, "yyyy-MM") <> Format(DTPicker_month, "yyyy-MM") And TDBCombo_company.Text = "GPN" Then
    '    MsgBox "Periode To Doesn't Match With Month!" & Chr(13) & _
    '        "Please Check Your Overtime Periode Date.", vbExclamation, headerMSG
    '    Exit Sub
    'End If
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    'entry manual will be resetted for this employee
    
    If chkEmployee.Value = 1 Then
        SQL = "SELECT * FROM h_salary_manual WHERE month = '" & Format(DTPicker_month, "yyyy-MM") & "' " & _
                "AND employee_code = '" & txt_employee_code.Text & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            i = MsgBox("Entry Manual will be resetted for this employee..." _
                        & Chr(13) & "Are you sure to proccess salary for this employee?", vbYesNo + vbQuestion, headerMSG)
            If Not i = vbYes Then
                rscari.Close
                Exit Sub
            End If
        Else
            i = MsgBox("Are you sure that the data you enter is correct?", vbYesNo + vbQuestion, headerMSG)
            If Not i = vbYes Then
                rscari.Close
                Exit Sub
            End If
        End If
        rscari.Close
    Else
        i = MsgBox("Are you sure that the data you enter is correct?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
    End If
    
    d1 = Format(DTPicker_month.Value, "yyyy-MM-01"): dx = DateAdd("m", 1, d1)
    d2 = Format(d1, "yyyy-MM-") & Format(DateDiff("d", d1, dx), "0#")

    Call days_func(DTPicker_periode_from.Value, DTPicker_periode_to.Value)
    
    ProgressBar1.Visible = True
    Label3.Visible = True
    lblKet.Visible = True
    
    If chkEmployee.Value = 0 Then
        SQL = "select LAST_DAY('" & d1 & "') tgl_akhir, a.employee_code, a.employee_name," & _
            "a.no_jamsostek,a.npwp,a.end_working,a.flag_active,a.nik," & _
            "(SELECT pph21_type FROM m_salary_standard WHERE employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) pph21_type," & _
            "(SELECT jstk_type FROM m_salary_standard WHERE employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) jstk_type " & _
            "from m_employee a " & _
            "JOIN (SELECT employee_code FROM h_attendance WHERE DATE(att_date) BETWEEN '" & Format(DTPicker_periode_from, "yyyy-MM-dd") & "'  AND '" & Format(DTPicker_periode_to, "yyyy-MM-dd") & "' " & _
                 "GROUP BY employee_code) d ON a.employee_code = d.employee_code " & _
            "where a.company_code = '" & TDBCombo_company.Text & "' AND CASE WHEN a.end_working = '00:00:00' OR ISNULL(a.end_working) THEN DATE(NOW()) " & _
                "ELSE DATE(a.end_working) END > '" & Format(DTPicker_periode_from, "yyyy-MM-dd") & "' "
    Else
        SQL = "select LAST_DAY('" & d1 & "') tgl_akhir, a.employee_code, a.employee_name," & _
            "a.no_jamsostek,a.npwp,a.end_working,a.flag_active,a.nik," & _
            "(SELECT pph21_type FROM m_salary_standard WHERE employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) pph21_type," & _
            "(SELECT jstk_type FROM m_salary_standard WHERE employee_code = a.employee_code AND date(salary_date) <= '" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1) jstk_type " & _
            "from m_employee a " & _
            "JOIN (SELECT employee_code FROM h_attendance WHERE DATE(att_date) BETWEEN '" & Format(DTPicker_periode_from, "yyyy-MM-dd") & "'  AND '" & Format(DTPicker_periode_to, "yyyy-MM-dd") & "' " & _
                 "GROUP BY employee_code) d ON a.employee_code = d.employee_code " & _
            "where a.company_code = '" & TDBCombo_company.Text & "' AND CASE WHEN a.end_working = '00:00:00' OR ISNULL(a.end_working) THEN DATE(NOW()) " & _
                "ELSE DATE(a.end_working) END > '" & Format(DTPicker_periode_from, "yyyy-MM-dd") & "' " & _
                "AND a.employee_code = '" & txt_employee_code.Text & "'"
    End If
    rsEmployee.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    If rsEmployee.RecordCount > 0 Then
        v_tgl_akhir = rsEmployee!tgl_akhir
'        v_end_mc = IIf(IsNull(rsEmployee!end_mc), "00:00:00", rsEmployee!end_mc)
        v_pph21_type = IIf(IsNull(rsEmployee!pph21_type), "STD", rsEmployee!pph21_type)
        v_jstk_type = IIf(IsNull(rsEmployee!jstk_type), "STD", rsEmployee!jstk_type)
    End If
        
'    CnG.BeginTrans
            
    If rsEmployee.RecordCount > 0 Then
        rsEmployee.MoveFirst
        While Not rsEmployee.EOF
            
            If chkManual.Value = 0 Then
                If chkEmployee.Value = 0 Then
                    str1 = "DELETE a FROM h_salary a JOIN m_employee b on a.employee_code = b.employee_code " & _
                        "WHERE LEFT(a.month,7) = '" & Format(d1, "yyyy-MM") & "' " & _
                            "AND a.company_code = '" & TDBCombo_company.Text & "' " & _
                            "AND a.employee_code = '" & rsEmployee!employee_code & "' " & _
                            "AND ifnull(a.flag_manual,0) = 0"
                    CnG.Execute str1
                Else
                    str1 = "DELETE a FROM h_salary a JOIN m_employee b on a.employee_code = b.employee_code " & _
                        "WHERE LEFT(a.month,7) = '" & Format(d1, "yyyy-MM") & "' " & _
                            "AND a.company_code = '" & TDBCombo_company.Text & "' " & _
                            "AND a.employee_code = '" & rsEmployee!employee_code & "'"
                    CnG.Execute str1
                End If
            Else
                str1 = "DELETE a FROM h_salary a JOIN m_employee b on a.employee_code = b.employee_code " & _
                    "WHERE LEFT(a.month,7) = '" & Format(d1, "yyyy-MM") & "' " & _
                        "AND a.company_code = '" & TDBCombo_company.Text & "' " & _
                        "AND a.employee_code = '" & rsEmployee!employee_code & "'"
                CnG.Execute str1
            End If
            
            str1 = "DELETE FROM t_spl_auto " & _
                   "WHERE Date(date) BETWEEN '" & Format(DTPicker_periode_from.Value, "yyyy-mm-dd") & "' AND '" & Format(DTPicker_periode_to.Value, "yyyy-mm-dd") & "' " & _
                        "AND employee_code = '" & rsEmployee!employee_code & "'"
            CnG.Execute str1
            
            rsEmployee.MoveNext
        Wend
    End If
            
    If rsEmployee.RecordCount > 0 Then
        ProgressBar1.Max = rsEmployee.RecordCount
        ProgressBar1.Value = 0
        
        Label3.Visible = True
        ProgressBar1.Visible = True
        
        
        '------------------------------- Hitung Overtime -----------------------------
        rsEmployee.MoveFirst
        While Not rsEmployee.EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblKet.Caption = "CALCULATE OVERTIME....."
            Label3.Caption = "( " & rsEmployee!nik & " ) " & rsEmployee!EMPLOYEE_NAME
            
            Call auto_overtime(rsEmployee!employee_code, DTPicker_periode_from.Value, DTPicker_periode_to.Value)
            
            rsEmployee.MoveNext
            DoEvents
        Wend
        '-----------------------------------------------------------------------------
        
        '--------------------------------- Hitung Late -------------------------------
        rsEmployee.MoveFirst
        While Not rsEmployee.EOF
'            ProgressBar1.Value = ProgressBar1.Value + 1
'            lblKet.Caption = "CALCULATE OVERTIME....."
'            Label3.Caption = "( " & rsEmployee!nik & " ) " & rsEmployee!EMPLOYEE_NAME
            
            Call check_late(rsEmployee!employee_code, DTPicker_periode_from.Value, DTPicker_periode_to.Value)
            
            rsEmployee.MoveNext
            DoEvents
        Wend
        '-----------------------------------------------------------------------------
        
        If flagPLAuto() <> 0 Then
            '--------------------------------- Hitung PL -------------------------------
            rsEmployee.MoveFirst
            While Not rsEmployee.EOF
    '            ProgressBar1.Value = ProgressBar1.Value + 1
    '            lblKet.Caption = "CALCULATE OVERTIME....."
    '            Label3.Caption = "( " & rsEmployee!nik & " ) " & rsEmployee!EMPLOYEE_NAME
                
                Call check_pl(rsEmployee!employee_code, DTPicker_periode_from.Value, DTPicker_periode_to.Value)
                
                rsEmployee.MoveNext
                DoEvents
            Wend
            '-----------------------------------------------------------------------------
        End If
                       
        '------------------------------- Hitung Salary -------------------------------
        ProgressBar1.Value = 0
        rsEmployee.MoveFirst
        While Not rsEmployee.EOF
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblKet.Caption = "CALCULATE SALARY PROCCESS...."
            Label3.Caption = "( " & rsEmployee!nik & " ) " & rsEmployee!EMPLOYEE_NAME
            'Label11.Caption = "( " & rsemployee!employee_code & " ) " & rsemployee!employee_name
            
            '+++++++++++++++++++++++++++++++++++++++ MC +++++++++++++++++++++++++++++++++++++++
'            v_tgl_mc = IIf(IsNull(rsEmployee!start_mc), 0, rsEmployee!start_mc)
            int_month = month(v_tgl_akhir)
            int_year = year(v_tgl_akhir)
            
'            If rsEmployee!flag_active = "2" Then
'
'                SQL = "DELETE FROM h_attendance " & _
'                        "WHERE employee_code = '" & rs!employee_code & "' AND DATE(att_date) >= '" & Format(v_tgl_mc, "yyyy-MM-dd") & "' " & _
'                        "AND month(att_date) = '" & Format(DTPicker_month.Value, "yyyy-MM") & "' AND year(att_date) = '" & year(DTPicker_month.Value) & "'"
'                CnG.Execute SQL
'
'                rsEmployee.MoveFirst
'                If Format(v_tgl_mc, "yyyy-MM") = Format(v_tgl_akhir, "yyyy-MM") Then
'                    v_tgl_mc = DateValue(v_tgl_mc)
'                    While v_tgl_mc <= v_tgl_akhir
''                        v_tgl_mc = v_tgl_mc + 1
'                        SQL = "INSERT INTO h_attendance (employee_code, att_date," & _
'                            "shift_number, shift_code, start_time," & _
'                            "end_time," & _
'                            "flag_present,description, entry_date, absent_status) " & _
'                        "VALUES " & _
'                            "('" & rs!employee_code & "','" & Format(v_tgl_mc, "yyyy-MM-dd 00:00:00") & "'," & _
'                            "'STF', 'MC','" & Format(v_tgl_mc, "yyyy-MM-dd 08:00:00") & "'," & _
'                            "'" & Format(v_tgl_mc, "yyyy-MM-dd 17:00:00") & "'," & _
'                            "0,'MC', Now(),8)"
'
'                        CnG.Execute SQL
'                        'rs.MoveNext
'                        v_tgl_mc = v_tgl_mc + 1
'                    Wend
'                Else
''                    Dim i As Integer
'                    For i = 1 To Day(v_tgl_akhir)
''                        i = i + 1
''                        v_tgl_mc = Format(v_tgl_akhir, "yyyy-MM") + "-" + i
'                        v_tgl_mc = DateSerial(int_year, int_month, i)
'                        SQL = "INSERT INTO h_attendance (employee_code, att_date," & _
'                            "shift_number, shift_code, start_time," & _
'                            "end_time," & _
'                            "flag_present,description, entry_date,absent_status) " & _
'                        "VALUES " & _
'                            "('" & rs!employee_code & "','" & Format(v_tgl_mc, "yyyy-MM-dd 00:00:00") & "'," & _
'                            "1, 'MC','" & Format(v_tgl_mc, "yyyy-MM-dd 08:00:00") & "'," & _
'                            "'" & Format(v_tgl_mc, "yyyy-MM-dd 17:00:00") & "'," & _
'                            "0,'MC', Now(),8)"
'
'                        CnG.Execute SQL
'                    Next
'                End If
'            End If
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            
            Call proses_su(rsEmployee!employee_code, _
                Format(DTPicker_periode_from.Value, "yyyy-MM-dd"), Format(DTPicker_periode_to.Value, "yyyy-MM-dd"), _
                IIf(IsNull(rsEmployee!no_jamsostek), "", rsEmployee!no_jamsostek), _
                IIf(IsNull(rsEmployee!npwp), 0, rsEmployee!npwp), TDBCombo_company.Text, _
                IIf(IsNull(Format(rsEmployee!end_working, "yyyyMM")), "0", Format(rsEmployee!end_working, "yyyyMM")), _
                rsEmployee!flag_active, Format(DTPicker_month.Value, "yyyy-MM"), v_pph21_type, v_jstk_type)

            '++++++++++++++++++++++++++++++ Update Data Loan +++++++++++++++++++++++++++
            
            SQL = "UPDATE td_loan SET flag_paid = 1 " _
                & "Where employee_code = '" & rsEmployee!employee_code & "' " _
                & "AND Month(installment_month) = '" & month(DTPicker_month.Value) & "' " _
                & "AND Year(installment_month) = '" & year(DTPicker_month.Value) & "'"
            CnG.Execute (SQL)
            '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
            rsEmployee.MoveNext
            DoEvents
        Wend
        '-------------------------------------------------------------------------------
        
        SQL = "UPDATE h_salary a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                "Set a.salary_value = 0 " & _
                "WHERE a.month = LEFT('" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "', 7) AND a.salary_code = 'SU-052'"
        CnG.Execute SQL
        
        SQL = "DELETE FROM h_d_salary WHERE company_code = '" & TDBCombo_company.Text & "' and left(month,7) = '" & Format(DTPicker_month.Value, "yyyy-MM") & "'"
        CnG.Execute SQL
        
        SQL = "INSERT INTO h_d_salary (month,periode_from,periode_to,company_code,company_name) " & _
            "VALUES " & _
            "('" & Format(DTPicker_month.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_periode_from.Value, "yyyy-MM-dd") & "','" & Format(DTPicker_periode_to.Value, "yyyy-MM-dd") & "'," & _
            "'" & TDBCombo_company.Text & "','" & Replace(txt_company_name.Text, "'", "''") & "')"
        CnG.Execute SQL
        
        'Update Temp Salary Proses ++++++++++++++++++++++++++++++++
        SQL = "UPDATE temp_sal_proses set salary_proses = 1 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute (SQL)
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    End If
    
    MsgBox "Calculate Salary Process Success...!!", vbInformation, "Information"
    ProgressBar1.Visible = False
    Label3.Visible = False
    lblKet.Visible = False
    fra_entry.Visible = False
    TDBCombo_company.Text = ""
    txt_company_name.Text = ""
    cmd_new.Enabled = True
    cmd_process.Enabled = False
    
'    CnG.CommitTrans
    
    Call load_data
End Sub

Private Sub process_delete()
    CnG.Execute "delete from h_d_salary where left(month,7) = '" & Format(DTPicker_month, "yyyy-mm") & "'"
    CnG.Execute "delete from h_m_salary where left(month,7) = '" & Format(DTPicker_month, "yyyy-mm") & "'"
End Sub

Private Sub CmdExit_Click()
    Unload Me
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

Private Sub cmdImport_Click()
    frm_import_salary_manual.Show (1)
End Sub

Private Sub cmdSearch_Click()
    frm_list_salary.tgl1 = TDBGrid1.Columns("periode_from_").Value
    frm_list_salary.tgl2 = TDBGrid1.Columns("periode_to_").Value
    frm_list_salary.str_company_code = TDBGrid1.Columns("company_code").Value
    frm_list_salary.Show (1)
'    Call frm_list_salary.listSalary(Format(DTPicker_periode_from.Value, "yyyy-MM-dd"), _
'                            Format(DTPicker_periode_to.Value, "yyyy-MM-dd"), _
'                            Format(DTPicker_month.Value, "yyyy-MM"), _
'                            TDBCombo_company.Text)
End Sub

Private Sub DTPicker_month_Change()
    DTPicker_periode_from = Format(DateAdd("m", -1, DTPicker_month), "yyyy-MM-") & "21"
    DTPicker_periode_to = Format(DTPicker_month, "yyyy-MM-") & "20"
End Sub

Private Sub Form_Load()
    DTPicker_month.Value = Date
    DTPicker_periode_from.Value = Date
    DTPicker_periode_to.Value = Date
    cmd_new.Enabled = True
    cmd_process.Enabled = False
    
    Call load_data
    Call load_data_company
    Call createGridKar
    
    Call load_data_user_access(Me)
    int_mode = 0
    'Call load_mode
 '   timer1.Enabled = True
End Sub

Public Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub


Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsSalProses.Filter = adFilterNone
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
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields
    
    rsSalProses.Filter = getFilter()
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

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
    End If

    DTPicker_periode_from = Format(DateAdd("m", -1, DTPicker_month), "yyyy-MM-") & "21"
    DTPicker_periode_to = Format(DTPicker_month, "yyyy-MM-") & "20"
End Sub

Public Sub load_data()
    If rsSalProses.State Then rsSalProses.Close
    SQL = "select *, cast(left(month,7) as char) as month_, periode_from as periode_from_, periode_to as periode_to_ " & _
          "from h_d_salary order by month desc"
    rsSalProses.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsSalProses
End Sub

Public Sub proses_su(pEmployee_code As String, pTgl1 As String, _
    pTgl2 As String, pnoJamsostek As String, pNpwp As String, _
    pCompany_Code As String, pEndWorking As String, _
    pFlag_Active As Integer, pBulan As String, pPph21 As String, pJstk As String)
Dim SQL As String
'Dim rsemployee As New ADODB.Recordset
Dim i As Integer
Dim vIntManual As Integer
Dim clsCalcSUFormula As New clsCalcSUFormula

On Error GoTo Err
        SQL = "SELECT IFNULL(flag_manual,0) flag_manual FROM h_salary WHERE LEFT(MONTH,7) = LEFT('" & pTgl2 & "',7) AND employee_code = '" & pEmployee_code & "'"
        rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rscari.RecordCount > 0 Then
            vIntManual = rscari!flag_manual
        Else
            vIntManual = 0
        End If
        rscari.Close
        
        If vIntManual = 0 Then
            SQL = "DELETE FROM h_salary_manual WHERE LEFT(MONTH,7) = LEFT('" & pTgl2 & "',7) AND employee_code = '" & pEmployee_code & "'"
            CnG.Execute SQL
            
            SQL = "DELETE FROM h_salary WHERE flag_type = 'SU' AND LEFT(MONTH,7) = LEFT('" & pTgl2 & "',7) AND employee_code = '" & pEmployee_code & "'"
            CnG.Execute SQL
        End If

        SQL = "insert into h_salary " & _
            "(MONTH, employee_code, salary_code, company_code, salary_name," & _
            "date_from, date_to, flag_main_salary, flag_sign,flag_detail," & _
            "flag_use_formula, formula_salary_code, flag_ptkp, ptkp_salary_code, flag_pkp," & _
            "flag_pph21, pph21_number, flag_tax, tax_salary_code, flag_type," & _
            "flag_visible, salary_value, Description) " & _
            "SELECT " & _
                "LEFT('" & pTgl2 & "',7) AS MONTH, '" & pEmployee_code & "', salary_code," & _
                "'" & pCompany_Code & "', salary_name, '" & pTgl1 & "', '" & pTgl2 & "', " & _
                "flag_main_salary, flag_sign, flag_detail, flag_use_formula, " & _
                "formula_salary_code, flag_ptkp, ptkp_salary_code, flag_pkp," & _
                "flag_pph21, pph21_number, 0 AS flag_tax, '' AS tax_salary_code," & _
                "'SU' AS flag_type, flag_visible," & _
                "f_get_sum_dsum('" & pEmployee_code & "',salary_code,'" & pTgl1 & "','" & pTgl2 & "'," & _
                        "'" & pCompany_Code & "','" & pBulan & "'," & _
                        "'" & pEndWorking & "','" & pFlag_Active & "')," & _
                "Description " & _
            "FROM m_salary_summary;"
        CnG.Execute SQL

        'SQL = "CALL sp_calc_su_formula('" & pTgl2 & "','" & pEmployee_code & "');"
        Call clsCalcSUFormula.CalcSuFormula(pTgl2, pEmployee_code, pnoJamsostek, pNpwp, pCompany_Code, pPph21, pJstk)
        'CnG.Execute SQL
        
        Exit Sub

Err:
If Err.Number = "-2147217900" Then
    Resume Next
Else
    MsgBox Err.Description
End If
End Sub

Private Sub days_func(start_time As String, end_time As String)
Dim v_tgl_awal, v_tgl_akhir As Date

    v_tgl_awal = Format(start_time, "yyyy-MM-dd")
    v_tgl_akhir = Format(end_time, "yyyy-MM-dd")
    
    v_tgl_awal = DateValue(v_tgl_awal)
    v_tgl_akhir = DateValue(v_tgl_akhir)
    
    SQL = "delete from m_days where dt between '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' and '" & Format(v_tgl_akhir, "yyyy-MM-dd") & "'"
    CnG.Execute SQL
            
        While v_tgl_awal <= v_tgl_akhir
            SQL = "SELECT holiday_date FROM t_holiday WHERE date(holiday_date) = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "'"
            rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs.RecordCount > 0 Then
                SQL = "INSERT INTO m_days (dt,status,description) " & _
                      "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','L','HOLIDAY')"
                CnG.Execute SQL
            Else
                If Format(v_tgl_awal, "dddd") = "Sunday" Then
                    SQL = "INSERT INTO m_days (dt,status,description) " & _
                          "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','M','SUNDAY')"
                    CnG.Execute SQL
                ElseIf Format(v_tgl_awal, "dddd") = "Saturday" Then
                    SQL = "INSERT INTO m_days (dt,status,description) " & _
                          "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','S','SATURDAY')"
                    CnG.Execute SQL
                Else
                    SQL = "INSERT INTO m_days (dt,status,description) " & _
                          "VALUES ('" & Format(v_tgl_awal, "yyyy-MM-dd") & "','W','WORK DAY')"
                    CnG.Execute SQL
                End If
            End If
            rs.Close
            
            
            v_tgl_awal = v_tgl_awal + 1
        Wend
End Sub

Public Sub auto_overtime(str_employee_code As String, start_time As String, end_time As String)
Dim rs2 As New ADODB.Recordset
Dim v_tgl_awal, v_tgl_akhir As Date
Dim ot_otomatis As Double
Dim vFlagOT As Integer
Dim vGroupCode As String
Dim vSatOT As Double
Dim vWorkOT As Double
Dim vOTSPL As Double
Dim v_total_ot As Double

Dim vOTCode As String

Dim vFlagMeal As Integer
Dim vGetMeal As Integer
Dim vOTMax As Double
    
    vTotOT = 0
    
    v_tgl_awal = Format(start_time, "yyyy-MM-dd")
    v_tgl_akhir = Format(end_time, "yyyy-MM-dd")
    
    v_tgl_awal = DateValue(v_tgl_awal)
    v_tgl_akhir = DateValue(v_tgl_akhir)
    
    SQL = "SELECT * FROM m_days " & _
        "WHERE date(dt) BETWEEN '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' AND '" & Format(v_tgl_akhir, "yyyy-MM-dd") & "' " & _
        "ORDER BY dt"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            SQL = "SELECT a.flag_ot,a.group_code,a.saturday_ot,a.work_ot FROM m_shift_group a JOIN h_attendance b ON a.group_code = b.group_code " & _
                  "WHERE b.employee_code = '" & str_employee_code & "' AND Date(b.att_date) = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "'"
            rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            
            If rs2.RecordCount > 0 Then
                vFlagOT = IIf(IsNull(rs2!flag_ot), 0, rs2!flag_ot)
                vGroupCode = IIf(IsNull(rs2!group_code), 0, rs2!group_code)
                vSatOT = IIf(IsNull(rs2!saturday_ot), 0, rs2!saturday_ot)
                vWorkOT = IIf(IsNull(rs2!work_ot), 0, rs2!work_ot)
            Else
                vFlagOT = 0
                vGroupCode = ""
                vSatOT = 0
                vWorkOT = 0
            End If
            rs2.Close
            
'            SQL = "SELECT ot_spl,ot_15,ot_20,ot_30,ot_40,ot_60,total_ot,ot_code,meal_allowance,transport_allowance,flag_on_call " & _
'                  "FROM t_spl WHERE employee_code = '" & str_employee_code & "' " & _
'                    "AND Date(date) = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' AND flag_approval = 1"
'            rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'            Select Case rs!Status
'                Case "L"
'                    ot_otomatis = 0
'                Case "M"
'                    If vFlagOT = 1 Then
'                        If rs2.RecordCount > 0 Then
'                            ot_otomatis = 0
'                        Else
'                            ot_otomatis = 1
'                        End If
'                    Else
'                        ot_otomatis = 0
'                    End If
'                Case "S"
'                    If vFlagOT = 1 Then
'                        ot_otomatis = vSatOT
'                    Else
'                        ot_otomatis = 0
'                    End If
'                Case "W"
'                    If vFlagOT = 1 Then
'                        ot_otomatis = vWorkOT
'                    Else
'                        ot_otomatis = 0
'                    End If
'            End Select
'
'            If rs2.RecordCount > 0 Then
'                vOTSPL = rs2!ot_spl
'                vOTCode = rs2!ot_code
''                vFlagHoliday = rs2!flag_holiday
'            Else
'                vOTSPL = 0
'                vOTCode = ""
''                vFlagHoliday = 0
'            End If
'
''            If vFlagHoliday = 0 Then
''                vTotOT = ot_otomatis + vOTSPL
''
''                SQL = "SELECT ot_max, flag_meal FROM m_pref_ot"
''                rsCari.Open SQL, CnG, adOpenForwardOnly
''
''                If rs.RecordCount > 0 Then
''                    vOTMax = rsCari!ot_max
''                    vFlagMeal = rsCari!flag_meal
''                Else
''                    vOTMax = 0
''                    vFlagMeal = 0
''                End If
''                rsCari.Close
''
''                If vFlagMeal <> 0 Then
''                    If vTotOT > vOTMax Then
''                        vGetMeal = 1
''                    Else
''                        vGetMeal = 0
''                    End If
''                Else
''                    vGetMeal = 0
''                End If
''
''                Call hitungLembur
''
''                v_total_ot = (1.5 * v_15) + (2 * v_20) + (3 * v_30) + (4 * v_40) + (6 * v_60)
''
''                SQL = "INSERT INTO t_spl_auto(company_code,employee_code,date," & _
''                          "ot_spl,ot_15,ot_20,ot_30,ot_40,total_ot,ot_code,meal_allowance,entry_date,entry_user) " & _
''                        "VALUES( " & _
''                          "'" & TDBCombo_company.Text & "','" & str_employee_code & "'," & _
''                          "'" & Format(v_tgl_awal, "yyyy-MM-dd") & "'," & _
''                          "'" & vTotOT & "','" & v_15 & "'," & _
''                          "'" & v_20 & "','" & v_30 & "','" & v_40 & "'," & _
''                          "'" & v_total_ot & "','" & vOTCode & "','" & vGetMeal & "',now(),'" & LOGIN_NAME & "')"
''
''            Else
'
''            SQL = "SELECT ot_max, flag_meal FROM m_pref_ot"
''            rscari.Open SQL, CnG, adOpenForwardOnly
''
''            If rs.RecordCount > 0 Then
''                vOTMax = rscari!ot_max
''                vFlagMeal = rscari!flag_meal
''            Else
''                vOTMax = 0
''                vFlagMeal = 0
''            End If
''            rscari.Close
''
''            If vFlagMeal <> 0 Then
''                If vTotOT > vOTMax Then
''                    vGetMeal = 1
''                Else
''                    vGetMeal = 0
''                End If
''            Else
''                vGetMeal = 0
''            End If
            
            SQL = "SELECT SUM(ot_spl) ot_spl, SUM(ot_15) tot_ot_15, SUM(ot_20) tot_ot_20, SUM(ot_30) tot_ot_30," & _
                    "SUM(ot_40) tot_ot_40, SUM(ot_60) tot_ot_60, SUM(total_ot) total_ot,ot_code," & _
                    "SUM(meal_allowance) meal_allowance,SUM(transport_allowance) transport_allowance,flag_on_call," & _
                    "SUM(flag_shift2) flag_shift2,SUM(flag_shift3) flag_shift3 " & _
                  "FROM t_spl WHERE employee_code = '" & str_employee_code & "' " & _
                    "AND Date(date) = '" & Format(v_tgl_awal, "yyyy-MM-dd") & "' AND flag_approval = 1 " & _
                  "GROUP BY employee_code"
            rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

            Select Case rs!Status
                Case "L"
                    ot_otomatis = 0
                Case "M"
                    If vFlagOT = 1 Then
                        If rs2.RecordCount > 0 Then
                            ot_otomatis = 0
                        Else
                            ot_otomatis = 1
                        End If
                    Else
                        ot_otomatis = 0
                    End If
                Case "S"
                    If vFlagOT = 1 Then
                        ot_otomatis = vSatOT
                    Else
                        ot_otomatis = 0
                    End If
                Case "W"
                    If vFlagOT = 1 Then
                        ot_otomatis = vWorkOT
                    Else
                        ot_otomatis = 0
                    End If
            End Select

            If rs2.RecordCount > 0 Then
                vOTSPL = rs2!ot_spl
                vOTCode = rs2!ot_code
'                vFlagHoliday = rs2!flag_holiday
            Else
                vOTSPL = 0
                vOTCode = ""
'                vFlagHoliday = 0
            End If
                        
            If rs2.RecordCount > 0 Then
                Call hitungLembur(rs2!tot_ot_15, rs2!tot_ot_20, rs2!tot_ot_30, rs2!tot_ot_40, rs2!tot_ot_60, _
                                rs2!ot_code, str_employee_code)

                SQL = "INSERT INTO t_spl_auto(company_code,employee_code,date," & _
                          "ot_spl,ot_15,ot_20,ot_30,ot_40,ot_60,total_ot,ot_code,meal_allowance,transport_allowance," & _
                          "flag_on_call,flag_shift2,flag_shift3,entry_date,entry_user) " & _
                        "VALUES( " & _
                          "'" & TDBCombo_company.Text & "','" & str_employee_code & "'," & _
                          "'" & Format(v_tgl_awal, "yyyy-MM-dd") & "'," & _
                          "'" & rs2!ot_spl & "','" & v_15 & "'," & _
                          "'" & v_20 & "','" & v_30 & "','" & v_40 & "','" & v_60 & "'," & _
                          "'" & vTotOT & "','" & vOTCode & "','" & rs2!meal_allowance & "','" & IIf(IsNull(rs2!transport_allowance), 0, rs2!transport_allowance) & "'," & _
                          "'" & IIf(IsNull(rs2!flag_on_call), 0, rs2!flag_on_call) & "'," & _
                          "'" & IIf(IsNull(rs2!flag_shift2), 0, rs2!flag_shift2) & "'," & _
                          "'" & IIf(IsNull(rs2!flag_shift3), 0, rs2!flag_shift3) & "',now(),'" & LOGIN_NAME & "')"
                CnG.Execute SQL
            Else
                SQL = "INSERT INTO t_spl_auto(company_code,employee_code,date," & _
                          "ot_spl,ot_15,ot_20,ot_30,ot_40,ot_60,total_ot,ot_code," & _
                          "meal_allowance,transport_allowance,flag_on_call," & _
                          "flag_shift2,flag_shift3,entry_date,entry_user) " & _
                        "VALUES( " & _
                          "'" & TDBCombo_company.Text & "','" & str_employee_code & "'," & _
                          "'" & Format(v_tgl_awal, "yyyy-MM-dd") & "'," & _
                          "0,0,0,0,0,0," & _
                          "'" & vOTSPL & "','',0,0,0,0,0,now(),'" & LOGIN_NAME & "')"
                CnG.Execute SQL
            End If
'            End If
            rs2.Close
            
            v_tgl_awal = v_tgl_awal + 1
            rs.MoveNext
        Wend
        
    End If
    rs.Close
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
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


Private Sub hitungLembur(pOT1 As Double, pOT2 As Double, pOT3 As Double, pOT4 As Double, pOT6 As Double, _
                        pOT_Code As String, pEmployee_code As String)
Dim rscari As New ADODB.Recordset
Dim vTime1, vTime2 As String
Dim vFlagTransportEmp As Integer
Dim jmlJam As Double
        
    jmlJam = pOT1 + pOT2 + pOT3 + pOT4 + pOT6
    
    SQL = "SELECT * FROM m_ot_detail WHERE ot_code = '" & pOT_Code & "' AND from_value < " & jmlJam & ""
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    If rscari.RecordCount > 0 Then
        While Not rscari.EOF
            If jmlJam > rscari!to_value Then
                If rscari!pengali = 1.5 Then
                    v_15 = rscari!to_value
                    v_20 = "0"
                    v_30 = "0"
                    v_40 = "0"
                    v_60 = "0"
                ElseIf rscari!pengali = 2 Then
                    If pOT_Code <> "HK" Then v_15 = "0" Else v_15 = v_15
                    If rscari!ot_number = 1 Then
                        v_20 = rscari!to_value
                    Else
                        v_20 = jmlJam - rscari!from_value
                    End If
                    v_30 = "0"
                    v_40 = "0"
                    v_60 = "0"
                ElseIf rscari!pengali = 3 Then
                    If pOT_Code <> "HK" Then v_15 = "0" Else v_15 = v_15
                    v_30 = "1"
                    v_40 = "0"
                    v_60 = "0"
                ElseIf rscari!pengali = 4 Then
                    If pOT_Code <> "HK" Then v_15 = "0" Else v_15 = v_15
                    If rscari!ot_number = 1 Then
                        v_40 = rscari!to_value
                    Else
                        v_40 = jmlJam - rscari!from_value
                    End If
                    v_60 = "0"
                ElseIf rscari!pengali = 6 Then
                    If rscari!ot_number = 1 Then
                        v_60 = rscari!to_value
                    Else
                        v_60 = jmlJam - rscari!from_value
                    End If
                End If
            Else
                If rscari!pengali = 1.5 Then
                    v_15 = jmlJam
                    v_20 = "0"
                    v_30 = "0"
                    v_40 = "0"
                    v_60 = "0"
                ElseIf rscari!pengali = 2 Then
                    If pOT_Code <> "HK" Then v_15 = "0" Else v_15 = v_15
                    v_20 = jmlJam - rscari!from_value
                    If Val(v_20) < 0 Then v_20 = 0
                    v_30 = "0"
                    v_40 = "0"
                    v_60 = "0"
                ElseIf rscari!pengali = 3 Then
                    If pOT_Code <> "HK" Then v_15 = "0" Else v_15 = v_15
                    v_30 = jmlJam - v_20 - v_15
                    If Val(v_30) < 0 Then v_30 = 0
                    v_40 = "0"
                    v_60 = "0"
                ElseIf rscari!pengali = 4 Then
                    If pOT_Code <> "HK" Then v_15 = "0" Else v_15 = v_15
                    v_40 = jmlJam - v_30 - v_20 - v_15
                    v_40 = Round(v_40, 2)
                    If Val(v_40) < 0 Then v_40 = 0
                    v_60 = "0"
                ElseIf rscari!pengali = 6 Then
                    If pOT_Code <> "HK" Then v_15 = "0" Else v_15 = v_15
                    v_60 = jmlJam - v_40 - v_30 - v_20 - v_15
                    If Val(v_60) < 0 Then v_60 = 0
                End If
            End If
        rscari.MoveNext
        Wend
    End If
    rscari.Close
    
    vTotOT = (1.5 * v_15) + (2 * v_20) + (3 * v_30) + (4 * v_40) + (6 * v_60)
        
    SQL = "SELECT ot_max_meal, ot_max_trans, flag_meal, flag_trans FROM m_pref_ot"
    rscari.Open SQL, CnG, adOpenForwardOnly
    
    If rscari.RecordCount > 0 Then
        vOTMax_Meal = rscari!ot_max_meal
        vFlagMeal = rscari!flag_meal
        vOTMax_Trans = rscari!ot_max_trans
        vFlagTrans = rscari!flag_trans
    Else
        vOTMax_Meal = 0
        vFlagMeal = 0
        vOTMax_Trans = 0
        vFlagTrans = 0
    End If
    rscari.Close
    
    SQL = "SELECT flag_transport FROM m_employee WHERE employee_code = '" & pEmployee_code & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vFlagTransportEmp = rscari!flag_transport
    End If
    rscari.Close
    
    If vFlagMeal <> 0 Then
        If jmlJam >= vOTMax_Meal Then
            vTotMeal = 1
        Else
            vTotMeal = 0
        End If
    Else
        vTotMeal = 0
    End If
    
    If vFlagTransportEmp <> 0 Then
        If vFlagTrans <> 0 Then
            If pOT_Code <> "HK" Then
                If jmlJam > 0 Then
                    vTotTransport = 1
                Else
                    vTotTransport = 0
                End If
            Else
                vTotTransport = 0
            End If
        Else
            vTotTransport = 0
        End If
    Else
        vTotTransport = 0
    End If
End Sub
