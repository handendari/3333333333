VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_proses_bonus 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BONUS PROCESS"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14910
   Icon            =   "frm_Proses_Bonus.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   22
      Top             =   8100
      Width           =   14325
      Begin prj_tpc.vbButton cmdExit 
         Height          =   705
         Left            =   12780
         TabIndex        =   11
         Top             =   390
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
         MICON           =   "frm_Proses_Bonus.frx":058A
         PICN            =   "frm_Proses_Bonus.frx":05A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_tpc.vbButton cmdNew 
         Height          =   705
         Left            =   7290
         TabIndex        =   24
         Top             =   300
         Visible         =   0   'False
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
         MICON           =   "frm_Proses_Bonus.frx":1638
         PICN            =   "frm_Proses_Bonus.frx":1654
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
         Left            =   840
         TabIndex        =   8
         Top             =   270
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
         MICON           =   "frm_Proses_Bonus.frx":26E6
         PICN            =   "frm_Proses_Bonus.frx":2702
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
         Left            =   1860
         TabIndex        =   9
         Top             =   270
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
         MICON           =   "frm_Proses_Bonus.frx":3794
         PICN            =   "frm_Proses_Bonus.frx":37B0
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
         Left            =   8280
         TabIndex        =   25
         Top             =   300
         Visible         =   0   'False
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
         MICON           =   "frm_Proses_Bonus.frx":4842
         PICN            =   "frm_Proses_Bonus.frx":485E
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
         Left            =   2880
         TabIndex        =   10
         Top             =   270
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
         MICON           =   "frm_Proses_Bonus.frx":58F0
         PICN            =   "frm_Proses_Bonus.frx":590C
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
   Begin VB.Frame fra_entry 
      Height          =   4455
      Left            =   240
      TabIndex        =   21
      Top             =   3570
      Width           =   14325
      Begin VB.TextBox txt_pengali_entry 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   6480
         TabIndex        =   5
         Top             =   900
         Width           =   435
      End
      Begin VB.CheckBox chkPresence_Entry 
         Caption         =   "Presence Allowance"
         Height          =   195
         Left            =   7410
         TabIndex        =   6
         Top             =   1290
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkPosition_Entry 
         Caption         =   "Position Allowance"
         Height          =   195
         Left            =   9360
         TabIndex        =   7
         Top             =   1290
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Frame Frame1 
         Caption         =   "List Karyawan"
         Height          =   4065
         Left            =   3360
         TabIndex        =   26
         Top             =   180
         Width           =   2895
         Begin VB.ListBox list_employee 
            Height          =   3765
            ItemData        =   "frm_Proses_Bonus.frx":699E
            Left            =   90
            List            =   "frm_Proses_Bonus.frx":69A0
            TabIndex        =   27
            Top             =   210
            Width           =   2715
         End
      End
      Begin VB.Label Label7 
         Caption         =   "X Basic Salary"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7020
         TabIndex        =   29
         Top             =   930
         Width           =   1755
      End
      Begin VB.Label Label6 
         Caption         =   "INCLUDE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6480
         TabIndex        =   28
         Top             =   1290
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CheckBox chkPosition 
      Caption         =   "Position Allowance"
      Height          =   195
      Left            =   4080
      TabIndex        =   2
      Top             =   2190
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chkPresence 
      Caption         =   "Presence Allowance"
      Height          =   195
      Left            =   2130
      TabIndex        =   1
      Top             =   2190
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txt_pengali 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   1800
      Width           =   435
   End
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1170
      OleObjectBlob   =   "frm_Proses_Bonus.frx":69A2
      TabIndex        =   13
      Top             =   1230
      Width           =   1785
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   14
      Top             =   1230
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1170
      TabIndex        =   12
      Top             =   810
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM"
      Format          =   93519875
      UpDown          =   -1  'True
      CurrentDate     =   39278
   End
   Begin prj_tpc.vbButton vbButton3 
      Height          =   495
      Left            =   3000
      TabIndex        =   17
      Top             =   690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Refresh"
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
      MICON           =   "frm_Proses_Bonus.frx":8908
      PICN            =   "frm_Proses_Bonus.frx":8924
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton vbButton2 
      Height          =   465
      Left            =   11250
      TabIndex        =   4
      Top             =   2010
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
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
      MICON           =   "frm_Proses_Bonus.frx":99B6
      PICN            =   "frm_Proses_Bonus.frx":99D2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_tpc.vbButton vbButton1 
      Height          =   465
      Left            =   9810
      TabIndex        =   3
      Top             =   2010
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Proccess"
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
      MICON           =   "frm_Proses_Bonus.frx":AA64
      PICN            =   "frm_Proses_Bonus.frx":AA80
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   5445
      Left            =   240
      TabIndex        =   23
      Top             =   2580
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   9604
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "EMPLOYEE CODE"
      Columns(0).DataField=   "employee_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EMP. ID NO."
      Columns(1).DataField=   "nik"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "EMPLOYEE NAME"
      Columns(2).DataField=   "employee_name"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "TITLE CODE"
      Columns(3).DataField=   "title_code"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TITLE NAME"
      Columns(4).DataField=   "title_name"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "BASIC SALARY"
      Columns(5).DataField=   "basic_salary"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "MULTIPLIER"
      Columns(6).DataField=   "pengali"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "BONUS"
      Columns(7).DataField=   "jml_bonus"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "TAX"
      Columns(8).DataField=   "pph21_value"
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "ACTUAL"
      Columns(9).DataField=   "actual"
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2990"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2910"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2223"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2143"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=5662"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=5583"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=5292"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=5212"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(28)=   "Column(5).Width=3069"
      Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=2990"
      Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(33)=   "Column(6).Width=1720"
      Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=1640"
      Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(38)=   "Column(7).Width=3307"
      Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=3228"
      Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(43)=   "Column(8).Width=2937"
      Splits(0)._ColumnProps(44)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(8)._WidthInPix=2858"
      Splits(0)._ColumnProps(46)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(47)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(48)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(49)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(51)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(52)=   "Column(9).Order=10"
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
      Caption         =   "LIST OF BONUS"
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
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=0"
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
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=1"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=70,.parent=13,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(74)  =   "Named:id=33:Normal"
      _StyleDefs(75)  =   ":id=33,.parent=0"
      _StyleDefs(76)  =   "Named:id=34:Heading"
      _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(78)  =   ":id=34,.wraptext=-1"
      _StyleDefs(79)  =   "Named:id=35:Footing"
      _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(81)  =   "Named:id=36:Selected"
      _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(83)  =   "Named:id=37:Caption"
      _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(85)  =   "Named:id=38:HighlightRow"
      _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(87)  =   "Named:id=39:EvenRow"
      _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(89)  =   "Named:id=40:OddRow"
      _StyleDefs(90)  =   ":id=40,.parent=33"
      _StyleDefs(91)  =   "Named:id=41:RecordSelector"
      _StyleDefs(92)  =   ":id=41,.parent=34"
      _StyleDefs(93)  =   "Named:id=42:FilterBar"
      _StyleDefs(94)  =   ":id=42,.parent=33"
   End
   Begin prj_tpc.vbButton vbButton4 
      Height          =   465
      Left            =   13170
      TabIndex        =   30
      Top             =   2010
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Tax"
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
      MICON           =   "frm_Proses_Bonus.frx":BB12
      PICN            =   "frm_Proses_Bonus.frx":BB2E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label5 
      Caption         =   "INCLUDE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1200
      TabIndex        =   20
      Top             =   2190
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "X Basic Salary"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1740
      TabIndex        =   19
      Top             =   1830
      Width           =   1755
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BONUS PROCESS"
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
      Left            =   390
      TabIndex        =   18
      Top             =   150
      Width           =   4845
   End
   Begin VB.Label Label1 
      Caption         =   "DATE :"
      Height          =   195
      Left            =   270
      TabIndex        =   16
      Top             =   870
      Width           =   705
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COMPANY :"
      Height          =   195
      Left            =   270
      TabIndex        =   15
      Top             =   1290
      Width           =   885
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_Proses_Bonus.frx":CBC0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16860
   End
End
Attribute VB_Name = "frm_proses_bonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompany As New ADODB.Recordset
    
Dim rs As New ADODB.Recordset
Dim rsBonus As New ADODB.Recordset

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Dim int_mode As Integer
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Dim rs2 As New ADODB.Recordset
Dim vBasicSalary As Double
Dim vPresenceAllowance As Double
Dim vPositionAllowance As Double
Dim vBonus As Double
Dim vTgl As String

Private Sub cmdSave_Click()
    Me.MousePointer = vbHourglass
    
    strsql = "SELECT employee_code,employee_name," _
            & "IFNULL((SELECT basic_salary FROM m_salary_standard " _
                        & "Where employee_code = a.employee_code And salary_date <= '" & Format(DTPicker1, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1),0) as basic_salary," _
            & "IFNULL((SELECT presence_allowance FROM m_salary_standard " _
                        & "Where employee_code = a.employee_code And salary_date <= '" & Format(DTPicker1, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1),0) as presence_allowance," _
            & "IFNULL((SELECT position_allowance FROM m_salary_standard " _
                        & "Where employee_code = a.employee_code And salary_date <= '" & Format(DTPicker1, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1),0) as position_allowance " _
        & "FROM temp_list a"
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        CnG.BeginTrans
        rs.MoveFirst
        While Not rs.EOF
            vBasicSalary = rs!basic_salary
            vPresenceAllowance = (rs!presence_allowance / 100) * rs!basic_salary
            vPositionAllowance = rs!position_allowance
            
            If chkPosition_Entry.Value Then
                vBonus = txt_pengali_entry * (vBasicSalary + vPositionAllowance)
            ElseIf chkPresence_Entry.Value Then
                vBonus = txt_pengali_entry * (vBasicSalary + vPresenceAllowance)
            ElseIf chkPosition_Entry.Value And chkPresence_Entry.Value Then
                vBonus = txt_pengali_entry * (vBasicSalary + vPresenceAllowance + vPositionAllowance)
            Else
                vBonus = (txt_pengali_entry * vBasicSalary)
            End If
            
            vTgl = Format(DTPicker1, "yyyy-MM-20")
            
            SQL = "UPDATE t_bonus SET pengali = " & Val(txt_pengali_entry.Text) & "," & _
                    "jml_bonus = '" & vBonus & "'," & _
                    "tgledit = Now()," & _
                    "useredit = '" & LOGIN_NAME & "' " & _
                  "where employee_code = '" & rs!employee_code & "' " & _
                    "AND DATE(tgltrans) = '" & Format(vTgl, "yyyy-MM-20") & "'"
            CnG.Execute SQL
            rs.MoveNext
        Wend
        CnG.CommitTrans
    End If
    rs.Close
    
    Me.MousePointer = vbNormal
    
    
    Call isiGridBonus
    int_mode = 0
    Call load_mode
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    chkPosition.Value = 0
    chkPresence.Value = 0
    txt_pengali.Text = 0
    
    timer1.Enabled = True
    Call isiCompany
    
    oClause = ""
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
End Sub

Private Sub isiCompany()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    Set TDBCombo_company.RowSource = rsCompany
End Sub

Public Sub isiBonus()

    Me.MousePointer = vbHourglass
    
    SQL = "DELETE FROM t_bonus WHERE DATE_FORMAT(tgltrans,'%Y-%m') = '" & Format(DTPicker1, "yyyy-MM") & "'"
    CnG.Execute SQL
    
    strsql = "SELECT employee_code,employee_name,a.title_code,c.title_name,a.department_code,d.department_name," _
            & "DATE(a.start_working) start_working,a.religion," _
            & "IFNULL((SELECT basic_salary FROM m_salary_standard " _
                        & "Where employee_code = a.employee_code And salary_date <= '" & Format(DTPicker1, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1),0) as basic_salary," _
            & "IFNULL((SELECT presence_allowance FROM m_salary_standard " _
                        & "Where employee_code = a.employee_code And salary_date <= '" & Format(DTPicker1, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1),0) as presence_allowance," _
            & "IFNULL((SELECT position_allowance FROM m_salary_standard " _
                        & "Where employee_code = a.employee_code And salary_date <= '" & Format(DTPicker1, "yyyy-MM-dd") & "' ORDER BY salary_date DESC LIMIT 1),0) as position_allowance," _
            & "DATE(start_working) start_working " _
        & "FROM m_employee a JOIN m_department d ON a.department_code = d.department_code " _
        & "JOIN m_division b ON a.division_code = b.division_code AND a.company_code = b.company_code " _
        & "JOIN m_title c ON a.title_code = c.title_code " _
        & "WHERE a.company_code = '" & TDBCombo_company.Text & "' "
    rs.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        CnG.BeginTrans
        rs.MoveFirst
        While Not rs.EOF
            vBasicSalary = rs!basic_salary
            vPresenceAllowance = (rs!presence_allowance / 100) * rs!basic_salary
            vPositionAllowance = rs!position_allowance
            
            If chkPosition.Value Then
                vBonus = txt_pengali * (vBasicSalary + vPositionAllowance)
            ElseIf chkPresence.Value Then
                vBonus = txt_pengali * (vBasicSalary + vPresenceAllowance)
            ElseIf chkPosition.Value And chkPresence.Value Then
                vBonus = txt_pengali * (vBasicSalary + vPresenceAllowance + vPositionAllowance)
            Else
                vBonus = (txt_pengali * vBasicSalary)
            End If
            
            vTgl = Format(DTPicker1, "yyyy-MM-20")
            
            SQL = "INSERT INTO t_bonus (tgltrans,employee_code,basic_salary,pengali,jml_bonus,tglinput,userinput) " _
                & "VALUES " _
                & "('" & vTgl & "','" & rs!employee_code & "','" & rs!basic_salary & "'," _
                & "" & Val(txt_pengali.Text) & ",'" & vBonus & "',now(),'" & LOGIN_CODE & "')"
            CnG.Execute SQL
            rs.MoveNext
        Wend
        CnG.CommitTrans
    End If
    rs.Close
    
    isiGridBonus
    Me.MousePointer = vbNormal
End Sub

Private Sub isiGridBonus()
    If rsBonus.State Then rsBonus.Close
    SQL = "SELECT a.*,b.nik,b.employee_name,b.title_code,c.title_name, (a.jml_bonus - a.pph21_value) actual " _
        & "FROM t_bonus a JOIN m_employee b ON a.employee_code = b.employee_code " _
        & "JOIN m_title c ON b.title_code = c.title_code " _
        & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
            & "AND DATE_FORMAT(tgltrans,'%Y-%m') = '" & Format(DTPicker1, "yyyy-MM") & "' " & oClause
    rsBonus.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsBonus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_proses_bonus = Nothing
End Sub

Private Sub LynxGrid2_RowColChanged()
    isiGridBonus
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
        
        isiGridBonus
    End If
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub vbButton1_Click()
    isiBonus
End Sub

Private Sub vbbutton2_Click()
Dim i As Integer
Dim tanya As Integer

    Me.MousePointer = vbHourglass
    
    Set SelBks = TDBGrid1.SelBookmarks
    i = MsgBox("Are you sure want to delete " _
        & SelBks.Count & " bonus data ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        SQL = "DELETE FROM t_bonus " & _
              "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").CellText(item) & "' " & _
                "AND DATE_FORMAT(tgltrans,'%Y-%m') = '" & Format(DTPicker1, "yyyy-MM") & "'"
        CnG.Execute SQL
    Next
    CnG.CommitTrans
    
    Call isiGridBonus
    
    Me.MousePointer = vbNormal
End Sub

Private Sub vbButton3_Click()
    isiGridBonus
End Sub


Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsBonus.Filter = adFilterNone
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
    
    rsBonus.Filter = getFilter()
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

Private Sub cmdCancel_Click()
    int_mode = 0
    Call load_mode
End Sub

Public Sub set_edit_data()
Dim i As Integer
Dim item
On Error GoTo Err
    If Not TDBGrid1.ApproxCount > 0 Then
        Exit Sub
    End If
    
    Set SelBks = TDBGrid1.SelBookmarks
    i = 0
    CnG.BeginTrans
    
    SQL = "DELETE FROM temp_list"
    CnG.Execute SQL
    
    For Each item In SelBks
        i = i + 1
        
        SQL = "INSERT INTO temp_list(employee_code,nik,employee_name) VALUES " _
                & "('" & TDBGrid1.Columns("employee_code").CellText(item) & "'," _
                & "'" & TDBGrid1.Columns("nik").CellText(item) & "'," _
                & "'" & TDBGrid1.Columns("employee_name").CellText(item) & "')"
        CnG.Execute SQL
                
    Next
    CnG.CommitTrans
    
    list_employee.Clear
    strsql = "select employee_name from temp_list"
    rscari.Open strsql, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        rscari.MoveFirst
        While Not rscari.EOF
            list_employee.AddItem rscari!EMPLOYEE_NAME
            rscari.MoveNext
        Wend
    End If
    rscari.Close
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdEdit_Click()
    txt_pengali_entry.Text = 0
    chkPosition_Entry.Value = 0
    chkPresence_Entry.Value = 0
    
    int_mode = 2
    Call load_mode
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
    cmdNew.Enabled = a And blnUser_Add
    cmdSave.Enabled = b
    cmdEdit.Enabled = c And blnUser_Edit
    cmdDelete.Enabled = d And blnUser_Delete
    cmdCancel.Enabled = e
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        fra_entry.Visible = True
        TDBGrid1.Enabled = False
        
    ElseIf int_mode = 0 Then    'VIEW
        fra_entry.Visible = False
        TDBGrid1.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
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

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
    
    x = x + 1
    
    If x Mod 2 <> 1 And vSubject = TDBGrid1.Columns(ColIndex).DataField Then
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " DESC"
    Else
        oClause = " ORDER BY " + TDBGrid1.Columns(ColIndex).DataField + " ASC"
    End If
    
    vSubject = TDBGrid1.Columns(ColIndex).DataField
    Call isiGridBonus

End Sub


Private Sub HitungPPh(str_employee_code As String)
Dim vJmlBln As Integer
Dim vJmlBruto As Double
Dim vJmlJamsostek As Double
Dim vAvgBruto As Double
Dim vAvgJamsostek As Double

Dim vPTKP As String
Dim vPTKP_Value As Double
Dim vMarital As Integer
Dim vSex As Integer
Dim vNoChild As Integer
Dim vStartWorking As String
Dim vJmlBlnKerja As Integer

Dim vTHR_Value As Double
Dim vBonus_Value As Double
Dim vBruto_Setahun As Double
Dim vPengurang_Setahun As Double
Dim vNetto_Setahun As Double
Dim vBiayaJabatan As Double
Dim vPKP As Double

Dim vPPh5 As Double
Dim vPPh15 As Double
Dim vPPh25 As Double
Dim vPPh30 As Double
Dim vPPh21_Bonus_Setahun As Double
Dim vPPh21_Gaji_THR_Setahun As Double
Dim vPPh21 As Double

Dim i As Integer

'On Error GoTo Err
    
    CnG.BeginTrans
    If rs.State Then rs.Close
    SQL = "SELECT SUM(CASE WHEN salary_code='SU-28' THEN 1 ELSE 0 END) jmlBln," & _
            "SUM(CASE WHEN salary_code='SU-281' THEN salary_value ELSE 0 END) jamsostek," & _
            "SUM(CASE WHEN salary_code='SU-28' THEN salary_value ELSE 0 END) bruto " & _
          "FROM h_salary " & _
          "WHERE employee_code = '" & str_employee_code & "' " & _
            "AND LEFT(MONTH,4) = '" & Format(DTPicker1.Value, "yyyy") & "' " & _
            "AND (RIGHT(MONTH,2) BETWEEN '01' " & _
                "AND '" & Format(DTPicker1.Value, "MM") & "') " & _
          "GROUP BY employee_code"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vJmlBln = rs.Fields(0).Value
        vJmlJamsostek = rs.Fields(1).Value
        vJmlBruto = rs.Fields(2).Value
    Else
        vJmlBln = 0
        vJmlJamsostek = 0
        vJmlBruto = 0
    End If
    rs.Close
    
    vAvgJamsostek = vJmlJamsostek / vJmlBln
    vAvgBruto = vJmlBruto / vJmlBln
    
    SQL = "SELECT a.ptkp_type, b.marital_status, b.sex, b.no_of_children, DATE(b.start_working) " & _
          "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
          "WHERE a.employee_code = '" & str_employee_code & "' " & _
            "AND a.salary_date <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
          "ORDER BY a.salary_date DESC LIMIT 1"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vPTKP = rs.Fields(0).Value
        vMarital = rs.Fields(1).Value
        vSex = rs.Fields(2).Value
        vNoChild = rs.Fields(3).Value
        vStartWorking = rs.Fields(4).Value
    Else
        vPTKP = "STD"
        vMarital = 0
        vSex = 0
        vNoChild = 0
        vStartWorking = Now
    End If
    rs.Close
    
    SQL = "SELECT f_get_ptkp(" & vMarital & ", " & vNoChild & "," & vSex & ", 1,'" & vPTKP & "') ptkp_value"
    rs.Open SQL, CnG, adOpenForwardOnly
    
    If rs.RecordCount > 0 Then
        vPTKP_Value = rs!ptkp_value
    End If
    rs.Close
    
    SQL = "SELECT SUM(jmlthr) FROM t_thr " & _
          "WHERE employee_code = '" & str_employee_code & "' " & _
            "AND YEAR(tgltrans) = '" & Format(DTPicker1.Value, "yyyy") & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vTHR_Value = rs.Fields(0).Value
    Else
        vTHR_Value = 0
    End If
    rs.Close
    
    SQL = "SELECT SUM(jml_bonus) FROM t_bonus " & _
          "WHERE employee_code = '" & str_employee_code & "' " & _
            "AND YEAR(tgltrans) = '" & Format(DTPicker1.Value, "yyyy") & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vBonus_Value = rs.Fields(0).Value
    Else
        vBonus_Value = 0
    End If
    rs.Close
    
    For i = 0 To 1
        vJmlBlnKerja = (DateDiff("m", vStartWorking, year(DTPicker1.Value) & "-12-31")) + 1
        If vJmlBlnKerja < 12 And Format(vStartWorking, "yyyy") = Format(DTPicker1.Value, "yyyy") Then
            vBruto_Setahun = IIf(i = 0, ((vAvgBruto * vJmlBlnKerja) + vTHR_Value + vBonus_Value), ((vAvgBruto * vJmlBlnKerja) + vTHR_Value))
            vBiayaJabatan = IIf(0.05 * vBruto_Setahun > 6000000, 6000000, 0.05 * vBruto_Setahun)
            vPengurang_Setahun = vBiayaJabatan + (vJmlBlnKerja * vAvgJamsostek)
        Else
            vBruto_Setahun = IIf(i = 0, ((vAvgBruto * 12) + vTHR_Value + vBonus_Value), ((vAvgBruto * 12) + vTHR_Value))
            vBiayaJabatan = IIf(0.05 * vBruto_Setahun > 6000000, 6000000, 0.05 * vBruto_Setahun)
            vPengurang_Setahun = vBiayaJabatan + (12 * vAvgJamsostek)
        End If
        
        vNetto_Setahun = vBruto_Setahun - vPengurang_Setahun
        vPKP = vNetto_Setahun - vPTKP_Value
        vPKP = Int(vPKP / 1000) * 1000
        
        If vPKP < 50000000 Then
            vPPh5 = 0.05 * vPKP
            vPPh15 = 0
            vPPh25 = 0
            vPPh30 = 0
        ElseIf vPKP > 50000000 And vPKP < 250000000 Then
            vPPh5 = 0.05 * 50000000
            vPPh15 = 0.15 * (vPKP - 50000000)
            vPPh25 = 0
            vPPh30 = 0
        ElseIf vPKP > 250000000 And vPKP < 500000000 Then
            vPPh5 = 0.05 * 50000000
            vPPh15 = 0.15 * 200000000
            vPPh25 = 0.25 * (vPKP - 250000000)
            vPPh30 = 0
        Else
            vPPh5 = 0.05 * 50000000
            vPPh15 = 0.15 * 200000000
            vPPh25 = 0.25 * 250000000
            vPPh30 = 0.35 * (vPKP - 500000000)
        End If
        
        If i = 0 Then
            vPPh21_Bonus_Setahun = vPPh5 + vPPh15 + vPPh25 + vPPh30
        Else
            vPPh21_Gaji_THR_Setahun = vPPh5 + vPPh15 + vPPh25 + vPPh30
        End If
                
    Next
    
    vPPh21 = vPPh21_Bonus_Setahun - vPPh21_Gaji_THR_Setahun
    
    '------------------------------- Update Bonus ------------------------------
    SQL = "UPDATE t_bonus SET pph21_value = '" & vPPh21 & "' " & _
          "WHERE employee_code = '" & str_employee_code & "' " & _
            "AND YEAR(tgltrans) = '" & Format(DTPicker1.Value, "yyyy") & "' " & _
            "AND MONTH(tgltrans) = '" & Format(DTPicker1.Value, "MM") & "'"
    CnG.Execute SQL
    '-------------------------------------------------------------------------
    
    CnG.CommitTrans
    Exit Sub
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub vbButton4_Click()
    Me.MousePointer = vbHourglass
    
    If rscari.State Then rscari.Close
    SQL = "SELECT employee_code " & _
             "FROM m_employee " & _
             "WHERE company_code = '" & TDBCombo_company.Text & "' " & _
                "AND flag_active <> 0"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        rscari.MoveFirst
        While Not rscari.EOF
            Call HitungPPh(rscari.Fields(0).Value)
            rscari.MoveNext
        Wend
    End If
    rscari.Close
    
    Call isiGridBonus
    
    Me.MousePointer = vbNormal
End Sub
