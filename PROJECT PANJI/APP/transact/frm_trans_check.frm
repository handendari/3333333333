VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_manual_check 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IJIN KHUSUS"
   ClientHeight    =   9090
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
   Icon            =   "frm_trans_check.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9090
   ScaleWidth      =   14685
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_division_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3540
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   28
      Top             =   1230
      Width           =   3855
   End
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   3465
      Left            =   2280
      TabIndex        =   24
      Top             =   5130
      Visible         =   0   'False
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   6112
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
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3540
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   12
      Top             =   810
      Width           =   3855
   End
   Begin VB.Frame fra_entry 
      Height          =   3135
      Left            =   240
      TabIndex        =   5
      Top             =   4320
      Width           =   14175
      Begin prj_panji.vbButton cmdBrowse 
         Height          =   315
         Left            =   3420
         TabIndex        =   25
         Top             =   480
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         BTYPE           =   14
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         MICON           =   "frm_trans_check.frx":058A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cbo_flag_io 
         Height          =   315
         ItemData        =   "frm_trans_check.frx":05A6
         Left            =   8400
         List            =   "frm_trans_check.frx":05B0
         TabIndex        =   2
         Text            =   "..."
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   1395
         Left            =   8400
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1200
         Width           =   4695
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txt_nik 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2040
         MaxLength       =   10
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker_check 
         Height          =   315
         Left            =   8400
         TabIndex        =   3
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   157286403
         CurrentDate     =   39270
      End
      Begin VB.TextBox txt_employee_code 
         Height          =   315
         Left            =   3690
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label lblVerify 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   750
         TabIndex        =   27
         Top             =   1680
         Width           =   6645
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "STATUS"
         Height          =   195
         Left            =   7620
         TabIndex        =   15
         Top             =   480
         Width           =   570
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TANGGAL"
         Height          =   195
         Left            =   7500
         TabIndex        =   14
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KETERANGAN"
         Height          =   195
         Left            =   7200
         TabIndex        =   11
         Top             =   1200
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KODE KARY."
         Height          =   195
         Left            =   720
         TabIndex        =   10
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NAMA KARY."
         Height          =   195
         Left            =   675
         TabIndex        =   9
         Top             =   840
         Width           =   930
      End
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   8
      Top             =   7560
      Width           =   14175
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin prj_panji.vbButton cmdNew 
         Height          =   705
         Left            =   660
         TabIndex        =   16
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Tambah"
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
         MICON           =   "frm_trans_check.frx":05C9
         PICN            =   "frm_trans_check.frx":05E5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdSave 
         Height          =   705
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Simpan"
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
         MICON           =   "frm_trans_check.frx":1677
         PICN            =   "frm_trans_check.frx":1693
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdEdit 
         Height          =   705
         Left            =   2700
         TabIndex        =   18
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Ubah"
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
         MICON           =   "frm_trans_check.frx":2725
         PICN            =   "frm_trans_check.frx":2741
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdDelete 
         Height          =   705
         Left            =   3720
         TabIndex        =   19
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Hapus"
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
         MICON           =   "frm_trans_check.frx":37D3
         PICN            =   "frm_trans_check.frx":37EF
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdCancel 
         Height          =   705
         Left            =   4740
         TabIndex        =   20
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Batal"
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
         MICON           =   "frm_trans_check.frx":4881
         PICN            =   "frm_trans_check.frx":489D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdExit 
         Height          =   705
         Left            =   9780
         TabIndex        =   21
         Top             =   390
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Keluar"
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
         MICON           =   "frm_trans_check.frx":592F
         PICN            =   "frm_trans_check.frx":594B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton btnApprove 
         Height          =   705
         Left            =   7170
         TabIndex        =   26
         Top             =   330
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1244
         BTYPE           =   14
         TX              =   "&Setujui"
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
         MICON           =   "frm_trans_check.frx":69DD
         PICN            =   "frm_trans_check.frx":69F9
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   10186
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
      Columns(3).Caption=   "DEPARTEMEN"
      Columns(3).DataField=   "department_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "DIV. KODE"
      Columns(4).DataField=   "division_code"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "DIVISI"
      Columns(5).DataField=   "division_name"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "TITLE CODE"
      Columns(6).DataField=   "title_code"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "JABATAN"
      Columns(7).DataField=   "title_name"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "EMPLOYEE CODE"
      Columns(8).DataField=   "employee_code"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "KODE KARY."
      Columns(9).DataField=   "nik"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "NAMA KARY."
      Columns(10).DataField=   "employee_name"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   17
      Columns(11)._MaxComboItems=   5
      Columns(11).ValueItems(0)._DefaultItem=   0
      Columns(11).ValueItems(0).Value=   "0"
      Columns(11).ValueItems(0).Value.vt=   8
      Columns(11).ValueItems(0).DisplayValue=   "Check In"
      Columns(11).ValueItems(0).DisplayValue.vt=   8
      Columns(11).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
      Columns(11).ValueItems(1)._DefaultItem=   0
      Columns(11).ValueItems(1).Value=   "1"
      Columns(11).ValueItems(1).Value.vt=   8
      Columns(11).ValueItems(1).DisplayValue=   "Check Out"
      Columns(11).ValueItems(1).DisplayValue.vt=   8
      Columns(11).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
      Columns(11).ValueItems.Count=   2
      Columns(11).Caption=   "STATUS"
      Columns(11).DataField=   "flag_io"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "TANGGAL"
      Columns(12).DataField=   "check_date"
      Columns(12).NumberFormat=   "yyyy-MM-dd"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   4
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "DISETUJUI"
      Columns(13).DataField=   "flag_approval"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "KETERANGAN"
      Columns(14).DataField=   "description"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "FLAG IO"
      Columns(15).DataField=   "flag_io"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "USER APPROVE"
      Columns(16).DataField=   "user_approval"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   17
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   1
      Splits(0).Size  =   3000.189
      Splits(0).Size.vt=   4
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   3
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=17"
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
      Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(23)=   "Column(3).Width=3519"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=3440"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(29)=   "Column(4).Width=1879"
      Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1799"
      Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(34)=   "Column(5).Width=2434"
      Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=2355"
      Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(39)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(49)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(52)=   "Column(8)._ColStyle=516"
      Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(54)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(55)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(57)=   "Column(9).AllowSizing=0"
      Splits(0)._ColumnProps(58)=   "Column(9)._ColStyle=516"
      Splits(0)._ColumnProps(59)=   "Column(9).Visible=0"
      Splits(0)._ColumnProps(60)=   "Column(9).AllowFocus=0"
      Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(62)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(65)=   "Column(10).AllowSizing=0"
      Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=516"
      Splits(0)._ColumnProps(67)=   "Column(10).Visible=0"
      Splits(0)._ColumnProps(68)=   "Column(10).AllowFocus=0"
      Splits(0)._ColumnProps(69)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(70)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(71)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(72)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(73)=   "Column(11)._ColStyle=516"
      Splits(0)._ColumnProps(74)=   "Column(11).Button=1"
      Splits(0)._ColumnProps(75)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(76)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(77)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(78)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(79)=   "Column(12).AllowSizing=0"
      Splits(0)._ColumnProps(80)=   "Column(12)._ColStyle=516"
      Splits(0)._ColumnProps(81)=   "Column(12).Visible=0"
      Splits(0)._ColumnProps(82)=   "Column(12).AllowFocus=0"
      Splits(0)._ColumnProps(83)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(84)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(85)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(86)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(87)=   "Column(13)._ColStyle=516"
      Splits(0)._ColumnProps(88)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(89)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(90)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(92)=   "Column(14).AllowSizing=0"
      Splits(0)._ColumnProps(93)=   "Column(14)._ColStyle=516"
      Splits(0)._ColumnProps(94)=   "Column(14).Visible=0"
      Splits(0)._ColumnProps(95)=   "Column(14).AllowFocus=0"
      Splits(0)._ColumnProps(96)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(97)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(98)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(100)=   "Column(15)._ColStyle=516"
      Splits(0)._ColumnProps(101)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(102)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(103)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(104)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(105)=   "Column(16)._ColStyle=516"
      Splits(0)._ColumnProps(106)=   "Column(16).Order=17"
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
      Splits(1)._ColumnProps(0)=   "Columns.Count=17"
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
      Splits(1)._ColumnProps(65)=   "Column(8).Width=2725"
      Splits(1)._ColumnProps(66)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(67)=   "Column(8)._WidthInPix=2646"
      Splits(1)._ColumnProps(68)=   "Column(8)._ColStyle=516"
      Splits(1)._ColumnProps(69)=   "Column(8).Visible=0"
      Splits(1)._ColumnProps(70)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(71)=   "Column(9).Width=1852"
      Splits(1)._ColumnProps(72)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(73)=   "Column(9)._WidthInPix=1773"
      Splits(1)._ColumnProps(74)=   "Column(9)._ColStyle=516"
      Splits(1)._ColumnProps(75)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(76)=   "Column(10).Width=4524"
      Splits(1)._ColumnProps(77)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(78)=   "Column(10)._WidthInPix=4445"
      Splits(1)._ColumnProps(79)=   "Column(10)._ColStyle=516"
      Splits(1)._ColumnProps(80)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(81)=   "Column(11).Width=3069"
      Splits(1)._ColumnProps(82)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(11)._WidthInPix=2990"
      Splits(1)._ColumnProps(84)=   "Column(11)._ColStyle=513"
      Splits(1)._ColumnProps(85)=   "Column(11).Button=1"
      Splits(1)._ColumnProps(86)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(87)=   "Column(12).Width=3016"
      Splits(1)._ColumnProps(88)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(89)=   "Column(12)._WidthInPix=2937"
      Splits(1)._ColumnProps(90)=   "Column(12)._ColStyle=513"
      Splits(1)._ColumnProps(91)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(92)=   "Column(13).Width=1852"
      Splits(1)._ColumnProps(93)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(94)=   "Column(13)._WidthInPix=1773"
      Splits(1)._ColumnProps(95)=   "Column(13)._ColStyle=513"
      Splits(1)._ColumnProps(96)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(97)=   "Column(14).Width=5212"
      Splits(1)._ColumnProps(98)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(99)=   "Column(14)._WidthInPix=5133"
      Splits(1)._ColumnProps(100)=   "Column(14)._ColStyle=516"
      Splits(1)._ColumnProps(101)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(102)=   "Column(15).Width=2725"
      Splits(1)._ColumnProps(103)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(104)=   "Column(15)._WidthInPix=2646"
      Splits(1)._ColumnProps(105)=   "Column(15)._ColStyle=516"
      Splits(1)._ColumnProps(106)=   "Column(15).Visible=0"
      Splits(1)._ColumnProps(107)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(108)=   "Column(16).Width=2725"
      Splits(1)._ColumnProps(109)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(110)=   "Column(16)._WidthInPix=2646"
      Splits(1)._ColumnProps(111)=   "Column(16)._ColStyle=516"
      Splits(1)._ColumnProps(112)=   "Column(16).Visible=0"
      Splits(1)._ColumnProps(113)=   "Column(16).Order=17"
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
      Caption         =   "DAFTAR IJIN KHUSUS"
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
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=90,.parent=13"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=87,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=88,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=89,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=82,.parent=13"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=79,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=80,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=81,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=46,.parent=13"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=43,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=44,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=45,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=110,.parent=13"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=107,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=108,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=109,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=98,.parent=13"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=95,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=96,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=97,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=114,.parent=13"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=111,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=112,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=113,.parent=17"
      _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=162,.parent=13"
      _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=159,.parent=14"
      _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=160,.parent=15"
      _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=161,.parent=17"
      _StyleDefs(102) =   "Splits(1).Style:id=99,.parent=1"
      _StyleDefs(103) =   "Splits(1).CaptionStyle:id=116,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(104) =   ":id=116,.fgcolor=&H80000009&"
      _StyleDefs(105) =   "Splits(1).HeadingStyle:id=100,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(106) =   ":id=100,.fgcolor=&H80000002&"
      _StyleDefs(107) =   "Splits(1).FooterStyle:id=101,.parent=3"
      _StyleDefs(108) =   "Splits(1).InactiveStyle:id=102,.parent=5"
      _StyleDefs(109) =   "Splits(1).SelectedStyle:id=104,.parent=6"
      _StyleDefs(110) =   "Splits(1).EditorStyle:id=103,.parent=7"
      _StyleDefs(111) =   "Splits(1).HighlightRowStyle:id=105,.parent=8"
      _StyleDefs(112) =   "Splits(1).EvenRowStyle:id=106,.parent=9"
      _StyleDefs(113) =   "Splits(1).OddRowStyle:id=115,.parent=10"
      _StyleDefs(114) =   "Splits(1).RecordSelectorStyle:id=117,.parent=11"
      _StyleDefs(115) =   "Splits(1).FilterBarStyle:id=118,.parent=12"
      _StyleDefs(116) =   "Splits(1).Columns(0).Style:id=122,.parent=99"
      _StyleDefs(117) =   "Splits(1).Columns(0).HeadingStyle:id=119,.parent=100"
      _StyleDefs(118) =   "Splits(1).Columns(0).FooterStyle:id=120,.parent=101"
      _StyleDefs(119) =   "Splits(1).Columns(0).EditorStyle:id=121,.parent=103"
      _StyleDefs(120) =   "Splits(1).Columns(1).Style:id=126,.parent=99"
      _StyleDefs(121) =   "Splits(1).Columns(1).HeadingStyle:id=123,.parent=100"
      _StyleDefs(122) =   "Splits(1).Columns(1).FooterStyle:id=124,.parent=101"
      _StyleDefs(123) =   "Splits(1).Columns(1).EditorStyle:id=125,.parent=103"
      _StyleDefs(124) =   "Splits(1).Columns(2).Style:id=130,.parent=99"
      _StyleDefs(125) =   "Splits(1).Columns(2).HeadingStyle:id=127,.parent=100"
      _StyleDefs(126) =   "Splits(1).Columns(2).FooterStyle:id=128,.parent=101"
      _StyleDefs(127) =   "Splits(1).Columns(2).EditorStyle:id=129,.parent=103"
      _StyleDefs(128) =   "Splits(1).Columns(3).Style:id=134,.parent=99"
      _StyleDefs(129) =   "Splits(1).Columns(3).HeadingStyle:id=131,.parent=100"
      _StyleDefs(130) =   "Splits(1).Columns(3).FooterStyle:id=132,.parent=101"
      _StyleDefs(131) =   "Splits(1).Columns(3).EditorStyle:id=133,.parent=103"
      _StyleDefs(132) =   "Splits(1).Columns(4).Style:id=146,.parent=99"
      _StyleDefs(133) =   "Splits(1).Columns(4).HeadingStyle:id=143,.parent=100"
      _StyleDefs(134) =   "Splits(1).Columns(4).FooterStyle:id=144,.parent=101"
      _StyleDefs(135) =   "Splits(1).Columns(4).EditorStyle:id=145,.parent=103"
      _StyleDefs(136) =   "Splits(1).Columns(5).Style:id=150,.parent=99"
      _StyleDefs(137) =   "Splits(1).Columns(5).HeadingStyle:id=147,.parent=100"
      _StyleDefs(138) =   "Splits(1).Columns(5).FooterStyle:id=148,.parent=101"
      _StyleDefs(139) =   "Splits(1).Columns(5).EditorStyle:id=149,.parent=103"
      _StyleDefs(140) =   "Splits(1).Columns(6).Style:id=154,.parent=99"
      _StyleDefs(141) =   "Splits(1).Columns(6).HeadingStyle:id=151,.parent=100"
      _StyleDefs(142) =   "Splits(1).Columns(6).FooterStyle:id=152,.parent=101"
      _StyleDefs(143) =   "Splits(1).Columns(6).EditorStyle:id=153,.parent=103"
      _StyleDefs(144) =   "Splits(1).Columns(7).Style:id=158,.parent=99"
      _StyleDefs(145) =   "Splits(1).Columns(7).HeadingStyle:id=155,.parent=100"
      _StyleDefs(146) =   "Splits(1).Columns(7).FooterStyle:id=156,.parent=101"
      _StyleDefs(147) =   "Splits(1).Columns(7).EditorStyle:id=157,.parent=103"
      _StyleDefs(148) =   "Splits(1).Columns(8).Style:id=94,.parent=99"
      _StyleDefs(149) =   "Splits(1).Columns(8).HeadingStyle:id=91,.parent=100"
      _StyleDefs(150) =   "Splits(1).Columns(8).FooterStyle:id=92,.parent=101"
      _StyleDefs(151) =   "Splits(1).Columns(8).EditorStyle:id=93,.parent=103"
      _StyleDefs(152) =   "Splits(1).Columns(9).Style:id=170,.parent=99"
      _StyleDefs(153) =   "Splits(1).Columns(9).HeadingStyle:id=167,.parent=100"
      _StyleDefs(154) =   "Splits(1).Columns(9).FooterStyle:id=168,.parent=101"
      _StyleDefs(155) =   "Splits(1).Columns(9).EditorStyle:id=169,.parent=103"
      _StyleDefs(156) =   "Splits(1).Columns(10).Style:id=174,.parent=99"
      _StyleDefs(157) =   "Splits(1).Columns(10).HeadingStyle:id=171,.parent=100"
      _StyleDefs(158) =   "Splits(1).Columns(10).FooterStyle:id=172,.parent=101"
      _StyleDefs(159) =   "Splits(1).Columns(10).EditorStyle:id=173,.parent=103"
      _StyleDefs(160) =   "Splits(1).Columns(11).Style:id=54,.parent=99,.alignment=2"
      _StyleDefs(161) =   "Splits(1).Columns(11).HeadingStyle:id=51,.parent=100"
      _StyleDefs(162) =   "Splits(1).Columns(11).FooterStyle:id=52,.parent=101"
      _StyleDefs(163) =   "Splits(1).Columns(11).EditorStyle:id=53,.parent=103"
      _StyleDefs(164) =   "Splits(1).Columns(12).Style:id=178,.parent=99,.alignment=2"
      _StyleDefs(165) =   "Splits(1).Columns(12).HeadingStyle:id=175,.parent=100"
      _StyleDefs(166) =   "Splits(1).Columns(12).FooterStyle:id=176,.parent=101"
      _StyleDefs(167) =   "Splits(1).Columns(12).EditorStyle:id=177,.parent=103"
      _StyleDefs(168) =   "Splits(1).Columns(13).Style:id=142,.parent=99,.alignment=2"
      _StyleDefs(169) =   "Splits(1).Columns(13).HeadingStyle:id=139,.parent=100"
      _StyleDefs(170) =   "Splits(1).Columns(13).FooterStyle:id=140,.parent=101"
      _StyleDefs(171) =   "Splits(1).Columns(13).EditorStyle:id=141,.parent=103"
      _StyleDefs(172) =   "Splits(1).Columns(14).Style:id=190,.parent=99"
      _StyleDefs(173) =   "Splits(1).Columns(14).HeadingStyle:id=187,.parent=100"
      _StyleDefs(174) =   "Splits(1).Columns(14).FooterStyle:id=188,.parent=101"
      _StyleDefs(175) =   "Splits(1).Columns(14).EditorStyle:id=189,.parent=103"
      _StyleDefs(176) =   "Splits(1).Columns(15).Style:id=138,.parent=99"
      _StyleDefs(177) =   "Splits(1).Columns(15).HeadingStyle:id=135,.parent=100"
      _StyleDefs(178) =   "Splits(1).Columns(15).FooterStyle:id=136,.parent=101"
      _StyleDefs(179) =   "Splits(1).Columns(15).EditorStyle:id=137,.parent=103"
      _StyleDefs(180) =   "Splits(1).Columns(16).Style:id=166,.parent=99"
      _StyleDefs(181) =   "Splits(1).Columns(16).HeadingStyle:id=163,.parent=100"
      _StyleDefs(182) =   "Splits(1).Columns(16).FooterStyle:id=164,.parent=101"
      _StyleDefs(183) =   "Splits(1).Columns(16).EditorStyle:id=165,.parent=103"
      _StyleDefs(184) =   "Named:id=33:Normal"
      _StyleDefs(185) =   ":id=33,.parent=0"
      _StyleDefs(186) =   "Named:id=34:Heading"
      _StyleDefs(187) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(188) =   ":id=34,.wraptext=-1"
      _StyleDefs(189) =   "Named:id=35:Footing"
      _StyleDefs(190) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(191) =   "Named:id=36:Selected"
      _StyleDefs(192) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(193) =   "Named:id=37:Caption"
      _StyleDefs(194) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(195) =   "Named:id=38:HighlightRow"
      _StyleDefs(196) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(197) =   "Named:id=39:EvenRow"
      _StyleDefs(198) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(199) =   "Named:id=40:OddRow"
      _StyleDefs(200) =   ":id=40,.parent=33"
      _StyleDefs(201) =   "Named:id=41:RecordSelector"
      _StyleDefs(202) =   ":id=41,.parent=34"
      _StyleDefs(203) =   "Named:id=42:FilterBar"
      _StyleDefs(204) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1740
      OleObjectBlob   =   "frm_trans_check.frx":7A8B
      TabIndex        =   1
      Top             =   810
      Width           =   1695
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_division 
      Height          =   375
      Left            =   1740
      OleObjectBlob   =   "frm_trans_check.frx":9A49
      TabIndex        =   29
      Top             =   1230
      Width           =   1695
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "DIVISI"
      Height          =   195
      Left            =   270
      TabIndex        =   30
      Top             =   1290
      Width           =   465
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "IJIN KHUSUS"
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
      Left            =   360
      TabIndex        =   23
      Top             =   150
      Width           =   2775
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   270
      TabIndex        =   13
      Top             =   870
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_check.frx":BA08
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14760
   End
End
Attribute VB_Name = "frm_trans_manual_check"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCompany As New ADODB.Recordset
Dim rsDivision As New ADODB.Recordset
Dim rsCheck As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Function check_validate_exist_new() As Boolean
    check_validate_exist_new = False
    
    SQL = "select count(employee_code) as rec_count from t_check where employee_code = '" & txt_employee_code.Text & "' " & _
                "and left(check_date,10)= '" & Format(DTPicker_check.Value, "yyyy-MM-dd") & "' " & _
                "and company_code= '" & TDBCombo_company.Text & "' " & _
                "and flag_io= '" & cbo_flag_io.ListIndex & "'"
    rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        rs.Close
        Exit Function
    End If
    rs.Close
    
    '+++++++++++++++++++APAKAH KARYAWAN SUDAH PERNAH DI INPUT++++++++++++++++++++
    SQL = "SELECT att_date,shift_code,shift_number,status FROM h_attendance WHERE employee_code = '" & txt_employee_code.Text & "' " _
        & "AND DATE(att_date) = '" & Format(DTPicker_check.Value, "yyyy-MM-dd") & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
    If rs.RecordCount > 0 Then
        Dim v_shift As String
        v_shift = IIf(rs!Status = "H", "HADIR", IIf(rs!Status = "A", "ALPHA", IIf(rs!Status = "T", "TUGAS DINAS", _
                IIf(rs!Status = "L", "LIBUR", IIf(rs!Status = "I", "IJIN", "SAKIT")))))
        
        Dim i As Integer
        i = MsgBox("Data Karyawan Sudah Ada Dengan Status " & v_shift & Chr(13) & _
                    "Apakah Yakin Akan Menginputkan Ijin Untuk Karyawan Ini?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then
            check_validate_exist_new = False
            rs.Close
            Exit Function
        Else
            check_validate_exist_new = False
        End If
        rs.Close
        Exit Function
    End If
    rs.Close
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
End Function

Private Sub check_invalid()
    MsgBox "Data Sudah Ada...", vbCritical, headerMSG
    txt_nik = ""
    txt_employee_code.Text = ""
    txt_employee_name.Text = ""
    If txt_nik.Enabled = True Then txt_nik.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If Not txt_employee_code = rsCheck.Fields("employee_code").Value And _
    check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    'validasi employee code
    If Trim(txt_nik) = "" Then
        MsgBox "Kode Karyawan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        txt_nik.SetFocus
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
    
    'validasi description
    If Trim(txt_description) = "" Then
        MsgBox "Deskripsi Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        txt_description.SetFocus
        check_validate_new = False
        Exit Function
    End If
End Function

Private Sub load_data()
    timer1.Enabled = True
End Sub

Private Sub cmd_browse_Click()
    frm_lookup_mst_employee.public_int_mode = 10
    frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
    frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_refresh_Click()
    Call load_data_check
End Sub

Private Sub btnApprove_Click()
    Dim i As Integer
    i = MsgBox("Apakah Yakin Akan Menyetujui Ijin Ini", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    CnG.BeginTrans
        SQL = "UPDATE t_check SET flag_approval = 1," & _
                "user_approval = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
                    "and left(check_date,10)= '" & Format(rsCheck.Fields("check_date").Value, "yyyy-mm-dd") & "' " & _
                    "and company_code= '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
    CnG.CommitTrans
    
    btnApprove.Visible = False
    Call load_data_check
End Sub

Private Sub cmdCancel_Click()
    int_mode = 0
    Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer

On Error GoTo Err
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
        & TDBGrid1.Columns("employee_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    CnG.BeginTrans
    CnG.Execute "delete from t_check where employee_code = '" & rsCheck.Fields("employee_code").Value & "' " & _
                "and left(check_date,10)= '" & Format(rsCheck.Fields("check_date").Value, "yyyy-MM-dd") & "' " & _
                "and company_code= '" & TDBCombo_company.Text & "' " & _
                "and flag_io= '" & TDBGrid1.Columns("flag_io").Value & "'"
    CnG.CommitTrans
    
    Call load_data_check
    int_mode = 0
    Call load_mode
    
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
    vSetData = 1
    
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    With rsCheck
        txt_employee_code = .Fields("employee_code").Value
        txt_nik = .Fields("nik").Value
        txt_employee_name = .Fields("employee_name").Value
        cbo_flag_io.ListIndex = .Fields("flag_io").Value
        DTPicker_check.Value = .Fields("check_date").Value
        txt_description = .Fields("description").Value
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
    
    btnApprove.Visible = False
    lblVerify.Caption = ""
End Sub

Private Sub insert_new_data()
On Error GoTo Err

    CnG.BeginTrans
    
    SQL = "INSERT INTO t_check(employee_code,check_date,company_code,flag_io," & _
            "description,entry_date,entry_user) " & _
          "VALUES ( " & _
            "'" & txt_employee_code & "','" & Format(DTPicker_check.Value, "yyyy-MM-dd hh:mm:ss") & "'," & _
            "'" & TDBCombo_company.Text & "','" & cbo_flag_io.ListIndex & "','" & txt_description.Text & "'," & _
            "now(),'" & LOGIN_NAME & "')"
    CnG.Execute SQL
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err

    SQL = "delete from t_check where employee_code = '" & rsCheck.Fields("employee_code").Value & "' " & _
                "and left(check_date,10)= '" & Format(rsCheck.Fields("check_date").Value, "yyyy-MM-dd") & "' " & _
                "and company_code= '" & TDBCombo_company.Text & "' " & _
                "and flag_io= '" & TDBGrid1.Columns("flag_io").Value & "'"
    CnG.Execute SQL
    Call insert_new_data

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
    
    Call load_data_check
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
            If Not LCase(Ctr.name) = "txt_company_name" And Not LCase(Ctr.name) = "txt_division_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" And Not LCase(Ctr.name) = "tdbcombo_division" Then Ctr.Text = ""
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
    txt_employee_code.Text = ""
    txt_nik.Text = ""
    txt_employee_name.Text = ""
    cbo_flag_io.ListIndex = 0
    DTPicker_check.Value = Now
    txt_description.Text = ""
    lblVerify.Caption = ""
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        If TDBCombo_company.Text = "" Then
            MsgBox "Perusahaan Belum Dipilih...", vbExclamation, headerMSG
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
        TDBGrid1.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
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

Private Sub Form_Load()
    Call createGridKar
    Call load_data_company
    
    btnApprove.Visible = False
        
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    timer1.Enabled = True
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsCheck.Filter = adFilterNone
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
Dim i As Integer

On Error GoTo Err
    Set Cols = TDBGrid1.Columns
    i = TDBGrid1.Col
    TDBGrid1.HoldFields
    
    rsCheck.Filter = getFilter()
    TDBGrid1.Col = i
    TDBGrid1.EditActive = True
    
    TDBGrid1.SelStart = Len(TDBGrid1.Columns(i).FilterText)
    If TDBGrid1.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid1.Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
        
        Call load_data_division
        Call load_data_check
    End If
End Sub

Private Sub tdbcombo_division_itemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name = TDBCombo_division.Columns("division_name").Value
        
        Call load_data_check
    End If
End Sub

Private Sub tdbcombo_division_Change()
    If TDBCombo_division.Text = "" Then
        txt_division_name.Text = ""
        Call load_data_check
    Else
        Call load_data_check
    End If
End Sub

Private Sub load_data_check()
    If rsCheck.State Then rsCheck.Close
    SQL = "SELECT b.nik,b.employee_name,b.company_code,c.company_name," & _
            "b.division_code,e.division_name,b.title_code," & _
            "f.title_name,a.employee_code,a.check_date,a.flag_io,a.description," & _
            "a.entry_date,a.flag_approval,a.user_approval " & _
          "FROM t_check a JOIN m_employee b ON a.employee_code = b.employee_code " & _
            "JOIN m_company c ON b.company_code = c.company_code " & _
            "JOIN m_division e ON b.division_code = e.division_code " & _
                "AND b.company_code = e.company_code " & _
            "JOIN m_title f ON b.title_code = f.title_code " & _
          "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
            "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " & _
          "ORDER BY a.check_date, a.flag_io"
    rsCheck.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsCheck
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data_division()
    If rsDivision.State Then rsDivision.Close
    SQL = "select * from m_division where company_code = '" & TDBCombo_company.Text & "' order by division_code"
    rsDivision.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_division.RowSource = rsDivision
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If TDBGrid1.Columns(ColIndex).Caption = "DATE" Then
        Value = Format(Value, "yyyy-MM-dd hh:mm:ss")
    End If
End Sub

Private Sub TDBGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rs2 As New ADODB.Recordset
Dim vFlagApproval As Integer
Dim vUserApproval As String
Dim a As String
    
    a = IIf(IsNull(TDBGrid1.Columns("flag_approval").Value), 0, TDBGrid1.Columns("flag_approval").Value)
    vFlagApproval = IIf(a = "", 0, a)
    
    SQL = "SELECT a.employee_code FROM m_form_approval_dtl a join m_user b on a.employee_code = b.employee_code " & _
            "WHERE a.form_name = 'frm_trans_manual_check' and b.user_name = '" & LOGIN_NAME & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Or LOGIN_LEVEL = 100 Then
        If vFlagApproval = 0 Then
            btnApprove.Visible = True
            lblVerify.Visible = False
        Else
            btnApprove.Visible = False
            lblVerify.Visible = True
            
            SQL = "SELECT a.*,b.employee_name FROM m_user a LEFT JOIN m_employee b ON a.employee_code = b.employee_code " & _
                    "WHERE a.user_name = '" & TDBGrid1.Columns("user_approval").Value & "'"
            rs2.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
            
            If rs2.RecordCount > 0 Then
                If rs2!user_level = 100 Then
                    vUserApproval = "ADMINISTRATOR"
                Else
                    vUserApproval = rs2!EMPLOYEE_NAME
                End If
            End If
            rs2.Close
            
            lblVerify.Caption = "Disetujui Oleh " & vUserApproval
        End If
    Else
        btnApprove.Visible = False
    End If
    rs.Close
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub createGridKar()
   With LynxGrid2
      .AddColumn "KODE KARY.", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "NAMA KARY.", 2000, , , , , , , , , True
      .AddColumn "Div. code", , , , , , , , , False
      .AddColumn "DIVISI", 1300, , , , , , , , , True
      .AddColumn "title code", , , , , , , , , False
      .AddColumn "JABATAN", 1300, , , , , , , , , True
      .AddColumn "Employee Code", 3000, , , , , , , , False
      .BackColorBkg = &HFCE1CB
      .Redraw = True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub isiGridKar(pilihan As Integer)
    If pilihan = 1 Then
        LynxGrid2.Clear
        If LOGIN_LEVEL = 100 Then
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                            "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                        & "AND (a.nik LIKE '%" & txt_nik.Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_nik.Text & "%') " _
                        & "AND a.flag_active <> 0"
        Else
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE " & IIf(TDBCombo_division.Text = "", "b.company_code = '" & TDBCombo_company.Text & "'", _
                            "b.company_code = '" & TDBCombo_company.Text & "' AND b.division_code = '" & TDBCombo_division.Text & "'") & " " _
                        & "AND (a.nik LIKE '%" & txt_nik.Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_nik.Text & "%') " _
                        & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                        & "ORDER BY a.employee_name ASC"

        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid2.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid2.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME _
                                & vbTab & rs!division_code & vbTab & rs!division_name _
                                & vbTab & rs!title_code & vbTab & rs!title_name _
                                & vbTab & rs!employee_code
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
            txt_employee_code.Text = LynxGrid2.CellText(LynxGrid2.Row, 6)
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

