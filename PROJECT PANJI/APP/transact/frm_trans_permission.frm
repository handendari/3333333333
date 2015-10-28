VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_permission 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FORM IJIN"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12540
   Icon            =   "frm_trans_permission.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   2115
      Left            =   5040
      TabIndex        =   28
      Top             =   4080
      Visible         =   0   'False
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   3731
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
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   90
      TabIndex        =   21
      Top             =   5370
      Width           =   12375
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   30
         Top             =   150
      End
      Begin prj_panji.vbButton cmdNew 
         Height          =   705
         Left            =   660
         TabIndex        =   22
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
         MICON           =   "frm_trans_permission.frx":058A
         PICN            =   "frm_trans_permission.frx":05A6
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
         TabIndex        =   23
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
         MICON           =   "frm_trans_permission.frx":1638
         PICN            =   "frm_trans_permission.frx":1654
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
         TabIndex        =   24
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
         MICON           =   "frm_trans_permission.frx":26E6
         PICN            =   "frm_trans_permission.frx":2702
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
         TabIndex        =   25
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
         MICON           =   "frm_trans_permission.frx":3794
         PICN            =   "frm_trans_permission.frx":37B0
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
         TabIndex        =   26
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
         MICON           =   "frm_trans_permission.frx":4842
         PICN            =   "frm_trans_permission.frx":485E
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
         Left            =   10830
         TabIndex        =   27
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
         MICON           =   "frm_trans_permission.frx":58F0
         PICN            =   "frm_trans_permission.frx":590C
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
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   18
      Top             =   780
      Width           =   2865
   End
   Begin VB.Frame fra_entry 
      Height          =   2385
      Left            =   90
      TabIndex        =   6
      Top             =   2940
      Width           =   12375
      Begin VB.TextBox txt_description 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1800
         Width           =   4575
      End
      Begin VB.TextBox txt_permission 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6600
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1290
         Width           =   495
      End
      Begin VB.TextBox txt_employee_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   6750
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   8
         Top             =   810
         Width           =   2805
      End
      Begin VB.TextBox txt_nik 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   1
         Top             =   810
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   4950
         TabIndex        =   0
         Top             =   360
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   87752707
         CurrentDate     =   40823
      End
      Begin prj_panji.vbButton cmdBrowse 
         Height          =   315
         Left            =   6315
         TabIndex        =   10
         Top             =   810
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
         MICON           =   "frm_trans_permission.frx":699E
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
         Left            =   6120
         TabIndex        =   9
         Top             =   810
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "="
         Height          =   225
         Left            =   7680
         TabIndex        =   36
         Top             =   1320
         Width           =   165
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KETERANGAN :"
         Height          =   195
         Left            =   3675
         TabIndex        =   30
         Top             =   1860
         Width           =   1200
      End
      Begin VB.Label lblPotongan 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7980
         TabIndex        =   17
         Top             =   1350
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "X"
         Height          =   225
         Left            =   6330
         TabIndex        =   16
         Top             =   1350
         Width           =   165
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "8"
         Height          =   195
         Left            =   5520
         TabIndex        =   15
         Top             =   1530
         Width           =   105
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   4950
         X2              =   6210
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Label lblBasicDaily 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4980
         TabIndex        =   14
         Top             =   1200
         Width           =   1185
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "POTONGAN :"
         Height          =   195
         Left            =   3870
         TabIndex        =   13
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "JAM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7170
         TabIndex        =   12
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KODE KARY :"
         Height          =   195
         Left            =   3840
         TabIndex        =   11
         Top             =   840
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TANGGAL :"
         Height          =   195
         Left            =   3840
         TabIndex        =   7
         Top             =   420
         Width           =   1035
      End
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1560
      OleObjectBlob   =   "frm_trans_permission.frx":69BA
      TabIndex        =   19
      Top             =   780
      Width           =   1695
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   3645
      Left            =   90
      TabIndex        =   29
      Top             =   1680
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   6429
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "TANGGAL"
      Columns(0).DataField=   "permission_date"
      Columns(0).NumberFormat=   "yyyy-MM-dd"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EMPLOYEE CODE"
      Columns(1).DataField=   "employee_code"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "KODE KARY."
      Columns(2).DataField=   "nik"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "NAMA KARY."
      Columns(3).DataField=   "employee_name"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "JAM IJIN (JAM)"
      Columns(4).DataField=   "permission_hour"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "POTONGAN"
      Columns(5).DataField=   "permission_value"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "KETERANGAN"
      Columns(6).DataField=   "description"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2170"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2090"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2434"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2355"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=5292"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=5212"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=5450"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=5371"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=512"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
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
      Caption         =   "DAFTAR IJIN"
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
      _StyleDefs(21)  =   "Splits(0).Style:id=99,.parent=1"
      _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=116,.parent=4,.bgcolor=&H80000002&"
      _StyleDefs(23)  =   ":id=116,.fgcolor=&H80000009&"
      _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=100,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(25)  =   ":id=100,.fgcolor=&H80000002&"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=101,.parent=3"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=102,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=104,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=103,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=105,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=106,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=115,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=117,.parent=11"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=118,.parent=12"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=54,.parent=99,.alignment=2"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=100"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=101"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=103"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=16,.parent=99,.alignment=2"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=13,.parent=100"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=14,.parent=101"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=15,.parent=103"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=32,.parent=99,.alignment=2"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=100"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=101"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=103"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=28,.parent=99"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=100"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=101"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=103"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=20,.parent=99,.alignment=2"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=17,.parent=100"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=18,.parent=101"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=19,.parent=103"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=24,.parent=99,.alignment=1"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=21,.parent=100"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=22,.parent=101"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=23,.parent=103"
      _StyleDefs(59)  =   "Splits(0).Columns(6).Style:id=62,.parent=99,.alignment=0"
      _StyleDefs(60)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=100"
      _StyleDefs(61)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=101"
      _StyleDefs(62)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=103"
      _StyleDefs(63)  =   "Named:id=33:Normal"
      _StyleDefs(64)  =   ":id=33,.parent=0"
      _StyleDefs(65)  =   "Named:id=34:Heading"
      _StyleDefs(66)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   ":id=34,.wraptext=-1"
      _StyleDefs(68)  =   "Named:id=35:Footing"
      _StyleDefs(69)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   "Named:id=36:Selected"
      _StyleDefs(71)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(72)  =   "Named:id=37:Caption"
      _StyleDefs(73)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(74)  =   "Named:id=38:HighlightRow"
      _StyleDefs(75)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(76)  =   "Named:id=39:EvenRow"
      _StyleDefs(77)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(78)  =   "Named:id=40:OddRow"
      _StyleDefs(79)  =   ":id=40,.parent=33"
      _StyleDefs(80)  =   "Named:id=41:RecordSelector"
      _StyleDefs(81)  =   ":id=41,.parent=34"
      _StyleDefs(82)  =   "Named:id=42:FilterBar"
      _StyleDefs(83)  =   ":id=42,.parent=33"
   End
   Begin MSComCtl2.DTPicker DTPicker_from 
      Height          =   315
      Left            =   1560
      TabIndex        =   31
      Top             =   1200
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   87752707
      CurrentDate     =   40794
   End
   Begin MSComCtl2.DTPicker DTPicker_to 
      Height          =   315
      Left            =   3300
      TabIndex        =   32
      Top             =   1200
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MM-yyyy"
      Format          =   87752707
      CurrentDate     =   40794
   End
   Begin prj_panji.vbButton cmdSearch 
      Height          =   465
      Left            =   4770
      TabIndex        =   35
      Top             =   1140
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Lihat"
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
      MICON           =   "frm_trans_permission.frx":8920
      PICN            =   "frm_trans_permission.frx":893C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "TANGGAL :"
      Height          =   195
      Left            =   630
      TabIndex        =   34
      Top             =   1230
      Width           =   855
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
      Left            =   2970
      TabIndex        =   33
      Top             =   1230
      Width           =   285
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN :"
      Height          =   195
      Left            =   300
      TabIndex        =   20
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FORM IJIN"
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
      Left            =   300
      TabIndex        =   5
      Top             =   180
      Width           =   4845
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_permission.frx":99CE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12540
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT MANUAL ATTENDANCE"
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
      TabIndex        =   4
      Top             =   150
      Width           =   4845
   End
End
Attribute VB_Name = "frm_trans_permission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCompany As New ADODB.Recordset
Dim rsPermission As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Private Function check_validate_exist_new() As Boolean
    check_validate_exist_new = False
    
    SQL = "select count(employee_code) as rec_count from t_permission where employee_code = '" & txt_employee_code.Text & "' " & _
                "and left(permission_date,10)= '" & Format(DTPicker1.Value, "yyyy-mm-dd") & "'"
    rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        rs.Close
        Exit Function
    End If
    
    rs.Close
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
    
    If Not txt_employee_code = rsPermission.Fields("employee_code").Value And _
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
End Function

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
    CnG.Execute "delete from t_permission where employee_code = '" & rsPermission.Fields("employee_code").Value & "' " & _
                    "and left(permission_date,10)= '" & Format(rsPermission.Fields("permission_date").Value, "yyyy-mm-dd") & "'"
    CnG.CommitTrans
    
    Call load_data_permission
    int_mode = 0
    Call load_mode
    
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
Dim v_verify As String
Dim v_level As String

    vSetData = 1
    
    If Not (TDBGrid1.ApproxCount > 0 And TDBGrid1.Bookmark > 0) Then
        MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    With rsPermission
        DTPicker1 = .Fields("permission_date").Value
        txt_employee_code = .Fields("employee_code").Value
        txt_nik = .Fields("nik").Value
        txt_employee_name = .Fields("employee_name").Value
        txt_permission.Text = .Fields("permission_hour").Value
        lblBasicDaily.Caption = FormatNumber(.Fields("basic_daily").Value)
        lblPotongan.Caption = FormatNumber(.Fields("permission_value").Value)
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
End Sub

Private Sub insert_new_data()
Dim v_total_ot As Double
Dim vEndTime As String

On Error GoTo Err
    CnG.BeginTrans
    SQL = "INSERT INTO t_permission(permission_date,employee_code,permission_hour,basic_daily," & _
            "permission_value,description,entry_date,entry_user) " & _
          "VALUES( " & _
            "'" & Format(DTPicker1.Value, "yyyy-MM-dd") & "','" & txt_employee_code.Text & "','" & txt_permission.Text & "'," & _
            "'" & Val(DropAllComma(lblBasicDaily.Caption)) & "','" & Val(DropAllComma(lblPotongan.Caption)) & "'," & _
            "'" & txt_description & "',now(),'" & LOGIN_NAME & "')"
    CnG.Execute SQL

    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err

    SQL = "DELETE FROM t_permission WHERE employee_code = '" & TDBGrid1.Columns("employee_code").Value & "' " & _
            "AND date(permission_date) = '" & Format(TDBGrid1.Columns("permission_date").Value, "yyyy-MM-dd") & "'"
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
    
    Call load_data_permission
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
            If Not LCase(Ctr.name) = "dtpicker_from" And Not LCase(Ctr.name) = "dtpicker_to" Then Ctr.Value = Now
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
    DTPicker1.Value = Now
    txt_employee_code.Text = ""
    txt_nik.Text = ""
    txt_employee_name.Text = ""
    txt_permission.Text = 0
    txt_description.Text = ""
    
    lblBasicDaily.Caption = ""
    lblPotongan.Caption = ""
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

Private Sub cmdSearch_Click()
    If TDBCombo_company.Text = "" Then
        MsgBox "Perusahaan Masih Kosong!", vbExclamation, headerMSG
        TDBCombo_company.SetFocus
        Exit Sub
    End If
    
    Call load_data_permission
End Sub

Private Sub Form_Load()
    Call createGridKar
    Call load_data_company
    
    DTPicker_from.Value = Now
    DTPicker_to.Value = Now
        
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    timer1.Enabled = True
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsPermission.Filter = adFilterNone
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
    
    rsPermission.Filter = getFilter()
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
    End If
End Sub

Private Sub load_data_permission()
    If rsPermission.State Then rsPermission.Close
    
    SQL = "SELECT a.*,b.nik,b.employee_name " & _
          "FROM t_permission a JOIN m_employee b ON a.employee_code = b.employee_code " & _
          "WHERE b.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                "AND DATE(permission_date) BETWEEN '" & Format(DTPicker_from.Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to, "yyyy-MM-dd") & "' " & _
                "ORDER BY permission_date ASC"
    rsPermission.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsPermission
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub TDBGrid1_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If TDBGrid1.Columns(ColIndex).Caption = "DATE FROM" _
    Or TDBGrid1.Columns(ColIndex).Caption = "DATE TO" Then
        Value = Format(Value, "yyyy-mm-dd")
    End If
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
                    & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
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
                    & "WHERE b.company_code = '" & TDBCombo_company.Text & "' " _
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
        
    SQL = "SELECT CASE WHEN '" & Format(DTPicker1.Value, "dddd") & "' = 'Sunday' THEN basic_salary_sunday " & _
            "ELSE CASE WHEN flag_basic = 0 THEN basic_salary ELSE (basic_salary/30) END END basic_salary " & _
          "FROM m_salary_standard " & _
          "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
            "AND date(salary_date) <= '" & Format(DTPicker1.Value, "yyyy-MM-dd") & "' " & _
          "ORDER BY salary_date DESC LIMIT 1"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        lblBasicDaily.Caption = FormatNumber(rscari!basic_salary)
    Else
        lblBasicDaily.Caption = FormatNumber(0)
    End If
    rscari.Close
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

Private Sub txt_permission_Validate(Cancel As Boolean)
    lblPotongan.Caption = (lblBasicDaily.Caption / 8) * txt_permission.Text
    lblPotongan.Caption = FormatNumber(lblPotongan.Caption)
End Sub
