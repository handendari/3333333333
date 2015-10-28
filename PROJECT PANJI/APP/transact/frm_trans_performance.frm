VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_performance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PERFORMA KARYAWAN"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12675
   Icon            =   "frm_trans_performance.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10155
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   Begin prj_panji.vbButton cmdBrowse 
      Height          =   285
      Left            =   3060
      TabIndex        =   33
      Top             =   1890
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
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
      MICON           =   "frm_trans_performance.frx":058A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.LynxGrid LynxGrid2 
      Height          =   2925
      Left            =   1500
      TabIndex        =   28
      Top             =   2190
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
   Begin VB.TextBox txt_employee_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      DragMode        =   1  'Automatic
      Height          =   285
      Left            =   3420
      TabIndex        =   30
      Top             =   1890
      Width           =   3495
   End
   Begin VB.TextBox txt_nik 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1500
      TabIndex        =   29
      Top             =   1890
      Width           =   1515
   End
   Begin VB.TextBox txt_division_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   1530
      Width           =   3855
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3060
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   22
      Top             =   1170
      Width           =   3855
   End
   Begin VB.Frame Frame5 
      Caption         =   "Data Control Button"
      Height          =   1335
      Index           =   4
      Left            =   210
      TabIndex        =   8
      Top             =   8130
      Width           =   11745
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   0
         Top             =   0
      End
      Begin prj_panji.vbButton cmdNew 
         Height          =   705
         Left            =   540
         TabIndex        =   9
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
         MICON           =   "frm_trans_performance.frx":05A6
         PICN            =   "frm_trans_performance.frx":05C2
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
         Left            =   1560
         TabIndex        =   10
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
         MICON           =   "frm_trans_performance.frx":1654
         PICN            =   "frm_trans_performance.frx":1670
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
         Left            =   2580
         TabIndex        =   11
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
         MICON           =   "frm_trans_performance.frx":2702
         PICN            =   "frm_trans_performance.frx":271E
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
         Left            =   3600
         TabIndex        =   12
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
         MICON           =   "frm_trans_performance.frx":37B0
         PICN            =   "frm_trans_performance.frx":37CC
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
         Left            =   4620
         TabIndex        =   13
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
         MICON           =   "frm_trans_performance.frx":485E
         PICN            =   "frm_trans_performance.frx":487A
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
         Left            =   10590
         TabIndex        =   21
         Top             =   360
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
         MICON           =   "frm_trans_performance.frx":590C
         PICN            =   "frm_trans_performance.frx":5928
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
      Height          =   2325
      Left            =   210
      TabIndex        =   1
      Top             =   2790
      Width           =   11745
      Begin VB.TextBox txt_perf_description 
         Appearance      =   0  'Flat
         Height          =   555
         Left            =   5160
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1290
         Width           =   3705
      End
      Begin VB.TextBox txt_perf_value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5160
         TabIndex        =   2
         Top             =   960
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker DTPicker_perf 
         Height          =   315
         Left            =   5160
         TabIndex        =   3
         Top             =   600
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         MousePointer    =   99
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   107610115
         CurrentDate     =   39270
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TANGGAL*"
         Height          =   195
         Left            =   3390
         TabIndex        =   7
         Top             =   600
         Width           =   825
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KETERANGAN"
         Height          =   195
         Left            =   3165
         TabIndex        =   6
         Top             =   1320
         Width           =   1110
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "NILAI*"
         Height          =   195
         Left            =   3750
         TabIndex        =   5
         Top             =   930
         Width           =   465
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid_Perf 
      Height          =   2805
      Left            =   210
      TabIndex        =   14
      Top             =   2310
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   4948
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
      Columns(1).Caption=   "SEQ NO"
      Columns(1).DataField=   "seq_no"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TANGGAL"
      Columns(2).DataField=   "perf_date"
      Columns(2).NumberFormat=   "yyyy-MM-dd"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "NILAI"
      Columns(3).DataField=   "perf_value"
      Columns(3).NumberFormat=   "General Number"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "KETERANGAN"
      Columns(4).DataField=   "description"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2963"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2884"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=3413"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=3334"
      Splits(0)._ColumnProps(16)=   "Column(2)._ColStyle=513"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=4551"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=4471"
      Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=11642"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=11562"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
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
      Caption         =   "DAFTAR PERFORMA KARYAWAN"
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
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=2"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=3"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
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
   Begin TrueOleDBGrid70.TDBGrid TDBGrid_Perf_Result 
      Height          =   2805
      Left            =   210
      TabIndex        =   15
      Top             =   5280
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   4948
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
      Columns(1).Caption=   "TAHUN"
      Columns(1).DataField=   "perf_year"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "RATA - RATA"
      Columns(2).DataField=   "perf_avg"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "GRADE"
      Columns(3).DataField=   "perf_grade"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2963"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2884"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2117"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=4815"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=4736"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=514"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=3678"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=3598"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos(0).PreviewInitHeight=   -1
      PrintInfos(0).PreviewInitScreenFill=   -1
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "DAFTAR HASIL PERFORMA KARYAWAN"
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
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=54,.parent=13,.alignment=2"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=2"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(50)  =   "Named:id=33:Normal"
      _StyleDefs(51)  =   ":id=33,.parent=0"
      _StyleDefs(52)  =   "Named:id=34:Heading"
      _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   ":id=34,.wraptext=-1"
      _StyleDefs(55)  =   "Named:id=35:Footing"
      _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=36:Selected"
      _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(59)  =   "Named:id=37:Caption"
      _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(61)  =   "Named:id=38:HighlightRow"
      _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(63)  =   "Named:id=39:EvenRow"
      _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(65)  =   "Named:id=40:OddRow"
      _StyleDefs(66)  =   ":id=40,.parent=33"
      _StyleDefs(67)  =   "Named:id=41:RecordSelector"
      _StyleDefs(68)  =   ":id=41,.parent=34"
      _StyleDefs(69)  =   "Named:id=42:FilterBar"
      _StyleDefs(70)  =   ":id=42,.parent=33"
   End
   Begin MSComCtl2.DTPicker DTPicker_perf_from 
      Height          =   315
      Left            =   1500
      TabIndex        =   16
      Top             =   810
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy"
      Format          =   107610115
      CurrentDate     =   39270
   End
   Begin MSComCtl2.DTPicker DTPicker_perf_to 
      Height          =   315
      Left            =   3420
      TabIndex        =   17
      Top             =   810
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
      _Version        =   393216
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy"
      Format          =   107610115
      CurrentDate     =   39270
   End
   Begin prj_panji.vbButton cmdSearch 
      Height          =   465
      Left            =   7080
      TabIndex        =   18
      Top             =   1710
      Width           =   1275
      _ExtentX        =   2249
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
      MICON           =   "frm_trans_performance.frx":69BA
      PICN            =   "frm_trans_performance.frx":69D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1500
      OleObjectBlob   =   "frm_trans_performance.frx":7A68
      TabIndex        =   23
      Top             =   1170
      Width           =   1545
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_division 
      Height          =   375
      Left            =   1500
      OleObjectBlob   =   "frm_trans_performance.frx":9A26
      TabIndex        =   27
      Top             =   1530
      Width           =   1545
   End
   Begin VB.TextBox txt_employee_code 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6780
      TabIndex        =   32
      Top             =   1890
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "KARYAWAN*"
      Height          =   195
      Left            =   390
      TabIndex        =   31
      Top             =   1890
      Width           =   990
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "DIVISI"
      Height          =   195
      Left            =   900
      TabIndex        =   25
      Top             =   1560
      Width           =   465
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN*"
      Height          =   195
      Left            =   195
      TabIndex        =   24
      Top             =   1230
      Width           =   1170
   End
   Begin VB.Label Label60 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "TAHUN*"
      Height          =   195
      Left            =   750
      TabIndex        =   20
      Top             =   870
      Width           =   630
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "S/D"
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
      Left            =   3060
      TabIndex        =   19
      Top             =   870
      Width           =   360
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "PERFORMA KARYAWAN"
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
      TabIndex        =   0
      Top             =   150
      Width           =   4365
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_performance.frx":B9E5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14850
   End
End
Attribute VB_Name = "frm_trans_performance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmployee As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsDiv As New ADODB.Recordset
Dim rsPerf As New ADODB.Recordset
Dim rsPerf_Result As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim i_lang As Integer
Dim v_company As String
Dim dst As String

Dim v_empCode As String
Dim v_seq As Integer
Dim v_perf_number As Integer
Dim v_count_perf As Integer
Dim v_sum_perf As Double
Dim v_avg_perf As Double

Private Function check_validate_exist_new() As Boolean
    check_validate_exist_new = False
End Function

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    'validasi value
    If txt_perf_value.Text = "" Then
        MsgBox "Nilai Prestasi Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        txt_perf_value.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    'validasi value
    If Not IsNumeric(txt_perf_value.Text) Or Val(txt_perf_value.Text) < 0 Then
        MsgBox "Nilai Prestasi Tidak Valid...", vbOKOnly + vbInformation, headerMSG
        txt_perf_value.Text = 0
        txt_perf_value.SetFocus
        check_validate_new = False
        Exit Function
    End If

End Function
Private Sub cancel_data()
    int_mode = 0
    Call load_mode
End Sub

Private Sub delete_data()
Dim i As Integer
Dim v_perf_year As Integer

On Error GoTo Err
    If Not (TDBGrid_Perf.ApproxCount > 0 And TDBGrid_Perf.Bookmark > 0) Then
        MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
        & TDBGrid_Perf.Columns("perf_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
    
    CnG.BeginTrans
    v_perf_year = year(TDBGrid_Perf.Columns("perf_date").Value)
    
    CnG.Execute "delete from t_employee_perf where employee_code = '" _
        & TDBGrid_Perf.Columns("employee_code").Value & "' " & _
        "AND seq_no = '" & TDBGrid_Perf.Columns("seq_no").Value & "'"
    
    v_count_perf = getCountValue()
    v_sum_perf = getSumValue()
    If v_count_perf <> 0 Then
        v_avg_perf = v_sum_perf / v_count_perf
    End If
    
    SQL = "SELECT perf_number FROM m_performance " & _
          "WHERE perf_under <= '" & v_avg_perf & "' " & _
            "AND perf_upper >= '" & v_avg_perf & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        v_perf_number = rs!perf_number
    End If
    rs.Close
    
    If v_count_perf = 0 Then
        CnG.Execute "delete from t_employee_perf_result where employee_code = '" _
            & TDBGrid_Perf_Result.Columns("employee_code").Value & "' " & _
            "AND perf_year = '" & v_perf_year & "'"
    Else
        SQL = "UPDATE t_employee_perf_result SET perf_year = '" & year(DTPicker_perf.Value) & "'," & _
                  "perf_avg = '" & v_avg_perf & "',perf_number = '" & v_perf_number & "'," & _
                  "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsPerf_Result.Fields("employee_code").Value & "' " & _
                  "AND perf_year = '" & rsPerf_Result.Fields("perf_year").Value & "'"
    CnG.Execute SQL
    End If
    
    CnG.CommitTrans

    Call load_data_employee_perf
    Call load_data_employee_perf_result
    
    int_mode = 0
    Call load_mode
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Function get_data_obj(ByRef Ctr As CONTROL, ByVal str As Variant) As Variant
    If TypeOf Ctr Is ComboBox Then
        If Ctr.name = "cbo_sex" Or Ctr.name = "cbo_marital_status" Or _
            Ctr.name = "cbo_religion" Or Ctr.name = "cbo_nationality" Or _
            Ctr.name = "cbo_tax_method" Then
            get_data_obj = IIf(IsNull(str) = True, 1, str)
        End If
    ElseIf TypeOf Ctr Is DTPicker Then
        get_data_obj = IIf(IsNull(str) = True, Now, str)
    ElseIf TypeOf Ctr Is TextBox Then
        get_data_obj = IIf(IsNull(str) = True, "", str)
    ElseIf TypeOf Ctr Is Image Then
        get_data_obj = IIf(IsNull(str) = True, "", str)
    End If
End Function

Public Sub set_edit_data()
    vSetData = 1
    
    If Not (TDBGrid_Perf.ApproxCount > 0 And TDBGrid_Perf.Bookmark > 0) Then
        MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
    
    With rsPerf
        DTPicker_perf.Value = .Fields("perf_date").Value
        txt_perf_value.Text = FormatNumber(.Fields("perf_value").Value)
        txt_perf_description.Text = .Fields("description").Value
    End With
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub new_data()
    int_mode = 1
    Call load_mode
End Sub

Private Sub insert_new_data()

On Error GoTo Err
    SQL = "SELECT MAX(seq_no) jmlSeq FROM t_employee_perf " & _
            "WHERE employee_code = '" & txt_employee_code.Text & "' AND perf_year = '" & year(DTPicker_perf.Value) & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        v_seq = IIf(IsNull(rs!jmlSeq), 0, rs!jmlSeq) + 1
    Else
        v_seq = 1
    End If
    rs.Close
    
    CnG.BeginTrans
    
    SQL = "INSERT INTO t_employee_perf (employee_code,seq_no,perf_year,perf_date," & _
            "perf_value,description,entry_date,entry_user) " & _
          "VALUES (" & _
            "'" & txt_employee_code.Text & "','" & v_seq & "','" & year(DTPicker_perf.Value) & "'," & _
            "'" & Format(DTPicker_perf.Value, "yyyy-MM-dd") & "'," & _
            "'" & txt_perf_value.Text & "','" & txt_perf_description.Text & "',now(),'" & LOGIN_NAME & "')"
    CnG.Execute SQL
        
    v_count_perf = getCountValue()
    v_sum_perf = getSumValue()
    v_avg_perf = v_sum_perf / v_count_perf
    
    SQL = "SELECT perf_number FROM m_performance " & _
          "WHERE perf_under <= '" & v_avg_perf & "' " & _
            "AND perf_upper >= '" & v_avg_perf & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        v_perf_number = rs!perf_number
    End If
    rs.Close
    
    If v_count_perf = 1 Then
        SQL = "INSERT INTO t_employee_perf_result (employee_code,perf_year," & _
                "perf_avg,perf_number,entry_date,entry_user) " & _
              "VALUES (" & _
                "'" & txt_employee_code.Text & "'," & _
                "'" & year(DTPicker_perf.Value) & "','" & v_avg_perf & "'," & _
                "'" & v_perf_number & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    Else
        SQL = "UPDATE t_employee_perf_result SET perf_year = '" & year(DTPicker_perf.Value) & "'," & _
                "perf_avg = '" & v_avg_perf & "',perf_number = '" & v_perf_number & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE employee_code = '" & rsPerf_Result.Fields("employee_code").Value & "' " & _
                "AND perf_year = '" & rsPerf_Result.Fields("perf_year").Value & "'"
        CnG.Execute SQL
    End If
    
    CnG.CommitTrans
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
Dim rscari As New ADODB.Recordset

On Error GoTo Err
    CnG.BeginTrans
    SQL = "UPDATE t_employee_perf SET perf_date = '" & Format(DTPicker_perf.Value, "yyyy-MM-dd") & "'," & _
            "perf_year = '" & year(DTPicker_perf.Value) & "'," & _
            "perf_value = '" & txt_perf_value.Text & "',description = '" & txt_perf_description.Text & "'," & _
            "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
          "WHERE employee_code = '" & rsPerf.Fields("employee_code").Value & "' " & _
            "AND seq_no = '" & rsPerf.Fields("seq_no").Value & "'"
    CnG.Execute SQL
    
    v_count_perf = getCountValue()
    v_sum_perf = getSumValue()
    v_avg_perf = v_sum_perf / v_count_perf
    
    SQL = "SELECT perf_number FROM m_performance " & _
          "WHERE perf_under <= '" & v_avg_perf & "' " & _
            "AND perf_upper >= '" & v_avg_perf & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        v_perf_number = rs!perf_number
    End If
    rs.Close
    
    SQL = "UPDATE t_employee_perf_result SET perf_year = '" & year(DTPicker_perf.Value) & "'," & _
            "perf_avg = '" & v_avg_perf & "',perf_number = '" & v_perf_number & "'," & _
            "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
          "WHERE employee_code = '" & rsPerf_Result.Fields("employee_code").Value & "' " & _
            "AND perf_year = '" & rsPerf_Result.Fields("perf_year").Value & "'"
    CnG.Execute SQL
    
    CnG.CommitTrans

    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub simpan_data()
Dim clsFunc As New clsFunction
Dim int_proses As Integer
Dim SQL As String

    If int_mode = 1 Then
        If Not check_validate_new Then Exit Sub
'        If check_validate_exist_new Then
'            Call check_invalid: Exit Sub
'        End If
        Call insert_new_data
        
        clsFunc.InsertLog ("Insert Performance : " & txt_nik.Text & "-" & Format(DTPicker_perf.Value, "yyyy-MM-dd"))
    ElseIf int_mode = 2 Then
        If Not check_validate_new Then Exit Sub
    '    If check_validate_exist_edit Then
    '        Call check_invalid: Exit Sub
    '    End If
    
        Call edit_old_data
        clsFunc.InsertLog ("Edit Performance : " & txt_nik.Text & "-" & Format(DTPicker_perf.Value, "yyyy-MM-dd"))
    End If
    
    Call load_data_employee_perf
    Call load_data_employee_perf_result
    
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
            If Not LCase(Ctr.name) = "txt_company_name" And _
               Not LCase(Ctr.name) = "txt_nik" And _
               Not LCase(Ctr.name) = "txt_employee_code" And _
               Not LCase(Ctr.name) = "txt_employee_name" And _
               Not LCase(Ctr.name) = "txt_division_name" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_company" And _
               Not LCase(Ctr.name) = "tdbcombo_division" Then Ctr.Text = ""
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
    DTPicker_perf.Value = Now
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
        
        If txt_nik.Text = "" Then
            MsgBox "Kode Karyawan Masih Kosong...", vbExclamation, headerMSG
            txt_nik.SetFocus
            
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        Call clear_view_data
        Call set_new_data
        
        fra_entry.Visible = True
        txt_perf_value.Enabled = True
        TDBGrid_Perf.Enabled = False
        
        If txt_perf_value.Enabled = True Then
            txt_perf_value.SetFocus
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        
        fra_entry.Visible = False
        TDBGrid_Perf.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        fra_entry.Visible = True
        TDBGrid_Perf.Enabled = False
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
        MsgBox "Perusahaan Belum Dipilih...", vbExclamation, headerMSG
        TDBCombo_company.SetFocus
        Exit Sub
    End If
    
    If txt_nik.Text = "" Then
        MsgBox "Karyawan Belum Dipilih...", vbExclamation, headerMSG
        txt_nik.SetFocus
        Exit Sub
    End If
    
    Call load_data_employee_perf
    Call load_data_employee_perf_result
End Sub

Private Sub Form_Load()
    
    Call load_data_company
    Call createGridKar
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_trans_performance = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    DTPicker_perf_from.Value = Now
    DTPicker_perf_to.Value = Now
    
    int_mode = 0
    Call load_mode
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value

        Call load_data_division
    End If
End Sub

Private Sub tdbcombo_division_Change()
    If TDBCombo_division.Text = "" Then
        txt_division_name.Text = ""
    End If
End Sub

Private Sub TDBCombo_division_itemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name.Text = TDBCombo_division.Columns("division_name").Value
    End If
End Sub

Public Sub load_data_company()
    TDBCombo_company.Text = "": txt_company_name = ""
    
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Public Sub load_data_division()
    TDBCombo_division.Text = "": txt_division_name.Text = ""
    
    If rsDiv.State Then rsDiv.Close
    SQL = "select * from m_division where company_code = '" & TDBCombo_company.Text & "' order by company_code"
    rsDiv.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_division.RowSource = rsDiv
End Sub

Private Sub load_data_employee_perf()
    If rsPerf.State Then rsPerf.Close
    SQL = "select * from t_employee_perf " & _
          "where employee_code = '" & txt_employee_code.Text & "' " & _
            "AND (year(perf_date) between '" & year(DTPicker_perf_from.Value) & "' and '" & year(DTPicker_perf_to.Value) & "') " & _
          "order by perf_date DESC"
    rsPerf.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Perf.DataSource = rsPerf
End Sub

Private Sub load_data_employee_perf_result()
    If rsPerf_Result.State Then rsPerf_Result.Close
    SQL = "select a.*, b.perf_grade from t_employee_perf_result a join m_performance b on a.perf_number = b.perf_number " & _
          "where employee_code = '" & txt_employee_code.Text & "' " & _
            "AND (perf_year between '" & year(DTPicker_perf_from.Value) & "' and '" & year(DTPicker_perf_to.Value) & "') " & _
          "order by perf_year DESC"
    rsPerf_Result.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Perf_Result.DataSource = rsPerf_Result
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid_Perf_Result.Columns
        Col.FilterText = ""
    Next Col
    rsPerf_Result.Filter = adFilterNone
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

Private Sub grid_filter()
On Error GoTo Err

Dim i As Integer
    
    Set Cols = TDBGrid_Perf_Result.Columns
    i = TDBGrid_Perf_Result.Col
    TDBGrid_Perf_Result.HoldFields
    
    rsPerf_Result.Filter = getFilter()
    TDBGrid_Perf_Result.Col = i
    TDBGrid_Perf_Result.EditActive = True
    
    TDBGrid_Perf_Result.SelStart = Len(TDBGrid_Perf_Result.Columns(i).FilterText)
    If TDBGrid_Perf_Result.ApproxCount < 1 Then
        Call clear_filter
        TDBGrid_Perf_Result.Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Function getCountValue() As Integer
    SQL = "SELECT COUNT(employee_code) jmlData " & _
          "FROM t_employee_perf " & _
          "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
            "AND YEAR(perf_date) = '" & year(DTPicker_perf.Value) & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        getCountValue = IIf(IsNull(rs!jmlData), 0, rs!jmlData)
    Else
        getCountValue = 0
    End If
    rs.Close
End Function

Private Function getSumValue() As Double
    SQL = "SELECT SUM(perf_value) sumData " & _
          "FROM t_employee_perf " & _
          "WHERE employee_code = '" & txt_employee_code.Text & "' " & _
            "AND YEAR(perf_date) = '" & year(DTPicker_perf.Value) & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        getSumValue = IIf(IsNull(rs!sumData), 0, rs!sumData)
    Else
        getSumValue = 0
    End If
    rs.Close
End Function

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
    
    Call load_data_employee_perf
    Call load_data_employee_perf_result
End Sub

Private Sub LynxGrid2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid2.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
        
        Call load_data_employee_perf
        Call load_data_employee_perf_result
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
        
        Call load_data_employee_perf
        Call load_data_employee_perf_result
    End If
End Sub

Private Sub cmdBrowse_Click()
    isiGridKar (1)
    
    Call load_data_employee_perf
    Call load_data_employee_perf_result
End Sub

Private Sub cmdNew_Click()
    Call new_data
End Sub

Private Sub cmdSave_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Click()
    Call cancel_data
End Sub


Private Sub TDBGrid_Employee_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Perf_FilterChange()
    Call grid_filter
End Sub
