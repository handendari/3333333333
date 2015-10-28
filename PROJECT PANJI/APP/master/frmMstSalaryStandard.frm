VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_salary_standard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER GAJI"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12435
   Icon            =   "frmMstSalaryStandard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   12435
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   120
      TabIndex        =   26
      Top             =   9120
      Width           =   12165
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Interval        =   600
         Left            =   120
         Top             =   360
      End
      Begin prj_panji.vbButton cmdNew 
         Height          =   705
         Left            =   690
         TabIndex        =   27
         Top             =   300
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
         MICON           =   "frmMstSalaryStandard.frx":058A
         PICN            =   "frmMstSalaryStandard.frx":05A6
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
         Left            =   1710
         TabIndex        =   28
         Top             =   300
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
         MICON           =   "frmMstSalaryStandard.frx":1638
         PICN            =   "frmMstSalaryStandard.frx":1654
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
         Left            =   2730
         TabIndex        =   29
         Top             =   300
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
         MICON           =   "frmMstSalaryStandard.frx":26E6
         PICN            =   "frmMstSalaryStandard.frx":2702
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
         Left            =   3750
         TabIndex        =   30
         Top             =   300
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
         MICON           =   "frmMstSalaryStandard.frx":3794
         PICN            =   "frmMstSalaryStandard.frx":37B0
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
         Left            =   4770
         TabIndex        =   31
         Top             =   300
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
         MICON           =   "frmMstSalaryStandard.frx":4842
         PICN            =   "frmMstSalaryStandard.frx":485E
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
         Left            =   9810
         TabIndex        =   32
         Top             =   330
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
         MICON           =   "frmMstSalaryStandard.frx":58F0
         PICN            =   "frmMstSalaryStandard.frx":590C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdImport 
         Height          =   705
         Left            =   8700
         TabIndex        =   79
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
         MICON           =   "frmMstSalaryStandard.frx":699E
         PICN            =   "frmMstSalaryStandard.frx":69BA
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
   Begin VB.TextBox txt_division_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3090
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   22
      Top             =   1170
      Width           =   3855
   End
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   3090
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   19
      Top             =   780
      Width           =   3855
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid_Emp 
      Height          =   3705
      Left            =   120
      TabIndex        =   6
      Top             =   1590
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   6535
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
      Columns(4).Caption=   "EMPLOYEE CODE"
      Columns(4).DataField=   "employee_code"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "KODE KARY."
      Columns(5).DataField=   "nik"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "NAMA KARY."
      Columns(6).DataField=   "employee_name"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "JABATAN"
      Columns(7).DataField=   "title_name"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1826"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1746"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
      Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Width=1826"
      Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1746"
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
      Splits(0)._ColumnProps(23)=   "Column(3).Width=3731"
      Splits(0)._ColumnProps(24)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(3)._WidthInPix=3651"
      Splits(0)._ColumnProps(26)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(27)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(29)=   "Column(4).Width=1217"
      Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1138"
      Splits(0)._ColumnProps(32)=   "Column(4).AllowSizing=0"
      Splits(0)._ColumnProps(33)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(34)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(35)=   "Column(4).AllowFocus=0"
      Splits(0)._ColumnProps(36)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(37)=   "Column(5).Width=3360"
      Splits(0)._ColumnProps(38)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(5)._WidthInPix=3281"
      Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(42)=   "Column(6).Width=6350"
      Splits(0)._ColumnProps(43)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(6)._WidthInPix=6271"
      Splits(0)._ColumnProps(45)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(47)=   "Column(7).Width=6324"
      Splits(0)._ColumnProps(48)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(49)=   "Column(7)._WidthInPix=6244"
      Splits(0)._ColumnProps(50)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(51)=   "Column(7).Order=8"
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
      Caption         =   "DAFTAR KARYAWAN"
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
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=70,.parent=13"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(66)  =   "Named:id=33:Normal"
      _StyleDefs(67)  =   ":id=33,.parent=0"
      _StyleDefs(68)  =   "Named:id=34:Heading"
      _StyleDefs(69)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(70)  =   ":id=34,.wraptext=-1"
      _StyleDefs(71)  =   "Named:id=35:Footing"
      _StyleDefs(72)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(73)  =   "Named:id=36:Selected"
      _StyleDefs(74)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(75)  =   "Named:id=37:Caption"
      _StyleDefs(76)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(77)  =   "Named:id=38:HighlightRow"
      _StyleDefs(78)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(79)  =   "Named:id=39:EvenRow"
      _StyleDefs(80)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(81)  =   "Named:id=40:OddRow"
      _StyleDefs(82)  =   ":id=40,.parent=33"
      _StyleDefs(83)  =   "Named:id=41:RecordSelector"
      _StyleDefs(84)  =   ":id=41,.parent=34"
      _StyleDefs(85)  =   "Named:id=42:FilterBar"
      _StyleDefs(86)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1320
      OleObjectBlob   =   "frmMstSalaryStandard.frx":7A4C
      TabIndex        =   20
      Top             =   780
      Width           =   1695
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_division 
      Height          =   375
      Left            =   1320
      OleObjectBlob   =   "frmMstSalaryStandard.frx":9A0A
      TabIndex        =   23
      Top             =   1170
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   25
      Top             =   5340
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "UMUM"
      TabPicture(0)   =   "frmMstSalaryStandard.frx":B9C9
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label27"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_pengali"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_ket"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblGaji"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TDBCombo_pph"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TDBCombo_ptkp"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "DTPicker_salary"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TDBCombo_jstk"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_jstk_name"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_ptkp_name"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_pph21_name"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_main_salary"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame3"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_main_salary_sunday"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt_bulanan"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cboGaji"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "TUNJANGAN"
      TabPicture(1)   =   "frmMstSalaryStandard.frx":B9E5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_incentive_allowance"
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(2)=   "txt_pph21_allowance"
      Tab(1).Control(3)=   "Frame1"
      Tab(1).Control(4)=   "txt_shift2_allowance"
      Tab(1).Control(5)=   "txt_meal_allowance"
      Tab(1).Control(6)=   "txt_shift3_allowance"
      Tab(1).Control(7)=   "txt_title_allowance"
      Tab(1).Control(8)=   "txt_presence_allowance"
      Tab(1).Control(9)=   "txt_performance_allowance"
      Tab(1).Control(10)=   "txt_family_allowance"
      Tab(1).Control(11)=   "Frame4"
      Tab(1).Control(12)=   "Frame5"
      Tab(1).Control(13)=   "Frame2"
      Tab(1).Control(14)=   "Frame7"
      Tab(1).Control(15)=   "Label18"
      Tab(1).Control(16)=   "Label16"
      Tab(1).Control(17)=   "Label12"
      Tab(1).Control(18)=   "Label11"
      Tab(1).Control(19)=   "Label10"
      Tab(1).Control(20)=   "Label8"
      Tab(1).Control(21)=   "Label9"
      Tab(1).Control(22)=   "Label4"
      Tab(1).Control(23)=   "Label5"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "POTONGAN"
      TabPicture(2)   =   "frmMstSalaryStandard.frx":BA01
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label7"
      Tab(2).Control(1)=   "Label13"
      Tab(2).Control(2)=   "Label15"
      Tab(2).Control(3)=   "txt_late_minutes"
      Tab(2).Control(4)=   "txt_late_amount"
      Tab(2).ControlCount=   5
      Begin VB.ComboBox cboGaji 
         Height          =   315
         ItemData        =   "frmMstSalaryStandard.frx":BA1D
         Left            =   8880
         List            =   "frmMstSalaryStandard.frx":BA27
         TabIndex        =   81
         Top             =   2250
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.TextBox txt_incentive_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -71790
         MaxLength       =   10
         TabIndex        =   8
         Top             =   1680
         Width           =   2325
      End
      Begin VB.Frame Frame6 
         Height          =   525
         Left            =   -66360
         TabIndex        =   72
         Top             =   570
         Width           =   2355
         Begin VB.OptionButton opt_monthly_family 
            Caption         =   "BULANAN"
            Height          =   225
            Left            =   1110
            TabIndex        =   73
            Top             =   210
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton opt_daily_family 
            Caption         =   "HARIAN"
            Height          =   225
            Left            =   90
            TabIndex        =   74
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.TextBox txt_pph21_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -66360
         MaxLength       =   10
         TabIndex        =   15
         Top             =   3000
         Width           =   2325
      End
      Begin VB.TextBox txt_bulanan 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   7500
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   62
         Top             =   2940
         Width           =   2865
      End
      Begin VB.TextBox txt_main_salary_sunday 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Frame Frame3 
         Height          =   525
         Left            =   4680
         TabIndex        =   59
         Top             =   2070
         Width           =   2355
         Begin VB.OptionButton opt_daily_basic 
            Caption         =   "HARIAN"
            Height          =   225
            Left            =   90
            TabIndex        =   61
            Top             =   210
            Width           =   945
         End
         Begin VB.OptionButton opt_monthly_basic 
            Caption         =   "BULANAN"
            Height          =   225
            Left            =   1110
            TabIndex        =   60
            Top             =   210
            Width           =   1125
         End
      End
      Begin VB.Frame Frame1 
         Height          =   525
         Left            =   -71790
         TabIndex        =   52
         Top             =   330
         Width           =   2355
         Begin VB.OptionButton opt_monthly_presence 
            Caption         =   "BULANAN"
            Height          =   225
            Left            =   1110
            TabIndex        =   54
            Top             =   210
            Width           =   1125
         End
         Begin VB.OptionButton opt_daily_presence 
            Caption         =   "HARIAN"
            Height          =   225
            Left            =   90
            TabIndex        =   53
            Top             =   210
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.TextBox txt_late_amount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -69660
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txt_late_minutes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -69660
         MaxLength       =   10
         TabIndex        =   16
         Top             =   1320
         Width           =   1575
      End
      Begin VB.TextBox txt_shift2_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -66360
         MaxLength       =   10
         TabIndex        =   13
         Top             =   2280
         Width           =   2325
      End
      Begin VB.TextBox txt_meal_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -66360
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1920
         Width           =   2325
      End
      Begin VB.TextBox txt_shift3_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -66360
         MaxLength       =   10
         TabIndex        =   14
         Top             =   2640
         Width           =   2325
      End
      Begin VB.TextBox txt_title_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -71790
         MaxLength       =   10
         TabIndex        =   9
         Top             =   2490
         Width           =   2325
      End
      Begin VB.TextBox txt_presence_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -71790
         MaxLength       =   10
         TabIndex        =   7
         Top             =   870
         Width           =   2325
      End
      Begin VB.TextBox txt_performance_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -71790
         MaxLength       =   10
         TabIndex        =   10
         Top             =   3300
         Width           =   2325
      End
      Begin VB.TextBox txt_family_allowance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   -66360
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1110
         Width           =   2325
      End
      Begin VB.TextBox txt_main_salary 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0;(0)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   4680
         MaxLength       =   10
         TabIndex        =   4
         Top             =   2700
         Width           =   1665
      End
      Begin VB.TextBox txt_pph21_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   35
         Top             =   900
         Width           =   3915
      End
      Begin VB.TextBox txt_ptkp_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   34
         Top             =   1320
         Width           =   3915
      End
      Begin VB.TextBox txt_jstk_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   6480
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   33
         Top             =   1740
         Width           =   3915
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_jstk 
         Height          =   375
         Left            =   4680
         OleObjectBlob   =   "frmMstSalaryStandard.frx":BA3E
         TabIndex        =   3
         Top             =   1740
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPicker_salary 
         Height          =   315
         Left            =   4680
         TabIndex        =   0
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   90505219
         CurrentDate     =   40987
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_ptkp 
         Height          =   375
         Left            =   4680
         OleObjectBlob   =   "frmMstSalaryStandard.frx":DA01
         TabIndex        =   2
         Top             =   1320
         Width           =   1695
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_pph 
         Height          =   375
         Left            =   4680
         OleObjectBlob   =   "frmMstSalaryStandard.frx":F9B4
         TabIndex        =   1
         Top             =   900
         Width           =   1695
      End
      Begin VB.Frame Frame4 
         Height          =   525
         Left            =   -71790
         TabIndex        =   66
         Top             =   1950
         Width           =   2355
         Begin VB.OptionButton opt_monthly_title 
            Caption         =   "BULANAN"
            Height          =   225
            Left            =   1110
            TabIndex        =   67
            Top             =   210
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton opt_daily_title 
            Caption         =   "HARIAN"
            Height          =   225
            Left            =   90
            TabIndex        =   68
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame5 
         Height          =   525
         Left            =   -71790
         TabIndex        =   69
         Top             =   2760
         Width           =   2355
         Begin VB.OptionButton opt_monthly_performance 
            Caption         =   "BULANAN"
            Height          =   225
            Left            =   1110
            TabIndex        =   70
            Top             =   210
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton opt_daily_performance 
            Caption         =   "HARIAN"
            Height          =   225
            Left            =   90
            TabIndex        =   71
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   -66360
         TabIndex        =   55
         Top             =   1380
         Width           =   2355
         Begin VB.OptionButton opt_daily_meal 
            Caption         =   "HARIAN"
            Height          =   225
            Left            =   90
            TabIndex        =   57
            Top             =   210
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton opt_monthly_meal 
            Caption         =   "BULANAN"
            Height          =   225
            Left            =   1110
            TabIndex        =   56
            Top             =   210
            Width           =   1125
         End
      End
      Begin VB.Frame Frame7 
         Height          =   525
         Left            =   -71790
         TabIndex        =   75
         Top             =   1140
         Width           =   2355
         Begin VB.OptionButton opt_monthly_insentif 
            Caption         =   "BULANAN"
            Height          =   225
            Left            =   1110
            TabIndex        =   76
            Top             =   210
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton opt_daily_insentif 
            Caption         =   "HARIAN"
            Height          =   225
            Left            =   90
            TabIndex        =   77
            Top             =   210
            Width           =   1065
         End
      End
      Begin VB.Label lblGaji 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "GAJI DITERIMA"
         Height          =   195
         Left            =   7485
         TabIndex        =   80
         Top             =   2280
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "INSENTIF KERJA"
         Height          =   195
         Left            =   -73305
         TabIndex        =   78
         Top             =   1590
         Width           =   1290
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PPh 21"
         Height          =   195
         Left            =   -67125
         TabIndex        =   65
         Top             =   3060
         Width           =   525
      End
      Begin VB.Label lbl_ket 
         Caption         =   "Label18"
         Height          =   195
         Left            =   7500
         TabIndex        =   64
         Top             =   2700
         Width           =   1755
      End
      Begin VB.Label lbl_pengali 
         Alignment       =   2  'Center
         Caption         =   "Label16"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6450
         TabIndex        =   63
         Top             =   2730
         Width           =   405
      End
      Begin VB.Label Label15 
         Caption         =   "Menit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -68040
         TabIndex        =   58
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "POTONGAN TERLAMBAT"
         Height          =   195
         Left            =   -71730
         TabIndex        =   50
         Top             =   1740
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TOLERANSI TERLAMBAT"
         Height          =   195
         Left            =   -71730
         TabIndex        =   49
         Top             =   1380
         Width           =   1935
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SHIFT 2"
         Height          =   195
         Left            =   -68400
         TabIndex        =   48
         Top             =   2340
         Width           =   1740
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TRNS/MAKAN"
         Height          =   195
         Left            =   -68535
         TabIndex        =   47
         Top             =   1800
         Width           =   1905
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SHIFT 3"
         Height          =   195
         Left            =   -68400
         TabIndex        =   46
         Top             =   2700
         Width           =   1740
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "JABATAN"
         Height          =   195
         Left            =   -74100
         TabIndex        =   45
         Top             =   2370
         Width           =   2070
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KEHADIRAN"
         Height          =   195
         Left            =   -73905
         TabIndex        =   44
         Top             =   750
         Width           =   1875
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PRESTASI"
         Height          =   195
         Left            =   -72825
         TabIndex        =   43
         Top             =   3180
         Width           =   795
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "KELUARGA"
         Height          =   195
         Left            =   -67500
         TabIndex        =   42
         Top             =   1020
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "GAJI POKOK (MINGGU)"
         Height          =   195
         Left            =   2475
         TabIndex        =   41
         Top             =   3180
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "GAJI POKOK*"
         Height          =   195
         Left            =   3225
         TabIndex        =   40
         Top             =   2760
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TIPE PPh21*"
         Height          =   195
         Left            =   3300
         TabIndex        =   39
         Top             =   930
         Width           =   945
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TIPE PTKP*"
         Height          =   195
         Left            =   3360
         TabIndex        =   38
         Top             =   1350
         Width           =   885
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TANGGAL"
         Height          =   195
         Left            =   3480
         TabIndex        =   37
         Top             =   540
         Width           =   765
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "TIPE JAMSOSTEK*"
         Height          =   195
         Left            =   2820
         TabIndex        =   36
         Top             =   1770
         Width           =   1425
      End
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid_Salary 
      Height          =   3615
      Left            =   120
      TabIndex        =   51
      Top             =   5370
      Width           =   12195
      _ExtentX        =   21511
      _ExtentY        =   6376
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "DATE"
      Columns(0).DataField=   "salary_date"
      Columns(0).NumberFormat=   "yyyy-MM-dd"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "EMPLOYEE CODE"
      Columns(1).DataField=   "employee_code"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "EMP. CODE"
      Columns(2).DataField=   "nik"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "BASIC SALARY"
      Columns(3).DataField=   "basic_salary"
      Columns(3).NumberFormat=   "Standard"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "BASIC SUNDAY"
      Columns(4).DataField=   "basic_salary_sunday"
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "PRESENCE ALLOW."
      Columns(5).DataField=   "presence_allowance"
      Columns(5).NumberFormat=   "Standard"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "INSENTIF KERJA"
      Columns(6).DataField=   "incentive_allowance"
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "TITLE ALLOW."
      Columns(7).DataField=   "title_allowance"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "PERFORMANCE ALLOW."
      Columns(8).DataField=   "performance_allowance"
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "FAMILY ALLOW."
      Columns(9).DataField=   "family_allowance"
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "TRNS/MEAL ALLOW."
      Columns(10).DataField=   "meal_allowance"
      Columns(10).NumberFormat=   "Standard"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "SHIFT 2 ALLOW."
      Columns(11).DataField=   "shift2_allowance"
      Columns(11).NumberFormat=   "Standard"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "SHIFT 3 ALLOW."
      Columns(12).DataField=   "shift3_allowance"
      Columns(12).NumberFormat=   "Standard"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "PPh21 TYPE"
      Columns(13).DataField=   "pph21_name"
      Columns(13).NumberFormat=   "Standard"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "PTKP TYPE"
      Columns(14).DataField=   "ptkp_name"
      Columns(14).NumberFormat=   "Standard"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "JAMSOSTEK TYPE"
      Columns(15).DataField=   "jamsostek_name"
      Columns(15).NumberFormat=   "Standard"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).Size  =   2
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=1217"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1138"
      Splits(0)._ColumnProps(9)=   "Column(1).AllowSizing=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(12)=   "Column(1).AllowFocus=0"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=3360"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=3281"
      Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(24)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=514"
      Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(37)=   "Column(6)._ColStyle=514"
      Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(39)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=514"
      Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(44)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=514"
      Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(49)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(50)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(52)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(53)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(54)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(55)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(57)=   "Column(10)._ColStyle=514"
      Splits(0)._ColumnProps(58)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(59)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(60)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(61)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(62)=   "Column(11)._ColStyle=514"
      Splits(0)._ColumnProps(63)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(64)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(65)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(66)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(67)=   "Column(12)._ColStyle=514"
      Splits(0)._ColumnProps(68)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(69)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(70)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(72)=   "Column(13)._ColStyle=514"
      Splits(0)._ColumnProps(73)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(74)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(75)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(76)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(77)=   "Column(14)._ColStyle=514"
      Splits(0)._ColumnProps(78)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(79)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(80)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(81)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(82)=   "Column(15)._ColStyle=514"
      Splits(0)._ColumnProps(83)=   "Column(15).Order=16"
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
      Caption         =   "DAFTAR GAJI"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=126,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=123,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=124,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=125,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=82,.parent=13,.alignment=1"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=79,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=80,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=81,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=86,.parent=13,.alignment=1"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
      _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=90,.parent=13,.alignment=1"
      _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=87,.parent=14"
      _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=88,.parent=15"
      _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=89,.parent=17"
      _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=94,.parent=13,.alignment=1"
      _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=91,.parent=14"
      _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=92,.parent=15"
      _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=93,.parent=17"
      _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=98,.parent=13,.alignment=1"
      _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=95,.parent=14"
      _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=96,.parent=15"
      _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=97,.parent=17"
      _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=102,.parent=13,.alignment=1"
      _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=99,.parent=14"
      _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=100,.parent=15"
      _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=101,.parent=17"
      _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=106,.parent=13,.alignment=1"
      _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=103,.parent=14"
      _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=104,.parent=15"
      _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=105,.parent=17"
      _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=110,.parent=13,.alignment=1"
      _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=107,.parent=14"
      _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=108,.parent=15"
      _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=109,.parent=17"
      _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=114,.parent=13,.alignment=1"
      _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=111,.parent=14"
      _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=112,.parent=15"
      _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=113,.parent=17"
      _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=118,.parent=13,.alignment=1"
      _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=115,.parent=14"
      _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=116,.parent=15"
      _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=117,.parent=17"
      _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=122,.parent=13,.alignment=1"
      _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=119,.parent=14"
      _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=120,.parent=15"
      _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=121,.parent=17"
      _StyleDefs(98)  =   "Named:id=33:Normal"
      _StyleDefs(99)  =   ":id=33,.parent=0"
      _StyleDefs(100) =   "Named:id=34:Heading"
      _StyleDefs(101) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(102) =   ":id=34,.wraptext=-1"
      _StyleDefs(103) =   "Named:id=35:Footing"
      _StyleDefs(104) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(105) =   "Named:id=36:Selected"
      _StyleDefs(106) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(107) =   "Named:id=37:Caption"
      _StyleDefs(108) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(109) =   "Named:id=38:HighlightRow"
      _StyleDefs(110) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(111) =   "Named:id=39:EvenRow"
      _StyleDefs(112) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(113) =   "Named:id=40:OddRow"
      _StyleDefs(114) =   ":id=40,.parent=33"
      _StyleDefs(115) =   "Named:id=41:RecordSelector"
      _StyleDefs(116) =   ":id=41,.parent=34"
      _StyleDefs(117) =   "Named:id=42:FilterBar"
      _StyleDefs(118) =   ":id=42,.parent=33"
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "DIVISI"
      Height          =   195
      Left            =   750
      TabIndex        =   24
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   1110
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER GAJI"
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
      Left            =   270
      TabIndex        =   18
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frmMstSalaryStandard.frx":11966
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12450
   End
End
Attribute VB_Name = "frm_mst_salary_standard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCompany As New ADODB.Recordset
Dim rsDiv As New ADODB.Recordset
Dim rsEmployee As New ADODB.Recordset
Dim rsSalary As New ADODB.Recordset
Dim rspph As New ADODB.Recordset
Dim rsPTKP As New ADODB.Recordset
Dim rsJSTK As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SQL As String
Dim v_main As Double
Dim v_salary_date As String
Dim vCompany As String
Dim vDivision As String


Private Function check_validate_exist_new() As Boolean
Dim str_sql As String
    check_validate_exist_new = False
    
    If rsSalary.State Then rsSalary.Close
    str_sql = "select count(employee_code) as rec_count from m_salary_standard " & _
                "where employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
                    "AND date(salary_date) = '" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'"
    rsSalary.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rsSalary.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        rsSalary.Close
        Exit Function
    End If
    
    rsSalary.Close
End Function

Private Sub check_invalid()
    MsgBox "Data Sudah Ada...", vbCritical, headerMSG
    txt_main_salary.Text = ""
    If DTPicker_salary.Enabled = True Then DTPicker_salary.SetFocus
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If Not Format(DTPicker_salary, "yyyy-MM-dd") = Format(rsSalary.Fields("salary_date").Value, "yyyy-MM-dd") And _
    check_validate_exist_new Then
        check_validate_exist_edit = True
        Exit Function
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True

    'validasi salary
    If Trim(TDBCombo_pph.Text) = "" Then
        MsgBox "Tipe PPh Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        TDBCombo_pph.SetFocus
        check_validate_new = False
        Exit Function
    End If

    'validasi salary
    If Trim(TDBCombo_ptkp.Text) = "" Then
        MsgBox "Tipe PTKP Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        TDBCombo_ptkp.SetFocus
        check_validate_new = False
        Exit Function
    End If
    
    'validasi salary
    If Trim(TDBCombo_jstk.Text) = "" Then
        MsgBox "Tipe Jamsostek Masih Kosong...", vbOKOnly + vbInformation, headerMSG
        TDBCombo_jstk.SetFocus
        check_validate_new = False
        Exit Function
    End If
End Function

Private Sub load_data()
    timer1.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    int_mode = 0
    Call load_mode
End Sub

Private Sub cmdDelete_Click()
Dim i As Integer
Dim vSalaryDate As String
    SQL = "SELECT salary_date FROM m_salary_standard " & _
            "WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
            "ORDER BY salary_date DESC LIMIT 1"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vSalaryDate = Format(rs!salary_date, "yyyy-MM-dd")
    End If
    rs.Close
    
    If vSalaryDate = Format(TDBGrid_Salary.Columns("salary_date").Value, "yyyy-MM-dd") Then
        If Not (TDBGrid_Salary.ApproxCount > 0 And TDBGrid_Salary.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Salary.Columns("salary_date").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_salary_standard " & _
                    "where employee_code = '" & TDBGrid_Salary.Columns("employee_code").Value & "' " & _
                        "AND date(salary_date) = '" & Format(TDBGrid_Salary.Columns("salary_date").Value, "yyyy-MM-dd") & "'"
        
        '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
        SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CnG.CommitTrans
        
        Call load_data_salary
        int_mode = 0
        Call load_mode
    Else
        MsgBox "Data Tidak Bisa Dihapus Karena Bukan Data Upah Terakhir...", vbExclamation, headerMSG
        Exit Sub
    End If
    
End Sub

Public Sub set_edit_data()
Dim v_flag_basic As Integer
Dim v_flag_presence As Integer
Dim v_flag_incentive As Integer
Dim v_flag_title As Integer
Dim v_flag_performance As Integer
Dim v_flag_family As Integer
Dim v_flag_meal As Integer
Dim jhk As Integer
    
    vSetData = 1
    
    If Not (TDBGrid_Salary.ApproxCount > 0 And TDBGrid_Salary.Bookmark > 0) Then
        MsgBox "Data Gaji Belum Ada...", vbInformation, headerMSG
        vSetData = 0
        Exit Sub
    End If
        
    With rsSalary
        DTPicker_salary.Value = Format(.Fields("salary_date").Value, "yyyy-MM-dd")
        TDBCombo_pph.Text = .Fields("pph21_type").Value
        txt_pph21_name.Text = .Fields("pph21_name").Value
        TDBCombo_ptkp.Text = .Fields("ptkp_type").Value
        txt_ptkp_name.Text = .Fields("ptkp_name").Value
        TDBCombo_jstk.Text = .Fields("jstk_type").Value
        txt_jstk_name.Text = .Fields("jamsostek_name").Value
        
        v_flag_basic = IIf(IsNull(.Fields("flag_basic").Value), 0, .Fields("flag_basic").Value)
        opt_daily_basic.Value = IIf(v_flag_basic = 0, True, False)
        opt_monthly_basic.Value = IIf(v_flag_basic = 0, False, True)
        txt_main_salary.Text = FormatNumber(IIf(IsNull(.Fields("basic_salary").Value), 0, .Fields("basic_salary").Value))
        txt_main_salary_sunday.Text = FormatNumber(IIf(IsNull(.Fields("basic_salary_sunday").Value), 0, .Fields("basic_salary_sunday").Value))
        
        v_flag_presence = IIf(IsNull(.Fields("flag_presence").Value), 0, .Fields("flag_presence").Value)
        opt_daily_presence.Value = IIf(v_flag_presence = 0, True, False)
        opt_monthly_presence.Value = IIf(v_flag_presence = 0, False, True)
        txt_presence_allowance.Text = FormatNumber(IIf(IsNull(.Fields("presence_allowance").Value), 0, .Fields("presence_allowance").Value))
        
        v_flag_incentive = IIf(IsNull(.Fields("flag_incentive").Value), 0, .Fields("flag_incentive").Value)
        opt_daily_insentif.Value = IIf(v_flag_incentive = 0, True, False)
        opt_monthly_insentif.Value = IIf(v_flag_incentive = 0, False, True)
        txt_incentive_allowance.Text = FormatNumber(IIf(IsNull(.Fields("incentive_allowance").Value), 0, .Fields("incentive_allowance").Value))
        
        v_flag_title = IIf(IsNull(.Fields("flag_title").Value), 0, .Fields("flag_title").Value)
        opt_daily_title.Value = IIf(v_flag_title = 0, True, False)
        opt_monthly_title.Value = IIf(v_flag_title = 0, False, True)
        txt_title_allowance.Text = FormatNumber(IIf(IsNull(.Fields("title_allowance").Value), 0, .Fields("title_allowance").Value))
        
        v_flag_performance = IIf(IsNull(.Fields("flag_performance").Value), 0, .Fields("flag_performance").Value)
        opt_daily_performance.Value = IIf(v_flag_performance = 0, True, False)
        opt_monthly_performance.Value = IIf(v_flag_performance = 0, False, True)
        txt_performance_allowance.Text = FormatNumber(IIf(IsNull(.Fields("performance_allowance").Value), 0, .Fields("performance_allowance").Value))
        
        v_flag_family = IIf(IsNull(.Fields("flag_family").Value), 0, .Fields("flag_family").Value)
        opt_daily_family.Value = IIf(v_flag_family = 0, True, False)
        opt_monthly_family.Value = IIf(v_flag_family = 0, False, True)
        txt_family_allowance.Text = FormatNumber(IIf(IsNull(.Fields("family_allowance").Value), 0, .Fields("family_allowance").Value))
        
        v_flag_meal = IIf(IsNull(.Fields("flag_meal").Value), 0, .Fields("flag_meal").Value)
        opt_daily_meal.Value = IIf(v_flag_meal = 0, 1, 0)
        opt_monthly_meal.Value = IIf(v_flag_meal = 0, 0, 1)
        txt_meal_allowance.Text = FormatNumber(IIf(IsNull(.Fields("meal_allowance").Value), 0, .Fields("meal_allowance").Value))
        
        txt_shift2_allowance.Text = FormatNumber(IIf(IsNull(.Fields("shift2_allowance").Value), 0, .Fields("shift2_allowance").Value))
        txt_shift3_allowance.Text = FormatNumber(IIf(IsNull(.Fields("shift3_allowance").Value), 0, .Fields("shift3_allowance").Value))
        txt_pph21_allowance.Text = FormatNumber(IIf(IsNull(.Fields("pph21_allowance").Value), 0, .Fields("pph21_allowance").Value))
        
        txt_late_minutes.Text = IIf(IsNull(.Fields("late_time_tolerance").Value), 0, .Fields("late_time_tolerance").Value)
        txt_late_amount.Text = FormatNumber(IIf(IsNull(.Fields("late_amount").Value), 0, .Fields("late_amount").Value))
        
        cboGaji.ListIndex = IIf(IsNull(.Fields("flag_gaji").Value), 0, .Fields("flag_gaji").Value)
    End With

    v_main = txt_main_salary.Text
    v_salary_date = Format(DTPicker_salary.Value, "yyyy-MM-dd")
    
    SQL = "SELECT company_code, division_code FROM m_employee WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vCompany = rs!COMPANY_CODE
        vDivision = rs!division_code
    End If
    rs.Close
    
    SQL = "SELECT jhk_value FROM m_pref_jhk WHERE company_code = '" & vCompany & "' AND division_code = '" & vDivision & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        jhk = IIf(IsNull(rs!jhk_value), Val(Format(getEndDay(month(DTPicker_salary.Value), year(DTPicker_salary.Value)), "dd")), rs!jhk_value)
    Else
        jhk = Val(Format(getEndDay(month(DTPicker_salary.Value), year(DTPicker_salary.Value)), "dd"))
    End If
    rs.Close
    
    If opt_daily_basic.Value Then
        txt_bulanan.Text = FormatNumber(Round(Val(DropAllComma(txt_main_salary.Text) * jhk)))
    Else
        txt_bulanan.Text = FormatNumber(Round(Val(DropAllComma(txt_main_salary.Text) / jhk)))
    End If
End Sub

Private Sub cmdEdit_Click()
Dim vSalaryDate As String
    SQL = "SELECT salary_date FROM m_salary_standard " & _
            "WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
            "ORDER BY salary_date DESC LIMIT 1"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vSalaryDate = Format(rs!salary_date, "yyyy-MM-dd")
    End If
    rs.Close
    
    TDBGrid_Emp.Enabled = False
    
    If vSalaryDate = Format(TDBGrid_Salary.Columns("salary_date").Value, "yyyy-MM-dd") Then
        int_mode = 2
        Call load_mode
        
        SSTab1.Tab = 0
    Else
        MsgBox "Data Tidak Bisa Diubah Karena Bukan Data Upah Terakhir...", vbExclamation, headerMSG
        Exit Sub
    End If
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
    frm_import_salary_standard.Show
End Sub

Private Sub cmdNew_Click()
    int_mode = 1
    Call load_mode
    
    SSTab1.Tab = 0
    DTPicker_salary.Value = Now
    opt_monthly_basic.Value = True
    cboGaji.ListIndex = 1
    opt_daily_presence.Value = True
    opt_monthly_insentif.Value = True
    opt_monthly_title.Value = True
    opt_monthly_performance.Value = True
    opt_monthly_family.Value = True
    opt_monthly_meal.Value = True
End Sub

Private Sub insert_new_data()
On Error GoTo Err
    CnG.BeginTrans

    SQL = "INSERT INTO m_salary_standard (employee_code, salary_date," & _
            "flag_basic, basic_salary, basic_salary_sunday, flag_presence,presence_allowance," & _
            "flag_incentive, incentive_allowance, flag_title, title_allowance," & _
            "flag_performance, performance_allowance, flag_family, family_allowance," & _
            "flag_meal,meal_allowance,shift2_allowance,shift3_allowance,pph21_allowance," & _
            "late_time_tolerance,late_amount,pph21_type,ptkp_type,jstk_type,flag_gaji,entry_date,entry_user) " & _
            "VALUES " & _
            "('" & TDBGrid_Emp.Columns("employee_code").Value & "','" & Format(DTPicker_salary.Value, "yyyy-MM-dd") & "'," & _
            "'" & IIf(opt_daily_basic.Value, 0, 1) & "'," & Val(DropAllComma(txt_main_salary.Text)) & "," & Val(DropAllComma(txt_main_salary.Text)) & "," & _
            "'" & IIf(opt_daily_presence.Value, 0, 1) & "'," & Val(DropAllComma(txt_presence_allowance.Text)) & "," & _
            "'" & IIf(opt_daily_insentif.Value, 0, 1) & "'," & Val(DropAllComma(txt_incentive_allowance.Text)) & "," & _
            "'" & IIf(opt_daily_title.Value, 0, 1) & "'," & Val(DropAllComma(txt_title_allowance.Text)) & "," & _
            "'" & IIf(opt_daily_performance.Value, 0, 1) & "'," & Val(DropAllComma(txt_performance_allowance.Text)) & "," & _
            "'" & IIf(opt_daily_family.Value, 0, 1) & "'," & Val(DropAllComma(txt_family_allowance.Text)) & "," & _
            "'" & IIf(opt_daily_meal.Value, 0, 1) & "'," & Val(DropAllComma(txt_meal_allowance.Text)) & "," & _
            "" & Val(DropAllComma(txt_shift2_allowance.Text)) & ",'" & Val(DropAllComma(txt_shift3_allowance.Text)) & "','" & Val(DropAllComma(txt_pph21_allowance.Text)) & "'," & _
            "" & Val(txt_late_minutes.Text) & ", " & Val(DropAllComma(txt_late_amount.Text)) & "," & _
            "'" & TDBCombo_pph.Text & "','" & TDBCombo_ptkp.Text & "','" & TDBCombo_jstk.Text & "'," & _
            "'" & cboGaji.ListIndex & "',now(),'" & LOGIN_NAME & "')"
    CnG.Execute SQL

    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err
    CnG.BeginTrans

    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    If v_main <> txt_main_salary.Text Then
            SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company.Text & "'"
            CnG.Execute SQL
    End If
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    SQL = "UPDATE m_salary_standard SET " & _
            "flag_basic = '" & IIf(opt_daily_basic.Value, 0, 1) & "'," & _
            "basic_salary = '" & Val(DropAllComma(txt_main_salary.Text)) & "'," & _
            "basic_salary_sunday = '" & Val(DropAllComma(txt_main_salary_sunday.Text)) & "'," & _
            "flag_presence = '" & IIf(opt_daily_presence.Value, 0, 1) & "'," & _
            "presence_allowance = '" & Val(DropAllComma(txt_presence_allowance.Text)) & "'," & _
            "flag_incentive = '" & IIf(opt_daily_insentif.Value, 0, 1) & "'," & _
            "incentive_allowance = '" & Val(DropAllComma(txt_incentive_allowance.Text)) & "'," & _
            "flag_title = '" & IIf(opt_daily_title.Value, 0, 1) & "'," & _
            "title_allowance = '" & Val(DropAllComma(txt_title_allowance.Text)) & "'," & _
            "flag_performance = '" & IIf(opt_daily_performance.Value, 0, 1) & "'," & _
            "performance_allowance = '" & Val(DropAllComma(txt_performance_allowance.Text)) & "'," & _
            "flag_family = '" & IIf(opt_daily_family.Value, 0, 1) & "'," & _
            "family_allowance = '" & Val(DropAllComma(txt_family_allowance.Text)) & "'," & _
            "flag_meal = '" & IIf(opt_daily_meal.Value, 0, 1) & "'," & _
            "meal_allowance = '" & Val(DropAllComma(txt_meal_allowance.Text)) & "'," & _
            "shift2_allowance = '" & Val(DropAllComma(txt_shift2_allowance.Text)) & "'," & _
            "shift3_allowance = '" & Val(DropAllComma(txt_shift3_allowance.Text)) & "'," & _
            "pph21_allowance = '" & Val(DropAllComma(txt_pph21_allowance.Text)) & "'," & _
            "late_time_tolerance = '" & txt_late_minutes.Text & "'," & _
            "late_amount = '" & Val(DropAllComma(txt_late_amount.Text)) & "'," & _
            "pph21_type = '" & TDBCombo_pph.Text & "', " & _
            "ptkp_type = '" & TDBCombo_ptkp.Text & "', " & _
            "jstk_type = '" & TDBCombo_jstk.Text & "',flag_gaji = '" & cboGaji.ListIndex & "' " & _
            "WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' AND date(salary_date) = '" & v_salary_date & "'"
    CnG.Execute SQL
    
    CnG.CommitTrans

    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdSave_Click()
Dim clsFunc As New clsFunction

    If int_mode = 1 Then
        If Not check_validate_new Then Exit Sub
        If check_validate_exist_new Then
            MsgBox "Data Tidak Valid...", vbInformation, headerMSG
            Exit Sub
        End If
        Call insert_new_data
        clsFunc.InsertLog ("Insert Salary Standard : " & TDBGrid_Emp.Columns("employee_code").Value)
    ElseIf int_mode = 2 Then
        If Not check_validate_new Then Exit Sub
        If check_validate_exist_edit Then
            Call check_invalid: Exit Sub
        End If
        Call edit_old_data
        clsFunc.InsertLog ("Edit Salary Standard : " & TDBGrid_Emp.Columns("employee_code").Value)
    End If
    
    TDBGrid_Emp.Enabled = True
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
Dim vNPWP_Methode As String
    DTPicker_salary.Value = Now
    DTPicker_salary.Enabled = True

    SQL = "SELECT npwp_method FROM m_employee WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "'"
    rs.Open SQL, CnG, adOpenForwardOnly
    
    If rs.RecordCount > 0 Then
        vNPWP_Methode = IIf(IsNull(rs!npwp_method), 0, rs!npwp_method)
        If vNPWP_Methode = 2 Then
            txt_pph21_allowance.Enabled = True
        Else
            txt_pph21_allowance.Enabled = False
            txt_pph21_allowance.Text = 0
        End If
    Else
        txt_pph21_allowance.Enabled = False
        txt_pph21_allowance.Text = 0
    End If
    rs.Close
    
    SQL = "SELECT distinct c.time_tolerance, c.late_value " & _
            "FROM m_salary_standard a JOIN td_shift b ON a.employee_code = b.employee_code " & _
                "JOIN m_shift_group c ON b.group_code = c.group_code " & _
            "WHERE a.employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    If rs.RecordCount > 0 Then
        txt_late_minutes.Text = rs.Fields(0).Value
        txt_late_amount.Text = FormatNumber(rs.Fields(1).Value)
    End If
    rs.Close
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        SSTab1.Visible = True
        TDBGrid_Salary.Enabled = False
        TDBGrid_Emp.Enabled = False
        Call set_new_data
        
        If DTPicker_salary.Enabled = True Then
            DTPicker_salary.SetFocus
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        SSTab1.Visible = False
        TDBGrid_Salary.Enabled = True
        TDBGrid_Emp.Enabled = True
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        DTPicker_salary.Enabled = False
        SSTab1.Visible = True
        TDBGrid_Salary.Enabled = False
        TDBGrid_Emp.Enabled = False
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
    Call load_data_company
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    timer1.Enabled = True
End Sub

Private Sub clear_filter_employee()
    For Each Col In TDBGrid_Emp.Columns
        Col.FilterText = ""
    Next Col
    rsEmployee.Filter = adFilterNone
End Sub

Private Sub clear_filter_salary()
    For Each Col In TDBGrid_Salary.Columns
        Col.FilterText = ""
    Next Col
    rsSalary.Filter = adFilterNone
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

Private Sub opt_daily_basic_Click()
Dim jhk As Integer

    SQL = "SELECT company_code, division_code FROM m_employee WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vCompany = rs!COMPANY_CODE
        vDivision = rs!division_code
    End If
    rs.Close

    Label2.Visible = True
    txt_main_salary_sunday.Visible = True
    
    SQL = "SELECT jhk_value FROM m_pref_jhk WHERE company_code = '" & vCompany & "' AND division_code = '" & vDivision & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        jhk = IIf(IsNull(rs!jhk_value), Val(Format(getEndDay(month(DTPicker_salary.Value), year(DTPicker_salary.Value)), "dd")), rs!jhk_value)
    Else
        jhk = Val(Format(getEndDay(month(DTPicker_salary.Value), year(DTPicker_salary.Value)), "dd"))
    End If
    rs.Close
    
    lbl_pengali.Caption = "x " & jhk
    lbl_ket.Caption = "Asumsi Gaji Per Bulan"
    
    txt_bulanan.Text = FormatNumber(Val(txt_main_salary.Text) * jhk)
    
    lblGaji.Visible = True
    cboGaji.Visible = True
    cboGaji.ListIndex = 0
End Sub

Private Sub opt_monthly_basic_Click()
Dim jhk As Integer

On Error GoTo Err

    SQL = "SELECT company_code, division_code FROM m_employee WHERE employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vCompany = rs!COMPANY_CODE
        vDivision = rs!division_code
    End If
    rs.Close

    Label2.Visible = False
    txt_main_salary_sunday.Visible = False
    txt_main_salary_sunday.Text = ""
    
    SQL = "SELECT jhk_value FROM m_pref_jhk WHERE company_code = '" & vCompany & "' AND division_code = '" & vDivision & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        jhk = IIf(IsNull(rs!jhk_value), Val(Format(getEndDay(month(DTPicker_salary.Value), year(DTPicker_salary.Value)), "dd")), rs!jhk_value)
    Else
        jhk = Val(Format(getEndDay(month(DTPicker_salary.Value), year(DTPicker_salary.Value)), "dd"))
    End If
    rs.Close
    
    lbl_pengali.Caption = "/ " & jhk
    lbl_ket.Caption = "Gaji Per Hari"
    
    txt_bulanan.Text = FormatNumber(Val(txt_main_salary) / jhk)
    
    lblGaji.Visible = False
    cboGaji.Visible = False
    cboGaji.ListIndex = 1
    Exit Sub

Err:
    txt_bulanan.Text = 0
End Sub

Private Sub TDBGrid_Emp_FilterChange()
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid_Emp.Columns
    i = TDBGrid_Emp.Col
    TDBGrid_Emp.HoldFields
    
    rsEmployee.Filter = getFilter()
    TDBGrid_Emp.Col = i
    TDBGrid_Emp.EditActive = True
    
    TDBGrid_Emp.SelStart = Len(TDBGrid_Emp.Columns(i).FilterText)
    If TDBGrid_Emp.ApproxCount < 1 Then
        Call clear_filter_employee
        TDBGrid_Emp.Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter_employee
End Sub

Private Sub TDBGrid_Salary_FilterChange()
On Error GoTo Err

Dim i As Integer

    Set Cols = TDBGrid_Salary.Columns
    i = TDBGrid_Salary.Col
    TDBGrid_Salary.HoldFields
    
    rsSalary.Filter = getFilter()
    TDBGrid_Salary.Col = i
    TDBGrid_Salary.EditActive = True
    
    TDBGrid_Salary.SelStart = Len(TDBGrid_Salary.Columns(i).FilterText)
    If TDBGrid_Salary.ApproxCount < 1 Then
        Call clear_filter_salary
        TDBGrid_Salary.Col = i
    End If
    
    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter_salary
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
        
        Call load_data_pph21
        Call load_data_ptkp
        Call load_data_pph21
        Call load_data_jstk
        Call load_data_division
        Call load_data_employee
    End If
End Sub

Private Sub tdbcombo_division_itemChange()
    If TDBCombo_division.ApproxCount > 0 Then
        TDBCombo_division.Text = TDBCombo_division.Columns("division_code").Value
        txt_division_name.Text = TDBCombo_division.Columns("division_name").Value
        
        Call load_data_employee
    End If
End Sub

Private Sub TDBCombo_pph_ItemChange()
    If TDBCombo_pph.ApproxCount > 0 Then
        TDBCombo_pph.Text = TDBCombo_pph.Columns("pph21_code").Value
        txt_pph21_name = TDBCombo_pph.Columns("pph21_name").Value
        
    End If
End Sub

Private Sub TDBCombo_ptkp_ItemChange()
    If TDBCombo_ptkp.ApproxCount > 0 Then
        TDBCombo_ptkp.Text = TDBCombo_ptkp.Columns("ptkp_code").Value
        txt_ptkp_name = TDBCombo_ptkp.Columns("ptkp_name").Value
        
    End If
End Sub

Private Sub TDBCombo_jstk_ItemChange()
    If TDBCombo_jstk.ApproxCount > 0 Then
        TDBCombo_jstk.Text = TDBCombo_jstk.Columns("jamsostek_code").Value
        txt_jstk_name = TDBCombo_jstk.Columns("jamsostek_name").Value
        
    End If
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data_division()
    If rsDiv.State Then rsDiv.Close
    SQL = "select * from m_division where company_code = '" & TDBCombo_company.Text & "' order by division_code"
    rsDiv.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_division.RowSource = rsDiv
End Sub

Public Sub load_data_employee()
    If rsEmployee.State Then rsEmployee.Close
    If LOGIN_LEVEL = 100 Then
        SQL = "SELECT a.*,c.division_name,d.title_name " & _
              "FROM m_employee a JOIN m_division c ON a.division_code = c.division_code " & _
                "JOIN m_title d ON a.title_code = d.title_code " & _
              "WHERE " & IIf(TDBCombo_division.Text <> "", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND a.division_code = '" & TDBCombo_division.Text & "'", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "'") & " " & _
                "AND flag_active <> 0 order by employee_name"
    Else
        SQL = "SELECT a.*,c.division_name,d.title_name " & _
              "FROM m_employee a JOIN m_division c ON a.division_code = c.division_code " & _
                "JOIN m_title d ON a.title_code = d.title_code " & _
              "WHERE " & IIf(TDBCombo_division.Text <> "", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "' AND a.division_code = '" & TDBCombo_division.Text & "'", "a.company_code = '" & TDBCombo_company.Columns("company_code").Value & "'") & " " & _
                "AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) AND flag_active <> 0 order by employee_name"
    End If

    rsEmployee.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid_Emp.DataSource = rsEmployee
End Sub

Private Sub load_data_salary()
'SQL = "SELECT b.*," & _
'            "(SELECT number FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) number," & _
'            "(SELECT main_salary FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) main_salary," & _
'            "(SELECT functional_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) functional_allowance," & _
'            "(SELECT staff_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) staff_allowance," & _
'            "(SELECT acting_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) acting_allowance," & _
'            "(SELECT skill_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) skill_allowance," & _
'            "(SELECT transport_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) transport_allowance," & _
'            "(SELECT presence_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) presence_allowance," & _
'            "(SELECT meal_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) meal_allowance," & _
'            "(SELECT phone_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) phone_allowance," & _
'            "(SELECT driver_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) driver_allowance," & _
'            "(SELECT special_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) special_allowance," & _
'            "(SELECT other_allowance FROM m_salary_standard WHERE employee_code = b.employee_code ORDER BY salary_date DESC LIMIT 1) other_allowance " & _
'        "FROM m_employee b WHERE b.flag_active <> 0 AND b.company_code = '" & TDBCombo_company.Text & "' AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) ORDER BY 1"
    
    If rsSalary.State Then rsSalary.Close
    SQL = "SELECT a.*,b.nik,c.pph21_name,d.ptkp_name,e.jamsostek_name " & _
          "FROM m_salary_standard a JOIN m_employee b ON a.employee_code = b.employee_code " & _
            "JOIN m_pph21 c ON a.pph21_type = c.pph21_code " & _
            "JOIN m_ptkp d ON a.ptkp_type = d.ptkp_code " & _
            "JOIN m_jamsostek e ON a.jstk_type = e.jamsostek_code " & _
          "WHERE a.employee_code = '" & TDBGrid_Emp.Columns("employee_code").Value & "' " & _
          "ORDER BY salary_date DESC"
    rsSalary.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Salary.DataSource = rsSalary
End Sub

Private Sub load_data_pph21()
    If rspph.State Then rspph.Close
    SQL = "select * from m_pph21 order by pph21_code"
    rspph.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_pph.RowSource = rspph
End Sub

Private Sub load_data_ptkp()
    If rsPTKP.State Then rsPTKP.Close
    SQL = "select * from m_ptkp order by ptkp_code"
    rsPTKP.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_ptkp.RowSource = rsPTKP
End Sub

Private Sub load_data_jstk()
    If rsJSTK.State Then rsJSTK.Close
    SQL = "select * from m_jamsostek order by jamsostek_code"
    rsJSTK.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_jstk.RowSource = rsJSTK
End Sub

Private Sub TDBGrid_Salary_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    If TDBGrid_Salary.Columns(ColIndex).Caption = "DATE" Then
        Value = Format(Value, "yyyy-mm-dd")
    End If
End Sub

Private Sub Timer1_Timer()
    timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub TDBGrid_Emp_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (TDBGrid_Emp.Row + 1) > 0 And (TDBGrid_Emp.Row + 1) <> LastRow Then
        'MsgBox "LETS..."
        Call load_data_salary
    End If
End Sub

Private Sub tdbcombo_division_Change()
    If TDBCombo_division.Text = "" Then txt_division_name.Text = ""
    
    Call load_data_employee
End Sub

Private Sub txt_main_salary_Change()
Dim vJHK As Double
    SQL = "SELECT jhk FROM m_company WHERE company_code = '" & TDBCombo_company.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vJHK = rs!jhk
    End If
    rs.Close
    
    If opt_daily_basic.Value Then
        txt_bulanan.Text = Round(Val(DropAllComma(txt_main_salary.Text)) * vJHK)
        txt_bulanan.Text = FormatNumber(DropAllComma(txt_bulanan.Text))
    Else
        txt_bulanan.Text = Round(Val(DropAllComma(txt_main_salary.Text)) / vJHK)
        txt_bulanan.Text = FormatNumber(DropAllComma(txt_bulanan.Text))
    End If
End Sub

Private Sub txt_main_salary_Validate(Cancel As Boolean)
    If Not Trim(txt_main_salary) = "" Then
        txt_main_salary = FormatNumber(DropAllComma(txt_main_salary))
    End If
End Sub

Private Sub txt_main_salary_sunday_Validate(Cancel As Boolean)
    If Not Trim(txt_main_salary_sunday) = "" Then
        txt_main_salary_sunday = FormatNumber(DropAllComma(txt_main_salary_sunday))
    End If
End Sub

Private Sub txt_presence_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_presence_allowance) = "" Then
        txt_presence_allowance = FormatNumber(DropAllComma(txt_presence_allowance))
    End If
End Sub

Private Sub txt_title_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_title_allowance) = "" Then
        txt_title_allowance = FormatNumber(DropAllComma(txt_title_allowance))
    End If
End Sub

Private Sub txt_performance_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_performance_allowance) = "" Then
        txt_performance_allowance = FormatNumber(DropAllComma(txt_performance_allowance))
    End If
End Sub

Private Sub txt_family_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_family_allowance) = "" Then
        txt_family_allowance = FormatNumber(DropAllComma(txt_family_allowance))
    End If
End Sub

Private Sub txt_meal_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_meal_allowance) = "" Then
        txt_meal_allowance = FormatNumber(DropAllComma(txt_meal_allowance))
    End If
End Sub

Private Sub txt_shift2_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_shift2_allowance) = "" Then
        txt_shift2_allowance = FormatNumber(DropAllComma(txt_shift2_allowance))
    End If
End Sub

Private Sub txt_shift3_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_shift3_allowance) = "" Then
        txt_shift3_allowance = FormatNumber(DropAllComma(txt_shift3_allowance))
    End If
End Sub

Private Sub txt_pph21_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_pph21_allowance) = "" Then
        txt_pph21_allowance = FormatNumber(DropAllComma(txt_pph21_allowance))
    End If
End Sub

Private Sub txt_late_amount_Validate(Cancel As Boolean)
    If Not Trim(txt_late_amount) = "" Then
        txt_late_amount = FormatNumber(DropAllComma(txt_late_amount))
    End If
End Sub

Private Sub txt_incentive_allowance_Validate(Cancel As Boolean)
    If Not Trim(txt_incentive_allowance) = "" Then
        txt_incentive_allowance = FormatNumber(DropAllComma(txt_incentive_allowance))
    End If
End Sub
