VERSION 5.00
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_pph21 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "POTONGAN PPH21 BULANAN"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14730
   Icon            =   "frm_trans_pph21.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9615
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_company_name 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Height          =   315
      Left            =   2850
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   14
      Top             =   1230
      Width           =   3855
   End
   Begin VB.Frame frmTombol 
      Caption         =   "Data Control Button"
      Height          =   1335
      Left            =   240
      TabIndex        =   6
      Top             =   8100
      Width           =   14325
      Begin prj_panji.vbButton cmdSave 
         Height          =   705
         Left            =   870
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
         MICON           =   "frm_trans_pph21.frx":058A
         PICN            =   "frm_trans_pph21.frx":05A6
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
         Left            =   1890
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
         MICON           =   "frm_trans_pph21.frx":1638
         PICN            =   "frm_trans_pph21.frx":1654
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
         Left            =   2910
         TabIndex        =   12
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
         MICON           =   "frm_trans_pph21.frx":26E6
         PICN            =   "frm_trans_pph21.frx":2702
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
         Left            =   11910
         TabIndex        =   13
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
         MICON           =   "frm_trans_pph21.frx":3794
         PICN            =   "frm_trans_pph21.frx":37B0
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
      TabIndex        =   5
      Top             =   3570
      Width           =   14325
      Begin VB.TextBox txt_pph21_value 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8070
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   990
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         Caption         =   "List Karyawan"
         Height          =   4065
         Left            =   4770
         TabIndex        =   8
         Top             =   180
         Width           =   2895
         Begin VB.ListBox list_employee 
            Height          =   3765
            Left            =   90
            TabIndex        =   9
            Top             =   210
            Width           =   2715
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "POT. PPH21"
         Height          =   195
         Left            =   8070
         TabIndex        =   2
         Top             =   720
         Width           =   930
      End
   End
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   300
      Left            =   1530
      TabIndex        =   1
      Top             =   810
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   529
      _Version        =   393216
      CustomFormat    =   "yyyy-MM"
      Format          =   94109699
      UpDown          =   -1  'True
      CurrentDate     =   39278
   End
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   6375
      Left            =   240
      TabIndex        =   7
      Top             =   1650
      Width           =   14325
      _ExtentX        =   25268
      _ExtentY        =   11245
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "KODE DIV."
      Columns(0).DataField=   "division_code"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "DIVISI"
      Columns(1).DataField=   "division_name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "PERIODE"
      Columns(2).DataField=   "month"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "KODE KARY"
      Columns(3).DataField=   "employee_code"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "NIK"
      Columns(4).DataField=   "nik"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "NAMA KARY."
      Columns(5).DataField=   "employee_name"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "POT. PPH21"
      Columns(6).DataField=   "pph21_value"
      Columns(6).NumberFormat=   "Standard"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5133"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5054"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2990"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2910"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(23)=   "Column(4).Width=2858"
      Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2778"
      Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=513"
      Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(33)=   "Column(5).Visible=0"
      Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(35)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(38)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(39)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(40)=   "Column(6).Order=7"
      Splits(1)._UserFlags=   0
      Splits(1).Size  =   2
      Splits(1).Size.vt=   2
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).ScrollBars=   3
      Splits(1).DividerColor=   13160660
      Splits(1).FilterBar=   -1  'True
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=7"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(1)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(1)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(1)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(7)=   "Column(1).Width=2725"
      Splits(1)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(9)=   "Column(1)._WidthInPix=2646"
      Splits(1)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(1)._ColumnProps(11)=   "Column(1).Visible=0"
      Splits(1)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(13)=   "Column(2).Width=2170"
      Splits(1)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(15)=   "Column(2)._WidthInPix=2090"
      Splits(1)._ColumnProps(16)=   "Column(2)._ColStyle=513"
      Splits(1)._ColumnProps(17)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(18)=   "Column(3).Width=2990"
      Splits(1)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(20)=   "Column(3)._WidthInPix=2910"
      Splits(1)._ColumnProps(21)=   "Column(3)._ColStyle=516"
      Splits(1)._ColumnProps(22)=   "Column(3).Visible=0"
      Splits(1)._ColumnProps(23)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(24)=   "Column(4).Width=2858"
      Splits(1)._ColumnProps(25)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(26)=   "Column(4)._WidthInPix=2778"
      Splits(1)._ColumnProps(27)=   "Column(4)._ColStyle=513"
      Splits(1)._ColumnProps(28)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(29)=   "Column(5).Width=7488"
      Splits(1)._ColumnProps(30)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(31)=   "Column(5)._WidthInPix=7408"
      Splits(1)._ColumnProps(32)=   "Column(5)._ColStyle=516"
      Splits(1)._ColumnProps(33)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(34)=   "Column(6).Width=2831"
      Splits(1)._ColumnProps(35)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(36)=   "Column(6)._WidthInPix=2752"
      Splits(1)._ColumnProps(37)=   "Column(6)._ColStyle=514"
      Splits(1)._ColumnProps(38)=   "Column(6).Order=7"
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
      Caption         =   "LIST OF POTONGAN PPH21"
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
      _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
      _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
      _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
      _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
      _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=94,.parent=13"
      _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=91,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=92,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=93,.parent=17"
      _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=110,.parent=13"
      _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=14"
      _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=15"
      _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=17"
      _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
      _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
      _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=102,.parent=13"
      _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=99,.parent=14"
      _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=100,.parent=15"
      _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=101,.parent=17"
      _StyleDefs(62)  =   "Splits(1).Style:id=43,.parent=1"
      _StyleDefs(63)  =   "Splits(1).CaptionStyle:id=56,.parent=4,.bgcolor=&H80000002&,.fgcolor=&H80000009&"
      _StyleDefs(64)  =   "Splits(1).HeadingStyle:id=44,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(65)  =   ":id=44,.fgcolor=&H80000002&"
      _StyleDefs(66)  =   "Splits(1).FooterStyle:id=45,.parent=3"
      _StyleDefs(67)  =   "Splits(1).InactiveStyle:id=46,.parent=5"
      _StyleDefs(68)  =   "Splits(1).SelectedStyle:id=52,.parent=6"
      _StyleDefs(69)  =   "Splits(1).EditorStyle:id=51,.parent=7"
      _StyleDefs(70)  =   "Splits(1).HighlightRowStyle:id=53,.parent=8"
      _StyleDefs(71)  =   "Splits(1).EvenRowStyle:id=54,.parent=9"
      _StyleDefs(72)  =   "Splits(1).OddRowStyle:id=55,.parent=10"
      _StyleDefs(73)  =   "Splits(1).RecordSelectorStyle:id=57,.parent=11"
      _StyleDefs(74)  =   "Splits(1).FilterBarStyle:id=58,.parent=12"
      _StyleDefs(75)  =   "Splits(1).Columns(0).Style:id=90,.parent=43"
      _StyleDefs(76)  =   "Splits(1).Columns(0).HeadingStyle:id=87,.parent=44"
      _StyleDefs(77)  =   "Splits(1).Columns(0).FooterStyle:id=88,.parent=45"
      _StyleDefs(78)  =   "Splits(1).Columns(0).EditorStyle:id=89,.parent=51"
      _StyleDefs(79)  =   "Splits(1).Columns(1).Style:id=98,.parent=43"
      _StyleDefs(80)  =   "Splits(1).Columns(1).HeadingStyle:id=95,.parent=44"
      _StyleDefs(81)  =   "Splits(1).Columns(1).FooterStyle:id=96,.parent=45"
      _StyleDefs(82)  =   "Splits(1).Columns(1).EditorStyle:id=97,.parent=51"
      _StyleDefs(83)  =   "Splits(1).Columns(2).Style:id=114,.parent=43,.alignment=2"
      _StyleDefs(84)  =   "Splits(1).Columns(2).HeadingStyle:id=111,.parent=44"
      _StyleDefs(85)  =   "Splits(1).Columns(2).FooterStyle:id=112,.parent=45"
      _StyleDefs(86)  =   "Splits(1).Columns(2).EditorStyle:id=113,.parent=51"
      _StyleDefs(87)  =   "Splits(1).Columns(3).Style:id=62,.parent=43"
      _StyleDefs(88)  =   "Splits(1).Columns(3).HeadingStyle:id=59,.parent=44"
      _StyleDefs(89)  =   "Splits(1).Columns(3).FooterStyle:id=60,.parent=45"
      _StyleDefs(90)  =   "Splits(1).Columns(3).EditorStyle:id=61,.parent=51"
      _StyleDefs(91)  =   "Splits(1).Columns(4).Style:id=66,.parent=43,.alignment=2"
      _StyleDefs(92)  =   "Splits(1).Columns(4).HeadingStyle:id=63,.parent=44"
      _StyleDefs(93)  =   "Splits(1).Columns(4).FooterStyle:id=64,.parent=45"
      _StyleDefs(94)  =   "Splits(1).Columns(4).EditorStyle:id=65,.parent=51"
      _StyleDefs(95)  =   "Splits(1).Columns(5).Style:id=82,.parent=43"
      _StyleDefs(96)  =   "Splits(1).Columns(5).HeadingStyle:id=79,.parent=44"
      _StyleDefs(97)  =   "Splits(1).Columns(5).FooterStyle:id=80,.parent=45"
      _StyleDefs(98)  =   "Splits(1).Columns(5).EditorStyle:id=81,.parent=51"
      _StyleDefs(99)  =   "Splits(1).Columns(6).Style:id=106,.parent=43,.alignment=1"
      _StyleDefs(100) =   "Splits(1).Columns(6).HeadingStyle:id=103,.parent=44"
      _StyleDefs(101) =   "Splits(1).Columns(6).FooterStyle:id=104,.parent=45"
      _StyleDefs(102) =   "Splits(1).Columns(6).EditorStyle:id=105,.parent=51"
      _StyleDefs(103) =   "Named:id=33:Normal"
      _StyleDefs(104) =   ":id=33,.parent=0"
      _StyleDefs(105) =   "Named:id=34:Heading"
      _StyleDefs(106) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(107) =   ":id=34,.wraptext=-1"
      _StyleDefs(108) =   "Named:id=35:Footing"
      _StyleDefs(109) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(110) =   "Named:id=36:Selected"
      _StyleDefs(111) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(112) =   "Named:id=37:Caption"
      _StyleDefs(113) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(114) =   "Named:id=38:HighlightRow"
      _StyleDefs(115) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(116) =   "Named:id=39:EvenRow"
      _StyleDefs(117) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(118) =   "Named:id=40:OddRow"
      _StyleDefs(119) =   ":id=40,.parent=33"
      _StyleDefs(120) =   "Named:id=41:RecordSelector"
      _StyleDefs(121) =   ":id=41,.parent=34"
      _StyleDefs(122) =   "Named:id=42:FilterBar"
      _StyleDefs(123) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBList60.TDBCombo TDBCombo_company 
      Height          =   375
      Left            =   1530
      OleObjectBlob   =   "frm_trans_pph21.frx":4842
      TabIndex        =   15
      Top             =   1230
      Width           =   1275
   End
   Begin prj_panji.vbButton cmdRefresh 
      Height          =   495
      Left            =   2850
      TabIndex        =   19
      Top             =   690
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   873
      BTYPE           =   14
      TX              =   "Lihat"
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
      MICON           =   "frm_trans_pph21.frx":67A8
      PICN            =   "frm_trans_pph21.frx":67C4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton vbButton2 
      Height          =   465
      Left            =   13170
      TabIndex        =   17
      Top             =   1140
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
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
      MICON           =   "frm_trans_pph21.frx":7856
      PICN            =   "frm_trans_pph21.frx":7872
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prj_panji.vbButton vbButton1 
      Height          =   465
      Left            =   11730
      TabIndex        =   18
      Top             =   1140
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   820
      BTYPE           =   14
      TX              =   "&Proses"
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
      MICON           =   "frm_trans_pph21.frx":8904
      PICN            =   "frm_trans_pph21.frx":8920
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PERUSAHAAN"
      Height          =   195
      Left            =   270
      TabIndex        =   16
      Top             =   1290
      Width           =   1185
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "POTONGAN PPH 21 BULANAN"
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
   Begin VB.Label Label1 
      Caption         =   "DATE"
      Height          =   195
      Left            =   270
      TabIndex        =   3
      Top             =   870
      Width           =   705
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_trans_pph21.frx":99B2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14760
   End
End
Attribute VB_Name = "frm_trans_pph21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompany As New ADODB.Recordset
    
Dim rs As New ADODB.Recordset
Dim rsPPh21 As New ADODB.Recordset

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns

Dim int_mode As Integer
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Dim oClause As String

Private Sub cmdSave_Click()
    Me.MousePointer = vbHourglass
    
    SQL = "SELECT employee_code FROM temp_list"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        CnG.BeginTrans
        rs.MoveFirst
        While Not rs.EOF
            SQL = "UPDATE t_pph21 SET pph21_value = '" & DropAllComma(txt_pph21_value.Text) & "' " & _
                  "WHERE employee_code = '" & rs!employee_code & "' " & _
                    "AND month = '" & Format(DTPicker1.Value, "yyyy-MM") & "'"
            CnG.Execute SQL
            rs.MoveNext
        Wend
        CnG.CommitTrans
    End If
    rs.Close
    
    Me.MousePointer = vbNormal
    
    Call load_data_pph21
    int_mode = 0
    Call load_mode
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    
'    timer1.Enabled = True
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

Public Sub isiPPh21()
    
    SQL = "SELECT month FROM t_pph21 WHERE month = '" & Format(DTPicker1.Value, "yyyy-MM") & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        MsgBox "Data PPh21 Untuk Periode " & "'" & Format(DTPicker1.Value, "yyyy-MM") & "'" & " Sudah Ada...", vbExclamation, headerMSG
        Exit Sub
    End If
    rs.Close

    Me.MousePointer = vbHourglass
    
    SQL = "SELECT employee_code FROM m_employee WHERE company_code = '" & TDBCombo_company.Text & "' " & _
            "AND flag_active <> 0"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        CnG.BeginTrans
        rs.MoveFirst
        While Not rs.EOF
            SQL = "INSERT INTO t_pph21 (month, employee_code, pph21_value, entry_date, entry_user) " & _
                  "VALUES " & _
                    "('" & Format(DTPicker1.Value, "yyyy-MM") & "','" & rs!employee_code & "',0," & _
                    "now(),'" & LOGIN_CODE & "')"
            CnG.Execute SQL
            rs.MoveNext
        Wend
        CnG.CommitTrans
    End If
    rs.Close
    
    load_data_pph21
    Me.MousePointer = vbNormal
End Sub

Private Sub load_data_pph21()
    If rsPPh21.State Then rsPPh21.Close
    SQL = "SELECT a.employee_code, b.nik, b.employee_name, " & _
            "b.division_code, c.division_name, a.pph21_value, a.month " & _
          "FROM t_pph21 a JOIN m_employee b ON a.employee_code = b.employee_code " & _
            "JOIN m_division c ON b.division_code = c.division_code " & _
          "WHERE a.month = '" & Format(DTPicker1.Value, "yyyy-MM") & "' " & _
            "AND b.company_code = '" & TDBCombo_company.Text & "' " & oClause
    rsPPh21.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid1.DataSource = rsPPh21
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_proses_bonus = Nothing
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name.Text = TDBCombo_company.Columns("company_name").Value
        
        load_data_pph21
    End If
End Sub

'Private Sub Timer1_Timer()
'    timer1.Enabled = False
'
'    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
'End Sub

Private Sub vbButton1_Click()
    isiPPh21
End Sub

Private Sub vbbutton2_Click()
Dim tanya As Integer

    Me.MousePointer = vbHourglass
    
    tanya = MsgBox("Apakah Anda Yakin Akan Menghapus Data PPh21 Periode " & "'" & Format(DTPicker1.Value, "yyyy-MM") & "'?", vbExclamation + vbYesNo, "Warning")
    If tanya = vbYes Then
        SQL = "DELETE a FROM t_pph21 a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                 "WHERE a.month = '" & Format(DTPicker1, "yyyy-MM") & "' " & _
                    "AND b.company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
        load_data_pph21
    End If
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdRefresh_Click()
    load_data_pph21
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid1.Columns
        Col.FilterText = ""
    Next Col
    rsPPh21.Filter = adFilterNone
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
    
    rsPPh21.Filter = getFilter()
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
                & "'" & Replace(TDBGrid1.Columns("employee_name").CellText(item), "'", "''") & "')"
        CnG.Execute SQL
                
    Next
    CnG.CommitTrans
    
    list_employee.Clear
    SQL = "select employee_name from temp_list"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        While Not rs.EOF
            list_employee.AddItem rs!EMPLOYEE_NAME
            rs.MoveNext
        Wend
    End If
    rs.Close
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub cmdEdit_Click()
    txt_pph21_value.Text = FormatNumber(0)
    
    int_mode = 2
    Call load_mode
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
'    cmdNew.Enabled = a And blnUser_Add
    cmdSave.Enabled = b
    cmdEdit.Enabled = c And blnUser_Edit
'    cmdDelete.Enabled = d And blnUser_Delete
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
    Call load_data_pph21

End Sub

Private Sub txt_pph21_value_Validate(Cancel As Boolean)
    If Not Trim(txt_pph21_value) = "" Then
        txt_pph21_value = FormatNumber(DropAllComma(txt_pph21_value))
    End If
End Sub

