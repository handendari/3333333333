VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_list_manual_att 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KEHADIRAN MANUAL"
   ClientHeight    =   11385
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15360
   Icon            =   "frmListManualAtt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   11385
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin prj_panji.LynxGrid LynxGrid3 
      Height          =   3465
      Left            =   6960
      TabIndex        =   29
      Top             =   1710
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   9645
      Left            =   180
      TabIndex        =   0
      Top             =   810
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   17013
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "INPUT MANUAL"
      TabPicture(0)   =   "frmListManualAtt.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame4(0)"
      Tab(0).Control(1)=   "timer1(0)"
      Tab(0).Control(2)=   "cmdPrev(0)"
      Tab(0).Control(3)=   "Frame1(0)"
      Tab(0).Control(4)=   "txt_group_shift(0)"
      Tab(0).Control(5)=   "txt_company_name(0)"
      Tab(0).Control(6)=   "DTPicker1(0)"
      Tab(0).Control(7)=   "TDBCombo_Group_Shift(0)"
      Tab(0).Control(8)=   "TDBCombo_company(0)"
      Tab(0).Control(9)=   "cmdRefresh(0)"
      Tab(0).Control(10)=   "cmdNext(0)"
      Tab(0).Control(11)=   "cmdDelete(0)"
      Tab(0).Control(12)=   "cmdEdit(0)"
      Tab(0).Control(13)=   "cmdNew(0)"
      Tab(0).Control(14)=   "Label2(0)"
      Tab(0).Control(15)=   "Label4(0)"
      Tab(0).Control(16)=   "Label1(0)"
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "DAFTAR ABSENSI TUNGGAL"
      TabPicture(1)   =   "frmListManualAtt.frx":05A6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdNew(1)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdEdit(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdDelete(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdNext(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdRefresh(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdPrev(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "TDBCombo_company(1)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "TDBCombo_Group_Shift(1)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "DTPicker1(1)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Frame1(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txt_group_shift(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txt_company_name(1)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame4(1)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "timer1(1)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "IMPORT DATA KEHADIRAN"
      TabPicture(2)   =   "frmListManualAtt.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(1)=   "ProgressBar1"
      Tab(2).Control(2)=   "cmdImport"
      Tab(2).Control(3)=   "cmdBrowse"
      Tab(2).Control(4)=   "cmdDownload"
      Tab(2).Control(5)=   "Label5"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "IMPORT DATA KEHADIRAN FORMAT MESIN"
      TabPicture(3)   =   "frmListManualAtt.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(1)=   "ProgressBar2"
      Tab(3).Control(2)=   "cmdImport_M"
      Tab(3).Control(3)=   "cmdBrowae_M"
      Tab(3).Control(4)=   "Label6"
      Tab(3).ControlCount=   5
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Index           =   1
         Interval        =   600
         Left            =   630
         Top             =   6270
      End
      Begin VB.Frame Frame4 
         Caption         =   "Advanced Filter"
         Height          =   1365
         Index           =   1
         Left            =   5670
         TabIndex        =   58
         Top             =   330
         Width           =   5025
         Begin VB.TextBox txt_nik 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1110
            TabIndex        =   62
            Top             =   270
            Width           =   855
         End
         Begin VB.TextBox txt_employee_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            DragMode        =   1  'Automatic
            Height          =   285
            Index           =   1
            Left            =   2370
            TabIndex        =   61
            Top             =   270
            Width           =   2535
         End
         Begin VB.TextBox txt_employee_code 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   4710
            TabIndex        =   60
            Top             =   270
            Visible         =   0   'False
            Width           =   255
         End
         Begin prj_panji.vbButton cmdBrowse_Emp 
            Height          =   285
            Index           =   1
            Left            =   1980
            TabIndex        =   59
            Top             =   270
            Width           =   375
            _ExtentX        =   661
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
            MICON           =   "frmListManualAtt.frx":05FA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComCtl2.DTPicker DTPicker_from 
            Height          =   315
            Index           =   1
            Left            =   1110
            TabIndex        =   63
            Top             =   750
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   91291651
            CurrentDate     =   40794
         End
         Begin MSComCtl2.DTPicker DTPicker_to 
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   64
            Top             =   750
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   91291651
            CurrentDate     =   40794
         End
         Begin prj_panji.vbButton cmdSearch 
            Height          =   675
            Index           =   1
            Left            =   4140
            TabIndex        =   65
            Top             =   600
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   1191
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
            MICON           =   "frmListManualAtt.frx":0616
            PICN            =   "frmListManualAtt.frx":0632
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "TANGGAL"
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   780
            Width           =   765
         End
         Begin VB.Label Label10 
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
            Left            =   2430
            TabIndex        =   67
            Top             =   780
            Width           =   285
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "KARYAWAN"
            Height          =   195
            Left            =   120
            TabIndex        =   66
            Top             =   300
            Width           =   930
         End
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Index           =   1
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   50
         Top             =   930
         Width           =   2775
      End
      Begin VB.TextBox txt_group_shift 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Index           =   1
         Left            =   2610
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   49
         Top             =   1290
         Width           =   2775
      End
      Begin VB.Frame Frame1 
         Height          =   7665
         Index           =   1
         Left            =   90
         TabIndex        =   43
         Top             =   1800
         Width           =   14835
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_Shift 
            Height          =   3735
            Index           =   1
            Left            =   60
            TabIndex        =   44
            Top             =   180
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   6588
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE SHIFT"
            Columns(0).DataField=   "shift_code"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "NAMA SHIFT"
            Columns(1).DataField=   "shift_name"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "FLAG DAY OVER"
            Columns(2).DataField=   "flag_day_over"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3545"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3466"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
            Caption         =   "DAFTAR SHIFT"
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Named:id=33:Normal"
            _StyleDefs(47)  =   ":id=33,.parent=0"
            _StyleDefs(48)  =   "Named:id=34:Heading"
            _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   ":id=34,.wraptext=-1"
            _StyleDefs(51)  =   "Named:id=35:Footing"
            _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   "Named:id=36:Selected"
            _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(55)  =   "Named:id=37:Caption"
            _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(57)  =   "Named:id=38:HighlightRow"
            _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(59)  =   "Named:id=39:EvenRow"
            _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(61)  =   "Named:id=40:OddRow"
            _StyleDefs(62)  =   ":id=40,.parent=33"
            _StyleDefs(63)  =   "Named:id=41:RecordSelector"
            _StyleDefs(64)  =   ":id=41,.parent=34"
            _StyleDefs(65)  =   "Named:id=42:FilterBar"
            _StyleDefs(66)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_Att 
            Height          =   7395
            Index           =   1
            Left            =   3720
            TabIndex        =   45
            Top             =   180
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   13044
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "TGL"
            Columns(0).DataField=   "_date"
            Columns(0).NumberFormat=   "yyyy-MM-dd"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "EMPLOYEE CODE"
            Columns(1).DataField=   "employee_code"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "DIV."
            Columns(2).DataField=   "division_name"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "KODE KARY."
            Columns(3).DataField=   "nik"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "ENROLLNUMBER"
            Columns(4).DataField=   "enrollnumber"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "NAMA KARY."
            Columns(5).DataField=   "employee_name"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "STATUS"
            Columns(6).DataField=   "status"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "TITLE CODE"
            Columns(7).DataField=   "title_code"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "JABATAN"
            Columns(8).DataField=   "title_name"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "SHIFT"
            Columns(9).DataField=   "shift_code"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "SHIFT NAME"
            Columns(10).DataField=   "shift_name"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "MASUK"
            Columns(11).DataField=   "time_in"
            Columns(11).NumberFormat=   "hh:mm"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "PULANG"
            Columns(12).DataField=   "time_out"
            Columns(12).NumberFormat=   "hh:mm"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "TGL. INPUT"
            Columns(13).DataField=   "entry_date"
            Columns(13).NumberFormat=   "yyyy-MM-dd"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "DEPT. CODE"
            Columns(14).DataField=   "department_code"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "DEPT. NAME"
            Columns(15).DataField=   "department_name"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "DIV. CODE"
            Columns(16).DataField=   "division_code"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "DIV. NAME"
            Columns(17).DataField=   "division_name"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(18)._VlistStyle=   0
            Columns(18)._MaxComboItems=   5
            Columns(18).Caption=   "DESCRIPTION"
            Columns(18).DataField=   "description"
            Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(19)._VlistStyle=   0
            Columns(19)._MaxComboItems=   5
            Columns(19).Caption=   "STATUS NAME"
            Columns(19).DataField=   "absent_name"
            Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(20)._VlistStyle=   0
            Columns(20)._MaxComboItems=   5
            Columns(20).Caption=   "ATT DATE"
            Columns(20).DataField=   "att_date"
            Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   21
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=21"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2037"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1958"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2143"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2064"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=2037"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1958"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=513"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(28)=   "Column(5).Width=3440"
            Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=3360"
            Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(33)=   "Column(6).Width=1376"
            Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=1296"
            Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(38)=   "Column(7).Width=2540"
            Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2461"
            Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(42)=   "Column(7).Visible=0"
            Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(44)=   "Column(8).Width=2381"
            Splits(0)._ColumnProps(45)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(8)._WidthInPix=2302"
            Splits(0)._ColumnProps(47)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(48)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(49)=   "Column(8)._MinWidth=10"
            Splits(0)._ColumnProps(50)=   "Column(9).Width=1244"
            Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=1164"
            Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=513"
            Splits(0)._ColumnProps(54)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(55)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(56)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(58)=   "Column(10)._ColStyle=516"
            Splits(0)._ColumnProps(59)=   "Column(10).Visible=0"
            Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(61)=   "Column(11).Width=1535"
            Splits(0)._ColumnProps(62)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(63)=   "Column(11)._WidthInPix=1455"
            Splits(0)._ColumnProps(64)=   "Column(11)._ColStyle=513"
            Splits(0)._ColumnProps(65)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(66)=   "Column(11)._MinWidth=54215968"
            Splits(0)._ColumnProps(67)=   "Column(12).Width=1588"
            Splits(0)._ColumnProps(68)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(69)=   "Column(12)._WidthInPix=1508"
            Splits(0)._ColumnProps(70)=   "Column(12)._ColStyle=513"
            Splits(0)._ColumnProps(71)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(72)=   "Column(12)._MinWidth=54215968"
            Splits(0)._ColumnProps(73)=   "Column(13).Width=2275"
            Splits(0)._ColumnProps(74)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(75)=   "Column(13)._WidthInPix=2196"
            Splits(0)._ColumnProps(76)=   "Column(13)._ColStyle=516"
            Splits(0)._ColumnProps(77)=   "Column(13).Visible=0"
            Splits(0)._ColumnProps(78)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(79)=   "Column(14).Width=2725"
            Splits(0)._ColumnProps(80)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(81)=   "Column(14)._WidthInPix=2646"
            Splits(0)._ColumnProps(82)=   "Column(14)._ColStyle=516"
            Splits(0)._ColumnProps(83)=   "Column(14).Visible=0"
            Splits(0)._ColumnProps(84)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(85)=   "Column(15).Width=2725"
            Splits(0)._ColumnProps(86)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(87)=   "Column(15)._WidthInPix=2646"
            Splits(0)._ColumnProps(88)=   "Column(15)._ColStyle=516"
            Splits(0)._ColumnProps(89)=   "Column(15).Visible=0"
            Splits(0)._ColumnProps(90)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(91)=   "Column(16).Width=2725"
            Splits(0)._ColumnProps(92)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(93)=   "Column(16)._WidthInPix=2646"
            Splits(0)._ColumnProps(94)=   "Column(16)._ColStyle=516"
            Splits(0)._ColumnProps(95)=   "Column(16).Visible=0"
            Splits(0)._ColumnProps(96)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(97)=   "Column(17).Width=2725"
            Splits(0)._ColumnProps(98)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(99)=   "Column(17)._WidthInPix=2646"
            Splits(0)._ColumnProps(100)=   "Column(17)._ColStyle=516"
            Splits(0)._ColumnProps(101)=   "Column(17).Visible=0"
            Splits(0)._ColumnProps(102)=   "Column(17).Order=18"
            Splits(0)._ColumnProps(103)=   "Column(18).Width=2725"
            Splits(0)._ColumnProps(104)=   "Column(18).DividerColor=0"
            Splits(0)._ColumnProps(105)=   "Column(18)._WidthInPix=2646"
            Splits(0)._ColumnProps(106)=   "Column(18)._ColStyle=516"
            Splits(0)._ColumnProps(107)=   "Column(18).Visible=0"
            Splits(0)._ColumnProps(108)=   "Column(18).Order=19"
            Splits(0)._ColumnProps(109)=   "Column(19).Width=2725"
            Splits(0)._ColumnProps(110)=   "Column(19).DividerColor=0"
            Splits(0)._ColumnProps(111)=   "Column(19)._WidthInPix=2646"
            Splits(0)._ColumnProps(112)=   "Column(19)._ColStyle=516"
            Splits(0)._ColumnProps(113)=   "Column(19).Visible=0"
            Splits(0)._ColumnProps(114)=   "Column(19).Order=20"
            Splits(0)._ColumnProps(115)=   "Column(20).Width=2725"
            Splits(0)._ColumnProps(116)=   "Column(20).DividerColor=0"
            Splits(0)._ColumnProps(117)=   "Column(20)._WidthInPix=2646"
            Splits(0)._ColumnProps(118)=   "Column(20)._ColStyle=516"
            Splits(0)._ColumnProps(119)=   "Column(20).Visible=0"
            Splits(0)._ColumnProps(120)=   "Column(20).Order=21"
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
            Caption         =   "DAFTAR KEHADIRAN"
            AllowArrows     =   0   'False
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=90,.parent=13,.alignment=2"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=118,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=115,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=116,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=117,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=94,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=102,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=99,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=100,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=101,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=98,.parent=13,.alignment=2"
            _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=95,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=96,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=97,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=106,.parent=13"
            _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=103,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=104,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=105,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=110,.parent=13,.alignment=2"
            _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=107,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=108,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=109,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=74,.parent=13,.alignment=2"
            _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
            _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=28,.parent=13"
            _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=25,.parent=14"
            _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=26,.parent=15"
            _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=27,.parent=17"
            _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=46,.parent=13"
            _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=43,.parent=14"
            _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=44,.parent=15"
            _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=45,.parent=17"
            _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=58,.parent=13"
            _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=55,.parent=14"
            _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=56,.parent=15"
            _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=57,.parent=17"
            _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=70,.parent=13"
            _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=67,.parent=14"
            _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=68,.parent=15"
            _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=69,.parent=17"
            _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=78,.parent=13"
            _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=75,.parent=14"
            _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=76,.parent=15"
            _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=77,.parent=17"
            _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=82,.parent=13"
            _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=79,.parent=14"
            _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=80,.parent=15"
            _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=81,.parent=17"
            _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=86,.parent=13"
            _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=83,.parent=14"
            _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=84,.parent=15"
            _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=85,.parent=17"
            _StyleDefs(114) =   "Splits(0).Columns(20).Style:id=114,.parent=13"
            _StyleDefs(115) =   "Splits(0).Columns(20).HeadingStyle:id=111,.parent=14"
            _StyleDefs(116) =   "Splits(0).Columns(20).FooterStyle:id=112,.parent=15"
            _StyleDefs(117) =   "Splits(0).Columns(20).EditorStyle:id=113,.parent=17"
            _StyleDefs(118) =   "Named:id=33:Normal"
            _StyleDefs(119) =   ":id=33,.parent=0"
            _StyleDefs(120) =   "Named:id=34:Heading"
            _StyleDefs(121) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(122) =   ":id=34,.wraptext=-1"
            _StyleDefs(123) =   "Named:id=35:Footing"
            _StyleDefs(124) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(125) =   "Named:id=36:Selected"
            _StyleDefs(126) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(127) =   "Named:id=37:Caption"
            _StyleDefs(128) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(129) =   "Named:id=38:HighlightRow"
            _StyleDefs(130) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(131) =   "Named:id=39:EvenRow"
            _StyleDefs(132) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(133) =   "Named:id=40:OddRow"
            _StyleDefs(134) =   ":id=40,.parent=33"
            _StyleDefs(135) =   "Named:id=41:RecordSelector"
            _StyleDefs(136) =   ":id=41,.parent=34"
            _StyleDefs(137) =   "Named:id=42:FilterBar"
            _StyleDefs(138) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   3615
            Index           =   1
            Left            =   60
            TabIndex        =   46
            Top             =   3960
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   6376
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
            Columns(1).Caption=   "DATE"
            Columns(1).DataField=   "date_att"
            Columns(1).NumberFormat=   "yyyy-MM-dd"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "TIME"
            Columns(2).DataField=   "time_att"
            Columns(2).NumberFormat=   "HH:mm:ss"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "TYPE"
            Columns(3).DataField=   "type"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "ATT DATE"
            Columns(4).DataField=   "att_date"
            Columns(4).NumberFormat=   "yyyy-MM-dd hh:mm:ss"
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
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=1984"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1905"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=1164"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1085"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(22)=   "Column(4).Width=3678"
            Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3598"
            Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
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
            Caption         =   "DAFTAR CHECK LOG"
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
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
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
      Begin VB.Frame Frame2 
         Height          =   7980
         Left            =   -74850
         TabIndex        =   36
         Top             =   480
         Width           =   14715
         Begin prj_panji.LynxGrid LynxGrid1 
            Height          =   7635
            Left            =   120
            TabIndex        =   37
            Top             =   210
            Width           =   14475
            _ExtentX        =   25532
            _ExtentY        =   13467
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
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   330
            Left            =   30
            Top             =   120
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
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
      End
      Begin VB.Frame Frame3 
         Height          =   7980
         Left            =   -74880
         TabIndex        =   30
         Top             =   420
         Width           =   14715
         Begin prj_panji.LynxGrid LynxGrid2 
            Height          =   7635
            Left            =   120
            TabIndex        =   31
            Top             =   210
            Width           =   14475
            _ExtentX        =   25532
            _ExtentY        =   13467
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
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   330
            Left            =   30
            Top             =   120
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   582
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
      End
      Begin VB.Frame Frame4 
         Caption         =   "Advanced Filter"
         Height          =   1365
         Index           =   0
         Left            =   -69330
         TabIndex        =   17
         Top             =   330
         Width           =   5025
         Begin prj_panji.vbButton cmdBrowse_Emp 
            Height          =   285
            Index           =   0
            Left            =   1980
            TabIndex        =   26
            Top             =   270
            Width           =   375
            _ExtentX        =   661
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
            MICON           =   "frmListManualAtt.frx":16C4
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
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   4710
            TabIndex        =   22
            Top             =   270
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txt_employee_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            DragMode        =   1  'Automatic
            Height          =   285
            Index           =   0
            Left            =   2370
            TabIndex        =   19
            Top             =   270
            Width           =   2535
         End
         Begin VB.TextBox txt_nik 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1110
            TabIndex        =   18
            Top             =   270
            Width           =   855
         End
         Begin MSComCtl2.DTPicker DTPicker_from 
            Height          =   315
            Index           =   0
            Left            =   1110
            TabIndex        =   20
            Top             =   750
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   91291651
            CurrentDate     =   40794
         End
         Begin MSComCtl2.DTPicker DTPicker_to 
            Height          =   315
            Index           =   0
            Left            =   2790
            TabIndex        =   21
            Top             =   750
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   91291651
            CurrentDate     =   40794
         End
         Begin prj_panji.vbButton cmdSearch 
            Height          =   675
            Index           =   0
            Left            =   4140
            TabIndex        =   27
            Top             =   600
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   1191
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
            MICON           =   "frmListManualAtt.frx":16E0
            PICN            =   "frmListManualAtt.frx":16FC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "KARYAWAN"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   300
            Width           =   930
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
            Left            =   2430
            TabIndex        =   24
            Top             =   780
            Width           =   285
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "TANGGAL"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   780
            Width           =   765
         End
      End
      Begin VB.Timer timer1 
         Enabled         =   0   'False
         Index           =   0
         Interval        =   600
         Left            =   -75000
         Top             =   7530
      End
      Begin prj_panji.vbButton cmdPrev 
         Height          =   285
         Index           =   0
         Left            =   -72390
         TabIndex        =   13
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "<"
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
         MICON           =   "frmListManualAtt.frx":278E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame1 
         Height          =   7665
         Index           =   0
         Left            =   -74910
         TabIndex        =   5
         Top             =   1800
         Width           =   14835
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_Shift 
            Height          =   3735
            Index           =   0
            Left            =   60
            TabIndex        =   12
            Top             =   180
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   6588
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE SHIFT"
            Columns(0).DataField=   "shift_code"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "NAMA SHIFT"
            Columns(1).DataField=   "shift_name"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "FLAG DAY OVER"
            Columns(2).DataField=   "flag_day_over"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1799"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1720"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3545"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3466"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
            Caption         =   "DAFTAR SHIFT"
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
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
            _StyleDefs(46)  =   "Named:id=33:Normal"
            _StyleDefs(47)  =   ":id=33,.parent=0"
            _StyleDefs(48)  =   "Named:id=34:Heading"
            _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(50)  =   ":id=34,.wraptext=-1"
            _StyleDefs(51)  =   "Named:id=35:Footing"
            _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(53)  =   "Named:id=36:Selected"
            _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(55)  =   "Named:id=37:Caption"
            _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(57)  =   "Named:id=38:HighlightRow"
            _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(59)  =   "Named:id=39:EvenRow"
            _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(61)  =   "Named:id=40:OddRow"
            _StyleDefs(62)  =   ":id=40,.parent=33"
            _StyleDefs(63)  =   "Named:id=41:RecordSelector"
            _StyleDefs(64)  =   ":id=41,.parent=34"
            _StyleDefs(65)  =   "Named:id=42:FilterBar"
            _StyleDefs(66)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid_Att 
            Height          =   7395
            Index           =   0
            Left            =   3720
            TabIndex        =   15
            Top             =   180
            Width           =   10995
            _ExtentX        =   19394
            _ExtentY        =   13044
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "TGL"
            Columns(0).DataField=   "_date"
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
            Columns(3).Caption=   "ENROLLNUMBER"
            Columns(3).DataField=   "enrollnumber"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "NAMA KARY."
            Columns(4).DataField=   "employee_name"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "STATUS"
            Columns(5).DataField=   "status"
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
            Columns(8).Caption=   "SHIFT"
            Columns(8).DataField=   "shift_code"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "SHIFT NAME"
            Columns(9).DataField=   "shift_name"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "MASUK"
            Columns(10).DataField=   "time_in"
            Columns(10).NumberFormat=   "hh:mm"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "PULANG"
            Columns(11).DataField=   "time_out"
            Columns(11).NumberFormat=   "hh:mm"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "TGL. INPUT"
            Columns(12).DataField=   "entry_date"
            Columns(12).NumberFormat=   "yyyy-MM-dd"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "DEPT. CODE"
            Columns(13).DataField=   "department_code"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "DEPT. NAME"
            Columns(14).DataField=   "department_name"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "DIV. CODE"
            Columns(15).DataField=   "division_code"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   0
            Columns(16)._MaxComboItems=   5
            Columns(16).Caption=   "DIV. NAME"
            Columns(16).DataField=   "division_name"
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(17)._VlistStyle=   0
            Columns(17)._MaxComboItems=   5
            Columns(17).Caption=   "DESCRIPTION"
            Columns(17).DataField=   "description"
            Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(18)._VlistStyle=   0
            Columns(18)._MaxComboItems=   5
            Columns(18).Caption=   "STATUS NAME"
            Columns(18).DataField=   "absent_name"
            Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(19)._VlistStyle=   0
            Columns(19)._MaxComboItems=   5
            Columns(19).Caption=   "ATT DATE"
            Columns(19).DataField=   "att_date"
            Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   20
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=20"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2223"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2143"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2143"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2064"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=2223"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2143"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=3440"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=3360"
            Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(28)=   "Column(5).Width=1429"
            Splits(0)._ColumnProps(29)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(30)=   "Column(5)._WidthInPix=1349"
            Splits(0)._ColumnProps(31)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(32)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(33)=   "Column(6).Width=2540"
            Splits(0)._ColumnProps(34)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(6)._WidthInPix=2461"
            Splits(0)._ColumnProps(36)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(38)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(39)=   "Column(7).Width=2540"
            Splits(0)._ColumnProps(40)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(7)._WidthInPix=2461"
            Splits(0)._ColumnProps(42)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(44)=   "Column(7)._MinWidth=10"
            Splits(0)._ColumnProps(45)=   "Column(8).Width=1376"
            Splits(0)._ColumnProps(46)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(8)._WidthInPix=1296"
            Splits(0)._ColumnProps(48)=   "Column(8)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(50)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(51)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(53)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(54)=   "Column(9).Visible=0"
            Splits(0)._ColumnProps(55)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(56)=   "Column(10).Width=1535"
            Splits(0)._ColumnProps(57)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(10)._WidthInPix=1455"
            Splits(0)._ColumnProps(59)=   "Column(10)._ColStyle=513"
            Splits(0)._ColumnProps(60)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(61)=   "Column(10)._MinWidth=54215968"
            Splits(0)._ColumnProps(62)=   "Column(11).Width=1588"
            Splits(0)._ColumnProps(63)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(11)._WidthInPix=1508"
            Splits(0)._ColumnProps(65)=   "Column(11)._ColStyle=513"
            Splits(0)._ColumnProps(66)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(67)=   "Column(11)._MinWidth=54215968"
            Splits(0)._ColumnProps(68)=   "Column(12).Width=2275"
            Splits(0)._ColumnProps(69)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(70)=   "Column(12)._WidthInPix=2196"
            Splits(0)._ColumnProps(71)=   "Column(12)._ColStyle=516"
            Splits(0)._ColumnProps(72)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(73)=   "Column(13).Width=2725"
            Splits(0)._ColumnProps(74)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(75)=   "Column(13)._WidthInPix=2646"
            Splits(0)._ColumnProps(76)=   "Column(13)._ColStyle=516"
            Splits(0)._ColumnProps(77)=   "Column(13).Visible=0"
            Splits(0)._ColumnProps(78)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(79)=   "Column(14).Width=2725"
            Splits(0)._ColumnProps(80)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(81)=   "Column(14)._WidthInPix=2646"
            Splits(0)._ColumnProps(82)=   "Column(14)._ColStyle=516"
            Splits(0)._ColumnProps(83)=   "Column(14).Visible=0"
            Splits(0)._ColumnProps(84)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(85)=   "Column(15).Width=2725"
            Splits(0)._ColumnProps(86)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(87)=   "Column(15)._WidthInPix=2646"
            Splits(0)._ColumnProps(88)=   "Column(15)._ColStyle=516"
            Splits(0)._ColumnProps(89)=   "Column(15).Visible=0"
            Splits(0)._ColumnProps(90)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(91)=   "Column(16).Width=2725"
            Splits(0)._ColumnProps(92)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(93)=   "Column(16)._WidthInPix=2646"
            Splits(0)._ColumnProps(94)=   "Column(16)._ColStyle=516"
            Splits(0)._ColumnProps(95)=   "Column(16).Visible=0"
            Splits(0)._ColumnProps(96)=   "Column(16).Order=17"
            Splits(0)._ColumnProps(97)=   "Column(17).Width=2725"
            Splits(0)._ColumnProps(98)=   "Column(17).DividerColor=0"
            Splits(0)._ColumnProps(99)=   "Column(17)._WidthInPix=2646"
            Splits(0)._ColumnProps(100)=   "Column(17)._ColStyle=516"
            Splits(0)._ColumnProps(101)=   "Column(17).Visible=0"
            Splits(0)._ColumnProps(102)=   "Column(17).Order=18"
            Splits(0)._ColumnProps(103)=   "Column(18).Width=2725"
            Splits(0)._ColumnProps(104)=   "Column(18).DividerColor=0"
            Splits(0)._ColumnProps(105)=   "Column(18)._WidthInPix=2646"
            Splits(0)._ColumnProps(106)=   "Column(18)._ColStyle=516"
            Splits(0)._ColumnProps(107)=   "Column(18).Visible=0"
            Splits(0)._ColumnProps(108)=   "Column(18).Order=19"
            Splits(0)._ColumnProps(109)=   "Column(19).Width=2725"
            Splits(0)._ColumnProps(110)=   "Column(19).DividerColor=0"
            Splits(0)._ColumnProps(111)=   "Column(19)._WidthInPix=2646"
            Splits(0)._ColumnProps(112)=   "Column(19)._ColStyle=516"
            Splits(0)._ColumnProps(113)=   "Column(19).Visible=0"
            Splits(0)._ColumnProps(114)=   "Column(19).Order=20"
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
            Caption         =   "DAFTAR KEHADIRAN"
            AllowArrows     =   0   'False
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
            _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=90,.parent=13,.alignment=2"
            _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
            _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=94,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=91,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=92,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=93,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=102,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=99,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=100,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=101,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=98,.parent=13,.alignment=2"
            _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=106,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=103,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=104,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=105,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=110,.parent=13,.alignment=2"
            _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=107,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=108,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=109,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=74,.parent=13,.alignment=2"
            _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=28,.parent=13"
            _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=25,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=26,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=27,.parent=17"
            _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=46,.parent=13"
            _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=43,.parent=14"
            _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=44,.parent=15"
            _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=45,.parent=17"
            _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=58,.parent=13"
            _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=55,.parent=14"
            _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=56,.parent=15"
            _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=57,.parent=17"
            _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=70,.parent=13"
            _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=67,.parent=14"
            _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=68,.parent=15"
            _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=69,.parent=17"
            _StyleDefs(98)  =   "Splits(0).Columns(16).Style:id=78,.parent=13"
            _StyleDefs(99)  =   "Splits(0).Columns(16).HeadingStyle:id=75,.parent=14"
            _StyleDefs(100) =   "Splits(0).Columns(16).FooterStyle:id=76,.parent=15"
            _StyleDefs(101) =   "Splits(0).Columns(16).EditorStyle:id=77,.parent=17"
            _StyleDefs(102) =   "Splits(0).Columns(17).Style:id=82,.parent=13"
            _StyleDefs(103) =   "Splits(0).Columns(17).HeadingStyle:id=79,.parent=14"
            _StyleDefs(104) =   "Splits(0).Columns(17).FooterStyle:id=80,.parent=15"
            _StyleDefs(105) =   "Splits(0).Columns(17).EditorStyle:id=81,.parent=17"
            _StyleDefs(106) =   "Splits(0).Columns(18).Style:id=86,.parent=13"
            _StyleDefs(107) =   "Splits(0).Columns(18).HeadingStyle:id=83,.parent=14"
            _StyleDefs(108) =   "Splits(0).Columns(18).FooterStyle:id=84,.parent=15"
            _StyleDefs(109) =   "Splits(0).Columns(18).EditorStyle:id=85,.parent=17"
            _StyleDefs(110) =   "Splits(0).Columns(19).Style:id=114,.parent=13"
            _StyleDefs(111) =   "Splits(0).Columns(19).HeadingStyle:id=111,.parent=14"
            _StyleDefs(112) =   "Splits(0).Columns(19).FooterStyle:id=112,.parent=15"
            _StyleDefs(113) =   "Splits(0).Columns(19).EditorStyle:id=113,.parent=17"
            _StyleDefs(114) =   "Named:id=33:Normal"
            _StyleDefs(115) =   ":id=33,.parent=0"
            _StyleDefs(116) =   "Named:id=34:Heading"
            _StyleDefs(117) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(118) =   ":id=34,.wraptext=-1"
            _StyleDefs(119) =   "Named:id=35:Footing"
            _StyleDefs(120) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(121) =   "Named:id=36:Selected"
            _StyleDefs(122) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(123) =   "Named:id=37:Caption"
            _StyleDefs(124) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(125) =   "Named:id=38:HighlightRow"
            _StyleDefs(126) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(127) =   "Named:id=39:EvenRow"
            _StyleDefs(128) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(129) =   "Named:id=40:OddRow"
            _StyleDefs(130) =   ":id=40,.parent=33"
            _StyleDefs(131) =   "Named:id=41:RecordSelector"
            _StyleDefs(132) =   ":id=41,.parent=34"
            _StyleDefs(133) =   "Named:id=42:FilterBar"
            _StyleDefs(134) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
            Height          =   3615
            Index           =   0
            Left            =   60
            TabIndex        =   28
            Top             =   3960
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   6376
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
            Columns(1).Caption=   "DATE"
            Columns(1).DataField=   "date_att"
            Columns(1).NumberFormat=   "yyyy-MM-dd"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "TIME"
            Columns(2).DataField=   "time_att"
            Columns(2).NumberFormat=   "HH:mm:ss"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "TYPE"
            Columns(3).DataField=   "type"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "ATT DATE"
            Columns(4).DataField=   "att_date"
            Columns(4).NumberFormat=   "yyyy-MM-dd hh:mm:ss"
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
            Splits(0)._ColumnProps(7)=   "Column(1).Width=2196"
            Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2117"
            Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(12)=   "Column(2).Width=1984"
            Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1905"
            Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=1164"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=1085"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(22)=   "Column(4).Width=3678"
            Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=3598"
            Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(27)=   "Column(4).Order=5"
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
            Caption         =   "DAFTAR CHECK LOG"
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
            _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
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
      Begin VB.TextBox txt_group_shift 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Index           =   0
         Left            =   -72390
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1290
         Width           =   2775
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Index           =   0
         Left            =   -72390
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         Top             =   930
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   0
         Left            =   -73650
         TabIndex        =   4
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   91291651
         CurrentDate     =   40794
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_Group_Shift 
         Height          =   375
         Index           =   0
         Left            =   -73650
         OleObjectBlob   =   "frmListManualAtt.frx":27AA
         TabIndex        =   6
         Top             =   1290
         Width           =   1245
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Index           =   0
         Left            =   -73650
         OleObjectBlob   =   "frmListManualAtt.frx":470B
         TabIndex        =   7
         Top             =   930
         Width           =   1245
      End
      Begin prj_panji.vbButton cmdRefresh 
         Height          =   495
         Index           =   0
         Left            =   -71550
         TabIndex        =   8
         Top             =   390
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
         MICON           =   "frmListManualAtt.frx":6674
         PICN            =   "frmListManualAtt.frx":6690
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdNext 
         Height          =   285
         Index           =   0
         Left            =   -72030
         TabIndex        =   14
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   ">"
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
         MICON           =   "frmListManualAtt.frx":7722
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   225
         Left            =   -68940
         TabIndex        =   32
         Top             =   8670
         Visible         =   0   'False
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin prj_panji.vbButton cmdImport_M 
         Height          =   555
         Left            =   -72900
         TabIndex        =   33
         Top             =   8640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   979
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
         MICON           =   "frmListManualAtt.frx":773E
         PICN            =   "frmListManualAtt.frx":775A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdBrowae_M 
         Height          =   555
         Left            =   -74580
         TabIndex        =   34
         Top             =   8640
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "&Browse"
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
         MICON           =   "frmListManualAtt.frx":87EC
         PICN            =   "frmListManualAtt.frx":8808
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   -69060
         TabIndex        =   38
         Top             =   8730
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin prj_panji.vbButton cmdImport 
         Height          =   555
         Left            =   -73020
         TabIndex        =   39
         Top             =   8700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   979
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
         MICON           =   "frmListManualAtt.frx":989A
         PICN            =   "frmListManualAtt.frx":98B6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdBrowse 
         Height          =   555
         Left            =   -74700
         TabIndex        =   40
         Top             =   8700
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "&Browse"
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
         MICON           =   "frmListManualAtt.frx":A948
         PICN            =   "frmListManualAtt.frx":A964
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdDownload 
         Height          =   555
         Left            =   -71160
         TabIndex        =   41
         Top             =   8700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   979
         BTYPE           =   14
         TX              =   "&Contoh File"
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
         MICON           =   "frmListManualAtt.frx":B9F6
         PICN            =   "frmListManualAtt.frx":BA12
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Index           =   1
         Left            =   1350
         TabIndex        =   47
         Top             =   480
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   91291651
         CurrentDate     =   40794
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_Group_Shift 
         Height          =   375
         Index           =   1
         Left            =   1350
         OleObjectBlob   =   "frmListManualAtt.frx":CAA4
         TabIndex        =   51
         Top             =   1290
         Width           =   1245
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Index           =   1
         Left            =   1350
         OleObjectBlob   =   "frmListManualAtt.frx":EA05
         TabIndex        =   52
         Top             =   930
         Width           =   1245
      End
      Begin prj_panji.vbButton cmdPrev 
         Height          =   285
         Index           =   1
         Left            =   2610
         TabIndex        =   55
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "<"
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
         MICON           =   "frmListManualAtt.frx":1096E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdRefresh 
         Height          =   495
         Index           =   1
         Left            =   3450
         TabIndex        =   56
         Top             =   390
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
         MICON           =   "frmListManualAtt.frx":1098A
         PICN            =   "frmListManualAtt.frx":109A6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdNext 
         Height          =   285
         Index           =   1
         Left            =   2970
         TabIndex        =   57
         Top             =   480
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   ">"
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
         MICON           =   "frmListManualAtt.frx":11A38
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdDelete 
         Height          =   465
         Index           =   0
         Left            =   -61500
         TabIndex        =   69
         Top             =   1290
         Width           =   1305
         _ExtentX        =   2302
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
         MICON           =   "frmListManualAtt.frx":11A54
         PICN            =   "frmListManualAtt.frx":11A70
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdEdit 
         Height          =   465
         Index           =   0
         Left            =   -62850
         TabIndex        =   70
         Top             =   1290
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
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
         MICON           =   "frmListManualAtt.frx":12B02
         PICN            =   "frmListManualAtt.frx":12B1E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdNew 
         Height          =   465
         Index           =   0
         Left            =   -64200
         TabIndex        =   71
         Top             =   1290
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
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
         MICON           =   "frmListManualAtt.frx":13BB0
         PICN            =   "frmListManualAtt.frx":13BCC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdDelete 
         Height          =   465
         Index           =   1
         Left            =   13470
         TabIndex        =   72
         Top             =   1290
         Width           =   1305
         _ExtentX        =   2302
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
         MICON           =   "frmListManualAtt.frx":14C5E
         PICN            =   "frmListManualAtt.frx":14C7A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdEdit 
         Height          =   465
         Index           =   1
         Left            =   12120
         TabIndex        =   73
         Top             =   1290
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
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
         MICON           =   "frmListManualAtt.frx":15D0C
         PICN            =   "frmListManualAtt.frx":15D28
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prj_panji.vbButton cmdNew 
         Height          =   465
         Index           =   1
         Left            =   10770
         TabIndex        =   74
         Top             =   1290
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   820
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
         MICON           =   "frmListManualAtt.frx":16DBA
         PICN            =   "frmListManualAtt.frx":16DD6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PERUSAHAAN"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   54
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "GRUP SHIFT"
         Height          =   195
         Index           =   1
         Left            =   330
         TabIndex        =   53
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TANGGAL"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   48
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Label1"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   -69000
         TabIndex        =   42
         Top             =   9030
         Visible         =   0   'False
         Width           =   5985
      End
      Begin VB.Label Label6 
         Caption         =   "Label1"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   -68880
         TabIndex        =   35
         Top             =   8970
         Visible         =   0   'False
         Width           =   5985
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "TANGGAL"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   11
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "GRUP SHIFT"
         Height          =   195
         Index           =   0
         Left            =   -74670
         TabIndex        =   10
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PERUSAHAAN"
         Height          =   195
         Index           =   0
         Left            =   -74820
         TabIndex        =   9
         Top             =   960
         Width           =   1110
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prj_panji.vbButton cmdExit 
      Height          =   705
      Left            =   14100
      TabIndex        =   16
      Top             =   10530
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
      MICON           =   "frmListManualAtt.frx":17E68
      PICN            =   "frmListManualAtt.frx":17E84
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "KEHADIRAN MANUAL"
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
      TabIndex        =   1
      Top             =   150
      Width           =   4845
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frmListManualAtt.frx":18F16
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16860
   End
End
Attribute VB_Name = "frm_list_manual_att"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Dim rs As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsGroupShift As New ADODB.Recordset
Dim rsshift As New ADODB.Recordset
Dim rsAtt As New ADODB.Recordset
Dim rsLogAtt As New ADODB.Recordset
Dim rsImportAtt As New ADODB.Recordset

Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Dim abstatus As Integer
Dim v_flag_present As Integer
Dim v_flag_duty As Integer

Dim vMode As Integer

Private Sub cmdDownload_Click()
Dim vFileName As String
Dim vSourceFile As String
    
    CommonDialog1.Filter = "XLS|*.xls"
    CommonDialog1.initDir = App.Path
    CommonDialog1.ShowSave
    
    vSourceFile = App.Path & "\files\FormatImportAttendance.xls"
    If Right(CommonDialog1.FileName, 4) <> ".xls" Then
        vFileName = CommonDialog1.FileName & ".xls"
    Else
        vFileName = CommonDialog1.FileName
    End If
           
    fso.GetFile vSourceFile
    
    fso.CopyFile vSourceFile, vFileName
    
    MsgBox "Format Absensi Berhasil Disimpan di " & vFileName, vbInformation, headerMSG
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click(Index As Integer)
    vMode = 0
    DTPicker1(SSTab1.Tab).Value = DTPicker1(SSTab1.Tab).Value + 1

    Call load_data_att(Index)
End Sub

Private Sub cmdPrev_Click(Index As Integer)
    vMode = 0
    DTPicker1(SSTab1.Tab).Value = DTPicker1(SSTab1.Tab).Value - 1

    Call load_data_att(Index)
End Sub

Private Sub cmdSearch_Click(Index As Integer)
    'validasi group_shift
    If Trim(TDBCombo_Group_Shift(SSTab1.Tab).Text) = "" Then
        MsgBox "Group Shift not selected!", vbOKOnly + vbInformation, headerMSG
        TDBCombo_Group_Shift(SSTab1.Tab).SetFocus
        Exit Sub
    End If
    
    vMode = 1
        
    Call days_func(DTPicker_from(SSTab1.Tab).Value, DTPicker_to(SSTab1.Tab).Value)
    Call load_data_att(Index)
    
'    TDBGrid_Att(SSTab1.Tab).FetchRowStyle = True
'    TDBGrid_Att(SSTab1.Tab).Refresh
End Sub

Private Sub Form_Load()
    DTPicker1(SSTab1.Tab).Value = Now
    SSTab1.Tab = 0
    
    Call load_data_company
    Call createGridKar
    
    DTPicker_from(SSTab1.Tab).Value = Now
    DTPicker_to(SSTab1.Tab).Value = Now

    timer1(SSTab1.Tab).Enabled = True

    cmdNew(SSTab1.Tab).Enabled = False
    cmdEdit(SSTab1.Tab).Enabled = False
    cmdDelete(SSTab1.Tab).Enabled = False
    cmdRefresh(SSTab1.Tab).Enabled = False
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBCombo_company(SSTab1.Tab).RowSource = rsCompany
End Sub

Private Sub load_data_group_shift()
    If rsGroupShift.State Then rsGroupShift.Close
    SQL = "select * from m_shift_group order by group_code"
    rsGroupShift.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    Set TDBCombo_Group_Shift(SSTab1.Tab).RowSource = rsGroupShift
End Sub

Private Sub load_data_shift()
    If rsshift.State Then rsshift.Close
    SQL = "select shift_code,shift_name,flag_day_over " & _
            "from m_shift where group_code = '" & TDBCombo_Group_Shift(SSTab1.Tab).Text & "' " & _
          "order by shift_name"
    rsshift.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    Set TDBGrid_Shift(SSTab1.Tab).DataSource = rsshift
End Sub

Public Sub load_data_att(Index As Integer)
Dim vParameter As String
    
    If Index = 0 Then
        vParameter = IIf(txt_nik(SSTab1.Tab).Text <> "", _
                        "a.employee_code = '" & txt_employee_code(SSTab1.Tab).Text & "' AND (DATE(b.att_date) BETWEEN '" & Format(DTPicker_from(SSTab1.Tab).Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to(SSTab1.Tab).Value, "yyyy-MM-dd") & "')", _
                        "(DATE(b.att_date) BETWEEN '" & Format(DTPicker_from(SSTab1.Tab).Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to(SSTab1.Tab).Value, "yyyy-MM-dd") & "')")
    Else
        vParameter = IIf(txt_nik(SSTab1.Tab).Text <> "", _
                        "a.employee_code = '" & txt_employee_code(SSTab1.Tab).Text & "' AND (DATE(b.att_date) BETWEEN '" & Format(DTPicker_from(SSTab1.Tab).Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to(SSTab1.Tab).Value, "yyyy-MM-dd") & "') AND (ISNULL(time_in) OR ISNULL(time_out))", _
                        "(DATE(b.att_date) BETWEEN '" & Format(DTPicker_from(SSTab1.Tab).Value, "yyyy-MM-dd") & "' AND '" & Format(DTPicker_to(SSTab1.Tab).Value, "yyyy-MM-dd") & "') AND (ISNULL(time_in) OR ISNULL(time_out))")
    End If
                        
    If rsAtt.State Then rsAtt.Close
    SQL = "select b.att_date,b.att_date _date,b.employee_code,a.nik,a.employee_name,b.status,e.absent_name,a.title_code,c.title_name, " & _
            "b.time_in,b.time_out,b.entry_date,a.division_code,d.division_name,b.description,b.shift_code,f.shift_name,b.enrollnumber " & _
          "FROM m_employee a JOIN h_attendance b ON a.employee_code = b.employee_code " & _
                "LEFT JOIN m_title c ON a.title_code = c.title_code " & _
                "LEFT JOIN m_division d ON a.division_code = d.division_code AND a.company_code = d.company_code " & _
                "LEFT JOIN m_absent_status e ON b.status = e.absent_code " & _
                "LEFT JOIN m_shift f ON b.shift_code = f.shift_code "

    If vMode = 0 Then
        SQL = SQL & _
                "where date(b.att_date) = '" & Format(DTPicker1(SSTab1.Tab).Value, "yyyy-MM-dd") & "' " & _
            "AND b.group_code = '" & TDBCombo_Group_Shift(SSTab1.Tab).Text & "' " & _
            "AND b.shift_code = '" & TDBGrid_Shift(SSTab1.Tab).Columns("shift_code").Value & "' ORDER BY b.att_date"
    Else
        SQL = SQL & _
                "where b.group_code = '" & TDBCombo_Group_Shift(SSTab1.Tab).Text & "' " & _
                "AND " & vParameter & " ORDER BY b.att_date"
    End If
    rsAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid_Att(SSTab1.Tab).DataSource = rsAtt

    cmdNew(SSTab1.Tab).Enabled = IIf(TDBCombo_company(SSTab1.Tab).Columns("company_code").Text = "", False, True)
    cmdEdit(SSTab1.Tab).Enabled = IIf(rsAtt.RecordCount = 0, False, True)
    cmdDelete(SSTab1.Tab).Enabled = IIf(rsAtt.RecordCount = 0, False, True)
    cmdRefresh(SSTab1.Tab).Enabled = IIf(rsAtt.RecordCount = 0, False, True)
End Sub

Private Sub load_data_log_att()
    If rsLogAtt.State Then rsLogAtt.Close
    SQL = "SELECT DISTINCT a.att_date, date(a.att_date) AS date_att, time(a.att_date) AS time_att, " & _
                "CASE WHEN a.flag_io = 0 THEN 'IN' ELSE 'OUT' END type " & _
            "FROM h_log_attendance a LEFT JOIN h_attendance b ON date(a.att_date) = date(b.att_date) AND a.employee_code = b.employee_code " & _
            "LEFT JOIN m_shift c ON b.shift_code = c.shift_code " & _
          "WHERE a.enrollnumber = '" & TDBGrid_Att(SSTab1.Tab).Columns("enrollnumber").Value & "' " & _
            "AND CASE WHEN c.flag_day_over = 1 THEN " & _
                    "(date(a.att_date) BETWEEN '" & Format(TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value, "yyyy-MM-dd") & "' AND ADDDATE('" & Format(TDBGrid_Att(SSTab1.Tab).Columns("tgl").Value, "yyyy-MM-dd") & "',1)) " & _
                "ELSE date(a.att_date) = '" & Format(TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value, "yyyy-MM-dd") & "' END " & _
          "ORDER BY att_date"
    rsLogAtt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid1(SSTab1.Tab).DataSource = rsLogAtt
End Sub

Private Sub showData()
Dim vTimeIn, vTimeOut As String
    
    SQL = "SELECT start_time, end_time FROM m_working_day WHERE group_code = '" & TDBCombo_Group_Shift(SSTab1.Tab).Text & "' " & _
            "AND shift_code = '" & TDBGrid_Shift(SSTab1.Tab).Columns("shift_code").Value & "'"
    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rscari.RecordCount > 0 Then
        vTimeIn = Format(rscari!start_time, "HH:mm")
        vTimeOut = Format(rscari!end_time, "HH:mm")
    Else
        vTimeIn = "00:00"
        vTimeOut = "00:00"
    End If
    rscari.Close

    If TDBGrid_Att(SSTab1.Tab).ApproxCount > 0 Then
        With frm_trans_att_man
            .DTPicker1.Value = IIf(TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value = "", TDBGrid_Att(SSTab1.Tab).Columns("tgl").Value, TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value)
            .txt_employee_code.Text = IIf(IsNull(TDBGrid_Att(SSTab1.Tab).Columns("employee_code").Value), txt_employee_code(SSTab1.Tab).Text, TDBGrid_Att(SSTab1.Tab).Columns("employee_code").Value)
            .txt_nik.Text = IIf(IsNull(TDBGrid_Att(SSTab1.Tab).Columns("nik").Value), txt_nik(SSTab1.Tab).Text, TDBGrid_Att(SSTab1.Tab).Columns("nik").Value)
            .txt_employee_name.Text = IIf(IsNull(TDBGrid_Att(SSTab1.Tab).Columns("employee_name").Value), txt_employee_name(SSTab1.Tab).Text, TDBGrid_Att(SSTab1.Tab).Columns("employee_name").Value)
            
            .TDBCombo_shift.Text = IIf(TDBGrid_Att(SSTab1.Tab).Columns("shift_code").Value = "", TDBGrid_Shift(SSTab1.Tab).Columns("shift_code").Value, TDBGrid_Att(SSTab1.Tab).Columns("shift_code").Value)
            .txt_shift_name.Text = IIf(TDBGrid_Att(SSTab1.Tab).Columns("shift_name").Value = "", TDBGrid_Shift(SSTab1.Tab).Columns("shift_name").Value, TDBGrid_Att(SSTab1.Tab).Columns("shift_name").Value)
            
            .TDBCombo_division.Text = IIf(IsNull(TDBGrid_Att(SSTab1.Tab).Columns("division_code").Value), "", TDBGrid_Att(SSTab1.Tab).Columns("division_code").Value)
            .txt_division_name.Text = IIf(IsNull(TDBGrid_Att(SSTab1.Tab).Columns("division_name").Value), "", TDBGrid_Att(SSTab1.Tab).Columns("division_name").Value)
            .txt_title_code.Text = IIf(IsNull(TDBGrid_Att(SSTab1.Tab).Columns("title_code").Value), "", TDBGrid_Att(SSTab1.Tab).Columns("title_code").Value)
            .txt_title_name.Text = IIf(IsNull(TDBGrid_Att(SSTab1.Tab).Columns("title_name").Value), "", TDBGrid_Att(SSTab1.Tab).Columns("title_name").Value)
            
            .TDBCombo_status.Text = IIf(TDBGrid_Att(SSTab1.Tab).Columns("status").Value = "", "P", TDBGrid_Att(SSTab1.Tab).Columns("status").Value)
            .txt_status_name.Text = IIf(TDBGrid_Att(SSTab1.Tab).Columns("absent_name").Value = "", "PRESENT", TDBGrid_Att(SSTab1.Tab).Columns("absent_name").Value)
            
            .DTPicker_from.Value = IIf(TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value = "", TDBGrid_Att(SSTab1.Tab).Columns("tgl").Value, TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value)
            .DTPicker_to.Value = IIf(TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value = "", TDBGrid_Att(SSTab1.Tab).Columns("tgl").Value, TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value)
            
            If TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value = "" Then
                .DTPicker_from.Enabled = True
                .DTPicker_to.Enabled = True
                
                .ttin.Value = vTimeIn
                .ttout.Value = vTimeOut
            Else
                .DTPicker_from.Enabled = False
                .DTPicker_to.Enabled = False
                
                If .TDBCombo_status.Text = "H" Or .TDBCombo_status.Text = "DT" Then
                .ttin.Value = Format(TDBGrid_Att(SSTab1.Tab).Columns("time_in").Value, "hh:mm")
                .ttout.Value = Format(TDBGrid_Att(SSTab1.Tab).Columns("time_out").Value, "hh:mm")
                Else
                    .ttin.Enabled = False
                    .ttout.Enabled = False
                    .ttin.Value = "00:00"
                    .ttout.Value = "00:00"
                End If
            End If
            
            .txt_description = IIf(IsNull(TDBGrid_Att(SSTab1.Tab).Columns("description").Value), "", TDBGrid_Att(SSTab1.Tab).Columns("description").Value)
            
            .chk_all_employee.Enabled = False
            
            If TDBGrid_Att(SSTab1.Tab).Columns("att_date").Value = "" Then
                .editTrans = False
            Else
                .editTrans = True
            End If
            .Show vbModal
        End With
    End If
End Sub

Private Sub newData()
    With frm_trans_att_man
        .DTPicker1.Value = DTPicker1(SSTab1.Tab).Value
        .DTPicker_from.Value = DTPicker1(SSTab1.Tab).Value
        .DTPicker_to.Value = DTPicker1(SSTab1.Tab).Value
        .TDBCombo_status.Text = "H"
        .txt_status_name.Text = "HADIR"
        .TDBCombo_shift.Text = TDBGrid_Shift(SSTab1.Tab).Columns("shift_code").Value
        .txt_shift_name.Text = TDBGrid_Shift(SSTab1.Tab).Columns("shift_name").Value
        .editTrans = False
        .Show vbModal
    End With
End Sub

'Private Sub Form_Resize()
'    Frame1.Width = Me.Width - 500
'    Frame1.Height = Me.Height - 2500
'    TDBGrid_Shift(SSTab1.Tab).Height = Frame1.Height - 400
'    TDBGrid_Shift(SSTab1.Tab).Width = Frame1.Width - (TDBGrid_Att(SSTab1.Tab).Width + 400)
'    TDBGrid_Att(SSTab1.Tab).Height = Frame1.Height - 400
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_list_manual_att = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Or SSTab1.Tab = 1 Then
        TDBCombo_company(SSTab1.Tab).Text = ""
        txt_company_name(SSTab1.Tab).Text = ""
        
        TDBCombo_Group_Shift(SSTab1.Tab).Text = ""
        txt_group_shift(SSTab1.Tab).Text = ""
        
        txt_nik(SSTab1.Tab).Text = ""
        txt_employee_code(SSTab1.Tab).Text = ""
        txt_employee_name(SSTab1.Tab).Text = ""
        
        DTPicker1(SSTab1.Tab).Value = Now
        DTPicker_from(SSTab1.Tab).Value = Now
        DTPicker_to(SSTab1.Tab).Value = Now
        
'        TDBGrid_Att(SSTab1.Tab).Columns.Remove 0
'        TDBGrid_Shift(SSTab1.Tab).Columns.Remove 0
'        TDBGrid1(SSTab1.Tab).Columns.Remove 0
        
        Call load_data_company
        Call createGridKar
    
        timer1(SSTab1.Tab).Enabled = True
        cmdNew(SSTab1.Tab).Enabled = False
        cmdEdit(SSTab1.Tab).Enabled = False
        cmdDelete(SSTab1.Tab).Enabled = False
        cmdRefresh(SSTab1.Tab).Enabled = False
    ElseIf SSTab1.Tab = 2 Then
        LynxGrid1.ClearAll
        createGrid
    ElseIf SSTab1.Tab = 3 Then
        LynxGrid2.ClearAll
        createGridM
    End If
End Sub

Private Sub TDBGrid_Att_DblClick(Index As Integer)
    showData
End Sub

Private Sub TDBGrid_Att_KeyPress(KeyAscii As Integer, Index As Integer)
    If KeyAscii = 13 Then
        showData
    End If
End Sub

Private Sub TDBGrid_Att_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If Not IIf(IsNull(TDBGrid_Att(SSTab1.Tab).Bookmark), 0, TDBGrid_Att(SSTab1.Tab).Bookmark) > 0 Then
        Set TDBGrid1(SSTab1.Tab).DataSource = Nothing
        Exit Sub
    End If
    
    Call load_data_log_att
End Sub

Private Sub TDBGrid_Shift_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    Call load_data_att(Index)
End Sub

Private Sub TDBCombo_company_ItemChange(Index As Integer)
    If TDBCombo_company(SSTab1.Tab).ApproxCount > 0 Then
        TDBCombo_company(SSTab1.Tab).Text = TDBCombo_company(SSTab1.Tab).Columns("company_code").Value
        txt_company_name(SSTab1.Tab).Text = TDBCombo_company(SSTab1.Tab).Columns("company_name").Value

        Call load_data_group_shift
    End If
End Sub

Private Sub TDBCombo_Group_Shift_ItemChange(Index As Integer)
    If TDBCombo_Group_Shift(SSTab1.Tab).ApproxCount > 0 Then
        TDBCombo_Group_Shift(SSTab1.Tab).Text = TDBCombo_Group_Shift(SSTab1.Tab).Columns("group_code").Value
        txt_group_shift(SSTab1.Tab).Text = TDBCombo_Group_Shift(SSTab1.Tab).Columns("group_name").Value

        Call load_data_shift
        Call load_data_att(Index)
    End If
End Sub

Private Sub cmdNew_Click(Index As Integer)
    If TDBGrid_Shift(SSTab1.Tab).ApproxCount = 0 Then
        MsgBox "Shift Code Tidak Sesuai..." & Chr(13) & "Please Check Your Transaction Again...", vbExclamation, headerMSG
        Exit Sub
    End If
    newData
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    showData
End Sub

Private Sub cmdRefresh_Click(Index As Integer)
    Call load_data_att(Index)
End Sub

Private Sub Timer1_Timer(Index As Integer)
    timer1(SSTab1.Tab).Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company(SSTab1.Tab), txt_company_name(SSTab1.Tab))
End Sub

Private Sub clear_filter()
    For Each Col In TDBGrid_Att(SSTab1.Tab).Columns
        Col.FilterText = ""
    Next Col
    rsAtt.Filter = adFilterNone
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

Private Sub TDBGrid_Att_FilterChange(Index As Integer)
On Error GoTo Err

    Dim i As Integer
    
    Set Cols = TDBGrid_Att(SSTab1.Tab).Columns
    i = TDBGrid_Att(SSTab1.Tab).Col
    TDBGrid_Att(SSTab1.Tab).HoldFields
    
    rsAtt.Filter = getFilter()
    TDBGrid_Att(SSTab1.Tab).Col = i
    TDBGrid_Att(SSTab1.Tab).EditActive = True
    
    TDBGrid_Att(SSTab1.Tab).SelStart = Len(TDBGrid_Att(SSTab1.Tab).Columns(i).FilterText)
    If TDBGrid_Att(SSTab1.Tab).ApproxCount < 1 Then
        Call clear_filter
        TDBGrid_Att(SSTab1.Tab).Col = i
    End If

    Exit Sub

Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub cmdDelete_Click(Index As Integer)
Dim i As Integer
Dim item
    
On Error GoTo Err
    If Not TDBGrid_Att(SSTab1.Tab).ApproxCount > 0 Then
        Exit Sub
    End If
        
    Set SelBks = TDBGrid_Att(SSTab1.Tab).SelBookmarks
    i = MsgBox("Apakah Anda Yakin Untu Menghapus " _
        & SelBks.Count & " Data Kehadiran ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
                
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        CnG.Execute "DELETE FROM h_attendance where employee_code = '" & TDBGrid_Att(SSTab1.Tab).Columns("employee_code").CellText(item) & "' " _
            & "and att_date = '" & Format(TDBGrid_Att(SSTab1.Tab).Columns("att_date").CellText(item), "yyyy-MM-dd HH:mm:ss") & "'"
                
    Next
    CnG.CommitTrans
    Call load_data_att(Index)
    MsgBox i & " Data Kehadiran Berhasil Dihapus...", vbInformation, headerMSG
    
    '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
    SQL = "Update temp_sal_proses set salary_proses = 0 where company_code = '" & TDBCombo_company(SSTab1.Tab).Text & "'"
    CnG.Execute SQL
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub createGrid()
   With LynxGrid1
      .AddColumn "TGL. ABSENSI", 1200, lgAlignRightCenter, lgDate, "yyyy-MM-dd", , , , , , True
      .AddColumn "KODE KARY.", 1300, lgAlignCenterCenter, , , , , , , True
      .AddColumn "NAMA KARY.", 2500, , , , , , , , True
      .AddColumn "KODE GROUP", 800, lgAlignCenterCenter, , , , , , , , True
      .AddColumn "KODE SHIFT", 800, , , , , , , , , True
      .AddColumn "STS ABSENSI", 800, lgAlignCenterCenter, , , , , , , True
      .AddColumn "MASUK", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "PULANG", 1000, lgAlignCenterCenter, lgDate, "hh:mm", , , , , True
      .AddColumn "DESKRIPSI", 3000, lgAlignCenterCenter, , , , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub createGridM()
   With LynxGrid2
      .AddColumn "KODE KARY.", 1200, lgAlignCenterCenter, , , , , , , True
      .AddColumn "TGL. ABSENSI", 2000, lgAlignCenterCenter, lgDate, "yyyy-MM-dd hh:mm:ss", , , , , , True
      .AddColumn "TIPE LOG", 1000, lgAlignCenterCenter, , , , , , , , True
      .BackColorBkg = &HFCE1CB
      .Redraw = True
   End With
    
End Sub

Private Sub cmdBrowse_Click()
    CommonDialog1.Filter = "XLS|*.xls"
    CommonDialog1.initDir = App.Path
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        Call fill_grid_excel_m(CommonDialog1.FileName)
    End If
End Sub

Private Sub cmdBrowae_M_Click()
    CommonDialog1.Filter = "XLS|*.xls"
    CommonDialog1.initDir = App.Path
    CommonDialog1.ShowOpen
    
    If CommonDialog1.FileName <> "" Then
        Call fill_grid_excelM(CommonDialog1.FileName)
    End If
End Sub

Private Sub fill_grid_excel_m(str_file_name As String)
On Error GoTo Err
    Dim strWorksheet As String
    strWorksheet = "data_attendance"
    
    Adodc1.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & str_file_name & ";Extended Properties=Excel 8.0"
    
    Adodc1.RecordSource = "select * from [" & strWorksheet & "$]"
    Adodc1.Refresh
    LynxGrid1.Redraw = False
    LynxGrid1.Clear
    With Adodc1.Recordset
        If .RecordCount > 0 Then
            Me.MousePointer = vbHourglass
            .MoveFirst
            While Not .EOF
                LynxGrid1.AddItem .Fields(0) & vbTab & .Fields(1) _
                        & vbTab & .Fields(2) & vbTab & .Fields(3) _
                        & vbTab & .Fields(4) & vbTab & .Fields(5) _
                        & vbTab & .Fields(6) & vbTab & .Fields(7) _
                        & vbTab & .Fields(8)
                .MoveNext
            Wend
            Me.MousePointer = vbNormal
        End If
    End With
    LynxGrid1.Redraw = True
    Exit Sub
    
Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub fill_grid_excelM(str_file_name As String)
On Error GoTo Err
    Dim strWorksheet As String
    strWorksheet = "data_log"
    
    Adodc2.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=" _
    & str_file_name & ";Extended Properties=Excel 8.0"
    
    Adodc2.RecordSource = "select * from [" & strWorksheet & "$]"
    Adodc2.Refresh
    LynxGrid2.Redraw = False
    LynxGrid2.Clear
    With Adodc2.Recordset
        If .RecordCount > 0 Then
            Me.MousePointer = vbHourglass
            .MoveFirst
            While Not .EOF
                LynxGrid2.AddItem .Fields(0) & vbTab & .Fields(1) & vbTab & .Fields(2)
                .MoveNext
            Wend
            Me.MousePointer = vbNormal
        End If
    End With
    LynxGrid2.Redraw = True
    Exit Sub
    
Err:
MsgBox Err.Description, vbExclamation, "Message Error!"
End Sub

Private Sub cmdImport_Click()
Dim aa As Integer
Dim rsnumber As New ADODB.Recordset
Dim nourut As Long
Dim v_employee_code As String
Dim SQL As String
Dim waktu_in As String, waktu_out As String
Dim time_in As String, time_out As String
Dim start_time As String, end_time As String

'On Error Resume Next
    With LynxGrid1
        If .Rows > 0 Then
            ProgressBar1.Visible = True
            Label5.Visible = True
            ProgressBar1.Max = .Rows - 1
            ProgressBar1.Value = 0
            
            DoEvents
            
            For aa = 0 To .Rows - 1
                
                ProgressBar1.Value = aa
                Label5.Caption = .CellText(aa, 0) & " - " & .CellText(aa, 1)
                             
                SQL = "SELECT employee_code FROM m_employee WHERE nik = '" & .CellText(aa, 1) & "' " _
                            & "AND flag_active <> 0"
                rsnumber.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsnumber.RecordCount > 0 Then
                    v_employee_code = rsnumber!employee_code
                    
                    SQL = "DELETE FROM h_attendance " & _
                          "WHERE att_date = '" & .CellText(aa, 0) & "' " & _
                            "AND employee_code = '" & .CellText(aa, 1) & "'"
                    CnG.Execute SQL
                    
                    SQL = "SELECT start_time, end_time FROM m_shift " & _
                      "WHERE group_code = '" & .CellText(aa, 3) & "' " & _
                        "AND shift_code = '" & .CellText(aa, 4) & "'"
                    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rs.RecordCount > 0 Then
                        waktu_in = Format(.CellText(aa, 0), "yyyy-MM-dd") & " " & Format(rs!start_time, "hh:mm") & ":00"
                        waktu_out = Format(.CellText(aa, 0), "yyyy-MM-dd") & " " & Format(rs!end_time, "hh:mm") & ":00"
                    Else
                        waktu_in = Format(.CellText(aa, 0), "yyyy-MM-dd") & " 07:00:00"
                        waktu_out = Format(.CellText(aa, 0), "yyyy-MM-dd") & " 15:00:00"
                    End If
                    rs.Close
                    
                    '+++++++++++++++++++++MENCARI TANGGAL BATAS JAM MASUK,ISTIRAHAT,KELUAR++++++++
                    SQL = "Select CAST(concat('" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "',' ', time(start_time)) as datetime) start_time," _
                            & "CAST(concat('" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "',' ', time(end_time)) as datetime) end_time," _
                            & "now() tglserver " _
                            & "from m_shift where shift_code = '" & .CellText(aa, 4) & "' " & _
                                "AND group_code = '" & .CellText(aa, 3) & "'"
                '        & "from m_shift where shift_code = '" & txtkdshift.Text & "'"
                    rs.Open SQL, CnG, adOpenDynamic, adLockReadOnly
                    
                    start_time = Format(rs!start_time, "yyyy-MM-dd hh:mm:ss")
                    If frm_list_manual_att.TDBGrid_Shift(SSTab1.Tab).Columns("flag_day_over").Value = 1 Then
                        end_time = Format(DateAdd("d", 1, rs!end_time), "yyyy-MM-dd hh:mm:ss")
                    Else
                        end_time = Format(rs!end_time, "yyyy-MM-dd hh:mm:ss")
                    End If
                    
                    time_in = Format(.CellText(aa, 0), "yyyy-MM-dd") & " " & Format(.CellText(aa, 6), "hh:mm:ss")
                    If Format(.CellText(aa, 7), "hh:mm:ss") <= Format(.CellText(aa, 6), "hh:mm:ss") Then
                        time_out = Format(DateAdd("d", 1, .CellText(aa, 0)), "yyyy-MM-dd") & " " & Format(.CellText(aa, 7), "hh:mm") & ":00"
                    Else
                        time_out = Format(.CellText(aa, 0), "yyyy-MM-dd") & " " & Format(.CellText(aa, 7), "hh:mm") & ":00"
                    End If
                    
                    rs.Close
    
                    If .CellText(aa, 0) <> "" And .CellText(aa, 1) <> "" Then
                        Select Case .CellText(aa, 5)
                        Case "H"
                            SQL = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & v_employee_code & "'"
                            CnG.Execute SQL
                            
                            SQL = "INSERT INTO h_attendance (employee_code,att_date,group_code,shift_code,status," & _
                                      "shift_number,start_time,end_time,time_in,time_out,description,entry_date,userinput," & _
                                      "absent_status,flag_present,flag_duty) " & _
                                    "VALUES (" & _
                                      "'" & v_employee_code & "','" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "','" & .CellText(aa, 3) & "'," & _
                                      "'" & .CellText(aa, 4) & "','" & .CellText(aa, 5) & "',1,'" & start_time & "','" & end_time & "'," & _
                                      "'" & time_in & "','" & time_out & "'," & _
                                      "'" & .CellText(aa, 8) & "',now(),'" & LOGIN_NAME & "',0,1,0)"
                            CnG.Execute SQL
                        Case "A", "L", "I", "S", "CT"
                            If .CellText(aa, 5) = "A" Then
                                abstatus = 2
                            ElseIf .CellText(aa, 5) = "L" Then
                                abstatus = 6
                            ElseIf .CellText(aa, 5) = "I" Then
                                abstatus = 0
                            ElseIf .CellText(aa, 5) = "S" Then
                                abstatus = 1
                            ElseIf .CellText(aa, 5) = "CT" Then
                                abstatus = 3
                            End If
                            
                            SQL = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & v_employee_code & "'"
                            CnG.Execute SQL
                            
                            SQL = "INSERT INTO h_attendance (att_date,employee_code,group_code,shift_code,status,flag_present,absent_status," & _
                                    "description,entry_date,userinput,shift_number) " & _
                                  "VALUES (" & _
                                    "'" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "','" & v_employee_code & "','" & .CellText(aa, 3) & "'," & _
                                    "'" & .CellText(aa, 4) & "','" & .CellText(aa, 5) & "',0,'" & abstatus & "'," & _
                                    "'" & .CellText(aa, 8) & "',now(),'" & LOGIN_NAME & "',1)"
                            CnG.Execute SQL
                            
                            Call generate_summary_leave(Format(.CellText(aa, 0), "yyyy-MM-dd"))
                        Case "T"
                            SQL = "DELETE FROM h_attendance WHERE date(att_date) = '" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "' " & _
                                    "AND employee_code = '" & v_employee_code & "'"
                            CnG.Execute SQL
                                
                            SQL = "INSERT INTO h_attendance (att_date,employee_code,group_code,shift_code,status,time_in,time_out," & _
                                    "flag_present,flag_duty,absent_status,description,entry_date,userinput,shift_number) " & _
                                  "VALUES (" & _
                                    "'" & Format(.CellText(aa, 0), "yyyy-MM-dd") & "','" & v_employee_code & "','" & .CellText(aa, 3) & "'," & _
                                    "'" & .CellText(aa, 4) & "','" & .CellText(aa, 5) & "'," & _
                                    "'" & waktu_in & "','" & waktu_out & "'," & _
                                    "1,1,0,'" & .CellText(aa, 8) & "',now(),'" & LOGIN_NAME & "',1)"
                            CnG.Execute SQL
                        End Select
                    End If
                End If
                rsnumber.Close
                
                
                DoEvents
            Next
            MsgBox "Import Data Berhasil...", vbInformation, headerMSG
            
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            SQL = "Update temp_sal_proses set salary_proses = 0"
            CnG.Execute SQL
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            ProgressBar1.Visible = False
            Label5.Visible = False
        End If
    End With
End Sub

Private Sub cmdImport_M_Click()
Dim aa As Long
Dim rsnumber As New ADODB.Recordset
Dim nourut As Long
Dim v_employee_code As String
Dim SQL As String

On Error Resume Next
    With LynxGrid2
        If .Rows > 0 Then
            ProgressBar2.Visible = True
            Label6.Visible = True
            ProgressBar2.Max = .Rows - 1
            ProgressBar2.Value = 0
            
            DoEvents
            
            For aa = 0 To .Rows - 1
                
                ProgressBar2.Value = aa
                Label6.Caption = .CellText(aa, 0) & " - " & .CellText(aa, 1)
                             
                SQL = "SELECT a.employee_code FROM m_employee a JOIN m_enroll_link b ON a.employee_code = b.employee_code " & _
                      "WHERE b.enrollnumber = '" & .CellText(aa, 0) & "' " _
                            & "AND a.flag_active <> 0 LIMIT 1"
                rsnumber.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                
                If rsnumber.RecordCount > 0 Then
                    v_employee_code = rsnumber!employee_code
                    
                    SQL = "DELETE a FROM h_attendance a JOIN m_enroll_link b ON a.enrollnumber = b.enrollnumber " _
                             & "WHERE a.att_date = '" & Format(.CellText(aa, 1), "yyyy-MM-dd HH:mm:ss") & "' " _
                                & "AND a.enrollnumber = '" & .CellText(aa, 0) & "'"
                    CnG.Execute SQL
                    
                    SQL = "DELETE a FROM h_log_attendance a JOIN m_enroll_link b ON a.enrollnumber = b.enrollnumber " _
                            & "WHERE a.att_date = '" & Format(.CellText(aa, 1), "yyyy-MM-dd HH:mm:ss") & "' " _
                                & "AND a.enrollnumber = '" & .CellText(aa, 0) & "'"
                    CnG.Execute SQL
                    
                    SQL = "SELECT a.* FROM h_log_attendance a JOIN m_enroll_link b ON a.enrollnumber = b.enrollnumber " & _
                          "WHERE a.enrollnumber = '" & .CellText(aa, 0) & "' " & _
                            "AND a.att_date = '" & Format(.CellText(aa, 1), "yyyy-MM-dd HH:mm:ss") & "'"
                    rscari.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
                    
                    If rscari.RecordCount = 0 Then
                        SQL = "INSERT INTO h_log_attendance(att_date,enrollnumber,employee_code,verifymode,flag_io) " & _
                                "VALUES ( " & _
                                "'" & Format(.CellText(aa, 1), "yyyy-MM-dd HH:mm:ss") & "','" & .CellText(aa, 0) & "'," & _
                                    "'" & v_employee_code & "',0,'" & IIf(.CellText(aa, 2) = "I", 0, 1) & "')"
                        CnG.Execute SQL
                    End If
                    rscari.Close
                End If
                rsnumber.Close
                
                
                DoEvents
            Next
            MsgBox "Import Data Success...!!!", vbInformation, headerMSG
            
            '+++++++++++++++++++++++++++++++++ Update Temp Salary Proses ++++++++++++++
            SQL = "Update temp_sal_proses set salary_proses = 0"
            CnG.Execute SQL
            '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

            ProgressBar2.Visible = False
            Label6.Visible = False
        End If
    End With
End Sub

Public Sub generate_summary_leave(ByRef dtp1 As DTPicker)
    CnG.BeginTrans
        CnG.Execute "call spg_leave_periode ('" & Format(dtp1, "yyyy-MM-dd HH:mm:ss") & "')"
    CnG.CommitTrans
End Sub

Private Sub createGridKar()
   With LynxGrid3
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
        LynxGrid3.Clear
        If LOGIN_LEVEL = 100 Then
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE b.company_code = '" & TDBCombo_company(SSTab1.Tab).Text & "' " _
                        & "AND (a.nik LIKE '%" & txt_nik(SSTab1.Tab).Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_nik(SSTab1.Tab).Text & "%') " _
                        & "AND a.flag_active <> 0"
        Else
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE b.company_code = '" & TDBCombo_company(SSTab1.Tab).Text & "' " _
                        & "AND (a.nik LIKE '%" & txt_nik(SSTab1.Tab).Text & "%' " _
                        & "OR a.employee_name LIKE '%" & txt_nik(SSTab1.Tab).Text & "%') " _
                        & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                        & "ORDER BY a.employee_name ASC"

        End If
        
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        If rs.RecordCount > 0 Then
            LynxGrid3.Redraw = False
            rs.MoveFirst
            While Not rs.EOF
                LynxGrid3.AddItem rs!nik & vbTab & rs!EMPLOYEE_NAME _
                                & vbTab & rs!division_code & vbTab & rs!division_name _
                                & vbTab & rs!title_code & vbTab & rs!title_name _
                                & vbTab & rs!employee_code
                rs.MoveNext
            Wend
            LynxGrid3.Redraw = True
            If rs.RecordCount = 1 Then
                rs.MoveFirst
                txt_employee_code(SSTab1.Tab).Text = rs!employee_code
                txt_employee_name(SSTab1.Tab).Text = rs!EMPLOYEE_NAME
                txt_nik(SSTab1.Tab).Text = rs!nik
'                TDBCombo1.SetFocus
            Else
                LynxGrid3.Visible = True
                LynxGrid3.SetFocus
            End If
        Else
            
        End If
        rs.Close
    Else
        If LynxGrid3.Rows > 0 Then
            txt_nik(SSTab1.Tab).Text = LynxGrid3.CellText(LynxGrid2.Row, 0)
            txt_employee_name(SSTab1.Tab).Text = LynxGrid3.CellText(LynxGrid2.Row, 1)
            txt_employee_code(SSTab1.Tab).Text = LynxGrid3.CellText(LynxGrid2.Row, 6)
        End If
        LynxGrid3.Visible = False
    End If
End Sub

Private Sub LynxGrid3_DblClick()
    isiGridKar (2)
End Sub

Private Sub LynxGrid3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        LynxGrid3.Visible = False
    End If
    If KeyAscii = 13 Then
        isiGridKar (2)
    End If
End Sub

Private Sub LynxGrid3_LostFocus()
    LynxGrid2.Visible = False
End Sub

Private Sub txt_nik_Change(Index As Integer)
    If txt_nik(SSTab1.Tab).Text = "" Then
        txt_employee_code(SSTab1.Tab).Text = ""
        txt_employee_name(SSTab1.Tab).Text = ""
    End If
End Sub

Private Sub txt_nik_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        isiGridKar (1)
    End If
End Sub

Private Sub cmdBrowse_Emp_Click(Index As Integer)
    isiGridKar (1)
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

