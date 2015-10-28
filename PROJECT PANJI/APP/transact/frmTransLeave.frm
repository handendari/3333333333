VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_trans_leave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CUTI"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12795
   Icon            =   "frmTransLeave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   12795
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame fra_company 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   150
      TabIndex        =   30
      Top             =   690
      Width           =   8235
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3210
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   31
         Top             =   120
         Width           =   3855
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   1440
         OleObjectBlob   =   "frmTransLeave.frx":058A
         TabIndex        =   32
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "PERUSAHAAN"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   180
         Width           =   1110
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7365
      Left            =   150
      TabIndex        =   0
      Top             =   1320
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   12991
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "CUTI KARY."
      TabPicture(0)   =   "frmTransLeave.frx":24F0
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGrid_Emp"
      Tab(0).Control(1)=   "fra_entry_Emp"
      Tab(0).Control(2)=   "frmTombol"
      Tab(0).Control(3)=   "LynxGrid2"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "CUTI BERSAMA"
      TabPicture(1)   =   "frmTransLeave.frx":250C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TDBGrid_Gen"
      Tab(1).Control(1)=   "fra_entry_Gen"
      Tab(1).Control(2)=   "Frame1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "REKAP CUTI"
      TabPicture(2)   =   "frmTransLeave.frx":2528
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "TDBGrid_Sum"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin prj_panji.LynxGrid LynxGrid2 
         Height          =   3465
         Left            =   -73170
         TabIndex        =   42
         Top             =   3780
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
      Begin VB.Frame Frame2 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   120
         TabIndex        =   35
         Top             =   5880
         Width           =   12255
         Begin prj_panji.vbButton cmdLoad_Sum 
            Height          =   705
            Left            =   7770
            TabIndex        =   37
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Load"
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
            MICON           =   "frmTransLeave.frx":2544
            PICN            =   "frmTransLeave.frx":2560
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
      Begin VB.Frame Frame1 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   23
         Top             =   5910
         Width           =   12285
         Begin prj_panji.vbButton cmdNew_Gen 
            Height          =   705
            Left            =   600
            TabIndex        =   25
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
            MICON           =   "frmTransLeave.frx":35F2
            PICN            =   "frmTransLeave.frx":360E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Gen 
            Height          =   705
            Left            =   1620
            TabIndex        =   26
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
            MICON           =   "frmTransLeave.frx":46A0
            PICN            =   "frmTransLeave.frx":46BC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Gen 
            Height          =   705
            Left            =   2640
            TabIndex        =   27
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
            MICON           =   "frmTransLeave.frx":574E
            PICN            =   "frmTransLeave.frx":576A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Gen 
            Height          =   705
            Left            =   3660
            TabIndex        =   28
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
            MICON           =   "frmTransLeave.frx":67FC
            PICN            =   "frmTransLeave.frx":6818
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Gen 
            Height          =   705
            Left            =   4680
            TabIndex        =   29
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
            MICON           =   "frmTransLeave.frx":78AA
            PICN            =   "frmTransLeave.frx":78C6
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
      Begin VB.Frame fra_entry_Gen 
         Height          =   2805
         Left            =   -74880
         TabIndex        =   17
         Top             =   3030
         Width           =   12285
         Begin VB.CommandButton Command1 
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
            TabIndex        =   19
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_desc 
            Appearance      =   0  'Flat
            Height          =   1035
            Left            =   5160
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   18
            Top             =   1110
            Width           =   4695
         End
         Begin MSComCtl2.DTPicker DTPicker_date_general_leave 
            Height          =   315
            Left            =   5160
            TabIndex        =   20
            Top             =   750
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   97124355
            CurrentDate     =   39270
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN*"
            Height          =   195
            Left            =   3825
            TabIndex        =   22
            Top             =   1110
            Width           =   1170
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TANGGAL*"
            Height          =   195
            Left            =   4170
            TabIndex        =   21
            Top             =   750
            Width           =   825
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   10
         Top             =   5880
         Width           =   12285
         Begin prj_panji.vbButton cmdNew_Emp 
            Height          =   705
            Left            =   630
            TabIndex        =   12
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
            MICON           =   "frmTransLeave.frx":8958
            PICN            =   "frmTransLeave.frx":8974
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Emp 
            Height          =   705
            Left            =   1650
            TabIndex        =   13
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
            MICON           =   "frmTransLeave.frx":9A06
            PICN            =   "frmTransLeave.frx":9A22
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Emp 
            Height          =   705
            Left            =   2670
            TabIndex        =   14
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
            MICON           =   "frmTransLeave.frx":AAB4
            PICN            =   "frmTransLeave.frx":AAD0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Emp 
            Height          =   705
            Left            =   3690
            TabIndex        =   15
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
            MICON           =   "frmTransLeave.frx":BB62
            PICN            =   "frmTransLeave.frx":BB7E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Emp 
            Height          =   705
            Left            =   4710
            TabIndex        =   16
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
            MICON           =   "frmTransLeave.frx":CC10
            PICN            =   "frmTransLeave.frx":CC2C
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
      Begin VB.Frame fra_entry_Emp 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   1
         Top             =   3030
         Width           =   12285
         Begin VB.TextBox txt_employee_code 
            Height          =   315
            Left            =   3510
            TabIndex        =   41
            Top             =   450
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txt_employee_name 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   1710
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   40
            Top             =   810
            Width           =   3495
         End
         Begin VB.TextBox txt_nik 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1710
            MaxLength       =   10
            TabIndex        =   39
            Top             =   420
            Width           =   1335
         End
         Begin VB.ComboBox cbo_date_to 
            Height          =   315
            ItemData        =   "frmTransLeave.frx":DCBE
            Left            =   9240
            List            =   "frmTransLeave.frx":DCC8
            TabIndex        =   3
            Text            =   "..."
            Top             =   390
            Width           =   1095
         End
         Begin VB.TextBox txt_description 
            Appearance      =   0  'Flat
            Height          =   1395
            Left            =   7440
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   870
            Width           =   4695
         End
         Begin MSComCtl2.DTPicker DTPicker_date_from 
            Height          =   315
            Left            =   7440
            TabIndex        =   4
            Top             =   390
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   97124355
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_date_to 
            Height          =   315
            Left            =   10440
            TabIndex        =   5
            Top             =   390
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "yyyy-MM-dd"
            Format          =   97124355
            CurrentDate     =   39270
         End
         Begin prj_panji.vbButton cmdBrowse 
            Height          =   315
            Left            =   3090
            TabIndex        =   43
            Top             =   420
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
            MICON           =   "frmTransLeave.frx":DCD6
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
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TANGGAL*"
            Height          =   195
            Left            =   6450
            TabIndex        =   9
            Top             =   390
            Width           =   825
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN*"
            Height          =   195
            Left            =   6105
            TabIndex        =   8
            Top             =   870
            Width           =   1170
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE KARY.*"
            Height          =   195
            Left            =   630
            TabIndex        =   7
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA KARY."
            Height          =   195
            Left            =   645
            TabIndex        =   6
            Top             =   840
            Width           =   990
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Emp 
         Height          =   5385
         Left            =   -74880
         TabIndex        =   11
         Top             =   420
         Width           =   12285
         _ExtentX        =   21669
         _ExtentY        =   9499
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
         Columns(2).Caption=   "KODE DEPT."
         Columns(2).DataField=   "department_code"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "DEPARTEMEN"
         Columns(3).DataField=   "department_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DIV. CODE"
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
         Columns(8).Caption=   "NO"
         Columns(8).DataField=   "leave_number"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "EMPLOYEE CODE"
         Columns(9).DataField=   "employee_code"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "KODE KARY."
         Columns(10).DataField=   "nik"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "NAMA KARY."
         Columns(11).DataField=   "employee_name"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "TGL AWAL"
         Columns(12).DataField=   "leave_date_from"
         Columns(12).NumberFormat=   "yyyy-MM-dd"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   4
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "S/D"
         Columns(13).DataField=   "flag_date_to"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "TGL AKHIR"
         Columns(14).DataField=   "leave_date_to"
         Columns(14).NumberFormat=   "FormatText Event"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "KETERANGAN"
         Columns(15).DataField=   "description"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   16
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=16"
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
         Splits(0)._ColumnProps(29)=   "Column(4).Width=1852"
         Splits(0)._ColumnProps(30)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(4)._WidthInPix=1773"
         Splits(0)._ColumnProps(32)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(33)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(34)=   "Column(5).Width=2487"
         Splits(0)._ColumnProps(35)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._WidthInPix=2408"
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
         Splits(0)._ColumnProps(52)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=516"
         Splits(0)._ColumnProps(54)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(55)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(56)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(57)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(58)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=516"
         Splits(0)._ColumnProps(61)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(62)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(63)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(64)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(65)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(66)=   "Column(10).AllowSizing=0"
         Splits(0)._ColumnProps(67)=   "Column(10)._ColStyle=516"
         Splits(0)._ColumnProps(68)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(69)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(70)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(71)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(72)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(73)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(74)=   "Column(11).AllowSizing=0"
         Splits(0)._ColumnProps(75)=   "Column(11)._ColStyle=516"
         Splits(0)._ColumnProps(76)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(77)=   "Column(11).AllowFocus=0"
         Splits(0)._ColumnProps(78)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(79)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(80)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(82)=   "Column(12).AllowSizing=0"
         Splits(0)._ColumnProps(83)=   "Column(12)._ColStyle=516"
         Splits(0)._ColumnProps(84)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(85)=   "Column(12).AllowFocus=0"
         Splits(0)._ColumnProps(86)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(87)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(88)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(89)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(90)=   "Column(13).AllowSizing=0"
         Splits(0)._ColumnProps(91)=   "Column(13)._ColStyle=516"
         Splits(0)._ColumnProps(92)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(93)=   "Column(13).AllowFocus=0"
         Splits(0)._ColumnProps(94)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(95)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(96)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(97)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(98)=   "Column(14).AllowSizing=0"
         Splits(0)._ColumnProps(99)=   "Column(14)._ColStyle=516"
         Splits(0)._ColumnProps(100)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(101)=   "Column(14).AllowFocus=0"
         Splits(0)._ColumnProps(102)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(103)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(104)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(105)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(106)=   "Column(15).AllowSizing=0"
         Splits(0)._ColumnProps(107)=   "Column(15)._ColStyle=516"
         Splits(0)._ColumnProps(108)=   "Column(15).Visible=0"
         Splits(0)._ColumnProps(109)=   "Column(15).AllowFocus=0"
         Splits(0)._ColumnProps(110)=   "Column(15).Order=16"
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
         Splits(1)._ColumnProps(0)=   "Columns.Count=16"
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
         Splits(1)._ColumnProps(65)=   "Column(8).Width=1429"
         Splits(1)._ColumnProps(66)=   "Column(8).DividerColor=0"
         Splits(1)._ColumnProps(67)=   "Column(8)._WidthInPix=1349"
         Splits(1)._ColumnProps(68)=   "Column(8)._ColStyle=514"
         Splits(1)._ColumnProps(69)=   "Column(8).Order=9"
         Splits(1)._ColumnProps(70)=   "Column(9).Width=2725"
         Splits(1)._ColumnProps(71)=   "Column(9).DividerColor=0"
         Splits(1)._ColumnProps(72)=   "Column(9)._WidthInPix=2646"
         Splits(1)._ColumnProps(73)=   "Column(9)._ColStyle=516"
         Splits(1)._ColumnProps(74)=   "Column(9).Visible=0"
         Splits(1)._ColumnProps(75)=   "Column(9).Order=10"
         Splits(1)._ColumnProps(76)=   "Column(10).Width=2090"
         Splits(1)._ColumnProps(77)=   "Column(10).DividerColor=0"
         Splits(1)._ColumnProps(78)=   "Column(10)._WidthInPix=2011"
         Splits(1)._ColumnProps(79)=   "Column(10)._ColStyle=516"
         Splits(1)._ColumnProps(80)=   "Column(10).Order=11"
         Splits(1)._ColumnProps(81)=   "Column(11).Width=5106"
         Splits(1)._ColumnProps(82)=   "Column(11).DividerColor=0"
         Splits(1)._ColumnProps(83)=   "Column(11)._WidthInPix=5027"
         Splits(1)._ColumnProps(84)=   "Column(11)._ColStyle=516"
         Splits(1)._ColumnProps(85)=   "Column(11).Order=12"
         Splits(1)._ColumnProps(86)=   "Column(12).Width=2143"
         Splits(1)._ColumnProps(87)=   "Column(12).DividerColor=0"
         Splits(1)._ColumnProps(88)=   "Column(12)._WidthInPix=2064"
         Splits(1)._ColumnProps(89)=   "Column(12)._ColStyle=513"
         Splits(1)._ColumnProps(90)=   "Column(12).Order=13"
         Splits(1)._ColumnProps(91)=   "Column(13).Width=1005"
         Splits(1)._ColumnProps(92)=   "Column(13).DividerColor=0"
         Splits(1)._ColumnProps(93)=   "Column(13)._WidthInPix=926"
         Splits(1)._ColumnProps(94)=   "Column(13)._ColStyle=513"
         Splits(1)._ColumnProps(95)=   "Column(13).Order=14"
         Splits(1)._ColumnProps(96)=   "Column(14).Width=2090"
         Splits(1)._ColumnProps(97)=   "Column(14).DividerColor=0"
         Splits(1)._ColumnProps(98)=   "Column(14)._WidthInPix=2011"
         Splits(1)._ColumnProps(99)=   "Column(14)._ColStyle=513"
         Splits(1)._ColumnProps(100)=   "Column(14).Order=15"
         Splits(1)._ColumnProps(101)=   "Column(15).Width=5398"
         Splits(1)._ColumnProps(102)=   "Column(15).DividerColor=0"
         Splits(1)._ColumnProps(103)=   "Column(15)._WidthInPix=5318"
         Splits(1)._ColumnProps(104)=   "Column(15)._ColStyle=516"
         Splits(1)._ColumnProps(105)=   "Column(15).Order=16"
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
         Caption         =   "DAFTAR CUTI KARYAWAN"
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
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=110,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=107,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=108,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=109,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=46,.parent=13"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
         _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
         _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
         _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
         _StyleDefs(86)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
         _StyleDefs(87)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
         _StyleDefs(88)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
         _StyleDefs(89)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
         _StyleDefs(90)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
         _StyleDefs(91)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
         _StyleDefs(92)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
         _StyleDefs(93)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
         _StyleDefs(94)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
         _StyleDefs(95)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
         _StyleDefs(96)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
         _StyleDefs(97)  =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
         _StyleDefs(98)  =   "Splits(1).Style:id=99,.parent=1"
         _StyleDefs(99)  =   "Splits(1).CaptionStyle:id=116,.parent=4,.bgcolor=&H80000002&"
         _StyleDefs(100) =   ":id=116,.fgcolor=&H80000009&"
         _StyleDefs(101) =   "Splits(1).HeadingStyle:id=100,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(102) =   ":id=100,.fgcolor=&H80000002&"
         _StyleDefs(103) =   "Splits(1).FooterStyle:id=101,.parent=3"
         _StyleDefs(104) =   "Splits(1).InactiveStyle:id=102,.parent=5"
         _StyleDefs(105) =   "Splits(1).SelectedStyle:id=104,.parent=6"
         _StyleDefs(106) =   "Splits(1).EditorStyle:id=103,.parent=7"
         _StyleDefs(107) =   "Splits(1).HighlightRowStyle:id=105,.parent=8"
         _StyleDefs(108) =   "Splits(1).EvenRowStyle:id=106,.parent=9"
         _StyleDefs(109) =   "Splits(1).OddRowStyle:id=115,.parent=10"
         _StyleDefs(110) =   "Splits(1).RecordSelectorStyle:id=117,.parent=11"
         _StyleDefs(111) =   "Splits(1).FilterBarStyle:id=118,.parent=12"
         _StyleDefs(112) =   "Splits(1).Columns(0).Style:id=122,.parent=99"
         _StyleDefs(113) =   "Splits(1).Columns(0).HeadingStyle:id=119,.parent=100"
         _StyleDefs(114) =   "Splits(1).Columns(0).FooterStyle:id=120,.parent=101"
         _StyleDefs(115) =   "Splits(1).Columns(0).EditorStyle:id=121,.parent=103"
         _StyleDefs(116) =   "Splits(1).Columns(1).Style:id=126,.parent=99"
         _StyleDefs(117) =   "Splits(1).Columns(1).HeadingStyle:id=123,.parent=100"
         _StyleDefs(118) =   "Splits(1).Columns(1).FooterStyle:id=124,.parent=101"
         _StyleDefs(119) =   "Splits(1).Columns(1).EditorStyle:id=125,.parent=103"
         _StyleDefs(120) =   "Splits(1).Columns(2).Style:id=130,.parent=99"
         _StyleDefs(121) =   "Splits(1).Columns(2).HeadingStyle:id=127,.parent=100"
         _StyleDefs(122) =   "Splits(1).Columns(2).FooterStyle:id=128,.parent=101"
         _StyleDefs(123) =   "Splits(1).Columns(2).EditorStyle:id=129,.parent=103"
         _StyleDefs(124) =   "Splits(1).Columns(3).Style:id=134,.parent=99"
         _StyleDefs(125) =   "Splits(1).Columns(3).HeadingStyle:id=131,.parent=100"
         _StyleDefs(126) =   "Splits(1).Columns(3).FooterStyle:id=132,.parent=101"
         _StyleDefs(127) =   "Splits(1).Columns(3).EditorStyle:id=133,.parent=103"
         _StyleDefs(128) =   "Splits(1).Columns(4).Style:id=146,.parent=99"
         _StyleDefs(129) =   "Splits(1).Columns(4).HeadingStyle:id=143,.parent=100"
         _StyleDefs(130) =   "Splits(1).Columns(4).FooterStyle:id=144,.parent=101"
         _StyleDefs(131) =   "Splits(1).Columns(4).EditorStyle:id=145,.parent=103"
         _StyleDefs(132) =   "Splits(1).Columns(5).Style:id=150,.parent=99"
         _StyleDefs(133) =   "Splits(1).Columns(5).HeadingStyle:id=147,.parent=100"
         _StyleDefs(134) =   "Splits(1).Columns(5).FooterStyle:id=148,.parent=101"
         _StyleDefs(135) =   "Splits(1).Columns(5).EditorStyle:id=149,.parent=103"
         _StyleDefs(136) =   "Splits(1).Columns(6).Style:id=154,.parent=99"
         _StyleDefs(137) =   "Splits(1).Columns(6).HeadingStyle:id=151,.parent=100"
         _StyleDefs(138) =   "Splits(1).Columns(6).FooterStyle:id=152,.parent=101"
         _StyleDefs(139) =   "Splits(1).Columns(6).EditorStyle:id=153,.parent=103"
         _StyleDefs(140) =   "Splits(1).Columns(7).Style:id=158,.parent=99"
         _StyleDefs(141) =   "Splits(1).Columns(7).HeadingStyle:id=155,.parent=100"
         _StyleDefs(142) =   "Splits(1).Columns(7).FooterStyle:id=156,.parent=101"
         _StyleDefs(143) =   "Splits(1).Columns(7).EditorStyle:id=157,.parent=103"
         _StyleDefs(144) =   "Splits(1).Columns(8).Style:id=162,.parent=99,.alignment=1"
         _StyleDefs(145) =   "Splits(1).Columns(8).HeadingStyle:id=159,.parent=100"
         _StyleDefs(146) =   "Splits(1).Columns(8).FooterStyle:id=160,.parent=101"
         _StyleDefs(147) =   "Splits(1).Columns(8).EditorStyle:id=161,.parent=103"
         _StyleDefs(148) =   "Splits(1).Columns(9).Style:id=54,.parent=99"
         _StyleDefs(149) =   "Splits(1).Columns(9).HeadingStyle:id=51,.parent=100"
         _StyleDefs(150) =   "Splits(1).Columns(9).FooterStyle:id=52,.parent=101"
         _StyleDefs(151) =   "Splits(1).Columns(9).EditorStyle:id=53,.parent=103"
         _StyleDefs(152) =   "Splits(1).Columns(10).Style:id=170,.parent=99"
         _StyleDefs(153) =   "Splits(1).Columns(10).HeadingStyle:id=167,.parent=100"
         _StyleDefs(154) =   "Splits(1).Columns(10).FooterStyle:id=168,.parent=101"
         _StyleDefs(155) =   "Splits(1).Columns(10).EditorStyle:id=169,.parent=103"
         _StyleDefs(156) =   "Splits(1).Columns(11).Style:id=174,.parent=99"
         _StyleDefs(157) =   "Splits(1).Columns(11).HeadingStyle:id=171,.parent=100"
         _StyleDefs(158) =   "Splits(1).Columns(11).FooterStyle:id=172,.parent=101"
         _StyleDefs(159) =   "Splits(1).Columns(11).EditorStyle:id=173,.parent=103"
         _StyleDefs(160) =   "Splits(1).Columns(12).Style:id=178,.parent=99,.alignment=2"
         _StyleDefs(161) =   "Splits(1).Columns(12).HeadingStyle:id=175,.parent=100"
         _StyleDefs(162) =   "Splits(1).Columns(12).FooterStyle:id=176,.parent=101"
         _StyleDefs(163) =   "Splits(1).Columns(12).EditorStyle:id=177,.parent=103"
         _StyleDefs(164) =   "Splits(1).Columns(13).Style:id=182,.parent=99,.alignment=2"
         _StyleDefs(165) =   "Splits(1).Columns(13).HeadingStyle:id=179,.parent=100"
         _StyleDefs(166) =   "Splits(1).Columns(13).FooterStyle:id=180,.parent=101"
         _StyleDefs(167) =   "Splits(1).Columns(13).EditorStyle:id=181,.parent=103"
         _StyleDefs(168) =   "Splits(1).Columns(14).Style:id=186,.parent=99,.alignment=2"
         _StyleDefs(169) =   "Splits(1).Columns(14).HeadingStyle:id=183,.parent=100"
         _StyleDefs(170) =   "Splits(1).Columns(14).FooterStyle:id=184,.parent=101"
         _StyleDefs(171) =   "Splits(1).Columns(14).EditorStyle:id=185,.parent=103"
         _StyleDefs(172) =   "Splits(1).Columns(15).Style:id=190,.parent=99"
         _StyleDefs(173) =   "Splits(1).Columns(15).HeadingStyle:id=187,.parent=100"
         _StyleDefs(174) =   "Splits(1).Columns(15).FooterStyle:id=188,.parent=101"
         _StyleDefs(175) =   "Splits(1).Columns(15).EditorStyle:id=189,.parent=103"
         _StyleDefs(176) =   "Named:id=33:Normal"
         _StyleDefs(177) =   ":id=33,.parent=0"
         _StyleDefs(178) =   "Named:id=34:Heading"
         _StyleDefs(179) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(180) =   ":id=34,.wraptext=-1"
         _StyleDefs(181) =   "Named:id=35:Footing"
         _StyleDefs(182) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(183) =   "Named:id=36:Selected"
         _StyleDefs(184) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(185) =   "Named:id=37:Caption"
         _StyleDefs(186) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(187) =   "Named:id=38:HighlightRow"
         _StyleDefs(188) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(189) =   "Named:id=39:EvenRow"
         _StyleDefs(190) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(191) =   "Named:id=40:OddRow"
         _StyleDefs(192) =   ":id=40,.parent=33"
         _StyleDefs(193) =   "Named:id=41:RecordSelector"
         _StyleDefs(194) =   ":id=41,.parent=34"
         _StyleDefs(195) =   "Named:id=42:FilterBar"
         _StyleDefs(196) =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Gen 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   24
         Top             =   420
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9551
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO"
         Columns(0).DataField=   "general_leave_number"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "TANGGAL"
         Columns(1).DataField=   "general_leave_date"
         Columns(1).NumberFormat=   "yyyy-MM-dd"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KETERANGAN"
         Columns(2).DataField=   "description"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).Size  =   2
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   1
         Splits(0).DividerColor=   13160660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3228"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3149"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=514"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=4260"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4180"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=13547"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=13467"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
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
         Caption         =   "DAFTAR CUTI BERSAMA"
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
         _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=162,.parent=99,.alignment=1"
         _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=159,.parent=100"
         _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=160,.parent=101"
         _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=161,.parent=103"
         _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=186,.parent=99,.alignment=2"
         _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=183,.parent=100"
         _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=184,.parent=101"
         _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=185,.parent=103"
         _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=190,.parent=99"
         _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=187,.parent=100"
         _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=188,.parent=101"
         _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=189,.parent=103"
         _StyleDefs(47)  =   "Named:id=33:Normal"
         _StyleDefs(48)  =   ":id=33,.parent=0"
         _StyleDefs(49)  =   "Named:id=34:Heading"
         _StyleDefs(50)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   ":id=34,.wraptext=-1"
         _StyleDefs(52)  =   "Named:id=35:Footing"
         _StyleDefs(53)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   "Named:id=36:Selected"
         _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(56)  =   "Named:id=37:Caption"
         _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(58)  =   "Named:id=38:HighlightRow"
         _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(60)  =   "Named:id=39:EvenRow"
         _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(62)  =   "Named:id=40:OddRow"
         _StyleDefs(63)  =   ":id=40,.parent=33"
         _StyleDefs(64)  =   "Named:id=41:RecordSelector"
         _StyleDefs(65)  =   ":id=41,.parent=34"
         _StyleDefs(66)  =   "Named:id=42:FilterBar"
         _StyleDefs(67)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Sum 
         Height          =   5115
         Left            =   120
         TabIndex        =   36
         Top             =   420
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   9022
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "company_code"
         Columns(0).DataField=   "company_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "DEPT. CODE"
         Columns(1).DataField=   "department_code"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "DEPT. NAME"
         Columns(2).DataField=   "department_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "EMPLOYEE CODE"
         Columns(3).DataField=   "employee_code"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "KODE KARY."
         Columns(4).DataField=   "nik"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "NAMA KARY."
         Columns(5).DataField=   "employee_name"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "PERIODE"
         Columns(6).DataField=   "leave_periode"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "PERIODE AWAL"
         Columns(7).DataField=   "start_periode"
         Columns(7).NumberFormat=   "yyyy-MM-dd"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "PERIODE AKHIR"
         Columns(8).DataField=   "end_periode"
         Columns(8).NumberFormat=   "yyyy-MM-dd"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "CUTI MAKS."
         Columns(9).DataField=   "max_leave"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "DIGUNAKAN"
         Columns(10).DataField=   "actual_leave"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "SISA CUTI"
         Columns(11).DataField=   "leave_available"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   16
         Columns(12)._MaxComboItems=   5
         Columns(12).ValueItems(0)._DefaultItem=   0
         Columns(12).ValueItems(0).Value=   "0"
         Columns(12).ValueItems(0).Value.vt=   8
         Columns(12).ValueItems(0).DisplayValue=   "NO"
         Columns(12).ValueItems(0).DisplayValue.vt=   8
         Columns(12).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(12).ValueItems(1)._DefaultItem=   0
         Columns(12).ValueItems(1).Value=   "1"
         Columns(12).ValueItems(1).Value.vt=   8
         Columns(12).ValueItems(1).DisplayValue=   "YES"
         Columns(12).ValueItems(1).DisplayValue.vt=   8
         Columns(12).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(12).ValueItems.Count=   2
         Columns(12).Caption=   "CHAIN OVER"
         Columns(12).DataField=   "flag_close"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   13
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=13"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=3942"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3863"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(14)=   "Column(1).Visible=0"
         Splits(0)._ColumnProps(15)=   "Column(1).AllowFocus=0"
         Splits(0)._ColumnProps(16)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(17)=   "Column(2).Width=7408"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=7329"
         Splits(0)._ColumnProps(20)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(21)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(23)=   "Column(2).AllowFocus=0"
         Splits(0)._ColumnProps(24)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(25)=   "Column(3).Width=2619"
         Splits(0)._ColumnProps(26)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(3)._WidthInPix=2540"
         Splits(0)._ColumnProps(28)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(29)=   "Column(3).Visible=0"
         Splits(0)._ColumnProps(30)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(31)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(32)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(35)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(36)=   "Column(4)._MinWidth=74771376"
         Splits(0)._ColumnProps(37)=   "Column(5).Width=5477"
         Splits(0)._ColumnProps(38)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(5)._WidthInPix=5398"
         Splits(0)._ColumnProps(40)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(41)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(42)=   "Column(5)._MinWidth=74771376"
         Splits(0)._ColumnProps(43)=   "Column(6).Width=2408"
         Splits(0)._ColumnProps(44)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(6)._WidthInPix=2328"
         Splits(0)._ColumnProps(46)=   "Column(6)._ColStyle=513"
         Splits(0)._ColumnProps(47)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(48)=   "Column(6)._MinWidth=74771280"
         Splits(0)._ColumnProps(49)=   "Column(7).Width=2487"
         Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=2408"
         Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=513"
         Splits(0)._ColumnProps(53)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(54)=   "Column(7)._MinWidth=74746528"
         Splits(0)._ColumnProps(55)=   "Column(8).Width=2434"
         Splits(0)._ColumnProps(56)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(8)._WidthInPix=2355"
         Splits(0)._ColumnProps(58)=   "Column(8)._ColStyle=513"
         Splits(0)._ColumnProps(59)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(60)=   "Column(8)._MinWidth=74745312"
         Splits(0)._ColumnProps(61)=   "Column(9).Width=1905"
         Splits(0)._ColumnProps(62)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(9)._WidthInPix=1826"
         Splits(0)._ColumnProps(64)=   "Column(9)._ColStyle=513"
         Splits(0)._ColumnProps(65)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(66)=   "Column(9)._MinWidth=74744864"
         Splits(0)._ColumnProps(67)=   "Column(10).Width=1905"
         Splits(0)._ColumnProps(68)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(69)=   "Column(10)._WidthInPix=1826"
         Splits(0)._ColumnProps(70)=   "Column(10)._ColStyle=513"
         Splits(0)._ColumnProps(71)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(72)=   "Column(10)._MinWidth=74714320"
         Splits(0)._ColumnProps(73)=   "Column(11).Width=1905"
         Splits(0)._ColumnProps(74)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(11)._WidthInPix=1826"
         Splits(0)._ColumnProps(76)=   "Column(11)._ColStyle=513"
         Splits(0)._ColumnProps(77)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(78)=   "Column(11)._MinWidth=74678208"
         Splits(0)._ColumnProps(79)=   "Column(12).Width=1879"
         Splits(0)._ColumnProps(80)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(12)._WidthInPix=1799"
         Splits(0)._ColumnProps(82)=   "Column(12).AllowSizing=0"
         Splits(0)._ColumnProps(83)=   "Column(12)._ColStyle=513"
         Splits(0)._ColumnProps(84)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(85)=   "Column(12).AllowFocus=0"
         Splits(0)._ColumnProps(86)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(87)=   "Column(12)._MinWidth=74671168"
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
         Caption         =   "REKAP CUTI"
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
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=86,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=82,.parent=13,.alignment=2"
         _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
         _StyleDefs(86)  =   "Named:id=33:Normal"
         _StyleDefs(87)  =   ":id=33,.parent=0"
         _StyleDefs(88)  =   "Named:id=34:Heading"
         _StyleDefs(89)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(90)  =   ":id=34,.wraptext=-1"
         _StyleDefs(91)  =   "Named:id=35:Footing"
         _StyleDefs(92)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(93)  =   "Named:id=36:Selected"
         _StyleDefs(94)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(95)  =   "Named:id=37:Caption"
         _StyleDefs(96)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(97)  =   "Named:id=38:HighlightRow"
         _StyleDefs(98)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(99)  =   "Named:id=39:EvenRow"
         _StyleDefs(100) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(101) =   "Named:id=40:OddRow"
         _StyleDefs(102) =   ":id=40,.parent=33"
         _StyleDefs(103) =   "Named:id=41:RecordSelector"
         _StyleDefs(104) =   ":id=41,.parent=34"
         _StyleDefs(105) =   "Named:id=42:FilterBar"
         _StyleDefs(106) =   ":id=42,.parent=33"
      End
   End
   Begin prj_panji.vbButton cmdExit 
      Height          =   705
      Left            =   11670
      TabIndex        =   38
      Top             =   8730
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
      MICON           =   "frmTransLeave.frx":DCF2
      PICN            =   "frmTransLeave.frx":DD0E
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
      Caption         =   "CUTI"
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
      TabIndex        =   34
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frmTransLeave.frx":EDA0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12690
   End
End
Attribute VB_Name = "frm_trans_leave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsCompany As New ADODB.Recordset
Dim rsEmpLeave As New ADODB.Recordset
Dim rsGenLeave As New ADODB.Recordset
Dim rsSumLeave As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim vEmpLeave_Number As Long
Dim vGenLeave_Number As Long

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    If SSTab1.Tab = 0 Then
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
            MsgBox "Description Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_description.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        'validasi description
        If Trim(txt_desc) = "" Then
            MsgBox "Description Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_desc.SetFocus
            check_validate_new = False
            Exit Function
        End If
    End If
End Function

Private Sub load_data()
    timer1.Enabled = True
End Sub

Private Sub date_event()
    If cbo_date_to.ListIndex = 0 Then
        DTPicker_date_to.Visible = False
    Else
        DTPicker_date_to.Visible = True
        DTPicker_date_to.Value = DTPicker_date_from.Value
    End If
End Sub

Private Sub cbo_date_to_Click()
    Call date_event
End Sub

Private Sub cmd_browse_Click()
    frm_lookup_mst_employee.public_int_mode = 1
    frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
    frm_lookup_mst_employee.Show 1
End Sub

Private Sub cancel_data()
    int_mode = 0
    Call load_mode
End Sub

Private Sub delete_data()
Dim i As Integer
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_Emp.ApproxCount > 0 And TDBGrid_Emp.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Emp.Columns("employee_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from t_leave where leave_number = " _
            & TDBGrid_Emp.Columns("leave_number").Value
        CnG.CommitTrans
        
        Call load_data_leave
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_Gen.ApproxCount > 0 And TDBGrid_Gen.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Gen.Columns("description").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from t_general_leave where general_leave_number = " _
            & TDBGrid_Gen.Columns("general_leave_number").Value
        CnG.CommitTrans
        
        Call load_data_general_leave
    End If
    
    int_mode = 0
    Call load_mode
End Sub

Public Sub set_edit_data()
    If SSTab1.Tab = 0 Then
        With rsEmpLeave
            txt_employee_code.Text = .Fields("employee_code").Value
            txt_nik.Text = .Fields("nik").Value
            txt_employee_name.Text = .Fields("employee_name").Value
            DTPicker_date_from.Value = .Fields("leave_date_from").Value
            cbo_date_to.ListIndex = .Fields("flag_date_to").Value
            If cbo_date_to.ListIndex = 1 Then _
                DTPicker_date_to.Value = .Fields("leave_date_to").Value
            txt_description.Text = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 1 Then
        With rsGenLeave
            DTPicker_date_general_leave.Value = .Fields("general_leave_date").Value
            txt_desc.Text = .Fields("description").Value
        End With
    End If
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

    CnG.BeginTrans
    If SSTab1.Tab = 0 Then
        SQL = "select ifnull(max(leave_number),0)+1 as leave_number from t_leave"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            vEmpLeave_Number = rs!leave_number
        End If
        rs.Close
        
        SQL = "INSERT INTO t_leave(company_code,leave_number,employee_code,leave_date_from," & _
                "flag_date_to,leave_date_to,description,entry_date,entry_user) " & _
              "VALUES( " & _
                "'" & TDBCombo_company.Text & "','" & vEmpLeave_Number & "','" & Trim(txt_employee_code.Text) & "'," & _
                "'" & Format(DTPicker_date_from.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "'" & cbo_date_to.ListIndex & "'," & _
                "'" & IIf(cbo_date_to.ListIndex = 1, Format(DTPicker_date_to.Value, "yyyy-MM-dd HH:mm:ss"), "00:00:00") & "'," & _
                "'" & Trim(txt_description.Text) & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "select ifnull(max(general_leave_number),0)+1 as general_leave_number from t_general_leave"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            vGenLeave_Number = rs!general_leave_number
        End If
        rs.Close
        
        SQL = "INSERT INTO t_general_leave(company_code,general_leave_number,general_leave_date," & _
                "description,entry_date,entry_user) " & _
              "VALUES(" & _
                "'" & TDBCombo_company.Text & "','" & vGenLeave_Number & "','" & Format(DTPicker_date_general_leave.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "'" & Trim(txt_desc.Text) & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    End If
    CnG.CommitTrans
    
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_old_data()
On Error GoTo Err

    CnG.BeginTrans
    If SSTab1.Tab = 0 Then
        SQL = "UPDATE t_leave SET employee_code = '" & Trim(txt_employee_code) & "'," & _
                "leave_date_from = '" & Format(DTPicker_date_from.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "flag_date_to = '" & cbo_date_to.ListIndex & "'," & _
                "leave_date_to = '" & IIf(cbo_date_to.ListIndex = 1, Format(DTPicker_date_to.Value, "yyyy-MM-dd HH:mm:ss"), "00:00:00") & "'," & _
                "description = '" & Trim(txt_description.Text) & "',edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE leave_number = '" & TDBGrid_Emp.Columns("leave_number").Value & "' " & _
                "AND company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "UPDATE t_general_leave SET general_leave_date = '" & Format(DTPicker_date_general_leave.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "description = '" & Trim(txt_desc.Text) & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE general_leave_number = '" & TDBGrid_Gen.Columns("general_leave_number").Value & "' " & _
                "AND company_code = '" & TDBCombo_company.Text & "'"
        CnG.Execute SQL
    End If
    CnG.CommitTrans
    
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub simpan_data()
    If int_mode = 1 Then
        If Not check_validate_new Then Exit Sub
    '    If check_validate_exist_new Then
    '        Call check_invalid: Exit Sub
    '    End If
        Call insert_new_data
    ElseIf int_mode = 2 Then
        If Not check_validate_new Then Exit Sub
    '    If check_validate_exist_edit Then
    '        Call check_invalid: Exit Sub
    '    End If
        Call edit_old_data
    End If
    
    If SSTab1.Tab = 0 Then
        Call load_data_leave
    ElseIf SSTab1.Tab = 1 Then
        Call load_data_general_leave
    End If
    
    int_mode = 0
    Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
    If SSTab1.Tab = 0 Then
        cmdNew_Emp.Enabled = a And blnUser_Add
        cmdSave_Emp.Enabled = b
        cmdEdit_Emp.Enabled = c And blnUser_Edit
        cmdDelete_Emp.Enabled = d And blnUser_Delete
        cmdCancel_Emp.Enabled = e
    ElseIf SSTab1.Tab = 1 Then
        cmdNew_Gen.Enabled = a And blnUser_Add
        cmdSave_Gen.Enabled = b
        cmdEdit_Gen.Enabled = c And blnUser_Edit
        cmdDelete_Gen.Enabled = d And blnUser_Delete
        cmdCancel_Gen.Enabled = e
    End If
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
    If SSTab1.Tab = 0 Then
        txt_employee_code = ""
        txt_nik = ""
        txt_employee_name = ""
        cbo_date_to.ListIndex = 0
        DTPicker_date_from.Value = Now: DTPicker_date_to.Value = Now
        txt_description = ""
    ElseIf SSTab1.Tab = 1 Then
        DTPicker_date_general_leave.Value = Now
        txt_desc = ""
    End If
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_Emp.Visible = True
            txt_nik.Enabled = True
            TDBGrid_Emp.Enabled = False
            Call set_new_data
            
            If txt_nik.Enabled = True Then
                txt_nik.SetFocus
            End If
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_Gen.Visible = True
            DTPicker_date_general_leave.Enabled = True
            TDBGrid_Gen.Enabled = False
            Call set_new_data
            
            If DTPicker_date_general_leave.Enabled = True Then
                DTPicker_date_general_leave.SetFocus
            End If
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_Emp.Visible = False
            TDBGrid_Emp.Enabled = True
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_Gen.Visible = False
            TDBGrid_Gen.Enabled = True
        End If
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If SSTab1.Tab = 0 Then
            txt_nik.Enabled = False
            fra_entry_Emp.Visible = True
            TDBGrid_Emp.Enabled = False
        ElseIf SSTab1.Tab = 1 Then
            DTPicker_date_general_leave.Enabled = False
            fra_entry_Gen.Visible = True
            TDBGrid_Gen.Enabled = False
        End If
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

Private Sub cmdLoad_Sum_Click()
    frm_lookup_leave_periode.Show 1
End Sub

Private Sub Form_Load()
    Call load_data_company
    Call createGridKar
    
    SSTab1.Tab = 0
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
    timer1.Enabled = True
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 0 Then
        For Each Col In TDBGrid_Emp.Columns
            Col.FilterText = ""
        Next Col
        rsEmpLeave.Filter = adFilterNone
    ElseIf SSTab1.Tab = 1 Then
        For Each Col In TDBGrid_Gen.Columns
            Col.FilterText = ""
        Next Col
        rsGenLeave.Filter = adFilterNone
    ElseIf SSTab1.Tab = 2 Then
        For Each Col In TDBGrid_Sum.Columns
            Col.FilterText = ""
        Next Col
        rsSumLeave.Filter = adFilterNone
    End If
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

Private Sub filter_change()
Dim i As Integer

On Error GoTo Err
    
    If SSTab1.Tab = 0 Then
        Set Cols = TDBGrid_Emp.Columns
        i = TDBGrid_Emp.Col
        TDBGrid_Emp.HoldFields
        
        rsEmpLeave.Filter = getFilter()
        TDBGrid_Emp.Col = i
        TDBGrid_Emp.EditActive = True
        
        TDBGrid_Emp.SelStart = Len(TDBGrid_Emp.Columns(i).FilterText)
        If TDBGrid_Emp.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Emp.Col = i
        End If
    ElseIf SSTab1.Tab = 1 Then
        Set Cols = TDBGrid_Gen.Columns
        i = TDBGrid_Gen.Col
        TDBGrid_Gen.HoldFields
        
        rsGenLeave.Filter = getFilter()
        TDBGrid_Gen.Col = i
        TDBGrid_Gen.EditActive = True
        
        TDBGrid_Gen.SelStart = Len(TDBGrid_Gen.Columns(i).FilterText)
        If TDBGrid_Gen.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Gen.Col = i
        End If
    ElseIf SSTab1.Tab = 2 Then
        Set Cols = TDBGrid_Sum.Columns
        i = TDBGrid_Gen.Col
        TDBGrid_Sum.HoldFields
        
        rsSumLeave.Filter = getFilter()
        TDBGrid_Sum.Col = i
        TDBGrid_Sum.EditActive = True
        
        TDBGrid_Sum.SelStart = Len(TDBGrid_Sum.Columns(i).FilterText)
        If TDBGrid_Sum.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Sum.Col = i
        End If
    End If

    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        Call load_data_leave
    ElseIf SSTab1.Tab = 1 Then
        Call load_data_general_leave
    ElseIf SSTab1.Tab = 2 Then
        Call load_data_summary_leave
    End If
    
    int_mode = 0
    Call load_mode
End Sub

Public Sub generate_summary_leave(ByRef dtp1 As DTPicker)
    CnG.BeginTrans
        CnG.Execute "call spg_leave_periode ('" & Format(dtp1, "yyyy-MM-dd HH:mm:ss") & "')"
    CnG.CommitTrans
    
    Call load_data_summary_leave
End Sub

Private Sub TDBCombo_company_ItemChange()
If TDBCombo_company.ApproxCount > 0 Then
    TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
    txt_company_name = TDBCombo_company.Columns("company_name").Value
    
    If SSTab1.Tab = 0 Then
        Call load_data_leave
    ElseIf SSTab1.Tab = 1 Then
        Call load_data_general_leave
    ElseIf SSTab1.Tab = 2 Then
        Call load_data_summary_leave
    End If
End If
End Sub

Private Sub load_data_leave()
    If rsEmpLeave.State Then rsEmpLeave.Close
    SQL = "SELECT a.*,b.employee_name,b.company_code,c.company_name," & _
            "b.division_code,e.division_name,b.title_code," & _
            "f.title_name,b.level_code,b.nik " & _
          "FROM t_leave a LEFT JOIN m_employee b ON a.employee_code = b.employee_code " & _
            "JOIN m_company c ON b.company_code = c.company_code " & _
            "JOIN m_division e ON b.division_code = e.division_code " & _
                "AND b.company_code = e.company_code " & _
            "JOIN m_title f ON b.title_code = f.title_code " & _
          "WHERE a.company_code = '" & TDBCombo_company.Text & "' " & _
          "order by leave_number"
    rsEmpLeave.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Emp.DataSource = rsEmpLeave
End Sub

Private Sub load_data_general_leave()
    If rsGenLeave.State Then rsGenLeave.Close
    SQL = "select * from t_general_leave " & _
          "where company_code = '" & TDBCombo_company.Text & "' order by general_leave_number"
    rsGenLeave.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_Gen.DataSource = rsGenLeave
End Sub

Private Sub load_data_summary_leave()
    If rsSumLeave.State Then rsSumLeave.Close
    SQL = "SELECT b.start_periode,b.end_periode, " & _
            "CAST(CONCAT(YEAR(b.start_periode),_latin1'-',YEAR(b.end_periode)) AS CHAR CHARSET latin1) AS leave_periode," & _
            "b.max_leave,b.actual_leave,b.over_leave,(-(1) * b.over_leave) AS leave_available," & _
            "b.flag_close,a.*,e.division_name,c.company_name,f.title_name " & _
          "FROM m_employee a LEFT JOIN t_leave_periode b ON a.employee_code = b.employee_code " & _
            "JOIN m_company c ON a.company_code = c.company_code " & _
            "JOIN m_division e ON a.division_code = e.division_code " & _
                "AND a.company_code = e.company_code " & _
            "JOIN m_title f ON a.title_code = f.title_code " & _
          "WHERE a.company_code = '" & TDBCombo_company.Text & "' " & _
          "order by a.division_code,a.employee_code"
    rsSumLeave.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly

    TDBGrid_Sum.DataSource = rsSumLeave
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
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
                    & "WHERE (a.nik LIKE '%" & txt_nik.Text & "%' " _
                    & "OR a.employee_name LIKE '%" & txt_nik.Text & "%') " _
                    & "AND a.flag_active <> 0 AND a.company_code = '" & TDBCombo_company.Text & "'"
        Else
            SQL = "SELECT a.nik,a.employee_name," _
                        & "a.division_code,b.division_name," _
                        & "a.title_code,c.title_name,a.employee_code " _
                    & "FROM m_employee a JOIN m_division b ON a.division_code = b.division_code and a.company_code = b.company_code " _
                    & "JOIN m_title c ON a.title_code = c.title_code " _
                    & "JOIN m_company e ON a.company_code = e.company_code " _
                    & "WHERE (a.nik LIKE '%" & txt_nik.Text & "%' " _
                    & "OR a.employee_name LIKE '%" & txt_nik.Text & "%') " _
                    & "AND a.flag_active <> 0 AND (level_code = ANY (SELECT access_level_code FROM t_user_access_level WHERE level_code = '" & LOGIN_CODE & "' AND allow_access <> 0)) " _
                    & "AND a.company_code = '" & TDBCombo_company.Text & "' " _
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

Private Sub cmdNew_Emp_Click()
    Call new_data
End Sub

Private Sub cmdSave_Emp_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Emp_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Emp_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Emp_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_Gen_Click()
    Call new_data
End Sub

Private Sub cmdSave_Gen_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Gen_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Gen_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Gen_Click()
    Call cancel_data
End Sub


Private Sub TDBGrid_Emp_FilterChange()
    Call filter_change
End Sub

Private Sub TDBGrid_Gen_FilterChange()
    Call filter_change
End Sub

Private Sub TDBGrid_Sum_FilterChange()
    Call filter_change
End Sub
