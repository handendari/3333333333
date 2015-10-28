VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_company 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER PERUSAHAAN"
   ClientHeight    =   8295
   ClientLeft      =   -735
   ClientTop       =   345
   ClientWidth     =   11850
   Icon            =   "frm_mst_company.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_company 
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   150
      TabIndex        =   66
      Top             =   660
      Visible         =   0   'False
      Width           =   6945
      Begin VB.TextBox txt_company 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3150
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   67
         Top             =   120
         Width           =   3645
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   1410
         OleObjectBlob   =   "frm_mst_company.frx":058A
         TabIndex        =   70
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "PERUSAHAAN"
         Height          =   195
         Left            =   120
         TabIndex        =   69
         Top             =   180
         Width           =   1110
      End
   End
   Begin prj_panji.vbButton cmdExit 
      Height          =   705
      Left            =   10740
      TabIndex        =   37
      Top             =   7410
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
      MICON           =   "frm_mst_company.frx":2548
      PICN            =   "frm_mst_company.frx":2564
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6105
      Left            =   90
      TabIndex        =   15
      Top             =   1260
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10769
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      Tab             =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "PERUSAHAAN"
      TabPicture(0)   =   "frm_mst_company.frx":35F6
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "TDBGrid_Company"
      Tab(0).Control(1)=   "fra_entry_Company"
      Tab(0).Control(2)=   "frmTombol"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "DEPARTEMEN"
      TabPicture(1)   =   "frm_mst_company.frx":3612
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "TDBGrid_Department"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fra_entry_Department"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "DIVISI"
      TabPicture(2)   =   "frm_mst_company.frx":362E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "fra_entry_Division"
      Tab(2).Control(2)=   "TDBGrid_Division"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "GRADE"
      TabPicture(3)   =   "frm_mst_company.frx":364A
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).Control(1)=   "fra_entry_Grade"
      Tab(3).Control(2)=   "TDBGrid_Grade"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "LEVEL"
      TabPicture(4)   =   "frm_mst_company.frx":3666
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).Control(1)=   "fra_entry_level"
      Tab(4).Control(2)=   "TDBGrid_Level"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "PERFORMA"
      TabPicture(5)   =   "frm_mst_company.frx":3682
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "TDBGrid_Perf"
      Tab(5).Control(1)=   "fra_entry_perf"
      Tab(5).Control(2)=   "Frame5"
      Tab(5).ControlCount=   3
      Begin VB.Frame fra_entry_Department 
         Height          =   2175
         Left            =   240
         TabIndex        =   113
         Top             =   2280
         Visible         =   0   'False
         Width           =   11205
         Begin VB.TextBox txt_department_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   116
            Top             =   1320
            Width           =   3495
         End
         Begin VB.TextBox txt_department_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   115
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txt_department_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   10
            TabIndex        =   114
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   3645
            TabIndex        =   119
            Top             =   1320
            Width           =   1110
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE*"
            Height          =   195
            Left            =   4230
            TabIndex        =   118
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA*"
            Height          =   195
            Left            =   3450
            TabIndex        =   117
            Top             =   960
            Width           =   1305
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   240
         TabIndex        =   107
         Top             =   4590
         Width           =   11205
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew_Department 
            Height          =   705
            Left            =   540
            TabIndex        =   108
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
            MICON           =   "frm_mst_company.frx":369E
            PICN            =   "frm_mst_company.frx":36BA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Department 
            Height          =   705
            Left            =   1560
            TabIndex        =   109
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
            MICON           =   "frm_mst_company.frx":474C
            PICN            =   "frm_mst_company.frx":4768
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Department 
            Height          =   705
            Left            =   2580
            TabIndex        =   110
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
            MICON           =   "frm_mst_company.frx":57FA
            PICN            =   "frm_mst_company.frx":5816
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Department 
            Height          =   705
            Left            =   3600
            TabIndex        =   111
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
            MICON           =   "frm_mst_company.frx":68A8
            PICN            =   "frm_mst_company.frx":68C4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Department 
            Height          =   705
            Left            =   4620
            TabIndex        =   112
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "s"
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
            MICON           =   "frm_mst_company.frx":7956
            PICN            =   "frm_mst_company.frx":7972
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
      Begin VB.Frame Frame5 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   97
         Top             =   4560
         Width           =   11295
         Begin prj_panji.vbButton cmdNew_Perf 
            Height          =   705
            Left            =   540
            TabIndex        =   98
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
            MICON           =   "frm_mst_company.frx":8A04
            PICN            =   "frm_mst_company.frx":8A20
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Perf 
            Height          =   705
            Left            =   1560
            TabIndex        =   99
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
            MICON           =   "frm_mst_company.frx":9AB2
            PICN            =   "frm_mst_company.frx":9ACE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Perf 
            Height          =   705
            Left            =   2580
            TabIndex        =   100
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
            MICON           =   "frm_mst_company.frx":AB60
            PICN            =   "frm_mst_company.frx":AB7C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Perf 
            Height          =   705
            Left            =   3600
            TabIndex        =   101
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
            MICON           =   "frm_mst_company.frx":BC0E
            PICN            =   "frm_mst_company.frx":BC2A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Perf 
            Height          =   705
            Left            =   4620
            TabIndex        =   102
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
            MICON           =   "frm_mst_company.frx":CCBC
            PICN            =   "frm_mst_company.frx":CCD8
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
      Begin VB.Frame fra_entry_perf 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   90
         Top             =   1650
         Visible         =   0   'False
         Width           =   11295
         Begin VB.TextBox txt_perf_grade 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   63
            Top             =   1380
            Width           =   1695
         End
         Begin VB.TextBox txt_perf_upper 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   62
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txt_perf_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   64
            Top             =   1740
            Width           =   3495
         End
         Begin VB.TextBox txt_perf_under 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   61
            Top             =   660
            Width           =   1695
         End
         Begin VB.TextBox txt_perf_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   60
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "GRADE*"
            Height          =   195
            Left            =   4800
            TabIndex        =   95
            Top             =   1380
            Width           =   1200
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NILAI AKHIR*"
            Height          =   195
            Left            =   5010
            TabIndex        =   94
            Top             =   1020
            Width           =   1005
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   4890
            TabIndex        =   93
            Top             =   1740
            Width           =   1110
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "NO.*"
            Height          =   195
            Left            =   600
            TabIndex        =   92
            Top             =   630
            Width           =   345
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NILAI AWAL*"
            Height          =   195
            Left            =   5040
            TabIndex        =   91
            Top             =   660
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74730
         TabIndex        =   59
         Top             =   4530
         Width           =   11295
         Begin VB.Timer Timer5 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew_Level 
            Height          =   705
            Left            =   540
            TabIndex        =   86
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
            MICON           =   "frm_mst_company.frx":DD6A
            PICN            =   "frm_mst_company.frx":DD86
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Level 
            Height          =   705
            Left            =   1560
            TabIndex        =   87
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
            MICON           =   "frm_mst_company.frx":EE18
            PICN            =   "frm_mst_company.frx":EE34
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Level 
            Height          =   705
            Left            =   2580
            TabIndex        =   88
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
            MICON           =   "frm_mst_company.frx":FEC6
            PICN            =   "frm_mst_company.frx":FEE2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Level 
            Height          =   705
            Left            =   3600
            TabIndex        =   89
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
            MICON           =   "frm_mst_company.frx":10F74
            PICN            =   "frm_mst_company.frx":10F90
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Level 
            Height          =   705
            Left            =   4620
            TabIndex        =   68
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
            MICON           =   "frm_mst_company.frx":12022
            PICN            =   "frm_mst_company.frx":1203E
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
      Begin VB.Frame fra_entry_level 
         Height          =   2775
         Left            =   -74730
         TabIndex        =   52
         Top             =   1650
         Visible         =   0   'False
         Width           =   11175
         Begin VB.TextBox txt_level_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   10
            TabIndex        =   53
            Top             =   810
            Width           =   1215
         End
         Begin VB.TextBox txt_level_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   50
            TabIndex        =   55
            Top             =   1230
            Width           =   3795
         End
         Begin VB.TextBox txt_level_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   50
            TabIndex        =   57
            Top             =   1620
            Width           =   5145
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA*"
            Height          =   195
            Left            =   2820
            TabIndex        =   58
            Top             =   1260
            Width           =   525
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE*"
            Height          =   195
            Left            =   2820
            TabIndex        =   56
            Top             =   840
            Width           =   510
         End
         Begin VB.Label tlabel26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   2235
            TabIndex        =   54
            Top             =   1650
            Width           =   1110
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74790
         TabIndex        =   50
         Top             =   4560
         Width           =   11175
         Begin VB.Timer Timer4 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew_Grade 
            Height          =   705
            Left            =   540
            TabIndex        =   81
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
            MICON           =   "frm_mst_company.frx":130D0
            PICN            =   "frm_mst_company.frx":130EC
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Grade 
            Height          =   705
            Left            =   1560
            TabIndex        =   82
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
            MICON           =   "frm_mst_company.frx":1417E
            PICN            =   "frm_mst_company.frx":1419A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Grade 
            Height          =   705
            Left            =   2580
            TabIndex        =   83
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
            MICON           =   "frm_mst_company.frx":1522C
            PICN            =   "frm_mst_company.frx":15248
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Grade 
            Height          =   705
            Left            =   3600
            TabIndex        =   84
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
            MICON           =   "frm_mst_company.frx":162DA
            PICN            =   "frm_mst_company.frx":162F6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Grade 
            Height          =   705
            Left            =   4620
            TabIndex        =   85
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
            MICON           =   "frm_mst_company.frx":17388
            PICN            =   "frm_mst_company.frx":173A4
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
      Begin VB.Frame fra_entry_Grade 
         Height          =   2655
         Left            =   -74790
         TabIndex        =   43
         Top             =   1770
         Visible         =   0   'False
         Width           =   11175
         Begin VB.TextBox txt_grade_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   46
            Top             =   1530
            Width           =   3495
         End
         Begin VB.TextBox txt_grade_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   45
            Top             =   1170
            Width           =   3495
         End
         Begin VB.TextBox txt_grade_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   10
            TabIndex        =   44
            Top             =   810
            Width           =   1695
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   2985
            TabIndex        =   49
            Top             =   1530
            Width           =   1110
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE*"
            Height          =   195
            Left            =   3600
            TabIndex        =   48
            Top             =   810
            Width           =   510
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA*"
            Height          =   195
            Left            =   3570
            TabIndex        =   47
            Top             =   1170
            Width           =   525
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   41
         Top             =   4560
         Width           =   11175
         Begin VB.Timer Timer3 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew_Division 
            Height          =   705
            Left            =   540
            TabIndex        =   76
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
            MICON           =   "frm_mst_company.frx":18436
            PICN            =   "frm_mst_company.frx":18452
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Division 
            Height          =   705
            Left            =   1560
            TabIndex        =   77
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
            MICON           =   "frm_mst_company.frx":194E4
            PICN            =   "frm_mst_company.frx":19500
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Division 
            Height          =   705
            Left            =   2580
            TabIndex        =   78
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
            MICON           =   "frm_mst_company.frx":1A592
            PICN            =   "frm_mst_company.frx":1A5AE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Division 
            Height          =   705
            Left            =   3600
            TabIndex        =   79
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
            MICON           =   "frm_mst_company.frx":1B640
            PICN            =   "frm_mst_company.frx":1B65C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Division 
            Height          =   705
            Left            =   4620
            TabIndex        =   80
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
            MICON           =   "frm_mst_company.frx":1C6EE
            PICN            =   "frm_mst_company.frx":1C70A
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
      Begin VB.Frame fra_entry_Division 
         Height          =   2355
         Left            =   -74760
         TabIndex        =   33
         Top             =   2070
         Visible         =   0   'False
         Width           =   11175
         Begin VB.TextBox txt_division_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   10
            TabIndex        =   34
            Top             =   660
            Width           =   1695
         End
         Begin VB.TextBox txt_division_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   35
            Top             =   1020
            Width           =   3495
         End
         Begin VB.TextBox txt_division_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   36
            Top             =   1380
            Width           =   3495
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA*"
            Height          =   195
            Left            =   4170
            TabIndex        =   40
            Top             =   1020
            Width           =   525
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE*"
            Height          =   195
            Left            =   4200
            TabIndex        =   39
            Top             =   660
            Width           =   510
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   3615
            TabIndex        =   38
            Top             =   1380
            Width           =   1110
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74820
         TabIndex        =   30
         Top             =   4590
         Width           =   11295
         Begin VB.Timer timer1 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin VB.CommandButton CmdExit1 
            Caption         =   "E&xit"
            Height          =   645
            Left            =   11520
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   360
            Width           =   975
         End
         Begin prj_panji.vbButton cmdNew_Company 
            Height          =   705
            Left            =   540
            TabIndex        =   71
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
            MICON           =   "frm_mst_company.frx":1D79C
            PICN            =   "frm_mst_company.frx":1D7B8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Company 
            Height          =   705
            Left            =   1560
            TabIndex        =   72
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
            MICON           =   "frm_mst_company.frx":1E84A
            PICN            =   "frm_mst_company.frx":1E866
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Company 
            Height          =   705
            Left            =   2580
            TabIndex        =   73
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
            MICON           =   "frm_mst_company.frx":1F8F8
            PICN            =   "frm_mst_company.frx":1F914
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Company 
            Height          =   705
            Left            =   3600
            TabIndex        =   74
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
            MICON           =   "frm_mst_company.frx":209A6
            PICN            =   "frm_mst_company.frx":209C2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Company 
            Height          =   705
            Left            =   4620
            TabIndex        =   75
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
            MICON           =   "frm_mst_company.frx":21A54
            PICN            =   "frm_mst_company.frx":21A70
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Division 
         Height          =   3915
         Left            =   -74760
         TabIndex        =   42
         Top             =   510
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   6906
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
         Columns(1).Caption=   "KODE"
         Columns(1).DataField=   "division_code"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA"
         Columns(2).DataField=   "division_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KETERANGAN"
         Columns(3).DataField=   "description"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=4022"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3942"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=5821"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=5741"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=8784"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=8705"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
         Caption         =   "DAFTAR DIVISI"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Grade 
         Height          =   3915
         Left            =   -74790
         TabIndex        =   51
         Top             =   510
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   6906
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE"
         Columns(0).DataField=   "grade_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA"
         Columns(1).DataField=   "grade_name"
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
         Splits(0).ScrollBars=   3
         Splits(0).DividerColor=   13160660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2566"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=5741"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5662"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=10266"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=10186"
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
         Caption         =   "DAFTAR GRADE"
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
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Level 
         Height          =   3885
         Left            =   -74730
         TabIndex        =   65
         Top             =   540
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   6853
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE"
         Columns(0).DataField=   "level_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA"
         Columns(1).DataField=   "level_name"
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KETERANGAN"
         Columns(2).DataField=   "description"
         Columns(2).NumberFormat=   "FormatText Event"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3069"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2990"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=7064"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6985"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=513"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=8467"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=8387"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
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
         Caption         =   "DAFTAR LEVEL"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.alignment=0,.bgcolor=&H80000002&"
         _StyleDefs(10)  =   ":id=4,.fgcolor=&H80000009&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
         _StyleDefs(11)  =   ":id=4,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=4,.fontname=Tahoma"
         _StyleDefs(13)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(16)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(17)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(18)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(19)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(20)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(21)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(22)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(23)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(24)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(25)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(26)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(27)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(28)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&H80000002&,.fgcolor=&H80000009&"
         _StyleDefs(29)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
         _StyleDefs(30)  =   ":id=14,.fgcolor=&H80000002&"
         _StyleDefs(31)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(32)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(33)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(34)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(35)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(36)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(37)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(38)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(39)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(40)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin VB.Frame fra_entry_Company 
         Height          =   4125
         Left            =   -74820
         TabIndex        =   16
         Top             =   450
         Visible         =   0   'False
         Width           =   11295
         Begin VB.TextBox txt_pict_location 
            Appearance      =   0  'Flat
            Height          =   435
            Left            =   9030
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   104
            Top             =   1260
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.PictureBox pic 
            Height          =   1845
            Left            =   7560
            ScaleHeight     =   1785
            ScaleWidth      =   1335
            TabIndex        =   103
            Top             =   300
            Width           =   1395
            Begin VB.Image img 
               Height          =   1485
               Left            =   120
               Stretch         =   -1  'True
               Top             =   150
               Width           =   1095
            End
         End
         Begin VB.TextBox txtNPP 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7560
            MaxLength       =   30
            TabIndex        =   13
            Top             =   3270
            Width           =   3225
         End
         Begin VB.TextBox msk_npwp 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7560
            MaxLength       =   30
            TabIndex        =   10
            Top             =   2190
            Width           =   3225
         End
         Begin VB.TextBox msk_npwp_pimpinan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7560
            MaxLength       =   30
            TabIndex        =   12
            Top             =   2910
            Width           =   3225
         End
         Begin VB.TextBox txt_pimpinan 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   7560
            MaxLength       =   30
            TabIndex        =   11
            Top             =   2550
            Width           =   3225
         End
         Begin VB.TextBox txt_city_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1770
            Width           =   3495
         End
         Begin VB.TextBox txt_email_address 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   8
            Top             =   3600
            Width           =   3495
         End
         Begin VB.TextBox txt_web_address 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   7
            Top             =   3240
            Width           =   3495
         End
         Begin VB.TextBox txt_fax_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   6
            Top             =   2880
            Width           =   3495
         End
         Begin VB.TextBox txt_phone_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   5
            Top             =   2490
            Width           =   3495
         End
         Begin VB.TextBox txt_postal_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   4
            Top             =   2130
            Width           =   3495
         End
         Begin VB.TextBox txt_company_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1860
            MaxLength       =   50
            TabIndex        =   1
            Top             =   690
            Width           =   3495
         End
         Begin VB.TextBox txt_address 
            Appearance      =   0  'Flat
            Height          =   675
            Left            =   1860
            MaxLength       =   100
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   2
            Top             =   1050
            Width           =   3495
         End
         Begin VB.TextBox txt_company_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1860
            MaxLength       =   10
            TabIndex        =   0
            Top             =   330
            Width           =   1875
         End
         Begin prj_panji.vbButton cmd_brows_pict 
            Height          =   585
            Left            =   9030
            TabIndex        =   9
            Top             =   600
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   1032
            BTYPE           =   14
            TX              =   "Browse"
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
            MICON           =   "frm_mst_company.frx":22B02
            PICN            =   "frm_mst_company.frx":22B1E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label34 
            Caption         =   "(Max 1 Mb)"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   9000
            TabIndex        =   106
            Top             =   300
            Width           =   1245
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LOGO"
            Height          =   195
            Left            =   5790
            TabIndex        =   105
            Top             =   390
            Width           =   1620
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "N P P"
            Height          =   195
            Left            =   5850
            TabIndex        =   29
            Top             =   3330
            Width           =   1605
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NPWP PEMOTONG PAJAK"
            Height          =   585
            Left            =   5100
            TabIndex        =   28
            Top             =   2940
            Width           =   2355
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "PEMOTONG PAJAK"
            Height          =   195
            Left            =   5970
            TabIndex        =   27
            Top             =   2610
            Width           =   1470
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KOTA"
            Height          =   195
            Left            =   510
            TabIndex        =   26
            Top             =   1770
            Width           =   1125
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "N P W P"
            Height          =   195
            Left            =   5850
            TabIndex        =   25
            Top             =   2220
            Width           =   1560
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "EMAIL"
            Height          =   195
            Left            =   540
            TabIndex        =   24
            Top             =   3630
            Width           =   1140
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "WEBSITE"
            Height          =   195
            Left            =   540
            TabIndex        =   23
            Top             =   3270
            Width           =   1140
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NO. FAX"
            Height          =   195
            Left            =   510
            TabIndex        =   22
            Top             =   2910
            Width           =   1155
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NO. TELP."
            Height          =   195
            Left            =   570
            TabIndex        =   21
            Top             =   2520
            Width           =   1110
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE POS"
            Height          =   195
            Left            =   480
            TabIndex        =   20
            Top             =   2130
            Width           =   1155
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ALAMAT"
            Height          =   195
            Left            =   450
            TabIndex        =   19
            Top             =   1050
            Width           =   1215
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE*"
            Height          =   195
            Left            =   540
            TabIndex        =   18
            Top             =   330
            Width           =   1110
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA*"
            Height          =   195
            Left            =   450
            TabIndex        =   17
            Top             =   690
            Width           =   1185
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Company 
         Height          =   4095
         Left            =   -74820
         TabIndex        =   32
         Top             =   480
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   7223
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE"
         Columns(0).DataField=   "company_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA"
         Columns(1).DataField=   "company_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "ALAMAT"
         Columns(2).DataField=   "address"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KODE POS"
         Columns(3).DataField=   "postal_code"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "NO. TELP."
         Columns(4).DataField=   "phone_number"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "NO. FAX."
         Columns(5).DataField=   "fax_number"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "WEBSITE"
         Columns(6).DataField=   "web_address"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "EMAIL"
         Columns(7).DataField=   "email_address"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2143"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2064"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=6773"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=6694"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4630"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4551"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2355"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2540"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2461"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2566"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2487"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(5)._MinWidth=10"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=2566"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2487"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(37)=   "Column(6)._MinWidth=54215968"
         Splits(0)._ColumnProps(38)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(39)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(41)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(42)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(43)=   "Column(7)._MinWidth=54215968"
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
         Caption         =   "DAFTAR PERUSAHAAN"
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
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
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
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=102,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=110,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=107,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=108,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=109,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Perf 
         Height          =   3885
         Left            =   -74760
         TabIndex        =   96
         Top             =   540
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6853
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO"
         Columns(0).DataField=   "perf_number"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NILAI AWAL"
         Columns(1).DataField=   "perf_under"
         Columns(1).NumberFormat=   "Standard"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NILAI AKHIR"
         Columns(2).DataField=   "perf_upper"
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "GRADE"
         Columns(3).DataField=   "perf_grade"
         Columns(3).NumberFormat=   "Standard"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2381"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3784"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3704"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=514"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3916"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3836"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2355"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2275"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=6376"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=6297"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=512"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
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
         Caption         =   "PENILAIAN PERFORMA KARYAWAN"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Department 
         Height          =   3915
         Left            =   240
         TabIndex        =   120
         Top             =   540
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   6906
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
         Columns(1).Caption=   "KODE"
         Columns(1).DataField=   "department_code"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA"
         Columns(2).DataField=   "department_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "KETERANGAN"
         Columns(3).DataField=   "description"
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
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=7408"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=7329"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=7382"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=7303"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
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
         Caption         =   "DAFTAR DEPARTEMEN"
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER PERUSAHAAN"
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
      Left            =   180
      TabIndex        =   14
      Top             =   150
      Width           =   6315
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_company.frx":23BB0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12690
   End
End
Attribute VB_Name = "frm_mst_company"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsComp As New ADODB.Recordset
Dim rsDiv As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsDepartment As New ADODB.Recordset
Dim rsDivision As New ADODB.Recordset
Dim rsGrade As New ADODB.Recordset
Dim rsLevel As New ADODB.Recordset
Dim rsPerf As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim strsql As String
Dim src As String

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_validate_exist_new = False
    
    If SSTab1.Tab = 0 Then
        str_sql = "select count(company_code) as rec_count from m_company where company_code = '" _
        & Replace$(Trim$(txt_company_code), Chr$(39), Chr$(96)) & "'"
    ElseIf SSTab1.Tab = 1 Then
        str_sql = "select count(department_code) as rec_count from m_department where department_code = '" _
        & Replace$(Trim$(txt_department_code), Chr$(39), Chr$(96)) & "'"
    ElseIf SSTab1.Tab = 2 Then
        str_sql = "select count(division_code) as rec_count from m_division where division_code = '" _
        & Replace(Trim(txt_division_code), "'", "''") & "'"
    ElseIf SSTab1.Tab = 3 Then
        str_sql = "select count(grade_code) as rec_count from m_grade where grade_code = '" & Trim(txt_grade_code) & "'"
    ElseIf SSTab1.Tab = 4 Then
        str_sql = "select count(level_code) as rec_count from m_level where level_code = '" & Trim(txt_level_code) & "'"
    ElseIf SSTab1.Tab = 5 Then
        str_sql = "select count(perf_number) as rec_count from m_performance where perf_number = '" & Trim(txt_perf_number) & "'"
    End If
    
    rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
    
    If rs.Fields("rec_count").Value > 0 Then
        check_validate_exist_new = True
        Exit Function
    End If
End Function

Private Sub check_invalid()
    MsgBox "Data Sudah Ada...", vbCritical, headerMSG
    If SSTab1.Tab = 0 Then
        txt_company_code = ""
        If txt_company_code.Enabled = True Then txt_company_code.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        txt_department_code = ""
        If txt_department_code.Enabled = True Then txt_department_code.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        txt_division_code = ""
        If txt_division_code.Enabled = True Then txt_division_code.SetFocus
    ElseIf SSTab1.Tab = 3 Then
        txt_grade_code = ""
        If txt_grade_code.Enabled = True Then txt_grade_code.SetFocus
    ElseIf SSTab1.Tab = 4 Then
        txt_level_code = ""
        If txt_level_code.Enabled = True Then txt_level_code.SetFocus
    ElseIf SSTab1.Tab = 5 Then
        txt_perf_number = ""
        If txt_perf_number.Enabled = True Then txt_perf_number.SetFocus
    End If
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If SSTab1.Tab = 0 Then
        If Not txt_company_code = rsCompany.Fields("company_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        If Not txt_department_code = rsDepartment.Fields("department_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        If Not txt_division_code = rsDivision.Fields("division_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        If Not txt_grade_code.Text = rsGrade.Fields("grade_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 4 Then
        If Not txt_level_code.Text = rsLevel.Fields("level_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 5 Then
        If Not txt_perf_number.Text = rsPerf.Fields("perf_number").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    If SSTab1.Tab = 0 Then
        'validasi company code
        If Trim(txt_company_code) = "" Then
            MsgBox "Kode Perusahaan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_company_code.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi company name
        If Trim(txt_company_name) = "" Then
            MsgBox "Nama Perusahaan Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_company_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        'validasi department code
        If Trim(txt_department_code) = "" Then
            MsgBox "Kode Department Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_department_code.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi department name
        If Trim(txt_department_name) = "" Then
            MsgBox "Nama Department Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_department_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        'validasi division code
        If Trim(txt_division_code) = "" Then
            MsgBox "Kode Divisi Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_division_code.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi division name
        If Trim(txt_division_name) = "" Then
            MsgBox "Nama Divisi Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_division_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        'validasi grade code
        If Trim(txt_grade_code.Text) = "" Then
            MsgBox "Kode Grade Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_grade_code.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi grade name
        If Trim(txt_grade_name) = "" Then
            MsgBox "Nama Grade Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_grade_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 4 Then
        'validasi level code
        If Trim(txt_level_code.Text) = "" Then
            MsgBox "Kode Level Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_level_code.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi level name
        If Trim(txt_level_name) = "" Then
            MsgBox "Nama Level Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_level_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 5 Then
        'validasi performance number
        If Trim(txt_perf_number.Text) = "" Then
            MsgBox "No. Performa Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_perf_number.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi performance under
        If Trim(txt_perf_under.Text) = "" Then
            MsgBox "Nilai Awal Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_perf_under.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi performance upper
        If Trim(txt_perf_upper.Text) = "" Then
            MsgBox "Performance Upper Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_perf_upper.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi performance grade
        If Trim(txt_perf_grade.Text) = "" Then
            MsgBox "Performance Grade Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_perf_grade.SetFocus
            check_validate_new = False
            Exit Function
        End If
    End If
End Function

Private Sub load_data()
    If SSTab1.Tab = 0 Then
        If rsCompany.State Then rsCompany.Close
        SQL = "select * from m_company order by company_code"
        rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Company.DataSource = rsCompany
        
        If IsNull(TDBGrid_Company.Columns("company_code").Value) Then
            SSTab1.TabEnabled(1) = False
            SSTab1.TabEnabled(2) = False
            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(4) = False
            SSTab1.TabEnabled(5) = False
        Else
            SSTab1.TabEnabled(1) = True
            SSTab1.TabEnabled(2) = True
            SSTab1.TabEnabled(3) = True
            SSTab1.TabEnabled(4) = True
            SSTab1.TabEnabled(5) = True
        End If
    ElseIf SSTab1.Tab = 1 Then
        If rsDepartment.State Then rsDepartment.Close
        SQL = "select * from m_department " & _
                "where company_code = '" & TDBCombo_company.Text & "' " & _
                "order by department_code"
        rsDepartment.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Department.DataSource = rsDepartment
    ElseIf SSTab1.Tab = 2 Then
        If rsDivision.State Then rsDivision.Close
        SQL = "select * from m_division " & _
                "where company_code = '" & TDBCombo_company.Text & "' " & _
                "order by division_code"
        rsDivision.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Division.DataSource = rsDivision
    ElseIf SSTab1.Tab = 3 Then
        If rsGrade.State Then rsGrade.Close
        SQL = "select * from m_grade order by grade_code"
        rsGrade.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Grade.DataSource = rsGrade
    ElseIf SSTab1.Tab = 4 Then
        If rsLevel.State Then rsLevel.Close
        SQL = "select * from m_level order by level_code"
        rsLevel.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Level.DataSource = rsLevel
    ElseIf SSTab1.Tab = 5 Then
        If rsPerf.State Then rsPerf.Close
        SQL = "select * from m_performance order by perf_number"
        rsPerf.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Perf.DataSource = rsPerf
    End If
End Sub

Private Sub cancel_data()
    int_mode = 0
    Call load_mode
End Sub

Private Sub delete_data()
On Error GoTo Err
Dim i As Integer
    CnG.BeginTrans
            
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_Company.ApproxCount > 0 And TDBGrid_Company.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
    
        i = MsgBox("All data about this company will be deleted!" & Chr(13) & _
                    "Apakah Yakin Akan Menghapus Data '" _
                    & TDBGrid_Company.Columns("company_name").Value & "' ?", vbYesNo + vbCritical + vbQuestion, headerMSG)
        If Not i = vbYes Then
            CnG.CommitTrans
            Exit Sub
        End If
        
        CnG.Execute "delete from m_company where company_code = '" & TDBGrid_Company.Columns("company_code").Value & "'"
        '+++++++++++++++++++++++++++++++ Delete temp salary Proses +++++++++++++++++++++++++++++++++++++++++++++++
        CnG.Execute "delete from temp_sal_proses where company_code = '" & TDBGrid_Company.Columns("company_code").Value & "'"
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_Department.ApproxCount > 0 And TDBGrid_Department.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Department.Columns("department_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then
            CnG.CommitTrans
            Exit Sub
        End If
        
        CnG.Execute "delete from m_department where department_code = '" _
                    & Replace(TDBGrid_Department.Columns("department_code").Value, "'", "''") & "'"
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_Division.ApproxCount > 0 And TDBGrid_Division.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Division.Columns("division_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then
            CnG.CommitTrans
            Exit Sub
        End If
        
        CnG.Execute "delete from m_division where division_code = '" _
                        & Replace(TDBGrid_Division.Columns("division_code").Value, "'", "''") & "'"
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_Grade.ApproxCount > 0 And TDBGrid_Grade.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Grade.Columns("grade_code").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then
            CnG.CommitTrans
            Exit Sub
        End If
        
        CnG.Execute "delete from m_grade where grade_code = '" _
                        & Replace(TDBGrid_Grade.Columns("grade_code").Value, "'", "''") & "'"
    ElseIf SSTab1.Tab = 4 Then
        If Not (TDBGrid_Level.ApproxCount > 0 And TDBGrid_Level.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Level.Columns("level_code").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then
            CnG.CommitTrans
            Exit Sub
        End If
        
        CnG.Execute "delete from m_level where level_code = '" _
                        & Replace(TDBGrid_Level.Columns("level_code").Value, "'", "''") & "'"
    ElseIf SSTab1.Tab = 5 Then
        If Not (TDBGrid_Perf.ApproxCount > 0 And TDBGrid_Perf.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Perf.Columns("perf_grade").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then
            CnG.CommitTrans
            Exit Sub
        End If
        
        CnG.Execute "delete from m_performance where perf_number = '" _
                        & Replace(TDBGrid_Perf.Columns("perf_number").Value, "'", "''") & "'"
    End If
    
    CnG.CommitTrans
        
    Call load_data
    int_mode = 0
    Call load_mode
    Exit Sub
    
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Public Sub set_edit_data()
'On Error GoTo Err
    vSetData = 1
    
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_Company.ApproxCount > 0 And TDBGrid_Company.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsCompany
            txt_company_code = .Fields("company_code").Value
            txt_company_name = .Fields("company_name").Value
            txt_address = .Fields("address").Value
            txt_city_name = "" & .Fields("city_name").Value
            txt_postal_code = .Fields("postal_code").Value
            txt_phone_number = .Fields("phone_number").Value
            txt_fax_number = .Fields("fax_number").Value
            txt_web_address = .Fields("web_address").Value
            txt_email_address = .Fields("email_address").Value
            
            '------------------------------- Show Image ---------------------------
            SQL = "SELECT picture FROM m_company WHERE company_code = '" & .Fields("company_code").Value & "'"
            Set img.Picture = getImageFromDB(SQL)
            
            img.Width = img.Picture.Width
            img.Height = img.Picture.Height
            If pic.Width < img.Width Then
                img.Width = pic.Width
'                img.Height = img.Height / (img.Picture.Width / img.Width)
            End If

            If pic.Height < img.Height Then
                img.Height = pic.Height
'                img.Width = img.Width / (img.Picture.Height / img.Height)
            End If

            img.Left = 0
            img.Top = 0
            '---------------------------------------------------------------------
            
            msk_npwp = "" & .Fields("npwp").Value
            txt_pimpinan = .Fields("pimpinan").Value
            msk_npwp_pimpinan = .Fields("pimpinan_npwp").Value
            txtNPP = IIf(IsNull(.Fields("npp").Value), "", .Fields("npp").Value)
        End With
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_Department.ApproxCount > 0 And TDBGrid_Department.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsDepartment
            txt_department_code = .Fields("department_code").Value
            txt_department_name = .Fields("department_name").Value
            txt_department_description = IIf(IsNull(.Fields("description").Value), "", .Fields("description").Value)
        End With
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_Division.ApproxCount > 0 And TDBGrid_Division.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsDivision
            txt_division_code = .Fields("division_code").Value
            txt_division_name = .Fields("division_name").Value
            txt_division_description = IIf(IsNull(.Fields("description").Value), "", .Fields("description").Value)
        End With
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_Grade.ApproxCount > 0 And TDBGrid_Grade.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsGrade
            txt_grade_code = .Fields("grade_code").Value
            txt_grade_name = .Fields("grade_name").Value
            txt_grade_description = IIf(IsNull(.Fields("description").Value), "", .Fields("description").Value)
        End With
    ElseIf SSTab1.Tab = 4 Then
        If Not (TDBGrid_Level.ApproxCount > 0 And TDBGrid_Level.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsLevel
            txt_level_code = .Fields("level_code").Value
            txt_level_name = .Fields("level_name").Value
            txt_level_description = IIf(IsNull(.Fields("description").Value), "", .Fields("description").Value)
        End With
    ElseIf SSTab1.Tab = 5 Then
        If Not (TDBGrid_Perf.ApproxCount > 0 And TDBGrid_Perf.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsPerf
            txt_perf_number = .Fields("perf_number").Value
            txt_perf_under = FormatNumber(.Fields("perf_under").Value)
            txt_perf_upper = FormatNumber(.Fields("perf_upper").Value)
            txt_perf_grade = .Fields("perf_grade").Value
            txt_perf_description = IIf(IsNull(.Fields("description").Value), "", .Fields("description").Value)
        End With
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub cmd_brows_pict_Click()
Dim cls As New clsDlg
Dim i As Double
    
    src = cls.OpenFlDlg(Me.hwnd, "Images File(*.jpg)|*.jpg", "Open File", vbNullString, True)
    
    If src <> "" Then
        i = Round(FileLen(src) / 1024, 0)
        
        If i > 1024 Then
            MsgBox "Gambar Terlalu Besar...", vbExclamation, headerMSG
            Exit Sub
        Else
            img.Picture = LoadPicture(src)
    
            img.Width = img.Picture.Width
            img.Height = img.Picture.Height
            If pic.Width < img.Width Then
                img.Width = pic.Width
            End If
        
            If pic.Height < img.Height Then
                img.Height = pic.Height
            End If
        End If
    End If
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
        SQL = "INSERT INTO m_company(company_code,company_name," & _
                "address,city_name,postal_code,phone_number," & _
                "fax_number,web_address,email_address,picture,npwp,pimpinan," & _
                "pimpinan_npwp,npp,entry_date,entry_user) " & _
              "VALUES( " & _
                "'" & Trim(txt_company_code.Text) & "','" & Trim(txt_company_name.Text) & "'," & _
                "'" & Trim(txt_address.Text) & "','" & Trim(txt_city_name.Text) & "'," & _
                "'" & Trim(txt_postal_code.Text) & "','" & Trim(txt_phone_number.Text) & "'," & _
                "'" & Trim(txt_fax_number.Text) & "','" & Trim(txt_web_address.Text) & "'," & _
                "'" & Trim(txt_email_address.Text) & "','" & img & "','" & Trim(msk_npwp.Text) & "'," & _
                "'" & Trim(txt_pimpinan.Text) & "','" & Trim(msk_npwp_pimpinan.Text) & "'," & _
                "'" & Trim(txtNPP.Text) & "',Now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
        
        If fileExists(src) Then
            SQL = "SELECT company_code, picture FROM m_company WHERE company_code = '" & txt_company_code.Text & "'"
            If Not addImageToDB(SQL, src, "picture") Then MsgBox "Gambar Gagal Disimpan...", vbExclamation, headerMSG
        End If
        
        '+++++++++++ Insert into temp salary Proses +++++++++++++++++++++++++++++++++++++
        SQL = "INSERT into temp_sal_proses VALUES('" & Trim(txt_company_code) & "',0)"
        CnG.Execute SQL
        '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
        Call load_data_company
    ElseIf SSTab1.Tab = 1 Then
        SQL = "INSERT INTO m_department(company_code,department_code,department_name,description) " & _
              "VALUES( " & _
                "'" & TDBCombo_company.Text & "','" & Trim(txt_department_code.Text) & "'," & _
                "'" & Trim(txt_department_name.Text) & "','" & Trim(txt_department_description.Text) & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "INSERT INTO m_division(division_code,company_code," & _
                "division_name,description) " & _
              "VALUES( " & _
                "'" & Trim(txt_division_code.Text) & "','" & TDBCombo_company.Text & "'," & _
                "'" & Trim(txt_division_name.Text) & "'," & _
                "'" & Trim(txt_division_description.Text) & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "INSERT INTO m_grade(grade_code,grade_name,description," & _
                "date_entry,user_entry) " & _
              "VALUES( " & _
                "'" & Trim(txt_grade_code.Text) & "','" & Trim(txt_grade_name.Text) & "'," & _
                "'" & Trim(txt_grade_description.Text) & "',Now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 4 Then
        SQL = "INSERT INTO m_level(level_code,level_name,description," & _
                "date_entry,user_entry) " & _
              "VALUES( " & _
                "'" & Val(txt_level_code.Text) & "','" & Trim(txt_level_name.Text) & "'," & _
                "'" & Trim(txt_level_description.Text) & "',Now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 5 Then
        SQL = "INSERT INTO m_performance (perf_number,perf_under," & _
                "perf_upper,perf_grade,description,entry_date,entry_user) " & _
              "VALUES( " & _
                "'" & Trim(txt_perf_number.Text) & "','" & Val(DropAllComma(txt_perf_under)) & "'," & _
                "'" & Val(DropAllComma(txt_perf_upper)) & "','" & txt_perf_grade & "'," & _
                "'" & Trim(txt_perf_description) & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    End If
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub edit_old_data()
On Error GoTo Err

    CnG.BeginTrans
    
    If SSTab1.Tab = 0 Then
        SQL = "UPDATE m_company SET company_name = '" & Trim(txt_company_name.Text) & "'," & _
                "address = '" & Trim(txt_address.Text) & "',city_name = '" & Trim(txt_city_name.Text) & "'," & _
                "postal_code = '" & Trim(txt_postal_code.Text) & "',phone_number = '" & Trim(txt_phone_number.Text) & "'," & _
                "fax_number = '" & Trim(txt_fax_number.Text) & "',web_address = '" & Trim(txt_web_address.Text) & "'," & _
                "email_address = '" & Trim(txt_email_address.Text) & "',picture = '" & img & "',npwp = '" & Trim(msk_npwp.Text) & "'," & _
                "pimpinan = '" & Trim(txt_pimpinan.Text) & "',pimpinan_npwp = '" & Trim(msk_npwp_pimpinan.Text) & "'," & _
                "npp = '" & Trim(txtNPP.Text) & "',edit_date = Now(), edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE company_code = '" & Trim(txt_company_code.Text) & "'"
        CnG.Execute SQL
        
        If fileExists(src) Then
            SQL = "SELECT company_code, picture FROM m_company WHERE company_code = '" & rsCompany.Fields("company_code").Value & "'"
            If Not addImageToDB(SQL, src, "picture") Then MsgBox "Gambar Gagal Disimpan...", vbExclamation, headerMSG
        End If
    ElseIf SSTab1.Tab = 1 Then
        SQL = "UPDATE m_department SET department_name = '" & Trim(txt_department_name.Text) & "'," & _
                "description = '" & Trim(txt_department_description.Text) & "' " & _
              "WHERE company_code = '" & TDBCombo_company.Text & "' " & _
                "AND department_code = '" & Trim(txt_department_code.Text) & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "UPDATE m_division SET division_name = '" & Trim(txt_division_name.Text) & "'," & _
                "description = '" & Trim(txt_division_description.Text) & "' " & _
              "WHERE company_code = '" & TDBCombo_company.Text & "' " & _
                "AND division_code = '" & Trim(txt_division_code.Text) & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "UPDATE m_grade SET grade_name = '" & txt_grade_name & "'," & _
                "description = '" & Trim(txt_grade_description.Text) & "'," & _
                "date_edit = Now(),user_edit = '" & LOGIN_NAME & "' " & _
              "WHERE grade_code = '" & Trim(txt_grade_code.Text) & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 4 Then
        SQL = "UPDATE m_level SET level_name = '" & txt_level_name & "'," & _
                "description = '" & Trim(txt_level_description.Text) & "'," & _
                "date_edit = Now(),user_edit = '" & LOGIN_NAME & "' " & _
              "WHERE level_code = '" & Trim(txt_level_code.Text) & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 5 Then
        SQL = "UPDATE m_performance SET perf_number = '" & Trim(txt_perf_number.Text) & "'," & _
                "perf_under = '" & Val(DropAllComma(txt_perf_under.Text)) & "'," & _
                "perf_upper = '" & Val(DropAllComma(txt_perf_upper.Text)) & "'," & _
                "perf_grade = '" & txt_perf_grade.Text & "'," & _
                "description = '" & Trim(txt_perf_description.Text) & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_CODE & "' " & _
              "WHERE perf_number = '" & Trim(txt_perf_number.Text) & "'"
        CnG.Execute SQL
    End If
    
    CnG.CommitTrans
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
CnG.RollbackTrans
End Sub

Private Sub simpan_data()
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
    
    Call load_data
    int_mode = 0
    Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
    If SSTab1.Tab = 0 Then
        cmdNew_Company.Enabled = a And blnUser_Add
        cmdSave_Company.Enabled = b
        cmdEdit_Company.Enabled = c And blnUser_Edit
        cmdDelete_Company.Enabled = d And blnUser_Delete
        cmdCancel_Company.Enabled = e
    ElseIf SSTab1.Tab = 1 Then
        cmdNew_Department.Enabled = a And blnUser_Add
        cmdSave_Department.Enabled = b
        cmdEdit_Department.Enabled = c And blnUser_Edit
        cmdDelete_Department.Enabled = d And blnUser_Delete
        cmdCancel_Department.Enabled = e
    ElseIf SSTab1.Tab = 2 Then
        cmdNew_Division.Enabled = a And blnUser_Add
        cmdSave_Division.Enabled = b
        cmdEdit_Division.Enabled = c And blnUser_Edit
        cmdDelete_Division.Enabled = d And blnUser_Delete
        cmdCancel_Division.Enabled = e
    ElseIf SSTab1.Tab = 3 Then
        cmdNew_Grade.Enabled = a And blnUser_Add
        cmdSave_Grade.Enabled = b
        cmdEdit_Grade.Enabled = c And blnUser_Edit
        cmdDelete_Grade.Enabled = d And blnUser_Delete
        cmdCancel_Grade.Enabled = e
    ElseIf SSTab1.Tab = 4 Then
        cmdNew_Level.Enabled = a And blnUser_Add
        cmdSave_Level.Enabled = b
        cmdEdit_Level.Enabled = c And blnUser_Edit
        cmdDelete_Level.Enabled = d And blnUser_Delete
        cmdCancel_Level.Enabled = e
    ElseIf SSTab1.Tab = 5 Then
        cmdNew_Perf.Enabled = a And blnUser_Add
        cmdSave_Perf.Enabled = b
        cmdEdit_Perf.Enabled = c And blnUser_Edit
        cmdDelete_Perf.Enabled = d And blnUser_Delete
        cmdCancel_Perf.Enabled = e
    End If
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_company" Then Ctr.Text = ""
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
        '--------------------------- Employee Picture Default -----------------------
        txt_pict_location.Text = App.Path & "\default-company.jpg"
        img.Picture = LoadPicture(txt_pict_location.Text)
        src = txt_pict_location.Text
        
        img.Width = img.Picture.Width
        img.Height = img.Picture.Height
        If pic.Width < img.Width Then
            img.Width = pic.Width
'                img.Height = img.Height / (img.Picture.Width / img.Width)
        End If
        
        If pic.Height < img.Height Then
            img.Height = pic.Height
'                img.Width = img.Width / (img.Picture.Height / img.Height)
        End If
        
        img.Left = 0
        img.Top = 0
        '----------------------------------------------------------------------------
    End If
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        If SSTab1.Tab = 0 Then
            fra_entry_Company.Visible = True
            txt_company_code.Enabled = True
            TDBGrid_Company.Enabled = False
            
            Call set_new_data

            If txt_company_code.Enabled = True Then
                txt_company_code.SetFocus
            End If
        ElseIf SSTab1.Tab = 1 Then
            If TDBCombo_company.Text = "" Then
                MsgBox "Perusahaan Belum Dipilih...", vbExclamation, headerMSG
                TDBCombo_company.SetFocus
                
                int_mode = 0
                Call load_mode
                Exit Sub
            End If
            
            fra_entry_Department.Visible = True
            txt_department_code.Enabled = True
            TDBGrid_Department.Enabled = False
            
            Call set_new_data

            If txt_department_code.Enabled = True Then
                txt_department_code.SetFocus
            End If
        ElseIf SSTab1.Tab = 2 Then
            If TDBCombo_company.Text = "" Then
                MsgBox "Perusahaan Belum Dipilih...", vbExclamation, headerMSG
                TDBCombo_company.SetFocus
                
                int_mode = 0
                Call load_mode
                Exit Sub
            End If
        
            fra_entry_Division.Visible = True
            txt_division_code.Enabled = True
            TDBGrid_Division.Enabled = False
            
            Call set_new_data

            If txt_division_code.Enabled = True Then
                txt_division_code.SetFocus
            End If
        ElseIf SSTab1.Tab = 3 Then
            fra_entry_grade.Visible = True
            txt_grade_code.Enabled = True
            TDBGrid_Grade.Enabled = False
            
            Call set_new_data

            If txt_grade_code.Enabled = True Then
                txt_grade_code.SetFocus
            End If
        ElseIf SSTab1.Tab = 4 Then
            fra_entry_level.Visible = True
            txt_level_code.Enabled = True
            TDBGrid_Level.Enabled = False
            
            Call set_new_data

            If txt_level_code.Enabled = True Then
                txt_level_code.SetFocus
            End If
        ElseIf SSTab1.Tab = 5 Then
            fra_entry_perf.Visible = True
            txt_perf_number.Enabled = True
            TDBGrid_Perf.Enabled = False
            
            Call set_new_data

            If txt_perf_number.Enabled = True Then
                txt_perf_number.SetFocus
            End If
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        If SSTab1.Tab = 0 Then
            fra_entry_Company.Visible = False
            TDBGrid_Company.Enabled = True
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_Department.Visible = False
            TDBGrid_Department.Enabled = True
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_Division.Visible = False
            TDBGrid_Division.Enabled = True
        ElseIf SSTab1.Tab = 3 Then
            fra_entry_grade.Visible = False
            TDBGrid_Grade.Enabled = True
        ElseIf SSTab1.Tab = 4 Then
            fra_entry_level.Visible = False
            TDBGrid_Level.Enabled = True
        ElseIf SSTab1.Tab = 5 Then
            fra_entry_perf.Visible = False
            TDBGrid_Perf.Enabled = True
        End If
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        If SSTab1.Tab = 0 Then
            txt_company_code.Enabled = False
            fra_entry_Company.Visible = True
            TDBGrid_Company.Enabled = False
        ElseIf SSTab1.Tab = 1 Then
            txt_department_code.Enabled = False
            fra_entry_Department.Visible = True
            TDBGrid_Department.Enabled = False
        ElseIf SSTab1.Tab = 2 Then
            txt_division_code.Enabled = False
            fra_entry_Division.Visible = True
            TDBGrid_Division.Enabled = False
        ElseIf SSTab1.Tab = 3 Then
            txt_grade_code.Enabled = False
            fra_entry_grade.Visible = True
            TDBGrid_Grade.Enabled = False
        ElseIf SSTab1.Tab = 4 Then
            txt_level_code.Enabled = False
            fra_entry_level.Visible = True
            TDBGrid_Level.Enabled = False
        ElseIf SSTab1.Tab = 5 Then
            txt_perf_number.Enabled = False
            fra_entry_perf.Visible = True
            TDBGrid_Perf.Enabled = False
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

Private Sub Form_Load()
    SSTab1.TabVisible(1) = False
    SSTab1.Tab = 0
        
    Call load_data
    Call load_data_company
    
    Call load_data_user_access(Me)
    int_mode = 0

    Call load_mode
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 0 Then
        For Each Col In TDBGrid_Company.Columns
            Col.FilterText = ""
        Next Col
        rsCompany.Filter = adFilterNone
    ElseIf SSTab1.Tab = 1 Then
        For Each Col In TDBGrid_Department.Columns
            Col.FilterText = ""
        Next Col
        rsDepartment.Filter = adFilterNone
    ElseIf SSTab1.Tab = 2 Then
        For Each Col In TDBGrid_Division.Columns
            Col.FilterText = ""
        Next Col
        rsDivision.Filter = adFilterNone
    ElseIf SSTab1.Tab = 3 Then
        For Each Col In TDBGrid_Level.Columns
            Col.FilterText = ""
        Next Col
        rsLevel.Filter = adFilterNone
    ElseIf SSTab1.Tab = 4 Then
        For Each Col In TDBGrid_Grade.Columns
            Col.FilterText = ""
        Next Col
        rsGrade.Filter = adFilterNone
    ElseIf SSTab1.Tab = 5 Then
        For Each Col In TDBGrid_Perf.Columns
            Col.FilterText = ""
        Next Col
        rsPerf.Filter = adFilterNone
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

Private Sub grid_filter()
On Error GoTo Err

Dim i As Integer
    
    If SSTab1.Tab = 0 Then
        Set Cols = TDBGrid_Company.Columns
        i = TDBGrid_Company.Col
        TDBGrid_Company.HoldFields
        
        rsCompany.Filter = getFilter()
        TDBGrid_Company.Col = i
        TDBGrid_Company.EditActive = True
        
        TDBGrid_Company.SelStart = Len(TDBGrid_Company.Columns(i).FilterText)
        If TDBGrid_Company.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Company.Col = i
        End If
    ElseIf SSTab1.Tab = 1 Then
        Set Cols = TDBGrid_Department.Columns
        i = TDBGrid_Department.Col
        TDBGrid_Department.HoldFields
        
        rsDepartment.Filter = getFilter()
        TDBGrid_Department.Col = i
        TDBGrid_Department.EditActive = True
        
        TDBGrid_Department.SelStart = Len(TDBGrid_Department.Columns(i).FilterText)
        If TDBGrid_Department.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Department.Col = i
        End If
    ElseIf SSTab1.Tab = 2 Then
        Set Cols = TDBGrid_Division.Columns
        i = TDBGrid_Division.Col
        TDBGrid_Division.HoldFields
        
        rsDivision.Filter = getFilter()
        TDBGrid_Division.Col = i
        TDBGrid_Division.EditActive = True
        
        TDBGrid_Division.SelStart = Len(TDBGrid_Division.Columns(i).FilterText)
        If TDBGrid_Division.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Division.Col = i
        End If
    ElseIf SSTab1.Tab = 3 Then
        Set Cols = TDBGrid_Grade.Columns
        i = TDBGrid_Grade.Col
        TDBGrid_Grade.HoldFields
        
        rsGrade.Filter = getFilter()
        TDBGrid_Grade.Col = i
        TDBGrid_Grade.EditActive = True
        
        TDBGrid_Grade.SelStart = Len(TDBGrid_Grade.Columns(i).FilterText)
        If TDBGrid_Grade.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Grade.Col = i
        End If
    ElseIf SSTab1.Tab = 4 Then
        Set Cols = TDBGrid_Level.Columns
        i = TDBGrid_Level.Col
        TDBGrid_Level.HoldFields
        
        rsLevel.Filter = getFilter()
        TDBGrid_Level.Col = i
        TDBGrid_Level.EditActive = True
        
        TDBGrid_Level.SelStart = Len(TDBGrid_Level.Columns(i).FilterText)
        If TDBGrid_Level.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Level.Col = i
        End If
    ElseIf SSTab1.Tab = 5 Then
        Set Cols = TDBGrid_Perf.Columns
        i = TDBGrid_Perf.Col
        TDBGrid_Perf.HoldFields
        
        rsPerf.Filter = getFilter()
        TDBGrid_Perf.Col = i
        TDBGrid_Perf.EditActive = True
        
        TDBGrid_Perf.SelStart = Len(TDBGrid_Perf.Columns(i).FilterText)
        If TDBGrid_Perf.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Perf.Col = i
        End If
    End If
    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

'Private Sub set_data_department(ByVal str_code As String)
'    rsDiv.MoveFirst
'    rsDiv.Find ("department_code='" & str_code & "'") ', 0, adSearchForward, 1)
'    If Not (rsDiv.EOF = True Or rsDiv.BOF = True) Then
'        rsDiv.Bookmark = rsDiv.AbsolutePosition
'        Call TDBCombo_department_itemChange
'    Else
'        TDBCombo_department.Text = "": txt_department.Text = ""
'    End If
'End Sub


Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_company = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
'Dim i As Integer
'
'    If fra_entry_Company.Visible = True Or fra_entry_Department.Visible = True Or _
'        fra_entry_Department.Visible = True Or fra_entry_Division.Visible = True Or _
'        fra_entry_Grade.Visible = True Or fra_entry_level.Visible = True Or _
'        fra_entry_perf.Visible = True Then
'        i = MsgBox("Your Form Entry Is Active!" & Chr(13) & _
'                    "Are You Sure To Discard This Change?", vbYesNo + vbQuestion, headerMSG)
'        If i = vbNo Then Exit Sub
'    End If
    
    If SSTab1.Tab = 0 Or SSTab1.Tab = 3 Or SSTab1.Tab = 4 Or SSTab1.Tab = 5 Then
        fra_company.Visible = False
        Call load_data
    Else
        fra_company.Visible = True
        TDBCombo_company.Text = TDBGrid_Company.Columns("company_code").Value
        txt_company.Text = TDBGrid_Company.Columns("company_name").Value
        
        Call load_data
                    
'        If SSTab1.Tab = 2 Then
'            Call load_data_department
'        End If
    End If
    
    int_mode = 0
    Call load_mode
End Sub

'Private Sub load_data_department()
'    If rsDiv.State Then rsDiv.Close
'    SQL = "select * from m_department where company_code = '" _
'    & TDBCombo_company.Columns("company_code").Value & "' order by department_code"
'    rsDiv.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
'
'    TDBCombo_department.RowSource = rsDiv
'End Sub

Private Sub load_data_company()
    If rsComp.State Then rsComp.Close
    SQL = "select * from m_company order by company_code"
    rsComp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsComp
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company.Text = TDBCombo_company.Columns("company_name").Value
        
        Call load_data
        
'        If SSTab1.Tab = 2 Then
'            Call load_data_department
'        End If
    End If
End Sub

'Private Sub TDBCombo_department_itemChange()
'    If TDBCombo_department.ApproxCount > 0 Then
'        TDBCombo_department.Text = TDBCombo_department.Columns("department_code").Value
'        txt_department.Text = TDBCombo_department.Columns("department_name").Value
'    End If
'End Sub

Private Sub cmdNew_Company_Click()
    Call new_data
End Sub

Private Sub cmdSave_Company_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Company_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Company_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Company_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_department_Click()
    Call new_data
End Sub

Private Sub cmdSave_department_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_department_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_department_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_department_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_division_Click()
    Call new_data
End Sub

Private Sub cmdSave_division_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_division_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_division_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_division_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_grade_Click()
    Call new_data
End Sub

Private Sub cmdSave_grade_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_grade_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_grade_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_grade_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_level_Click()
    Call new_data
End Sub

Private Sub cmdSave_level_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_level_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_level_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_level_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_Perf_Click()
    Call new_data
End Sub

Private Sub cmdSave_Perf_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Perf_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Perf_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Perf_Click()
    Call cancel_data
End Sub


Private Sub TDBGrid_Company_FilterChange()
    Call grid_filter
End Sub
Private Sub TDBGrid_Department_FilterChange()
    Call grid_filter
End Sub
Private Sub TDBGrid_Division_FilterChange()
    Call grid_filter
End Sub
Private Sub TDBGrid_Grade_FilterChange()
    Call grid_filter
End Sub
Private Sub TDBGrid_Level_FilterChange()
    Call grid_filter
End Sub
Private Sub TDBGrid_Perf_FilterChange()
    Call grid_filter
End Sub


Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Call set_company_mode(rsComp, TDBCombo_company, txt_company)
End Sub

Private Sub Timer3_Timer()
    Timer3.Enabled = False
    Call set_company_mode(rsComp, TDBCombo_company, txt_company)
End Sub


Private Sub txt_company_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_company_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_department_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_department_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_division_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_division_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_grade_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_grade_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_level_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_level_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_perf_grade_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_perf_under_Validate(Cancel As Boolean)
    If Not Trim(txt_perf_under) = "" Then
        txt_perf_under = FormatNumber(DropAllComma(txt_perf_under))
    End If
End Sub

Private Sub txt_perf_upper_Validate(Cancel As Boolean)
    If Not Trim(txt_perf_upper) = "" Then
        txt_perf_upper = FormatNumber(DropAllComma(txt_perf_upper))
    End If
End Sub
