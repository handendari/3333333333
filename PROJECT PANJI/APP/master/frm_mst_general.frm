VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_working_time 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER SHIFT"
   ClientHeight    =   10575
   ClientLeft      =   -15
   ClientTop       =   300
   ClientWidth     =   12930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_mst_general.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   9045
      Left            =   120
      TabIndex        =   0
      Top             =   690
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   15954
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "SHIFT"
      TabPicture(0)   =   "frm_mst_general.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txt_group"
      Tab(0).Control(1)=   "fra_entry_WT"
      Tab(0).Control(2)=   "frmTombol"
      Tab(0).Control(3)=   "TDBGrid_WT"
      Tab(0).Control(4)=   "TDBCombo_Group"
      Tab(0).Control(5)=   "Label45"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "HARI KERJA"
      TabPicture(1)   =   "frm_mst_general.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_group_wd"
      Tab(1).Control(1)=   "fra_entry_WD"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "txt_working_time_name_wd"
      Tab(1).Control(4)=   "TDBGrid_WD"
      Tab(1).Control(5)=   "TDBCombo_working_time_wd"
      Tab(1).Control(6)=   "TDBCombo_group_wd"
      Tab(1).Control(7)=   "Label47"
      Tab(1).Control(8)=   "Label24"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "SHIFT KERJA KARYAWAN"
      TabPicture(2)   =   "frm_mst_general.frx":05C2
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label42"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label46"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "TDBCombo_group_wt"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "TDBCombo_company"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "TDBGrid_ListEmp"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "TDBGrid_EmpWT"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "fra_entry_empWT"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame3"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "txt_company_name"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txt_group_wt"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "GRUP SHIFT"
      TabPicture(3)   =   "frm_mst_general.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).Control(1)=   "fra_entry_group"
      Tab(3).Control(2)=   "TDBGrid_Group"
      Tab(3).ControlCount=   3
      Begin VB.TextBox txt_group_wd 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -71640
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   185
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox txt_group_wt 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   182
         Top             =   810
         Width           =   3855
      End
      Begin VB.TextBox txt_company_name 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   179
         Top             =   450
         Width           =   3855
      End
      Begin VB.TextBox txt_group 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -71670
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   176
         Top             =   540
         Width           =   3855
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   169
         Top             =   5640
         Width           =   12135
         Begin prj_panji.vbButton cmdNew_Group 
            Height          =   705
            Left            =   540
            TabIndex        =   170
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
            MICON           =   "frm_mst_general.frx":05FA
            PICN            =   "frm_mst_general.frx":0616
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Group 
            Height          =   705
            Left            =   1560
            TabIndex        =   171
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
            MICON           =   "frm_mst_general.frx":16A8
            PICN            =   "frm_mst_general.frx":16C4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Group 
            Height          =   705
            Left            =   2580
            TabIndex        =   172
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
            MICON           =   "frm_mst_general.frx":2756
            PICN            =   "frm_mst_general.frx":2772
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Group 
            Height          =   705
            Left            =   3600
            TabIndex        =   173
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
            MICON           =   "frm_mst_general.frx":3804
            PICN            =   "frm_mst_general.frx":3820
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Group 
            Height          =   705
            Left            =   4620
            TabIndex        =   174
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
            MICON           =   "frm_mst_general.frx":48B2
            PICN            =   "frm_mst_general.frx":48CE
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
      Begin VB.Frame fra_entry_group 
         Height          =   3975
         Left            =   -74760
         TabIndex        =   164
         Top             =   1410
         Width           =   12135
         Begin VB.TextBox txt_late_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5640
            MaxLength       =   50
            TabIndex        =   196
            Top             =   2790
            Width           =   1065
         End
         Begin VB.TextBox txt_time_tolerance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5640
            MaxLength       =   50
            TabIndex        =   195
            Top             =   2430
            Width           =   1065
         End
         Begin VB.TextBox txt_work_ot 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5640
            MaxLength       =   50
            TabIndex        =   194
            Top             =   2070
            Width           =   1065
         End
         Begin VB.TextBox txt_sat_ot 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5640
            MaxLength       =   50
            TabIndex        =   193
            Top             =   1710
            Width           =   1065
         End
         Begin VB.CheckBox chk_overtime 
            Caption         =   "NO"
            Height          =   255
            Left            =   5640
            TabIndex        =   192
            Top             =   1410
            Width           =   1005
         End
         Begin VB.CheckBox chk_rollable 
            Caption         =   "NO"
            Height          =   255
            Left            =   5640
            TabIndex        =   190
            Top             =   1140
            Width           =   1005
         End
         Begin VB.TextBox txt_group_description 
            Appearance      =   0  'Flat
            Height          =   615
            Left            =   5640
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   198
            Top             =   3150
            Width           =   4335
         End
         Begin VB.TextBox txt_group_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5640
            MaxLength       =   10
            TabIndex        =   188
            Top             =   420
            Width           =   1095
         End
         Begin VB.CommandButton Command2 
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
            TabIndex        =   165
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_group_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   5640
            MaxLength       =   50
            TabIndex        =   189
            Top             =   780
            Width           =   2655
         End
         Begin VB.Label Label62 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "POTONGAN KETERLAMBATAN"
            Height          =   195
            Left            =   3030
            TabIndex        =   207
            Top             =   2820
            Width           =   2145
         End
         Begin VB.Label Label61 
            Caption         =   "MENIT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6780
            TabIndex        =   206
            Top             =   2460
            Width           =   2925
         End
         Begin VB.Label Label60 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TOLERANSI KETERLAMBATAN"
            Height          =   195
            Left            =   3015
            TabIndex        =   205
            Top             =   2460
            Width           =   2145
         End
         Begin VB.Label Label59 
            Caption         =   "JAM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6780
            TabIndex        =   204
            Top             =   2130
            Width           =   405
         End
         Begin VB.Label Label53 
            Caption         =   "JAM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6780
            TabIndex        =   203
            Top             =   1770
            Width           =   405
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LEMBUR HARI BIASA OTOMATIS"
            Height          =   195
            Left            =   2820
            TabIndex        =   202
            Top             =   2100
            Width           =   2340
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LEMBUR SABTU OTOMATIS"
            Height          =   195
            Left            =   3195
            TabIndex        =   201
            Top             =   1740
            Width           =   1950
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DAPAT LEMBUR?"
            Height          =   195
            Left            =   3945
            TabIndex        =   197
            Top             =   1410
            Width           =   1200
         End
         Begin VB.Label Label48 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ROLLING?"
            Height          =   195
            Left            =   4425
            TabIndex        =   191
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label Label58 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA GRUP*"
            Height          =   195
            Left            =   4185
            TabIndex        =   168
            Top             =   780
            Width           =   975
         End
         Begin VB.Label Label57 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE GRUP*"
            Height          =   195
            Left            =   4215
            TabIndex        =   167
            Top             =   420
            Width           =   945
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   4110
            TabIndex        =   166
            Top             =   3150
            Width           =   990
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   180
         TabIndex        =   151
         Top             =   7500
         Width           =   12225
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew_empWT 
            Height          =   705
            Left            =   540
            TabIndex        =   156
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
            MICON           =   "frm_mst_general.frx":5960
            PICN            =   "frm_mst_general.frx":597C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_empWT 
            Height          =   705
            Left            =   1560
            TabIndex        =   157
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
            MICON           =   "frm_mst_general.frx":6A0E
            PICN            =   "frm_mst_general.frx":6A2A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_empWT 
            Height          =   705
            Left            =   2580
            TabIndex        =   158
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
            MICON           =   "frm_mst_general.frx":7ABC
            PICN            =   "frm_mst_general.frx":7AD8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_empWT 
            Height          =   705
            Left            =   3600
            TabIndex        =   159
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
            MICON           =   "frm_mst_general.frx":8B6A
            PICN            =   "frm_mst_general.frx":8B86
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmd_delete_dtl 
            Height          =   705
            Left            =   8340
            TabIndex        =   160
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Hapus Dtl"
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
            MICON           =   "frm_mst_general.frx":9C18
            PICN            =   "frm_mst_general.frx":9C34
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmd_add_dtl 
            Height          =   705
            Left            =   7350
            TabIndex        =   161
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Tambah Dtl"
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
            MICON           =   "frm_mst_general.frx":ACC6
            PICN            =   "frm_mst_general.frx":ACE2
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_empWT 
            Height          =   705
            Left            =   4620
            TabIndex        =   154
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
            MICON           =   "frm_mst_general.frx":BD74
            PICN            =   "frm_mst_general.frx":BD90
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
      Begin VB.Frame fra_entry_empWT 
         Height          =   1575
         Left            =   180
         TabIndex        =   137
         Top             =   1860
         Width           =   12195
         Begin VB.TextBox txt_working_time_name_emp 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   140
            Top             =   1080
            Width           =   3015
         End
         Begin VB.TextBox txt_shift_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1740
            TabIndex        =   162
            Top             =   1080
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.ComboBox cbo_type 
            ForeColor       =   &H80000002&
            Height          =   315
            ItemData        =   "frm_mst_general.frx":CE22
            Left            =   10170
            List            =   "frm_mst_general.frx":CE2C
            Locked          =   -1  'True
            TabIndex        =   139
            Text            =   "..."
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cbo_enable 
            Height          =   315
            ItemData        =   "frm_mst_general.frx":CE40
            Left            =   10170
            List            =   "frm_mst_general.frx":CE4A
            TabIndex        =   138
            Text            =   "..."
            Top             =   720
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_entry_date 
            Height          =   315
            Left            =   1920
            TabIndex        =   141
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   94633987
            CurrentDate     =   39270
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_working_time_emp 
            Height          =   375
            Left            =   1920
            OleObjectBlob   =   "frm_mst_general.frx":CE57
            TabIndex        =   142
            Top             =   720
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker_start_date 
            Height          =   315
            Left            =   6720
            TabIndex        =   143
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   94633987
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_end_date 
            Height          =   315
            Left            =   6720
            TabIndex        =   144
            Top             =   720
            Visible         =   0   'False
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "dd-MM-yyyy"
            Format          =   94633987
            CurrentDate     =   39270
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SHIFT"
            Height          =   195
            Left            =   1365
            TabIndex        =   150
            Top             =   720
            Width           =   435
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL INPUT"
            Height          =   195
            Left            =   1020
            TabIndex        =   149
            Top             =   360
            Width           =   765
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL MULAI"
            Height          =   195
            Left            =   5535
            TabIndex        =   148
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TGL BERAKHIR"
            Height          =   195
            Left            =   5430
            TabIndex        =   147
            Top             =   720
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TIPE"
            ForeColor       =   &H80000002&
            Height          =   195
            Left            =   9690
            TabIndex        =   146
            Top             =   360
            Width           =   330
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "ENABLE"
            Height          =   195
            Left            =   9210
            TabIndex        =   145
            Top             =   720
            Width           =   795
         End
      End
      Begin VB.Frame fra_entry_WD 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   113
         Top             =   3270
         Width           =   12135
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
            TabIndex        =   118
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.ComboBox cbo_active 
            Height          =   315
            ItemData        =   "frm_mst_general.frx":F25E
            Left            =   2040
            List            =   "frm_mst_general.frx":F268
            TabIndex        =   116
            Text            =   "..."
            Top             =   1080
            Width           =   1215
         End
         Begin VB.ComboBox cbo_working_day 
            Height          =   315
            ItemData        =   "frm_mst_general.frx":F275
            Left            =   2040
            List            =   "frm_mst_general.frx":F28E
            TabIndex        =   115
            Text            =   "..."
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txt_break_interval_minute_wd 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9960
            MaxLength       =   10
            TabIndex        =   114
            Top             =   1440
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker_start_time_wd 
            Height          =   315
            Left            =   6120
            TabIndex        =   117
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270.5
         End
         Begin MSComCtl2.DTPicker DTPicker_end_time_wd 
            Height          =   315
            Left            =   6120
            TabIndex        =   119
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_min_break_in_wd 
            Height          =   315
            Left            =   9960
            TabIndex        =   120
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270.5
         End
         Begin MSComCtl2.DTPicker DTPicker_max_break_out_wd 
            Height          =   315
            Left            =   9960
            TabIndex        =   121
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JAM PULANG"
            Height          =   195
            Left            =   5055
            TabIndex        =   128
            Top             =   1110
            Width           =   930
         End
         Begin VB.Label Label22 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "HARI KERJA"
            Height          =   195
            Left            =   720
            TabIndex        =   127
            Top             =   720
            Width           =   1125
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JAM MASUK"
            Height          =   195
            Left            =   5130
            TabIndex        =   126
            Top             =   750
            Width           =   855
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "AKTIF"
            Height          =   195
            Left            =   1395
            TabIndex        =   125
            Top             =   1080
            Width           =   435
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "INTERVAL"
            Height          =   195
            Left            =   9120
            TabIndex        =   124
            Top             =   1470
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JAM AKHIR ISTIRAHAT"
            Height          =   195
            Left            =   8160
            TabIndex        =   123
            Top             =   1080
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JAM AWAL ISTIRAHAT"
            Height          =   195
            Left            =   8190
            TabIndex        =   122
            Top             =   720
            Visible         =   0   'False
            Width           =   1635
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   112
         Top             =   5820
         Width           =   12135
         Begin prj_panji.vbButton cmdNew_WD 
            Height          =   705
            Left            =   540
            TabIndex        =   132
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
            MICON           =   "frm_mst_general.frx":F2D2
            PICN            =   "frm_mst_general.frx":F2EE
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_WD 
            Height          =   705
            Left            =   1560
            TabIndex        =   133
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
            MICON           =   "frm_mst_general.frx":10380
            PICN            =   "frm_mst_general.frx":1039C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_WD 
            Height          =   705
            Left            =   2580
            TabIndex        =   134
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
            MICON           =   "frm_mst_general.frx":1142E
            PICN            =   "frm_mst_general.frx":1144A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_WD 
            Height          =   705
            Left            =   3600
            TabIndex        =   135
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
            MICON           =   "frm_mst_general.frx":124DC
            PICN            =   "frm_mst_general.frx":124F8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_WD 
            Height          =   705
            Left            =   4620
            TabIndex        =   136
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
            MICON           =   "frm_mst_general.frx":1358A
            PICN            =   "frm_mst_general.frx":135A6
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
      Begin VB.TextBox txt_working_time_name_wd 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -71640
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   111
         Top             =   840
         Width           =   3855
      End
      Begin VB.Frame fra_entry_WT 
         Height          =   2415
         Left            =   -74790
         TabIndex        =   79
         Top             =   3360
         Width           =   12135
         Begin VB.ComboBox cbo_moving_number 
            Height          =   315
            ItemData        =   "frm_mst_general.frx":14638
            Left            =   2400
            List            =   "frm_mst_general.frx":14645
            TabIndex        =   200
            Text            =   "1"
            Top             =   1530
            Width           =   885
         End
         Begin VB.TextBox txt_moving_number 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   570
            MaxLength       =   10
            TabIndex        =   199
            Top             =   1800
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CheckBox chk_flag_moving 
            Height          =   255
            Left            =   2160
            TabIndex        =   86
            Top             =   1560
            Width           =   615
         End
         Begin VB.TextBox txt_break_interval_minute_wt 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   9720
            MaxLength       =   10
            TabIndex        =   85
            Top             =   1440
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.ComboBox cbo_tolerance 
            Height          =   315
            ItemData        =   "frm_mst_general.frx":14652
            Left            =   7560
            List            =   "frm_mst_general.frx":1465C
            TabIndex        =   84
            Text            =   "..."
            Top             =   2010
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cbo_day_over 
            Height          =   315
            ItemData        =   "frm_mst_general.frx":14669
            Left            =   6480
            List            =   "frm_mst_general.frx":14673
            TabIndex        =   83
            Text            =   "No"
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txt_shift_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   82
            Top             =   1080
            Width           =   2655
         End
         Begin VB.CommandButton CmdBrowse 
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
            TabIndex        =   81
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_shift_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   80
            Top             =   720
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker DTPicker_start_time_wt 
            Height          =   315
            Left            =   6480
            TabIndex        =   87
            Top             =   720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270.5
         End
         Begin MSComCtl2.DTPicker DTPicker_end_time_wt 
            Height          =   315
            Left            =   6480
            TabIndex        =   88
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_in_duration 
            Height          =   315
            Left            =   10560
            TabIndex        =   89
            Top             =   1920
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270.5
         End
         Begin MSComCtl2.DTPicker DTPicker_out_duration 
            Height          =   315
            Left            =   10560
            TabIndex        =   90
            Top             =   2160
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270
         End
         Begin MSComCtl2.DTPicker DTPicker_min_break_in_wt 
            Height          =   315
            Left            =   9720
            TabIndex        =   91
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270.5
         End
         Begin MSComCtl2.DTPicker DTPicker_max_break_out_wt 
            Height          =   315
            Left            =   9720
            TabIndex        =   92
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            MousePointer    =   99
            CustomFormat    =   "HH:mm"
            Format          =   94633987
            UpDown          =   -1  'True
            CurrentDate     =   39270
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SHIFT ROLLING? / NO"
            Height          =   195
            Left            =   480
            TabIndex        =   104
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "INTERVAL"
            Height          =   195
            Left            =   8910
            TabIndex        =   103
            Top             =   1470
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JAM AKHIR ISTIRAHAT"
            Height          =   195
            Left            =   7980
            TabIndex        =   102
            Top             =   1110
            Visible         =   0   'False
            Width           =   1665
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JAM AWAL ISTIRAHAT"
            Height          =   195
            Left            =   7680
            TabIndex        =   101
            Top             =   750
            Visible         =   0   'False
            Width           =   1965
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "TOLERANCE"
            Height          =   195
            Left            =   6120
            TabIndex        =   100
            Top             =   2010
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "OUT TOLERANCE"
            Height          =   195
            Left            =   9120
            TabIndex        =   99
            Top             =   2160
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "IN TOLERANCE"
            Height          =   195
            Left            =   9120
            TabIndex        =   98
            Top             =   1920
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "LEWAT HARI"
            Height          =   195
            Left            =   5490
            TabIndex        =   97
            Top             =   1440
            Width           =   930
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JAM MASUK"
            Height          =   195
            Left            =   4830
            TabIndex        =   96
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE SHIFT"
            Height          =   195
            Left            =   600
            TabIndex        =   95
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA SHIFT"
            Height          =   195
            Left            =   600
            TabIndex        =   94
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "JAM PULANG"
            Height          =   195
            Left            =   4905
            TabIndex        =   93
            Top             =   1080
            Width           =   1500
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74790
         TabIndex        =   78
         Top             =   6030
         Width           =   12135
         Begin prj_panji.vbButton cmdNew_WT 
            Height          =   705
            Left            =   540
            TabIndex        =   106
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
            MICON           =   "frm_mst_general.frx":14680
            PICN            =   "frm_mst_general.frx":1469C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_WT 
            Height          =   705
            Left            =   1560
            TabIndex        =   107
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
            MICON           =   "frm_mst_general.frx":1572E
            PICN            =   "frm_mst_general.frx":1574A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_WT 
            Height          =   705
            Left            =   2580
            TabIndex        =   108
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
            MICON           =   "frm_mst_general.frx":167DC
            PICN            =   "frm_mst_general.frx":167F8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_WT 
            Height          =   705
            Left            =   3600
            TabIndex        =   109
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
            MICON           =   "frm_mst_general.frx":1788A
            PICN            =   "frm_mst_general.frx":178A6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_WT 
            Height          =   705
            Left            =   4620
            TabIndex        =   110
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
            MICON           =   "frm_mst_general.frx":18938
            PICN            =   "frm_mst_general.frx":18954
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
      Begin VB.Frame fra_entry_dtl_pph 
         Height          =   2775
         Left            =   -74730
         TabIndex        =   63
         Top             =   2610
         Width           =   11295
         Begin VB.CheckBox chk_flag_over 
            Height          =   255
            Left            =   6120
            TabIndex        =   69
            Top             =   1200
            Width           =   375
         End
         Begin VB.TextBox txt_pph21_percentage 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   68
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txt_pph21_upper 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   67
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txt_pph_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   66
            Top             =   1920
            Width           =   3495
         End
         Begin VB.TextBox txt_pph21_under 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   65
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txt_pph21_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   64
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "PERCENTAGE"
            Height          =   195
            Left            =   4800
            TabIndex        =   75
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "UP"
            Height          =   195
            Left            =   4800
            TabIndex        =   74
            Top             =   1200
            Width           =   195
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "TO VALUE"
            Height          =   195
            Left            =   4800
            TabIndex        =   73
            Top             =   840
            Width           =   720
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   4800
            TabIndex        =   72
            Top             =   1920
            Width           =   1020
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "NO."
            Height          =   195
            Left            =   600
            TabIndex        =   71
            Top             =   480
            Width           =   285
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "FROM VALUE"
            Height          =   195
            Left            =   4800
            TabIndex        =   70
            Top             =   480
            Width           =   945
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74730
         TabIndex        =   55
         Top             =   5490
         Width           =   11295
         Begin prj_panji.vbButton cmdNew_Pph 
            Height          =   705
            Left            =   540
            TabIndex        =   56
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&New Dtl"
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
            MICON           =   "frm_mst_general.frx":199E6
            PICN            =   "frm_mst_general.frx":19A02
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Pph 
            Height          =   705
            Left            =   1560
            TabIndex        =   57
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Save Dtl"
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
            MICON           =   "frm_mst_general.frx":1AA94
            PICN            =   "frm_mst_general.frx":1AAB0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Pph 
            Height          =   705
            Left            =   2580
            TabIndex        =   58
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Edit Dtl"
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
            MICON           =   "frm_mst_general.frx":1BB42
            PICN            =   "frm_mst_general.frx":1BB5E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Pph 
            Height          =   705
            Left            =   3600
            TabIndex        =   59
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Delete Dtl"
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
            MICON           =   "frm_mst_general.frx":1CBF0
            PICN            =   "frm_mst_general.frx":1CC0C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Pph 
            Height          =   705
            Left            =   4620
            TabIndex        =   60
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Cancel Dtl"
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
            MICON           =   "frm_mst_general.frx":1DC9E
            PICN            =   "frm_mst_general.frx":1DCBA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton CmdNew_Master_Pph 
            Height          =   705
            Left            =   8850
            TabIndex        =   61
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":1ED4C
            PICN            =   "frm_mst_general.frx":1ED68
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_All_Pph 
            Height          =   705
            Left            =   9870
            TabIndex        =   62
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":1FDFA
            PICN            =   "frm_mst_general.frx":1FE16
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
         TabIndex        =   43
         Top             =   1650
         Width           =   11175
         Begin VB.TextBox txt_level_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   10
            TabIndex        =   46
            Top             =   810
            Width           =   1215
         End
         Begin VB.TextBox txt_level_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   50
            TabIndex        =   45
            Top             =   1230
            Width           =   3795
         End
         Begin VB.TextBox txt_level_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3630
            MaxLength       =   50
            TabIndex        =   44
            Top             =   1620
            Width           =   5145
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   1950
            TabIndex        =   49
            Top             =   1260
            Width           =   525
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   1950
            TabIndex        =   48
            Top             =   840
            Width           =   510
         End
         Begin VB.Label tlabel26 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   1950
            TabIndex        =   47
            Top             =   1650
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74790
         TabIndex        =   37
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
            TabIndex        =   38
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":20EA8
            PICN            =   "frm_mst_general.frx":20EC4
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
            TabIndex        =   39
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":21F56
            PICN            =   "frm_mst_general.frx":21F72
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
            TabIndex        =   40
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":23004
            PICN            =   "frm_mst_general.frx":23020
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
            TabIndex        =   41
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":240B2
            PICN            =   "frm_mst_general.frx":240CE
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
            TabIndex        =   42
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":25160
            PICN            =   "frm_mst_general.frx":2517C
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
         TabIndex        =   30
         Top             =   1770
         Width           =   11175
         Begin VB.TextBox txt_grade_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   33
            Top             =   1530
            Width           =   3495
         End
         Begin VB.TextBox txt_grade_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   50
            TabIndex        =   32
            Top             =   1170
            Width           =   3495
         End
         Begin VB.TextBox txt_grade_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4230
            MaxLength       =   10
            TabIndex        =   31
            Top             =   810
            Width           =   1695
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   2910
            TabIndex        =   36
            Top             =   1530
            Width           =   1095
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   2910
            TabIndex        =   35
            Top             =   810
            Width           =   510
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   2910
            TabIndex        =   34
            Top             =   1170
            Width           =   525
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   24
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
            TabIndex        =   25
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":2620E
            PICN            =   "frm_mst_general.frx":2622A
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
            TabIndex        =   26
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":272BC
            PICN            =   "frm_mst_general.frx":272D8
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
            TabIndex        =   27
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":2836A
            PICN            =   "frm_mst_general.frx":28386
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
            TabIndex        =   28
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":29418
            PICN            =   "frm_mst_general.frx":29434
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
            TabIndex        =   29
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":2A4C6
            PICN            =   "frm_mst_general.frx":2A4E2
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
         Height          =   2655
         Left            =   -74760
         TabIndex        =   14
         Top             =   1770
         Width           =   11175
         Begin VB.TextBox txt_division_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   18
            Top             =   1050
            Width           =   1695
         End
         Begin VB.TextBox txt_division_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   17
            Top             =   1410
            Width           =   3495
         End
         Begin VB.TextBox txt_division_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   16
            Top             =   1770
            Width           =   3495
         End
         Begin VB.TextBox txt_department 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   315
            Left            =   5400
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   15
            Top             =   660
            Width           =   3855
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_department 
            Height          =   375
            Left            =   3600
            OleObjectBlob   =   "frm_mst_general.frx":2B574
            TabIndex        =   19
            Top             =   660
            Width           =   1725
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   2280
            TabIndex        =   23
            Top             =   1410
            Width           =   525
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   2280
            TabIndex        =   22
            Top             =   1050
            Width           =   510
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   2280
            TabIndex        =   21
            Top             =   1770
            Width           =   1095
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "DEPARTMENT*"
            Height          =   195
            Left            =   2280
            TabIndex        =   20
            Top             =   690
            Width           =   1185
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74760
         TabIndex        =   8
         Top             =   4590
         Width           =   11205
         Begin VB.Timer Timer7 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew_Department 
            Height          =   705
            Left            =   540
            TabIndex        =   9
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":2D4DD
            PICN            =   "frm_mst_general.frx":2D4F9
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
            TabIndex        =   10
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":2E58B
            PICN            =   "frm_mst_general.frx":2E5A7
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
            TabIndex        =   11
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":2F639
            PICN            =   "frm_mst_general.frx":2F655
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
            TabIndex        =   12
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":306E7
            PICN            =   "frm_mst_general.frx":30703
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
            TabIndex        =   13
            Top             =   360
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
            MICON           =   "frm_mst_general.frx":31795
            PICN            =   "frm_mst_general.frx":317B1
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
      Begin VB.Frame fra_entry_Department 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   1
         Top             =   2280
         Width           =   11205
         Begin VB.TextBox txt_department_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   10
            TabIndex        =   4
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txt_department_name 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   3
            Top             =   960
            Width           =   3495
         End
         Begin VB.TextBox txt_department_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4860
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1320
            Width           =   3495
         End
         Begin VB.Label Label54 
            AutoSize        =   -1  'True
            Caption         =   "NAME*"
            Height          =   195
            Left            =   3480
            TabIndex        =   7
            Top             =   960
            Width           =   525
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            Caption         =   "CODE*"
            Height          =   195
            Left            =   3480
            TabIndex        =   6
            Top             =   600
            Width           =   510
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "DESCRIPTION"
            Height          =   195
            Left            =   3480
            TabIndex        =   5
            Top             =   1320
            Width           =   1095
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid3 
         Height          =   3915
         Left            =   -74760
         TabIndex        =   50
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
         Columns(1).Caption=   "department_code"
         Columns(1).DataField=   "department_code"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "DEPT. NAME"
         Columns(2).DataField=   "department_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "DIV. CODE"
         Columns(3).DataField=   "division_code"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "DIV. NAME"
         Columns(4).DataField=   "division_name"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DESCRIPTION"
         Columns(5).DataField=   "description"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
         Splits(0)._ColumnProps(17)=   "Column(2).Width=3969"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=3889"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(21)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(22)=   "Column(3).Width=2646"
         Splits(0)._ColumnProps(23)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._WidthInPix=2566"
         Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(27)=   "Column(4).Width=5741"
         Splits(0)._ColumnProps(28)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._WidthInPix=5662"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=6297"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=6218"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
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
         Caption         =   "LIST OF DIVISION"
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
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=82,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=79,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=80,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=81,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=86,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=83,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=84,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=85,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(58)  =   "Named:id=33:Normal"
         _StyleDefs(59)  =   ":id=33,.parent=0"
         _StyleDefs(60)  =   "Named:id=34:Heading"
         _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   ":id=34,.wraptext=-1"
         _StyleDefs(63)  =   "Named:id=35:Footing"
         _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=36:Selected"
         _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(67)  =   "Named:id=37:Caption"
         _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(69)  =   "Named:id=38:HighlightRow"
         _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=39:EvenRow"
         _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(73)  =   "Named:id=40:OddRow"
         _StyleDefs(74)  =   ":id=40,.parent=33"
         _StyleDefs(75)  =   "Named:id=41:RecordSelector"
         _StyleDefs(76)  =   ":id=41,.parent=34"
         _StyleDefs(77)  =   "Named:id=42:FilterBar"
         _StyleDefs(78)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid4 
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
         Columns(0).Caption=   "GRADE CODE"
         Columns(0).DataField=   "grade_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "GRADE NAME"
         Columns(1).DataField=   "grade_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "DESCRIPTION"
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
         Caption         =   "LIST OF DIVISION"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid5 
         Height          =   3885
         Left            =   -74730
         TabIndex        =   52
         Top             =   540
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   6853
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "LEVEL CODE"
         Columns(0).DataField=   "level_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "LEVEL NAME"
         Columns(1).DataField=   "level_name"
         Columns(1).NumberFormat=   "FormatText Event"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "DESCRIPTION"
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
         Caption         =   "LIST OF GROUP"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Department 
         Height          =   3915
         Left            =   -74760
         TabIndex        =   53
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
         Columns(3).Caption=   "DESCRIPTION"
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
         Caption         =   "LIST OF DEPARTMENT"
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
      Begin TrueOleDBList60.TDBCombo TDBCombo_pph 
         Height          =   375
         Left            =   -73740
         OleObjectBlob   =   "frm_mst_general.frx":32843
         TabIndex        =   54
         Top             =   630
         Width           =   1695
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_PPh 
         Height          =   4335
         Left            =   -74730
         TabIndex        =   76
         Top             =   1050
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NUMBER"
         Columns(0).DataField=   "pph21_number"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "FROM"
         Columns(1).DataField=   "pph21_under"
         Columns(1).NumberFormat=   "Standard"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "TO"
         Columns(2).DataField=   "pph21_upper"
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "UP"
         Columns(3).DataField=   "flag_over"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "PERCENTAGE"
         Columns(4).DataField=   "pph21_percentage"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "DESCRIPTION"
         Columns(5).DataField=   "description"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
         Splits(0)._ColumnProps(16)=   "Column(3).Width=1429"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1349"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2355"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2275"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=4948"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=4868"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=512"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
         Caption         =   "LIST OF PPh PASAL 21"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(58)  =   "Named:id=33:Normal"
         _StyleDefs(59)  =   ":id=33,.parent=0"
         _StyleDefs(60)  =   "Named:id=34:Heading"
         _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   ":id=34,.wraptext=-1"
         _StyleDefs(63)  =   "Named:id=35:Footing"
         _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=36:Selected"
         _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(67)  =   "Named:id=37:Caption"
         _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(69)  =   "Named:id=38:HighlightRow"
         _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=39:EvenRow"
         _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(73)  =   "Named:id=40:OddRow"
         _StyleDefs(74)  =   ":id=40,.parent=33"
         _StyleDefs(75)  =   "Named:id=41:RecordSelector"
         _StyleDefs(76)  =   ":id=41,.parent=34"
         _StyleDefs(77)  =   "Named:id=42:FilterBar"
         _StyleDefs(78)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_WT 
         Height          =   4665
         Left            =   -74790
         TabIndex        =   105
         Top             =   1110
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   8229
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
         Columns(2).Caption=   "JAM MASUK"
         Columns(2).DataField=   "start_time"
         Columns(2).NumberFormat=   "hh:nn"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "JAM PULANG"
         Columns(3).DataField=   "end_time"
         Columns(3).NumberFormat=   "hh:nn"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "MIN BREAK IN"
         Columns(4).DataField=   "min_break_in"
         Columns(4).NumberFormat=   "FormatText Event"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "MAX BREAK OUT"
         Columns(5).DataField=   "max_break_out"
         Columns(5).NumberFormat=   "FormatText Event"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "BREAK INTERVAL (M)"
         Columns(6).DataField=   "break_interval_minute"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "LEWAT HARI"
         Columns(7).DataField=   "flag_day_over"
         Columns(7).NumberFormat=   "FormatText Event"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   4
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TOLERANCE"
         Columns(8).DataField=   "flag_tolerance"
         Columns(8).NumberFormat=   "FormatText Event"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "IN TOLERANCE"
         Columns(9).DataField=   "start_time_tolerance"
         Columns(9).NumberFormat=   "FormatText Event"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "OUT TOLERANCE"
         Columns(10).DataField=   "end_time_tolerance"
         Columns(10).NumberFormat=   "FormatText Event"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   4
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "SHIFT ROLLING?"
         Columns(11).DataField=   "flag_moving"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "NO ROLLING"
         Columns(12).DataField=   "moving_number"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   13
         Splits(0)._UserFlags=   0
         Splits(0).Size  =   2
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=13"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3149"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3069"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=5794"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5715"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2514"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2434"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2487"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2408"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2461"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2381"
         Splits(0)._ColumnProps(24)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(26)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(27)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(4)._MinWidth=10"
         Splits(0)._ColumnProps(30)=   "Column(5).Width=2408"
         Splits(0)._ColumnProps(31)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(5)._WidthInPix=2328"
         Splits(0)._ColumnProps(33)=   "Column(5).AllowSizing=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(35)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(36)=   "Column(5).AllowFocus=0"
         Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(38)=   "Column(5)._MinWidth=10"
         Splits(0)._ColumnProps(39)=   "Column(6).Width=3122"
         Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=3043"
         Splits(0)._ColumnProps(42)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=513"
         Splits(0)._ColumnProps(44)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(45)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(46)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(47)=   "Column(6)._MinWidth=10"
         Splits(0)._ColumnProps(48)=   "Column(7).Width=2143"
         Splits(0)._ColumnProps(49)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(50)=   "Column(7)._WidthInPix=2064"
         Splits(0)._ColumnProps(51)=   "Column(7)._ColStyle=513"
         Splits(0)._ColumnProps(52)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(53)=   "Column(7)._MinWidth=10"
         Splits(0)._ColumnProps(54)=   "Column(8).Width=1931"
         Splits(0)._ColumnProps(55)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(8)._WidthInPix=1852"
         Splits(0)._ColumnProps(57)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(58)=   "Column(8)._ColStyle=513"
         Splits(0)._ColumnProps(59)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(60)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(61)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(62)=   "Column(8)._MinWidth=54215968"
         Splits(0)._ColumnProps(63)=   "Column(9).Width=2434"
         Splits(0)._ColumnProps(64)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(65)=   "Column(9)._WidthInPix=2355"
         Splits(0)._ColumnProps(66)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(67)=   "Column(9)._ColStyle=513"
         Splits(0)._ColumnProps(68)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(69)=   "Column(9).AllowFocus=0"
         Splits(0)._ColumnProps(70)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(71)=   "Column(10).Width=2540"
         Splits(0)._ColumnProps(72)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(73)=   "Column(10)._WidthInPix=2461"
         Splits(0)._ColumnProps(74)=   "Column(10).AllowSizing=0"
         Splits(0)._ColumnProps(75)=   "Column(10)._ColStyle=513"
         Splits(0)._ColumnProps(76)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(77)=   "Column(10).AllowFocus=0"
         Splits(0)._ColumnProps(78)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(79)=   "Column(10)._MinWidth=60129312"
         Splits(0)._ColumnProps(80)=   "Column(11).Width=2302"
         Splits(0)._ColumnProps(81)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(82)=   "Column(11)._WidthInPix=2223"
         Splits(0)._ColumnProps(83)=   "Column(11)._ColStyle=513"
         Splits(0)._ColumnProps(84)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(85)=   "Column(12).Width=1852"
         Splits(0)._ColumnProps(86)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(12)._WidthInPix=1773"
         Splits(0)._ColumnProps(88)=   "Column(12)._ColStyle=513"
         Splits(0)._ColumnProps(89)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(90)=   "Column(12)._MinWidth=119546496"
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
         Caption         =   "SHIFT"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=98,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=102,.parent=13,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=99,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=100,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=101,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=110,.parent=13,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=107,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=108,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=109,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=17"
         _StyleDefs(74)  =   "Splits(0).Columns(10).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(75)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
         _StyleDefs(76)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
         _StyleDefs(77)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
         _StyleDefs(78)  =   "Splits(0).Columns(11).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(79)  =   "Splits(0).Columns(11).HeadingStyle:id=55,.parent=14"
         _StyleDefs(80)  =   "Splits(0).Columns(11).FooterStyle:id=56,.parent=15"
         _StyleDefs(81)  =   "Splits(0).Columns(11).EditorStyle:id=57,.parent=17"
         _StyleDefs(82)  =   "Splits(0).Columns(12).Style:id=78,.parent=13,.alignment=2"
         _StyleDefs(83)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=14"
         _StyleDefs(84)  =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=15"
         _StyleDefs(85)  =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_WD 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   129
         Top             =   1350
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "SHIFT CODE"
         Columns(0).DataField=   "shift_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "KODE"
         Columns(1).DataField=   "day_code"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NAMA HARI"
         Columns(2).DataField=   "day_name"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "JAM MASUK"
         Columns(3).DataField=   "start_time"
         Columns(3).NumberFormat=   "hh:nn"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "JAM PULANG"
         Columns(4).DataField=   "end_time"
         Columns(4).NumberFormat=   "hh:nn"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   4
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "LEWAT HARI"
         Columns(5).DataField=   "flag_day_over"
         Columns(5).NumberFormat=   "FormatText Event"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "MIN BREAK IN"
         Columns(6).DataField=   "min_break_in"
         Columns(6).NumberFormat=   "FormatText Event"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "MAX BREAK OUT"
         Columns(7).DataField=   "max_break_out"
         Columns(7).NumberFormat=   "FormatText Event"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "INTERVAL (M)"
         Columns(8).DataField=   "break_interval_minute"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   4
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "AKTIF"
         Columns(9).DataField=   "flag_active"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).Size  =   2
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=3731"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=3651"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=6297"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=6218"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2699"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2619"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=2381"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2302"
         Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(34)=   "Column(5)._MinWidth=10"
         Splits(0)._ColumnProps(35)=   "Column(6).Width=2223"
         Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=2143"
         Splits(0)._ColumnProps(38)=   "Column(6).AllowSizing=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._ColStyle=513"
         Splits(0)._ColumnProps(40)=   "Column(6).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(6).AllowFocus=0"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(6)._MinWidth=54215968"
         Splits(0)._ColumnProps(44)=   "Column(7).Width=2381"
         Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2302"
         Splits(0)._ColumnProps(47)=   "Column(7).AllowSizing=0"
         Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
         Splits(0)._ColumnProps(49)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(50)=   "Column(7).AllowFocus=0"
         Splits(0)._ColumnProps(51)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(52)=   "Column(7)._MinWidth=54215968"
         Splits(0)._ColumnProps(53)=   "Column(8).Width=2117"
         Splits(0)._ColumnProps(54)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(8)._WidthInPix=2037"
         Splits(0)._ColumnProps(56)=   "Column(8).AllowSizing=0"
         Splits(0)._ColumnProps(57)=   "Column(8)._ColStyle=513"
         Splits(0)._ColumnProps(58)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(59)=   "Column(8).AllowFocus=0"
         Splits(0)._ColumnProps(60)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(61)=   "Column(8)._MinWidth=54215968"
         Splits(0)._ColumnProps(62)=   "Column(9).Width=2355"
         Splits(0)._ColumnProps(63)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(64)=   "Column(9)._WidthInPix=2275"
         Splits(0)._ColumnProps(65)=   "Column(9)._ColStyle=513"
         Splits(0)._ColumnProps(66)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(67)=   "Column(9)._MinWidth=54215968"
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
         Caption         =   "DAFTAR HARI KERJA"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=98,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=102,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=99,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=100,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=101,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
         _StyleDefs(70)  =   "Splits(0).Columns(9).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(71)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
         _StyleDefs(72)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
         _StyleDefs(73)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
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
      Begin TrueOleDBList60.TDBCombo TDBCombo_working_time_wd 
         Height          =   375
         Left            =   -73440
         OleObjectBlob   =   "frm_mst_general.frx":3479D
         TabIndex        =   130
         Top             =   840
         Width           =   1695
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_EmpWT 
         Height          =   2175
         Left            =   180
         TabIndex        =   152
         Top             =   1260
         Width           =   12165
         _ExtentX        =   21458
         _ExtentY        =   3836
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
         Columns(1).Caption=   "NO"
         Columns(1).DataField=   "shift_number"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KODE SHIFT"
         Columns(2).DataField=   "shift_code"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NAMA SHIFT"
         Columns(3).DataField=   "nm_shift"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "TANGGAL MULAI"
         Columns(4).DataField=   "start_date"
         Columns(4).NumberFormat=   "yyyy-MM-dd"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "END DATE"
         Columns(5).DataField=   "end_date"
         Columns(5).NumberFormat=   "yyyy-MM-dd"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   16
         Columns(6)._MaxComboItems=   5
         Columns(6).ValueItems(0)._DefaultItem=   0
         Columns(6).ValueItems(0).Value=   "0"
         Columns(6).ValueItems(0).Value.vt=   8
         Columns(6).ValueItems(0).DisplayValue=   "General"
         Columns(6).ValueItems(0).DisplayValue.vt=   8
         Columns(6).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems(1)._DefaultItem=   0
         Columns(6).ValueItems(1).Value=   "1"
         Columns(6).ValueItems(1).Value.vt=   8
         Columns(6).ValueItems(1).DisplayValue=   "Shift"
         Columns(6).ValueItems(1).DisplayValue.vt=   8
         Columns(6).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(6).ValueItems.Count=   2
         Columns(6).Caption=   "TIPE"
         Columns(6).DataField=   "flag_shift"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   4
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "ENABLE"
         Columns(7).DataField=   "flag_enable"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=1455"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1376"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=4101"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=4022"
         Splits(0)._ColumnProps(17)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=7250"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=7170"
         Splits(0)._ColumnProps(22)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(23)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(24)=   "Column(4).Width=2593"
         Splits(0)._ColumnProps(25)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(4)._WidthInPix=2514"
         Splits(0)._ColumnProps(27)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(29)=   "Column(5).Width=2514"
         Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2434"
         Splits(0)._ColumnProps(32)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(33)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(34)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(35)=   "Column(6).Width=3149"
         Splits(0)._ColumnProps(36)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(6)._WidthInPix=3069"
         Splits(0)._ColumnProps(38)=   "Column(6)._ColStyle=513"
         Splits(0)._ColumnProps(39)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(40)=   "Column(7).Width=2275"
         Splits(0)._ColumnProps(41)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(7)._WidthInPix=2196"
         Splits(0)._ColumnProps(43)=   "Column(7)._ColStyle=513"
         Splits(0)._ColumnProps(44)=   "Column(7).Order=8"
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
         Caption         =   "HEADER - DAFTAR SHIFT"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_ListEmp 
         Height          =   3855
         Left            =   180
         TabIndex        =   153
         Top             =   3540
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   6800
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
         Columns(2).Caption=   "DIV. CODE"
         Columns(2).DataField=   "division_code"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "DIVISI"
         Columns(3).DataField=   "division_name"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "NUMBER"
         Columns(4).DataField=   "shift_number"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "EMP. ID"
         Columns(5).DataField=   "employee_code"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "KODE KARY."
         Columns(6).DataField=   "nik"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "NAMA KARY."
         Columns(7).DataField=   "employee_name"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "TGL MULAI"
         Columns(8).DataField=   "start_date"
         Columns(8).NumberFormat=   "yyyy-MM-dd"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).Size  =   2
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
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
         Splits(0)._ColumnProps(17)=   "Column(2).Width=2778"
         Splits(0)._ColumnProps(18)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(2)._WidthInPix=2699"
         Splits(0)._ColumnProps(20)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(21)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(22)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(23)=   "Column(2)._MinWidth=74338528"
         Splits(0)._ColumnProps(24)=   "Column(3).Width=4895"
         Splits(0)._ColumnProps(25)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(3)._WidthInPix=4815"
         Splits(0)._ColumnProps(27)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(28)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(29)=   "Column(3)._MinWidth=74338544"
         Splits(0)._ColumnProps(30)=   "Column(4).Width=1217"
         Splits(0)._ColumnProps(31)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(4)._WidthInPix=1138"
         Splits(0)._ColumnProps(33)=   "Column(4).AllowSizing=0"
         Splits(0)._ColumnProps(34)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(35)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(36)=   "Column(4).AllowFocus=0"
         Splits(0)._ColumnProps(37)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(38)=   "Column(5).Width=3360"
         Splits(0)._ColumnProps(39)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(5)._WidthInPix=3281"
         Splits(0)._ColumnProps(41)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(42)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(43)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(44)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(45)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(47)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(48)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(49)=   "Column(7).Width=6879"
         Splits(0)._ColumnProps(50)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(7)._WidthInPix=6800"
         Splits(0)._ColumnProps(52)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(53)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(54)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(55)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(57)=   "Column(8)._ColStyle=513"
         Splits(0)._ColumnProps(58)=   "Column(8).Order=9"
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
         Caption         =   "DETAILS - DAFTAR KARYAWAN"
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
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
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
         _StyleDefs(66)  =   "Splits(0).Columns(8).Style:id=78,.parent=13,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
         _StyleDefs(70)  =   "Named:id=33:Normal"
         _StyleDefs(71)  =   ":id=33,.parent=0"
         _StyleDefs(72)  =   "Named:id=34:Heading"
         _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(74)  =   ":id=34,.wraptext=-1"
         _StyleDefs(75)  =   "Named:id=35:Footing"
         _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(77)  =   "Named:id=36:Selected"
         _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(79)  =   "Named:id=37:Caption"
         _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(81)  =   "Named:id=38:HighlightRow"
         _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(83)  =   "Named:id=39:EvenRow"
         _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(85)  =   "Named:id=40:OddRow"
         _StyleDefs(86)  =   ":id=40,.parent=33"
         _StyleDefs(87)  =   "Named:id=41:RecordSelector"
         _StyleDefs(88)  =   ":id=41,.parent=34"
         _StyleDefs(89)  =   "Named:id=42:FilterBar"
         _StyleDefs(90)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Group 
         Height          =   4665
         Left            =   -74760
         TabIndex        =   175
         Top             =   720
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   8229
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE GRUP"
         Columns(0).DataField=   "group_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA GRUP"
         Columns(1).DataField=   "group_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   4
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "ROLLING?"
         Columns(2).DataField=   "flag_rollable"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "LEMBUR?"
         Columns(3).DataField=   "flag_ot"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "TOLERANSI (MENIT)"
         Columns(4).DataField=   "time_tolerance"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "POT TRLMBT"
         Columns(5).DataField=   "late_value"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "KETERANGAN"
         Columns(6).DataField=   "description"
         Columns(6).NumberFormat=   "hh:nn"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).Size  =   2
         Splits(0).Size.vt=   2
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   13160660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2196"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2117"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3254"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=1614"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1535"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=513"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2249"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2170"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=513"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=513"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=5636"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=5556"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=513"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
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
         Caption         =   "DAFTAR GRUP SHIFT"
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
         _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=98,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=95,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=96,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=97,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(62)  =   "Named:id=33:Normal"
         _StyleDefs(63)  =   ":id=33,.parent=0"
         _StyleDefs(64)  =   "Named:id=34:Heading"
         _StyleDefs(65)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(66)  =   ":id=34,.wraptext=-1"
         _StyleDefs(67)  =   "Named:id=35:Footing"
         _StyleDefs(68)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(69)  =   "Named:id=36:Selected"
         _StyleDefs(70)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=37:Caption"
         _StyleDefs(72)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(73)  =   "Named:id=38:HighlightRow"
         _StyleDefs(74)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(75)  =   "Named:id=39:EvenRow"
         _StyleDefs(76)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(77)  =   "Named:id=40:OddRow"
         _StyleDefs(78)  =   ":id=40,.parent=33"
         _StyleDefs(79)  =   "Named:id=41:RecordSelector"
         _StyleDefs(80)  =   ":id=41,.parent=34"
         _StyleDefs(81)  =   "Named:id=42:FilterBar"
         _StyleDefs(82)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_Group 
         Height          =   375
         Left            =   -73470
         OleObjectBlob   =   "frm_mst_general.frx":366FF
         TabIndex        =   177
         Top             =   540
         Width           =   1695
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_company 
         Height          =   375
         Left            =   1290
         OleObjectBlob   =   "frm_mst_general.frx":38657
         TabIndex        =   180
         Top             =   450
         Width           =   1695
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_group_wt 
         Height          =   375
         Left            =   1290
         OleObjectBlob   =   "frm_mst_general.frx":3A5BD
         TabIndex        =   183
         Top             =   810
         Width           =   1695
      End
      Begin TrueOleDBList60.TDBCombo TDBCombo_group_wd 
         Height          =   375
         Left            =   -73440
         OleObjectBlob   =   "frm_mst_general.frx":3C518
         TabIndex        =   186
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "GRUP SHIFT"
         Height          =   195
         Left            =   -74610
         TabIndex        =   187
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "GRUP SHIFT"
         Height          =   195
         Left            =   300
         TabIndex        =   184
         Top             =   810
         Width           =   885
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "PERUSAHAAN"
         Height          =   195
         Left            =   210
         TabIndex        =   181
         Top             =   480
         Width           =   1005
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "GROUP SHIFT"
         Height          =   195
         Left            =   -74790
         TabIndex        =   178
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "SHIFT"
         Height          =   195
         Left            =   -73995
         TabIndex        =   131
         Top             =   840
         Width           =   435
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "PPH21 TYPE"
         Height          =   195
         Left            =   -74730
         TabIndex        =   77
         Top             =   660
         Width           =   870
      End
   End
   Begin prj_panji.vbButton cmdExit 
      Height          =   705
      Left            =   11610
      TabIndex        =   163
      Top             =   9780
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
      MICON           =   "frm_mst_general.frx":3E473
      PICN            =   "frm_mst_general.frx":3E48F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label43 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER SHIFT"
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
      Left            =   150
      TabIndex        =   155
      Top             =   150
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   585
      Left            =   0
      Picture         =   "frm_mst_general.frx":3F521
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12960
   End
End
Attribute VB_Name = "frm_mst_working_time"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsWT As New ADODB.Recordset
Dim rsWT_tdb As New ADODB.Recordset
Dim rsWD As New ADODB.Recordset
Dim rsCompany As New ADODB.Recordset
Dim rsEmpWT As New ADODB.Recordset
Dim rsWorkTime As New ADODB.Recordset
Dim rsListEmp As New ADODB.Recordset
Dim rsGroup As New ADODB.Recordset
Dim rsGroup_tdb As New ADODB.Recordset
Dim rsGroup_tdb_wt As New ADODB.Recordset
Dim rsGroup_tdb_wd As New ADODB.Recordset

Dim vStartDate As String
Dim vFlagRollable As Integer
Dim FlagNew As Boolean
Dim int_mode As Integer
Dim str_kode_rekening As String
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim SelBks As TrueOleDBGrid70.SelBookmarks

Private Sub check_invalid()
    MsgBox "Data Sudah Ada...", vbCritical, headerMSG
    
    If SSTab1.Tab = 0 Then
        txt_shift_code = ""
        If txt_shift_code.Enabled = True Then txt_shift_code.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        cbo_working_day.Text = ""
        If cbo_working_day.Enabled = True Then cbo_working_day.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        DTPicker_entry_date.Value = Now
        If DTPicker_entry_date.Enabled = True Then DTPicker_entry_date.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        txt_group_code = ""
        If txt_group_code.Enabled = True Then txt_group_code.SetFocus
    End If
End Sub

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset

    check_validate_exist_new = False
    
    If SSTab1.Tab = 0 Then
        SQL = "select count(shift_code) as rec_count from m_shift where shift_code = '" _
                & Replace$(Trim$(txt_shift_code), Chr$(39), Chr$(96)) & "'"

        rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
    
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        SQL = "select count(day_code) as rec_count from m_working_day where shift_code = '" _
                & TDBCombo_working_time_wd.Columns("shift_code").Value & "' and day_code=" _
                & cbo_working_day.ListIndex
        
        rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
    
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        SQL = "select count(group_code) as rec_count from m_shift_group where group_code = '" _
                & Trim(txt_group_code.Text) & "'"
        
        rs.Open SQL, CnG, adOpenStatic, adLockReadOnly
    
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            Exit Function
        End If
    End If
'    rs.Close
End Function

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If SSTab1.Tab = 0 Then
        If Not txt_shift_code = rsWT.Fields("shift_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        If Not txt_group_code = rsGroup.Fields("group_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    End If
    
End Function

Private Function check_validate_new() As Boolean
check_validate_new = True
    
    If SSTab1.Tab = 0 Then
        'validasi shift code
        If Trim(txt_shift_code) = "" Then
            MsgBox "Kode Shift Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_shift_code.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi shift name
        If Trim(txt_shift_name) = "" Then
            MsgBox "Nama Shift Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_shift_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        'validasi day code
        If cbo_working_day.ListIndex < 0 Then
            MsgBox "Kode Hari Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            cbo_working_day.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        'validasi working time tdbcombo
        If check_validate_tdbcombo(TDBCombo_working_time_emp) = False Then
            MsgBox "Shift Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_working_time_emp.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 3 Then
        'validasi group code
        If Trim(txt_group_code) = "" Then
            MsgBox "Kode Group Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_group_code.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi group name
        If Trim(txt_group_name) = "" Then
            MsgBox "Nama Group Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_group_name.SetFocus
            check_validate_new = False
            Exit Function
        End If
    End If

End Function

Private Sub load_data()
    If SSTab1.Tab = 0 Then
        If rsWT.State Then rsWT.Close
        SQL = "select * from m_shift where flag_shift=0 " & _
              "AND group_code = '" & TDBCombo_Group.Columns("group_code").Value & "' order by shift_code"
        rsWT.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_WT.DataSource = rsWT
    ElseIf SSTab1.Tab = 1 Then
        If rsWD.State Then rsWD.Close
        SQL = "select * from m_working_day where shift_code='" _
                & TDBCombo_working_time_wd.Columns("shift_code").Value & "' order by day_code"
        rsWD.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_WD.DataSource = rsWD
    ElseIf SSTab1.Tab = 2 Then
        If rsGroup_tdb_wt.State Then rsGroup_tdb_wt.Close
        SQL = "select * from m_shift_group order by group_code"
        rsGroup_tdb_wt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBCombo_group_wt.RowSource = rsGroup_tdb_wt
    ElseIf SSTab1.Tab = 3 Then
        If rsGroup.State Then rsGroup.Close
        SQL = "select * from m_shift_group order by group_code"
        rsGroup.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Group.DataSource = rsGroup
    End If
'    timer1.Enabled = True
End Sub

Private Sub load_data_company()
    If rsCompany.State Then rsCompany.Close
    SQL = "select * from m_company order by company_code"
    rsCompany.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_company.RowSource = rsCompany
End Sub

Private Sub load_data_shift()
    If rsWT_tdb.State Then rsWT_tdb.Close
    SQL = "select * from m_shift where group_code = '" & TDBCombo_group_wd.Text & "' order by shift_code"
    rsWT_tdb.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_working_time_wd.RowSource = rsWT_tdb
End Sub

Private Sub load_data_shift_group()
    If rsGroup_tdb.State Then rsGroup_tdb.Close
    SQL = "select * from m_shift_group order by group_code"
    rsGroup_tdb.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_Group.RowSource = rsGroup_tdb
End Sub

Private Sub load_data_shift_group_wt()
    If rsGroup_tdb_wt.State Then rsGroup_tdb_wt.Close
    SQL = "select * from m_shift_group order by group_code"
    rsGroup_tdb_wt.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_group_wt.RowSource = rsGroup_tdb_wt
End Sub

Private Sub load_data_shift_group_wd()
    If rsGroup_tdb_wd.State Then rsGroup_tdb_wd.Close
    SQL = "select * from m_shift_group order by group_code"
    rsGroup_tdb_wd.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_group_wd.RowSource = rsGroup_tdb_wd
End Sub

Private Sub load_data_header_working_time()
    If rsEmpWT.State Then rsEmpWT.Close
    SQL = "select * " & _
          "from (select b.group_code,b.shift_code,b.shift_name nm_shift, " & _
                  "(SELECT x.shift_date FROM tm_shift x WHERE x.shift_code = b.shift_code ORDER BY x.shift_number DESC LIMIT 1) shift_date," & _
                  "(SELECT x.shift_number FROM tm_shift x WHERE x.shift_code = b.shift_code ORDER BY x.shift_number DESC LIMIT 1) shift_number," & _
                  "(SELECT x.start_date FROM tm_shift x WHERE x.shift_code = b.shift_code ORDER BY x.shift_number DESC LIMIT 1) start_date," & _
                  "(SELECT x.end_date FROM tm_shift x WHERE x.shift_code = b.shift_code ORDER BY x.shift_number DESC LIMIT 1) end_date," & _
                  "(SELECT x.flag_shift FROM tm_shift x WHERE x.shift_code = b.shift_code ORDER BY x.shift_number DESC LIMIT 1) flag_shift," & _
                  "(SELECT x.flag_enable FROM tm_shift x WHERE x.shift_code = b.shift_code ORDER BY x.shift_number DESC LIMIT 1) flag_enable " & _
                "from m_shift b where b.group_code = '" & TDBCombo_group_wt.Text & "') aa " & _
          "where ifnull(shift_number,0) > 0 order by shift_code"
    rsEmpWT.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_EmpWT.DataSource = rsEmpWT
End Sub

Private Sub load_data_working_time()
    If rsWorkTime.State Then rsWorkTime.Close
    SQL = "select * from m_shift where flag_shift = " & _
            "'" & get_flag_shift & "' " & _
            "AND group_code = '" & TDBCombo_group_wt.Text & "' " & _
            "order by shift_code"
    rsWorkTime.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_working_time_emp.RowSource = rsWorkTime
End Sub

Public Sub load_data_detail_working_time()
    If rsListEmp.State Then rsListEmp.Close
    
    If rs.State Then rs.Close
    SQL = "SELECT flag_rollable FROM m_shift_group WHERE group_code = '" & TDBCombo_group_wt.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        vFlagRollable = rs!flag_rollable
    End If
    rs.Close
    
    If vFlagRollable = 0 Then
        SQL = "SELECT b.employee_code,b.nik,b.employee_name,a.shift_number," & _
                "a.company_code,c.company_name,b.division_code,e.division_name,a.start_date " & _
              "FROM td_shift2 a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                "JOIN m_company c ON b.company_code = c.company_code " & _
                "JOIN m_division e ON b.division_code = e.division_code and b.company_code = e.company_code " & _
              "WHERE a.company_code='" _
                & TDBCombo_company.Text & "' and a.shift_code='" _
                & TDBGrid_EmpWT.Columns("shift_code").Value & "' and a.group_code ='" _
                & TDBCombo_group_wt.Text & "' order by b.division_code, b.employee_name"
    Else
        SQL = "SELECT b.employee_code,b.nik,b.employee_name,a.shift_number," & _
                "a.company_code,c.company_name,b.division_code,e.division_name,'" & Format(TDBGrid_EmpWT.Columns("start_date").Value, "yyyy-MM-dd") & "' start_date " & _
              "FROM td_shift a JOIN m_employee b ON a.employee_code = b.employee_code " & _
                "JOIN m_company c ON b.company_code = c.company_code " & _
                "JOIN m_division e ON b.division_code = e.division_code and b.company_code = e.company_code " & _
              "WHERE a.company_code='" _
                & TDBCombo_company.Text & "' and a.shift_number='" _
                & TDBGrid_EmpWT.Columns("shift_number").Value & "' and a.group_code ='" _
                & TDBCombo_group_wt.Text & "' order by b.division_code, b.employee_name"
    End If
    
    rsListEmp.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBGrid_ListEmp.DataSource = rsListEmp
End Sub


'Private Sub cbo_tolerance_Click()
'    If cbo_tolerance.ListIndex = 0 Then
'        DTPicker_in_duration.Enabled = False
'        DTPicker_out_duration.Enabled = False
'    Else
'        DTPicker_in_duration.Enabled = True
'        DTPicker_out_duration.Enabled = True
'    End If
'End Sub

Private Sub cancel_data()
    int_mode = 0
    Call load_mode
End Sub

Private Sub delete_data()
Dim i As Integer
On Error GoTo Err
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_WT.ApproxCount > 0 And TDBGrid_WT.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_WT.Columns("shift_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_shift where shift_code = '" _
            & TDBGrid_WT.Columns("shift_code").Value & "'"
            
        CnG.CommitTrans
    
        Call load_data
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_WD.ApproxCount > 0 And TDBGrid_WD.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_WD.Columns("day_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_working_day where shift_code = '" _
            & TDBCombo_working_time_wd.Columns("shift_code").Value & "' and day_code=" _
            & TDBGrid_WD.Columns("day_code").Value
        CnG.CommitTrans
        
        Call load_data
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_EmpWT.ApproxCount > 0 And TDBGrid_EmpWT.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_EmpWT.Columns("nm_shift").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from td_shift where company_code = '" _
            & TDBCombo_company.Columns("company_code").Value & "' and shift_number = " _
            & TDBGrid_EmpWT.Columns("shift_number").Value & ""
            
        CnG.Execute "delete from tm_shift where company_code = '" _
            & TDBCombo_company.Columns("company_code").Value & "' and shift_number='" _
            & TDBGrid_EmpWT.Columns("shift_number").Value & "' and group_code = '" _
            & TDBCombo_group_wt.Text & "'"
            
        CnG.Execute "delete from td_shift2 where company_code = '" _
            & TDBCombo_company.Columns("company_code").Value & "' and shift_number = " _
            & TDBGrid_EmpWT.Columns("shift_number").Value & ""
        CnG.CommitTrans
        
        Call load_data_header_working_time
        
        TDBGrid_ListEmp.DataSource = Nothing
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_Group.ApproxCount > 0 And TDBGrid_Group.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Group.Columns("group_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_shift_group where group_code = '" _
            & TDBGrid_Group.Columns("group_code").Value & "'"
        CnG.CommitTrans
        
        Call load_data
    End If
    
    
    int_mode = 0
    Call load_mode
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub edit_data()
    int_mode = 2
    Call load_mode
End Sub

Private Sub chk_overtime_Click()
    If chk_overtime.Value = 0 Then
        chk_overtime.Caption = "NO"
    Else
        chk_overtime.Caption = "YES"
    End If
End Sub

'Private Sub chk_flag_moving_Click()
'    cbo_moving_number.SetFocus
'End Sub

Private Sub chk_rollable_Click()
    If chk_rollable.Value = 0 Then
        chk_rollable.Caption = "NO"
    Else
        chk_rollable.Caption = "YES"
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
Dim rs As New ADODB.Recordset

'On Error GoTo Err
    CnG.BeginTrans
    
    If SSTab1.Tab = 0 Then
        SQL = "INSERT INTO m_shift (group_code,shift_code,shift_name,start_time,end_time," & _
                "flag_day_over,flag_shift,min_break_in,max_break_out,break_interval_minute," & _
                "flag_moving,moving_number) " & _
              "VALUES ( " & _
                "'" & TDBCombo_Group.Text & "'," & _
                "'" & Trim(txt_shift_code.Text) & "'," & _
                "'" & Trim(txt_shift_name.Text) & "'," & _
                "'" & Format(DTPicker_start_time_wt.Value, "yyyy-MM-dd HH:mm:00") & "'," & _
                "'" & Format(DTPicker_end_time_wt.Value, "yyyy-MM-dd HH:mm:00") & "'," & _
                "'" & cbo_day_over.ListIndex & "'," & _
                "0,'" & Format(DTPicker_min_break_in_wt, "yyyy-MM-dd HH:mm:00") & "'," & _
                "'" & Format(DTPicker_max_break_out_wt, "yyyy-MM-dd HH:mm:00") & "'," & _
                "'" & Val(DropAllComma(txt_break_interval_minute_wt)) & "'," & _
                "'" & IIf(chk_flag_moving.Value = vbChecked, 1, 0) & "'," & _
                "'" & Val("" & cbo_moving_number.Text) & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "INSERT INTO m_working_day (group_code,shift_code,day_code,day_name,start_time," & _
                "end_time,flag_active,flag_day_over) " & _
              "SELECT '" & TDBCombo_group_wd.Text & "','" & TDBCombo_working_time_wd.Columns("shift_code").Value & "'," & _
                "'" & cbo_working_day.ListIndex & "','" & cbo_working_day.Text & "'," & _
                "'" & Format(DTPicker_start_time_wd.Value, "yyyy-MM-dd HH:mm:00") & "'," & _
                "'" & Format(DTPicker_end_time_wd.Value, "yyyy-MM-dd HH:mm:00") & "'," & _
                "'" & cbo_active.ListIndex & "',flag_day_over " & _
              "FROM m_shift WHERE shift_code = '" & TDBCombo_working_time_wd.Columns("shift_code").Value & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        Dim get_next_shift_number As Long
        SQL = "select ifnull(max(shift_number),0)+1 as cur_rec from tm_shift " & _
                "where company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                "AND group_code = '" & TDBCombo_group_wt.Text & "'"
        rs.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        If rs.RecordCount > 0 Then
            get_next_shift_number = rs!cur_rec
        End If
        rs.Close
        
        SQL = "INSERT INTO tm_shift (company_code,shift_number,group_code,shift_date,shift_code," & _
                "start_date,end_date,flag_shift,flag_enable,entry_date) " & _
              "VALUES( " & _
                "'" & TDBCombo_company.Columns("company_code").Value & "'," & _
                "'" & get_next_shift_number & "'," & _
                "'" & TDBCombo_group_wt.Text & "'," & _
                "'" & Format(DTPicker_entry_date.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "'" & TDBCombo_working_time_emp.Columns("shift_code").Value & "'," & _
                "'" & Format(DTPicker_start_date.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "'" & Format(DTPicker_end_date.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "'" & TDBCombo_working_time_emp.Columns("flag_shift").Value & "'," & _
                "'" & cbo_enable.ListIndex & "',now())"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "INSERT INTO m_shift_group (group_code,group_name,flag_rollable,flag_ot,saturday_ot,work_ot,time_tolerance,late_value,description,entry_date,entry_user) " & _
              "VALUES ( " & _
                "'" & Trim(txt_group_code.Text) & "'," & _
                "'" & Trim(txt_group_name.Text) & "'," & _
                "'" & chk_rollable.Value & "'," & _
                "'" & chk_overtime.Value & "'," & _
                "'" & Val(txt_sat_ot.Text) & "'," & _
                "'" & Val(txt_work_ot.Text) & "'," & _
                "'" & Val(txt_time_tolerance.Text) & "'," & _
                "'" & Val(DropAllComma(txt_late_value.Text)) & "'," & _
                "'" & Trim(txt_group_description.Text) & "'," & _
                "now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
        
        SQL = "UPDATE m_salary_standard a join td_shift b on a.employee_code = b.employee_code " & _
                "SET a.late_time_tolerance = (SELECT time_tolerance FROM m_shift_group WHERE group_code = '" & frm_mst_working_time.TDBCombo_group_wt.Text & "') " & _
              "WHERE b.group_code = '" & txt_group_code.Text & "'"
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
        SQL = "UPDATE m_shift SET shift_code = '" & Trim(txt_shift_code.Text) & "'," & _
                "shift_name = '" & Trim(txt_shift_name) & "'," & _
                "start_time = '" & Format(DTPicker_start_time_wt.Value, "yyyy-MM-dd HH:mm:00") & "'," & _
                "end_time = '" & Format(DTPicker_end_time_wt.Value, "yyyy-MM-dd HH:mm:00") & "'," & _
                "flag_day_over = '" & cbo_day_over.ListIndex & "'," & _
                "min_break_in = '" & Format(DTPicker_min_break_in_wt, "yyyy-MM-dd HH:mm:00") & "'," & _
                "max_break_out = '" & Format(DTPicker_max_break_out_wt, "yyyy-MM-dd HH:mm:00") & "'," & _
                "break_interval_minute = '" & Val(DropAllComma(txt_break_interval_minute_wt.Text)) & "'," & _
                "flag_moving = '" & IIf(chk_flag_moving.Value = vbChecked, 1, 0) & "'," & _
                "moving_number = '" & Val("" & cbo_moving_number) & "' " & _
              "WHERE shift_code = '" & Trim(txt_shift_code.Text) & "' " & _
                "AND group_code = '" & Trim(TDBCombo_Group.Text) & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "UPDATE m_working_day SET shift_code = '" & TDBCombo_working_time_wd.Columns("shift_code").Value & "'," & _
                "day_code = '" & cbo_working_day.ListIndex & "'," & _
                "day_name = '" & cbo_working_day.Text & "'," & _
                "start_time = '" & Format(DTPicker_start_time_wd.Value, "yyyy-MM-dd HH:mm:00") & "'," & _
                "end_time = '" & Format(DTPicker_end_time_wd.Value, "yyyy-MM-dd HH:mm:00") & "'," & _
                "flag_active = '" & cbo_active.ListIndex & "' " & _
              "WHERE shift_code = '" & TDBCombo_working_time_wd.Columns("shift_code").Value & "' " & _
                "AND day_code = '" & cbo_working_day.ListIndex & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "UPDATE tm_shift SET company_code = '" & TDBCombo_company.Columns("company_code").Value & "'," & _
                "shift_number = '" & txt_shift_number.Text & "'," & _
                "group_code = '" & TDBCombo_group_wt.Text & "'," & _
                "shift_date = '" & Format(DTPicker_entry_date.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "shift_code = '" & TDBCombo_working_time_emp.Columns("shift_code").Value & "'," & _
                "start_date = '" & Format(DTPicker_start_date.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "end_date = '" & Format(DTPicker_end_date.Value, "yyyy-MM-dd HH:mm:ss") & "'," & _
                "flag_shift = '" & TDBCombo_working_time_emp.Columns("flag_shift").Value & "'," & _
                "flag_enable = '" & cbo_enable.ListIndex & "' " & _
              "WHERE shift_number = '" & txt_shift_number.Text & "' " & _
                "AND company_code = '" & TDBCombo_company.Columns("company_code").Value & "' " & _
                "AND group_code = '" & TDBCombo_group_wt.Text & "' " & _
                "AND date(start_date) = '" & Format(vStartDate, "yyyy-MM-dd") & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 3 Then
        SQL = "UPDATE m_shift_group SET group_code = '" & Trim(txt_group_code.Text) & "'," & _
                "group_name = '" & Trim(txt_group_name.Text) & "'," & _
                "flag_rollable = '" & chk_rollable.Value & "'," & _
                "flag_ot = '" & chk_overtime.Value & "'," & _
                "saturday_ot = '" & Val(txt_sat_ot.Text) & "'," & _
                "work_ot = '" & Val(txt_work_ot.Text) & "'," & _
                "time_tolerance = '" & Val(txt_time_tolerance.Text) & "'," & _
                "late_value = '" & Val(DropAllComma(txt_late_value.Text)) & "'," & _
                "description = '" & Trim(txt_group_description.Text) & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_NAME & "' " & _
              "WHERE group_code = '" & TDBGrid_Group.Columns("group_code").Value & "'"
        CnG.Execute SQL
        
        SQL = "UPDATE m_salary_standard a join td_shift b on a.employee_code = b.employee_code " & _
                "SET a.late_time_tolerance = '" & Val(txt_time_tolerance.Text) & "' " & _
              "WHERE b.group_code = '" & txt_group_code.Text & "'"
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
        If check_validate_exist_new Then
            Call check_invalid: Exit Sub
        End If
        Call insert_new_data
    ElseIf int_mode = 2 Then
        If Not check_validate_new Then Exit Sub
        Call edit_old_data
    End If
    
    If SSTab1.Tab = 2 Then
        Call load_data_header_working_time
    Else
        Call load_data
    End If
    
    int_mode = 0
    Call load_mode
End Sub

Private Sub set_buttons_enable(ByVal a As Boolean, ByVal b As Boolean, ByVal c As Boolean, _
ByVal d As Boolean, ByVal e As Boolean, ByVal f As Boolean, ByVal g As Boolean)
    If SSTab1.Tab = 0 Then
        cmdNew_WT.Enabled = a And blnUser_Add
        cmdSave_WT.Enabled = b
        cmdEdit_WT.Enabled = c And blnUser_Edit
        cmdDelete_WT.Enabled = d And blnUser_Delete
        cmdCancel_WT.Enabled = e
    ElseIf SSTab1.Tab = 1 Then
        cmdNew_WD.Enabled = a And blnUser_Add
        cmdSave_WD.Enabled = b
        cmdEdit_WD.Enabled = c And blnUser_Edit
        cmdDelete_WD.Enabled = d And blnUser_Delete
        cmdCancel_WD.Enabled = e
    ElseIf SSTab1.Tab = 2 Then
        cmdNew_empWT.Enabled = a And blnUser_Add
        cmdSave_empWT.Enabled = b
        cmdEdit_empWT.Enabled = c And blnUser_Edit
        cmdDelete_empWT.Enabled = d And blnUser_Delete
        cmdCancel_empWT.Enabled = e
    ElseIf SSTab1.Tab = 3 Then
        cmdNew_Group.Enabled = a And blnUser_Add
        cmdSave_Group.Enabled = b
        cmdEdit_Group.Enabled = c And blnUser_Edit
        cmdDelete_Group.Enabled = d And blnUser_Delete
        cmdCancel_Group.Enabled = e
    End If
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    If SSTab1.Tab = 0 Or SSTab1.Tab = 3 Then
        For Each Ctr In Me
            If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
                If Not LCase(Ctr.name) = "txt_group" Then Ctr.Text = ""
            ElseIf TypeOf Ctr Is TDBCombo Then
                If Not LCase(Ctr.name) = "tdbcombo_group" Then Ctr.Text = ""
            ElseIf TypeOf Ctr Is DTPicker Then
                Ctr.Value = Now
            End If
        Next
    ElseIf SSTab1.Tab = 1 Then
        For Each Ctr In Me
            If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
                If Not LCase(Ctr.name) = "txt_working_time_name_wd" And _
                   Not LCase(Ctr.name) = "txt_group_wd" Then Ctr.Text = ""
            ElseIf TypeOf Ctr Is TDBCombo Then
                If Not LCase(Ctr.name) = "tdbcombo_working_time_wd" And _
                   Not LCase(Ctr.name) = "tdbcombo_group_wd" Then Ctr.Text = ""
            ElseIf TypeOf Ctr Is DTPicker Then
                Ctr.Value = Now
            End If
        Next
    ElseIf SSTab1.Tab = 2 Then
        For Each Ctr In Me
            If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
                If Not LCase(Ctr.name) = "txt_company_name" And _
                   Not LCase(Ctr.name) = "txt_group_wt" Then Ctr.Text = ""
            ElseIf TypeOf Ctr Is TDBCombo Then
                If Not LCase(Ctr.name) = "tdbcombo_company" And _
                   Not LCase(Ctr.name) = "tdbcombo_group_wt" Then Ctr.Text = ""
            ElseIf TypeOf Ctr Is DTPicker Then
                Ctr.Value = Now
            End If
        Next
    End If
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

Public Sub set_edit_data()
    vSetData = 1
            
    If SSTab1.Tab = 0 Then
        If Not (TDBGrid_WT.ApproxCount > 0 And TDBGrid_WT.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
    
        With rsWT
            
            txt_shift_code = .Fields("shift_code").Value
            '----------------------------------------------------------------------
            txt_shift_name = .Fields("shift_name").Value
            DTPicker_start_time_wt.Value = .Fields("start_time").Value
            DTPicker_end_time_wt.Value = .Fields("end_time").Value
            cbo_day_over.ListIndex = .Fields("flag_day_over").Value
            
            '-- additional 4 BPP
            DTPicker_min_break_in_wt = IIf(IsNull(.Fields("min_break_in").Value) = True, Now, .Fields("min_break_in").Value)
            DTPicker_max_break_out_wt = IIf(IsNull(.Fields("max_break_out").Value) = True, Now, .Fields("max_break_out").Value)
            txt_break_interval_minute_wt = Val("" & .Fields("break_interval_minute").Value)
            
            chk_flag_moving.Value = .Fields("flag_moving").Value
            cbo_moving_number.Text = .Fields("moving_number").Value
            
        '    cbo_tolerance.ListIndex = .Fields("flag_tolerance").Value
        '    If cbo_tolerance.ListIndex Then
        '        DTPicker_in_duration.Value = .Fields("start_time_tolerance").Value
        '        DTPicker_out_duration.Value = .Fields("end_time_tolerance").Value
        '    End If
        End With
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_WD.ApproxCount > 0 And TDBGrid_WD.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsWD
            '.Fields("shift_code").Value = TDBCombo_working_time.Columns("shift_code").Value
            cbo_working_day.ListIndex = .Fields("day_code").Value
            '-----------------------------------------------------------------------------
            DTPicker_start_time_wd.Value = .Fields("start_time").Value
            DTPicker_end_time_wd.Value = .Fields("end_time").Value
            cbo_active.ListIndex = .Fields("flag_active").Value
            
'            '-- additional 4 BPP
'            DTPicker_min_break_in = IIf(IsNull(.Fields("min_break_in").Value) = True, Now, .Fields("min_break_in").Value)
'            DTPicker_max_break_out = IIf(IsNull(.Fields("max_break_out").Value) = True, Now, .Fields("max_break_out").Value)
'            txt_break_interval_minute = Val("" & .Fields("break_interval_minute").Value)
        End With
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_EmpWT.ApproxCount > 0 And TDBGrid_EmpWT.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsEmpWT
'            .Fields("shift_number").Value = get_next_shift_number
            '-------------------------------------------------------------------------------
            DTPicker_entry_date.Value = .Fields("shift_date").Value
            Call set_data_shift(.Fields("shift_code").Value)
            DTPicker_start_date.Value = .Fields("start_date").Value
            DTPicker_end_date.Value = .Fields("end_date").Value
            txt_shift_number.Text = .Fields("shift_number").Value
            
            cbo_type.ListIndex = .Fields("flag_shift").Value
            cbo_enable.ListIndex = .Fields("flag_enable").Value
            
            vStartDate = .Fields("start_date").Value
        End With
    ElseIf SSTab1.Tab = 3 Then
        If Not (TDBGrid_Group.ApproxCount > 0 And TDBGrid_Group.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsGroup
            
            txt_group_code.Text = .Fields("group_code").Value
            '----------------------------------------------------------------------
            txt_group_name.Text = .Fields("group_name").Value
            chk_rollable.Value = .Fields("flag_rollable").Value
            chk_overtime.Value = .Fields("flag_ot").Value
            txt_sat_ot.Text = IIf(IsNull(.Fields("saturday_ot").Value), 0, .Fields("saturday_ot").Value)
            txt_work_ot.Text = IIf(IsNull(.Fields("work_ot").Value), 0, .Fields("work_ot").Value)
            txt_time_tolerance.Text = IIf(IsNull(.Fields("time_tolerance").Value), 0, .Fields("time_tolerance").Value)
            txt_late_value.Text = FormatNumber(IIf(IsNull(.Fields("late_value").Value), 0, .Fields("late_value").Value))
            txt_group_description.Text = .Fields("description").Value
            
            Call chk_rollable_Click
        End With
    End If
End Sub

Private Sub set_new_data()
    If SSTab1.Tab = 0 Then
        DTPicker_start_time_wt.Value = Format(Now, "yyyy-MM-dd ") & "08:30:00"
        DTPicker_end_time_wt.Value = Format(Now, "yyyy-MM-dd ") & "17:00:00"
'        DTPicker_in_duration.Value = Format(Now, "yyyy-MM-dd ") & "08:35:00"
'        DTPicker_out_duration.Value = Format(Now, "yyyy-MM-dd ") & "16:55:00"
'
'        cbo_tolerance.ListIndex = 0
        ' additional 4 BPP
        DTPicker_min_break_in_wt.Value = Format(Now, "yyyy-mm-dd 12:00:00")
        DTPicker_max_break_out_wt.Value = Format(Now, "yyyy-mm-dd 15:00:00")
        txt_break_interval_minute_wt.Text = "1"
        cbo_day_over.ListIndex = 0
        chk_flag_moving.Value = 0
    ElseIf SSTab1.Tab = 1 Then
        DTPicker_start_time_wd.Value = Format(Now, "yyyy-MM-dd ") & "08:30:00"
        DTPicker_end_time_wd.Value = Format(Now, "yyyy-MM-dd ") & "17:00:00"
        
        cbo_working_day.ListIndex = 0
        cbo_active.ListIndex = 1
    ElseIf SSTab1.Tab = 3 Then
        DTPicker_entry_date.Value = Now
        DTPicker_start_date.Value = Now
        DTPicker_end_date.Value = Now
        cbo_type.ListIndex = get_flag_shift
        chk_rollable.Value = 0
        chk_overtime.Value = 0
        chk_rollable.Caption = "NO"
        chk_overtime.Caption = "NO"
        cbo_enable.ListIndex = 1
    End If
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_WT.Visible = True
            txt_shift_code.Enabled = True
            TDBGrid_WT.Enabled = False
            Call set_new_data
            
            If txt_shift_code.Enabled = True Then
                txt_shift_code.SetFocus
            End If
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_WD.Visible = True
            cbo_working_day.Enabled = True
            TDBGrid_WD.Enabled = False
            Call set_new_data
            
            If cbo_working_day.Enabled = True Then
                cbo_working_day.SetFocus
            End If
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_empWT.Visible = True
            DTPicker_entry_date.Enabled = True
            TDBGrid_EmpWT.Enabled = False
            Call set_new_data
            
            If DTPicker_entry_date.Enabled = True Then
                DTPicker_entry_date.SetFocus
            End If
        ElseIf SSTab1.Tab = 3 Then
            fra_entry_group.Visible = True
            txt_group_code.Enabled = True
            TDBGrid_Group.Enabled = False
            Call set_new_data
            
            If txt_group_code.Enabled = True Then
                txt_group_code.SetFocus
            End If
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_WT.Visible = False
            TDBGrid_WT.Enabled = True
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_WD.Visible = False
            TDBGrid_WD.Enabled = True
        ElseIf SSTab1.Tab = 2 Then
            fra_entry_empWT.Visible = False
            TDBGrid_EmpWT.Enabled = True
        ElseIf SSTab1.Tab = 3 Then
            fra_entry_group.Visible = False
            TDBGrid_Group.Enabled = True
        End If
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        If SSTab1.Tab = 0 Then
            txt_shift_code.Enabled = False
            fra_entry_WT.Visible = True
            TDBGrid_WT.Enabled = False
        ElseIf SSTab1.Tab = 1 Then
            cbo_working_day.Enabled = False
            fra_entry_WD.Visible = True
            TDBGrid_WD.Enabled = False
        ElseIf SSTab1.Tab = 2 Then
            DTPicker_entry_date.Enabled = False
            fra_entry_empWT.Visible = True
            TDBGrid_EmpWT.Enabled = False
        ElseIf SSTab1.Tab = 3 Then
            txt_group_code.Enabled = False
            fra_entry_group.Visible = True
            TDBGrid_Group.Enabled = False
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
    
    SSTab1.Tab = 3
    Call load_data
    
    Call load_data_user_access(Me)
    int_mode = 0
    Call load_mode
'    timer1.Enabled = True
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 0 Then
        For Each Col In TDBGrid_WT.Columns
            Col.FilterText = ""
        Next Col
        rsWT.Filter = adFilterNone
    ElseIf SSTab1.Tab = 1 Then
        For Each Col In TDBGrid_WD.Columns
            Col.FilterText = ""
        Next Col
        rsWD.Filter = adFilterNone
    ElseIf SSTab1.Tab = 2 Then
        For Each Col In TDBGrid_EmpWT.Columns
            Col.FilterText = ""
        Next Col
        rsEmpWT.Filter = adFilterNone
    ElseIf SSTab1.Tab = 3 Then
        For Each Col In TDBGrid_Group.Columns
            Col.FilterText = ""
        Next Col
        rsGroup.Filter = adFilterNone
    End If
End Sub

Private Sub clear_filter_detail_employee()
    For Each Col In TDBGrid_ListEmp.Columns
        Col.FilterText = ""
    Next Col
    rsListEmp.Filter = adFilterNone
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
        Set Cols = TDBGrid_WT.Columns
        i = TDBGrid_WT.Col
        TDBGrid_WT.HoldFields
        
        rsWT.Filter = getFilter()
        TDBGrid_WT.Col = i
        TDBGrid_WT.EditActive = True
        
        TDBGrid_WT.SelStart = Len(TDBGrid_WT.Columns(i).FilterText)
        If TDBGrid_WT.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_WT.Col = i
        End If
    ElseIf SSTab1.Tab = 1 Then
        Set Cols = TDBGrid_WD.Columns
        i = TDBGrid_WD.Col
        TDBGrid_WD.HoldFields
        
        rsWD.Filter = getFilter()
        TDBGrid_WD.Col = i
        TDBGrid_WD.EditActive = True
        
        TDBGrid_WD.SelStart = Len(TDBGrid_WD.Columns(i).FilterText)
        If TDBGrid_WD.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_WD.Col = i
        End If
    ElseIf SSTab1.Tab = 2 Then
        Set Cols = TDBGrid_EmpWT.Columns
        i = TDBGrid_EmpWT.Col
        TDBGrid_EmpWT.HoldFields
        
        rsEmpWT.Filter = getFilter()
        TDBGrid_EmpWT.Col = i
        TDBGrid_EmpWT.EditActive = True
        
        TDBGrid_EmpWT.SelStart = Len(TDBGrid_EmpWT.Columns(i).FilterText)
        If TDBGrid_EmpWT.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_EmpWT.Col = i
        End If
    ElseIf SSTab1.Tab = 3 Then
        Set Cols = TDBGrid_Group.Columns
        i = TDBGrid_Group.Col
        TDBGrid_Group.HoldFields
        
        rsGroup.Filter = getFilter()
        TDBGrid_Group.Col = i
        TDBGrid_Group.EditActive = True
        
        TDBGrid_Group.SelStart = Len(TDBGrid_Group.Columns(i).FilterText)
        If TDBGrid_Group.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Group.Col = i
        End If
    End If
    
    Exit Sub

Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_working_time = Nothing
End Sub

Private Sub TDBGrid_ListEmp_FilterChange()
Dim i As Integer

On Error GoTo Err

    Set Cols = TDBGrid_ListEmp.Columns
    i = TDBGrid_ListEmp.Col
    TDBGrid_ListEmp.HoldFields
    
    rsListEmp.Filter = getFilter()
    TDBGrid_ListEmp.Col = i
    TDBGrid_ListEmp.EditActive = True
    
    TDBGrid_ListEmp.SelStart = Len(TDBGrid_ListEmp.Columns(i).FilterText)
    If TDBGrid_ListEmp.ApproxCount < 1 Then
        Call clear_filter_detail_employee
        TDBGrid_ListEmp.Col = i
    End If
    
    Exit Sub

Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter_detail_employee
End Sub

'Private Sub TDBGrid1_FormatText _
'(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
'If TDBGrid1.Columns(ColIndex).Caption = "IN" _
'Or TDBGrid1.Columns(ColIndex).Caption = "OUT" _
'Or TDBGrid1.Columns(ColIndex).Caption = "MIN BREAK IN" _
'Or TDBGrid1.Columns(ColIndex).Caption = "MAX BREAK OUT" Then
'    Value = Format(Value, "hh:nn")
'End If
'End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    int_mode = 0
    Call load_mode
    
    If SSTab1.Tab = 0 Then
        Call load_data_shift_group
        TDBGrid_WD.DataSource = Nothing
        TDBGrid_EmpWT.DataSource = Nothing
        TDBGrid_ListEmp.DataSource = Nothing
        TDBGrid_Group.DataSource = Nothing
    ElseIf SSTab1.Tab = 1 Then
        Call load_data_shift_group_wd
        TDBGrid_WT.DataSource = Nothing
        TDBGrid_EmpWT.DataSource = Nothing
        TDBGrid_ListEmp.DataSource = Nothing
        TDBGrid_Group.DataSource = Nothing
    ElseIf SSTab1.Tab = 2 Then
        Call load_data_company
        Timer1.Enabled = True
        TDBGrid_WT.DataSource = Nothing
        TDBGrid_WD.DataSource = Nothing
        TDBGrid_Group.DataSource = Nothing
    ElseIf SSTab1.Tab = 3 Then
        Call load_data
        TDBGrid_WD.DataSource = Nothing
        TDBGrid_WT.DataSource = Nothing
        TDBGrid_EmpWT.DataSource = Nothing
        TDBGrid_ListEmp.DataSource = Nothing
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Call set_company_mode(rsCompany, TDBCombo_company, txt_company_name)
End Sub

Private Sub TDBCombo_working_time_wd_ItemChange()
    If TDBCombo_working_time_wd.ApproxCount > 0 Then
        TDBCombo_working_time_wd.Text = TDBCombo_working_time_wd.Columns("shift_code").Value
        txt_working_time_name_wd.Text = TDBCombo_working_time_wd.Columns("shift_name").Value
        
        Call load_data
    End If
End Sub

Private Sub TDBCombo_Group_ItemChange()
    If TDBCombo_Group.ApproxCount > 0 Then
        TDBCombo_Group.Text = TDBCombo_Group.Columns("group_code").Value
        txt_group.Text = TDBCombo_Group.Columns("group_name").Value
        
        Call load_data
    End If
End Sub

Private Sub TDBCombo_Group_wt_ItemChange()
    If TDBCombo_group_wt.ApproxCount > 0 Then
        TDBCombo_group_wt.Text = TDBCombo_group_wt.Columns("group_code").Value
        txt_group_wt.Text = TDBCombo_group_wt.Columns("group_name").Value
        
        Call load_data_header_working_time
        If TDBGrid_EmpWT.ApproxCount > 0 Then Call TDBGrid_EmpWT_RowColChange(-1, 0)
        Call load_data_working_time
    End If
End Sub

Private Sub TDBCombo_Group_wd_ItemChange()
    If TDBCombo_group_wd.ApproxCount > 0 Then
        TDBCombo_group_wd.Text = TDBCombo_group_wd.Columns("group_code").Value
        txt_group_wd.Text = TDBCombo_group_wd.Columns("group_name").Value
        
        Call load_data_shift
    End If
End Sub

Private Sub TDBCombo_company_ItemChange()
    If TDBCombo_company.ApproxCount > 0 Then
        TDBCombo_company.Text = TDBCombo_company.Columns("company_code").Value
        txt_company_name = TDBCombo_company.Columns("company_name").Value
        
        TDBGrid_EmpWT.Close: TDBGrid_ListEmp.Close
        
        Call load_data
    End If
End Sub

Private Sub TDBCombo_working_time_emp_ItemChange()
    If TDBCombo_working_time_emp.ApproxCount > 0 Then
        TDBCombo_working_time_emp.Text = TDBCombo_working_time_emp.Columns("shift_code").Value
        txt_working_time_name_emp.Text = TDBCombo_working_time_emp.Columns("shift_name").Value
        cbo_type.ListIndex = TDBCombo_working_time_emp.Columns("flag_shift").Value
    End If
End Sub

Private Sub TDBGrid_EmpWT_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If (TDBGrid_EmpWT.Row + 1) > 0 And (TDBGrid_EmpWT.Row + 1) <> LastRow Then
        'MsgBox "LETS..."
        Call load_data_detail_working_time
    End If
End Sub

Private Sub set_data_shift(ByVal str_code As String)
    rsWorkTime.MoveFirst
    rsWorkTime.Find ("shift_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (rsWorkTime.EOF = True Or rsWorkTime.BOF = True) Then
        TDBCombo_working_time_emp.Bookmark = rsWorkTime.AbsolutePosition
        Call TDBCombo_working_time_emp_ItemChange
    Else
        TDBCombo_working_time_emp.Text = ""
    End If
End Sub

Private Function get_flag_shift() As Integer
Dim rs As New ADODB.Recordset

    rs.Open "select flag_shift from s_working_time where company_code='" _
            & TDBCombo_company.Columns("company_code").Value & "'", CnG, adOpenStatic, adLockReadOnly
    
    If rs.RecordCount > 0 Then
        get_flag_shift = rs.Fields("flag_shift").Value
    Else
        'MsgBox "Setting data isn't found for this company!", vbCritical, headerMSG
        CnG.Execute "insert into s_working_time(company_code, flag_shift) values('" _
            & TDBCombo_company.Columns("company_code").Value & "',0)"
        get_flag_shift = 0
    End If
    rs.Close
End Function

Private Sub cmdNew_WT_Click()
    Call new_data
End Sub

Private Sub cmdSave_WT_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_WT_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_WT_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_WT_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_WD_Click()
    Call new_data
End Sub

Private Sub cmdSave_WD_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_WD_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_WD_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_WD_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_empWT_Click()
    Call new_data
End Sub

Private Sub cmdSave_empWT_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_empWT_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_empWT_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_empWT_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_Group_Click()
    Call new_data
End Sub

Private Sub cmdSave_Group_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Group_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Group_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Group_Click()
    Call cancel_data
End Sub

Private Sub cmd_add_dtl_Click()
    If Not TDBGrid_EmpWT.ApproxCount > 0 Then
        MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
        Exit Sub
    End If
    
    frm_lookup_mst_employee.public_int_mode = 0
    frm_lookup_mst_employee.public_str_company_code = TDBCombo_company.Columns("company_code").Value
    frm_lookup_mst_employee.DTPicker_start_date.Visible = True
    frm_lookup_mst_employee.lbl_start_date.Visible = True
    frm_lookup_mst_employee.DTPicker_start_date.Value = TDBGrid_EmpWT.Columns("start_date").Value
    frm_lookup_mst_employee.Show 1
End Sub

Private Sub cmd_delete_dtl_Click()
Dim i As Integer
Dim item
    
On Error GoTo Err
    If Not TDBGrid_ListEmp.ApproxCount > 0 Then
        Exit Sub
    End If
        
    Set SelBks = TDBGrid_ListEmp.SelBookmarks
    i = MsgBox("Apakah Yakin Akan Menghapus " _
        & SelBks.Count & " employee's data ?", vbYesNo + vbQuestion, headerMSG)
    If Not i = vbYes Then Exit Sub
        
    SQL = "SELECT flag_rollable FROM m_shift_group WHERE group_code = '" & TDBCombo_group_wt.Text & "'"
    rs.Open SQL, CnG, adOpenForwardOnly
    
    If rs.RecordCount > 0 Then
        vFlagRollable = rs!flag_rollable
    End If
    rs.Close
                
    i = 0
    CnG.BeginTrans
    For Each item In SelBks
        i = i + 1
        
        'MsgBox TDBGrid2.Columns("employee_name").CellText(item)
        
        If vFlagRollable = 1 Then
            CnG.Execute "delete from td_shift where company_code = '" _
                & TDBGrid_ListEmp.Columns("company_code").CellText(item) & "' and employee_code = '" _
                & TDBGrid_ListEmp.Columns("employee_code").CellText(item) & "'"
        Else
            CnG.Execute "delete from td_shift where company_code = '" _
                & TDBGrid_ListEmp.Columns("company_code").CellText(item) & "' and shift_number = " _
                & TDBGrid_ListEmp.Columns("shift_number").Value & " and employee_code = '" _
                & TDBGrid_ListEmp.Columns("employee_code").CellText(item) & "' and group_code = '" _
                & TDBCombo_group_wt.Text & "'"
        End If
        
        CnG.Execute "delete from td_shift2 where company_code = '" _
            & TDBGrid_ListEmp.Columns("company_code").CellText(item) & "' and shift_number = " _
            & TDBGrid_ListEmp.Columns("shift_number").Value & " and employee_code = '" _
            & TDBGrid_ListEmp.Columns("employee_code").CellText(item) & "' and date(start_date) = '" _
            & Format(TDBGrid_ListEmp.Columns("start_date").Value, "yyyy-MM-dd") & "'"
    Next
    CnG.CommitTrans
    Call load_data_detail_working_time
    MsgBox i & " Data Karyawan Berhasil Dihapus...", vbInformation, headerMSG
    
    Exit Sub

Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
End Sub


Private Sub TDBGrid_WT_FilterChange()
    Call filter_change
End Sub

Private Sub TDBGrid_WD_FilterChange()
    Call filter_change
End Sub

Private Sub TDBGrid_EmpWT_FilterChange()
    Call filter_change
End Sub

Private Sub TDBGrid_Group_FilterChange()
    Call filter_change
End Sub


Private Sub txt_group_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_group_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txt_shift_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_shift_name_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_late_value_Validate(Cancel As Boolean)
    If Not Trim(txt_late_value) = "" Then
        txt_late_value = FormatNumber(DropAllComma(txt_late_value))
    End If
End Sub

