VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D62356B-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODL6.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frm_mst_thr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MASTER THR"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11670
   Icon            =   "frmMstTHR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6435
      Left            =   90
      TabIndex        =   8
      Top             =   780
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   11351
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "MASTER THR"
      TabPicture(0)   =   "frmMstTHR.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TDBGrid_Thr"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_entry1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TDBCombo_thr"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txt_thr_name_type"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra_entry_thr"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "TUNJANGAN"
      TabPicture(1)   =   "frmMstTHR.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TDBGrid_Tunj"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "fra_entry_tunj"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "THR STATUS"
      TabPicture(2)   =   "frmMstTHR.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TDBGrid_Thr_Status"
      Tab(2).Control(1)=   "frmTombol"
      Tab(2).Control(2)=   "fra_entry"
      Tab(2).ControlCount=   3
      Begin VB.Frame fra_entry_tunj 
         Height          =   2655
         Left            =   -74730
         TabIndex        =   58
         Top             =   2040
         Width           =   10875
         Begin VB.TextBox txt_status_tunj_name 
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
            Height          =   375
            Left            =   4380
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   64
            Top             =   990
            Width           =   2655
         End
         Begin VB.ComboBox cboTunj 
            Height          =   315
            ItemData        =   "frmMstTHR.frx":05DE
            Left            =   4380
            List            =   "frmMstTHR.frx":05EE
            TabIndex        =   61
            Text            =   "Combo1"
            Top             =   1410
            Width           =   1725
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
            TabIndex        =   59
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_status_tunj 
            Height          =   375
            Left            =   4380
            OleObjectBlob   =   "frmMstTHR.frx":0619
            TabIndex        =   65
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "THR*"
            Height          =   195
            Left            =   3870
            TabIndex        =   66
            Top             =   660
            Width           =   405
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TUNJANGAN*"
            Height          =   195
            Left            =   3255
            TabIndex        =   62
            Top             =   1440
            Width           =   1050
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74730
         TabIndex        =   52
         Top             =   4800
         Width           =   10875
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew_Tunj 
            Height          =   705
            Left            =   720
            TabIndex        =   53
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
            MICON           =   "frmMstTHR.frx":257B
            PICN            =   "frmMstTHR.frx":2597
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Tunj 
            Height          =   705
            Left            =   1710
            TabIndex        =   54
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
            MICON           =   "frmMstTHR.frx":3629
            PICN            =   "frmMstTHR.frx":3645
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Tunj 
            Height          =   705
            Left            =   2730
            TabIndex        =   55
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
            MICON           =   "frmMstTHR.frx":46D7
            PICN            =   "frmMstTHR.frx":46F3
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Tunj 
            Height          =   705
            Left            =   3750
            TabIndex        =   56
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
            MICON           =   "frmMstTHR.frx":5785
            PICN            =   "frmMstTHR.frx":57A1
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Tunj 
            Height          =   705
            Left            =   4770
            TabIndex        =   57
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
            MICON           =   "frmMstTHR.frx":6833
            PICN            =   "frmMstTHR.frx":684F
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
         Height          =   2655
         Left            =   -74730
         TabIndex        =   43
         Top             =   2100
         Width           =   10875
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
            TabIndex        =   46
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_status_name 
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
            Height          =   375
            Left            =   4500
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   45
            Top             =   960
            Width           =   2655
         End
         Begin VB.TextBox txt_thr_status_name 
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
            Height          =   375
            Left            =   4500
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   1740
            Width           =   2655
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_status 
            Height          =   375
            Left            =   4500
            OleObjectBlob   =   "frmMstTHR.frx":78E1
            TabIndex        =   47
            Top             =   600
            Width           =   1695
         End
         Begin TrueOleDBList60.TDBCombo TDBCombo_thr_status 
            Height          =   375
            Left            =   4500
            OleObjectBlob   =   "frmMstTHR.frx":983E
            TabIndex        =   48
            Top             =   1380
            Width           =   1695
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "STATUS*"
            Height          =   195
            Left            =   3270
            TabIndex        =   50
            Top             =   630
            Width           =   1125
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "THR*"
            Height          =   195
            Left            =   3990
            TabIndex        =   49
            Top             =   1410
            Width           =   405
         End
      End
      Begin VB.Frame frmTombol 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   -74730
         TabIndex        =   37
         Top             =   4860
         Width           =   10875
         Begin VB.Timer timer1 
            Enabled         =   0   'False
            Interval        =   600
            Left            =   120
            Top             =   360
         End
         Begin prj_panji.vbButton cmdNew_Status 
            Height          =   705
            Left            =   720
            TabIndex        =   38
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
            MICON           =   "frmMstTHR.frx":B79F
            PICN            =   "frmMstTHR.frx":B7BB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdSave_Status 
            Height          =   705
            Left            =   1710
            TabIndex        =   39
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
            MICON           =   "frmMstTHR.frx":C84D
            PICN            =   "frmMstTHR.frx":C869
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdEdit_Status 
            Height          =   705
            Left            =   2730
            TabIndex        =   40
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
            MICON           =   "frmMstTHR.frx":D8FB
            PICN            =   "frmMstTHR.frx":D917
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_Status 
            Height          =   705
            Left            =   3750
            TabIndex        =   41
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
            MICON           =   "frmMstTHR.frx":E9A9
            PICN            =   "frmMstTHR.frx":E9C5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdCancel_Status 
            Height          =   705
            Left            =   4770
            TabIndex        =   42
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
            MICON           =   "frmMstTHR.frx":FA57
            PICN            =   "frmMstTHR.frx":FA73
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
      Begin VB.Frame fra_entry_thr 
         Height          =   2775
         Left            =   90
         TabIndex        =   9
         Top             =   2100
         Visible         =   0   'False
         Width           =   11295
         Begin VB.CheckBox chk_tunj 
            Caption         =   "DITAMBAH TUNJANGAN?"
            Height          =   225
            Left            =   8070
            TabIndex        =   63
            Top             =   390
            Width           =   2505
         End
         Begin VB.ComboBox cbo_umk 
            Height          =   315
            ItemData        =   "frmMstTHR.frx":10B05
            Left            =   9360
            List            =   "frmMstTHR.frx":10B0F
            TabIndex        =   36
            Text            =   "UMK"
            Top             =   1380
            Width           =   1095
         End
         Begin VB.CheckBox chk_proporsional 
            Caption         =   "PROPORSIONAL"
            Height          =   225
            Left            =   6120
            TabIndex        =   35
            Top             =   390
            Width           =   1755
         End
         Begin VB.TextBox txt_naik 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1740
            Width           =   1695
         End
         Begin VB.TextBox txt_thr_value 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1380
            Width           =   1695
         End
         Begin VB.TextBox txt_thr_upper 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1020
            Width           =   1695
         End
         Begin VB.TextBox txt_thr_description 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   7
            Top             =   2100
            Width           =   4305
         End
         Begin VB.TextBox txt_thr_under 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   6120
            MaxLength       =   50
            TabIndex        =   2
            Top             =   660
            Width           =   1695
         End
         Begin VB.TextBox txt_thr_number 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   1
            Top             =   630
            Width           =   1695
         End
         Begin VB.CheckBox chk_percentage 
            Caption         =   "PROSENTASE"
            Height          =   225
            Left            =   7890
            TabIndex        =   5
            Top             =   1410
            Width           =   1545
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KENAIKAN PER TAHUN*"
            Height          =   195
            Left            =   3945
            TabIndex        =   23
            Top             =   1740
            Width           =   1860
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NILAI*"
            Height          =   195
            Left            =   4590
            TabIndex        =   14
            Top             =   1380
            Width           =   1215
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "SAMPAI (BULAN)*"
            Height          =   195
            Left            =   4470
            TabIndex        =   13
            Top             =   1020
            Width           =   1335
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KETERANGAN"
            Height          =   195
            Left            =   4710
            TabIndex        =   12
            Top             =   2130
            Width           =   1110
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "NO.*"
            Height          =   195
            Left            =   600
            TabIndex        =   11
            Top             =   630
            Width           =   345
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "DARI (BULAN)*"
            Height          =   225
            Left            =   4470
            TabIndex        =   10
            Top             =   660
            Width           =   1335
         End
      End
      Begin VB.TextBox txt_thr_name_type 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000B&
         Height          =   315
         Left            =   3300
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   24
         Top             =   510
         Width           =   3855
      End
      Begin VB.Frame Frame5 
         Caption         =   "Data Control Button"
         Height          =   1335
         Left            =   90
         TabIndex        =   15
         Top             =   5010
         Width           =   11295
         Begin prj_panji.vbButton cmdNew 
            Height          =   705
            Left            =   540
            TabIndex        =   16
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
            MICON           =   "frmMstTHR.frx":10B1C
            PICN            =   "frmMstTHR.frx":10B38
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
            TabIndex        =   17
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Simpan Dtl"
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
            MICON           =   "frmMstTHR.frx":11BCA
            PICN            =   "frmMstTHR.frx":11BE6
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
            TabIndex        =   18
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Ubah Dtl"
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
            MICON           =   "frmMstTHR.frx":12C78
            PICN            =   "frmMstTHR.frx":12C94
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
            TabIndex        =   19
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
            MICON           =   "frmMstTHR.frx":13D26
            PICN            =   "frmMstTHR.frx":13D42
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
            TabIndex        =   20
            Top             =   360
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1244
            BTYPE           =   14
            TX              =   "&Batal Dtl"
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
            MICON           =   "frmMstTHR.frx":14DD4
            PICN            =   "frmMstTHR.frx":14DF0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton CmdNew_Master 
            Height          =   705
            Left            =   8850
            TabIndex        =   33
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
            MICON           =   "frmMstTHR.frx":15E82
            PICN            =   "frmMstTHR.frx":15E9E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin prj_panji.vbButton cmdDelete_All 
            Height          =   705
            Left            =   9870
            TabIndex        =   34
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
            MICON           =   "frmMstTHR.frx":16F30
            PICN            =   "frmMstTHR.frx":16F4C
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
      Begin TrueOleDBList60.TDBCombo TDBCombo_thr 
         Height          =   375
         Left            =   1500
         OleObjectBlob   =   "frmMstTHR.frx":17FDE
         TabIndex        =   25
         Top             =   510
         Width           =   1695
      End
      Begin VB.Frame fra_entry1 
         Height          =   1815
         Left            =   90
         TabIndex        =   27
         Top             =   3060
         Visible         =   0   'False
         Width           =   11295
         Begin VB.TextBox txt_name_thr 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4890
            MaxLength       =   50
            TabIndex        =   32
            Top             =   930
            Width           =   3495
         End
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
            TabIndex        =   28
            Top             =   120
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.TextBox txt_thr_code 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4890
            MaxLength       =   10
            TabIndex        =   30
            Top             =   570
            Width           =   1695
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "KODE THR"
            Height          =   195
            Left            =   3480
            TabIndex        =   31
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "NAMA THR"
            Height          =   195
            Left            =   3450
            TabIndex        =   29
            Top             =   960
            Width           =   1245
         End
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Thr 
         Height          =   3885
         Left            =   90
         TabIndex        =   21
         Top             =   990
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   6853
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NO."
         Columns(0).DataField=   "thr_number"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "DARI (BULAN)"
         Columns(1).DataField=   "thr_under"
         Columns(1).NumberFormat=   "Standard"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "SAMPAI (BULAN)"
         Columns(2).DataField=   "thr_upper"
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NILAI"
         Columns(3).DataField=   "thr_value"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   4
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "PROSENTASE"
         Columns(4).DataField=   "flag_percentage"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "KENAIKAN"
         Columns(5).DataField=   "thr_naik"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "KETERANGAN"
         Columns(6).DataField=   "description"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "FLAG TUNJ"
         Columns(7).DataField=   "flag_tunj"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1455"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1376"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=513"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2275"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2196"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=514"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2381"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2302"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=514"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2540"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2461"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2011"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=1931"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=513"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2672"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2593"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=5556"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=5477"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=516"
         Splits(0)._ColumnProps(40)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(41)=   "Column(7).Order=8"
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
         Caption         =   "DAFTAR THR"
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
         _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(55)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(6).Style:id=54,.parent=13,.alignment=3"
         _StyleDefs(59)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Thr_Status 
         Height          =   4245
         Left            =   -74730
         TabIndex        =   51
         Top             =   510
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE STATUS"
         Columns(0).DataField=   "status_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA STATUS"
         Columns(1).DataField=   "status_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KODE THR"
         Columns(2).DataField=   "thr_code"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NAMA THR"
         Columns(3).DataField=   "thr_name"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2990"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2910"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=5874"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5794"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=6826"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=6747"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
         Caption         =   "DAFTAR THR PER STATUS"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid TDBGrid_Tunj 
         Height          =   4245
         Left            =   -74730
         TabIndex        =   60
         Top             =   450
         Width           =   10875
         _ExtentX        =   19182
         _ExtentY        =   7488
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE THR"
         Columns(0).DataField=   "thr_code"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA THR"
         Columns(1).DataField=   "thr_name"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KD. TUNJ"
         Columns(2).DataField=   "flag_tunj"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NAMA TUNJANGAN"
         Columns(3).DataField=   "tunj_name"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2990"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2910"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=5874"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=5794"
         Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=6826"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=6747"
         Splits(0)._ColumnProps(21)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
         Caption         =   "DAFTAR TUNJANGAN UNTUK THR"
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
         _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "TIPE THR"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   570
         Width           =   1200
      End
   End
   Begin prj_panji.vbButton cmdKeluar 
      Height          =   705
      Left            =   10530
      TabIndex        =   22
      Top             =   7290
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
      MICON           =   "frmMstTHR.frx":19F38
      PICN            =   "frmMstTHR.frx":19F54
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MASTER THR"
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
      TabIndex        =   0
      Top             =   150
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   0
      Picture         =   "frmMstTHR.frx":1AFE6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12690
   End
End
Attribute VB_Name = "frm_mst_thr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsThr As New ADODB.Recordset
Dim rsThr_Detail As New ADODB.Recordset

Dim rsThrStatus_List As New ADODB.Recordset
Dim rsStatus As New ADODB.Recordset
Dim rsThrStatus As New ADODB.Recordset
Dim rsThrTunj As New ADODB.Recordset

Dim int_mode As Integer
Dim Col As TrueOleDBGrid70.Column
Dim Cols As TrueOleDBGrid70.Columns
Dim strsql As String

Dim vFlagTunj As Integer

Private Function check_validate_exist_new() As Boolean
Dim rs As New ADODB.Recordset
Dim str_sql As String
    check_validate_exist_new = False
    
    If SSTab1.Tab = 0 Then
        str_sql = "select count(thr_number) as rec_count from m_thr_detail where thr_number = '" & Trim(txt_thr_number) & "' " & _
                    "AND thr_code = '" & TDBCombo_thr.Text & "'"

        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
        
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        str_sql = "select count(thr_code) as rec_count from m_thr_tunj where thr_code = '" & TDBCombo_status_tunj.Text & "' " & _
                    "AND flag_tunj = '" & cboTunj.ListIndex & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
        
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        str_sql = "select count(status_code) as rec_count from m_thr_status where status_code = '" & TDBCombo_status.Text & "'"
        rs.Open str_sql, CnG, adOpenStatic, adLockReadOnly
        
        If rs.Fields("rec_count").Value > 0 Then
            check_validate_exist_new = True
            Exit Function
        End If
    End If
End Function

Private Sub check_invalid()
    MsgBox "Data Sudah Ada...", vbCritical, headerMSG
    
    If SSTab1.Tab = 0 Then
        txt_thr_number = ""
        If txt_thr_number.Enabled = True Then txt_thr_number.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        TDBCombo_status_tunj.Text = ""
        If TDBCombo_status_tunj.Enabled = True Then TDBCombo_status_tunj.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        TDBCombo_status.Text = ""
        If TDBCombo_status.Enabled = True Then TDBCombo_status.SetFocus
    End If
End Sub

Private Function check_validate_exist_edit() As Boolean
    check_validate_exist_edit = False
    
    If SSTab1.Tab = 0 Then
        If Not txt_thr_number.Text = rsThr_Detail.Fields("thr_number").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        If Not TDBCombo_status_tunj.Text = rsThrTunj.Fields("thr_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        If Not TDBCombo_status.Text = rsThrStatus_List.Fields("status_code").Value And _
        check_validate_exist_new Then
            check_validate_exist_edit = True
            Exit Function
        End If
    End If
End Function

Private Function check_validate_new() As Boolean
    check_validate_new = True
    
    If SSTab1.Tab = 0 Then
        'validasi thr number
        If Trim(txt_thr_number.Text) = "" Then
            MsgBox "No. THR Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_thr_number.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi thr under
        If Trim(txt_thr_under.Text) = "" Then
            MsgBox "Kolom Dari Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_thr_under.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi thr upper
        If Trim(txt_thr_upper.Text) = "" Then
            MsgBox "Kolom Sampai Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_thr_upper.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi thr value
        If Trim(txt_thr_value.Text) = "" Then
            MsgBox "Nilai THR Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_thr_value.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi thr naik
        If Trim(txt_naik.Text) = "" Then
            MsgBox "Kolom Kenaikan Per Tahun Masih Kosong...", vbOKOnly + vbInformation, headerMSG
            txt_naik.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 1 Then
        'validasi status
        If Trim(TDBCombo_status_tunj.Text) = "" Then
            MsgBox "Kode Status Masih Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_status_tunj.SetFocus
            check_validate_new = False
            Exit Function
        End If
    ElseIf SSTab1.Tab = 2 Then
        'validasi status
        If Trim(TDBCombo_status.Text) = "" Then
            MsgBox "Kode Status Masih Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_status.SetFocus
            check_validate_new = False
            Exit Function
        End If
        
        'validasi thr
        If Trim(TDBCombo_thr_status.Text) = "" Then
            MsgBox "Kode THR Masih Belum Dipilih...", vbOKOnly + vbInformation, headerMSG
            TDBCombo_thr_status.SetFocus
            check_validate_new = False
            Exit Function
        End If
    End If
End Function

Private Sub load_data()
    If SSTab1.Tab = 0 Then
        If rsThr_Detail.State Then rsThr_Detail.Close
        SQL = "select * from m_thr_detail " & _
                "where thr_code = '" & TDBCombo_thr.Text & "' " & _
                "order by thr_number"
        rsThr_Detail.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Thr.DataSource = rsThr_Detail
    ElseIf SSTab1.Tab = 1 Then
        If rsThrTunj.State Then rsThrTunj.Close
        SQL = "select a.*,b.thr_name from m_thr_tunj a join m_thr b on a.thr_code = b.thr_code " & _
              "order by thr_code"
        rsThrTunj.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Tunj.DataSource = rsThrTunj
    ElseIf SSTab1.Tab = 2 Then
        If rsThrStatus_List.State Then rsThrStatus_List.Close
        SQL = "select a.*,b.status_name,c.thr_name from m_thr_status a join m_emp_status b on a.status_code = b.status_code " & _
              "join m_thr c on a.thr_code = c.thr_code " & _
              "order by status_code"
        rsThrStatus_List.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
        
        TDBGrid_Thr_Status.DataSource = rsThrStatus_List
    End If
End Sub

Private Sub load_data_thr()
    If rsThr.State Then rsThr.Close
    SQL = "select * from m_thr order by thr_code"
    rsThr.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_thr.RowSource = rsThr
End Sub

Private Sub load_data_status()
    If rsStatus.State Then rsStatus.Close
    SQL = "select * from m_emp_status order by status_code"
    rsStatus.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_status.RowSource = rsStatus
End Sub

Private Sub load_data_thr_status()
    If rsThrStatus.State Then rsThrStatus.Close
    SQL = "select * from m_thr order by thr_code"
    rsThrStatus.Open SQL, CnG, adOpenForwardOnly, adLockReadOnly
    
    TDBCombo_thr_status.RowSource = rsThrStatus
    TDBCombo_status_tunj.RowSource = rsThrStatus
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
        If Not (TDBGrid_Thr.ApproxCount > 0 And TDBGrid_Thr.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Thr.Columns("thr_value").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.Execute "delete from m_thr_detail where thr_code = '" & TDBCombo_thr.Text & "' and thr_number = '" _
                        & Replace(TDBGrid_Thr.Columns("thr_number").Value, "'", "''") & "'"
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_Tunj.ApproxCount > 0 And TDBGrid_Tunj.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Tunj.Columns("tunj_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.Execute "delete from m_thr_tunj where thr_code = '" & TDBGrid_Tunj.Columns("thr_code").Value & "' " & _
                        "AND flag_tunj = '" & cboTunj.ListIndex & "'"
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_Thr_Status.ApproxCount > 0 And TDBGrid_Thr_Status.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            Exit Sub
        End If
        
        i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & TDBGrid_Thr_Status.Columns("status_name").Value & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.Execute "delete from m_thr_status where status_code = '" & TDBGrid_Thr_Status.Columns("status_code").Value & "'"
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
        If Not (TDBGrid_Thr.ApproxCount > 0 And TDBGrid_Thr.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsThr_Detail
            txt_thr_number.Text = .Fields("thr_number").Value
            txt_thr_under.Text = FormatNumber(.Fields("thr_under").Value)
            txt_thr_upper.Text = FormatNumber(.Fields("thr_upper").Value)
            chk_proporsional.Value = IIf(IsNull(.Fields("flag_proporsional").Value), 0, .Fields("flag_proporsional").Value)
            chk_tunj.Value = IIf(IsNull(.Fields("flag_tunj").Value), 0, .Fields("flag_tunj").Value)
            chk_percentage.Value = FormatNumber(.Fields("flag_percentage").Value)
            cbo_umk.ListIndex = IIf(IsNull(.Fields("flag_umk").Value), 0, .Fields("flag_umk").Value)
            txt_thr_value.Text = FormatNumber(.Fields("thr_value").Value)
            txt_naik.Text = FormatNumber(.Fields("thr_naik").Value)
            txt_thr_description.Text = .Fields("description").Value
        End With
    ElseIf SSTab1.Tab = 1 Then
        If Not (TDBGrid_Tunj.ApproxCount > 0 And TDBGrid_Tunj.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsThrTunj
            Call set_data_thr(.Fields("thr_code").Value)
            cboTunj.ListIndex = .Fields("flag_tunj").Value
            vFlagTunj = .Fields("flag_tunj").Value
        End With
    ElseIf SSTab1.Tab = 2 Then
        If Not (TDBGrid_Thr_Status.ApproxCount > 0 And TDBGrid_Thr_Status.Bookmark > 0) Then
            MsgBox "Tidak Ada Data Yang Dipilih...", vbInformation, headerMSG
            vSetData = 0
            Exit Sub
        End If
        
        With rsThrStatus_List
            Call set_data_status(.Fields("status_code").Value)
            Call set_data_thr(.Fields("thr_code").Value)
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

Private Sub cmdDelete_All_Click()
Dim i As Integer

On Error GoTo Err
    i = MsgBox("Apakah Yakin Akan Menghapus Data '" _
            & txt_thr_name_type.Text & "' ?", vbYesNo + vbQuestion, headerMSG)
        If Not i = vbYes Then Exit Sub
        
        CnG.BeginTrans
        CnG.Execute "delete from m_thr_detail where " _
                & "thr_code = '" & TDBCombo_thr.Text & "'"
        CnG.Execute "delete from m_thr where thr_code = " _
                & "'" & TDBCombo_thr.Text & "'"
        
        CnG.CommitTrans
        
        Call load_mode
        Call load_data_thr
        int_mode = 0
        Call load_mode
        
        TDBCombo_thr.Text = ""
        txt_thr_name_type.Text = ""
        Set TDBGrid_Thr.DataSource = Nothing
        
        Exit Sub
Err:
CnG.RollbackTrans: MsgBox Err.Description, vbExclamation, headerMSG
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
        SQL = "INSERT INTO m_thr_detail (thr_code,thr_number,thr_under," & _
                "thr_upper,flag_proporsional,flag_tunj,flag_percentage,flag_umk,thr_value,thr_naik,description,entry_date,entry_user) " & _
              "VALUES( " & _
                "'" & TDBCombo_thr.Text & "','" & Trim(txt_thr_number.Text) & "','" & Val(DropAllComma(txt_thr_under)) & "'," & _
                "'" & Val(DropAllComma(txt_thr_upper)) & "', '" & chk_tunj.Value & "','" & chk_proporsional.Value & "','" & chk_percentage.Value & "'," & _
                "'" & cbo_umk.ListIndex & "','" & Val(DropAllComma(txt_thr_value)) & "'," & _
                "'" & Val(DropAllComma(txt_naik.Text)) & "','" & Trim(txt_thr_description) & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "INSERT INTO m_thr_tunj (thr_code,flag_tunj,tunj_name,entry_date,entry_user) " & _
              "VALUES( " & _
                "'" & TDBCombo_status_tunj.Text & "','" & cboTunj.ListIndex & "','" & cboTunj.Text & "',now(),'" & LOGIN_NAME & "')"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "INSERT INTO m_thr_status (status_code,thr_code,entry_date,entry_user) " & _
              "VALUES( " & _
                "'" & TDBCombo_status.Text & "','" & TDBCombo_thr_status.Text & "',now(),'" & LOGIN_NAME & "')"
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
        SQL = "UPDATE m_thr_detail SET thr_number = '" & Trim(txt_thr_number.Text) & "'," & _
                "thr_under = '" & Val(DropAllComma(txt_thr_under.Text)) & "'," & _
                "thr_upper = '" & Val(DropAllComma(txt_thr_upper.Text)) & "'," & _
                "flag_proporsional = '" & chk_proporsional.Value & "'," & _
                "flag_tunj = '" & chk_tunj.Value & "'," & _
                "flag_percentage = '" & chk_percentage.Value & "'," & _
                "flag_umk = '" & cbo_umk.ListIndex & "'," & _
                "thr_value = '" & Val(DropAllComma(txt_thr_value.Text)) & "'," & _
                "thr_naik = '" & Val(DropAllComma(txt_naik.Text)) & "'," & _
                "description = '" & Trim(txt_thr_description.Text) & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_CODE & "' " & _
              "WHERE thr_code = '" & TDBCombo_thr.Text & "' AND thr_number = '" & Trim(txt_thr_number.Text) & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 1 Then
        SQL = "UPDATE m_thr_tunj SET thr_code = '" & TDBCombo_status_tunj.Text & "'," & _
                "flag_tunj = '" & cboTunj.ListIndex & "'," & _
                "tunj_name = '" & cboTunj.Text & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_CODE & "' " & _
              "WHERE thr_code = '" & TDBCombo_status_tunj.Text & "' " & _
                "AND flag_tunj = '" & vFlagTunj & "'"
        CnG.Execute SQL
    ElseIf SSTab1.Tab = 2 Then
        SQL = "UPDATE m_thr_status SET status_code = '" & TDBCombo_status.Text & "'," & _
                "thr_code = '" & TDBCombo_thr_status.Text & "'," & _
                "edit_date = now(),edit_user = '" & LOGIN_CODE & "' " & _
              "WHERE status_code = '" & TDBCombo_status.Text & "'"
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
        cmdNew.Enabled = a And blnUser_Add
        cmdSave.Enabled = b
        cmdEdit.Enabled = c And blnUser_Edit
        cmdDelete.Enabled = d And blnUser_Delete
        cmdCancel.Enabled = e
    ElseIf SSTab1.Tab = 2 Then
        cmdNew_Tunj.Enabled = a And blnUser_Add
        cmdSave_Tunj.Enabled = b
        cmdEdit_Tunj.Enabled = c And blnUser_Edit
        cmdDelete_Tunj.Enabled = d And blnUser_Delete
        cmdCancel_Tunj.Enabled = e
    ElseIf SSTab1.Tab = 2 Then
        cmdNew_Status.Enabled = a And blnUser_Add
        cmdSave_Status.Enabled = b
        cmdEdit_Status.Enabled = c And blnUser_Edit
        cmdDelete_Status.Enabled = d And blnUser_Delete
        cmdCancel_Status.Enabled = e
    End If
End Sub

Private Sub clear_view_data()
Dim Ctr As CONTROL
    For Each Ctr In Me
        If TypeOf Ctr Is TextBox Or TypeOf Ctr Is TDBText Then
            If Not LCase(Ctr.name) = "txt_thr_name_type" Then Ctr.Text = ""
        ElseIf TypeOf Ctr Is TDBCombo Then
            If Not LCase(Ctr.name) = "tdbcombo_thr" Then Ctr.Text = ""
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
    chk_percentage.Value = 0
    chk_proporsional.Value = 0
    cbo_umk.ListIndex = 0
End Sub

Private Sub set_data_mode()
    If int_mode = 1 Then        'NEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_thr.Visible = True
            txt_thr_number.Enabled = True
            TDBGrid_Thr.Enabled = False
            
            Call set_new_data
    
            If txt_thr_number.Enabled = True Then
                txt_thr_number.SetFocus
            End If
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_tunj.Visible = True
            TDBCombo_status_tunj.Enabled = True
            TDBGrid_Tunj.Enabled = False
            
            Call set_new_data
    
            If TDBCombo_status_tunj.Enabled = True Then
                TDBCombo_status_tunj.SetFocus
            End If
        ElseIf SSTab1.Tab = 2 Then
            fra_entry.Visible = True
            TDBCombo_status.Enabled = True
            TDBGrid_Thr_Status.Enabled = False
            
            Call set_new_data
    
            If TDBCombo_status.Enabled = True Then
                TDBCombo_status.SetFocus
            End If
        End If
        
    ElseIf int_mode = 0 Then    'VIEW
        Call clear_view_data
        
        If SSTab1.Tab = 0 Then
            fra_entry_thr.Visible = False
            TDBGrid_Thr.Enabled = True
        ElseIf SSTab1.Tab = 1 Then
            fra_entry_tunj.Visible = False
            TDBGrid_Tunj.Enabled = True
        ElseIf SSTab1.Tab = 2 Then
            fra_entry.Visible = False
            TDBGrid_Thr_Status.Enabled = True
        End If
    
    ElseIf int_mode = 2 Then    'EDIT
        Call set_edit_data
        
        If vSetData = 0 Then
            int_mode = 0
            Call load_mode
            Exit Sub
        End If
        
        If SSTab1.Tab = 0 Then
            txt_thr_number.Enabled = False
            fra_entry_thr.Visible = True
            TDBGrid_Thr.Enabled = False
        ElseIf SSTab1.Tab = 1 Then
            TDBCombo_status_tunj.Enabled = False
            fra_entry_tunj.Visible = True
            TDBGrid_Tunj.Enabled = False
        ElseIf SSTab1.Tab = 2 Then
            TDBCombo_status.Enabled = False
            fra_entry.Visible = True
            TDBGrid_Thr_Status.Enabled = False
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

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub CmdNew_Master_Click()
Dim SQL As String

    If CmdNew_Master.Caption = "&Tambah" Then
        CmdNew_Master.Caption = "&Simpan"
        cmdCancel.Caption = "&Batal"
        Call set_buttons_enable(False, False, False, False, True, False, False)
        fra_entry1.Visible = True
        fra_entry_thr.Visible = False
        
        txt_thr_code.Text = ""
        txt_name_thr.Text = ""
        txt_thr_code.SetFocus
        
    Else
        SQL = "INSERT INTO m_thr(thr_code,thr_name) " _
                & "VALUES ('" & txt_thr_code.Text & "','" & txt_name_thr.Text & "')"
        CnG.Execute SQL
        
        Call set_buttons_enable(True, False, True, True, False, True, True)
        CmdNew_Master.Caption = "&Tambah"
        cmdCancel.Caption = "&Batal Dtl"
            
        fra_entry1.Visible = False
        
        txt_thr_code.Text = ""
        txt_name_thr.Text = ""
        txt_thr_name_type.Text = ""
    
        Call load_data_thr
    End If
End Sub

Private Sub Form_Load()
    SSTab1.Tab = 0
    
    Call load_data_thr
    cboTunj.ListIndex = 0
    
    Call load_data_user_access(Me)
    int_mode = 0

    Call load_mode
End Sub

Private Sub clear_filter()
    If SSTab1.Tab = 0 Then
        For Each Col In TDBGrid_Thr.Columns
            Col.FilterText = ""
        Next Col
        rsThr_Detail.Filter = adFilterNone
    ElseIf SSTab1.Tab = 1 Then
        For Each Col In TDBGrid_Tunj.Columns
            Col.FilterText = ""
        Next Col
        rsThrTunj.Filter = adFilterNone
    ElseIf SSTab1.Tab = 2 Then
        For Each Col In TDBGrid_Thr_Status.Columns
            Col.FilterText = ""
        Next Col
        rsThrStatus_List.Filter = adFilterNone
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
        Set Cols = TDBGrid_Thr.Columns
        i = TDBGrid_Thr.Col
        TDBGrid_Thr.HoldFields
        
        rsThr_Detail.Filter = getFilter()
        TDBGrid_Thr.Col = i
        TDBGrid_Thr.EditActive = True
        
        TDBGrid_Thr.SelStart = Len(TDBGrid_Thr.Columns(i).FilterText)
        If TDBGrid_Thr.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Thr.Col = i
        End If
    ElseIf SSTab1.Tab = 1 Then
        Set Cols = TDBGrid_Tunj.Columns
        i = TDBGrid_Tunj.Col
        TDBGrid_Tunj.HoldFields
        
        rsThrTunj.Filter = getFilter()
        TDBGrid_Tunj.Col = i
        TDBGrid_Tunj.EditActive = True
        
        TDBGrid_Tunj.SelStart = Len(TDBGrid_Tunj.Columns(i).FilterText)
        If TDBGrid_Tunj.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Tunj.Col = i
        End If
    ElseIf SSTab1.Tab = 2 Then
        Set Cols = TDBGrid_Thr_Status.Columns
        i = TDBGrid_Thr_Status.Col
        TDBGrid_Thr_Status.HoldFields
        
        rsThrStatus_List.Filter = getFilter()
        TDBGrid_Thr_Status.Col = i
        TDBGrid_Thr_Status.EditActive = True
        
        TDBGrid_Thr_Status.SelStart = Len(TDBGrid_Thr_Status.Columns(i).FilterText)
        If TDBGrid_Thr_Status.ApproxCount < 1 Then
            Call clear_filter
            TDBGrid_Thr_Status.Col = i
        End If
    End If

    Exit Sub
    
Err:
MsgBox "Data Tidak Ditemukan Pada Kolom Ini " & vbCr _
& "Atau Filter Data Tidak Sesuai...", vbCritical, headerMSG
Call clear_filter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm_mst_thr = Nothing
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


Private Sub cmdNew_Status_Click()
    Call new_data
End Sub

Private Sub cmdSave_Status_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Status_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Status_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Status_Click()
    Call cancel_data
End Sub


Private Sub cmdNew_Tunj_Click()
    Call new_data
End Sub

Private Sub cmdSave_Tunj_Click()
    Call simpan_data
End Sub

Private Sub cmdEdit_Tunj_Click()
    Call edit_data
End Sub

Private Sub cmdDelete_Tunj_Click()
    Call delete_data
End Sub

Private Sub cmdCancel_Tunj_Click()
    Call cancel_data
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        txt_thr_name_type.Text = ""
        
        Call load_data_thr
        
        Call load_data_user_access(Me)
        int_mode = 0
    
        Call load_mode
    ElseIf SSTab1.Tab = 1 Then
        Call load_data_thr_status
        
        Call load_data
        
        Call load_data_user_access(Me)
        int_mode = 0
    
        Call load_mode
    ElseIf SSTab1.Tab = 2 Then
        Call load_data_status
        Call load_data_thr_status
        
        Call load_data
        
        Call load_data_user_access(Me)
        int_mode = 0
    
        Call load_mode
    End If
End Sub

Private Sub TDBCombo_status_ItemChange()
    If TDBCombo_status.ApproxCount > 0 Then
        TDBCombo_status.Text = TDBCombo_status.Columns("status_code").Value
        txt_status_tunj_name.Text = TDBCombo_status.Columns("status_name").Value
    End If
End Sub

Private Sub TDBCombo_status_tunj_ItemChange()
    If TDBCombo_status_tunj.ApproxCount > 0 Then
        TDBCombo_status_tunj.Text = TDBCombo_status_tunj.Columns("thr_code").Value
        txt_status_tunj_name.Text = TDBCombo_status_tunj.Columns("thr_name").Value
    End If
End Sub

Private Sub TDBCombo_thr_status_ItemChange()
    If TDBCombo_thr_status.ApproxCount > 0 Then
        TDBCombo_thr_status.Text = TDBCombo_thr_status.Columns("thr_code").Value
        txt_thr_status_name.Text = TDBCombo_thr_status.Columns("thr_name").Value
    End If
End Sub

Private Sub TDBCombo_thr_ItemChange()
    If TDBCombo_thr.ApproxCount > 0 Then
        TDBCombo_thr.Text = TDBCombo_thr.Columns("thr_code").Value
        txt_thr_name_type.Text = TDBCombo_thr.Columns("thr_name").Value
        
        Call load_data
    End If
End Sub

Private Sub TDBGrid_thr_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Thr_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim iFlagTunj As Integer
    iFlagTunj = TDBGrid_Thr.Columns("flag_tunj").Value
    If iFlagTunj = 0 Then
        SSTab1.TabEnabled(1) = False
    Else
        SSTab1.TabEnabled(1) = True
    End If
End Sub

Private Sub TDBGrid_thr_status_FilterChange()
    Call grid_filter
End Sub

Private Sub TDBGrid_Tunj_FilterChange()
    Call grid_filter
End Sub

Private Sub txt_thr_value_Validate(Cancel As Boolean)
    If Not Trim(txt_thr_value) = "" Then
        txt_thr_value = FormatNumber(DropAllComma(txt_thr_value))
    End If
End Sub

Private Sub txt_naik_Validate(Cancel As Boolean)
    If Not Trim(txt_naik) = "" Then
        txt_naik = FormatNumber(DropAllComma(txt_naik))
    End If
End Sub

Private Sub txt_thr_code_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txt_name_thr_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub set_data_status(ByVal str_code As String)
On Error GoTo Err
    
    rsStatus.MoveFirst
    rsStatus.Find ("status_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (rsStatus.EOF = True Or rsStatus.BOF = True) Then
        TDBCombo_status.Bookmark = rsStatus.AbsolutePosition
        Call TDBCombo_status_ItemChange
    Else
        TDBCombo_status.Text = ""
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub

Private Sub set_data_thr(ByVal str_code As String)
On Error GoTo Err
    
    rsThrStatus.MoveFirst
    rsThrStatus.Find ("thr_code='" & str_code & "'")   ', 0, adSearchForward, 1)
    If Not (rsThrStatus.EOF = True Or rsThrStatus.BOF = True) Then
        If SSTab1.Tab = 1 Then
            TDBCombo_status_tunj.Bookmark = rsThrStatus.AbsolutePosition
            Call TDBCombo_status_tunj_ItemChange
        ElseIf SSTab1.Tab = 2 Then
            TDBCombo_thr_status.Bookmark = rsThrStatus.AbsolutePosition
            Call TDBCombo_thr_status_ItemChange
        End If
    Else
        If SSTab1.Tab = 1 Then
            TDBCombo_status_tunj.Text = ""
        Else
            TDBCombo_thr_status.Text = ""
        End If
    End If
    Exit Sub

Err:
MsgBox Err.Description, vbExclamation, headerMSG
End Sub
